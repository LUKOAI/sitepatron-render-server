"""
SitePatron Deck Render Service v1.1
====================================

Flask endpoint który renderuje Sosenco-style PDF z HTML template'ów.
Apps Script wysyła POST z 6 wartościami per-klient + językiem,
serwer zwraca PDF jako binary.

Endpoints:
    GET  /health          — status + lista dostępnych template'ów
    POST /render          — render PDF (wymaga X-API-Key)

ZMIANY v1.1:
- Fallback B2B: jeśli wszystkie 4 pola B2B (B2B_QUANTITY, B2B_PRODUCT_PL,
  B2B_BUYERS_NOM, B2B_BUYERS_GEN) są puste/brakujące → endpoint
  podmienia 2 fragmenty HTML na uniwersalny tekst z kalkulatorem
  oszczędności (400 EUR / 1000 EUR vs prowizja Amazon 15%).
  Jeśli wszystkie 4 wypełnione → normalna substytucja jak v1.0.

Deploy:
    docker build -t sitepatron-render .
    docker run -p 8000:8000 -e RENDER_API_KEY=secret sitepatron-render

Lokalnie:
    pip install -r requirements.txt
    playwright install chromium
    RENDER_API_KEY=secret python3 server.py
"""

import os
import re
import json
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from playwright.sync_api import sync_playwright

app = Flask(__name__)

# ============================================================
# KONFIGURACJA
# ============================================================

API_KEY = os.environ.get("RENDER_API_KEY", "")
TEMPLATES_DIR = Path(__file__).parent / "templates"
PORT = int(os.environ.get("PORT", 8000))

# Defaulty cenowe — wpisuje endpoint, klient nie musi ich wysyłać
DEFAULT_VALUES = {
    "PRICE_MONTHLY_USD": "59",
    "PRICE_YEARLY_USD": "499",
    "PRICE_MIN_3M_USD": "177",
    "PRICE_SAVINGS_USD": "209",
    "PRICE_LOCAL_APPROX": "240",
}

# Lokalne nazwy miesięcy do auto-DECK_DATE
MONTHS = {
    "pl": ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
           "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"],
    "en": ["January", "February", "March", "April", "May", "June",
           "July", "August", "September", "October", "November", "December"],
    "de": ["Januar", "Februar", "März", "April", "Mai", "Juni",
           "Juli", "August", "September", "Oktober", "November", "Dezember"],
}

# ============================================================
# B2B FALLBACK — gdy klient nie wypełnił 4 pól B2B
# ============================================================

B2B_FIELDS = ["B2B_QUANTITY", "B2B_PRODUCT_PL", "B2B_BUYERS_NOM", "B2B_BUYERS_GEN"]

# Per-język fragmenty do podmiany.
# - bullet_old: oryginalny bullet w sekcji "Co Ci to praktycznie daje" (slide 10)
# - bullet_new: uniwersalna wersja bez {{B2B_BUYERS_NOM}}
# - example_old: cały paragraf example-box (slide 10)
# - example_new: nowy paragraf z kalkulatorem oszczędności
B2B_FALLBACK = {
    "pl": {
        "bullet_old": (
            'Klienci hurtowi ({{B2B_BUYERS_NOM}}) piszą bezpośrednio do Ciebie'
        ),
        "bullet_new": (
            'Klienci hurtowi (wspólnoty mieszkaniowe, hotele, biura, firmy z wieloma '
            'oddziałami) piszą bezpośrednio do Ciebie'
        ),
        "example_old": (
            'Klient zamawia komplet {{B2B_QUANTITY}} {{B2B_PRODUCT_PL}} (np. dla '
            '{{B2B_BUYERS_GEN}}). <strong>Na Amazon:</strong> kupuje sztuka po sztuce, '
            'każda obciążona 15% prowizją Amazon + ograniczenia platformy. '
            '<strong>Bezpośrednio od Patrona przez Stronę:</strong> jedna faktura B2B '
            'między firmami, bez prowizji, możliwość negocjacji ceny przy większej '
            'ilości. Czysty zysk + nowy kontakt biznesowy w portfolio.'
        ),
        "example_new": (
            'Klient firmowy zamawia bezpośrednio przez Stronę zamiast przez Amazon '
            '(np. wspólnota mieszkaniowa, biuro, hotel, firma z oddziałami). '
            '<strong>Na Amazon:</strong> każda sztuka obciążona ~15% prowizją + '
            'ograniczenia platformy. <strong>Bezpośrednio od Patrona:</strong> jedna '
            'faktura B2B między firmami, bez prowizji, możliwość negocjacji ceny. '
            '<strong>Liczby:</strong> zamówienie za 400 EUR = ~60 EUR oszczędności na '
            'prowizji Amazon — to pokrywa cały miesiąc Site Patron. Zamówienie za '
            '1000 EUR = ~150 EUR oszczędności = 2,5 miesiąca abonamentu z jednego '
            'zamówienia. Plus stały klient B2B w Twoim portfolio.'
        ),
    },
    # EN / DE — do dodania gdy będą template'y w tych językach.
    # Na razie EN/DE template'y nie istnieją, więc fallback EN/DE też nie jest potrzebny.
}


def get_deck_date(language: str) -> str:
    """Zwraca 'Miesiąc YYYY' w odpowiednim języku (fallback EN)."""
    now = datetime.now()
    months = MONTHS.get(language.lower(), MONTHS["en"])
    return f"{months[now.month - 1]} {now.year}"


def get_template_path(language: str) -> Path:
    """Znajdź template per-język. Fallback PL → EN → pierwszy dostępny."""
    candidates = [
        TEMPLATES_DIR / f"sitepatron-deck-{language.upper()}-template.html",
        TEMPLATES_DIR / "sitepatron-deck-PL-template.html",
        TEMPLATES_DIR / "sitepatron-deck-EN-template.html",
    ]
    for c in candidates:
        if c.exists():
            return c
    available = list(TEMPLATES_DIR.glob("*.html"))
    if available:
        return available[0]
    raise FileNotFoundError(f"Brak template'ów w {TEMPLATES_DIR}")


def render_html_to_pdf(html: str, output_path: Path) -> None:
    """Render HTML → PDF przez Playwright Chromium (headless)."""
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".html", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(html)
        tmp_path = Path(tmp.name)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(
                f"file://{tmp_path.absolute()}",
                wait_until="networkidle",
                timeout=30000,
            )
            page.wait_for_timeout(2000)  # czekaj na fonts CDN
            page.pdf(
                path=str(output_path),
                width="1280px",
                height="905px",
                print_background=True,
                margin={"top": "0", "bottom": "0", "left": "0", "right": "0"},
            )
            browser.close()
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass


def apply_b2b_fallback(html: str, values: dict, language: str) -> tuple:
    """
    Jeśli WSZYSTKIE 4 pola B2B są puste/brakujące → podmienia 2 fragmenty HTML
    na uniwersalny tekst (z kalkulatorem oszczędności).

    Jeśli choć jedno z 4 jest wypełnione → nic nie robi (zwraca HTML i values
    bez zmian, ale uzupełnia brakujące B2B_* pustym stringiem żeby walidacja
    fill_template przeszła).

    Returns: (modified_html, modified_values)
    """
    # Sprawdź czy WSZYSTKIE 4 są niepuste
    b2b_filled = {f: str(values.get(f, "")).strip() for f in B2B_FIELDS}
    all_empty = all(v == "" for v in b2b_filled.values())
    all_filled = all(v != "" for v in b2b_filled.values())

    new_values = dict(values)

    if all_empty:
        # Tryb fallback — podmień bloki HTML
        fallback = B2B_FALLBACK.get(language, B2B_FALLBACK["pl"])

        if fallback["bullet_old"] in html:
            html = html.replace(fallback["bullet_old"], fallback["bullet_new"])
        if fallback["example_old"] in html:
            html = html.replace(fallback["example_old"], fallback["example_new"])

        # Po replace HTML nie zawiera już {{B2B_*}} — fill_template będzie OK.
        # Ale dla pewności wrzucamy puste stringi do values.
        for f in B2B_FIELDS:
            new_values[f] = ""
    elif all_filled:
        # Tryb normalny — wszystko OK, fill_template podstawi 4 wartości
        pass
    else:
        # Tryb mieszany — niektóre wypełnione, inne nie. Też używamy fallback,
        # bo inaczej tekst będzie połamany ("Klient zamawia komplet  szyldów").
        fallback = B2B_FALLBACK.get(language, B2B_FALLBACK["pl"])

        if fallback["bullet_old"] in html:
            html = html.replace(fallback["bullet_old"], fallback["bullet_new"])
        if fallback["example_old"] in html:
            html = html.replace(fallback["example_old"], fallback["example_new"])

        for f in B2B_FIELDS:
            new_values[f] = ""

    return html, new_values


def fill_template(html: str, values: dict) -> str:
    """Podstaw {{KEY}} → wartości. Walidacja: wszystkie placeholdery podane."""
    placeholders_in_template = set(re.findall(r"\{\{([A-Z0-9_]+)\}\}", html))
    placeholders_in_values = set(values.keys())

    missing = placeholders_in_template - placeholders_in_values
    if missing:
        raise ValueError(f"Brakuje wartości w żądaniu: {sorted(missing)}")

    for key, value in values.items():
        html = html.replace(f"{{{{{key}}}}}", str(value))

    remaining = re.findall(r"\{\{[A-Z0-9_]+\}\}", html)
    if remaining:
        raise ValueError(f"Pozostały placeholdery po podstawieniu: {sorted(set(remaining))}")

    return html


# ============================================================
# ENDPOINTS
# ============================================================

@app.route("/health", methods=["GET"])
def health():
    """Status check + lista dostępnych template'ów."""
    templates = sorted([f.name for f in TEMPLATES_DIR.glob("*.html")])
    return jsonify({
        "status": "ok",
        "version": "1.1",
        "templates": templates,
        "api_key_required": bool(API_KEY),
        "playwright": "ready",
        "b2b_fallback_languages": sorted(B2B_FALLBACK.keys()),
    })


@app.route("/render", methods=["POST"])
def render():
    """
    Render PDF z HTML template'u.

    Body JSON:
        {
            "language": "pl",          # pl/en/de (fallback PL)
            "values": {                # per-klient
                "DEMO_URL_A": "...",         # WYMAGANE
                "DEMO_URL_B": "...",         # WYMAGANE
                "B2B_QUANTITY": "...",       # OPCJONALNE (jeśli puste → fallback)
                "B2B_PRODUCT_PL": "...",     # OPCJONALNE
                "B2B_BUYERS_NOM": "...",     # OPCJONALNE
                "B2B_BUYERS_GEN": "..."      # OPCJONALNE
            },
            "deck_date": "Kwiecień 2026"  # opcjonalne (override auto)
        }

    Headers:
        X-API-Key: secret  (wymagane jeśli RENDER_API_KEY ustawione)

    Response: PDF binary (Content-Type: application/pdf)
    """
    # Autoryzacja
    if API_KEY:
        provided_key = request.headers.get("X-API-Key", "")
        if provided_key != API_KEY:
            return jsonify({"error": "Unauthorized — invalid X-API-Key"}), 401

    # Parse JSON
    payload = request.get_json(silent=True)
    if not payload:
        return jsonify({"error": "Invalid JSON body"}), 400

    language = str(payload.get("language", "pl")).lower().strip()
    client_values = payload.get("values", {}) or {}
    deck_date_override = payload.get("deck_date")

    # Złóż finalny zestaw wartości
    final_values = dict(DEFAULT_VALUES)            # ceny default
    final_values.update(client_values)             # klient nadpisuje (jeśli chce)
    final_values["DECK_DATE"] = (
        str(deck_date_override) if deck_date_override else get_deck_date(language)
    )

    # Wczytaj template
    try:
        template_path = get_template_path(language)
        html = template_path.read_text(encoding="utf-8")
    except Exception as e:
        return jsonify({"error": f"Template error: {e}"}), 500

    # Apply B2B fallback (jeśli pola B2B puste → podmień bloki HTML)
    html, final_values = apply_b2b_fallback(html, final_values, language)

    # Podstaw placeholdery
    try:
        filled_html = fill_template(html, final_values)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    # Render PDF
    output_dir = Path(tempfile.gettempdir()) / "sitepatron-pdfs"
    output_dir.mkdir(exist_ok=True)
    pdf_path = output_dir / f"deck-{language}-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}.pdf"

    try:
        render_html_to_pdf(filled_html, pdf_path)
    except Exception as e:
        return jsonify({"error": f"Render failed: {e}"}), 500

    download_name = f"sitepatron-deck-{language}.pdf"
    return send_file(
        str(pdf_path),
        mimetype="application/pdf",
        as_attachment=False,
        download_name=download_name,
    )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print(f"Starting SitePatron Render Service v1.1 on port {PORT}")
    print(f"Templates dir: {TEMPLATES_DIR}")
    print(f"API key required: {bool(API_KEY)}")
    print(f"Available templates: {[f.name for f in TEMPLATES_DIR.glob('*.html')]}")
    print(f"B2B fallback languages: {sorted(B2B_FALLBACK.keys())}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
