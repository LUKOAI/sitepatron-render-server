"""
SitePatron Deck Render Service
==============================

Flask endpoint który renderuje Sosenco-style PDF z HTML template'ów.
Apps Script wysyła POST z 6 wartościami per-klient + językiem,
serwer zwraca PDF jako binary.

Endpoints:
    GET  /health          — status + lista dostępnych template'ów
    POST /render          — render PDF (wymaga X-API-Key)

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
        "templates": templates,
        "api_key_required": bool(API_KEY),
        "playwright": "ready",
    })


@app.route("/render", methods=["POST"])
def render():
    """
    Render PDF z HTML template'u.

    Body JSON:
        {
            "language": "pl",          # pl/en/de (fallback PL)
            "values": {                # per-klient (6 placeholderów)
                "DEMO_URL_A": "...",
                "DEMO_URL_B": "...",
                "B2B_QUANTITY": "...",
                "B2B_PRODUCT_PL": "...",
                "B2B_BUYERS_NOM": "...",
                "B2B_BUYERS_GEN": "..."
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
    print(f"Starting SitePatron Render Service on port {PORT}")
    print(f"Templates dir: {TEMPLATES_DIR}")
    print(f"API key required: {bool(API_KEY)}")
    print(f"Available templates: {[f.name for f in TEMPLATES_DIR.glob('*.html')]}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
