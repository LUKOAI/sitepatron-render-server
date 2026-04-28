"""
SitePatron Deck Render Service v1.3
====================================

Flask endpoint który renderuje Sosenco-style PDF z HTML template'ów.
Apps Script wysyła POST z wartościami per-klient + językiem,
serwer zwraca PDF jako binary LUB wgrywa do Drive (jeśli klient
prześle drive_folder_id).

Endpoints:
    GET  /health          — status + lista dostępnych template'ów
    POST /render          — render PDF (wymaga X-API-Key)

ZMIANY v1.3:
- Drive Upload Mode: jeśli klient prześle `drive_folder_id` w request
  + endpoint ma skonfigurowane GOOGLE_SERVICE_ACCOUNT_JSON,
  to PDF jest wgrywany bezpośrednio do Drive folderu klienta
  (przez service account), a endpoint zwraca JSON z file_id zamiast
  PDF binary. Powód: Apps Script ma dzienny limit 50MB UrlFetch
  bandwidth — wysyłanie 1.6MB PDF szybko go wyczerpuje.
- Backward compat: jeśli klient nie prześle drive_folder_id,
  endpoint zwraca PDF binary jak w v1.2.

ZMIANY v1.2:
- Fix polskich znaków diakrytycznych (ó, ą, ę): czekamy na
  document.fonts.ready PLUS 5s timeout zamiast 2s.
- Normalizacja URL DEMO_URL_A/DEMO_URL_B: jesli brak schemy http(s),
  dodajemy https:// automatycznie.

ZMIANY v1.1:
- Fallback B2B: jeśli wszystkie 4 pola B2B są puste/brakujące →
  endpoint podmienia 2 fragmenty HTML na uniwersalny tekst z
  kalkulatorem oszczędności.

Deploy:
    docker build -t sitepatron-render .
    docker run -p 8000:8000 \\
        -e RENDER_API_KEY=secret \\
        -e GOOGLE_SERVICE_ACCOUNT_JSON='{...}' \\
        sitepatron-render

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

# Drive upload (opcjonalny - tylko gdy GOOGLE_SERVICE_ACCOUNT_JSON jest ustawione)
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    DRIVE_LIBS_AVAILABLE = True
except ImportError:
    DRIVE_LIBS_AVAILABLE = False

app = Flask(__name__)

# ============================================================
# KONFIGURACJA
# ============================================================

API_KEY = os.environ.get("RENDER_API_KEY", "")
TEMPLATES_DIR = Path(__file__).parent / "templates"
PORT = int(os.environ.get("PORT", 8000))

# Service account JSON do uploadu plikow do Drive (opcjonalne)
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

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
}


# ============================================================
# DRIVE UPLOAD (przez service account)
# ============================================================

def get_drive_service():
    """
    Build Google Drive service from service account credentials in env var.
    Zwraca None jesli zmienna srodowiskowa nie jest ustawiona lub jest niepoprawna.
    """
    if not DRIVE_LIBS_AVAILABLE:
        return None
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        return None

    try:
        creds_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    except json.JSONDecodeError as e:
        print(f"ERROR: GOOGLE_SERVICE_ACCOUNT_JSON nie jest poprawnym JSONem: {e}")
        return None

    try:
        credentials = service_account.Credentials.from_service_account_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        return build("drive", "v3", credentials=credentials, cache_discovery=False)
    except Exception as e:
        print(f"ERROR: nie mogę zbudować Drive service: {e}")
        return None


def get_service_account_email() -> str:
    """Wyciagnij email service account z JSON env var (do /health endpoint)."""
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        return ""
    try:
        creds_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
        return creds_info.get("client_email", "")
    except Exception:
        return ""


def upload_pdf_to_drive(pdf_path: Path, folder_id: str, name: str) -> dict:
    """
    Upload PDF do Drive folderu i zwroc metadata.
    Wymaga zeby service account mial Editor access do folderu.
    """
    service = get_drive_service()
    if service is None:
        raise RuntimeError(
            "Drive service unavailable - missing or invalid GOOGLE_SERVICE_ACCOUNT_JSON"
        )

    file_metadata = {
        "name": name,
        "parents": [folder_id],
    }
    media = MediaFileUpload(
        str(pdf_path),
        mimetype="application/pdf",
        resumable=False,
    )

    try:
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, webViewLink, webContentLink",
            supportsAllDrives=True,
        ).execute()
    except Exception as e:
        raise RuntimeError(f"Drive API upload failed: {e}")

    return {
        "file_id": file.get("id"),
        "name": file.get("name"),
        "view_url": file.get("webViewLink"),
        "download_url": file.get("webContentLink"),
    }


# ============================================================
# RENDERING
# ============================================================

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
    """Render HTML → PDF przez Playwright Chromium (headless).

    Czeka na pełne załadowanie fontów (document.fonts.ready) przed
    renderowaniem PDF — bez tego polskie znaki (ó, ą, ę) gubią się
    w PDF, a layout się rozjeżdża.
    """
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
            try:
                page.evaluate("() => document.fonts.ready")
                page.wait_for_function(
                    "() => document.fonts.status === 'loaded'",
                    timeout=10000,
                )
            except Exception as e:
                print(f"Warning: document.fonts wait failed: {e}")

            page.wait_for_timeout(5000)

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


def normalize_url(url: str) -> str:
    """Jeśli URL nie zaczyna się od http:// ani https:// → dodaj https://."""
    if not url:
        return url
    url = url.strip()
    if url.startswith("http://") or url.startswith("https://"):
        return url
    return "https://" + url


def apply_b2b_fallback(html: str, values: dict, language: str) -> tuple:
    """B2B fallback - zob. v1.1."""
    b2b_filled = {f: str(values.get(f, "")).strip() for f in B2B_FIELDS}
    all_filled = all(v != "" for v in b2b_filled.values())

    new_values = dict(values)

    if not all_filled:
        fallback = B2B_FALLBACK.get(language, B2B_FALLBACK["pl"])

        if fallback["bullet_old"] in html:
            html = html.replace(fallback["bullet_old"], fallback["bullet_new"])
        if fallback["example_old"] in html:
            html = html.replace(fallback["example_old"], fallback["example_new"])

        for f in B2B_FIELDS:
            new_values[f] = ""

    return html, new_values


def fill_template(html: str, values: dict) -> str:
    """Podstaw {{KEY}} → wartości."""
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
    """Status check + lista dostępnych template'ów + status Drive upload."""
    templates = sorted([f.name for f in TEMPLATES_DIR.glob("*.html")])
    drive_service = get_drive_service()
    return jsonify({
        "status": "ok",
        "version": "1.3",
        "templates": templates,
        "api_key_required": bool(API_KEY),
        "playwright": "ready",
        "b2b_fallback_languages": sorted(B2B_FALLBACK.keys()),
        "drive_upload_enabled": drive_service is not None,
        "drive_service_account_email": get_service_account_email(),
        "drive_libs_available": DRIVE_LIBS_AVAILABLE,
    })


@app.route("/render", methods=["POST"])
def render():
    """
    Render PDF z HTML template'u.

    Body JSON:
        {
            "language": "pl",
            "values": {
                "DEMO_URL_A": "...",         # WYMAGANE
                "DEMO_URL_B": "...",         # WYMAGANE
                "B2B_QUANTITY": "...",       # OPCJONALNE
                "B2B_PRODUCT_PL": "...",     # OPCJONALNE
                "B2B_BUYERS_NOM": "...",     # OPCJONALNE
                "B2B_BUYERS_GEN": "..."      # OPCJONALNE
            },
            "deck_date": "Kwiecień 2026",   # opcjonalne
            "drive_folder_id": "1AbC..."    # OPCJONALNE - wgraj do Drive
        }

    Headers:
        X-API-Key: secret  (wymagane jeśli RENDER_API_KEY ustawione)

    Response:
        - PDF binary (Content-Type: application/pdf) — jesli brak drive_folder_id
        - JSON {file_id, name, view_url, download_url} — jesli drive_folder_id podane
    """
    if API_KEY:
        provided_key = request.headers.get("X-API-Key", "")
        if provided_key != API_KEY:
            return jsonify({"error": "Unauthorized — invalid X-API-Key"}), 401

    payload = request.get_json(silent=True)
    if not payload:
        return jsonify({"error": "Invalid JSON body"}), 400

    language = str(payload.get("language", "pl")).lower().strip()
    client_values = payload.get("values", {}) or {}
    deck_date_override = payload.get("deck_date")
    drive_folder_id = payload.get("drive_folder_id", "")

    final_values = dict(DEFAULT_VALUES)
    final_values.update(client_values)
    final_values["DECK_DATE"] = (
        str(deck_date_override) if deck_date_override else get_deck_date(language)
    )

    for url_key in ("DEMO_URL_A", "DEMO_URL_B"):
        if url_key in final_values and final_values[url_key]:
            final_values[url_key] = normalize_url(str(final_values[url_key]))

    try:
        template_path = get_template_path(language)
        html = template_path.read_text(encoding="utf-8")
    except Exception as e:
        return jsonify({"error": f"Template error: {e}"}), 500

    html, final_values = apply_b2b_fallback(html, final_values, language)

    try:
        filled_html = fill_template(html, final_values)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    output_dir = Path(tempfile.gettempdir()) / "sitepatron-pdfs"
    output_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    pdf_filename = f"deck-{language}-{timestamp}.pdf"
    pdf_path = output_dir / pdf_filename

    try:
        render_html_to_pdf(filled_html, pdf_path)
    except Exception as e:
        return jsonify({"error": f"Render failed: {e}"}), 500

    # === TRYB B: Upload do Drive (jesli klient prosi I Drive jest skonfigurowany) ===
    drive_service = get_drive_service() if drive_folder_id else None
    if drive_folder_id and drive_service is not None:
        drive_filename = f"sitepatron-deck-{language}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
        try:
            file_info = upload_pdf_to_drive(pdf_path, drive_folder_id, drive_filename)
        except Exception as e:
            # Jesli upload padl, zwracamy PDF binary jako fallback
            print(f"WARNING: Drive upload failed, falling back to PDF binary: {e}")
        else:
            # Upload OK — sprzatamy lokalny plik, zwracamy JSON
            try:
                pdf_path.unlink()
            except Exception:
                pass
            return jsonify(file_info)

    # === TRYB A: Zwróć PDF binary (legacy / fallback gdy Drive niedostepny) ===
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
    print(f"Starting SitePatron Render Service v1.3 on port {PORT}")
    print(f"Templates dir: {TEMPLATES_DIR}")
    print(f"API key required: {bool(API_KEY)}")
    print(f"Available templates: {[f.name for f in TEMPLATES_DIR.glob('*.html')]}")
    print(f"B2B fallback languages: {sorted(B2B_FALLBACK.keys())}")
    print(f"Drive libs available: {DRIVE_LIBS_AVAILABLE}")
    drive_svc = get_drive_service()
    print(f"Drive upload enabled: {drive_svc is not None}")
    if drive_svc is not None:
        print(f"Service account: {get_service_account_email()}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
