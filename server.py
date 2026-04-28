"""
SitePatron Deck Render Service v1.3.2
======================================

Flask endpoint który renderuje Sosenco-style PDF z HTML template'ów.
Apps Script wysyła POST z wartościami per-klient + językiem,
serwer zwraca PDF jako binary LUB wgrywa do Drive (jeśli klient
prześle drive_folder_id).

ZMIANY v1.3.2:
- Klient moze przeslac `pdf_filename` w payload (bez .pdf na koncu lub z).
  Endpoint uzyje go jako nazwe pliku w Drive. Jesli nie podano,
  fallback na auto-generowana nazwe.

ZMIANY v1.3.1:
- Drive upload error: zwraca JSON z traceback zamiast cichego fallback.
- Logging do stderr (gunicorn poprawnie loguje).

ZMIANY v1.3:
- Drive Upload Mode: jesli klient przesle drive_folder_id, endpoint
  wgrywa PDF do Drive przez service account, zwraca JSON z file_id.

Endpoints:
    GET  /health          — status + lista dostępnych template'ów
    POST /render          — render PDF (wymaga X-API-Key)
"""

import os
import re
import sys
import json
import logging
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from playwright.sync_api import sync_playwright

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    DRIVE_LIBS_AVAILABLE = True
except ImportError:
    DRIVE_LIBS_AVAILABLE = False

app = Flask(__name__)

logging.basicConfig(
    stream=sys.stderr,
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
)
log = logging.getLogger("sitepatron")

# ============================================================
# KONFIGURACJA
# ============================================================

API_KEY = os.environ.get("RENDER_API_KEY", "")
TEMPLATES_DIR = Path(__file__).parent / "templates"
PORT = int(os.environ.get("PORT", 8000))
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

DEFAULT_VALUES = {
    "PRICE_MONTHLY_USD": "59",
    "PRICE_YEARLY_USD": "499",
    "PRICE_MIN_3M_USD": "177",
    "PRICE_SAVINGS_USD": "209",
    "PRICE_LOCAL_APPROX": "240",
}

MONTHS = {
    "pl": ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
           "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"],
    "en": ["January", "February", "March", "April", "May", "June",
           "July", "August", "September", "October", "November", "December"],
    "de": ["Januar", "Februar", "März", "April", "Mai", "Juni",
           "Juli", "August", "September", "Oktober", "November", "Dezember"],
}

# ============================================================
# B2B FALLBACK
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
# DRIVE UPLOAD
# ============================================================

def get_drive_service():
    if not DRIVE_LIBS_AVAILABLE:
        return None
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        return None

    try:
        creds_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    except json.JSONDecodeError as e:
        log.error(f"GOOGLE_SERVICE_ACCOUNT_JSON nie jest poprawnym JSONem: {e}")
        return None

    try:
        credentials = service_account.Credentials.from_service_account_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        return build("drive", "v3", credentials=credentials, cache_discovery=False)
    except Exception as e:
        log.error(f"nie mogę zbudować Drive service: {e}")
        return None


def get_service_account_email() -> str:
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        return ""
    try:
        creds_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
        return creds_info.get("client_email", "")
    except Exception:
        return ""


def upload_pdf_to_drive(pdf_path: Path, folder_id: str, name: str) -> dict:
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

    log.info(f"Drive upload: file={name}, folder_id={folder_id}, size={pdf_path.stat().st_size}")
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name, webViewLink, webContentLink",
        supportsAllDrives=True,
    ).execute()
    log.info(f"Drive upload OK: file_id={file.get('id')}")

    return {
        "file_id": file.get("id"),
        "name": file.get("name"),
        "view_url": file.get("webViewLink"),
        "download_url": file.get("webContentLink"),
    }


def sanitize_filename(name: str) -> str:
    """Sanityzuj nazwe pliku - usun znaki ktore moga sprawic problem w Drive/OS."""
    if not name:
        return "untitled"
    # Usun control chars i znaki ktore Drive moze odrzucic
    name = re.sub(r'[\x00-\x1f\x7f<>:"/\\|?*]', '_', name)
    name = name.strip(' .')  # Drive nie lubi plikow ktore zaczynaja/koncza sie kropka lub spacja
    if not name:
        return "untitled"
    # Limit dlugosci
    if len(name) > 200:
        # Zachowaj rozszerzenie
        if name.lower().endswith('.pdf'):
            name = name[:196] + '.pdf'
        else:
            name = name[:200]
    return name


# ============================================================
# RENDERING
# ============================================================

def get_deck_date(language: str) -> str:
    now = datetime.now()
    months = MONTHS.get(language.lower(), MONTHS["en"])
    return f"{months[now.month - 1]} {now.year}"


def get_template_path(language: str) -> Path:
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
                log.warning(f"document.fonts wait failed: {e}")

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
    if not url:
        return url
    url = url.strip()
    if url.startswith("http://") or url.startswith("https://"):
        return url
    return "https://" + url


def apply_b2b_fallback(html: str, values: dict, language: str) -> tuple:
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
    templates = sorted([f.name for f in TEMPLATES_DIR.glob("*.html")])
    drive_service = get_drive_service()
    return jsonify({
        "status": "ok",
        "version": "1.3.2",
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
    custom_pdf_filename = str(payload.get("pdf_filename", "")).strip()

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
    pdf_path = output_dir / f"deck-{language}-{timestamp}.pdf"

    try:
        render_html_to_pdf(filled_html, pdf_path)
    except Exception as e:
        return jsonify({"error": f"Render failed: {e}"}), 500

    # === TRYB B: Upload do Drive ===
    drive_service = get_drive_service() if drive_folder_id else None
    if drive_folder_id and drive_service is not None:
        # Nazwa pliku w Drive: priorytet (1) klient, (2) auto
        if custom_pdf_filename:
            drive_filename = custom_pdf_filename
            if not drive_filename.lower().endswith(".pdf"):
                drive_filename += ".pdf"
        else:
            drive_filename = f"sitepatron-{language}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"

        drive_filename = sanitize_filename(drive_filename)

        try:
            file_info = upload_pdf_to_drive(pdf_path, drive_folder_id, drive_filename)
        except Exception as e:
            error_detail = f"{type(e).__name__}: {str(e)}"
            tb_str = traceback.format_exc()
            log.error(f"Drive upload failed: {error_detail}\n{tb_str}")
            try:
                pdf_path.unlink()
            except Exception:
                pass
            return jsonify({
                "error": "Drive upload failed",
                "detail": error_detail,
                "traceback": tb_str[-3000:] if len(tb_str) > 3000 else tb_str,
                "drive_folder_id_used": drive_folder_id,
                "service_account_email": get_service_account_email(),
            }), 500

        try:
            pdf_path.unlink()
        except Exception:
            pass
        return jsonify(file_info)

    # === TRYB A: Zwróć PDF binary ===
    download_name = custom_pdf_filename or f"sitepatron-{language}.pdf"
    if not download_name.lower().endswith(".pdf"):
        download_name += ".pdf"
    download_name = sanitize_filename(download_name)
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
    log.info(f"Starting SitePatron Render Service v1.3.2 on port {PORT}")
    log.info(f"Templates dir: {TEMPLATES_DIR}")
    log.info(f"API key required: {bool(API_KEY)}")
    log.info(f"Available templates: {[f.name for f in TEMPLATES_DIR.glob('*.html')]}")
    log.info(f"B2B fallback languages: {sorted(B2B_FALLBACK.keys())}")
    log.info(f"Drive libs available: {DRIVE_LIBS_AVAILABLE}")
    drive_svc = get_drive_service()
    log.info(f"Drive upload enabled: {drive_svc is not None}")
    if drive_svc is not None:
        log.info(f"Service account: {get_service_account_email()}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
