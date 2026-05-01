"""
SitePatron Deck Render Service v1.3.4
======================================

Flask endpoint który renderuje Sosenco-style PDF z HTML template'ów.
Apps Script wysyła POST z wartościami per-klient + językiem,
serwer zwraca PDF jako binary LUB wgrywa do Drive (jeśli klient
prześle drive_folder_id).

ZMIANY v1.3.4:
- FIX FALLBACK: gdy brak template dla danego języka, fallback to teraz
  EN (a nie PL jak wcześniej). PL pozostaje jako ostateczne fallback
  (gdyby EN też nie istniał - praktycznie niemożliwe).
- B2B_FALLBACK: zmieniony default z "pl" na "en" (linia w apply_b2b_fallback).
- MONTHS: zmieniony default z "en" na "en" (już było OK, bez zmian).
- Wersja bumped na 1.3.4 w /health.

ZMIANY v1.3.3:
- B2B_FALLBACK rozszerzony o klucze "en" i "de" (do tej pory tylko "pl").
  Teraz wszystkie 3 jezyki maja pelny fallback gdy fields B2B w arkuszu sa puste.

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
    "fr": ["janvier", "février", "mars", "avril", "mai", "juin",
           "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
    "es": ["enero", "febrero", "marzo", "abril", "mayo", "junio",
           "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
    "it": ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno",
           "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"],
    "nl": ["januari", "februari", "maart", "april", "mei", "juni",
           "juli", "augustus", "september", "oktober", "november", "december"],
    "sv": ["januari", "februari", "mars", "april", "maj", "juni",
           "juli", "augusti", "september", "oktober", "november", "december"],
    "da": ["januar", "februar", "marts", "april", "maj", "juni",
           "juli", "august", "september", "oktober", "november", "december"],
    "no": ["januar", "februar", "mars", "april", "mai", "juni",
           "juli", "august", "september", "oktober", "november", "desember"],
    "fi": ["tammikuu", "helmikuu", "maaliskuu", "huhtikuu", "toukokuu", "kesäkuu",
           "heinäkuu", "elokuu", "syyskuu", "lokakuu", "marraskuu", "joulukuu"],
    "et": ["jaanuar", "veebruar", "märts", "aprill", "mai", "juuni",
           "juuli", "august", "september", "oktoober", "november", "detsember"],
    "lt": ["sausis", "vasaris", "kovas", "balandis", "gegužė", "birželis",
           "liepa", "rugpjūtis", "rugsėjis", "spalis", "lapkritis", "gruodis"],
    "lv": ["janvāris", "februāris", "marts", "aprīlis", "maijs", "jūnijs",
           "jūlijs", "augusts", "septembris", "oktobris", "novembris", "decembris"],
    "cs": ["leden", "únor", "březen", "duben", "květen", "červen",
           "červenec", "srpen", "září", "říjen", "listopad", "prosinec"],
    "sk": ["január", "február", "marec", "apríl", "máj", "jún",
           "júl", "august", "september", "október", "november", "december"],
    "sl": ["januar", "februar", "marec", "april", "maj", "junij",
           "julij", "avgust", "september", "oktober", "november", "december"],
    "hr": ["siječanj", "veljača", "ožujak", "travanj", "svibanj", "lipanj",
           "srpanj", "kolovoz", "rujan", "listopad", "studeni", "prosinac"],
    "hu": ["január", "február", "március", "április", "május", "június",
           "július", "augusztus", "szeptember", "október", "november", "december"],
    "ro": ["ianuarie", "februarie", "martie", "aprilie", "mai", "iunie",
           "iulie", "august", "septembrie", "octombrie", "noiembrie", "decembrie"],
    "bg": ["януари", "февруари", "март", "април", "май", "юни",
           "юли", "август", "септември", "октомври", "ноември", "декември"],
    "el": ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος",
           "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος"],
    "tr": ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
           "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"],
    "pt": ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
           "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
    "ar": ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
           "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"],
    "ja": ["1月", "2月", "3月", "4月", "5月", "6月",
           "7月", "8月", "9月", "10月", "11月", "12月"],
    "ko": ["1월", "2월", "3월", "4월", "5월", "6월",
           "7월", "8월", "9월", "10월", "11월", "12월"],
    "zh": ["1月", "2月", "3月", "4月", "5月", "6月",
           "7月", "8月", "9月", "10月", "11月", "12月"],
}

# ============================================================
# B2B FALLBACK
#
# bullet_old/example_old MUSZA dokladnie pasowac do tekstu w odpowiednim
# template HTML (sitepatron-deck-{LANG}-template.html). Jesli zmienisz tekst
# w template, zaktualizuj tutaj odpowiednio.
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
    "en": {
        "bullet_old": (
            'Wholesale customers ({{B2B_BUYERS_NOM}}) contact you directly'
        ),
        "bullet_new": (
            'Wholesale customers (housing associations, hotels, offices, companies '
            'with multiple branches) contact you directly'
        ),
        "example_old": (
            'A customer orders a set of {{B2B_QUANTITY}} {{B2B_PRODUCT_PL}} (e.g. for '
            '{{B2B_BUYERS_GEN}}). <strong>On Amazon:</strong> they buy unit by unit, '
            'each subject to Amazon\'s 15% commission + platform restrictions. '
            '<strong>Directly from the Patron via the Site:</strong> one B2B invoice '
            'between companies, no commission, room to negotiate price for larger '
            'volumes. Pure profit + a new business contact in your portfolio.'
        ),
        "example_new": (
            'A business customer orders directly through the Site instead of through '
            'Amazon (e.g. housing association, office, hotel, company with branches). '
            '<strong>On Amazon:</strong> each unit subject to ~15% commission + '
            'platform restrictions. <strong>Directly from the Patron:</strong> one '
            'B2B invoice between companies, no commission, room to negotiate price. '
            '<strong>The numbers:</strong> a 400 EUR order = ~60 EUR savings on '
            'Amazon commission — that covers a whole month of Site Patron. A 1,000 '
            'EUR order = ~150 EUR savings = 2.5 months of subscription from a single '
            'order. Plus a permanent B2B customer in your portfolio.'
        ),
    },
    "de": {
        "bullet_old": (
            'Großkunden ({{B2B_BUYERS_NOM}}) wenden sich direkt an Sie'
        ),
        "bullet_new": (
            'Großkunden (Wohnungseigentümergemeinschaften, Hotels, Büros, Unternehmen '
            'mit mehreren Niederlassungen) wenden sich direkt an Sie'
        ),
        "example_old": (
            'Ein Kunde bestellt einen Satz {{B2B_QUANTITY}} {{B2B_PRODUCT_PL}} (z.B. '
            'für {{B2B_BUYERS_GEN}}). <strong>Auf Amazon:</strong> er kauft Stück für '
            'Stück, jedes mit 15% Amazon-Provision + Plattformeinschränkungen belastet. '
            '<strong>Direkt vom Patron über die Website:</strong> eine B2B-Rechnung '
            'zwischen Unternehmen, ohne Provision, Verhandlungsspielraum beim Preis '
            'für größere Mengen. Reiner Gewinn + ein neuer Geschäftskontakt in Ihrem '
            'Portfolio.'
        ),
        "example_new": (
            'Ein Geschäftskunde bestellt direkt über die Website statt über Amazon '
            '(z.B. Wohnungseigentümergemeinschaft, Büro, Hotel, Unternehmen mit '
            'mehreren Niederlassungen). <strong>Auf Amazon:</strong> jedes Stück mit '
            '~15% Provision belastet + Plattformeinschränkungen. <strong>Direkt vom '
            'Patron:</strong> eine B2B-Rechnung zwischen Unternehmen, ohne Provision, '
            'Verhandlungsspielraum beim Preis. <strong>Die Zahlen:</strong> eine '
            'Bestellung über 400 EUR = ~60 EUR Ersparnis bei der Amazon-Provision — '
            'das deckt einen ganzen Monat Site Patron ab. Eine Bestellung über 1.000 '
            'EUR = ~150 EUR Ersparnis = 2,5 Monate Abonnement aus einer einzigen '
            'Bestellung. Plus ein dauerhafter B2B-Kunde in Ihrem Portfolio.'
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
    """
    Find HTML template for the given language.

    Search order:
      1. Exact match: sitepatron-deck-{LANG}-template.html
      2. Fallback: EN template (was PL in v1.3.3 and earlier — fixed in v1.3.4)
      3. Last resort: PL template (kept for safety)
      4. Any HTML file in templates/
    """
    candidates = [
        TEMPLATES_DIR / f"sitepatron-deck-{language.upper()}-template.html",
        TEMPLATES_DIR / "sitepatron-deck-EN-template.html",   # ← v1.3.4: EN przed PL
        TEMPLATES_DIR / "sitepatron-deck-PL-template.html",   # ← ostatecznie PL
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
    """
    Apply B2B field fallback when fields are empty.

    v1.3.4: default fallback is now EN (was PL). This affects languages outside
    PL/EN/DE — they will get English B2B fallback text.
    Note: the fallback strings won't match the HTML for non-PL/EN/DE languages
    anyway (each template has its own B2B text), so this only matters for the
    "did we substitute or not" logic. Substitution will be a no-op for those
    languages because the EN strings won't match their HTML.
    For full proper handling: each template should use B2B_BUYERS_NOM placeholder
    consistently, and apply_b2b_fallback should match generic patterns.
    """
    b2b_filled = {f: str(values.get(f, "")).strip() for f in B2B_FIELDS}
    all_filled = all(v != "" for v in b2b_filled.values())

    new_values = dict(values)

    if not all_filled:
        # v1.3.4: fallback EN zamiast PL
        fallback = B2B_FALLBACK.get(language, B2B_FALLBACK["en"])

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
        "version": "1.3.4",
        "templates": templates,
        "api_key_required": bool(API_KEY),
        "playwright": "ready",
        "b2b_fallback_languages": sorted(B2B_FALLBACK.keys()),
        "months_languages": sorted(MONTHS.keys()),
        "fallback_chain": "EN -> PL (v1.3.4)",
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
        log.info(f"Render: language={language}, template={template_path.name}")
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
    log.info(f"Starting SitePatron Render Service v1.3.4 on port {PORT}")
    log.info(f"Templates dir: {TEMPLATES_DIR}")
    log.info(f"API key required: {bool(API_KEY)}")
    log.info(f"Available templates: {[f.name for f in TEMPLATES_DIR.glob('*.html')]}")
    log.info(f"B2B fallback languages: {sorted(B2B_FALLBACK.keys())}")
    log.info(f"MONTHS languages: {sorted(MONTHS.keys())}")
    log.info(f"Fallback chain: EN -> PL")
    log.info(f"Drive libs available: {DRIVE_LIBS_AVAILABLE}")
    drive_svc = get_drive_service()
    log.info(f"Drive upload enabled: {drive_svc is not None}")
    if drive_svc is not None:
        log.info(f"Service account: {get_service_account_email()}")
    app.run(host="0.0.0.0", port=PORT, debug=False)
