/**********************************************************
 *  SitePatron Email Automation v3.1
 *  System wysylki emaili z auto-detekcja jezyka
 *
 *  Zmiany vs v3.0:
 *  - Send_ColdPitch USUNIETY z BATCH (cold pitch wysylany TYLKO manualnie
 *    z menu Pitch + PDF, gdzie jest generowany PDF jako zalacznik)
 *  - Menu "Batch wysylka -> Cold Pitch -> wszyscy New" USUNIETE
 *  - firstSetupV30 NIE tworzy kolumn Send_ColdPitch / TS_ColdPitch
 *    (te zarzadza pitch_addon.gs v2.12 wspolnie z PDFLink/DraftLink)
 *  - migrateV2toV30: Send_Email1 -> Send_ColdPitch (kolumna z pitch_addon)
 *
 *  Zmiany vs v2.0:
 *  - 4 checkboxy batch: Send_Pitch / Send_Landing1 / Send_Landing2 / Send_Purchase
 *    Pasuje do nowych szablonow w SP_EmailTemplates:
 *    SP_PITCH_LANG / SP_LANDING1_LANG / SP_LANDING2_LANG / SP_PURCHASE_LANG
 *  - Cold Pitch (SP_COLDPITCH_LANG) - tylko manualnie z pitch_addon.gs
 *  - Fallback URL: jesli LANDING_URL_DE nie ma w SP_Config,
 *    automatycznie uzyj LANDING_URL_EN. Analogicznie dla
 *    CHECKOUT_URL_*. Klient nigdy nie zobaczy {{LANDING_URL_DE}}
 *    w mailu.
 *  - STOP w menu: 1 klik anuluje wszystkie pending z kolejki,
 *    odznacza checkbox-y, loguje CANCELLED_BY_USER.
 *  - Czyste menu z separatorami i grupami.
 *  - Reload Config bez restartu skryptu.
 *  - Migracja v2 -> v3: kopiuje stare kolumny do nowych.
 *
 *  Zachowane vs v2.0:
 *  - Wysylka z aliasu info@netanaliza.com
 *  - Re-check checkbox PRZED wysylka (bugfix v2)
 *  - Country -> Language mapping (50+ krajow)
 *  - Pitch + PDF Addon (osobny plik pitch_addon.gs v2.12)
 *
 *  Wersja: 3.1 | 2026-04-30 | NetAnaliza
 *
 *  ZAKLADKI:
 *  - SP_Sellers: dane sellerow, checkboxy, statusy
 *  - SP_EmailTemplates: szablony w wielu jezykach
 *  - SP_LOG: protokol
 *  - SP_Config: parametry
 **********************************************************/

/* =================== KONFIGURACJA =================== */
const SHEET_SELLERS    = 'SP_Sellers';
const SHEET_TEMPLATES  = 'SP_EmailTemplates';
const SHEET_LOG        = 'SP_LOG';
const SHEET_CONFIG     = 'SP_Config';

let CONFIG_CACHE = null;
const PROP_EMAIL_QUEUE = 'SP_EMAIL_QUEUE';
const EMAIL_DELAY_MS_DEFAULT = 120000;  // 2 minuty
const TZ_DEFAULT = 'Europe/Berlin';

// ===== NADAWCA =====
const DEFAULT_FROM_EMAIL = 'info@netanaliza.com';
const DEFAULT_FROM_NAME  = 'Lukasz from NetAnaliza';
const DEFAULT_REPLY_TO   = 'info@netanaliza.com';

// ===== MAPY KOLUMN -> TYP/TS/STATUS (TYLKO BATCH) =====
// Cold Pitch nie jest tutaj - obsluguje go pitch_addon.gs (manual + PDF)
const COLUMN_TO_TYPE = {
  'Send_Pitch':      'PITCH',
  'Send_Landing1':   'LANDING1',
  'Send_Landing2':   'LANDING2',
  'Send_Purchase':   'PURCHASE'
};

const COLUMN_TO_TS = {
  'Send_Pitch':      'TS_Pitch',
  'Send_Landing1':   'TS_Landing1',
  'Send_Landing2':   'TS_Landing2',
  'Send_Purchase':   'TS_Purchase'
};

const COLUMN_TO_STATUS = {
  'Send_Pitch':      'Pitch_Sent',
  'Send_Landing1':   'Landing1_Sent',
  'Send_Landing2':   'Landing2_Sent',
  'Send_Purchase':   'Purchase_Sent'
};

const VALID_SEND_COLS = Object.keys(COLUMN_TO_TYPE);

const STATUS_ORDER = [
  'New','Researched',
  'ColdPitch_Sent','Pitch_Sent','Lead',
  'Landing1_Sent','Landing2_Sent',
  'Interested','Registered',
  'Purchase_Sent','Purchased',
  'Declined','DoNotContact'
];

function getConfig_() {
  if (CONFIG_CACHE) return CONFIG_CACHE;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) {
    cfg[String(data[i][0]).trim()] = String(data[i][1]).trim();
  }
  CONFIG_CACHE = cfg;
  return cfg;
}

function getConfigValue_(key, fallback) {
  return getConfig_()[key] || fallback;
}

/* =================== COUNTRY -> LANGUAGE =================== */
const COUNTRY_TO_LANG = {
  // ===== Niemiecki =====
  'DE':'de','AT':'de','CH':'de','LI':'de','LU':'de',
  
  // ===== Angielski =====
  'GB':'en','UK':'en','US':'en','CA':'en','AU':'en','NZ':'en',
  'IE':'en','ZA':'en','SG':'en','IN':'en','PH':'en','MY':'en',
  'KE':'en','NG':'en','GH':'en','AE':'en',
  
  // ===== Polski =====
  'PL':'pl',
  
  // ===== Francuski =====
  'FR':'fr','BE':'fr','MC':'fr','SN':'fr','CI':'fr','MA':'fr','TN':'fr','DZ':'fr',
  
  // ===== Hiszpański =====
  'ES':'es','MX':'es','AR':'es','CO':'es','CL':'es','PE':'es',
  
  // ===== Włoski =====
  'IT':'it','SM':'it',
  
  // ===== Holenderski =====
  'NL':'nl',
  
  // ===== Skandynawia =====
  'SE':'sv','DK':'da','NO':'no','FI':'fi',
  
  // ===== Portugalski (UWAGA: BR przeniesiony nizej do pt-BR) =====
  'PT':'pt',
  
  // ===== Srodkowoeuropejskie =====
  'CZ':'cs','RO':'ro','HU':'hu','TR':'tr',
  
  // ===== CJK (UWAGA: TW i HK przeniesione nizej do zh-TW) =====
  'JP':'ja','KR':'ko','CN':'zh',
  
  // ===== Arabski =====
  'SA':'ar','EG':'ar','IQ':'ar','JO':'ar','KW':'ar','QA':'ar',
  
  // ===== Balkany / Grecja / Cypr =====
  'GR':'el','CY':'el','BG':'bg','HR':'hr','SK':'sk','SI':'sl',
  
  // ===== Baltyk =====
  'LT':'lt','LV':'lv','EE':'et',
  
  // ===== [PACZKA 8] Nowe jezyki =====
  'VN':'vi',                                 // Wietnam
  'ID':'id',                                 // Indonezja  
  'TH':'th',                                 // Tajlandia
  'TW':'zh-TW','HK':'zh-TW',                 // Tajwan + Hong Kong (traditional Chinese)
  'BR':'pt-BR',                              // Brazylia (NIE pt-PT)
  'RU':'ru','UA':'ru','BY':'ru','KZ':'ru'    // Rosja, Ukraina, Bialorus, Kazachstan
};

function getLanguageForRow_(headers, rowData) {
  // Priorytet 1: kolumna Language (reczne nadpisanie)
  const iLang = colIndex_(headers, 'Language');
  if (iLang > -1) {
    const lang = String(rowData[iLang] || '').trim().toLowerCase();
    if (lang && lang.length >= 2) {
      // [MINI-FIX 01.05.2026] Wyjatki dla kolizji kod_kraju == kod_jezyka
      // gdzie COUNTRY_TO_LANG mapuje na cos innego niz oczekiwane.
      // UWAGA: gdy dojdzie Paczka 7 (AR/JA/KO/ZH), tu beda kolejne wyjatki!
      if (lang === 'ar') return 'ar';  // arabski (NIE Argentyna->es)
      // [MINI-FIX] Jesli user wpisal kod KRAJU zamiast jezyka
      // (np. 'cz' zamiast 'cs', 'dk' zamiast 'da', 'gb' zamiast 'en')
      // - automatycznie mapuj na wlasciwy kod jezyka.
      const langUpper = lang.toUpperCase();
      if (COUNTRY_TO_LANG[langUpper]) {
        Logger.log('Auto-mapping Language="' + lang + '" -> ' + COUNTRY_TO_LANG[langUpper]);
        return COUNTRY_TO_LANG[langUpper];
      }
      return lang;
    }
  }
  // Priorytet 2: mapowanie Country -> Language
  const iCountry = colIndex_(headers, 'Country');
  if (iCountry > -1) {
    const country = String(rowData[iCountry] || '').trim().toUpperCase();
    if (country && COUNTRY_TO_LANG[country]) return COUNTRY_TO_LANG[country];
  }
  // Fallback: angielski
  return 'en';
}

/* =================== MENU =================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SitePatron')
    // ---- BATCH WYSYLKA (4 typy, BEZ Cold Pitch) ----
    .addSubMenu(ui.createMenu('Batch wysylka')
      .addItem('Pitch -> po Cold Pitch (X dni)', 'batchSend_Pitch_AfterDelay')
      .addItem('Landing 1 -> wszyscy Lead', 'batchSend_Landing1_Leads')
      .addItem('Landing 2 -> po Landing 1 (X dni)', 'batchSend_Landing2_AfterDelay'))
    .addSeparator()
    // ---- KONTROLA KOLEJKI ----
    .addItem('Sprawdz kolejke', 'showQueue')
    .addItem('STOP - anuluj wszystkie pending', 'stopAllInQueue')
    .addItem('Statystyki pipeline', 'showStats')
    .addSeparator()
    // ---- PITCH + PDF (drafty manual) - Cold Pitch i Pitch z PDF ----
    .addSubMenu(ui.createMenu('Pitch + PDF (drafty)')
      .addItem('Cold Pitch + PDF (zaznaczony wiersz)', 'generateColdPitchDraft')
      .addItem('Pitch po reakcji + PDF (zaznaczony wiersz)', 'generatePitchDraft')
      .addItem('Tylko PDF (do podgladu)', 'generatePdfOnlyForRow')
      .addSeparator()
      .addItem('Otworz folder Drive', 'openDriveFolder_v25'))
    .addSeparator()
    // ---- KONFIGURACJA ----
    .addSubMenu(ui.createMenu('Konfiguracja')
      .addItem('Pierwsza instalacja v3.1 (4 checkboxy batch)', 'firstSetupV30')
      .addItem('Reload Config (po zmianach SP_Config)', 'reloadConfig')
      .addItem('Instaluj triggery', 'installTriggers')
      .addItem('Zablokuj naglowki', 'protectHeaders')
      .addSeparator()
      .addItem('MIGRACJA v2 -> v3 (z 3 do 4 checkboxow)', 'migrateV2toV30')
      .addSeparator()
      .addItem('PDF Addon: instalacja v2.12 (Cold Pitch + Pitch)', 'firstSetupV25')
      .addItem('PDF Addon: konfiguracja render', 'firstSetupV26')
      .addItem('PDF Addon: MIGRACJA Pitsch -> ColdPitch', 'migratePitschToColdPitch')
      .addSeparator()
      .addItem('[STARE] Pierwsza instalacja v2', 'firstSetup'))
    .addToUi();
}

/* =================== FIRST SETUP (v2 - stary, zachowany) =================== */
function firstSetup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const required = [SHEET_SELLERS, SHEET_TEMPLATES, SHEET_LOG, SHEET_CONFIG];
  const missing = required.filter(name => !ss.getSheetByName(name));
  if (missing.length > 0) {
    ui.alert('Brak zakladek: ' + missing.join(', '));
    return;
  }
  const sellers = ss.getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sellers);
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
  ['Send_Email1','Send_Email2','Send_Reminder'].forEach(col => {
    const idx = colIndex_(headers, col);
    if (idx > -1) sellers.getRange(2, idx+1, 500, 1).setDataValidation(rule);
  });
  ui.alert('Setup v2 zakonczony! Nastepny krok: Instaluj triggery.\n\nUWAGA: ten setup uzywa starych kolumn (Email1/Email2/Reminder).\nDla nowego systemu uzyj: "Pierwsza instalacja v3.1".');
}

/* =================== FIRST SETUP V3.1 =================== */
// Tworzy 8 kolumn dla BATCH (4 typy, BEZ Cold Pitch).
// Cold Pitch zarzadza pitch_addon.gs v2.12 (firstSetupV25).
function firstSetupV30() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const required = [SHEET_SELLERS, SHEET_TEMPLATES, SHEET_LOG, SHEET_CONFIG];
  const missing = required.filter(name => !ss.getSheetByName(name));
  if (missing.length > 0) {
    ui.alert('Brak zakladek: ' + missing.join(', '));
    return;
  }

  const sellers = ss.getSheetByName(SHEET_SELLERS);
  let headers = getHeaders_(sellers);

  // Dodaj brakujace kolumny: 4 checkbox batch + 4 timestamp
  const newCols = [];
  Object.keys(COLUMN_TO_TYPE).forEach(c => {
    if (colIndex_(headers, c) === -1) newCols.push(c);
  });
  Object.keys(COLUMN_TO_TS).forEach(c => {
    const tsCol = COLUMN_TO_TS[c];
    if (colIndex_(headers, tsCol) === -1) newCols.push(tsCol);
  });

  newCols.forEach(col => {
    const last = sellers.getLastColumn();
    sellers.insertColumnAfter(last);
    sellers.getRange(1, last + 1).setValue(col);
  });

  if (newCols.length > 0) {
    headers = getHeaders_(sellers);
    SpreadsheetApp.flush();
  }

  // Aplikuj checkbox validation na 4 kolumnach Send_*
  const checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
  Object.keys(COLUMN_TO_TYPE).forEach(col => {
    const idx = colIndex_(headers, col);
    if (idx > -1) sellers.getRange(2, idx + 1, 500, 1).setDataValidation(checkRule);
  });

  ui.alert('Setup v3.1 zakonczony',
    'Dodano ' + newCols.length + ' nowych kolumn:\n' +
    (newCols.length > 0 ? '  ' + newCols.join('\n  ') : '  (zadnych - juz byly)') + '\n\n' +
    'UWAGA: kolumny Cold Pitch (Send_ColdPitch, TS_ColdPitch, PDFLink_ColdPitch,\n' +
    'DraftLink_ColdPitch) zarzadza pitch_addon.gs - uruchom "PDF Addon: instalacja v2.12".\n\n' +
    'Nastepne kroki:\n' +
    '  1. Konfiguracja -> PDF Addon: instalacja v2.12 (tworzy kolumny Cold Pitch)\n' +
    '  2. Konfiguracja -> Instaluj triggery\n' +
    '  3. Konfiguracja -> Zablokuj naglowki\n' +
    '  4. Jesli masz dane w starych kolumnach Email1/2/Reminder:\n' +
    '     Konfiguracja -> MIGRACJA v2 -> v3',
    ui.ButtonSet.OK);
}

/* =================== MIGRACJA v2 -> v3.1 =================== */
// Migracja danych ze starego systemu Email1/Email2/Reminder.
// Send_Email1 (cold pitch starszy) -> Send_ColdPitch (manual cold pitch + PDF).
function migrateV2toV30() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  if (!sheet) { ui.alert('Brak SP_Sellers'); return; }

  let headers = getHeaders_(sheet);

  const COL_MAP = {
    'Send_Email1':   'Send_ColdPitch',
    'Send_Email2':   'Send_Landing1',
    'Send_Reminder': 'Send_Landing2',
    'TS_Email1':     'TS_ColdPitch',
    'TS_Email2':     'TS_Landing1',
    'TS_Reminder':   'TS_Landing2'
  };

  const STATUS_MAP = {
    'Email1_Sent':   'ColdPitch_Sent',
    'Email2_Sent':   'Landing1_Sent',
    'Reminder_Sent': 'Landing2_Sent'
  };

  const legacyPresent = Object.keys(COL_MAP).filter(c => colIndex_(headers, c) > -1);
  if (legacyPresent.length === 0) {
    ui.alert('Migracja',
      'Brak starych kolumn Email1/Email2/Reminder w arkuszu - nie ma czego migrowac.\n\n' +
      'Jesli kolumny Send_Pitch/Landing1/2/Purchase tez nie ma, uruchom najpierw "Pierwsza instalacja v3.1".\n' +
      'Jesli kolumny Send_ColdPitch nie ma, uruchom "PDF Addon: instalacja v2.12".',
      ui.ButtonSet.OK);
    return;
  }

  const conf1 = ui.alert('Migracja v2 -> v3 KROK 1/3: kopiowanie kolumn',
    'Wykryto ' + legacyPresent.length + ' starych kolumn:\n  ' +
    legacyPresent.join('\n  ') + '\n\n' +
    'KROK 1: skopiuje wartosci do nowych kolumn:\n' +
    '  Send_Email1 -> Send_ColdPitch\n' +
    '  Send_Email2 -> Send_Landing1\n' +
    '  Send_Reminder -> Send_Landing2\n' +
    'KROK 2: zaktualizuje statusy (Email1_Sent -> ColdPitch_Sent itd.).\n' +
    'KROK 3: zapyta o usuniecie starych kolumn.\n\n' +
    'Kontynuowac KROK 1?',
    ui.ButtonSet.YES_NO);
  if (conf1 !== ui.Button.YES) return;

  // Utworz brakujace nowe kolumny (te ze STATUS_MAP + Pitch + Purchase)
  Object.values(COL_MAP).forEach(newCol => {
    if (colIndex_(headers, newCol) === -1) {
      const last = sheet.getLastColumn();
      sheet.insertColumnAfter(last);
      sheet.getRange(1, last + 1).setValue(newCol);
    }
  });
  ['Send_Pitch','TS_Pitch','Send_Purchase','TS_Purchase'].forEach(newCol => {
    if (colIndex_(headers, newCol) === -1) {
      const last = sheet.getLastColumn();
      sheet.insertColumnAfter(last);
      sheet.getRange(1, last + 1).setValue(newCol);
    }
  });
  SpreadsheetApp.flush();
  headers = getHeaders_(sheet);

  // KROK 1: kopiuj dane
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Brak wierszy danych'); return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let migratedRows = 0;

  for (let i = 0; i < data.length; i++) {
    let migratedThis = false;
    Object.keys(COL_MAP).forEach(oldCol => {
      const newCol = COL_MAP[oldCol];
      const oldIdx = colIndex_(headers, oldCol);
      const newIdx = colIndex_(headers, newCol);
      if (oldIdx === -1 || newIdx === -1) return;
      const oldVal = data[i][oldIdx];
      // Tylko skopiuj jesli nowa komorka jest pusta
      if (oldVal !== '' && oldVal !== null && oldVal !== undefined && data[i][newIdx] === '') {
        sheet.getRange(i + 2, newIdx + 1).setValue(oldVal);
        migratedThis = true;
      }
    });

    // Aktualizuj Status jesli pasuje do mapy
    const iStatus = colIndex_(headers, 'Status');
    if (iStatus > -1) {
      const oldStatus = String(data[i][iStatus] || '').trim();
      if (STATUS_MAP[oldStatus]) {
        sheet.getRange(i + 2, iStatus + 1).setValue(STATUS_MAP[oldStatus]);
        migratedThis = true;
      }
    }

    if (migratedThis) migratedRows++;
  }

  SpreadsheetApp.flush();

  // KROK 2: aplikuj data validation na nowych checkbox batch (BEZ Send_ColdPitch!)
  const checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
  Object.keys(COLUMN_TO_TYPE).forEach(col => {
    const idx = colIndex_(headers, col);
    if (idx > -1) sheet.getRange(2, idx + 1, 500, 1).setDataValidation(checkRule);
  });

  // KROK 3: pytanie o usuniecie starych kolumn
  const conf3 = ui.alert('Migracja v3 KROK 3/3: usuniecie starych kolumn',
    'Przeniesiono dane dla ' + migratedRows + ' wierszy.\n\n' +
    'KROK 3: usunac ' + legacyPresent.length + ' starych kolumn?\n  ' +
    legacyPresent.join('\n  ') + '\n\n' +
    'TO JEST DESTRUKCYJNE.\n' +
    'Kliknij NO jesli wolisz zachowac stare kolumny do weryfikacji.',
    ui.ButtonSet.YES_NO);

  if (conf3 !== ui.Button.YES) {
    ui.alert('Migracja czesciowa zakonczona',
      'Skopiowano: ' + migratedRows + ' wierszy.\n' +
      'Stare kolumny zachowano (mozna usunac recznie pozniej).',
      ui.ButtonSet.OK);
    return;
  }

  // Usun stare kolumny od prawej do lewej
  const toDelete = [];
  const headersFresh = getHeaders_(sheet);
  legacyPresent.forEach(col => {
    const idx = colIndex_(headersFresh, col);
    if (idx > -1) toDelete.push(idx + 1);
  });
  toDelete.sort((a, b) => b - a);

  let deleted = 0;
  toDelete.forEach(colNum => {
    sheet.deleteColumn(colNum);
    deleted++;
  });

  ui.alert('Migracja v3.1 zakonczona',
    'Skopiowano: ' + migratedRows + ' wierszy.\n' +
    'Usunieto starych kolumn: ' + deleted + '\n\n' +
    'System dziala teraz na 4 nowych checkboxach batch + Cold Pitch (manual).',
    ui.ButtonSet.OK);
}

/* =================== TRIGGERY =================== */
function installTriggers() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onEditTrigger','processEmailQueue'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('processEmailQueue').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert('Triggery zainstalowane! System gotowy.');
}

function protectHeaders() {
  [SHEET_SELLERS,SHEET_TEMPLATES,SHEET_LOG,SHEET_CONFIG].forEach(name => {
    const sh = SpreadsheetApp.getActive().getSheetByName(name);
    if (!sh) return;
    try {
      const p = sh.getRange(1,1,1,sh.getMaxColumns()).protect();
      p.setDescription('Naglowki');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    } catch(e) {}
  });
  SpreadsheetApp.getUi().alert('Naglowki zablokowane');
}

function reloadConfig() {
  CONFIG_CACHE = null;
  const cfg = getConfig_();
  const keys = Object.keys(cfg);
  const urlKeys = keys.filter(k => k.startsWith('LANDING_URL_') || k.startsWith('CHECKOUT_URL_'));
  SpreadsheetApp.getUi().alert('Config przeladowany',
    'Liczba kluczy: ' + keys.length + '\n\n' +
    'Klucze URL (LANDING/CHECKOUT):\n' +
    (urlKeys.length > 0
      ? urlKeys.map(k => '  ' + k + ' = ' + (cfg[k] || '(puste)').substring(0, 60)).join('\n')
      : '  (brak)'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* =================== onEdit - CHECKBOX =================== */
function onEditTrigger(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_SELLERS) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row === 1) return;

    const headers = getHeaders_(sheet);
    const colName = headers[col - 1];
    if (!VALID_SEND_COLS.includes(colName)) return;

    const isChecked = e.range.getValue() === true || e.value === 'TRUE';

    // === ODZNACZENIE: usun z kolejki ===
    if (!isChecked) {
      removeFromQueue_(row, colName);
      logAction_('UNCHECKED', '', '', '', '', 'W' + row + ' ' + colName + ' odznaczony - usuniety z kolejki');
      return;
    }

    // === ZAZNACZENIE: walidacja i dodanie do kolejki ===

    // Sprawdz Purchased
    const iPurchased = colIndex_(headers, 'Purchased');
    if (iPurchased > -1) {
      const p = sheet.getRange(row, iPurchased+1).getValue();
      if (p === true || String(p).toUpperCase() === 'TRUE') {
        SpreadsheetApp.getUi().alert('Ten seller juz kupil! Email nie wyslany.');
        e.range.setValue(false);
        return;
      }
    }

    const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

    // Email: priorytet Verified, fallback Amazon
    const iEV = colIndex_(headers, 'Email_Verified');
    const iEA = colIndex_(headers, 'Email_Amazon');
    let email = '';
    if (iEV > -1) email = String(rowData[iEV] || '').trim();
    if (!email && iEA > -1) email = String(rowData[iEA] || '').trim();
    if (!email) {
      SpreadsheetApp.getUi().alert('Brak emaila w wierszu ' + row);
      e.range.setValue(false);
      return;
    }

    const iSN = colIndex_(headers, 'SellerName');
    const iSI = colIndex_(headers, 'SellerID');
    const sellerName = iSN > -1 ? String(rowData[iSN] || '').trim() : '';
    const sellerID = iSI > -1 ? String(rowData[iSI] || '').trim() : '';
    const language = getLanguageForRow_(headers, rowData);
    const emailType = COLUMN_TO_TYPE[colName];

    const templateKey = 'SP_' + emailType + '_' + language.toUpperCase();
    const delayMs = parseInt(getConfigValue_('EMAIL_DELAY_MS', String(EMAIL_DELAY_MS_DEFAULT)));

    queueAdd_({
      row: row, column: colName, email: email,
      sellerName: sellerName, sellerID: sellerID,
      templateKey: templateKey, language: language,
      added: Date.now(), due: Date.now() + delayMs,
    });

    logAction_('QUEUED', sellerID, email, templateKey, sellerName,
      'W' + row + ', lang=' + language + ', za ' + Math.round(delayMs/1000) + 's');
  } catch (err) {
    console.error('onEditTrigger error:', err);
  }
}

/* =================== EMAIL QUEUE =================== */
function queueGet_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_EMAIL_QUEUE);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch(e) { return []; }
}

function queueSave_(arr) {
  PropertiesService.getScriptProperties().setProperty(PROP_EMAIL_QUEUE, JSON.stringify(arr || []));
}

function queueAdd_(item) {
  const lock = LockService.getScriptLock();
  lock.tryLock(1000);
  const queue = queueGet_();
  const key = item.row + ':' + item.column;
  const filtered = queue.filter(x => (x.row + ':' + x.column) !== key);
  filtered.push(item);
  queueSave_(filtered);
  if (lock.hasLock()) lock.releaseLock();
}

function removeFromQueue_(row, column) {
  const lock = LockService.getScriptLock();
  lock.tryLock(1000);
  const queue = queueGet_();
  const key = row + ':' + column;
  const filtered = queue.filter(x => (x.row + ':' + x.column) !== key);
  queueSave_(filtered);
  if (lock.hasLock()) lock.releaseLock();
}

function showQueue() {
  const queue = queueGet_();
  const ui = SpreadsheetApp.getUi();
  if (queue.length === 0) { ui.alert('Kolejka pusta.'); return; }
  const now = Date.now();
  const lines = queue.map(j => {
    const w = Math.max(0, Math.round((j.due - now) / 1000));
    return 'W' + j.row + ' | ' + j.templateKey + ' | ' + j.email + ' | ' + (w > 0 ? w + 's' : 'gotowy');
  });
  ui.alert('Kolejka (' + queue.length + ')', lines.join('\n'), ui.ButtonSet.OK);
}

/* =================== STOP - anuluj wszystkie pending =================== */
function stopAllInQueue() {
  const ui = SpreadsheetApp.getUi();
  const queue = queueGet_();

  if (queue.length === 0) {
    ui.alert('Kolejka jest pusta - nic do anulowania.');
    return;
  }

  const r = ui.alert('STOP - anuluj wszystkie pending',
    'W kolejce jest ' + queue.length + ' pending email(i).\n\n' +
    'Wszystkie zostana:\n' +
    '  1. USUNIETE z kolejki\n' +
    '  2. Checkbox zostanie ODZNACZONY (false) w SP_Sellers\n' +
    '  3. Wpis CANCELLED_BY_USER pojawi sie w SP_LOG\n\n' +
    'Maile JUZ wyslane (DONE) NIE sa cofniete.\n\n' +
    'Kontynuowac?',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  let cancelled = 0;

  queue.forEach(job => {
    try {
      const colIdx = colIndex_(headers, job.column);
      if (colIdx > -1) {
        sheet.getRange(job.row, colIdx + 1).setValue(false);
      }
      logAction_('CANCELLED_BY_USER', job.sellerID, job.email, job.templateKey, job.sellerName,
        'W' + job.row + ' ' + job.column + ' STOP zatrzymano');
      cancelled++;
    } catch (err) {
      console.error('STOP error for W' + job.row + ': ' + err.message);
    }
  });

  queueSave_([]);
  ui.alert('Zatrzymano ' + cancelled + ' email(i). Kolejka pusta.');
}

/* =================== PROCESS QUEUE =================== */
function processEmailQueue() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  let queue = queueGet_();
  const now = Date.now();
  const remain = [];
  CONFIG_CACHE = null;

  queue.forEach(job => {
    // Jeszcze nie czas
    if (now < job.due) { remain.push(job); return; }

    const colIdx = colIndex_(headers, job.column);
    if (colIdx === -1) return;

    // ===== BUGFIX v2: PONOWNE SPRAWDZENIE CHECKBOXA =====
    // Jesli uzytkownik odznaczy checkbox w czasie oczekiwania,
    // email NIE zostanie wyslany.
    const cellValue = sheet.getRange(job.row, colIdx+1).getValue();
    if (cellValue !== true) {
      logAction_('CANCELLED', job.sellerID, job.email, job.templateKey, job.sellerName,
        'W' + job.row + ' checkbox=' + String(cellValue) + ' (nie true) - wysylka wstrzymana');
      return;  // NIE dodajemy do remain = usuwamy z kolejki
    }

    // Dodatkowe sprawdzenie: czy nie kupil w miedzyczasie
    const iPurchased = colIndex_(headers, 'Purchased');
    if (iPurchased > -1) {
      const purchased = sheet.getRange(job.row, iPurchased+1).getValue();
      if (purchased === true || String(purchased).toUpperCase() === 'TRUE') {
        sheet.getRange(job.row, colIdx+1).setValue(false);
        logAction_('CANCELLED', job.sellerID, job.email, job.templateKey, job.sellerName,
          'W' + job.row + ' - seller kupil w miedzyczasie');
        return;
      }
    }

    // ===== WYSYLKA =====
    try {
      const rowData = sheet.getRange(job.row, 1, 1, headers.length).getValues()[0];
      sendSitePatronEmail_(job.templateKey, job.email, headers, rowData, job.language);

      // DONE
      sheet.getRange(job.row, colIdx+1).setValue('DONE');

      // Timestamp
      const tsCol = COLUMN_TO_TS[job.column];
      const iTsCol = colIndex_(headers, tsCol);
      if (iTsCol > -1) sheet.getRange(job.row, iTsCol+1).setValue(nowTZ_());

      // Status update (tylko w gore - nie nadpisuj wyzszego statusu)
      const iStatus = colIndex_(headers, 'Status');
      if (iStatus > -1) {
        const newStatus = COLUMN_TO_STATUS[job.column];
        if (newStatus) {
          const cur = STATUS_ORDER.indexOf(String(sheet.getRange(job.row, iStatus+1).getValue()));
          const nw = STATUS_ORDER.indexOf(newStatus);
          if (nw > cur) sheet.getRange(job.row, iStatus+1).setValue(newStatus);
        }
      }

      // Global timestamp
      const iTs = colIndex_(headers, 'Timestamp');
      if (iTs > -1) sheet.getRange(job.row, iTs+1).setValue(nowTZ_());

      logAction_('SENT', job.sellerID, job.email, job.templateKey, job.sellerName, 'W' + job.row);

    } catch (err) {
      sheet.getRange(job.row, colIdx+1).setValue('ERROR');
      logAction_('ERROR', job.sellerID, job.email, job.templateKey, job.sellerName,
        'W' + job.row + ': ' + err.message);
    }
  });

  queueSave_(remain);
  if (lock.hasLock()) lock.releaseLock();
}

/* =================== SEND EMAIL =================== */
function sendSitePatronEmail_(templateKey, toEmail, headers, rowData, language) {
  const templates = SpreadsheetApp.getActive().getSheetByName(SHEET_TEMPLATES);
  if (!templates) throw new Error('Brak arkusza ' + SHEET_TEMPLATES);

  const tData = templates.getDataRange().getValues();
  let subject = '', body = '';
  let fbSubject = '', fbBody = '';

  // Template key np. SP_PITCH_PL, fallback SP_PITCH_EN
  const fallbackKey = templateKey.replace(/_[A-Z]{2,}$/, '_EN');

  for (let i = 1; i < tData.length; i++) {
    const key = String(tData[i][0]).trim().toUpperCase();
    // Kolumna D(3) = Subject, E(4) = Body
    if (key === templateKey.toUpperCase()) {
      subject = String(tData[i][3] || '').trim();
      body = String(tData[i][4] || '').trim();
    }
    if (key === fallbackKey.toUpperCase()) {
      fbSubject = String(tData[i][3] || '').trim();
      fbBody = String(tData[i][4] || '').trim();
    }
  }

  // Fallback na EN jesli brak szablonu w danym jezyku
  if (!subject || !body) { subject = fbSubject; body = fbBody; }
  if (!subject || !body) throw new Error('Brak szablonu: ' + templateKey + ' i fallback ' + fallbackKey);

  subject = replacePlaceholders_(subject, headers, rowData);
  body = replacePlaceholders_(body, headers, rowData);

  // Cleanup: jesli ContactPerson pusty, "Hi ," -> "Hi," itd.
  body = body.replace(/Hi\s+,/g, 'Hi,');
  subject = subject.replace(/Hi\s+,/g, 'Hi,');

  // ===== WYSYLKA Z ALIASU info@netanaliza.com =====
  GmailApp.sendEmail(toEmail, subject, '', {
    from: getConfigValue_('EMAIL_FROM', DEFAULT_FROM_EMAIL),
    name: getConfigValue_('EMAIL_FROM_NAME', DEFAULT_FROM_NAME),
    replyTo: getConfigValue_('EMAIL_REPLY_TO', DEFAULT_REPLY_TO),
    htmlBody: body,
  });
}

/* =================== PLACEHOLDERS - z FALLBACK URL =================== */
function replacePlaceholders_(text, headers, rowData) {
  const colMap = {};
  headers.forEach((h, idx) => { colMap[String(h).trim()] = idx; });
  const cfg = getConfig_();

  // Aliasy: rożne nazwy kolumn mapuja na te same dane
  const aliases = {
    'Firma':'BusinessName','SellerNameAmazon':'SellerName',
    'Email':'Email_Verified','Telefon':'Phone','ASIN':'TopASIN',
  };

  return text.replace(/\{\{(\w+)\}\}/g, function(match, key) {
    // 1. Szukaj w kolumnach arkusza
    if (colMap.hasOwnProperty(key)) {
      const val = String(rowData[colMap[key]] || '');
      if (val) return val;
    }
    // 2. Szukaj w aliasach
    if (aliases[key] && colMap.hasOwnProperty(aliases[key])) {
      const val = String(rowData[colMap[aliases[key]]] || '');
      if (val) return val;
    }
    // 3. Szukaj w konfiguracji
    if (cfg[key]) return cfg[key];

    // 4. FALLBACK URL: LANDING_URL_DE -> LANDING_URL_EN, CHECKOUT_URL_DE_MONTHLY -> CHECKOUT_URL_EN_MONTHLY
    let fallbackKey = null;
    let m = key.match(/^(LANDING_URL)_[A-Z]{2}$/);
    if (m) {
      fallbackKey = m[1] + '_EN';
    } else {
      m = key.match(/^(CHECKOUT_URL)_[A-Z]{2}(_[A-Z]+)$/);
      if (m) fallbackKey = m[1] + '_EN' + m[2];
    }
    if (fallbackKey && fallbackKey !== key && cfg[fallbackKey]) {
      console.log('FALLBACK_URL: ' + key + ' -> ' + fallbackKey);
      return cfg[fallbackKey];
    }

    // 5. Nie znaleziono = pusty string (nie zostawiaj {{placeholder}})
    return '';
  });
}

/* =================== BATCH OPS (BEZ Cold Pitch) =================== */
function batchSend_Pitch_AfterDelay() {
  _batchSendAfterDelay_(
    'Send_Pitch', 'TS_ColdPitch', 'ColdPitch_Sent', 1,
    'Pitch (po Cold Pitch)', 'po Cold Pitch'
  );
}

function batchSend_Landing1_Leads() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Batch Landing 1',
    'Wyslac Landing 1 do WSZYSTKICH ze statusem "Lead"?\n\n' +
    '(Lead = klient odpowiedzial na Cold Pitch / Pitch i wyrazil zainteresowanie)',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const iSend = colIndex_(headers, 'Send_Landing1');
  if (iSend === -1) {
    ui.alert('Brak kolumny Send_Landing1. Uruchom "Pierwsza instalacja v3.1".');
    return;
  }
  const iEV = colIndex_(headers, 'Email_Verified');
  const iEA = colIndex_(headers, 'Email_Amazon');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow-1, headers.length).getValues();
  let count = 0;
  data.forEach((row, idx) => {
    const status = String(row[iStatus] || '').trim();
    const send = String(row[iSend] || '').trim();
    const email = String(row[iEV] || row[iEA] || '').trim();
    if (status === 'Lead' && send !== 'DONE' && send !== 'ERROR' && email) {
      sheet.getRange(idx+2, iSend+1).setValue(true);
      count++;
    }
  });
  ui.alert('Zaznaczono ' + count + ' sellerow. Landing 1 za ~2 min.');
}

function batchSend_Landing2_AfterDelay() {
  _batchSendAfterDelay_(
    'Send_Landing2', 'TS_Landing1', 'Landing1_Sent', 3,
    'Landing 2 (po Landing 1)', 'po Landing 1'
  );
}

// Helper dla batch po opoznieniu
function _batchSendAfterDelay_(sendCol, tsCol, requiredStatus, defaultDays, title, statusLabel) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Batch ' + title, 'Ile dni ' + statusLabel + '?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const days = parseInt(resp.getResponseText()) || defaultDays;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const iSend = colIndex_(headers, sendCol);
  if (iSend === -1) {
    ui.alert('Brak kolumny ' + sendCol + '. Uruchom "Pierwsza instalacja v3.1".');
    return;
  }
  const iTs = colIndex_(headers, tsCol);
  const iEV = colIndex_(headers, 'Email_Verified');
  const iEA = colIndex_(headers, 'Email_Amazon');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow-1, headers.length).getValues();
  const now = new Date();
  let count = 0;

  data.forEach((row, idx) => {
    const status = String(row[iStatus] || '');
    const send = String(row[iSend] || '');
    const ts = iTs > -1 ? row[iTs] : null;
    const email = String(row[iEV] || row[iEA] || '').trim();
    if (status === requiredStatus && send !== 'DONE' && send !== 'ERROR' && email && ts) {
      const sentDate = parseDate_(ts);
      if (sentDate && (now - sentDate) / 86400000 >= days) {
        sheet.getRange(idx+2, iSend+1).setValue(true);
        count++;
      }
    }
  });
  ui.alert('Zaznaczono ' + count + ' (' + statusLabel + ' >= ' + days + ' dni temu). Wysylka za ~2 min.');
}

/* =================== STATS =================== */
function showStats() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Brak danych'); return; }

  const data = sheet.getRange(2, iStatus+1, lastRow-1, 1).getValues();
  const counts = {};
  data.forEach(r => { const s = String(r[0] || 'Brak').trim(); counts[s] = (counts[s]||0)+1; });

  const lines = STATUS_ORDER.filter(s => counts[s]).map(s => s + ': ' + counts[s]);

  // Statusy ktore nie sa w STATUS_ORDER (np. stare wiersze sprzed migracji)
  Object.keys(counts).forEach(s => {
    if (!STATUS_ORDER.includes(s)) lines.push(s + ': ' + counts[s] + ' (nieznany)');
  });

  const queueLen = queueGet_().length;
  lines.push('---');
  lines.push('W kolejce: ' + queueLen);

  SpreadsheetApp.getUi().alert('Pipeline (' + data.length + ' total)', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/* =================== LOGGING =================== */
function logAction_(action, sellerID, email, template, sellerName, details) {
  try {
    const log = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
    if (log) log.appendRow([nowTZ_(), action, sellerID, email, template, sellerName, details]);
  } catch(e) { console.error('Log error:', e); }
}

/* =================== HELPERS =================== */
function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h));
}

function colIndex_(headers, name) {
  const n = String(name).toLowerCase().trim();
  let idx = headers.findIndex(h => String(h).toLowerCase().trim() === n);
  if (idx > -1) return idx;
  const fb = {
    'email_verified':['email','e-mail','mail','kontaktemail'],
    'email_amazon':['emailamazon','amazon email'],
    'sellername':['seller name','seller name amazon'],
    'sellerid':['seller id','amazon seller id'],
    'businessname':['firma','nazwa firmy','company'],
    'contactperson':['osoba','contact person'],
    'nichesite':['niche site','strona niszowa'],
    'nichesiteurl':['niche site url','url strony'],
    'patronpreviewurl':['patron preview url','preview url'],
    // Aliasy starych kolumn (uzywane przez migrateV2toV30)
    'send_email1':['send email1','wyslij email1'],
    'send_email2':['send email2','wyslij email2'],
    'send_reminder':['send reminder','wyslij reminder'],
    'ts_email1':['timestamp email1'],
    'ts_email2':['timestamp email2'],
    'ts_reminder':['timestamp reminder'],
    // Aliasy nowych kolumn v3.1
    'send_coldpitch':['send cold pitch','wyslij cold pitch'],
    'send_pitch':['send pitch','wyslij pitch'],
    'send_landing1':['send landing1','wyslij landing1'],
    'send_landing2':['send landing2','wyslij landing2'],
    'send_purchase':['send purchase','wyslij purchase'],
    'ts_coldpitch':['timestamp coldpitch','ts cold pitch'],
    'ts_pitch':['timestamp pitch'],
    'ts_landing1':['timestamp landing1'],
    'ts_landing2':['timestamp landing2'],
    'ts_purchase':['timestamp purchase'],
    'pdflink_coldpitch':['pdf link coldpitch','pdf coldpitch'],
    'pdflink_pitch':['pdf link pitch','pdf pitch'],
    'draftlink_coldpitch':['draft link coldpitch','draft coldpitch'],
    'draftlink_pitch':['draft link pitch','draft pitch'],
    'purchased':['kupil'],
  };
  if (fb[n]) {
    for (const v of fb[n]) {
      idx = headers.findIndex(h => String(h).toLowerCase().trim() === v);
      if (idx > -1) return idx;
    }
  }
  return -1;
}

function parseDate_(val) {
  if (val instanceof Date) return val;
  const s = String(val || '');
  if (!s) return null;
  // Format: dd.MM.yyyy HH:mm:ss
  const parts = s.split(' ')[0].split('.');
  if (parts.length === 3) return new Date(parts[2], parts[1]-1, parts[0]);
  // Fallback: proba natywnego parsowania
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function nowTZ_() {
  return Utilities.formatDate(new Date(), getConfigValue_('TIMEZONE', TZ_DEFAULT), 'dd.MM.yyyy HH:mm:ss');
}
