/**********************************************************
 *  SitePatron Email Automation v2
 *  System wysylki emaili z auto-detekcja jezyka
 *
 *  Zmiany vs v1:
 *  - Wysylka z info@netanaliza.com (alias Gmail)
 *  - Nazwa nadawcy: Lukasz from NetAnaliza
 *  - Re-check checkbox PRZED wysylka (bugfix)
 *  - Batch Reminder po X dniach od Email 2
 *  - Czystszy kod, lepsze logi
 *
 *  Wersja: 2.0 | 2026-03-03 | NetAnaliza / Dollar
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
  // NIE wysylamy po niemiecku! DE/AT/CH -> fallback EN
  // Jesli kiedys chcesz DE, odkomentuj linie ponizej:
  // 'DE':'de','AT':'de','CH':'de','LI':'de','LU':'de',
  'DE':'de','AT':'de','CH':'de','LI':'de','LU':'de',
  'GB':'en','UK':'en','US':'en','CA':'en','AU':'en','NZ':'en',
  'IE':'en','ZA':'en','SG':'en','IN':'en','PH':'en','MY':'en',
  'KE':'en','NG':'en','GH':'en','AE':'en','HK':'en',
  'PL':'pl',
  'FR':'fr','BE':'fr','MC':'fr','SN':'fr','CI':'fr','MA':'fr','TN':'fr','DZ':'fr',
  'ES':'es','MX':'es','AR':'es','CO':'es','CL':'es','PE':'es',
  'IT':'it','SM':'it',
  'NL':'nl',
  'SE':'sv',
  'PT':'pt','BR':'pt',
  'CZ':'cs','DK':'da','FI':'fi','NO':'no','RO':'ro','HU':'hu','TR':'tr',
  'JP':'ja','KR':'ko','CN':'zh','TW':'zh',
  'SA':'ar','EG':'ar','IQ':'ar','JO':'ar','KW':'ar','QA':'ar',
  'GR':'el','CY':'el','BG':'bg','HR':'hr','SK':'sk','SI':'sl',
  'LT':'lt','LV':'lv','EE':'et',
};

function getLanguageForRow_(headers, rowData) {
  // Priorytet 1: kolumna Language (reczne nadpisanie)
  const iLang = colIndex_(headers, 'Language');
  if (iLang > -1) {
    const lang = String(rowData[iLang] || '').trim().toLowerCase();
    if (lang && lang.length >= 2) return lang;
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
    .addItem('Instaluj triggery', 'installTriggers')
    .addItem('Sprawdz kolejke', 'showQueue')
    .addItem('Statystyki pipeline', 'showStats')
    .addSeparator()
    .addSubMenu(ui.createMenu('Batch Send')
      .addItem('Email 1 - wszyscy New', 'batchSendEmail1_NewSellers')
      .addItem('Email 2 - po X dniach od Email 1', 'batchSendEmail2_AfterDelay')
      .addItem('Reminder - po X dniach od Email 2', 'batchSendReminder_AfterDelay'))
    .addSeparator()
    .addItem('Pierwsza instalacja', 'firstSetup')
    .addItem('Zablokuj naglowki', 'protectHeaders')
    .addSeparator()
    .addSubMenu(ui.createMenu('Pitch + PDF (drafty)')
      .addItem('Cold Pitch + PDF (zaznaczony wiersz)', 'generateColdPitchDraft')
      .addItem('Pitch po reakcji + PDF (zaznaczony wiersz)', 'generatePitchDraft')
      .addItem('Tylko PDF (do podgladu)', 'generatePdfOnlyForRow')
      .addSeparator()
      .addItem('Otworz folder Drive', 'openDriveFolder_v25')
      .addItem('Pierwsza instalacja v2.5', 'firstSetupV25')
      .addItem('Pierwsza instalacja v2.6 (nowy render)', 'firstSetupV26'))
    .addToUi();
}

/* =================== FIRST SETUP =================== */
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
  ui.alert('Setup zakonczony! Nastepny krok: Instaluj triggery');
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
    const validCols = ['Send_Email1','Send_Email2','Send_Reminder'];
    if (!validCols.includes(colName)) return;

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

    let emailType;
    if (colName === 'Send_Email1') emailType = 'EMAIL1';
    else if (colName === 'Send_Email2') emailType = 'EMAIL2';
    else emailType = 'REMIND';

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
      const tsMap = {'Send_Email1':'TS_Email1','Send_Email2':'TS_Email2','Send_Reminder':'TS_Reminder'};
      const iTsCol = colIndex_(headers, tsMap[job.column]);
      if (iTsCol > -1) sheet.getRange(job.row, iTsCol+1).setValue(nowTZ_());

      // Status update (tylko w gore - nie nadpisuj wyzszego statusu)
      const iStatus = colIndex_(headers, 'Status');
      if (iStatus > -1) {
        const statusMap = {
          'Send_Email1':'Email1_Sent',
          'Send_Email2':'Email2_Sent',
          'Send_Reminder':'Reminder_Sent'
        };
        const newStatus = statusMap[job.column];
        if (newStatus) {
          const order = ['New','Researched','Email1_Sent','Email2_Sent','Reminder_Sent',
                         'Interested','Registered','Purchased','Declined','DoNotContact'];
          const cur = order.indexOf(String(sheet.getRange(job.row, iStatus+1).getValue()));
          const nw = order.indexOf(newStatus);
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

  // Template key np. SP_EMAIL1_PL, fallback SP_EMAIL1_EN
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

/* =================== PLACEHOLDERS =================== */
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
      return val;
    }
    // 2. Szukaj w aliasach
    if (aliases[key] && colMap.hasOwnProperty(aliases[key])) {
      return String(rowData[colMap[aliases[key]]] || '');
    }
    // 3. Szukaj w konfiguracji
    if (cfg[key]) return cfg[key];
    // 4. Nie znaleziono = pusty string (nie zostawiaj {{placeholder}})
    return '';
  });
}

/* =================== BATCH OPS =================== */
function batchSendEmail1_NewSellers() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Batch Email 1',
    'Wyslac Email 1 do WSZYSTKICH ze statusem "New"?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const iSend = colIndex_(headers, 'Send_Email1');
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
    if (status === 'New' && send !== 'DONE' && send !== 'ERROR' && email) {
      sheet.getRange(idx+2, iSend+1).setValue(true);
      count++;
    }
  });
  ui.alert('Zaznaczono ' + count + ' sellerow. Emaile za ~2 min.');
}

function batchSendEmail2_AfterDelay() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Batch Email 2', 'Ile dni po Email 1?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const days = parseInt(resp.getResponseText()) || 1;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const iSend2 = colIndex_(headers, 'Send_Email2');
  const iTs1 = colIndex_(headers, 'TS_Email1');
  const iEV = colIndex_(headers, 'Email_Verified');
  const iEA = colIndex_(headers, 'Email_Amazon');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow-1, headers.length).getValues();
  const now = new Date();
  let count = 0;

  data.forEach((row, idx) => {
    const status = String(row[iStatus] || '');
    const send = String(row[iSend2] || '');
    const ts = row[iTs1];
    const email = String(row[iEV] || row[iEA] || '').trim();
    if (status === 'Email1_Sent' && send !== 'DONE' && send !== 'ERROR' && email && ts) {
      const sentDate = parseDate_(ts);
      if (sentDate && (now - sentDate) / 86400000 >= days) {
        sheet.getRange(idx+2, iSend2+1).setValue(true);
        count++;
      }
    }
  });
  ui.alert('Zaznaczono ' + count + ' (Email1 >= ' + days + ' dni temu). Wysylka za ~2 min.');
}

function batchSendReminder_AfterDelay() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Batch Reminder', 'Ile dni po Email 2?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const days = parseInt(resp.getResponseText()) || 3;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  const headers = getHeaders_(sheet);
  const iStatus = colIndex_(headers, 'Status');
  const iSendR = colIndex_(headers, 'Send_Reminder');
  const iTs2 = colIndex_(headers, 'TS_Email2');
  const iEV = colIndex_(headers, 'Email_Verified');
  const iEA = colIndex_(headers, 'Email_Amazon');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow-1, headers.length).getValues();
  const now = new Date();
  let count = 0;

  data.forEach((row, idx) => {
    const status = String(row[iStatus] || '');
    const send = String(row[iSendR] || '');
    const ts = row[iTs2];
    const email = String(row[iEV] || row[iEA] || '').trim();
    if (status === 'Email2_Sent' && send !== 'DONE' && send !== 'ERROR' && email && ts) {
      const sentDate = parseDate_(ts);
      if (sentDate && (now - sentDate) / 86400000 >= days) {
        sheet.getRange(idx+2, iSendR+1).setValue(true);
        count++;
      }
    }
  });
  ui.alert('Zaznaczono ' + count + ' (Email2 >= ' + days + ' dni temu). Wysylka za ~2 min.');
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

  const order = ['New','Researched','Email1_Sent','Email2_Sent','Reminder_Sent',
                 'Interested','Registered','Purchased','Declined','DoNotContact'];
  const lines = order.filter(s => counts[s]).map(s => s + ': ' + counts[s]);

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
    'send_email1':['send email1','wyslij email1'],
    'send_email2':['send email2','wyslij email2'],
    'send_reminder':['send reminder','wyslij reminder'],
    'ts_email1':['timestamp email1'],
    'ts_email2':['timestamp email2'],
    'ts_reminder':['timestamp reminder'],
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
