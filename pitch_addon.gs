/**********************************************************
 *  SitePatron Pitch + PDF Addon (v2.12)
 *
 *  ZMIANY vs v2.11:
 *  - ROZDZIELENIE KOLUMN: Pitsch -> ColdPitch + Pitch (osobne 4-tki)
 *      COLDPITCH (manual cold pitch z PDF):
 *        Send_ColdPitch, TS_ColdPitch, PDFLink_ColdPitch, DraftLink_ColdPitch
 *      PITCH (batch bez PDF + manual case-by-case z PDF, wspolnie):
 *        Send_Pitch, TS_Pitch, PDFLink_Pitch, DraftLink_Pitch
 *  - generateColdPitchDraft -> kolumny *_ColdPitch
 *  - generatePitchDraft -> kolumny *_Pitch (wspolne z batch w Sitepatron emailautomation.gs)
 *  - generatePdfOnlyForRow -> PDFLink_ColdPitch + TS_ColdPitch (default)
 *  - NOWA FUNKCJA: migratePitschToColdPitch - przenosi dane z Send_Pitsch/...
 *    do Send_ColdPitch/... (bo Pitsch byl uzywany glownie dla cold pitch).
 *    Pyta przed kazdym destruktywnym krokiem.
 *
 *  ZMIANY vs v2.11 -> v2.12 SZCZEGOLOWO:
 *  - PITSCH_COLS (stare) USUNIETE
 *  - LEGACY_COLS rozszerzone: Send_Pitsch/TS_Pitsch/PDFLink_Pitsch/DraftLink_Pitsch
 *
 *  WYMAGA server.py v1.3.2+ (pdf_filename + Drive upload przez service account)
 *  + GOOGLE_SERVICE_ACCOUNT_JSON na Railway
 *  + service account jako Content Manager Shared Drive
 *  + DRIVE_FOLDER_ID w SP_Config = ID Shared Drive
 *
 *  Wersja: 2.12 | 2026-04-30 | NetAnaliza
 **********************************************************/

const SHEET_DECK_SLIDES = 'SP_DeckSlides';
const SHEET_PITCH_TEMPLATES = 'Sheet1';
const DRIVE_FOLDER_NAME_DEFAULT = 'SitePatron-Generated-PDFs';

const RENDER_ENDPOINT_KEY = 'RENDER_ENDPOINT_URL';
const RENDER_API_KEY_KEY = 'RENDER_API_KEY';

// Nowe kolumny v2.12: rozdzielenie ColdPitch / Pitch
const COLDPITCH_COLS = ['Send_ColdPitch', 'TS_ColdPitch', 'PDFLink_ColdPitch', 'DraftLink_ColdPitch'];
const PITCH_COLS     = ['Send_Pitch',     'TS_Pitch',     'PDFLink_Pitch',     'DraftLink_Pitch'];

// Legacy: stare kolumny ColdPitch/Pitch z v2.10 (przed Pitsch) + Pitsch z v2.11
const LEGACY_COLS = [
  // Z v2.10 (jesli ktos jest sprzed v2.11)
  'PDFLink_Cold', 'DraftLink_Cold',
  // Z v2.11 (Pitsch wspolny)
  'Send_Pitsch', 'TS_Pitsch', 'PDFLink_Pitsch', 'DraftLink_Pitsch'
];

/* =================== MENU EXTENSION =================== */
// Sugerowane wpisy menu (w Sitepatron emailautomation.gs onOpen):
//
//   .addSubMenu(SpreadsheetApp.getUi().createMenu('Pitch + PDF (drafty)')
//     .addItem('Cold Pitch + PDF (zaznaczony wiersz)', 'generateColdPitchDraft')
//     .addItem('Pitch po reakcji + PDF (zaznaczony wiersz)', 'generatePitchDraft')
//     .addItem('Tylko PDF (do podgladu)', 'generatePdfOnlyForRow')
//     .addSeparator()
//     .addItem('Otworz folder Drive', 'openDriveFolder_v25')
//     .addItem('Pierwsza instalacja v2.5', 'firstSetupV25')
//     .addItem('Pierwsza instalacja v2.6 (nowy render)', 'firstSetupV26')
//     .addItem('MIGRACJA Pitsch -> ColdPitch', 'migratePitschToColdPitch'))

/* =================== FIRST SETUP V2.5 (zaktualizowane v2.12) =================== */
// Tworzy 8 kolumn: 4 dla ColdPitch + 4 dla Pitch.
function firstSetupV25() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const sellers = ss.getSheetByName(SHEET_SELLERS);
  if (!sellers) {
    ui.alert('Blad', 'Brak SP_Sellers. Najpierw uruchom firstSetup z v2.', ui.ButtonSet.OK);
    return;
  }
  const templates = ss.getSheetByName(SHEET_TEMPLATES);
  if (!templates) {
    ui.alert('Blad', 'Brak SP_EmailTemplates.', ui.ButtonSet.OK);
    return;
  }

  let headers = getHeaders_(sellers);
  let addedCols = 0;
  // ColdPitch (manual cold pitch z PDF)
  COLDPITCH_COLS.forEach(col => {
    if (colIndex_(headers, col) === -1) {
      const last = sellers.getLastColumn();
      sellers.insertColumnAfter(last);
      sellers.getRange(1, last + 1).setValue(col);
      addedCols++;
    }
  });
  // Pitch (batch + manual case-by-case z PDF)
  PITCH_COLS.forEach(col => {
    if (colIndex_(headers, col) === -1) {
      const last = sellers.getLastColumn();
      sellers.insertColumnAfter(last);
      sellers.getRange(1, last + 1).setValue(col);
      addedCols++;
    }
  });

  headers = getHeaders_(sellers);
  // Send_ColdPitch - wartosci tekstowe (DRAFT_CREATED/DONE/ERROR), bez checkbox
  // Send_Pitch - checkbox dla batch
  const checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(true).build();
  const idxSendPitch = colIndex_(headers, 'Send_Pitch');
  if (idxSendPitch > -1) sellers.getRange(2, idxSendPitch + 1, 500, 1).setDataValidation(checkRule);
  // Send_ColdPitch - usuwamy checkbox jesli byl (wartosci tekstowe)
  const idxSendCold = colIndex_(headers, 'Send_ColdPitch');
  if (idxSendCold > -1) sellers.getRange(2, idxSendCold + 1, 500, 1).clearDataValidations();

  let deck = ss.getSheetByName(SHEET_DECK_SLIDES);
  if (!deck) {
    deck = ss.insertSheet(SHEET_DECK_SLIDES);
    deck.getRange(1, 1, 1, 5).setValues([['SlideKey', 'SlideNumber', 'Language', 'Title', 'Body']]);
    deck.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#37474F').setFontColor('#FFFFFF');
    deck.setFrozenRows(1);
  }

  const folderName = getConfigValue_('DRIVE_FOLDER_NAME', DRIVE_FOLDER_NAME_DEFAULT);
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) folder = folders.next();
  else folder = DriveApp.createFolder(folderName);

  const config = ss.getSheetByName(SHEET_CONFIG);
  if (config) {
    const data = config.getDataRange().getValues();
    let foundId = false, foundName = false;
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0]).trim();
      if (k === 'DRIVE_FOLDER_ID') foundId = true;
      if (k === 'DRIVE_FOLDER_NAME') foundName = true;
    }
    if (!foundId) config.appendRow(['DRIVE_FOLDER_ID', folder.getId(), 'ID Shared Drive lub folderu Drive na PDFy']);
    if (!foundName) config.appendRow(['DRIVE_FOLDER_NAME', folderName, 'Nazwa folderu Drive (legacy)']);
    CONFIG_CACHE = null;
  }

  ui.alert('Setup v2.12 OK',
    'Dodano kolumn: ' + addedCols + ' (ColdPitch + Pitch)\n' +
    'Folder Drive: ' + folder.getName() + '\n' + folder.getUrl() +
    '\n\nUWAGA: jesli masz stare kolumny Send_Pitsch (z literowka) - uruchom migratePitschToColdPitch.',
    ui.ButtonSet.OK);
}

/* =================== FIRST SETUP V2.6 =================== */
function firstSetupV26() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sellers = ss.getSheetByName(SHEET_SELLERS);
  if (!sellers) {
    ui.alert('Blad', 'Brak SP_Sellers. Uruchom najpierw firstSetupV25.', ui.ButtonSet.OK);
    return;
  }

  let headers = getHeaders_(sellers);
  const newCols = [
    'MultibrandURL',
    'B2B_Quantity',
    'B2B_Product_PL',
    'B2B_Buyers_NOM',
    'B2B_Buyers_GEN'
  ];
  let addedCols = 0;
  newCols.forEach(col => {
    if (colIndex_(headers, col) === -1) {
      const last = sellers.getLastColumn();
      sellers.insertColumnAfter(last);
      sellers.getRange(1, last + 1).setValue(col);
      addedCols++;
    }
  });

  const config = ss.getSheetByName(SHEET_CONFIG);
  let foundEndpoint = false, foundApiKey = false;
  if (config) {
    const data = config.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0]).trim();
      if (k === RENDER_ENDPOINT_KEY) foundEndpoint = true;
      if (k === RENDER_API_KEY_KEY) foundApiKey = true;
    }
    if (!foundEndpoint) config.appendRow([RENDER_ENDPOINT_KEY, '', 'URL endpointu Python']);
    if (!foundApiKey) config.appendRow([RENDER_API_KEY_KEY, '', 'Sekretny klucz endpointu']);
  }
  CONFIG_CACHE = null;

  ui.alert('Setup v2.6 OK',
    'Dodano kolumn: ' + addedCols + ' (z ' + newCols.length + ')\n\n' +
    'WAZNE: DRIVE_FOLDER_ID w SP_Config musi wskazywac Shared Drive\n' +
    '(np. "0AORrQ2XLr8xcUk9PVA"). Service account musi byc Content Manager.',
    ui.ButtonSet.OK);
}

/* =================== MIGRACJA Pitsch -> ColdPitch (v2.12) =================== */
/* Odpal RAZ. Przenosi dane z Send_Pitsch/TS_Pitsch/PDFLink_Pitsch/DraftLink_Pitsch
 * do Send_ColdPitch/TS_ColdPitch/PDFLink_ColdPitch/DraftLink_ColdPitch.
 * Zaklada ze Pitsch byl uzywany glownie dla cold pitch (95%+ przypadkow).
 */
function migratePitschToColdPitch() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_SELLERS);
  if (!sheet) { ui.alert('Brak SP_Sellers'); return; }

  let headers = getHeaders_(sheet);

  // Krok 1: sprawdz czy nowe kolumny ColdPitch juz istnieja
  let addedNew = 0;
  COLDPITCH_COLS.forEach(col => {
    if (colIndex_(headers, col) === -1) {
      const last = sheet.getLastColumn();
      sheet.insertColumnAfter(last);
      sheet.getRange(1, last + 1).setValue(col);
      addedNew++;
    }
  });
  if (addedNew > 0) {
    headers = getHeaders_(sheet);
    SpreadsheetApp.flush();
  }

  // Krok 2: sprawdz ile starych kolumn istnieje
  const PITSCH_LEGACY = ['Send_Pitsch', 'TS_Pitsch', 'PDFLink_Pitsch', 'DraftLink_Pitsch'];
  const legacyPresent = PITSCH_LEGACY.filter(c => colIndex_(headers, c) > -1);

  const newColsIdx = {
    Send: colIndex_(headers, 'Send_ColdPitch'),
    Ts: colIndex_(headers, 'TS_ColdPitch'),
    Pdf: colIndex_(headers, 'PDFLink_ColdPitch'),
    Draft: colIndex_(headers, 'DraftLink_ColdPitch')
  };

  if (legacyPresent.length === 0) {
    ui.alert('Migracja Pitsch -> ColdPitch',
      'Brak starych kolumn Send_Pitsch/TS_Pitsch/PDFLink_Pitsch/DraftLink_Pitsch.\n' +
      'Nic do migracji.\n\n' +
      'Nowe kolumny ColdPitch ' + (addedNew > 0 ? 'zostaly utworzone (' + addedNew + ').' : 'juz istnieja.'),
      ui.ButtonSet.OK);
    return;
  }

  const conf1 = ui.alert('Migracja Pitsch -> ColdPitch KROK 1/2: kopiowanie danych',
    'Zostalo wykryto ' + legacyPresent.length + ' starych kolumn:\n  ' +
    legacyPresent.join('\n  ') + '\n\n' +
    'KROK 1: skopiuje dane do nowych kolumn ColdPitch:\n' +
    '  Send_Pitsch -> Send_ColdPitch\n' +
    '  TS_Pitsch -> TS_ColdPitch\n' +
    '  PDFLink_Pitsch -> PDFLink_ColdPitch\n' +
    '  DraftLink_Pitsch -> DraftLink_ColdPitch\n' +
    'KROK 2: po Twoim potwierdzeniu - usunie stare kolumny.\n\n' +
    'Kontynuowac KROK 1 (kopiowanie)?',
    ui.ButtonSet.YES_NO);
  if (conf1 !== ui.Button.YES) return;

  // Krok 1: kopia danych
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Brak wierszy danych.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const idxOld = {
    Send: colIndex_(headers, 'Send_Pitsch'),
    Ts: colIndex_(headers, 'TS_Pitsch'),
    Pdf: colIndex_(headers, 'PDFLink_Pitsch'),
    Draft: colIndex_(headers, 'DraftLink_Pitsch')
  };

  // Pobierz formuly osobno (HYPERLINK musi byc zachowany jako formula)
  const formulas = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getFormulas();

  let migratedRows = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowFormulas = formulas[i];

    function pickFormulaOrValue(idx) {
      if (idx === -1) return null;
      const f = rowFormulas[idx];
      if (f && f.length > 0) return { formula: f };
      const v = row[idx];
      if (v !== '' && v !== null && v !== undefined) return { value: v };
      return null;
    }

    let migratedThisRow = false;

    // Send_Pitsch -> Send_ColdPitch
    const sendVal = pickFormulaOrValue(idxOld.Send);
    if (sendVal && newColsIdx.Send > -1 && (data[i][newColsIdx.Send] === '' || data[i][newColsIdx.Send] === null)) {
      const cell = sheet.getRange(i + 2, newColsIdx.Send + 1);
      if (sendVal.formula) cell.setFormula(sendVal.formula);
      else cell.setValue(sendVal.value);
      migratedThisRow = true;
    }

    // TS_Pitsch -> TS_ColdPitch
    const tsVal = pickFormulaOrValue(idxOld.Ts);
    if (tsVal && newColsIdx.Ts > -1 && (data[i][newColsIdx.Ts] === '' || data[i][newColsIdx.Ts] === null)) {
      const cell = sheet.getRange(i + 2, newColsIdx.Ts + 1);
      cell.setValue(tsVal.value !== undefined ? tsVal.value : tsVal.formula);
      migratedThisRow = true;
    }

    // PDFLink_Pitsch -> PDFLink_ColdPitch
    const pdfVal = pickFormulaOrValue(idxOld.Pdf);
    if (pdfVal && newColsIdx.Pdf > -1 && (data[i][newColsIdx.Pdf] === '' || data[i][newColsIdx.Pdf] === null)) {
      const cell = sheet.getRange(i + 2, newColsIdx.Pdf + 1);
      if (pdfVal.formula) cell.setFormula(pdfVal.formula);
      else cell.setValue(pdfVal.value);
      migratedThisRow = true;
    }

    // DraftLink_Pitsch -> DraftLink_ColdPitch
    const draftVal = pickFormulaOrValue(idxOld.Draft);
    if (draftVal && newColsIdx.Draft > -1 && (data[i][newColsIdx.Draft] === '' || data[i][newColsIdx.Draft] === null)) {
      const cell = sheet.getRange(i + 2, newColsIdx.Draft + 1);
      if (draftVal.formula) cell.setFormula(draftVal.formula);
      else cell.setValue(draftVal.value);
      migratedThisRow = true;
    }

    if (migratedThisRow) migratedRows++;
  }

  // Send_ColdPitch - WARTOSCI TEKSTOWE (bez checkbox)
  // Ale dla bezpieczenstwa usun data validation jesli istnial
  if (newColsIdx.Send > -1) {
    sheet.getRange(2, newColsIdx.Send + 1, Math.max(500, lastRow), 1).clearDataValidations();
  }

  SpreadsheetApp.flush();

  // Krok 2: pytanie o usuniecie starych kolumn
  const conf2 = ui.alert('Migracja Pitsch -> ColdPitch KROK 2/2: usuwanie starych kolumn',
    'Skopiowano dane dla ' + migratedRows + ' wierszy.\n\n' +
    'KROK 2: usunac ' + legacyPresent.length + ' starych kolumn?\n  ' +
    legacyPresent.join('\n  ') + '\n\n' +
    'TO JEST DESTRUKCYJNE - kolumny znikna z arkusza.\n' +
    'Jesli wolisz zachowac je jeszcze - kliknij NO.',
    ui.ButtonSet.YES_NO);

  if (conf2 !== ui.Button.YES) {
    ui.alert('Migracja zakonczona (krok 1)',
      'Skopiowano dane: ' + migratedRows + ' wierszy.\n' +
      'Stare kolumny Pitsch zachowano. Mozesz je usunac recznie pozniej.',
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

  ui.alert('Migracja Pitsch -> ColdPitch zakonczona',
    'Skopiowano dane: ' + migratedRows + ' wierszy.\n' +
    'Usunieto starych kolumn: ' + deleted + '\n\n' +
    'Cold Pitch dziala teraz na kolumnach Send_ColdPitch/TS_ColdPitch/...\n' +
    'Pitch (batch + manual) na kolumnach Send_Pitch/TS_Pitch/...',
    ui.ButtonSet.OK);
}

/* =================== GLOWNE FUNKCJE - DRAFT + PDF =================== */
function generateColdPitchDraft() {
  _generatePitchDraftCommon_(
    'SP_COLDPITCH',     // template prefix (Cold Pitch)
    'COLD',              // logActionLabel - szczegoly w SP_LOG
    'ColdPitch_Sent',    // wartosc Status (jesli aktualizujemy)
    'ColdPitch'          // sufiks kolumn (Send_ColdPitch, TS_ColdPitch, ...)
  );
}

function generatePitchDraft() {
  _generatePitchDraftCommon_(
    'SP_PITCH',          // template prefix (Pitch po reakcji)
    'PITCH',             // logActionLabel
    'Pitch_Sent',        // wartosc Status
    'Pitch'              // sufiks kolumn (Send_Pitch, TS_Pitch, ...)
  );
}

function generatePdfOnlyForRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  if (!sheet) { ui.alert('Brak SP_Sellers'); return; }
  const row = sheet.getActiveRange().getRow();
  if (row === 1) { ui.alert('Zaznacz wiersz sprzedawcy'); return; }

  const headers = getHeaders_(sheet);
  const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const iBrand = colIndex_(headers, 'Brand');
  const brand = String(rowData[iBrand] || 'Unknown').trim();
  const iSellerID = colIndex_(headers, 'SellerID');
  const sellerID = iSellerID > -1 ? String(rowData[iSellerID] || '').trim() : '';
  const iSellerName = colIndex_(headers, 'SellerName');
  const sellerName = iSellerName > -1 ? String(rowData[iSellerName] || '').trim() : '';
  const language = getLanguageForRow_(headers, rowData);

  try {
    SpreadsheetApp.getActive().toast('Generuje PDF deck...', 'Site Patron', 30);
    const pdfFile = generatePdfDeck_(brand, language, headers, rowData);

    // Zapis do PDFLink_ColdPitch + TS_ColdPitch (default - bo PDF only zwykle przed cold pitchem)
    const iPdf = colIndex_(headers, 'PDFLink_ColdPitch');
    const iTs = colIndex_(headers, 'TS_ColdPitch');
    if (iPdf > -1) {
      sheet.getRange(row, iPdf + 1).setFormula(
        '=HYPERLINK("' + pdfFile.getUrl() + '","PDF (' +
        Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM HH:mm') + ')")'
      );
    }
    if (iTs > -1) sheet.getRange(row, iTs + 1).setValue(nowTZ_());

    // Log do SP_LOG
    logAction_('PDF_ONLY', sellerID, '', '', sellerName,
      'W' + row + ' tylko PDF (podglad), brand=' + brand + ', file=' + pdfFile.getName());

    ui.alert('PDF utworzony',
      'Plik: ' + pdfFile.getName() + '\n' +
      (iPdf > -1 ? 'Link zapisany w PDFLink_ColdPitch (W' + row + ')\n\n' : '\n') +
      pdfFile.getUrl(),
      ui.ButtonSet.OK);
  } catch (err) {
    Logger.log('generatePdfOnlyForRow error: ' + err.message + '\n' + err.stack);
    ui.alert('Blad', err.message, ui.ButtonSet.OK);
  }
}

function _generatePitchDraftCommon_(templatePrefix, logLabel, statusValue, colSuffix) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SELLERS);
  if (!sheet) { ui.alert('Brak SP_Sellers'); return; }

  const row = sheet.getActiveRange().getRow();
  if (row === 1) { ui.alert('Zaznacz wiersz sprzedawcy (nie naglowek)'); return; }

  const headers = getHeaders_(sheet);
  const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  const iEV = colIndex_(headers, 'Email_Verified');
  const iEA = colIndex_(headers, 'Email_Amazon');
  const email = String(rowData[iEV] || rowData[iEA] || '').trim();
  if (!email) { ui.alert('Brak emaila w wierszu ' + row); return; }

  const iBrand = colIndex_(headers, 'Brand');
  const brand = String(rowData[iBrand] || 'Unknown').trim();
  const iSellerID = colIndex_(headers, 'SellerID');
  const sellerID = iSellerID > -1 ? String(rowData[iSellerID] || '').trim() : '';
  const iSellerName = colIndex_(headers, 'SellerName');
  const sellerName = iSellerName > -1 ? String(rowData[iSellerName] || '').trim() : '';
  const language = getLanguageForRow_(headers, rowData);
  const templateKey = templatePrefix + '_' + language.toUpperCase();

  const iPurchased = colIndex_(headers, 'Purchased');
  if (iPurchased > -1) {
    const p = rowData[iPurchased];
    if (p === true || String(p).toUpperCase() === 'TRUE') {
      ui.alert('Ten sprzedawca juz kupil. Draft nie utworzony.');
      return;
    }
  }

  const conf = ui.alert('Generuj draft + PDF',
    'Marka: ' + brand + '\n' +
    'Email: ' + email + '\n' +
    'Jezyk: ' + language + '\n' +
    'Template: ' + templateKey + '\n' +
    'Typ: ' + logLabel + '\n' +
    'Kolumny: Send_' + colSuffix + ', TS_' + colSuffix + ', PDFLink_' + colSuffix + ', DraftLink_' + colSuffix +
    '\n\nKontynuowac?',
    ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  CONFIG_CACHE = null;
  try {
    SpreadsheetApp.getActive().toast('Generuje PDF deck...', 'Site Patron', 30);
    const pdfFile = generatePdfDeck_(brand, language, headers, rowData);

    SpreadsheetApp.getActive().toast('Tworze email draft...', 'Site Patron', 10);
    const draft = createEmailDraftWithAttachment_(templateKey, email, headers, rowData, language, pdfFile);

    // Zapis do 4 kolumn z sufiksem (ColdPitch lub Pitch)
    const iSend = colIndex_(headers, 'Send_' + colSuffix);
    const iTs = colIndex_(headers, 'TS_' + colSuffix);
    const iPdf = colIndex_(headers, 'PDFLink_' + colSuffix);
    const iDraft = colIndex_(headers, 'DraftLink_' + colSuffix);
    const iStatus = colIndex_(headers, 'Status');

    if (iSend > -1) sheet.getRange(row, iSend + 1).setValue('DRAFT_CREATED');
    if (iTs > -1) sheet.getRange(row, iTs + 1).setValue(nowTZ_());
    if (iPdf > -1) {
      sheet.getRange(row, iPdf + 1).setFormula(
        '=HYPERLINK("' + pdfFile.getUrl() + '","PDF (' +
        Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM HH:mm') + ')")'
      );
    }
    if (iDraft > -1) {
      const draftUrl = 'https://mail.google.com/mail/u/0/#drafts/' + draft.getId();
      sheet.getRange(row, iDraft + 1).setFormula(
        '=HYPERLINK("' + draftUrl + '","' + logLabel + ' draft")'
      );
    }
    if (iStatus > -1 && _shouldUpdateStatusV25_(sheet.getRange(row, iStatus + 1).getValue(), statusValue)) {
      sheet.getRange(row, iStatus + 1).setValue(statusValue);
    }

    logAction_('DRAFT_CREATED_' + logLabel, sellerID, email, templateKey, sellerName,
      'W' + row + ' draft+PDF utworzony, brand=' + brand + ', file=' + pdfFile.getName() + ', kol=' + colSuffix);

    const r = ui.alert('Gotowe',
      'PDF: ' + pdfFile.getName() + '\nDraft: w Gmail Drafts\n' +
      'Zapisano w kolumnach *_' + colSuffix + '\n\nOtworzyc Drafts?',
      ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) openGmailDrafts_v25();

  } catch (err) {
    Logger.log('_generatePitchDraftCommon_ error: ' + err.message + '\n' + err.stack);
    logAction_('ERROR_' + logLabel, sellerID, email, templateKey, sellerName,
      'W' + row + ': ' + err.message);
    ui.alert('Blad', err.message + '\n\nView > Logs.', ui.ButtonSet.OK);
  }
}

function _shouldUpdateStatusV25_(currentStatus, newStatus) {
  const order = ['New','Researched',
                 'ColdPitch_Sent','Pitch_Sent','Lead',
                 'Landing1_Sent','Landing2_Sent',
                 'Interested','Registered',
                 'Purchase_Sent','Purchased',
                 'Declined','DoNotContact'];
  const cur = order.indexOf(String(currentStatus || ''));
  const nw = order.indexOf(newStatus);
  return nw > cur;
}

/* =================== HELPERS V2.10+ =================== */

function _extractDomainName_(url) {
  if (!url) return '';
  let s = String(url).trim().toLowerCase();
  s = s.replace(/^https?:\/\//, '');
  s = s.replace(/^www\./, '');
  s = s.split('/')[0];
  s = s.split('.')[0];
  s = s.replace(/[^a-z0-9-]/g, '');
  return s;
}

function _sanitizeForFilename_(s) {
  if (!s) return '';
  return String(s)
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9_-]/g, '')
    .substring(0, 50);
}

function _buildPdfFilename_(headers, rowData) {
  function getVal(colName) {
    const idx = colIndex_(headers, colName);
    if (idx === -1) return '';
    return String(rowData[idx] || '').trim();
  }

  const sellerName = _sanitizeForFilename_(getVal('SellerName')) || 'Unknown';
  const domain = _extractDomainName_(getVal('NicheSiteURL')) || 'no-domain';
  const sellerID = _sanitizeForFilename_(getVal('SellerID')) || 'no-id';
  const country = _sanitizeForFilename_(getVal('Country')).toUpperCase().substring(0, 2) || 'XX';
  const ts = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyyMMdd-HHmm');

  return 'sitepatron-' + sellerName + '-' + domain + '-' + sellerID + '-' + country + '-' + ts + '.pdf';
}

function _driveRetry_(fn, opName) {
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      Logger.log(opName + ' proba ' + attempt + ' nieudana: ' + e.message);
      if (attempt < 3) Utilities.sleep(2000);
    }
  }
  throw new Error(opName + ' nieudane po 3 probach. Ostatni blad: ' + (lastErr && lastErr.message));
}

/* =================== PDF DECK GENERATION (z v2.10) ============================================== */
function generatePdfDeck_(brand, language, headers, rowData) {
  const folderId = getConfigValue_('DRIVE_FOLDER_ID', '');
  if (!folderId) throw new Error('Brak DRIVE_FOLDER_ID w SP_Config. Uruchom firstSetupV25.');

  const endpointUrl = getConfigValue_(RENDER_ENDPOINT_KEY, '');
  if (!endpointUrl) {
    throw new Error('Brak ' + RENDER_ENDPOINT_KEY + ' w SP_Config. Uruchom firstSetupV26.');
  }

  const apiKey = getConfigValue_(RENDER_API_KEY_KEY, '');

  function getVal(colName) {
    const idx = colIndex_(headers, colName);
    if (idx === -1) return '';
    return String(rowData[idx] || '').trim();
  }

  const demoUrlA = getVal('NicheSiteURL');
  const demoUrlB = getVal('MultibrandURL');

  const missingRequired = [];
  if (!demoUrlA) missingRequired.push('NicheSiteURL (DEMO_URL_A)');
  if (!demoUrlB) missingRequired.push('MultibrandURL (DEMO_URL_B)');
  if (missingRequired.length > 0) {
    throw new Error('Brak WYMAGANYCH wartosci w wierszu SP_Sellers: ' + missingRequired.join(', '));
  }

  const values = {
    'DEMO_URL_A': demoUrlA,
    'DEMO_URL_B': demoUrlB,
    'B2B_QUANTITY': getVal('B2B_Quantity'),
    'B2B_PRODUCT_PL': getVal('B2B_Product_PL'),
    'B2B_BUYERS_NOM': getVal('B2B_Buyers_NOM'),
    'B2B_BUYERS_GEN': getVal('B2B_Buyers_GEN')
  };

  const pdfFilename = _buildPdfFilename_(headers, rowData);
  const requestHeaders = { 'Content-Type': 'application/json' };
  if (apiKey) requestHeaders['X-API-Key'] = apiKey;
  const renderUrl = endpointUrl.replace(/\/+$/, '') + '/render';

  const payload = {
    language: language,
    values: values,
    drive_folder_id: folderId,
    pdf_filename: pdfFilename
  };

  let response;
  try {
    response = UrlFetchApp.fetch(renderUrl, {
      method: 'post',
      headers: requestHeaders,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
  } catch (err) {
    throw new Error('Endpoint nieosiagalny (' + renderUrl + '): ' + err.message);
  }

  const code = response.getResponseCode();
  if (code !== 200) {
    let errBody = response.getContentText();
    if (errBody.length > 500) errBody = errBody.substring(0, 500) + '...';
    throw new Error('Endpoint zwrocil HTTP ' + code + ': ' + errBody);
  }

  const contentType = response.getHeaders()['Content-Type'] || '';

  if (contentType.indexOf('json') !== -1) {
    let result;
    try {
      result = JSON.parse(response.getContentText());
    } catch (e) {
      throw new Error('Endpoint zwrocil niepoprawny JSON: ' + e.message);
    }
    if (!result.file_id) {
      throw new Error('Endpoint zwrocil JSON bez file_id: ' + JSON.stringify(result).substring(0, 300));
    }
    return _driveRetry_(
      () => DriveApp.getFileById(result.file_id),
      'getFileById(' + result.file_id + ')'
    );
  }

  if (contentType.indexOf('pdf') !== -1) {
    Logger.log('UWAGA: endpoint zwrocil PDF binary zamiast JSON.');
    const pdfBlob = response.getBlob().setContentType('application/pdf');
    pdfBlob.setName(pdfFilename);
    const parentFolder = _driveRetry_(
      () => DriveApp.getFolderById(folderId),
      'getFolderById(' + folderId + ')'
    );
    return _driveRetry_(
      () => parentFolder.createFile(pdfBlob),
      'createFile(' + pdfFilename + ')'
    );
  }

  throw new Error('Endpoint zwrocil nieoczekiwany Content-Type: ' + contentType);
}

/* =================== EMAIL DRAFT Z ATTACHMENT =================== */
function createEmailDraftWithAttachment_(templateKey, toEmail, headers, rowData, language, pdfFile) {
  const ss = SpreadsheetApp.getActive();
  const fallbackKey = templateKey.replace(/_[A-Z]{2,}$/, '_EN');

  function findInSheet(sheetName, key) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return null;
    const tData = sh.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      const k = String(tData[i][0] || '').trim().toUpperCase();
      if (k === key.toUpperCase()) {
        const subject = String(tData[i][3] || '').trim();
        const body = String(tData[i][4] || '').trim();
        if (subject && body) return { subject: subject, body: body };
      }
    }
    return null;
  }

  let found = findInSheet(SHEET_TEMPLATES, templateKey)
    || findInSheet(SHEET_PITCH_TEMPLATES, templateKey)
    || findInSheet(SHEET_TEMPLATES, fallbackKey)
    || findInSheet(SHEET_PITCH_TEMPLATES, fallbackKey);

  if (!found) throw new Error('Brak templejtu: ' + templateKey + ' (ani ' + fallbackKey + ')');

  let subject = replacePlaceholders_(found.subject, headers, rowData);
  let body = replacePlaceholders_(found.body, headers, rowData);
  body = body.replace(/Hi\s+,/g, 'Hi,').replace(/Pan\s+,/g, 'Szanowny Panie,').replace(/Herr\s+,/g, 'Sehr geehrter Herr,');
  subject = subject.replace(/Hi\s+,/g, 'Hi,');

  const options = {
    from: getConfigValue_('EMAIL_FROM', DEFAULT_FROM_EMAIL),
    name: getConfigValue_('EMAIL_FROM_NAME', DEFAULT_FROM_NAME),
    replyTo: getConfigValue_('EMAIL_REPLY_TO', DEFAULT_REPLY_TO),
    htmlBody: body,
    attachments: [pdfFile.getAs('application/pdf')]
  };
  return GmailApp.createDraft(toEmail, subject, '', options);
}

/* =================== HELPERS UI =================== */
function openGmailDrafts_v25() {
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("https://mail.google.com/mail/u/0/#drafts","_blank");google.script.host.close();</script>'
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Otwieram Gmail Drafts...');
}

function openDriveFolder_v25() {
  const folderId = getConfigValue_('DRIVE_FOLDER_ID', '');
  if (!folderId) {
    SpreadsheetApp.getUi().alert('Brak DRIVE_FOLDER_ID w SP_Config.');
    return;
  }
  const url = 'https://drive.google.com/drive/folders/' + folderId;
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank");google.script.host.close();</script>'
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Otwieram folder Drive...');
}