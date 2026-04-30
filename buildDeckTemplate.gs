/**********************************************************
 *  BUILD DECK TEMPLATE v3 - 15 unikalnych layoutów
 *
 *  ZMIANA vs v2:
 *  - Cena $59/$499 wpisana na sztywno (nie placeholder)
 *    - bo placeholder ${{SP_PRICE_MONTHLY}} jest dłuższy
 *      od wartości i rozwala wielkie fonty
 *  - Slide 5 layout naprawiony (mniejsze elementy, value
 *    labels nie wyłażą)
 *  - Slide 6 prawa kolumna zmieniona (cena na sztywno + opis)
 *
 *  Placeholdery które ZOSTAJĄ:
 *  {{TITLE_S01}}..{{TITLE_S15}}, {{BODY_S01}}..{{BODY_S15}}
 *  {{Brand}}, {{ContactPerson}}, {{NicheSiteURL}}
 *
 *  Placeholdery USUNIĘTE z template (wpisane na sztywno):
 *  {{SP_PRICE_MONTHLY}} = 59, {{SP_PRICE_ANNUAL}} = 499
 *
 *  UZYCIE: Run > buildDeckTemplate
 **********************************************************/

const TEMPLATE_ID = '1Z3HRS4dqCetfXY4ow7bRJHq4uWNwT7K63i6RwA6NCSo';

const SW = 720;  // Slides width pt
const SH = 405;  // Slides height pt

const H = {
  NAVY: '#1A2942',
  CREAM: '#F8F4EC',
  GOLD: '#C9A961',
  WHITE: '#FFFFFF',
  DARK: '#1A2942',
  MUTED: '#6B7280',
  RED: '#D95252',
  GREEN: '#4D9966',
  ACCBG: '#EDE5D3',
  GOLD_LIGHT: '#D9CDB0',
  GREY_BORDER: '#E0E0E0'
};

// === HELPERS ===
function fillBoxHex_(slide, x, y, w, h, hexColor) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill(hexColor);
  shape.getBorder().setTransparent();
  return shape;
}
function strokeBoxHex_(slide, x, y, w, h, fillHex, strokeHex, weight) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill(fillHex);
  shape.getBorder().getLineFill().setSolidFill(strokeHex);
  shape.getBorder().setWeight(weight || 0.5);
  return shape;
}
function txHex_(slide, x, y, w, h, text, opts) {
  opts = opts || {};
  const tb = slide.insertTextBox(text, x, y, w, h);
  const ts = tb.getText().getTextStyle();
  ts.setFontFamily(opts.font || 'Inter');
  ts.setFontSize(opts.size || 11);
  ts.setForegroundColor(opts.color || H.DARK);
  if (opts.bold) ts.setBold(true);
  if (opts.italic) ts.setItalic(true);
  if (opts.align === 'center') tb.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  else if (opts.align === 'right') tb.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  return tb;
}
function header_(slide, badgeText) {
  fillBoxHex_(slide, 0, 0, SW, 32, H.ACCBG);
  if (badgeText) txHex_(slide, 24, 11, 200, 12, badgeText, { font: 'Inter', size: 7, bold: true, color: H.MUTED });
  txHex_(slide, SW - 220, 11, 200, 12, '✦ LUKO SITE PATRON', { font: 'Inter', size: 7, bold: true, color: H.NAVY, align: 'right' });
}
function pageNum_(slide, n) {
  txHex_(slide, SW - 60, SH - 18, 40, 12, n + ' / 15', { font: 'Inter', size: 6, color: H.MUTED, align: 'right' });
}

// === MAIN ===
function buildDeckTemplate() {
  const ui = SpreadsheetApp.getUi();
  let pres;
  try { pres = SlidesApp.openById(TEMPLATE_ID); }
  catch (e) { ui.alert('Blad otwarcia template: ' + e.message); return; }
  
  const existing = pres.getSlides();
  for (let i = existing.length - 1; i >= 0; i--) existing[i].remove();
  
  buildSlide01_(pres);
  buildSlide02_(pres);
  buildSlide03_(pres);
  buildSlide04_(pres);
  buildSlide05_(pres);
  buildSlide06_(pres);
  buildSlide07_(pres);
  buildSlide08_(pres);
  buildSlide09_(pres);
  buildSlide10_(pres);
  buildSlide11_(pres);
  buildSlide12_(pres);
  buildSlide13_(pres);
  buildSlide14_(pres);
  buildSlide15_(pres);
  
  pres.saveAndClose();
  ui.alert('Template zbudowany v3',
    'Otworz: https://docs.google.com/presentation/d/' + TEMPLATE_ID + '/edit',
    ui.ButtonSet.OK);
}

// === SLIDE 1 - COVER ===
function buildSlide01_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.NAVY);
  fillBoxHex_(slide, 45, 45, 70, 2, H.GOLD);
  txHex_(slide, 45, 56, 400, 14, '✦ LUKO SITE PATRON', { font: 'Inter', size: 8, bold: true, color: H.GOLD });
  txHex_(slide, 45, 110, 630, 130, '{{TITLE_S01}}', { font: 'Playfair Display', size: 32, bold: true, color: H.WHITE });
  txHex_(slide, 45, 260, 630, 80, '{{BODY_S01}}', { font: 'Inter', size: 11, color: H.GOLD_LIGHT });
  txHex_(slide, SW - 270, SH - 22, 250, 14, 'NETANALIZA LTD · KWIECIEŃ 2026', { font: 'Inter', size: 6, color: '#8A95A8', align: 'right' });
}

// === SLIDE 2 - OFERTA ===
function buildSlide02_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'CO OFERUJĘ');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S02}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 110, 650, 270, '{{BODY_S02}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 2);
}

// === SLIDE 3 - PROBLEM ===
function buildSlide03_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'PROBLEM');
  txHex_(slide, 34, 56, 650, 50, '{{TITLE_S03}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 113, 650, 1, H.GOLD);
  txHex_(slide, 34, 125, 650, 255, '{{BODY_S03}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 3);
}

// === SLIDE 4 - ROZWIĄZANIE ===
function buildSlide04_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'ROZWIĄZANIE');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S04}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  fillBoxHex_(slide, 34, 115, 650, 70, H.NAVY);
  txHex_(slide, 46, 122, 626, 56, '{{BODY_S04}}', { font: 'Inter', size: 9, color: H.WHITE });
  pageNum_(slide, 4);
}

// === SLIDE 5 - LICZBY (POPRAWIONE: mniejsze proporcje, czytelne) ===
function buildSlide05_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'LICZBY');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S05}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  
  // duza liczba 10-50% - dopasowana
  txHex_(slide, 100, 115, 520, 64, '10–50%',
    { font: 'Playfair Display', size: 56, bold: true, color: H.NAVY, align: 'center' });
  txHex_(slide, 100, 180, 520, 14, 'osób które weszły na Stronę Site Patron — kupiło na Amazon',
    { font: 'Inter', size: 9, bold: true, color: H.NAVY, align: 'center' });
  
  // 3 slupki - precyzyjne pozycje
  const barAreaY = 215;
  const barMaxH = 70;
  const barW = 90;
  const barGap = 50;
  const totalBarW = 3 * barW + 2 * barGap;
  const barStartX = (SW - totalBarW) / 2;
  
  const bars = [
    { label: 'Amazon Ads — średnia', value: '4–5%', h: 18, color: H.MUTED },
    { label: 'Amazon Ads — szczyt', value: '~10%', h: 35, color: H.MUTED },
    { label: 'Site Patron', value: '10–50%', h: 70, color: H.GOLD }
  ];
  
  bars.forEach(function(b, i) {
    const x = barStartX + i * (barW + barGap);
    const barTopY = barAreaY + (barMaxH - b.h);
    // value LABEL nad slupkiem (16pt nie 16+ żeby nie wyłaziło)
    txHex_(slide, x - 10, barTopY - 18, barW + 20, 14,
      b.value, { font: 'Playfair Display', size: 12, bold: true, color: b.color, align: 'center' });
    // slupek
    fillBoxHex_(slide, x, barTopY, barW, b.h, b.color);
    // label pod slupkiem
    txHex_(slide, x - 10, barAreaY + barMaxH + 6, barW + 20, 14,
      b.label, { font: 'Inter', size: 7, color: H.MUTED, align: 'center' });
  });
  
  // body opis pod
  txHex_(slide, 34, 340, 650, 40, '{{BODY_S05}}', { font: 'Inter', size: 7, color: H.MUTED, italic: true });
  pageNum_(slide, 5);
}

// === SLIDE 6 - CO ODRÓŻNIA (cena $59 na sztywno) ===
function buildSlide06_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'CO ODRÓŻNIA TEN MODEL');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S06}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  
  // LEWA - inne kanaly
  strokeBoxHex_(slide, 34, 115, 320, 245, H.WHITE, H.GREY_BORDER, 0.5);
  txHex_(slide, 46, 125, 296, 24, 'Inne kanały — koszt rośnie z każdym zamówieniem',
    { font: 'Inter', size: 8, bold: true, color: H.MUTED });
  const otherCh = [
    'Amazon Ads (PPC) — CPC × każde kliknięcie',
    'Influencerzy / afilianci — % prowizji',
    'Google / Facebook Ads — CPC × kliknięcia',
    'Reklama display — CPM × wyświetlenia',
    'E-mail marketing — opłata × wysyłka',
    'Agencja SEO — 2-5 tys/mies + długie umowy'
  ];
  otherCh.forEach(function(ch, i) {
    txHex_(slide, 46, 162 + i * 30, 12, 14, '✗', { font: 'Inter', size: 9, bold: true, color: H.RED });
    txHex_(slide, 64, 162 + i * 30, 280, 16, ch, { font: 'Inter', size: 7, color: H.NAVY });
  });
  
  // PRAWA - granat z cena $59 NA SZTYWNO
  fillBoxHex_(slide, 364, 115, 320, 245, H.NAVY);
  fillBoxHex_(slide, 364, 115, 320, 2, H.GOLD);
  txHex_(slide, 376, 125, 296, 14, 'LUKO SITE PATRON — opłata stała',
    { font: 'Inter', size: 8, bold: true, color: H.GOLD });
  // CENA NA SZTYWNO 59 (nie placeholder!)
  txHex_(slide, 376, 152, 296, 70, '$59',
    { font: 'Playfair Display', size: 60, bold: true, color: H.WHITE, align: 'center' });
  txHex_(slide, 376, 222, 296, 14, 'miesięcznie — i tyle',
    { font: 'Inter', size: 9, color: H.GOLD_LIGHT, align: 'center' });
  
  const benefits = [
    'Każde zamówienie ze Strony — bez kosztu',
    'Każde wyświetlenie — bezpłatne',
    'Każdy klient B2B — bez prowizji'
  ];
  benefits.forEach(function(b, i) {
    txHex_(slide, 388, 260 + i * 22, 12, 14, '✓', { font: 'Inter', size: 9, bold: true, color: H.GOLD });
    txHex_(slide, 408, 260 + i * 22, 264, 18, b, { font: 'Inter', size: 7, color: H.WHITE });
  });
  pageNum_(slide, 6);
}

// === SLIDE 7 - ALTERNATYWY ===
function buildSlide07_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'CO BYŚ MUSIAŁ ZROBIĆ SAM');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S07}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 110, 650, 270, '{{BODY_S07}}', { font: 'Inter', size: 8, color: H.DARK });
  pageNum_(slide, 7);
}

// === SLIDE 8 - KORZYŚCI 1-4 ===
function buildSlide08_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'KORZYŚCI 1—4');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S08}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 113, 650, 270, '{{BODY_S08}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 8);
}

// === SLIDE 9 - KORZYŚCI 5-8 ===
function buildSlide09_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'KORZYŚCI 5—8');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S09}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 113, 650, 270, '{{BODY_S09}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 9);
}

// === SLIDE 10 - BONUS B2B ===
function buildSlide10_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'BONUS B2B');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S10}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 113, 650, 270, '{{BODY_S10}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 10);
}

// === SLIDE 11 - WYŁĄCZNOŚĆ ===
function buildSlide11_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'WYŁĄCZNOŚĆ');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S11}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 113, 650, 270, '{{BODY_S11}}', { font: 'Inter', size: 9, color: H.DARK });
  pageNum_(slide, 11);
}

// === SLIDE 12 - PRAWNE (jeden duzy box ze zlotym lewym pasem) ===
function buildSlide12_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'ASPEKTY PRAWNE');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S12}}', { font: 'Playfair Display', size: 20, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  strokeBoxHex_(slide, 34, 115, 650, 250, H.WHITE, H.GOLD, 0.5);
  fillBoxHex_(slide, 34, 115, 4, 250, H.GOLD);
  txHex_(slide, 50, 130, 626, 240, '{{BODY_S12}}', { font: 'Inter', size: 8, color: H.DARK });
  pageNum_(slide, 12);
}

// === SLIDE 13 - CENNIK (ceny $59/$499 NA SZTYWNO) ===
function buildSlide13_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'CENNIK');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S13}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  txHex_(slide, 34, 113, 650, 18, 'Bez prowizji. Bez CPC. Bez ukrytych kosztów. 30-dniowa gwarancja zwrotu = zero ryzyka.',
    { font: 'Inter', size: 8, color: H.MUTED });
  
  // LEWA karta - Miesieczny
  strokeBoxHex_(slide, 80, 145, 270, 200, H.WHITE, H.GREY_BORDER, 0.5);
  txHex_(slide, 80, 165, 270, 18, 'Miesięczny', { font: 'Playfair Display', size: 14, bold: true, color: H.NAVY, align: 'center' });
  // CENA $59 NA SZTYWNO
  txHex_(slide, 80, 195, 270, 60, '$59',
    { font: 'Playfair Display', size: 50, bold: true, color: H.NAVY, align: 'center' });
  txHex_(slide, 80, 260, 270, 12, 'miesięcznie', { font: 'Inter', size: 8, color: H.MUTED, align: 'center' });
  fillBoxHex_(slide, 105, 280, 220, 0.5, H.GREY_BORDER);
  txHex_(slide, 80, 290, 270, 12, 'Min. okres: 3 miesiące', { font: 'Inter', size: 7, color: H.MUTED, align: 'center' });
  txHex_(slide, 80, 305, 270, 12, 'Auto-przedłużenie do anulowania', { font: 'Inter', size: 7, color: H.MUTED, align: 'center' });
  
  // PRAWA karta - Roczny POLECANE
  strokeBoxHex_(slide, 370, 145, 270, 200, H.WHITE, H.GOLD, 1);
  fillBoxHex_(slide, 410, 130, 190, 22, H.GOLD);
  txHex_(slide, 410, 134, 190, 14, 'POLECANE — OSZCZĘDZASZ 209 USD',
    { font: 'Inter', size: 7, bold: true, color: H.NAVY, align: 'center' });
  txHex_(slide, 370, 165, 270, 18, 'Roczny', { font: 'Playfair Display', size: 14, bold: true, color: H.NAVY, align: 'center' });
  // CENA $499 NA SZTYWNO
  txHex_(slide, 370, 195, 270, 60, '$499',
    { font: 'Playfair Display', size: 50, bold: true, color: H.NAVY, align: 'center' });
  txHex_(slide, 370, 260, 270, 12, 'rocznie', { font: 'Inter', size: 8, color: H.MUTED, align: 'center' });
  fillBoxHex_(slide, 395, 280, 220, 0.5, H.GREY_BORDER);
  txHex_(slide, 370, 290, 270, 12, 'Min. okres: 1 rok', { font: 'Inter', size: 7, color: H.MUTED, align: 'center' });
  txHex_(slide, 370, 305, 270, 12, 'Auto-przedłużenie do anulowania', { font: 'Inter', size: 7, color: H.MUTED, align: 'center' });
  
  fillBoxHex_(slide, 34, 358, 650, 30, H.ACCBG);
  txHex_(slide, 34, 364, 650, 12, '✓ 30-dniowa gwarancja  ✓ Anulowanie kliknięciem  ✓ Niski próg ~240 zł/mies',
    { font: 'Inter', size: 7, bold: true, color: H.NAVY, align: 'center' });
  txHex_(slide, 34, 376, 650, 12, '✓ Strona już istnieje  ✓ Każde zamówienie ze Strony — bez kosztu',
    { font: 'Inter', size: 7, bold: true, color: H.NAVY, align: 'center' });
  pageNum_(slide, 13);
}

// === SLIDE 14 - NASTĘPNE KROKI (timeline 5 kroków) ===
function buildSlide14_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.CREAM);
  header_(slide, 'NASTĘPNE KROKI');
  txHex_(slide, 34, 56, 650, 36, '{{TITLE_S14}}', { font: 'Playfair Display', size: 18, bold: true, color: H.NAVY });
  fillBoxHex_(slide, 34, 99, 650, 1, H.GOLD);
  
  const timelineY = 140;
  fillBoxHex_(slide, 80, timelineY + 18, 560, 1, H.GOLD);
  
  const steps = [
    { num: '01', title: 'Wybór planu', desc: 'Mailem (mies./rok)\nlub pytania', tag: 'TAG 0' },
    { num: '02', title: 'Płatność', desc: 'Stripe/PayPal,\nfaktura natychmiast', tag: 'TAG 1' },
    { num: '03', title: 'Logo + opis', desc: 'SVG/PNG\n+ 2-3 zdania', tag: 'TAG 1—2' },
    { num: '04', title: 'Wyłączność', desc: 'Strona pokazuje\ntylko {{Brand}}', tag: 'TAG 2—3' },
    { num: '05', title: 'Pierwszy raport', desc: 'Po 30 dniach\nliczby. Co miesiąc.', tag: '+30 DNI' }
  ];
  const stepW = 110, stepGap = 5;
  const totalW = 5 * stepW + 4 * stepGap;
  const startX = (SW - totalW) / 2;
  steps.forEach(function(s, i) {
    const x = startX + i * (stepW + stepGap);
    const circ = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x + stepW / 2 - 18, timelineY, 36, 36);
    circ.getFill().setSolidFill(H.NAVY);
    circ.getBorder().setTransparent();
    txHex_(slide, x + stepW / 2 - 18, timelineY + 8, 36, 20, s.num,
      { font: 'Playfair Display', size: 13, bold: true, color: H.GOLD, align: 'center' });
    txHex_(slide, x, timelineY + 50, stepW, 14, s.title, { font: 'Inter', size: 8, bold: true, color: H.NAVY, align: 'center' });
    txHex_(slide, x, timelineY + 66, stepW, 36, s.desc, { font: 'Inter', size: 6, color: H.MUTED, align: 'center' });
    txHex_(slide, x, timelineY + 110, stepW, 12, s.tag, { font: 'Inter', size: 6, bold: true, color: H.GOLD, align: 'center' });
  });
  
  fillBoxHex_(slide, 34, 320, 650, 65, H.ACCBG);
  txHex_(slide, 46, 328, 626, 56, '{{BODY_S14}}', { font: 'Inter', size: 7, color: H.NAVY });
  pageNum_(slide, 14);
}

// === SLIDE 15 - CLOSING ===
function buildSlide15_(pres) {
  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(H.NAVY);
  fillBoxHex_(slide, 45, 30, 70, 2, H.GOLD);
  txHex_(slide, 45, 40, 400, 14, '✦ LUKO SITE PATRON', { font: 'Inter', size: 8, bold: true, color: H.GOLD });
  txHex_(slide, 45, 80, 630, 50, '{{TITLE_S15}}', { font: 'Playfair Display', size: 30, bold: true, color: H.WHITE });
  txHex_(slide, 45, 145, 630, 180, '{{BODY_S15}}', { font: 'Inter', size: 9, color: H.GOLD_LIGHT });
  fillBoxHex_(slide, 45, 350, 630, 30, H.GOLD);
  txHex_(slide, 45, 358, 630, 16, 'ODPISZ NA TEN MAIL I RUSZAMY',
    { font: 'Inter', size: 11, bold: true, color: H.NAVY, align: 'center' });
}
