/**
 * מגיעים · מערכת אישורי הגעה לאירועים
 * רצה על גוגל שיטס. אין צורך בשרת, בדומיין או בחשבונות נוספים.
 *
 * © כל הזכויות שמורות ליוצרת המערכת. הקוד הזה מוגן בזכות יוצרים.
 * רישיון שימוש אישי, לא בלעדי ולא ניתן להעברה, לבעל הרישיון המפורט למטה בלבד.
 * אסור להעביר, למכור, להשכיר, להעתיק, לשכפל, לשתף או להפיץ את הקוד, כולו או חלקו.
 *
 * ---- פרטי הרישיון (למוכרת: למלא בכל מכירה, שתי שורות) ----
 */
var LICENSE_ID    = 'UNLICENSED';          // למשל: MG-2026-0001
var LICENSE_EMAIL = 'לא הוגדר';            // האימייל של הקונה, כפי שמופיע ברשומת המכירה

var TAB_GUESTS   = 'אורחים';
var TAB_SETTINGS = 'הגדרות';
var TAB_LOG      = 'יומן';

var COL = {
  NAME: 1, PHONE: 2, SIDE: 3, GROUP: 4, INVITED: 5,
  STATUS: 6, COMING: 7, ANSWERED: 8,
  CODE: 9, LINK: 10,
  WA1: 11, WA2: 12, WA3: 13,
  SENT1: 14, SENT2: 15, SENT3: 16
};

var HEADERS = [
  'שם האורח', 'טלפון', 'צד', 'קבוצה', 'כמה מוזמנים',
  'סטטוס', 'כמה מגיעים', 'מתי ענו',
  'קוד', 'הקישור האישי',
  'שליחה 1 · הזמנה', 'שליחה 2 · תזכורת', 'שליחה 3 · יום האירוע',
  'נשלח 1', 'נשלח 2', 'נשלח 3'
];

var DEFAULT_SETTINGS = [
  ['כתובת האתר', '', 'מתמלא לבד. אין צורך לגעת.'],
  ['כותרת', 'האירוע שלנו', 'מופיע בראש דף האורח'],
  ['תאריך', '', 'לכתוב כמו שרוצים שהאורחים יראו. למשל: יום שלישי, 12.5'],
  ['שעה', '', 'למשל: 09:30'],
  ['שם המקום', '', ''],
  ['כתובת', '', 'כתובת מלאה, כולל עיר'],
  ['קישור לתמונת ההזמנה', '', 'לא חובה. מופיע בראש דף האורח'],
  ['הרשימה סגורה', 'לא', 'כן / לא. כשכתוב כן, האורחים רואים את הפרטים אבל לא יכולים לשנות תשובה'],
  ['קישור קבוצתי פעיל', 'כן', 'כן / לא. קישור אחד שאפשר לשלוח לקבוצת וואטסאפ, וכל אחד מוסיף את עצמו'],
  ['הודעה 1 · הזמנה',
   'שלום {{שם}} 🤍\nאנחנו שמחים להזמין אתכם לאירוע שלנו!\n📅 {{תאריך}} בשעה {{שעה}}\n📍 {{מקום}}, {{כתובת}}\nנשמח מאוד אם תוכלו לאשר הגעה בקישור האישי:\n{{קישור}}\nמחכים לראותכם 🙏',
   'משתנים אפשריים: {{שם}} {{תאריך}} {{שעה}} {{מקום}} {{כתובת}} {{קישור}} {{כמה_מוזמנים}}'],
  ['הודעה 2 · תזכורת',
   'היי {{שם}}, מזכירים באהבה 🤍\nעדיין לא קיבלנו את אישור ההגעה שלכם לאירוע ב{{תאריך}}.\nנשמח לדעת אם תוכלו להגיע, זה לוקח 10 שניות:\n{{קישור}}\nתודה רבה!',
   ''],
  ['הודעה 3 · יום האירוע',
   'בוקר טוב {{שם}}! היום זה קורה 🎉\nהאירוע מתחיל בשעה {{שעה}} ב{{מקום}}.\nהכתובת: {{כתובת}}\nנתראה בקרוב 🤍',
   '']
];

/* ============================ תפריט ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🥂 ניהול אירוע')
    .addItem('1. הכנת הגיליון', 'setupSheets')
    .addSeparator()
    .addItem('2. יצירת קישורים לאורחים חדשים', 'buildLinks')
    .addSeparator()
    .addItem('📤 ייצוא לרשימת השליחה', 'exportSendLinks')
    .addSeparator()
    .addItem('📣 קישור קבוצתי להעתקה', 'showGroupLink')
    .addSeparator()
    .addItem('🩺 בדיקת מערכת', 'healthCheck')
    .addItem('ℹ️ פרטי רישיון', 'showLicense')
    .addSeparator()
    .addItem('סגירת רשימת האישורים עכשיו', 'lockNow')
    .addItem('פתיחת הרשימה מחדש', 'unlockNow')
    .addToUi();
}

/* ========================= פרטי רישיון ========================= */

function showLicense() {
  var ui = SpreadsheetApp.getUi();
  var who = '';
  try { who = Session.getEffectiveUser().getEmail() || ''; } catch (e) { who = ''; }
  ui.alert('פרטי רישיון',
    'מגיעים · רישיון שימוש אישי\n\n' +
    'מספר רישיון: ' + LICENSE_ID + '\n' +
    'הונפק ל: ' + LICENSE_EMAIL + '\n' +
    (who ? 'מותקן תחת החשבון: ' + who + '\n' : '') +
    '\nהשימוש אישי בלבד, באירועים של בעל הרישיון.\n' +
    'אין להעביר, למכור, להעתיק או להפיץ את הקוד.\n\n' +
    '© כל הזכויות שמורות ליוצרת המערכת.',
    ui.ButtonSet.OK);
}

/* ======================== הכנת הגיליון ========================= */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTheme && null;

  var st = ss.getSheetByName(TAB_SETTINGS) || ss.insertSheet(TAB_SETTINGS, 0);
  if (st.getLastRow() < 2) {
    st.clear();
    st.getRange(1, 1, 1, 3).setValues([['מה', 'הערך', 'הסבר']])
      .setFontWeight('bold').setBackground('#efe9dd');
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    st.setColumnWidth(1, 190).setColumnWidth(2, 430).setColumnWidth(3, 380);
    st.getRange(2, 2, DEFAULT_SETTINGS.length, 1).setWrap(true);
    st.getRange(2, 3, DEFAULT_SETTINGS.length, 1).setWrap(true).setFontColor('#7a7166');
  }
  ensureSettingRows(st);
  st.setRightToLeft(true);
  st.setFrozenRows(1);

  var g = ss.getSheetByName(TAB_GUESTS) || ss.insertSheet(TAB_GUESTS, 1);
  if (g.getLastRow() < 1 || g.getRange(1, 1).getValue() !== HEADERS[0]) {
    g.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#efe9dd');
  }
  g.setRightToLeft(true);
  g.setFrozenRows(1);
  g.setColumnWidth(COL.NAME, 190);
  g.setColumnWidth(COL.PHONE, 130);
  g.setColumnWidth(COL.LINK, 260);
  [COL.WA1, COL.WA2, COL.WA3].forEach(function (c) { g.setColumnWidth(c, 150); });
  g.getRange(2, COL.PHONE, 998, 1).setNumberFormat('@');

  var sideRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['צד האב', 'צד האם', 'משותף'], true).setAllowInvalid(true).build();
  g.getRange(2, COL.SIDE, 998, 1).setDataValidation(sideRule);
  g.getRange(2, COL.SENT1, 998, 3).insertCheckboxes();

  var rules = [];
  function tint(text, bg, fg) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text)
      .setBackground(bg).setFontColor(fg)
      .setRanges([g.getRange(2, COL.STATUS, 998, 1)]).build());
  }
  tint('אישרו', '#dcf3e3', '#14653a');
  tint('לא מגיעים', '#fbdedd', '#8f1d1a');
  tint('ממתינים', '#f0efec', '#5c574f');
  g.setConditionalFormatRules(rules);

  st.getRange(1, 5).setValue('מצב האישורים').setFontWeight('bold').setBackground('#efe9dd');
  st.getRange(2, 5, 6, 2).setValues([
    ['אורחים ברשימה',   "=COUNTA('" + TAB_GUESTS + "'!A2:A)"],
    ['אישרו',           "=COUNTIF('" + TAB_GUESTS + "'!F2:F;\"אישרו\")"],
    ['לא מגיעים',       "=COUNTIF('" + TAB_GUESTS + "'!F2:F;\"לא מגיעים\")"],
    ['טרם ענו',         "=COUNTIF('" + TAB_GUESTS + "'!F2:F;\"ממתינים\")"],
    ['סה״כ אנשים שמגיעים', "=SUM('" + TAB_GUESTS + "'!G2:G)"],
    ['אחוז שענו',       "=IFERROR(TEXT((F3+F4)/F2;\"0%\");\"—\")"]
  ]);
  st.getRange(2, 5, 6, 1).setFontColor('#7a7166');
  st.getRange(2, 6, 6, 1).setFontWeight('bold').setHorizontalAlignment('left');
  st.setColumnWidth(5, 170).setColumnWidth(6, 90);

  var log = ss.getSheetByName(TAB_LOG) || ss.insertSheet(TAB_LOG);
  if (log.getLastRow() < 1) {
    log.getRange(1, 1, 1, 4).setValues([['מתי', 'אורח', 'מה קרה', 'פרטים']])
      .setFontWeight('bold').setBackground('#efe9dd');
  }
  log.setRightToLeft(true);
  log.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'הגיליון מוכן.\n\nעכשיו: מלאי את לשונית "הגדרות", ואז הזיני אורחים בלשונית "אורחים".\n' +
    'אחרי שהאתר מפורסם, הדביקי את הכתובת שלו בשורה הראשונה בהגדרות ולחצי על "יצירת קישורים".');
}

/* ========================= קריאת הגדרות ======================== */

function readSettings() {
  var st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_SETTINGS);
  var out = {};
  if (!st) return out;
  var rows = Math.max(st.getLastRow() - 1, 1);
  var keys = st.getRange(2, 1, rows, 1).getValues();
  var vals = st.getRange(2, 2, rows, 1).getDisplayValues();
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0] || '').trim();
    if (!k) continue;
    var v = String(vals[i][0] == null ? '' : vals[i][0]).trim();
    if (k === 'שעה') v = v.replace(/^(\d{1,2}):(\d{2}):\d{2}$/, '$1:$2');
    out[k] = v;
  }
  return out;
}

/* ============== זיהוי עצמי של כתובת האתר ============== */

function detectWebAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  return url ? url.replace(/\/dev$/, '/exec').split('?')[0] : '';
}

function getWebAppUrl() {
  var auto = detectWebAppUrl();
  if (auto) return auto;
  var manual = String(readSettings()['כתובת האתר'] || '').trim();
  return manual ? manual.split('?')[0] : '';
}

function writeSetting(key, value) {
  var st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_SETTINGS);
  if (!st || st.getLastRow() < 2) return;
  var v = st.getRange(2, 1, st.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) { st.getRange(i + 2, 2).setValue(value); return; }
  }
}

function ensureSettingRows(st) {
  var existing = {};
  if (st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0]) existing[String(r[0]).trim()] = true; });
  }
  var add = DEFAULT_SETTINGS.filter(function (r) { return !existing[r[0]]; });
  if (!add.length) return;
  st.getRange(st.getLastRow() + 1, 1, add.length, 3).setValues(add);
  st.getRange(st.getLastRow() - add.length + 1, 2, add.length, 2).setWrap(true);
}

/* =================== קישור קבוצתי ==================== */

function groupLinkOn() {
  return String(readSettings()['קישור קבוצתי פעיל'] || 'כן').trim() !== 'לא';
}

function showGroupLink() {
  var ui = SpreadsheetApp.getUi();
  var url = getWebAppUrl();
  if (!url) { ui.alert('האתר עוד לא פורסם, אז אין עדיין קישור קבוצתי.'); return; }
  if (!groupLinkOn()) {
    ui.alert('הקישור הקבוצתי כבוי.\n\nבלשונית "הגדרות", בשורה "קישור קבוצתי פעיל", שני ל"כן".');
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Assistant,Arial,sans-serif;font-size:15px;line-height:1.6;padding:6px">' +
    '<p>זה הקישור לשליחה בקבוצות. כל מי שלוחץ מזין שם וטלפון ומצטרף לרשימה.</p>' +
    '<p style="margin:14px 0"><b>קישור כללי</b><br>' +
    '<input style="width:100%;padding:9px;font-size:13px;direction:ltr" ' +
    'value="' + url + '?join=1" onclick="this.select()" readonly></p>' +
    '<p style="margin:14px 0"><b>עם שם קבוצה</b>, כדי לדעת מאיפה הגיעו:<br>' +
    '<input style="width:100%;padding:9px;font-size:13px;direction:ltr" ' +
    'value="' + url + '?join=1&k=' + encodeURIComponent('מילואים') + '" onclick="this.select()" readonly></p>' +
    '<p style="color:#6b7683">החליפי את המילה שבסוף בשם הקבוצה שלך. היא תופיע בעמודת "קבוצה" בגיליון.</p>' +
    '<p style="color:#6b7683">לחיצה על התיבה בוחרת את כל הטקסט, ואז העתקה רגילה.</p></div>')
    .setWidth(520).setHeight(330);
  ui.showModalDialog(html, 'קישור קבוצתי');
}

function joinRsvp(name, phoneRaw, attending, count, groupLabel) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!groupLinkOn()) return { ok: false, message: 'ההרשמה הקבוצתית סגורה.' };
    if (isLocked())     return { ok: false, message: 'רשימת האורחים נסגרה. אנא צרו קשר עם המשפחה.' };

    name = String(name == null ? '' : name).trim().replace(/\s+/g, ' ');
    if (name.length < 2)  return { ok: false, message: 'יש להזין שם מלא.' };
    if (name.length > 60) return { ok: false, message: 'השם ארוך מדי.' };

    var phone = normalizePhone(phoneRaw);
    if (!/^972\d{8,9}$/.test(phone)) return { ok: false, message: 'מספר הטלפון אינו תקין. למשל 0501234567' };

    var n = 0;
    if (attending === true) {
      var raw = String(count == null ? '' : count).trim();
      if (!/^\d+$/.test(raw)) return { ok: false, message: 'יש להזין מספר, ספרות בלבד.' };
      n = parseInt(raw, 10);
      if (!(n >= 1))  return { ok: false, message: 'יש להזין לפחות אורח אחד.' };
      if (n > 20)     return { ok: false, message: 'מקסימום 20 אורחים. ליותר מזה, אנא צרו קשר ישירות.' };
    } else if (attending !== false) {
      return { ok: false, message: 'לא התקבלה תשובה תקינה.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var g  = ss.getSheetByName(TAB_GUESTS);
    var label = String(groupLabel || '').trim().slice(0, 40) || 'קישור קבוצתי';

    // כבר קיים עם אותו טלפון? מעדכנים את השורה הקיימת ולא יוצרים כפילות
    var row = 0, existingName = '';
    if (g.getLastRow() > 1) {
      var rows = g.getRange(2, 1, g.getLastRow() - 1, HEADERS.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (normalizePhone(rows[i][COL.PHONE - 1]) === phone) {
          row = i + 2; existingName = String(rows[i][COL.NAME - 1] || ''); break;
        }
      }
    }

    if (!row) {
      row = g.getLastRow() + 1;
      g.getRange(row, COL.NAME).setValue(name);
      g.getRange(row, COL.PHONE).setValue(phone);
      g.getRange(row, COL.SIDE).setValue('משותף');
      g.getRange(row, COL.GROUP).setValue(label);
      g.getRange(row, COL.CODE).setValue(makeCode());
      g.getRange(row, COL.INVITED).setValue(attending ? n : 1);
      g.getRange(row, COL.LINK).setValue(getWebAppUrl() + '?g=' + g.getRange(row, COL.CODE).getValue());
    }

    g.getRange(row, COL.STATUS).setValue(attending ? 'אישרו' : 'לא מגיעים');
    g.getRange(row, COL.COMING).setValue(attending ? n : 0);
    g.getRange(row, COL.ANSWERED).setValue(new Date());

    writeLog(name, existingName ? 'עדכון דרך קישור קבוצתי' : 'הצטרפות דרך קישור קבוצתי',
             label + ' · ' + (attending ? n + ' אורחים' : 'לא מגיעים') +
             (existingName && existingName !== name ? ' · היה רשום כ' + existingName : ''));

    return { ok: true, attending: attending, count: attending ? n : 0, name: name };
  } finally {
    lock.releaseLock();
  }
}

/* ============ ייצוא לכלי רשימת השליחה ============= */

function exportSendLinks() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('ייצוא לרשימת השליחה',
    'איזה סבב לייצא?\n\n' +
    '1 = הזמנה, לכל מי שיש לו טלפון\n' +
    '2 = תזכורת, רק למי שטרם ענה\n' +
    '3 = יום האירוע, רק למי שאישר\n\n' +
    'הקלידי 1, 2 או 3:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var round = parseInt(String(res.getResponseText()).trim(), 10);
  if ([1, 2, 3].indexOf(round) === -1) {
    ui.alert('יש להקליד 1, 2 או 3.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var g = ss.getSheetByName(TAB_GUESTS);
  if (!g || g.getLastRow() < 2) { ui.alert('אין אורחים ברשימה.'); return; }

  var s = readSettings();
  var base = getWebAppUrl();
  if (!base) { ui.alert('האתר עוד לא פורסם.'); return; }

  var msgs = [s['הודעה 1 · הזמנה'], s['הודעה 2 · תזכורת'], s['הודעה 3 · יום האירוע']];
  var rows = g.getRange(2, 1, g.getLastRow() - 1, HEADERS.length).getValues();
  var out = [], skipped = 0;

  rows.forEach(function (r) {
    var name = String(r[COL.NAME - 1] || '').trim();
    var phone = normalizePhone(r[COL.PHONE - 1]);
    var code = String(r[COL.CODE - 1] || '').trim();
    var status = String(r[COL.STATUS - 1] || 'ממתינים').trim();
    if (!name || !phone || !code) { if (name) skipped++; return; }
    if (round === 2 && status !== 'ממתינים') return;
    if (round === 3 && status !== 'אישרו') return;

    var body = fillTemplate(msgs[round - 1],
                            { name: name, invited: r[COL.INVITED - 1] || 1 },
                            s, base + '?g=' + code);
    out.push(name + ', https://web.whatsapp.com/send?phone=' + phone +
             '&text=' + encodeURIComponent(body));
  });

  if (!out.length) { ui.alert('אין נמענים מתאימים לסבב הזה.'); return; }

  var note = out.length + ' נמענים' + (skipped ? ' · ' + skipped + ' דולגו (בלי טלפון או בלי קוד)' : '');
  var html = HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font-family:Assistant,Arial,sans-serif;font-size:14px;padding:6px">' +
    '<p style="margin:0 0 8px">' + note + '. סמני הכל, העתיקי, והדביקי בקובץ <b>רשימת שליחה</b>.</p>' +
    '<textarea readonly onclick="this.select()" style="width:100%;height:300px;' +
    'direction:ltr;text-align:left;font-size:11px;padding:8px;border:1px solid #D7DEE6;border-radius:8px">' +
    out.join('\n').replace(/</g, '&lt;') + '</textarea></div>')
    .setWidth(620).setHeight(420);
  ui.showModalDialog(html, 'ייצוא סבב ' + round);
}

/* ==================== בדיקת מערכת ==================== */

function healthCheck() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [], fatal = [];

  function ok(t)  { lines.push('✓  ' + t); }
  function bad(t, fix) { lines.push('✗  ' + t); fatal.push(fix); }

  // 1. לשוניות
  var missing = [TAB_SETTINGS, TAB_GUESTS, TAB_LOG].filter(function (n) { return !ss.getSheetByName(n); });
  if (missing.length) bad('חסרות לשוניות: ' + missing.join(', '),
                          'הריצי "1. הכנת הגיליון" מהתפריט.');
  else ok('שלוש הלשוניות קיימות');

  // 2. חיבור לגיליון הנכון
  ok('הקוד מחובר לגיליון "' + ss.getName() + '"');
  ok('רישיון: ' + LICENSE_ID);

  // 3. כתובת האתר
  var url = detectWebAppUrl();
  if (!url) {
    bad('האתר עוד לא פורסם',
        'לחצי "לפריסה" ← "פריסה חדשה" ← אפליקציית אינטרנט ← גישה לכל אחד ← פריסה.');
  } else {
    ok('האתר פורסם');
    writeSetting('כתובת האתר', url);

    // 4. בדיקה אמיתית מקצה לקצה
    try {
      var res = UrlFetchApp.fetch(url + '?g=__healthcheck__',
                                  { muteHttpExceptions: true, followRedirects: true });
      var body = res.getContentText();
      if (body.indexOf('Script function not found') > -1) {
        bad('האתר מפרסם קוד ישן או ריק',
            'לחצי "לפריסה" ← "ניהול פריסות" ← עיפרון ← בשדה גרסה לבחור "גרסה חדשה" ← פריסה. אם זה לא עוזר, צרי פריסה חדשה.');
      } else if (body.indexOf('הקישור אינו תקין') > -1 || body.indexOf('אישור הגעה') > -1) {
        ok('הדף נטען ועובד');
      } else if (body.indexOf('accounts.google.com') > -1 || res.getResponseCode() === 401) {
        bad('האתר דורש התחברות, האורחים לא יוכלו להיכנס',
            'לחצי "לפריסה" ← "ניהול פריסות" ← עיפרון ← "למי יש גישה" לשנות ל"כל אחד" ← פריסה.');
      } else {
        bad('הדף החזיר תשובה לא מוכרת (קוד ' + res.getResponseCode() + ')',
            'צרי פריסה חדשה: "לפריסה" ← "פריסה חדשה" ← אפליקציית אינטרנט ← כל אחד.');
      }
    } catch (e) {
      lines.push('•  לא הצלחתי לבדוק את הדף מכאן. פתחי את הכתובת בדפדפן ידנית.');
    }
  }

  // 5. הגדרות חובה
  var s = readSettings();
  var need = ['כותרת', 'תאריך', 'שעה', 'שם המקום', 'כתובת'];
  var empty = need.filter(function (k) { return !String(s[k] || '').trim(); });
  if (empty.length) lines.push('•  שדות ריקים בהגדרות: ' + empty.join(', '));
  else ok('פרטי האירוע מלאים');

  // 6. אורחים
  var g = ss.getSheetByName(TAB_GUESTS);
  if (g && g.getLastRow() > 1) {
    var rows = g.getRange(2, 1, g.getLastRow() - 1, HEADERS.length).getValues();
    var withName = 0, noPhone = 0, noCode = 0;
    rows.forEach(function (r) {
      if (!String(r[COL.NAME - 1] || '').trim()) return;
      withName++;
      if (!normalizePhone(r[COL.PHONE - 1])) noPhone++;
      else if (!String(r[COL.CODE - 1] || '').trim()) noCode++;
    });
    ok(withName + ' אורחים ברשימה');
    if (noPhone) lines.push('•  ' + noPhone + ' בלי טלפון, הם ידולגו בשליחה');
    if (noCode) bad(noCode + ' אורחים עדיין בלי קישור',
                    'הריצי "2. יצירת קישורים לאורחים חדשים".');
  } else {
    lines.push('•  אין עדיין אורחים ברשימה');
  }

  var msg = lines.join('\n');
  msg += fatal.length
    ? '\n\n⚠️ מה לעשות עכשיו:\n' + fatal.map(function (f, i) { return (i + 1) + '. ' + f; }).join('\n')
    : '\n\n🎉 הכל תקין. אפשר לשלוח.';
  ui.alert('בדיקת מערכת', msg, ui.ButtonSet.OK);
}

function isLocked() {
  return String(readSettings()['הרשימה סגורה'] || '').trim() === 'כן';
}

function setLock(value) {
  var st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_SETTINGS);
  var v = st.getRange(2, 1, st.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'הרשימה סגורה') { st.getRange(i + 2, 2).setValue(value); return; }
  }
}

function lockNow() {
  setLock('כן');
  SpreadsheetApp.getUi().alert('רשימת האישורים סגורה. האורחים עדיין רואים את פרטי האירוע, אבל לא יכולים לשנות תשובה.');
}

function unlockNow() {
  setLock('לא');
  SpreadsheetApp.getUi().alert('הרשימה פתוחה שוב.');
}

/* ===================== טלפונים וקישורים ======================== */

function normalizePhone(raw) {
  var s = String(raw == null ? '' : raw).replace(/[\u200e\u200f]/g, '').replace(/[^\d+]/g, '');
  if (!s) return '';
  if (s.indexOf('+') === 0) s = s.slice(1);
  if (s.indexOf('00') === 0) s = s.slice(2);
  if (s.indexOf('972') === 0) return s;
  if (s.indexOf('0') === 0) return '972' + s.slice(1);
  if (s.length === 9 && s.charAt(0) === '5') return '972' + s;
  if (s.length >= 8 && s.length <= 10) return '972' + s;
  return s;
}

function makeCode() {
  var a = 'abcdefghjkmnpqrstuvwxyz23456789', s = '';
  for (var i = 0; i < 10; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}

function fillTemplate(text, guest, settings, link) {
  return String(text || '')
    .replace(/\{\{שם\}\}/g, guest.name)
    .replace(/\{\{תאריך\}\}/g, settings['תאריך'] || '')
    .replace(/\{\{שעה\}\}/g, settings['שעה'] || '')
    .replace(/\{\{מקום\}\}/g, settings['שם המקום'] || '')
    .replace(/\{\{כתובת\}\}/g, settings['כתובת'] || '')
    .replace(/\{\{כמה_מוזמנים\}\}/g, guest.invited || 1)
    .replace(/\{\{קישור\}\}/g, link || '');
}

function buildLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var g = ss.getSheetByName(TAB_GUESTS);
  var s = readSettings();
  var base = getWebAppUrl();

  if (!base) {
    SpreadsheetApp.getUi().alert(
      'האתר עוד לא פורסם, ולכן אין עדיין קישורים.\n\n' +
      'לחצי למעלה על "לפריסה" ← "פריסה חדשה" ← גלגל השיניים ← אפליקציית אינטרנט ← ' +
      '"למי יש גישה" = כל אחד ← פריסה.\n\nואז חזרי לכאן ולחצי שוב.');
    return;
  }
  writeSetting('כתובת האתר', base);

  var last = g.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('אין עדיין אורחים ברשימה.'); return; }

  var rows = g.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var seen = {}, added = 0, dupes = [], refreshed = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], rowNum = i + 2;
    var name = String(r[COL.NAME - 1] || '').trim();
    var phoneRaw = r[COL.PHONE - 1];
    if (!name && !phoneRaw) continue;

    var phone = normalizePhone(phoneRaw);
    if (!phone) continue;

    if (seen[phone]) { dupes.push(name + ' (שורה ' + rowNum + ')'); continue; }
    seen[phone] = true;

    g.getRange(rowNum, COL.PHONE).setValue(phone);
    if (!r[COL.SIDE - 1]) g.getRange(rowNum, COL.SIDE).setValue('משותף');
    if (!r[COL.INVITED - 1]) g.getRange(rowNum, COL.INVITED).setValue(1);
    if (!r[COL.STATUS - 1]) g.getRange(rowNum, COL.STATUS).setValue('ממתינים');

    var code = String(r[COL.CODE - 1] || '').trim();
    if (!code) { code = makeCode(); g.getRange(rowNum, COL.CODE).setValue(code); added++; }
    else refreshed++;

    var link = base + '?g=' + code;
    g.getRange(rowNum, COL.LINK).setValue(link);

    var guest = { name: name, invited: r[COL.INVITED - 1] || 1 };
    var msgs = [s['הודעה 1 · הזמנה'], s['הודעה 2 · תזכורת'], s['הודעה 3 · יום האירוע']];
    var labels = ['שליחת הזמנה', 'שליחת תזכורת', 'שליחת הודעת יום האירוע'];

    for (var k = 0; k < 3; k++) {
      var body = fillTemplate(msgs[k], guest, s, link);
      var wa = 'https://web.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(body);
      g.getRange(rowNum, COL.WA1 + k)
        .setFormula('=HYPERLINK("' + wa.replace(/"/g, '%22') + '";"' + labels[k] + '")');
    }
  }

  var msg = 'מוכן.\n\nנוצרו קישורים ל-' + added + ' אורחים חדשים.';
  if (refreshed) msg += '\nעודכנו הודעות ל-' + refreshed + ' אורחים קיימים.';
  if (dupes.length) {
    msg += '\n\nשימי לב, המספרים האלה כבר מופיעים אצל אורח אחר ולכן דילגתי עליהם:\n' +
           dupes.slice(0, 8).join('\n') +
           '\n\nאם זה זוג עם מספר משותף, מחקי שורה אחת ועדכני "כמה מוזמנים" בשורה שנשארה.';
  }
  SpreadsheetApp.getUi().alert(msg);
}

/* ========================== דף האורח =========================== */

function doGet(e) {
  var code = (e && e.parameter && e.parameter.g) ? String(e.parameter.g).trim() : '';
  var s = readSettings();
  var data;

  var isJoin = !!(e && e.parameter && e.parameter.join);

  if (isJoin) {
    data = groupLinkOn()
      ? { mode: 'join', group: (e.parameter.k ? String(e.parameter.k) : ''), locked: isLocked() }
      : { error: 'ההרשמה נסגרה. אנא פנו למשפחה.' };
  } else if (!code) {
    data = { error: 'הקישור אינו תקין. אנא פנו למשפחה לקבלת קישור חדש.' };
  } else {
    var found = findGuestByCode(code);
    if (!found) {
      data = { error: 'הקישור אינו תקין. אנא פנו למשפחה לקבלת קישור חדש.' };
    } else {
      data = {
        code: code,
        name: found.name,
        invited: found.invited,
        status: found.status,
        coming: found.coming,
        locked: isLocked(),
        title: s['כותרת'] || 'האירוע שלנו',
        date: s['תאריך'] || '',
        time: s['שעה'] || '',
        venue: s['שם המקום'] || '',
        address: s['כתובת'] || '',
        image: s['קישור לתמונת ההזמנה'] || '',
        waze: buildWaze(s['כתובת'] || ''),
        gmaps: buildGmaps(s['כתובת'] || '')
      };
    }
  }

  if (!data.error) {
    data.title   = data.title   || s['כותרת'] || 'אישור הגעה';
    data.date    = data.date    || s['תאריך'] || '';
    data.time    = data.time    || s['שעה'] || '';
    data.venue   = data.venue   || s['שם המקום'] || '';
    data.address = data.address || s['כתובת'] || '';
    data.image   = data.image   || s['קישור לתמונת ההזמנה'] || '';
    data.waze    = data.waze    || buildWaze(s['כתובת'] || '');
    data.gmaps   = data.gmaps   || buildGmaps(s['כתובת'] || '');
  }

  var t = HtmlService.createTemplateFromFile('rsvp');
  t.payload = JSON.stringify(data);
  return t.evaluate()
    .setTitle(data.title || 'אישור הגעה')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildWaze(address) {
  if (!address) return '';
  return 'https://waze.com/ul?q=' + encodeURIComponent(address) + '&navigate=yes';
}

function buildGmaps(address) {
  if (!address) return '';
  return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(address);
}

function findGuestByCode(code) {
  var g = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
  var last = g.getLastRow();
  if (last < 2) return null;
  var rows = g.getRange(2, 1, last - 1, HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.CODE - 1] || '').trim() === code) {
      return {
        row: i + 2,
        name: String(rows[i][COL.NAME - 1] || ''),
        invited: Number(rows[i][COL.INVITED - 1] || 1),
        status: String(rows[i][COL.STATUS - 1] || 'ממתינים'),
        coming: rows[i][COL.COMING - 1] === '' ? null : Number(rows[i][COL.COMING - 1])
      };
    }
  }
  return null;
}

/* ================== שמירת תשובה, כולל בדיקות ================== */

function submitRsvp(code, attending, count) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (isLocked()) return { ok: false, message: 'רשימת האורחים נסגרה. לשינויים אנא צרו קשר עם המשפחה.' };

    var guest = findGuestByCode(String(code || '').trim());
    if (!guest) return { ok: false, message: 'הקישור אינו תקין.' };

    var n = 0;
    if (attending === true) {
      var raw = String(count == null ? '' : count).trim();
      if (!/^\d+$/.test(raw)) return { ok: false, message: 'יש להזין מספר, ספרות בלבד.' };
      n = parseInt(raw, 10);
      if (!(n >= 1)) return { ok: false, message: 'יש להזין לפחות אורח אחד.' };
      if (n > 20) return { ok: false, message: 'מקסימום 20 אורחים. ליותר מזה, אנא צרו קשר ישירות.' };
    } else if (attending !== false) {
      return { ok: false, message: 'לא התקבלה תשובה תקינה.' };
    }

    var g = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
    var before = guest.status + (guest.coming == null ? '' : ' / ' + guest.coming);
    g.getRange(guest.row, COL.STATUS).setValue(attending ? 'אישרו' : 'לא מגיעים');
    g.getRange(guest.row, COL.COMING).setValue(attending ? n : 0);
    g.getRange(guest.row, COL.ANSWERED).setValue(new Date());

    writeLog(guest.name, attending ? 'אישרו הגעה' : 'לא מגיעים',
             'קודם: ' + before + ' · עכשיו: ' + (attending ? n : 0));

    return { ok: true, attending: attending, count: attending ? n : 0 };
  } finally {
    lock.releaseLock();
  }
}

function writeLog(name, action, details) {
  try {
    var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_LOG);
    if (log) log.appendRow([new Date(), name, action, details]);
  } catch (err) { /* היומן לא יפיל את השמירה */ }
}
