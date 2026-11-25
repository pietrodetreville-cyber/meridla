// ---------- doklady (paragony): foto, zápis do listu Doklady, odeslání účetní ----------
// Kdo má v listu Lidé ve sloupci Doklady "ano", smí fotit doklady (vidí jen své). Kdo má "ano" ve sloupci Odesílá účetní,
// vidí doklady všech a smí vybrané odeslat účetní z aplikace. V tabulce se vybírá zaškrtnutím ve sloupci Vybrat k odeslání.
// Odeslané doklady se zapíšou do listu Faktury (DUZP = datum dokladu, lze přepsat) a list se seřadí podle DUZP.
// Fotky se ukládají na Disk do složky Doklady (SVJ_app/OSTRE nebo TESTOVACI). E-mail účetní je ve vlastnostech skriptu (ACCOUNTANT),
// nastavuje se z menu Administrace → Doklady → Nastavit e-mail účetní.

const SHEET_DOCS = 'Doklady', DOCS_FOLDER = 'Doklady';
// Doklady a Faktury jsou v samostatné Google tabulce „Doklady SVJ Od Vysoké“ ve složce SVJ_app/OSTRE (TESTOVACI).
// Vytvoří se sama při prvním použití; její ID je ve vlastnosti skriptu DOCS_SS. Při otevření má vlastní menu Doklady.
function docsFileName_() { return 'Doklady SVJ Od Vysoké' + (props_().getProperty('ENV') === 'test' ? ' – TEST' : ''); }
function docsSSIfExists_() { const id = props_().getProperty('DOCS_SS'); if (!id) return null; try { return SpreadsheetApp.openById(id); } catch (e) { return null; } }
function docsSS_() {
  const have = docsSSIfExists_(); if (have) return have;
  const ss = SpreadsheetApp.create(docsFileName_());
  try { DriveApp.getFileById(ss.getId()).moveTo(driveRoot_()); } catch (e) {}
  props_().setProperty('DOCS_SS', ss.getId());
  // první list = Doklady; dřívější listy Doklady a Faktury z tabulky měřidel se zkopírují (původní zůstanou s příponou „(přesunuto)“)
  const first = ss.getSheets()[0]; first.setName(SHEET_DOCS);
  const old = SpreadsheetApp.getActive();
  [SHEET_DOCS, 'Faktury'].forEach(n => {
    const src = old && old.getSheetByName(n); if (!src || src.getLastRow() < 1) return;
    const dst = ss.getSheetByName(n) || ss.insertSheet(n), v = src.getDataRange().getValues();
    dst.getRange(1, 1, v.length, v[0].length).setValues(v); dst.setFrozenRows(1); dst.getRange(1, 1, 1, v[0].length).setFontWeight('bold');
    if (n === SHEET_DOCS) { dst.getRange('B:B').setNumberFormat('d. M. yyyy'); dst.getRange('C:C').setNumberFormat('#,##0.00'); if (v.length > 1 && v[0].length >= COL_DSEL) dst.getRange(2, COL_DSEL, v.length - 1, 1).insertCheckboxes(); }
    else { dst.getRange('A:B').setNumberFormat('d. M. yyyy'); dst.getRange('C:C').setNumberFormat('#,##0.00'); }
    try { src.setName(n + ' (přesunuto)'); } catch (e) {}
  });
  try {
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'docsOnOpen' && t.getTriggerSourceId() === ss.getId()))
      ScriptApp.newTrigger('docsOnOpen').forSpreadsheet(ss).onOpen().create();
  } catch (e) { Logger.log('Menu pro soubor dokladů se nepodařilo zapnout: ' + e.message); }
  return ss;
}
// menu v souboru dokladů (instalovaný spouštěč při otevření)
function docsOnOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Doklady')
    .addItem('Zaregistrovat došlou fakturu (PDF)', 'invoiceDialog')
    .addItem('Odeslat vybrané (zaškrtnuté) doklady účetní', 'sendReceipts')
    .addItem('Obnovit přehled pro výbor', 'refreshDash')
    .addItem('Seřadit Faktury podle DUZP', 'sortInvoices')
    .addSeparator()
    .addSubMenu(ui.createMenu('Nastavení').addItem('E-mail účetní', 'setAccountant').addItem('Kopie e-mailu (výbor)', 'setAccountantCc').addItem('Kdo z výboru vidí doklady a přehled', 'setDashViewers'))
    .addToUi();
}
// z menu tabulky měřidel: odkaz na soubor dokladů
function openDocsFile() {
  const ss = docsSS_(), url = ss.getUrl();
  const html = HtmlService.createHtmlOutput('<p style="font:14px Arial">Doklady a faktury jsou v samostatné tabulce:</p><p style="font:14px Arial"><a href="' + url + '" target="_blank">' + ss.getName() + '</a></p>').setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Soubor dokladů');
}
const DOCS_HEAD = ['Zapsáno', 'Datum dokladu', 'Částka (Kč)', 'Obchod / dodavatel', 'Účel', 'Budova', 'Kdo', 'Fotka', 'Stav', 'ID', 'Poznámka', 'Odesláno účetní', 'Vybrat k odeslání', 'Dávka', 'Zpráva pro účetní', 'Číslo dokladu', 'Typ', 'Číslo faktury', 'DUZP', 'Splatnost'];
const COL_DSEL = 13, COL_DBATCH = 14, COL_DMSG = 15, COL_DNUM = 16, COL_DTYPE = 17, COL_DINVNO = 18, COL_DDUZP = 19, COL_DDUE = 20;
// fotka v buňce: krátký text (📷 číslo dokladu) s odkazem místo dlouhé adresy
function setPhotoLink_(sh, row, col, label, url) {
  try { sh.getRange(row, col).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).build()); }
  catch (e) { sh.getRange(row, col).setValue(url); }
}
// řádky listu Doklady od 2. řádku; ve sloupci Fotka (H) je vždy adresa fotky (z odkazu, nebo z textu u starých řádků)
function docRows_(sh, width) {
  const n = sh.getLastRow() - 1; if (n < 1) return [];
  const v = sh.getRange(2, 1, n, width).getValues();
  try { const rt = sh.getRange(2, 8, n, 1).getRichTextValues(); v.forEach((r, i) => { const u = rt[i] && rt[i][0] && rt[i][0].getLinkUrl(); if (u) r[7] = u; }); } catch (e) {}
  return v;
}
// dřívější řádky s dlouhou adresou ve sloupci se převedou na krátký odkaz
function linkifyPhotos_(sh, col, numCol) {
  const n = sh.getLastRow() - 1; if (n < 1 || sh.getLastColumn() < col) return;
  const v = sh.getRange(2, col, n, 1).getValues(), nums = numCol && sh.getLastColumn() >= numCol ? sh.getRange(2, numCol, n, 1).getValues() : null;
  v.forEach((r, i) => { const t = String(r[0] || ''); if (/^https?:\/\//.test(t)) setPhotoLink_(sh, i + 2, col, '📷 ' + (nums && String(nums[i][0] || '').trim() ? String(nums[i][0]).trim() : 'fotka'), t); });
}

// číslo dokladu v interní evidenci: rrrr-mm-nnn (rok a měsíc z data dokladu, nnn = pořadí v tom měsíci)
function docNumPrefix_(date) { return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM'); }
function nextDocNum_(sh, date, taken) {
  const pre = docNumPrefix_(date) + '-';
  let max = 0;
  const nums = (sh.getLastRow() > 1 && sh.getLastColumn() >= COL_DNUM ? sh.getRange(2, COL_DNUM, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []).concat(taken || []);
  nums.forEach(n => { if (n.indexOf(pre) === 0) { const k = Number(n.slice(pre.length)); if (k > max) max = k; } });
  return pre + String(max + 1).padStart(3, '0');
}
// dřívějším dokladům bez čísla se čísla přidělí podle data dokladu a jejich fotky se přejmenují
function numberOldDocs_(sh) {
  if (sh.getLastRow() < 2) return;
  const v = docRows_(sh, COL_DNUM), todo = [];
  v.forEach((r, i) => { if (r[9] && !String(r[COL_DNUM - 1] || '').trim()) todo.push({ n: i + 2, r: r }); });
  todo.sort((a, b) => new Date(a.r[1]) - new Date(b.r[1]) || new Date(a.r[0]) - new Date(b.r[0]));
  const taken = [];
  todo.forEach(x => {
    const d = Object.prototype.toString.call(x.r[1]) === '[object Date]' ? x.r[1] : new Date(x.r[1]);
    const num = nextDocNum_(sh, isNaN(d) ? new Date() : d, taken); taken.push(num);
    sh.getRange(x.n, COL_DNUM).setValue(num);
    try { DriveApp.getFileById(driveId_(x.r[7])).setName(num + '.jpg'); } catch (e) {}
  });
}
const DOCS_MAX_MAIL = 15; // nejvýš dokladů v jednom e-mailu (velikost příloh)

function docsSheet_() {
  const ss = docsSS_();
  let sh = ss.getSheetByName(SHEET_DOCS);
  if (!sh) sh = ss.insertSheet(SHEET_DOCS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(DOCS_HEAD); sh.setFrozenRows(1); sh.getRange(1, 1, 1, DOCS_HEAD.length).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('d. M. yyyy'); sh.getRange('C:C').setNumberFormat('#,##0.00');
  }
  if (sh.getLastColumn() < COL_DDUE || !sh.getRange(1, COL_DDUE).getValue()) { sh.getRange('S:T').setNumberFormat('d. M. yyyy'); }
  [[COL_DBATCH, 'Dávka'], [COL_DMSG, 'Zpráva pro účetní'], [COL_DTYPE, 'Typ'], [COL_DINVNO, 'Číslo faktury'], [COL_DDUZP, 'DUZP'], [COL_DDUE, 'Splatnost']].forEach(([c, n]) => { if (sh.getLastColumn() < c || !sh.getRange(1, c).getValue()) sh.getRange(1, c).setValue(n).setFontWeight('bold'); });
  if (sh.getLastColumn() < COL_DNUM || !sh.getRange(1, COL_DNUM).getValue()) { sh.getRange(1, COL_DNUM).setValue('Číslo dokladu').setFontWeight('bold'); numberOldDocs_(sh); }
  linkifyPhotos_(sh, 8, COL_DNUM);
  if (sh.getLastColumn() < COL_DSEL || !sh.getRange(1, COL_DSEL).getValue()) {
    sh.getRange(1, COL_DSEL).setValue('Vybrat k odeslání').setFontWeight('bold');
    if (sh.getLastRow() > 1) sh.getRange(2, COL_DSEL, sh.getLastRow() - 1, 1).insertCheckboxes();
  }
  return sh;
}
const fmtKc_ = x => (Math.round((Number(x) || 0) * 100) / 100).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' Kč';
const docNum_ = v => { const t = String(v === undefined || v === null ? '' : v).replace(/\s/g, '').replace(',', '.'); return t === '' ? NaN : Number(t); };

// uložení dokladu z aplikace: fotka je povinná, datum a částka také, účel (na co) také
function receiptSave_(body, who) {
  const amount = Math.round(docNum_(body.amount) * 100) / 100, tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  if (!body.photo || !isFinite(amount) || !String(body.purpose || '').trim() || !/^\d{4}-\d{2}-\d{2}$/.test(String(body.date || ''))) return { ok: false, error: 'data' };
  let date = null; try { date = Utilities.parseDate(body.date + ' 12:00', tz, 'yyyy-MM-dd HH:mm'); } catch (e) {}
  if (!date || date.getFullYear() < 2000) return { ok: false, error: 'date' };
  if (date.getTime() > Date.now() + 36e5) return { ok: false, error: 'future' };
  let rid = String(body.clientId || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24);
  if (!rid) rid = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = docsSheet_();
    if (sh.getLastRow() > 1 && sh.getRange(2, 10, sh.getLastRow() - 1, 1).getValues().some(r => String(r[0]) === rid)) return { ok: true, dup: true, rid: rid };
    const num = nextDocNum_(sh, date), name = num + '.jpg';
    const url = child_(driveRoot_(), DOCS_FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', name)).getUrl();
    sh.appendRow([new Date(), date, amount, String(body.vendor || '').slice(0, 120), String(body.purpose).slice(0, 200), String(body.building || '').slice(0, 40), who, url, 'nový', rid, String(body.note || '').slice(0, 300), '', false, '', '', num, 'paragon']);
    sh.getRange(sh.getLastRow(), COL_DSEL).insertCheckboxes();
    setPhotoLink_(sh, sh.getLastRow(), 8, '📷 ' + num, url);
    return { ok: true, rid: rid, num: num };
  } finally { lock.releaseLock(); }
}

// posledních 40 dokladů (nejnovější nahoře)
function receiptsList_(P) {
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [], accountant: !!props_().getProperty('ACCOUNTANT') };
  const all = !P || P.acc, n = sh.getLastRow() - 1;
  const v = docRows_(sh, Math.max(COL_DDUE, sh.getLastColumn())).filter(r => all || String(r[6]) === P.name).slice(-60);
  const iso = x => x && Object.prototype.toString.call(x) === '[object Date]' ? x.toISOString() : String(x || '');
  const items = v.filter(r => r[9]).map(r => ({ rid: String(r[9]), num: String(r[COL_DNUM - 1] || ''), date: iso(r[1]), amount: Number(r[2]), vendor: String(r[3] || ''), purpose: String(r[4] || ''), building: String(r[5] || ''), who: String(r[6] || ''), pid: driveId_(r[7]), status: String(r[8] || ''), sent: iso(r[11]), type: String(r[COL_DTYPE - 1] || 'paragon'), invNo: String(r[COL_DINVNO - 1] || '') })).reverse();
  return { ok: true, items: items, total: n, accountant: !!props_().getProperty('ACCOUNTANT') };
}

// AI přečte z fotky paragonu datum, celkovou částku a obchod (jen předvyplnění, uživatel to zkontroluje)
function receiptRead_(b64) {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  if (!b64) return { ok: false, error: 'data' };
  const prompt = 'Na fotce je účtenka nebo paragon (případně faktura). Přečti z ní: datum vystavení (YYYY-MM-DD), celkovou částku k úhradě v Kč včetně DPH (číslo s desetinnou tečkou), název obchodu nebo dodavatele a jeho IČO, pokud je uvedeno. Nečitelné nech prázdné, nic nevymýšlej.';
  const schema = { type: 'OBJECT', properties: { datum: { type: 'STRING', nullable: true }, castka: { type: 'NUMBER', nullable: true }, obchod: { type: 'STRING', nullable: true }, ico: { type: 'STRING', nullable: true } } };
  const think = m => /gemini-2\.5/.test(m) ? { thinkingBudget: 0 } : /gemini-3/.test(m) ? { thinkingLevel: 'minimal' } : null;
  const pay = (m, withThink) => JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: Object.assign({ responseMimeType: 'application/json', responseSchema: schema, temperature: 0.1, maxOutputTokens: 1024 }, withThink && think(m) ? { thinkingConfig: think(m) } : {}) });
  const call = (m, withThink) => UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: pay(m, withThink), muteHttpExceptions: true });
  const det = [];
  for (const m of GEMINI_MODELS) {
    let r = call(m, true);
    if (r.getResponseCode() === 400 && think(m)) r = call(m, false);
    const code = r.getResponseCode();
    if (code !== 200) { det.push(m + ': ' + (code === 429 ? 'kvóta AI vyčerpána' : code === 404 ? 'model není dostupný' : code === 503 ? 'model je přetížený' : 'HTTP ' + code)); continue; }
    try {
      const j = JSON.parse(r.getContentText());
      const t = ((j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || []).map(p => p.text || '').join('');
      const o = JSON.parse(String(t).replace(/```json|```/g, ''));
      const date = /^\d{4}-\d{2}-\d{2}$/.test(String(o.datum || '')) ? String(o.datum) : '';
      const amount = isFinite(docNum_(o.castka)) ? Math.round(docNum_(o.castka) * 100) / 100 : null;
      const vendor = [String(o.obchod || '').trim(), o.ico ? 'IČO ' + String(o.ico).replace(/\s/g, '') : ''].filter(Boolean).join(', ');
      if (date || amount !== null || vendor) return { ok: true, date: date, amount: amount, vendor: vendor, model: m };
      det.push(m + ': nic nepřečteno');
    } catch (e) { det.push(m + ': ' + e.message); }
  }
  return { ok: false, error: 'ai', detail: det.join(' | ') };
}

// odeslání vybraných dokladů účetní (e-mail s fotkami v příloze); doklady se označí jako odeslané a zapíšou do listu Faktury
// pick(rowValues) vybírá doklady; už odeslané se přeskočí
// e-mail(y) oddělené čárkou nebo středníkem; neplatné se zahodí
const mailList_ = v => String(v || '').split(/[,;\s]+/).map(x => x.trim()).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
const escH_ = x => String(x === undefined || x === null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// HTML podoba e-mailu pro účetní: u každého dokladu nahoře krátký řádek (číslo · datum vlevo, částka vpravo),
// delší údaje (dodavatel, faktura, účel, poznámka) pod ním přes celou šířku → vejde se na telefon i s více údaji
// (vložené styly a tabulky kvůli e-mailovým programům)
function receiptsMailHtml_(rows, sum, msg, who, batchId) {
  const F = 'font-family:Arial,Helvetica,sans-serif;';
  const short = w => String(w || '').indexOf('@') > 0 ? String(w).split('@')[0] : String(w || '');
  const cleanNote = n => String(n || '').split(' · ').filter(p => !/^soubor:/.test(p)).join(' · ');
  const line = (txt, style) => txt ? '<div style="' + F + 'word-break:break-word;' + style + '">' + txt + '</div>' : '';
  const item = x => '<tr><td style="padding:12px 14px;border-bottom:1px solid #e5e9ef">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="' + F + 'font-size:13px;color:#8a96a3;vertical-align:top"><b style="color:#2563eb">' + escH_(x.num) + '</b> · ' + escH_(x.date) + '</td>' +
    '<td style="' + F + 'font-size:16px;font-weight:bold;color:#1f2933;text-align:right;white-space:nowrap;vertical-align:top;padding-left:10px">' + escH_(fmtKc_(x.amount)) + '</td>' +
    '</tr></table>' +
    line(escH_(x.vendor || '–'), 'font-size:15px;font-weight:bold;color:#1f2933;margin-top:4px') +
    line(x.inv ? '📄 ' + escH_(x.inv) : '', 'font-size:13px;color:#2563eb;margin-top:2px') +
    line(escH_(x.purpose) + (x.building ? ' · ' + escH_(x.building) : ''), 'font-size:14px;color:#516070;margin-top:2px') +
    line('zapsal(a) ' + escH_(short(x.who)) + (cleanNote(x.note) ? ' · ' + escH_(cleanNote(x.note)) : ''), 'font-size:12px;color:#8a96a3;margin-top:3px') +
    (x.nophoto ? line('přílohu se nepodařilo přiložit', 'font-size:12px;color:#b42318;margin-top:2px') : '') +
    '</td></tr>';
  const count = rows.length + (rows.length === 1 ? ' doklad' : rows.length < 5 ? ' doklady' : ' dokladů');
  return '<div style="background:#f4f6f9;padding:16px 8px;' + F + '">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;margin:0 auto;background:#ffffff;border-radius:12px;border:1px solid #dde3ea;border-collapse:separate;overflow:hidden">' +
    '<tr><td style="background:#2563eb;padding:16px 18px;color:#ffffff;' + F + '"><div style="font-size:19px;font-weight:bold">SVJ Od Vysoké</div><div style="font-size:14px;color:#dbe7ff;margin-top:3px">Doklady k zaúčtování · ' + count + '</div></td></tr>' +
    '<tr><td style="padding:16px 18px 4px;font-size:15px;line-height:1.45;color:#1f2933;' + F + '">Dobrý den,<br>posíláme doklady k zaúčtování. Fotky jsou v příloze, číslo přílohy odpovídá číslu dokladu v naší interní evidenci.</td></tr>' +
    (msg ? '<tr><td style="padding:10px 18px"><div style="background:#fff8e1;border-left:4px solid #e0a800;border-radius:6px;padding:10px 12px;font-size:15px;line-height:1.4;color:#1f2933;word-break:break-word;' + F + '"><b>Zpráva:</b> ' + escH_(msg) + '</div></td></tr>' : '') +
    '<tr><td style="padding:10px 18px 4px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border:1px solid #e5e9ef;border-radius:8px;border-collapse:separate">' +
    rows.map(item).join('') +
    '<tr><td style="padding:12px 14px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="' + F + 'font-size:15px;font-weight:bold;color:#1f2933">Celkem</td>' +
    '<td style="' + F + 'font-size:18px;font-weight:bold;color:#2563eb;text-align:right;white-space:nowrap">' + escH_(fmtKc_(sum)) + '</td>' +
    '</tr></table></td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:12px 18px 18px;font-size:12px;line-height:1.5;color:#8a96a3;' + F + '">Dávka ' + escH_(batchId) + ' · odeslal(a) ' + escH_(short(who)) + '<br>Odesláno z aplikace Fotoapp SVJ Od Vysoké.</td></tr>' +
    '</table></div>';
}

// odeslání vybraných dokladů účetní (e-mail s fotkami v příloze), v kopii výbor (ACCOUNTANT_CC) a případně další adresa;
// doklady se označí jako odeslané a zapíšou do listu Faktury. pick(rowValues) vybírá doklady; už odeslané se přeskočí
function receiptsSendCore_(who, pick, message, extra) {
  const to = props_().getProperty('ACCOUNTANT');
  if (!to) return { ok: false, error: 'noaccountant' };
  const extraList = mailList_(extra);
  if (String(extra || '').trim() && !extraList.length) return { ok: false, error: 'email' };
  const cc = mailList_(props_().getProperty('ACCOUNTANT_CC')).concat(extraList).filter((x, i, a) => a.indexOf(x) === i && x.toLowerCase() !== to.toLowerCase());
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, count: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const width = Math.max(DOCS_HEAD.length, sh.getLastColumn());
    const v = docRows_(sh, width), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const chosen = [], skipped = [];
    v.forEach((r, i) => { if (r[9] && pick(r)) (/^odesl/i.test(String(r[8] || '')) ? skipped : chosen).push({ n: i + 2, r: r }); });
    if (!chosen.length) return { ok: true, count: 0, skipped: skipped.length };
    const batch = chosen.slice(0, DOCS_MAX_MAIL), files = [], lines = [], items = [];
    let sum = 0;
    batch.forEach((x, k) => {
      const r = x.r, d = Object.prototype.toString.call(r[1]) === '[object Date]' ? Utilities.formatDate(r[1], tz, 'd. M. yyyy') : String(r[1]);
      sum += Number(r[2]) || 0;
      const num = String(r[COL_DNUM - 1] || '').trim() || String(k + 1);
      const fd = x => Object.prototype.toString.call(x) === '[object Date]' ? Utilities.formatDate(x, tz, 'd. M. yyyy') : String(x || '');
      const inv = String(r[COL_DTYPE - 1] || '') === 'faktura';
      const it = { num: num, date: d, amount: r[2], vendor: r[3], purpose: r[4], building: r[5], who: r[6], note: r[10], nophoto: false, inv: inv ? 'Faktura' + (r[COL_DINVNO - 1] ? ' č. ' + r[COL_DINVNO - 1] : '') + (r[COL_DDUZP - 1] ? ' · DUZP ' + fd(r[COL_DDUZP - 1]) : '') + (r[COL_DDUE - 1] ? ' · splatnost ' + fd(r[COL_DDUE - 1]) : '') : '' };
      try { const f = DriveApp.getFileById(driveId_(r[7])); files.push(f.getBlob().setName(num + (/pdf/i.test(f.getMimeType()) ? '.pdf' : '.jpg'))); } catch (e) { it.nophoto = true; }
      items.push(it);
      lines.push(num + ' · ' + d + ' · ' + fmtKc_(r[2]) + ' · ' + (r[3] || '–') + (it.inv ? ' · ' + it.inv : '') + ' · ' + r[4] + (r[5] ? ' · ' + r[5] : '') + ' · zapsal(a) ' + r[6] + (r[10] ? ' · ' + r[10] : '') + (it.nophoto ? ' (fotku se nepodařilo přiložit)' : ''));
    });
    sh.getRange(1, COL_DMSG).getValue() || docsSheet_(); // doplní hlavičky Dávka / Zpráva, pokud chybí
    const stamp = new Date(), msg = String(message || '').trim().slice(0, 500), batchId = Utilities.formatDate(stamp, tz, 'yyyyMMdd-HHmmss');
    const mail = {
      to: to, name: 'SVJ Od Vysoké', subject: envTag_() + 'SVJ Od Vysoké – doklady k zaúčtování (' + batch.length + ')',
      body: 'Dobrý den,\n\nposíláme doklady k zaúčtování. Fotky jsou v příloze, číslo přílohy odpovídá číslu dokladu v naší interní evidenci:\n\n' + lines.join('\n') + '\n\nCelkem: ' + fmtKc_(sum) + (msg ? '\n\nZpráva: ' + msg : '') + '\n\nDávka ' + batchId + ', odeslal(a): ' + who + '\nSVJ Od Vysoké',
      htmlBody: receiptsMailHtml_(items, sum, msg, who, batchId),
      attachments: files
    };
    if (cc.length) mail.cc = cc.join(',');
    MailApp.sendEmail(mail);
    const note = msg + (extraList.length ? (msg ? ' ' : '') + '(kopie: ' + extraList.join(', ') + ')' : '');
    batch.forEach(x => { sh.getRange(x.n, 9).setValue('odesláno'); sh.getRange(x.n, 12).setValue(stamp); sh.getRange(x.n, COL_DSEL).setValue(false); sh.getRange(x.n, COL_DBATCH).setValue(batchId); sh.getRange(x.n, COL_DMSG).setValue(note); });
    invoicesAppend_(batch.map(x => x.r), stamp, batchId, note);
    return { ok: true, count: batch.length, to: to, cc: cc, more: chosen.length - batch.length, skipped: skipped.length };
  } finally { lock.releaseLock(); }
}
// z aplikace: seznam ID dokladů (jen pro toho, kdo smí odesílat účetní)
function receiptsSend_(body, who) {
  const ids = (Array.isArray(body.rids) ? body.rids : []).map(String).filter(Boolean);
  if (!ids.length) return { ok: false, error: 'data' };
  return receiptsSendCore_(who, r => ids.indexOf(String(r[9])) >= 0, body.message, body.extra);
}

// ---------- list Faktury: odeslané doklady, řazené podle DUZP ----------
const SHEET_INV = 'Faktury';
const INV_HEAD = ['DUZP', 'Datum dokladu', 'Částka (Kč)', 'Dodavatel', 'Účel', 'Budova', 'Doklad (fotka)', 'ID dokladu', 'Odesláno účetní', 'Poznámka', 'Dávka', 'Zpráva pro účetní', 'Číslo dokladu', 'Typ', 'Číslo faktury', 'Splatnost'];
// list se založí, pokud chybí; pokud už existuje (třeba vlastní), zapisuje se podle názvů sloupců a chybějící sloupce se přidají na konec
function invoicesAppend_(rows, stamp, batchId, msg) {
  const ss = docsSS_();
  let sh = ss.getSheetByName(SHEET_INV);
  if (!sh) {
    sh = ss.insertSheet(SHEET_INV); sh.appendRow(INV_HEAD); sh.setFrozenRows(1); sh.getRange(1, 1, 1, INV_HEAD.length).setFontWeight('bold');
    sh.getRange('A:B').setNumberFormat('d. M. yyyy'); sh.getRange('C:C').setNumberFormat('#,##0.00'); sh.getRange('I:I').setNumberFormat('d. M. yyyy');
  }
  let head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).normalize('NFC').trim().toLowerCase()) : [];
  const col = name => { let i = head.indexOf(name.toLowerCase()); if (i < 0) { head.push(name.toLowerCase()); i = head.length - 1; sh.getRange(1, i + 1).setValue(name).setFontWeight('bold'); } return i; };
  const ci = INV_HEAD.map(col);
  const known = sh.getLastRow() > 1 ? sh.getRange(2, ci[7] + 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const out = rows.filter(r => known.indexOf(String(r[9])) < 0).map(r => {
    const row = new Array(head.length).fill('');
    const duzp = Object.prototype.toString.call(r[COL_DDUZP - 1]) === '[object Date]' ? r[COL_DDUZP - 1] : r[1];
    [duzp, r[1], r[2], r[3], r[4], r[5], r[7], String(r[9]), stamp, r[10], batchId || '', msg || '', String(r[COL_DNUM - 1] || ''), String(r[COL_DTYPE - 1] || 'paragon'), String(r[COL_DINVNO - 1] || ''), r[COL_DDUE - 1] || ''].forEach((val, k) => { row[ci[k]] = val; }); // DUZP = datum dokladu (lze přepsat)
    return row;
  });
  if (out.length) {
    const first = sh.getLastRow() + 1;
    sh.getRange(first, 1, out.length, head.length).setValues(out);
    out.forEach((row, j) => { const u = String(row[ci[6]] || ''); if (/^https?:/.test(u)) setPhotoLink_(sh, first + j, ci[6] + 1, (row[ci[13]] === 'faktura' ? '📄 ' : '📷 ') + (row[ci[12]] || 'fotka'), u); });
  }
  linkifyPhotos_(sh, ci[6] + 1, ci[12] + 1);
  sortInvoices_(sh, ci[0]);
}
function sortInvoices_(sh, duzpIdx) {
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort({ column: duzpIdx + 1, ascending: true });
}

// ---------- menu v tabulce ----------
function setAccountant() {
  const ui = SpreadsheetApp.getUi(), cur = props_().getProperty('ACCOUNTANT') || '';
  const r = ui.prompt('E-mail účetní', 'Na tuto adresu se budou posílat doklady k zaúčtování.' + (cur ? ' Teď: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = r.getResponseText().trim();
  if (v === '') { props_().deleteProperty('ACCOUNTANT'); ui.alert('E-mail účetní je smazaný.'); return; }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) { ui.alert('To nevypadá jako e-mail.'); return; }
  props_().setProperty('ACCOUNTANT', v); ui.alert('Uloženo: ' + v);
}
function selectedCount_() {
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < COL_DSEL) return 0;
  return sh.getRange(2, 1, sh.getLastRow() - 1, COL_DSEL).getValues().filter(r => r[COL_DSEL - 1] === true && r[9] && !/^odesl/i.test(String(r[8] || ''))).length;
}
function sendReceipts() {
  const ui = SpreadsheetApp.getUi(), n = selectedCount_(), to = props_().getProperty('ACCOUNTANT');
  if (!to) { ui.alert('Nejdřív nastav e-mail účetní (menu Doklady → Nastavit e-mail účetní).'); return; }
  if (!n) { ui.alert('Nic není vybráno. V listu Doklady zaškrtni ve sloupci Vybrat k odeslání doklady, které chceš poslat.'); return; }
  const cc = mailList_(props_().getProperty('ACCOUNTANT_CC'));
  const html = '<div style="font:14px Arial,Helvetica,sans-serif;color:#1f2933">' +
    '<p style="margin:0 0 10px">Odeslat <b>' + n + '</b> ' + (n === 1 ? 'vybraný doklad' : n < 5 ? 'vybrané doklady' : 'vybraných dokladů') + ' účetní <b>' + escH_(to) + '</b>' + (cc.length ? '<br><span style="color:#516070">v kopii: ' + escH_(cc.join(', ')) + '</span>' : '') + '</p>' +
    '<label style="font-weight:bold">Zpráva pro účetní <span style="font-weight:normal;color:#8a96a3">(nepovinné)</span></label>' +
    '<textarea id="m" rows="3" style="width:100%;box-sizing:border-box;margin:4px 0 10px;font:14px Arial;padding:6px" placeholder="např. Proplaceno z účtu SVJ jedním příkazem Janu Novákovi"></textarea>' +
    '<label style="font-weight:bold">Poslat také na e-mail <span style="font-weight:normal;color:#8a96a3">(nepovinné, víc adres odděl čárkou)</span></label>' +
    '<input id="e" type="text" style="width:100%;box-sizing:border-box;margin:4px 0 12px;font:14px Arial;padding:6px" placeholder="např. jan.novak@example.com">' +
    '<div id="s" style="min-height:18px;color:#b42318;margin-bottom:8px"></div>' +
    '<button id="b" style="background:#2563eb;color:#fff;border:0;border-radius:8px;padding:9px 16px;font:bold 14px Arial;cursor:pointer">Odeslat</button> ' +
    '<button onclick="google.script.host.close()" style="background:#e8edf5;border:0;border-radius:8px;padding:9px 16px;font:14px Arial;cursor:pointer">Zrušit</button></div>' +
    '<script>document.getElementById("b").onclick=function(){var b=this;b.disabled=true;document.getElementById("s").style.color="#516070";document.getElementById("s").textContent="Odesílám…";' +
    'google.script.run.withSuccessHandler(function(t){document.getElementById("s").style.color=/^Odesláno/.test(t)?"#166534":"#b42318";document.getElementById("s").textContent=t;if(/^Odesláno/.test(t)){b.textContent="Hotovo";setTimeout(function(){google.script.host.close()},2500)}else b.disabled=false})' +
    '.withFailureHandler(function(e){document.getElementById("s").textContent="Chyba: "+e.message;b.disabled=false}).receiptsDialogSend(document.getElementById("m").value,document.getElementById("e").value)};</script>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(360), 'Odeslat doklady účetní');
}
// volá okno Odeslat doklady účetní
function receiptsDialogSend(message, extra) {
  const r = receiptsSendCore_(Session.getActiveUser().getEmail() || 'správce (tabulka)', row => row[COL_DSEL - 1] === true, message, extra);
  if (!r.ok) return r.error === 'noaccountant' ? 'Nejdřív nastav e-mail účetní.' : r.error === 'email' ? 'E-mail navíc nevypadá jako platná adresa.' : 'Odeslání se nepovedlo.';
  if (!r.count) return r.skipped ? 'Vybrané doklady už byly odeslané.' : 'Nic není vybráno.';
  dashRefreshSafe_();
  return 'Odesláno ' + r.count + ' dokladů na ' + r.to + (r.cc && r.cc.length ? ' (v kopii ' + r.cc.join(', ') + ')' : '') + ' a zapsáno do listu Faktury.' + (r.more ? ' Zbývá ještě ' + r.more + ' vybraných, pošli je znovu.' : '');
}
function setAccountantCc() {
  const ui = SpreadsheetApp.getUi(), cur = props_().getProperty('ACCOUNTANT_CC') || '';
  const r = ui.prompt('Kopie e-mailu účetní (výbor)', 'Tyto adresy dostanou každý e-mail s doklady v kopii. Víc adres odděl čárkou. Prázdné = bez kopie.' + (cur ? ' Teď: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (t === '') { props_().deleteProperty('ACCOUNTANT_CC'); ui.alert('Kopie je zrušená.'); return; }
  const list = mailList_(t);
  if (!list.length) { ui.alert('Nenašel jsem žádnou platnou adresu.'); return; }
  props_().setProperty('ACCOUNTANT_CC', list.join(',')); ui.alert('Uloženo: ' + list.join(', '));
}
function sortInvoices() {
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_INV);
  if (!sh || sh.getLastRow() < 3) { SpreadsheetApp.getUi().alert('List Faktury je prázdný nebo má jen jeden řádek.'); return; }
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase()), i = head.indexOf('duzp');
  if (i < 0) { SpreadsheetApp.getUi().alert('V listu Faktury chybí sloupec DUZP.'); return; }
  sortInvoices_(sh, i); dss.toast('Faktury jsou seřazené podle DUZP.', 'Hotovo', 5);
}
