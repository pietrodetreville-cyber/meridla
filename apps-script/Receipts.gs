// ---------- doklady (paragony): foto, zápis do listu Doklady, odeslání účetní ----------
// Kdo má v listu Lidé ve sloupci Doklady "ano", smí fotit doklady (vidí jen své). Kdo má "ano" ve sloupci Odesílá účetní,
// vidí doklady všech a smí vybrané odeslat účetní z aplikace. V tabulce se vybírá zaškrtnutím ve sloupci Vybrat k odeslání.
// Odeslané doklady se zapíšou do listu Faktury (DUZP = datum dokladu, lze přepsat) a list se seřadí podle DUZP.
// Fotky se ukládají na Disk do složky Doklady (SVJ_app/OSTRE nebo TESTOVACI). E-mail účetní je ve vlastnostech skriptu (ACCOUNTANT),
// nastavuje se z menu Měřidla → Doklady → Nastavit e-mail účetní.

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
  SpreadsheetApp.getUi().createMenu('Doklady').addItem('Odeslat vybrané (zaškrtnuté) doklady účetní', 'sendReceipts').addItem('Seřadit Faktury podle DUZP', 'sortInvoices').addItem('Nastavit e-mail účetní', 'setAccountant').addToUi();
}
// z menu tabulky měřidel: odkaz na soubor dokladů
function openDocsFile() {
  const ss = docsSS_(), url = ss.getUrl();
  const html = HtmlService.createHtmlOutput('<p style="font:14px Arial">Doklady a faktury jsou v samostatné tabulce:</p><p style="font:14px Arial"><a href="' + url + '" target="_blank">' + ss.getName() + '</a></p>').setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Soubor dokladů');
}
const DOCS_HEAD = ['Zapsáno', 'Datum dokladu', 'Částka (Kč)', 'Obchod / dodavatel', 'Účel', 'Budova', 'Kdo', 'Fotka', 'Stav', 'ID', 'Poznámka', 'Odesláno účetní', 'Vybrat k odeslání', 'Dávka', 'Zpráva pro účetní'];
const COL_DSEL = 13, COL_DBATCH = 14, COL_DMSG = 15;
const DOCS_MAX_MAIL = 15; // nejvýš dokladů v jednom e-mailu (velikost příloh)

function docsSheet_() {
  const ss = docsSS_();
  let sh = ss.getSheetByName(SHEET_DOCS);
  if (!sh) sh = ss.insertSheet(SHEET_DOCS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(DOCS_HEAD); sh.setFrozenRows(1); sh.getRange(1, 1, 1, DOCS_HEAD.length).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('d. M. yyyy'); sh.getRange('C:C').setNumberFormat('#,##0.00');
  }
  [[COL_DBATCH, 'Dávka'], [COL_DMSG, 'Zpráva pro účetní']].forEach(([c, n]) => { if (sh.getLastColumn() < c || !sh.getRange(1, c).getValue()) sh.getRange(1, c).setValue(n).setFontWeight('bold'); });
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
    const name = body.date + '_' + String(Math.round(amount)) + 'Kc_' + rid + '.jpg';
    const url = child_(driveRoot_(), DOCS_FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', name)).getUrl();
    sh.appendRow([new Date(), date, amount, String(body.vendor || '').slice(0, 120), String(body.purpose).slice(0, 200), String(body.building || '').slice(0, 40), who, url, 'nový', rid, String(body.note || '').slice(0, 300), '', false]);
    sh.getRange(sh.getLastRow(), COL_DSEL).insertCheckboxes();
    return { ok: true, rid: rid };
  } finally { lock.releaseLock(); }
}

// posledních 40 dokladů (nejnovější nahoře)
function receiptsList_(P) {
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [], accountant: !!props_().getProperty('ACCOUNTANT') };
  const all = !P || P.acc, n = sh.getLastRow() - 1;
  const v = sh.getRange(2, 1, n, 12).getValues().filter(r => all || String(r[6]) === P.name).slice(-60);
  const iso = x => x && Object.prototype.toString.call(x) === '[object Date]' ? x.toISOString() : String(x || '');
  const items = v.filter(r => r[9]).map(r => ({ rid: String(r[9]), date: iso(r[1]), amount: Number(r[2]), vendor: String(r[3] || ''), purpose: String(r[4] || ''), building: String(r[5] || ''), who: String(r[6] || ''), pid: driveId_(r[7]), status: String(r[8] || ''), sent: iso(r[11]) })).reverse();
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
function receiptsSendCore_(who, pick, message) {
  const to = props_().getProperty('ACCOUNTANT');
  if (!to) return { ok: false, error: 'noaccountant' };
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, count: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const width = Math.max(DOCS_HEAD.length, sh.getLastColumn());
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const chosen = [], skipped = [];
    v.forEach((r, i) => { if (r[9] && pick(r)) (/^odesl/i.test(String(r[8] || '')) ? skipped : chosen).push({ n: i + 2, r: r }); });
    if (!chosen.length) return { ok: true, count: 0, skipped: skipped.length };
    const batch = chosen.slice(0, DOCS_MAX_MAIL), files = [], lines = [];
    let sum = 0;
    batch.forEach((x, k) => {
      const r = x.r, d = Object.prototype.toString.call(r[1]) === '[object Date]' ? Utilities.formatDate(r[1], tz, 'd. M. yyyy') : String(r[1]);
      sum += Number(r[2]) || 0;
      lines.push((k + 1) + '. ' + d + ' · ' + fmtKc_(r[2]) + ' · ' + (r[3] || '–') + ' · ' + r[4] + (r[5] ? ' · ' + r[5] : '') + ' · zapsal(a) ' + r[6] + (r[10] ? ' · ' + r[10] : ''));
      try { const f = DriveApp.getFileById(driveId_(r[7])); files.push(f.getBlob().setName((k + 1) + '_' + f.getName())); } catch (e) { lines[lines.length - 1] += ' (fotku se nepodařilo přiložit)'; }
    });
    sh.getRange(1, COL_DMSG).getValue() || docsSheet_(); // doplní hlavičky Dávka / Zpráva, pokud chybí
    const msgTxt = String(message || '').trim();
    MailApp.sendEmail({
      to: to, name: 'SVJ Od Vysoké', subject: envTag_() + 'SVJ Od Vysoké – doklady k zaúčtování (' + batch.length + ')',
      body: 'Dobrý den,\n\nposíláme doklady k zaúčtování (fotky v příloze, číslo přílohy odpovídá řádku):\n\n' + lines.join('\n') + '\n\nCelkem: ' + fmtKc_(sum) + (msgTxt ? '\n\nZpráva: ' + msgTxt : '') + '\n\nOdeslal(a): ' + who + '\nSVJ Od Vysoké',
      attachments: files
    });
    const stamp = new Date(), msg = String(message || '').trim().slice(0, 500), batchId = Utilities.formatDate(stamp, tz, 'yyyyMMdd-HHmmss');
    batch.forEach(x => { sh.getRange(x.n, 9).setValue('odesláno'); sh.getRange(x.n, 12).setValue(stamp); sh.getRange(x.n, COL_DSEL).setValue(false); sh.getRange(x.n, COL_DBATCH).setValue(batchId); sh.getRange(x.n, COL_DMSG).setValue(msg); });
    invoicesAppend_(batch.map(x => x.r), stamp, batchId, msg);
    return { ok: true, count: batch.length, to: to, more: chosen.length - batch.length, skipped: skipped.length };
  } finally { lock.releaseLock(); }
}
// z aplikace: seznam ID dokladů (jen pro toho, kdo smí odesílat účetní)
function receiptsSend_(body, who) {
  const ids = (Array.isArray(body.rids) ? body.rids : []).map(String).filter(Boolean);
  if (!ids.length) return { ok: false, error: 'data' };
  return receiptsSendCore_(who, r => ids.indexOf(String(r[9])) >= 0, body.message);
}

// ---------- list Faktury: odeslané doklady, řazené podle DUZP ----------
const SHEET_INV = 'Faktury';
const INV_HEAD = ['DUZP', 'Datum dokladu', 'Částka (Kč)', 'Dodavatel', 'Účel', 'Budova', 'Doklad (fotka)', 'ID dokladu', 'Odesláno účetní', 'Poznámka', 'Dávka', 'Zpráva pro účetní'];
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
    [r[1], r[1], r[2], r[3], r[4], r[5], r[7], String(r[9]), stamp, r[10], batchId || '', msg || ''].forEach((val, k) => { row[ci[k]] = val; }); // DUZP = datum dokladu (lze přepsat)
    return row;
  });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, head.length).setValues(out);
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
function sendReceipts() {
  const ui = SpreadsheetApp.getUi(), dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_DOCS);
  const n = sh && sh.getLastRow() > 1 && sh.getLastColumn() >= COL_DSEL ? sh.getRange(2, 1, sh.getLastRow() - 1, COL_DSEL).getValues().filter(r => r[COL_DSEL - 1] === true && r[9] && !/^odesl/i.test(String(r[8] || ''))).length : 0;
  if (!n) { ui.alert('Nic není vybráno. V listu Doklady zaškrtni ve sloupci Vybrat k odeslání doklady, které chceš poslat.'); return; }
  // krátká zpráva pro účetní (např. proplaceno z účtu jedním příkazem komu); doklady odeslané najednou dostanou společné číslo dávky
  const p = ui.prompt('Odeslat ' + n + ' dokladů účetní', 'Zpráva pro účetní (nepovinné), např. „Proplaceno z účtu SVJ jedním příkazem Janu Novákovi“:', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const r = receiptsSendCore_(Session.getActiveUser().getEmail() || 'správce (tabulka)', row => row[COL_DSEL - 1] === true, p.getResponseText());
  if (!r.ok) { ui.alert(r.error === 'noaccountant' ? 'Nejdřív nastav e-mail účetní (Měřidla → Doklady → Nastavit e-mail účetní).' : 'Odeslání se nepovedlo.'); return; }
  ui.alert(r.count ? 'Odesláno ' + r.count + ' dokladů na ' + r.to + ' a zapsáno do listu Faktury (soubor dokladů).' + (r.more ? ' Zbývá ještě ' + r.more + ' vybraných, pošli je dalším spuštěním.' : '') + (r.skipped ? ' Přeskočeno už odeslaných: ' + r.skipped + '.' : '') : (r.skipped ? 'Vybrané doklady už byly odeslané.' : 'Nic není vybráno. V listu Doklady zaškrtni ve sloupci Vybrat k odeslání doklady, které chceš poslat.'));
}
function sortInvoices() {
  const dss = docsSSIfExists_(), sh = dss && dss.getSheetByName(SHEET_INV);
  if (!sh || sh.getLastRow() < 3) { SpreadsheetApp.getUi().alert('List Faktury je prázdný nebo má jen jeden řádek.'); return; }
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase()), i = head.indexOf('duzp');
  if (i < 0) { SpreadsheetApp.getUi().alert('V listu Faktury chybí sloupec DUZP.'); return; }
  sortInvoices_(sh, i); dss.toast('Faktury jsou seřazené podle DUZP.', 'Hotovo', 5);
}
