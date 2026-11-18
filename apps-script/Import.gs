// ---------- Import hodnot z papírových archů ----------
// List „Import z papíru": Datum | Budova | Druh | Číslo měřidla | Stav | Poznámka | Fotka archu | Stav importu | Měřidlo (id)
// 1) hodnoty se do listu vloží ručně, nebo je z vyfocené stránky přepíše AI (správce, v appce),
// 2) menu Měřidla → Import z papíru → Náhled (nic nezapisuje) zkontroluje přiřazení k měřidlům,
// 3) menu … → Zapsat hodnoty je zapíše jako zpětné záznamy (opakované spuštění nic nezdvojí).

const SHEET_IMPORT = 'Import z papíru';
const IMPORT_HEAD = ['Datum', 'Budova', 'Druh', 'Číslo měřidla', 'Stav', 'Poznámka', 'Fotka archu', 'Stav importu', 'Měřidlo (id)'];
const PAPER_FOLDER = 'Papírové záznamy';

const paperKey_ = x => String(x === undefined || x === null ? '' : x).toLowerCase().replace(/[^a-z0-9]/g, '');

function ensureImportSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_IMPORT);
  if (!sh) { sh = ss.insertSheet(SHEET_IMPORT); sh.appendRow(IMPORT_HEAD); sh.setFrozenRows(1); }
  return sh;
}

// sjednocení názvu druhu (voda, vodoměr, plyn …)
function normType_(t) {
  const x = String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  if (!x) return '';
  if (/^vod/.test(x)) return /tuv/.test(x) ? 'TUV' : 'Vodoměr';
  if (/^tuv/.test(x)) return 'TUV';
  if (/^plyn/.test(x)) return 'Plynoměr';
  if (/^elek/.test(x) || x === 'el') return 'Elektroměr';
  if (/^kalor/.test(x)) return 'Kalorimetr';
  return String(t).trim();
}

// datum z buňky: skutečné datum, nebo text 30.9.2026 / 30. 9. 2026 / 2026-09-30; čas se nastaví na 12:00
function parseImportDate_(v, tz) {
  let ymd = '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) ymd = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  else {
    const t = String(v === undefined || v === null ? '' : v).trim();
    let m = t.match(/^(\d{1,2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{4})/);
    if (m) ymd = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    else { m = t.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) ymd = m[1] + '-' + m[2] + '-' + m[3]; }
  }
  if (!ymd) return null;
  try {
    const d = Utilities.parseDate(ymd + ' 12:00', tz, 'yyyy-MM-dd HH:mm');
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === ymd && d.getFullYear() >= 2000 ? d : null;
  } catch (e) { return null; }
}

// najde měřidlo: podle čísla, jinak podle budovy a druhu (musí být právě jedno)
function importMeter_(r, meters) {
  const num = paperKey_(r.cislo);
  if (num) {
    const m = meters.find(x => paperKey_(x.id) === num);
    return m ? { m: m } : { err: 'měřidlo s číslem „' + r.cislo + '“ není v listu Měřidla' };
  }
  const typ = normType_(r.druh), b = String(r.budova || '').toLowerCase().trim();
  if (!b || !typ) return { err: 'chybí číslo měřidla, nebo budova a druh' };
  const c = meters.filter(x => String(x.building).toLowerCase() === b && String(x.typeName).toLowerCase() === typ.toLowerCase());
  if (c.length === 1) return { m: c[0] };
  return { err: c.length ? 'v budově je víc měřidel tohoto druhu, doplň číslo měřidla' : 'v budově ' + r.budova + ' není měřidlo druhu ' + typ };
}

// odkaz na fotku archu na Disku podle názvu souboru (do sloupce Fotka v Odečtech); text už s odkazem se použije přímo
const archUrlCache_ = {};
function archUrl_(name) {
  const n = String(name || '').trim();
  if (!n) return '';
  if (n.indexOf('http') === 0) return n;
  if (archUrlCache_[n] !== undefined) return archUrlCache_[n];
  let url = '';
  try { const it = child_(driveRoot_(), PAPER_FOLDER).getFilesByName(n); if (it.hasNext()) url = it.next().getUrl(); } catch (e) {}
  return (archUrlCache_[n] = url);
}

function importRun_(write) {
  const sh = ensureImportSheet_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const data = sh.getDataRange().getValues();
  const say = t => { try { SpreadsheetApp.getActive().toast(t, write ? 'Import hotov' : 'Náhled hotov', 20); } catch (e) {} };
  if (data.length < 2) { Logger.log('List „' + SHEET_IMPORT + '“ je prázdný.'); say('List je prázdný.'); return 'List je prázdný.'; }
  const head = data[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
  const find = n => head.indexOf(n);
  const c = { date: find('datum'), b: find('budova'), t: find('druh'), n: find('číslo měřidla'), v: find('stav'), note: find('poznámka'), photo: find('fotka archu'), st: find('stav importu'), id: find('měřidlo (id)') };
  if (c.date < 0 || c.v < 0) { Logger.log('V listu chybí sloupec Datum nebo Stav.'); say('V listu chybí sloupec Datum nebo Stav.'); return 'V listu chybí sloupec Datum nebo Stav.'; }
  if (c.st < 0) { c.st = head.length; head.push('stav importu'); sh.getRange(1, c.st + 1).setValue('Stav importu'); }
  if (c.id < 0) { c.id = head.length; head.push('měřidlo (id)'); sh.getRange(1, c.id + 1).setValue('Měřidlo (id)'); }
  const meters = meters_(), cnt = { ok: 0, wrote: 0, dup: 0, err: 0, skip: 0 };
  const lock = LockService.getScriptLock();
  if (write) lock.waitLock(30000);
  try {
    for (let i = 1; i < data.length; i++) {
      const r = data[i], rowNo = i + 1, cell = k => (c[k] >= 0 && r[c[k]] !== undefined ? r[c[k]] : '');
      if (/^(ZAPSÁNO|VYŘAZENO)/.test(String(cell('st')))) { cnt.skip++; continue; }
      const raw = String(cell('v')).trim();
      if (raw === '' && String(cell('date')).trim() === '' && String(cell('n')).trim() === '' && String(cell('b')).trim() === '') continue;
      const when = parseImportDate_(cell('date'), tz), value = raw === '' ? NaN : Number(raw.replace(/\s/g, '').replace(',', '.'));
      let msg, mid = '';
      if (!when) msg = 'CHYBA: chybí nebo neplatné datum';
      else if (when.getTime() > Date.now() + 3600000) msg = 'CHYBA: datum v budoucnosti';
      else if (!isFinite(value)) msg = 'CHYBA: chybí nebo neplatný stav';
      else {
        const f = importMeter_({ cislo: cell('n'), budova: cell('b'), druh: cell('t') }, meters);
        if (!f.m) msg = 'CHYBA: ' + f.err;
        else {
          mid = f.m.id;
          const val = round_(value, f.m.decimals);
          if (!write) { msg = 'OK, zapíše se (' + f.m.name + ', ' + (f.m.buildingLabel || f.m.building) + ') = ' + val; cnt.ok++; }
          else {
            const note = [String(cell('note')).trim(), String(cell('photo')).trim() ? 'Arch: ' + String(cell('photo')).trim() : '', 'Import z papíru.'].filter(Boolean).join(' ');
            const res = backfillCore_(f.m, when, val, note, 'Správce', false, archUrl_(cell('photo')));
            if (res.ok) { msg = 'ZAPSÁNO ✓ ' + Utilities.formatDate(new Date(), tz, 'd. M. yyyy H:mm'); cnt.wrote++; }
            else if (res.error === 'dup') { msg = 'PŘESKOČENO: stejný stav k tomuto dni už je zapsaný'; cnt.dup++; }
            else { msg = 'CHYBA: ' + res.error; }
          }
        }
      }
      if (/^CHYBA/.test(msg)) cnt.err++;
      sh.getRange(rowNo, c.st + 1).setValue(msg);
      sh.getRange(rowNo, c.id + 1).setValue(mid);
    }
  } finally {
    if (write) lock.releaseLock();
  }
  const text = write
    ? 'Zapsáno: ' + cnt.wrote + ', už zapsáno: ' + cnt.dup + ', chyby: ' + cnt.err + (cnt.skip ? ', přeskočeno (zapsané a vyřazené): ' + cnt.skip : '')
    : 'Zapíše se: ' + cnt.ok + ', chyby: ' + cnt.err + (cnt.skip ? ', přeskočeno (zapsané a vyřazené): ' + cnt.skip : '') + '. Výsledek je ve sloupci Stav importu.';
  Logger.log(text);
  say(text);
  return text;
}

function importPreview() { importRun_(false); }
function importWrite() { importRun_(true); }

// ---------- foto papírového archu (správce v appce): uložení na Disk + AI přepis do listu ----------
function paperUpload_(body) {
  const chk = adminCheck_(body.pin);
  if (!chk.ok) return chk;
  return paperProcess_(body.photo);
}

// uloží fotku archu na Disk a nechá ji přepsat AI do listu (volá se z appky po ověření PINu i z okna v tabulce)
function paperProcess_(photo) {
  if (!photo) return { ok: false, error: 'data' };
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const name = 'arch-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss') + '.jpg';
  try { child_(driveRoot_(), PAPER_FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(photo), 'image/jpeg', name)); }
  catch (e) { return { ok: false, error: 'drive' }; }
  // čtení AI se nespouští: arch se jen uloží, hodnoty se z něj přečtou později v okně Import z papíru
  return { ok: true, name: name, saved: true, count: 0 };
}

// přečte arch AI a řádky přidá do listu (fotka je už na Disku pod názvem name)
function paperRead_(photo, name) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const ai = paperTranscribe_(photo);
  if (!ai.ok) return { ok: true, name: name, count: 0, aiError: ai.error, detail: ai.detail || '' };
  const sh = ensureImportSheet_(), c = importCols_(sh);
  ai.rows.forEach(r => {
    let d = '';
    if (r.datum) { try { d = Utilities.parseDate(r.datum + ' 12:00', tz, 'yyyy-MM-dd HH:mm'); } catch (e) {} }
    const row = new Array(c.n).fill('');
    row[c.date] = d; row[c.b] = r.budova; row[c.t] = r.druh; row[c.num] = r.cislo; row[c.v] = r.stav; row[c.note] = 'AI z archu, ověř'; row[c.photo] = name; row[c.st] = 'K ověření';
    sh.appendRow(row);
  });
  return { ok: true, name: name, count: ai.rows.length, rows: ai.rows };
}

// přepis tabulky z fotky (Gemini, klíč je jen ve vlastnostech skriptu)
// vytáhne z odpovědi pole řádků i když je JSON obalený, v objektu nebo uříznutý (došel výstup)
function parseRows_(t) {
  t = String(t || '').split('```json').join('').split('```').join('').trim();
  let v;
  try { v = JSON.parse(t); } catch (e) {
    const i = t.indexOf('[');
    if (i < 0) return null;
    const u = t.slice(i), k = u.lastIndexOf('}');
    if (k < 0) return null;
    try { v = JSON.parse(u.slice(0, k + 1) + ']'); } catch (e2) { return null; }
  }
  if (!Array.isArray(v) && v && typeof v === 'object') { const arr = Object.keys(v).map(x => v[x]).find(x => Array.isArray(x)); v = arr || [v]; }
  return Array.isArray(v) ? v : null;
}

// z odpovědi 429 vytáhne, který limit Google myslí (za minutu / za den, počet požadavků / tokenů)
function quotaWhy_(txt) {
  try {
    const j = JSON.parse(txt), e = j.error || {}, out = [];
    (e.details || []).forEach(d => (d.violations || []).forEach(v => out.push(String(v.quotaId || v.quotaMetric || '').split('/').pop())));
    (e.details || []).forEach(d => { if (d.retryDelay) out.push('zkus za ' + d.retryDelay); });
    return out.filter(Boolean).slice(0, 3).join(', ') || String(e.message || '').replace(/\s+/g, ' ').slice(0, 90);
  } catch (x) { return String(txt || '').replace(/\s+/g, ' ').slice(0, 90); }
}

function paperTranscribe_(b64) {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  let known = '';
  try { known = meters_().slice(0, 80).map(m => (m.building || '?') + ' ' + (m.typeName || m.name) + ' ' + m.id).join('; '); } catch (e) {}
  const prompt = 'Na fotce je papírový arch (tabulka) s ručně psanými odečty měřidel v bytovém domě. Přečti z něj všechny řádky s hodnotou, ať má arch jakékoli rozložení sloupců. Typické rozložení: Č.P. (číslo popisné domu 271 až 275; řádky se značkou /2, /4, /6, /8, /10 patří k domu uvedenému těsně nad nimi), plynoměr a vodoměr v m3, číslo elektroměru, stav elektroměru v kWh a datum. Ručně napsané datum se může táhnout přes víc řádků a platí pro všechny řádky vedle něj. Řádky kolem nápisu KOTELNA: elektroměry patří k domu 271, plynoměry, vodoměr a TUV (teplá užitková voda) patří kotelně. ' +
    (known ? 'Známá měřidla (budova, druh, id): ' + known + '. ' : '') +
    'Vrať pole objektů, jeden objekt pro každou přečtenou hodnotu: datum ve tvaru YYYY-MM-DD (nebo null), budova ("271".."275" nebo "Kotelna"), druh ("Vodoměr", "Plynoměr", "Elektroměr" nebo "TUV"), cislo (číslo elektroměru bez mezer, např. W060073, jinak prázdný řetězec), stav (číslo s desetinnou tečkou). Hodnoty nevymýšlej, nečitelné vynech.';
  // celý arch = JEDEN požadavek se strukturovanou odpovědí (pole řádků), ne jeden požadavek na řádek
  const schema = { type: 'ARRAY', items: { type: 'OBJECT', properties: { datum: { type: 'STRING', nullable: true }, budova: { type: 'STRING' }, druh: { type: 'STRING' }, cislo: { type: 'STRING' }, stav: { type: 'NUMBER' } }, required: ['budova', 'druh', 'stav'] } };
  // bez „přemýšlení“ modelu je čtení rychlejší (2.5: rozpočet 0, řada 3: úroveň minimal); když model nastavení odmítne (400), jde se bez něj
  const think = m => /gemini-2.5/.test(m) ? { thinkingBudget: 0 } : /gemini-3/.test(m) ? { thinkingLevel: 'minimal' } : null;
  const pay = (m, withThink) => JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: Object.assign({ responseMimeType: 'application/json', responseSchema: schema, temperature: 0.1, maxOutputTokens: 8192 }, withThink && think(m) ? { thinkingConfig: think(m) } : {}) });
  const call = (m, withThink) => UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: pay(m, withThink), muteHttpExceptions: true });
  const det = [];
  // hlavní model, při selhání záložní; vždy jen jeden požadavek najednou (záložní se nevolá zbytečně)
  for (const m of GEMINI_MODELS) {
    for (let attempt = 0; attempt < 2; attempt++) {
      const main = m === GEMINI_MODELS[0];
      let r = call(m, true);
      if (r.getResponseCode() === 400 && think(m)) r = call(m, false);
      const code = r.getResponseCode();
      // 503: hlavní model se po 2 s zkusí ještě jednou, u záložních se po 2 s přejde na další; 429 a 404: hned další model
      if (code === 503) { Utilities.sleep(2000); if (main && attempt === 0) continue; }
      if (code !== 200) { det.push(m + ': ' + (code === 429 ? 'kvóta AI vyčerpána (' + quotaWhy_(r.getContentText()) + ')' : code === 404 ? 'model není dostupný' : code === 503 ? 'model je přetížený' : 'HTTP ' + code + ' ' + r.getContentText().replace(/\s+/g, ' ').slice(0, 100))); break; }
      try {
        const j = JSON.parse(r.getContentText());
        if (!j.candidates || !j.candidates[0]) { det.push(m + ': bez odpovědi' + (j.promptFeedback && j.promptFeedback.blockReason ? ' (' + j.promptFeedback.blockReason + ')' : '')); break; }
        const t = ((j.candidates[0].content && j.candidates[0].content.parts) || []).map(p => p.text || '').join('');
        const arr = parseRows_(t);
        if (!arr) { det.push(m + ': odpověď není tabulka: ' + t.replace(/\s+/g, ' ').slice(0, 80)); break; }
        const rows = arr.map(o => ({
          datum: /^\d{4}-\d{2}-\d{2}$/.test(String(o.datum || '')) ? String(o.datum) : '',
          budova: String(o.budova === undefined || o.budova === null ? '' : o.budova).trim(),
          druh: normType_(o.druh),
          cislo: String(o.cislo || '').replace(/\s/g, '').toUpperCase(),
          stav: (o.stav === undefined || o.stav === null || String(o.stav).trim() === '') ? NaN : Number(String(o.stav).replace(/\s/g, '').replace(',', '.'))
        })).filter(o => isFinite(o.stav));
        if (rows.length) return { ok: true, rows: rows, model: m };
        det.push(m + ': žádná čitelná hodnota');
      } catch (e) { det.push(m + ': ' + e.message); }
      break;
    }
  }
  Logger.log('Přepis archu selhal: ' + det.join(' | '));
  return { ok: false, error: 'ai', detail: (det.some(d => /kvóta/.test(d)) ? 'Kvóta AI klíče je vyčerpaná, zkontroluj limity a fakturaci v Google AI Studiu, nebo počkej. ' : '') + det.join(' | ') };
}


// ---------- přehledné okno pro import (menu Měřidla → Import z papíru → Otevřít přehled) ----------
function importUi() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('ImportUi').setWidth(1200).setHeight(780), 'Import z papíru');
}

// sloupce listu podle názvu (chybějící se doplní na konec)
function importCols_(sh) {
  let head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).normalize('NFC').trim().toLowerCase()) : [];
  IMPORT_HEAD.forEach(h => { if (head.indexOf(h.toLowerCase()) < 0) { head.push(h.toLowerCase()); sh.getRange(1, head.length).setValue(h); } });
  const f = n => head.indexOf(n);
  return { n: head.length, date: f('datum'), b: f('budova'), t: f('druh'), num: f('číslo měřidla'), v: f('stav'), note: f('poznámka'), photo: f('fotka archu'), st: f('stav importu'), id: f('měřidlo (id)') };
}

function importUiData() {
  const sh = ensureImportSheet_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(), c = importCols_(sh);
  const data = sh.getLastRow() > 1 ? sh.getRange(1, 1, sh.getLastRow(), c.n).getValues() : [[]];
  const rows = [], arches = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i], g = k => (r[c[k]] === undefined || r[c[k]] === null ? '' : r[c[k]]);
    if (String(g('v')).trim() === '' && String(g('date')).trim() === '' && String(g('num')).trim() === '' && String(g('b')).trim() === '') continue;
    const d = parseImportDate_(g('date'), tz), photo = String(g('photo')).trim();
    if (photo && arches.indexOf(photo) < 0) arches.push(photo);
    rows.push({ row: i + 1, datum: d ? Utilities.formatDate(d, tz, 'd. M. yyyy') : String(g('date')), budova: String(g('b')).trim(), druh: String(g('t')).trim(), cislo: String(g('num')).trim(),
      stav: String(g('v')).trim(), poznamka: String(g('note')), foto: photo, status: String(g('st')), mid: String(g('id')) });
  }
  const ms = meters_();
  const uniq = a => a.filter((x, i) => x && a.indexOf(x) === i);
  return {
    rows: rows, arches: arches,
    meters: ms.map(m => ({ id: m.id, label: (m.buildingLabel || m.building || '?') + ' · ' + m.name + ' · ' + m.id })).sort((a, b) => a.label.localeCompare(b.label, 'cs')),
    buildings: uniq(ms.map(m => String(m.building || '')).concat(['Kotelna'])).sort((a, b) => a.localeCompare(b, 'cs')),
    types: uniq(ms.map(m => String(m.typeName || '')).concat(['Vodoměr', 'Plynoměr', 'Elektroměr', 'TUV']))
  };
}

// uloží úpravy z okna; řádky se nikdy nemažou, vyřazený řádek dostane stav „VYŘAZENO“ a import ho přeskočí
function importUiSave(rows) {
  const sh = ensureImportSheet_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(), c = importCols_(sh);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const last = sh.getLastRow(), old = last > 1 ? sh.getRange(1, 1, last, c.n).getValues() : [[]];
    let saved = 0;
    (rows || []).forEach(x => {
      const isNew = !x.row || x.row < 2 || x.row > last;
      const cur = isNew ? new Array(c.n).fill('') : old[x.row - 1].slice();
      const oldStatus = String(cur[c.st] || '');
      if (/^ZAPSÁNO/.test(oldStatus)) return;
      const d = parseImportDate_(x.datum, tz), vRaw = String(x.stav === undefined || x.stav === null ? '' : x.stav).trim();
      const vNum = vRaw === '' ? '' : Number(vRaw.replace(/\s/g, '').replace(',', '.'));
      const next = { date: d || String(x.datum || '').trim(), b: String(x.budova || '').trim(), t: String(x.druh || '').trim(), num: String(x.cislo || '').trim(), v: (vNum === '' || !isFinite(vNum)) ? vRaw : vNum, note: String(x.poznamka || ''), photo: String(x.foto || '') };
      const same = k => String(cur[c[k]] instanceof Date ? Utilities.formatDate(cur[c[k]], tz, 'yyyy-MM-dd') : cur[c[k]]).trim() === String(next[k] instanceof Date ? Utilities.formatDate(next[k], tz, 'yyyy-MM-dd') : next[k]).trim();
      const changed = isNew || !['date', 'b', 't', 'num', 'v', 'note', 'photo'].every(same);
      Object.keys(next).forEach(k => { cur[c[k]] = next[k]; });
      if (x.exclude) cur[c.st] = 'VYŘAZENO (nezapíše se)';
      else if (changed || /^VYŘAZENO/.test(oldStatus) || !oldStatus) cur[c.st] = 'K ověření';
      if (isNew) { sh.appendRow(cur); const rn = sh.getLastRow(); if (d) sh.getRange(rn, c.date + 1).setNumberFormat('d. M. yyyy'); }
      else { sh.getRange(x.row, 1, 1, c.n).setValues([cur]); if (d) sh.getRange(x.row, c.date + 1).setNumberFormat('d. M. yyyy'); }
      saved++;
    });
    return { ok: true, saved: saved };
  } finally { lock.releaseLock(); }
}

// nahrání fotky archu přímo z okna v tabulce (běží pod účtem správce tabulky, PIN netřeba)
function importUiUpload(photo) { return paperProcess_(photo); }

// seznam fotek archů ve složce Papírové záznamy na Disku (výchozí zdroj v okně)
function importUiFiles() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(), folder = child_(driveRoot_(), PAPER_FOLDER), out = [], it = folder.getFiles();
  while (it.hasNext() && out.length < 300) {
    const f = it.next();
    if (String(f.getMimeType()).indexOf('image/') !== 0) continue;
    out.push({ id: f.getId(), name: f.getName(), t: f.getDateCreated().getTime(), date: Utilities.formatDate(f.getDateCreated(), tz, 'd. M. yyyy H:mm'), kb: Math.round(f.getSize() / 1024) });
  }
  out.sort((a, b) => b.t - a.t);
  return { folderUrl: folder.getUrl(), files: out };
}

// přečte AI fotku ze složky podle id souboru
function importUiRereadId(id) { try { return importUiRereadId_(id); } catch (e) { return { ok: false, error: 'exc', detail: String(e && e.message || e) }; } }
function importUiRereadId_(id) {
  let f; try { f = DriveApp.getFileById(String(id || '')); } catch (e) { return { ok: false, error: 'nofile' }; }
  if (f.getSize() > 15 * 1024 * 1024) return { ok: false, error: 'toobig' };
  return paperRead_(Utilities.base64Encode(f.getBlob().getBytes()), f.getName());
}

// znovu pustí do čtení už uloženou fotku archu (např. když první čtení selhalo na přetížení AI)
function importUiReread(name) { try { return importUiReread_(name); } catch (e) { return { ok: false, error: 'exc', detail: String(e && e.message || e) }; } }
function importUiReread_(name) {
  const it = child_(driveRoot_(), PAPER_FOLDER).getFilesByName(String(name || ''));
  if (!it.hasNext()) return { ok: false, error: 'nofile' };
  return paperRead_(Utilities.base64Encode(it.next().getBlob().getBytes()), String(name));
}

function importUiRun(write) { return { ok: true, text: importRun_(!!write) }; }

function importUiPhoto(name) {
  const it = child_(driveRoot_(), PAPER_FOLDER).getFilesByName(String(name || ''));
  if (!it.hasNext()) return null;
  const b = it.next().getBlob();
  return 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
}
