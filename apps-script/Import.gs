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

function importRun_(write) {
  const sh = ensureImportSheet_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const data = sh.getDataRange().getValues();
  const say = t => { try { SpreadsheetApp.getActive().toast(t, write ? 'Import hotov' : 'Náhled hotov', 20); } catch (e) {} };
  if (data.length < 2) { Logger.log('List „' + SHEET_IMPORT + '“ je prázdný.'); say('List je prázdný.'); return; }
  const head = data[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
  const find = n => head.indexOf(n);
  const c = { date: find('datum'), b: find('budova'), t: find('druh'), n: find('číslo měřidla'), v: find('stav'), note: find('poznámka'), photo: find('fotka archu'), st: find('stav importu'), id: find('měřidlo (id)') };
  if (c.date < 0 || c.v < 0) { Logger.log('V listu chybí sloupec Datum nebo Stav.'); say('V listu chybí sloupec Datum nebo Stav.'); return; }
  if (c.st < 0) { c.st = head.length; head.push('stav importu'); sh.getRange(1, c.st + 1).setValue('Stav importu'); }
  if (c.id < 0) { c.id = head.length; head.push('měřidlo (id)'); sh.getRange(1, c.id + 1).setValue('Měřidlo (id)'); }
  const meters = meters_(), cnt = { ok: 0, wrote: 0, dup: 0, err: 0, skip: 0 };
  const lock = LockService.getScriptLock();
  if (write) lock.waitLock(30000);
  try {
    for (let i = 1; i < data.length; i++) {
      const r = data[i], rowNo = i + 1, cell = k => (c[k] >= 0 && r[c[k]] !== undefined ? r[c[k]] : '');
      if (/^ZAPSÁNO/.test(String(cell('st')))) { cnt.skip++; continue; }
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
            const res = backfillCore_(f.m, when, val, note, 'Správce', false);
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
    ? 'Zapsáno: ' + cnt.wrote + ', už zapsáno: ' + cnt.dup + ', chyby: ' + cnt.err + (cnt.skip ? ', dříve zapsané přeskočeno: ' + cnt.skip : '')
    : 'Zapíše se: ' + cnt.ok + ', chyby: ' + cnt.err + (cnt.skip ? ', dříve zapsané: ' + cnt.skip : '') + '. Výsledek je ve sloupci Stav importu.';
  Logger.log(text);
  say(text);
}

function importPreview() { importRun_(false); }
function importWrite() { importRun_(true); }

// ---------- foto papírového archu (správce v appce): uložení na Disk + AI přepis do listu ----------
function paperUpload_(body) {
  const chk = adminCheck_(body.pin);
  if (!chk.ok) return chk;
  if (!body.photo) return { ok: false, error: 'data' };
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const name = 'arch-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss') + '.jpg';
  try { child_(driveRoot_(), PAPER_FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', name)); }
  catch (e) { return { ok: false, error: 'drive' }; }
  const ai = paperTranscribe_(body.photo);
  if (!ai.ok) return { ok: true, name: name, count: 0, aiError: ai.error, detail: ai.detail || '' };
  const sh = ensureImportSheet_();
  ai.rows.forEach(r => {
    let d = '';
    if (r.datum) { try { d = Utilities.parseDate(r.datum + ' 12:00', tz, 'yyyy-MM-dd HH:mm'); } catch (e) {} }
    sh.appendRow([d, r.budova, r.druh, r.cislo, r.stav, 'AI z archu, ověř', name, 'K ověření', '']);
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

function paperTranscribe_(b64) {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  const models = geminiModels_();
  if (!models.length) return { ok: false, error: 'nomodel' };
  let known = '';
  try { known = meters_().slice(0, 80).map(m => (m.building || '?') + ' ' + (m.typeName || m.name) + ' ' + m.id).join('; '); } catch (e) {}
  const prompt = 'Na fotce je papírový arch (tabulka) s ručně psanými odečty měřidel v bytovém domě. Přečti z něj všechny řádky s hodnotou, ať má arch jakékoli rozložení sloupců. Typické rozložení: Č.P. (číslo popisné domu 271 až 275; řádky se značkou /2, /4, /6, /8, /10 patří k domu uvedenému těsně nad nimi), plynoměr a vodoměr v m3, číslo elektroměru, stav elektroměru v kWh a datum. Ručně napsané datum se může táhnout přes víc řádků a platí pro všechny řádky vedle něj. Řádky kolem nápisu KOTELNA: elektroměry patří k domu 271, plynoměry, vodoměr a TUV (teplá užitková voda) patří kotelně. ' +
    (known ? 'Známá měřidla (budova, druh, id): ' + known + '. ' : '') +
    'Vrať JSON pole, jeden objekt pro každou přečtenou hodnotu: {"datum":"YYYY-MM-DD" nebo null,"budova":"271".."275" nebo "Kotelna","druh":"Vodoměr" nebo "Plynoměr" nebo "Elektroměr" nebo "TUV","cislo":"číslo elektroměru bez mezer, např. W060073, jinak prázdný řetězec","stav":číslo s desetinnou tečkou}. Hodnoty nevymýšlej, nečitelné vynech.';
  const payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0, maxOutputTokens: 8192 } });
  const reqs = models.map(m => ({
    url: 'https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', method: 'post',
    contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: payload, muteHttpExceptions: true }));
  const det = [];
  for (let attempt = 0; attempt < 2; attempt++) {
    const res = UrlFetchApp.fetchAll(reqs);
    for (let i = 0; i < res.length; i++) {
      const r = res[i], code = r.getResponseCode();
      if (code !== 200) { det.push(models[i] + ' HTTP ' + code + ': ' + r.getContentText().replace(/\s+/g, ' ').slice(0, 110)); continue; }
      try {
        const j = JSON.parse(r.getContentText());
        if (!j.candidates || !j.candidates[0]) { det.push(models[i] + ': bez odpovědi' + (j.promptFeedback && j.promptFeedback.blockReason ? ' (' + j.promptFeedback.blockReason + ')' : '')); continue; }
        const t = ((j.candidates[0].content && j.candidates[0].content.parts) || []).map(p => p.text || '').join('');
        const arr = parseRows_(t);
        if (!arr) { det.push(models[i] + ': odpověď není tabulka: ' + t.replace(/\s+/g, ' ').slice(0, 80)); continue; }
        const rows = arr.map(o => ({
          datum: /^\d{4}-\d{2}-\d{2}$/.test(String(o.datum || '')) ? String(o.datum) : '',
          budova: String(o.budova === undefined || o.budova === null ? '' : o.budova).trim(),
          druh: normType_(o.druh),
          cislo: String(o.cislo || '').replace(/\s/g, '').toUpperCase(),
          stav: (o.stav === undefined || o.stav === null || String(o.stav).trim() === '') ? NaN : Number(String(o.stav).replace(/\s/g, '').replace(',', '.'))
        })).filter(o => isFinite(o.stav));
        if (rows.length) return { ok: true, rows: rows };
        det.push(models[i] + ': žádná čitelná hodnota');
      } catch (e) { det.push(models[i] + ': ' + e.message); }
    }
    if (det.some(d => /HTTP (429|5\d\d)/.test(d))) Utilities.sleep(2500); else break;
  }
  Logger.log('Přepis archu selhal: ' + det.join(' | '));
  return { ok: false, error: 'ai', detail: det.slice(0, 4).join(' | ') };
}
