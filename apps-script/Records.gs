// ---------- záznamy odečtů: úprava, vyřazení, přečtení AI znovu a zámek měřidla ----------
// Pravidla:
//  - autor záznamu ho smí upravit do 24 hodin od času odečtu
//  - správce (PIN) smí kdykoli upravit, vyřadit nebo znovu pustit AI
//  - vyřazený záznam zůstane v tabulce (sloupec Vyřazeno, červené pozadí), ale nepočítá se
//  - při focení / zápisu je měřidlo pro ostatní zablokované (zámek v cache vyprší sám za 10 minut)

const EDIT_WINDOW_MS = 24 * 3600 * 1000;
const LOCK_TTL_S = 600;
const COL_OUT = 14, COL_EDIT = 15, COL_SUP = 16, COL_ENTRY = 17, SUP_BG = '#e3efff';

function nowText_() { return Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'd. M. yyyy H:mm'); }

function ensureReadingCols_(sh) {
  if (!sh.getRange(1, COL_OUT).getValue()) sh.getRange(1, COL_OUT).setValue('Vyřazeno');
  if (!sh.getRange(1, COL_EDIT).getValue()) sh.getRange(1, COL_EDIT).setValue('Upraveno');
  if (!sh.getRange(1, COL_SUP).getValue()) sh.getRange(1, COL_SUP).setValue('Odečet dodavatele');
  if (!sh.getRange(1, COL_ENTRY).getValue()) sh.getRange(1, COL_ENTRY).setValue('Zapsáno');
}

// záznam podle ID záznamu; starší záznamy bez ID se hledají podle čísla řádku ("row-12")
function findReading_(sh, rid) {
  rid = String(rid || '');
  if (!rid) return null;
  const data = sh.getDataRange().getValues();
  let i = -1;
  if (/^r\d+$/.test(rid)) { i = Number(rid.slice(1)) - 1; if (i < 1 || i >= data.length || !data[i][0]) i = -1; }
  else i = data.findIndex((r, k) => k > 0 && String(r[12]) === rid);
  return i < 0 ? null : { n: i + 1, r: data[i] };
}

// přepočítá "Předchozí" a "Přírůstek" u všech záznamů měřidla (vyřazené a bez hodnoty se přeskočí)
function relink_(sh, meterId) {
  const data = sh.getDataRange().getValues(), rows = [];
  data.forEach((r, i) => { if (i > 0 && String(r[1]) === String(meterId) && r[3] !== '' && r[3] !== null && isFinite(Number(r[3])) && !r[COL_OUT - 1]) rows.push({ n: i + 1, t: new Date(r[0]).getTime(), v: Number(r[3]), e: r[4], f: r[5] }); });
  rows.sort((a, b) => a.t - b.t);
  let prev = null;
  rows.forEach(x => {
    const e = prev === null ? '' : prev, f = prev === null ? '' : Math.round((x.v - prev) * 1000) / 1000;
    if (x.e !== e || x.f !== f) sh.getRange(x.n, 5, 1, 2).setValues([[e, f]]);
    prev = x.v;
  });
}

// smí osoba záznam upravit? autor do 24 h, jinak správce s PINem
function recordAuth_(P, who, rec, body) {
  const meter = meters_().find(m => m.id === String(rec.r[1]));
  if (!meter || !canEdit_(P, meter)) return { ok: false, error: 'forbidden' };
  // okno 24 h se počítá od času odečtu, u odečtu dodavatele (zadává se zpětným datem z faktury) od chvíle zadání
  const entered = rec.r[COL_ENTRY - 1] ? new Date(rec.r[COL_ENTRY - 1]).getTime() : NaN;
  const age = Date.now() - (isNaN(entered) ? new Date(rec.r[0]).getTime() : entered);
  if (String(rec.r[9]) === who && age <= EDIT_WINDOW_MS) return { ok: true, meter: meter, admin: false };
  if (!body.pin) return { ok: false, error: String(rec.r[9]) === who ? 'expired' : 'forbidden' };
  const chk = adminCheck_(body.pin);
  return chk.ok ? { ok: true, meter: meter, admin: true } : chk;
}

function applyValue_(sh, rec, meter, value, who, how, extraNote) {
  const n = rec.n, old = rec.r[3];
  let note = String(rec.r[6] || '').replace(/^POZOR: nižší než minule\.\s*/, '').trim();
  note = (note + ' ' + how + ': ' + who + ' ' + nowText_() + (old !== '' && old !== null ? ' (bylo ' + old + ')' : '') + (extraNote ? ' – ' + String(extraNote).slice(0, 120) : '')).trim().slice(0, 400);
  sh.getRange(n, 4).setValue(value);
  sh.getRange(n, 7).setValue(note);
  if (how === 'AI znovu') sh.getRange(n, 9).setValue('AI (znovu)');
  else if (!rec.r[8] || rec.r[8] === 'bez hodnoty') sh.getRange(n, 9).setValue('ručně');
  sh.getRange(n, 12).setValue('');
  sh.getRange(n, COL_EDIT).setValue(how + ': ' + who + ' ' + nowText_());
  relink_(sh, meter.id);
  bustCache_(meter.id);
}

function editReading_(body, P, who) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS); ensureReadingCols_(sh);
    const rec = findReading_(sh, body.rid);
    if (!rec) return { ok: false, error: 'nofile' };
    if (rec.r[COL_OUT - 1]) return { ok: false, error: 'excluded' };
    const auth = recordAuth_(P, who, rec, body);
    if (!auth.ok) return auth;
    const raw = String(body.value === undefined || body.value === null ? '' : body.value).replace(/\s/g, '').replace(',', '.');
    const value = raw === '' ? NaN : round_(Number(raw), auth.meter.decimals);
    if (!isFinite(value)) return { ok: false, error: 'data' };
    if (typeof body.supplier === 'boolean' && body.supplier && !auth.meter.supplier) return { ok: false, error: 'nosupplier' };
    let newWhen = null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(body.when || ''))) {
      const willBeSup = typeof body.supplier === 'boolean' ? body.supplier : String(rec.r[COL_SUP - 1] || '') !== '';
      if (!willBeSup) return { ok: false, error: 'data' };
      try { newWhen = Utilities.parseDate(body.when + ' 12:00', SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'); } catch (e) {}
      if (!newWhen || newWhen.getFullYear() < 2000) return { ok: false, error: 'date' };
      if (newWhen.getTime() > Date.now() + 36e5) return { ok: false, error: 'future' };
    }
    applyValue_(sh, rec, auth.meter, value, who, auth.admin ? 'Upraveno správcem' : 'Upraveno', body.note);
    if (newWhen) { sh.getRange(rec.n, 1).setValue(newWhen); relink_(sh, auth.meter.id); bustCache_(auth.meter.id); }
    if (typeof body.supplier === 'boolean') {
      sh.getRange(rec.n, COL_SUP).setValue(body.supplier ? 'ANO' : '');
      sh.getRange(rec.n, 1, 1, COL_SUP).setBackground(body.supplier ? SUP_BG : null);
      bustCache_(auth.meter.id);
    }
    return { ok: true, value: value };
  } finally { lock.releaseLock(); }
}

// jen správce: označí záznam jako vyřazený (nebo ho vrátí, když restore = true)
function excludeReading_(body, P, who) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS); ensureReadingCols_(sh);
    const rec = findReading_(sh, body.rid);
    if (!rec) return { ok: false, error: 'nofile' };
    const meter = meters_().find(m => m.id === String(rec.r[1]));
    if (!meter || !canEdit_(P, meter)) return { ok: false, error: 'forbidden' };
    const chk = adminCheck_(body.pin); if (!chk.ok) return chk;
    const range = sh.getRange(rec.n, 1, 1, COL_EDIT);
    if (body.restore) { sh.getRange(rec.n, COL_OUT).setValue(''); range.setBackground(null); }
    else { sh.getRange(rec.n, COL_OUT).setValue('VYŘAZENO · ' + who + ' · ' + nowText_() + (body.reason ? ' · ' + String(body.reason).slice(0, 120) : '')); range.setBackground('#fde2e1'); }
    relink_(sh, meter.id);
    bustCache_(meter.id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// jen správce: znovu přečte hodnotu AI z uložené fotky záznamu
function rereadReading_(body, P, who) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS); ensureReadingCols_(sh);
  let rec = findReading_(sh, body.rid);
  if (!rec) return { ok: false, error: 'nofile' };
  if (rec.r[COL_OUT - 1]) return { ok: false, error: 'excluded' };
  const meter = meters_().find(m => m.id === String(rec.r[1]));
  if (!meter || !canEdit_(P, meter)) return { ok: false, error: 'forbidden' };
  const chk = adminCheck_(body.pin); if (!chk.ok) return chk;
  const fid = driveId_(rec.r[7]);
  if (!fid) return { ok: false, error: 'nophoto' };
  let bytes;
  try { const f = DriveApp.getFileById(fid); if (f.getSize() > 15 * 1024 * 1024) return { ok: false, error: 'toobig' }; bytes = f.getBlob().getBytes(); } catch (e) { return { ok: false, error: 'nophoto' }; }
  const ai = readWithAi_(meter, Utilities.base64Encode(bytes));
  if (!ai.ok) return { ok: false, error: 'ai', detail: ai.detail || ai.error };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    rec = findReading_(sh, body.rid); // mezitím se mohl změnit
    if (!rec) return { ok: false, error: 'nofile' };
    const value = round_(Number(ai.value), meter.decimals);
    applyValue_(sh, rec, meter, value, who, 'AI znovu', '');
    return { ok: true, value: value };
  } finally { lock.releaseLock(); }
}

// ---------- zámek měřidla (kdo právě fotí / zapisuje) ----------
const lockKey_ = id => 'L:' + id;
function meterLock_(body, who) {
  const cache = CacheService.getScriptCache(), key = lockKey_(body.id);
  let cur = null; try { cur = JSON.parse(cache.get(key) || 'null'); } catch (e) {}
  if (body.release) { if (cur && cur.by === who) cache.remove(key); return { ok: true }; }
  if (cur && cur.by !== who) return { ok: false, error: 'locked', by: cur.by };
  cache.put(key, JSON.stringify({ by: who, t: Date.now() }), LOCK_TTL_S);
  return { ok: true };
}
// kdo drží zámek u kterého měřidla (jen cizí zámky)
function lockedMap_(ids, who) {
  const out = {};
  try {
    const got = CacheService.getScriptCache().getAll(ids.map(lockKey_));
    ids.forEach(id => { const v = got[lockKey_(id)]; if (v) { const c = JSON.parse(v); if (c.by !== who) out[id] = c.by; } });
  } catch (e) {}
  return out;
}

// ---------- archivace starých záznamů (menu) ----------
// Záznamy starší než 2 roky se přesunou do listu Odečty_archiv (nic se nemaže). U každého měřidla zůstane v Odečtech nejnovější
// záznam z doby před hranicí, aby navazovaly přírůstky. V aplikaci se pak ukáže jen novější historie.
function archiveOldReadings() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SHEET_READINGS);
  const data = sh.getDataRange().getValues(), ncol = data[0].length, cut = new Date();
  cut.setMonth(cut.getMonth() - 24);
  const newest = {};
  data.forEach((r, i) => { if (i > 0 && r[0] && new Date(r[0]) < cut && !r[COL_OUT - 1]) { const k = String(r[1]); if (newest[k] === undefined || new Date(r[0]) > new Date(data[newest[k]][0])) newest[k] = i; } });
  const keep = {}; Object.keys(newest).forEach(k => { keep[newest[k]] = true; });
  const move = []; data.forEach((r, i) => { if (i > 0 && r[0] && new Date(r[0]) < cut && !keep[i]) move.push(i); });
  const when = Utilities.formatDate(cut, ss.getSpreadsheetTimeZone(), 'd. M. yyyy');
  if (!move.length) { ui.alert('Archivace', 'Nic staršího než ' + when + ' k archivaci není.', ui.ButtonSet.OK); return; }
  if (ui.alert('Archivace', 'Přesunout ' + move.length + ' záznamů starších než ' + when + ' do listu Odečty_archiv? Nic se nemaže, záznamy budou v archivu. V aplikaci se pak ukáže jen novější historie.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    let arch = ss.getSheetByName('Odečty_archiv');
    if (!arch) { arch = ss.insertSheet('Odečty_archiv'); arch.getRange(1, 1, 1, ncol).setValues([data[0]]); arch.setFrozenRows(1); }
    arch.getRange(arch.getLastRow() + 1, 1, move.length, ncol).setValues(move.map(i => data[i].slice(0, ncol)));
    // řádky se z Odečtů odstraní až po zápisu do archivu, od konce po souvislých blocích
    let end = move.length - 1;
    while (end >= 0) { let start = end; while (start > 0 && move[start - 1] === move[start] - 1) start--; sh.deleteRows(move[start] + 1, end - start + 1); end = start - 1; }
    bustAll_();
  } finally { lock.releaseLock(); }
  ui.alert('Hotovo', 'Přesunuto ' + move.length + ' záznamů do listu Odečty_archiv.', ui.ButtonSet.OK);
}
