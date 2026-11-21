/**
 * Příjem odečtů měřidel: fotka -> složka na Disku, hodnota -> tabulka, upozornění e-mailem.
 *
 * Jednorázově: spusť funkci setup() (založí listy, složku a tajný kód).
 * Potom: Nasadit -> Nová nasazení -> Webová aplikace (Spustit jako: já, Přístup: kdokoli).
 *
 * Aplikace musí POST posílat s Content-Type: text/plain (jinak prohlížeč vyžaduje preflight,
 * který Apps Script neumí).
 */

// verze skriptu (zvyšuje se při změně kódu); ověření: adresa nasazení + ?action=version
const SCRIPT_VERSION = '1.41.1';
const SHEET_METERS = 'Měřidla';
const SHEET_READINGS = 'Odečty';
const SHEET_PEOPLE = 'Lidé';
const SHEET_TYPES = 'Typy';
const SHEET_ALERTS = 'Upozornění';
const SHEET_BUILDINGS = 'Budovy';
const SHEET_PRICES = 'Ceny';
const FOLDER_NAME = 'Fotodokumentace odečtů';
const PLACES_NAME = 'Místa měřidel – plánky a fotky';
const DRIVE_ROOT = 'SVJ_app';
const DEFAULT_APP_URL = 'https://pietrodetreville-cyber.github.io/meridla/app/';

// ---------- jednorázové nastavení ----------
function setup() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  // typy měřidel: společné vlastnosti (jednotka, nápověda pro čtení, ikona, výchozí interval)
  let ty = ss.getSheetByName(SHEET_TYPES);
  if (!ty) {
    ty = ss.insertSheet(SHEET_TYPES);
    ty.appendRow(['Typ', 'Jednotka', 'Zobrazení', 'Číslic před čárkou', 'Nápověda pro čtení', 'Výchozí interval (dny)', 'Ikona']);
    ty.appendRow(['Elektroměr', 'kWh', 'digitální', 5, 'Přečti celkový stav v kWh před desetinnou čárkou. Číslice za čárkou (červené nebo v rámečku) ignoruj. Když displej střídá více hodnot, vezmi celkovou spotřebu (1.8.0).', 7, '⚡']);
    ty.appendRow(['Vodoměr', 'm³', 'kolečka', 5, 'Přečti černá kolečka (celé m³). Červená kolečka vpravo jsou desetiny, ty ignoruj.', 7, '💧']);
    ty.appendRow(['Plynoměr', 'm³', 'kolečka', 5, 'Přečti černá kolečka (celé m³). Červená kolečka vpravo jsou desetiny, ty ignoruj.', 7, '🔥']);
    ty.appendRow(['Kalorimetr', 'GJ', 'digitální', 5, 'Přečti celkovou naměřenou energii (spotřebu tepla) v jednotce uvedené na displeji. Ignoruj teploty, průtok, objem a výkon.', 7, '🌡️']);
    ty.setFrozenRows(1);
  }
  // budovy: vztah "je součástí" (např. kotelna je součástí budovy 271)
  let bd = ss.getSheetByName(SHEET_BUILDINGS);
  if (!bd) {
    bd = ss.insertSheet(SHEET_BUILDINGS);
    bd.appendRow(['Budova', 'Součást budovy', 'Popis']);
    bd.appendRow(['Kotelna', '271', 'Kotelna je součástí budovy 271']);
    bd.setFrozenRows(1);
  }
  // upozornění e-mailem, když není dodržen interval
  let al = ss.getSheetByName(SHEET_ALERTS);
  if (!al) {
    al = ss.insertSheet(SHEET_ALERTS);
    al.appendRow(['Komu', 'Budova', 'Dní po termínu', 'Poznámka']);
    al.appendRow([props.getProperty('NOTIFY') || Session.getEffectiveUser().getEmail(), '', 0, 'správce: dostane upozornění v den termínu']);
    al.appendRow(['', '', 3, 'sem doplňte e-mail výboru: upozornění po 3 dnech zpoždění']);
    al.setFrozenRows(1);
  }
  // měřidla: u každého stačí vybrat typ a budovu; id se doplní samo
  let m = ss.getSheetByName(SHEET_METERS);
  if (!m) {
    m = ss.insertSheet(SHEET_METERS);
    m.appendRow(['id', 'Název', 'Typ', 'Budova', 'Interval (dny)', 'Umístění', 'Nápověda_místo', 'Nápověda_focení', 'Zaokrouhlit_desetinné_místo']);
    ['271', '272', '273', '274', '275'].forEach(b => ['Elektroměr', 'Vodoměr', 'Plynoměr'].forEach(t => m.appendRow(['', t, t, b, '', ''])));
    ['Elektroměr', 'Vodoměr', 'Plynoměr', 'Kalorimetr'].forEach(t => m.appendRow(['', t, t, 'Kotelna', '', '']));
    m.setFrozenRows(1);
    m.getRange('C2:C500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ty.getRange('A2:A50'), true).setAllowInvalid(false).build());
    m.getRange('D2:D500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['271', '272', '273', '274', '275', 'Kotelna'], true).setAllowInvalid(true).build());
    fillIds_();
  }
  let r = ss.getSheetByName(SHEET_READINGS);
  if (!r) {
    r = ss.insertSheet(SHEET_READINGS);
    r.appendRow(['Čas', 'id', 'Měřidlo', 'Stav', 'Předchozí', 'Přírůstek', 'Poznámka', 'Fotka', 'Zdroj']);
    r.setFrozenRows(1);
  }

  // složka pro plánky a fotky míst (do ní se obrázky vkládají, v tabulce se uvede název souboru)
  if (!props.getProperty('PLACES_FOLDER_ID')) props.setProperty('PLACES_FOLDER_ID', child_(driveRoot_(), PLACES_NAME).getId());
  const mHead = m.getRange(1, 1, 1, Math.max(1, m.getLastColumn())).getValues()[0];
  if (!mHead.some(x => /^(fotka|plánek)[_ ]?místa/i.test(String(x).trim()))) {
    let c = mHead.findIndex(x => String(x).trim() === '');
    m.getRange(1, (c < 0 ? mHead.length : c) + 1).setValue('Fotka_místa');
  }
  Logger.log('Složka na plánky a fotky míst: ' + DriveApp.getFolderById(props.getProperty('PLACES_FOLDER_ID')).getUrl());
  // číslo měřidla jako id: vždy jako text (aby se neztratily nuly na začátku)
  r.getRange('B:B').setNumberFormat('@');
  m.getRange('A:A').setNumberFormat('@');
  let p = ss.getSheetByName(SHEET_PEOPLE);
  if (!p) {
    p = ss.insertSheet(SHEET_PEOPLE);
    p.appendRow(['Jméno', 'Tajný kód (negenerovat ručně)', 'Odkaz pro tuto osobu', 'Budovy', 'Oprávnění']);
    p.setFrozenRows(1);
  }
  if (!r.getRange(1, 10).getValue()) r.getRange(1, 10).setValue('Kdo');
  if (!r.getRange(1, 11).getValue()) r.getRange(1, 11).setValue('Budova');
  if (!r.getRange(1, 12).getValue()) r.getRange(1, 12).setValue('Chyba');
  if (!r.getRange(1, 13).getValue()) r.getRange(1, 13).setValue('ID záznamu');
  if (!al.getRange(1, 5).getValue()) {
    al.getRange(1, 5).setValue('Chybí hodnota');
    if (al.getLastRow() >= 2 && String(al.getRange(2, 1).getValue()).indexOf('@') > 0) al.getRange(2, 5).setValue('ano');
  }
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', child_(driveRoot_(), FOLDER_NAME).getId());
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  if (!props.getProperty('NOTIFY')) props.setProperty('NOTIFY', Session.getEffectiveUser().getEmail());

  Logger.log('TAJNÝ KÓD (token): ' + props.getProperty('TOKEN'));
  Logger.log('Upozornění chodí na: ' + props.getProperty('NOTIFY'));
  Logger.log('Složka s fotkami: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
}

// sloupec "Dodavatel" v listu Měřidla (ano = pod plombou dodavatele, prázdné = interní); založí se při otevření tabulky
function ensureSupplierCol_() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_METERS);
    if (!sh || sh.getLastColumn() < 1) return;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
    const hi = head.findIndex(h => /^(dodavatel|plomb)/.test(h));
    if (hi < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue('Dodavatel'); bustMeters_(); }
    else if (head[hi] !== 'dodavatel') { sh.getRange(1, hi + 1).setValue('Dodavatel'); bustMeters_(); } // starý název "Dodavatel (plomba)" -> "Dodavatel"
  } catch (e) {}
}
// podmíněné formátování listu Měřidla: zelený řádek = pod plombou dodavatele ("ano"), žlutý = interní (prázdné pole). Prázdné řádky bez id se nebarví.
// Pravidla se ukládají podle písmena sloupce, po přesunu sloupce Dodavatel stačí spustit menu Měřidla -> Obarvit měřidla znovu.
function formatMeterRows_() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_METERS);
    if (!sh || sh.getLastColumn() < 1) return;
    const lastCol = sh.getLastColumn(), head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
    const ci = head.findIndex(h => /^(dodavatel|plomb)/.test(h)), ii = head.indexOf('id');
    if (ci < 0 || ii < 0) return;
    const letter = n => { let t = '', x = n; while (x > 0) { const m = (x - 1) % 26; t = String.fromCharCode(65 + m) + t; x = Math.floor((x - 1) / 26); } return t; };
    const sup = '$' + letter(ci + 1) + '2', idc = '$' + letter(ii + 1) + '2';
    const fGreen = '=' + sup + '="ano"', fYellow = '=(' + idc + '<>"")*(' + sup + '="")';
    const range = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), lastCol);
    // vlastní pravidla se poznají podle barvy a typu (vlastní vzorec); cizí pravidla zůstanou
    const isOurs = r => { try { const c = r.getBooleanCondition(); return !!c && c.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA && (c.getBackground() === '#d9ead3' || c.getBackground() === '#fff2cc'); } catch (e) { return false; } };
    const keep = sh.getConditionalFormatRules().filter(r => !isOurs(r));
    const mk = (f, bg) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(bg).setRanges([range]).build();
    sh.setConditionalFormatRules(keep.concat([mk(fGreen, '#d9ead3'), mk(fYellow, '#fff2cc')]));
  } catch (e) { Logger.log('Obarvení měřidel: ' + e.message); }
}
function formatMeterRows() { formatMeterRows_(); try { SpreadsheetApp.getActive().toast('Měřidla jsou obarvená: zelená = dodavatel, žlutá = interní.', 'Hotovo', 6); } catch (e) {} }
function onOpen() {
  ensureSupplierCol_(); formatMeterRows_();
  SpreadsheetApp.getUi().createMenu('Měřidla').addItem('Vygenerovat odkazy pro lidi', 'generateLinks').addItem('Doplnit chybějící id měřidel', 'fillIds_').addItem('Uspořádat Disk (SVJ_app)', 'organizeDrive').addItem('Vytvořit testovací sadu', 'createTestSet').addSubMenu(SpreadsheetApp.getUi().createMenu('Import z papíru').addItem('Otevřít přehled (doporučeno)', 'importUi').addItem('Náhled (nic nezapisuje)', 'importPreview').addItem('Zapsat hodnoty', 'importWrite')).addItem('Pozastavit / znovu zapnout e-maily', 'toggleMails').addItem('Resetovat záznamy (vymaže list Odečty)', 'resetReadings').addItem('Vymazat list Import z papíru', 'resetImportSheet').addItem('Archivovat staré záznamy (starší než 2 roky)', 'archiveOldReadings').addItem('Obarvit měřidla (dodavatel zeleně, interní žlutě)', 'formatMeterRows').addItem('Zkontrolovat měřidla', 'checkMeters').addItem('Odeslat kontrolu termínů teď', 'reminder').addItem('Zapnout denní kontrolu (8:00)', 'installReminder').addToUi();
}

// V listu "Lidé" vyplň jména (sloupec A) a spusť: doplní kódy a hotové odkazy.
function generateLinks() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PEOPLE);
  if (!sh) { SpreadsheetApp.getUi().alert('Nejdřív spusť funkci setup.'); return; }
  const base = props_().getProperty('APP_URL') || DEFAULT_APP_URL;
  const n = sh.getLastRow();
  if (n < 2) return;
  const rows = sh.getRange(2, 1, n - 1, 3).getValues();
  const out = rows.map(r => {
    if (!String(r[0]).trim()) return [r[1], r[2]];
    const token = r[1] || Utilities.getUuid().replace(/-/g, '');
    return [token, base + '#k=' + token];
  });
  sh.getRange(2, 2, out.length, 2).setValues(out);
}

// Vygeneruje nový tajný kód (starý přestane platit). Kód se zobrazí jen v protokolu spuštění.
function resetToken() {
  const t = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', t);
  Logger.log('NOVÝ TAJNÝ KÓD (token): ' + t);
}

// ---------- pomocné ----------
function ok_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// testovací kopie skriptu má ve vlastnostech skriptu ENV = test: e-maily dostanou předponu [TEST]
function envTag_() { return PropertiesService.getScriptProperties().getProperty('ENV') === 'test' ? '[TEST] ' : ''; }
function props_() { return PropertiesService.getScriptProperties(); }
// Vrací jméno člověka podle tajného kódu, nebo '' když kód neplatí.
// Hlavní kód (vlastnost TOKEN) patří "Správci", ostatní kódy jsou v listu Lidé.
// ---------- lidé a jejich oprávnění ----------
// list Lidé: Jméno | Tajný kód | Odkaz | Vidí | Edituje
//  Vidí:    budovy, které osoba vidí (např. "271, 273"); prázdné = všechny
//  Edituje: budovy, kde smí zapisovat (např. "Kotelna"); prázdné nebo "všechny" = všechny, které vidí; "nic" nebo "vidět" = nic (jen prohlíží)
// Kotelna patří pod budovu 271 (stejně jako v aplikaci), takže kdo má 271, vidí i Kotelnu.
function normKey_(v) { return String(v === undefined || v === null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase(); }
function personList_(v) { return String(v === undefined || v === null ? '' : v).split(/[,;\n]+/).map(normKey_).filter(Boolean); }
function person_(t) {
  if (!t) return null;
  const master = props_().getProperty('TOKEN');
  if (master && t === master) return { name: 'Správce', buildings: null, editBuildings: null, edit: true, editAll: true };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PEOPLE);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(1, 1, sh.getLastRow(), 5).getValues();
  if (!v[0][3] || v[0][3] === 'Budovy') sh.getRange(1, 4).setValue('Vidí (budovy)');
  if (!v[0][4] || v[0][4] === 'Oprávnění') sh.getRange(1, 5).setValue('Edituje (budovy)');
  const hit = v.slice(1).find(r => String(r[0]).trim() && r[1] && String(r[1]) === t);
  if (!hit) return null;
  const see = personList_(hit[3]), en = normKey_(hit[4]);
  let editList = null; // null = všechny, které vidí
  if (/^(nic|zadn|vid|jen|cten|read|view|pouze|ne$|-$)/.test(en)) editList = [];
  else if (en && !/^(vse|vsech|edit|ano|all)/.test(en)) editList = personList_(hit[4]);
  const buildings = see.length ? see : null;
  return { name: String(hit[0]).trim(), buildings: buildings, editBuildings: editList, edit: editList === null || editList.length > 0, editAll: buildings === null && editList === null };
}
function who_(t) { const p = person_(t); return p ? p.name : ''; }
function inList_(list, m) { return list.indexOf(normKey_(m.building)) >= 0 || (!!m.parent && list.indexOf(normKey_(m.parent)) >= 0); }
function canSee_(p, m) { return !p || !p.buildings || inList_(p.buildings, m); }
function canEdit_(p, m) { return !p || (canSee_(p, m) && (p.editBuildings === null || inList_(p.editBuildings, m))); }

// přečte list jako tabulku podle názvů sloupců (pořadí sloupců není důležité)
function table_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  if (!v.length) return { sheet: sh, rows: [] };
  const head = v[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
  return { sheet: sh, rows: v.slice(1).map((r, i) => { const o = {}; head.forEach((h, k) => { if (h) o[h] = r[k]; }); return { row: i + 2, v: o }; }) };
}

function slug_(x) {
  return String(x).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

// doplní trvalé id u řádků, které ho nemají (např. nově přidané měřidlo)
function fillIds_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_METERS);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    const head = data[0].map(h => String(h).trim().toLowerCase());
    const ci = head.indexOf('id');
    if (ci < 0) return;
    const col = n => head.indexOf(n);
    const used = {};
    data.slice(1).forEach(r => { if (String(r[ci]).trim()) used[String(r[ci]).trim()] = 1; });
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (String(r[ci]).trim()) continue;
      const name = col('název') >= 0 ? String(r[col('název')] || '').trim() : '';
      const type = col('typ') >= 0 ? String(r[col('typ')] || '').trim() : '';
      const b = col('budova') >= 0 ? String(r[col('budova')] || '').trim() : '';
      if (!name && !type) continue;
      const base = slug_([b, type || name].filter(Boolean).join('-'));
      let n = 1, id;
      do { id = base + '-' + n++; } while (used[id]);
      used[id] = 1;
      sh.getRange(i + 1, ci + 1).setValue(id);
    }
    bustMeters_();
  } finally {
    lock.releaseLock();
  }
}

// měřidla: vlastnosti z listu Měřidla, doplněné výchozími hodnotami podle typu z listu Typy
// vztahy mezi budovami z listu Budovy (výchozí: kotelna je součástí budovy 271)
function buildings_() {
  const out = { kotelna: { name: 'Kotelna', parent: '271' } };
  const t = table_(SHEET_BUILDINGS);
  if (t) t.rows.forEach(r => {
    const n = String(r.v['budova'] || '').trim();
    if (n) out[n.toLowerCase()] = { name: n, parent: String(r.v['součást budovy'] || '').trim() };
  });
  return out;
}

function findKey_(v, re) { return Object.keys(v).find(k => re.test(k)); }

function round_(x, d) { const f = Math.pow(10, d); return Math.round(x * f) / f; }

// seznam měřidel se v rámci jednoho požadavku (doGet / doPost) čte z tabulky jen jednou
let MEMO_ON_ = false, METERS_MEMO_ = null, MISSING_IDS_ = false;
// ---------- mezipaměť (CacheService): souhrn posledních hodnot, seznam měřidel a historie měřidla ----------
// Zápis záznamu mezipaměť zneplatní. Ruční úprava přímo v tabulce se projeví nejpozději za dobu platnosti (TTL).
const CACHE_TTL = { meters: 180, sum: 300, hist: 600, prices: 300 };
function cacheGetJson_(k) { try { const v = CacheService.getScriptCache().get(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function cachePutJson_(k, obj, ttl) { try { const t = JSON.stringify(obj); if (t.length < 90000) CacheService.getScriptCache().put(k, t, ttl); } catch (e) {} }
function bustCache_(meterId) { try { CacheService.getScriptCache().removeAll(['SUM2', 'H:' + meterId]); } catch (e) {} }
function bustMeters_() { try { CacheService.getScriptCache().remove('METERS'); } catch (e) {} METERS_MEMO_ = null; }
function bustAll_() { try { const ids = metersRead_().map(m => m.id); CacheService.getScriptCache().removeAll(['SUM2', 'METERS'].concat(ids.map(i => 'H:' + i))); } catch (e) {} METERS_MEMO_ = null; }
function meters_() {
  if (MEMO_ON_ && METERS_MEMO_) return METERS_MEMO_;
  let r = MEMO_ON_ ? cacheGetJson_('METERS') : null;
  if (!r) {
    r = metersRead_();
    if (MEMO_ON_ && !MISSING_IDS_) cachePutJson_('METERS', r, CACHE_TTL.meters); // s chybějícím id se nic neukládá, dokud se id nedoplní
  }
  if (MEMO_ON_) METERS_MEMO_ = r;
  return r;
}
// poslední hodnota a stav "čeká na hodnotu" každého měřidla (z listu Odečty se čte jen při prázdné mezipaměti)
function summary_() {
  const c = cacheGetJson_('SUM2'); if (c) return c;
  const by = {};
  readings_().sort((a, b) => a.time - b.time).forEach(x => {
    const s = by[x.id] || (by[x.id] = { v: null, t: null, w: '', p: false, sv: null, st: null });
    s.p = !x.has;
    if (x.has) { s.v = x.value; s.t = x.time.toISOString(); s.w = x.who; if (x.sup) { s.sv = x.value; s.st = s.t; } }
  });
  cachePutJson_('SUM2', by, CACHE_TTL.sum);
  return by;
}
function metersRead_() {
  const types = {};
  const tt = table_(SHEET_TYPES);
  if (tt) tt.rows.forEach(r => { const k = String(r.v['typ'] || '').trim().toLowerCase(); if (k) types[k] = r.v; });
  const mt = table_(SHEET_METERS);
  const num = x => Number(x) > 0 ? Math.round(Number(x)) : 0;
  const seen = {}, B = buildings_();
  // sloupec "Dodavatel" se při prvním čtení doplní do hlavičky listu Měřidla (ano = pod plombou dodavatele, prázdné = interní měřidlo)
  if (mt.sheet && !Object.keys((mt.rows[0] || { v: {} }).v).some(k => /^(dodavatel|plomb)/.test(k))) ensureSupplierCol_();
  MISSING_IDS_ = mt.rows.some(r => !String(r.v['id'] || '').trim() && (String(r.v['název'] || '').trim() || String(r.v['typ'] || '').trim()));
  return mt.rows.filter(r => String(r.v['id'] || '').trim()).map(r => {
    const v = r.v, typeName = String(v['typ'] || '').trim(), t = types[typeName.toLowerCase()] || {};
    // stejné id u dvou měřidel by slilo jejich historii: druhé a další dostane příponu
    let id = String(v['id']).trim();
    if (seen[id]) { seen[id]++; id = id + '#' + seen[id]; } else seen[id] = 1;
    const dk = findKey_(v, /zaokrouh/), dr = dk ? v[dk] : '';
    const dec = (dr !== '' && dr !== null && isFinite(Number(dr)) && Number(dr) >= 0 && Number(dr) <= 4) ? Math.round(Number(dr)) : 1;
    const pk = findKey_(v, /^nápověda[_ ]místo/), fk = findKey_(v, /^nápověda[_ ]focení/);
    const bRaw = String(v['budova'] || '').trim();
    const bName = bRaw ? bRaw.charAt(0).toUpperCase() + bRaw.slice(1) : '';
    const parent = bName && B[bName.toLowerCase()] ? B[bName.toLowerCase()].parent : '';
    return {
      id: id,
      row: r.row,
      name: String(v['název'] || typeName).trim(),
      typeName: typeName,
      type: String(v['zobrazení'] || t['zobrazení'] || typeName),
      digits: v['číslic před čárkou'] || t['číslic před čárkou'] || '',
      hint: String(v['nápověda pro čtení'] || t['nápověda pro čtení'] || ''),
      unit: String(v['jednotka'] || t['jednotka'] || ''),
      building: bName,
      parent: parent,
      buildingLabel: parent ? bName + ' (budova ' + parent + ')' : bName,
      place: String(v['umístění'] || '').trim(),
      placeHelp: pk ? String(v[pk] || '').trim() : '',
      shootHelp: fk ? String(v[fk] || '').trim() : '',
      placePhoto: (() => { const k = findKey_(v, /^(fotka|plánek|plan)[_ ]?(místa|mista)?/); return k ? String(v[k] || '').trim() : ''; })(),
      decimals: dec,
      supplier: (() => { const k = findKey_(v, /^(dodavatel|plomb)/); return k ? /^(ano|a|1|x|true|yes)$/i.test(String(v[k] === undefined || v[k] === null ? '' : v[k]).trim()) : false; })(),
      icon: String(t['ikona'] || '').trim(),
      interval: num(v['interval (dny)'] || v['interval']) // 0 = rozhodující termín poslední den v měsíci; vyplněný interval (dny) má přednost
    };
  });
}

function readings_(withOut) {
  const v = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS).getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ r: r, row: i + 2 })).filter(o => o.r[0] && (withOut || !o.r[13])).map(o => {
    const r = o.r, has = r[3] !== '' && r[3] !== null && isFinite(Number(r[3]));
    return { time: new Date(r[0]), id: String(r[1]), value: has ? Number(r[3]) : NaN, has: has, who: String(r[9] || ''), rid: String(r[12] || ''), row: o.row, src: String(r[8] || ''), photo: String(r[7] || ''), sup: String(r[15] || '') !== '', ent: r[16] ? new Date(r[16]) : null };
  });
}

// číslo dne (podle časové zóny tabulky), aby se počítaly kalendářní dny
function dayNum_(d) {
  const p = Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-').map(Number);
  return Math.floor(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5);
}
// ---------- rozhodující termín odečtu: poslední den v měsíci ----------
// Platí pro měřidla bez vyplněného pole Interval (dny); vyplněný interval měřidla má přednost (denní rytmus od posledního odečtu).
// Odečty mezi termíny se zapisují, ale cyklus upozornění neovlivňují.
// Termín za měsíc je splněný odečtem v tento poslední den, nebo prvním odečtem po něm. Měřidlo je k odečtu (due),
// dokud od posledního uplynulého termínu (nebo dne, kdy termín je) nebyl žádný odečet s hodnotou.
const dayStr_ = dn => { const d = new Date(dn * 864e5); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0'); };
// poslední den měsíce, ve kterém leží den dn (dn = číslo dne od 1. 1. 1970)
function monthEnd_(dn) { const d = new Date(dn * 864e5); return Math.floor(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0) / 864e5); }
// nejnovější termín, který už nastal (dnes včetně): poslední den tohoto měsíce, pokud už je dnes, jinak poslední den předchozího měsíce
function lastTerm_(dn) {
  const e = monthEnd_(dn);
  if (dn >= e) return e;
  const d = new Date(dn * 864e5);
  return Math.floor(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 0) / 864e5);
}
// stav měřidla: due = chybí odečet za poslední nastalý termín; term/nextTerm jako yyyy-MM-dd
function dueInfo_(m, last) {
  const today = dayNum_(new Date());
  // měřidlo s vyplněným intervalem: dosavadní denní rytmus od posledního odečtu
  if (m && m.interval > 0) {
    if (!last) return { mode: 'interval', key: 'none', due: true, overdueDays: 0, remainingDays: 0, daysSince: null };
    const ds = today - dayNum_(last.time);
    return { mode: 'interval', key: last.time.toISOString(), due: ds >= m.interval, overdueDays: Math.max(0, ds - m.interval), remainingDays: Math.max(0, m.interval - ds), daysSince: ds };
  }
  const term = lastTerm_(today), next = monthEnd_(term + 1);
  const base = { mode: 'month', key: dayStr_(term), term: dayStr_(term), nextTerm: dayStr_(next) };
  if (!last) return Object.assign(base, { due: true, overdueDays: 0, remainingDays: 0, daysSince: null });
  const done = dayNum_(last.time) >= term;
  return Object.assign(base, { due: !done, overdueDays: done ? 0 : today - term, remainingDays: done ? next - today : 0, daysSince: today - dayNum_(last.time) });
}

// začátek aktuálního týdne (pondělí 0:00 v časové zóně tabulky)
function weekStart_() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const dow = (Number(Utilities.formatDate(now, tz, 'u')) + 6) % 7; // 0 = pondělí
  const midnight = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd'T'00:00:00"));
  return new Date(midnight.getTime() - dow * 864e5);
}

// ---------- seznam měřidel pro aplikaci ----------
function doGet(e) {
  MEMO_ON_ = true; METERS_MEMO_ = null; const T0 = Date.now();
  if (e.parameter.action === 'version') return ok_({ ok: true, version: SCRIPT_VERSION, env: props_().getProperty('ENV') === 'test' ? 'test' : 'prod' });
  if (e.parameter.action === 'form') return formPage_(e.parameter.rid, e.parameter.key);
  const P = person_(e.parameter.token), who = P ? P.name : '';
  if (!who) return ok_({ ok: false, error: 'token' });
  if (e.parameter.action === 'history') {
    const hm = meters_().find(m => m.id === e.parameter.id);
    if (!hm || !canSee_(P, hm)) return ok_({ ok: false, error: 'forbidden' });
    let hd = cacheGetJson_('H:' + hm.id);
    if (!hd) {
      const mine = readings_().filter(x => x.id === hm.id).sort((a, b) => a.time - b.time);
      hd = { r: mine.filter(x => x.has).slice(-800).map(x => ({ rid: x.rid || 'row-' + x.row, time: x.time.toISOString(), value: x.value, who: x.who, src: x.src, sup: x.sup, ent: x.ent && !isNaN(x.ent) ? x.ent.toISOString() : '', pid: driveId_(x.photo) })), p: mine.filter(x => !x.has).slice(-20).map(x => ({ rid: x.rid || 'row-' + x.row, time: x.time.toISOString(), who: x.who, pid: driveId_(x.photo) })) };
      cachePutJson_('H:' + hm.id, hd, CACHE_TTL.hist);
    }
    let pr = cacheGetJson_('PR:' + hm.typeName);
    if (!pr) { pr = pricesFor_(hm.typeName); cachePutJson_('PR:' + hm.typeName, pr, CACHE_TTL.prices); }
    return ok_({ ok: true, ms: Date.now() - T0, me: who, readings: hd.r, pending: hd.p, prices: pr });
  }
  if (e.parameter.action === 'rphoto') return ok_(recordImage_(e.parameter.f, P));
  if (e.parameter.action === 'place') {
    const mm = meters_().find(x => x.id === e.parameter.id);
    return ok_(mm && mm.placePhoto && canSee_(P, mm) ? placeImage_(mm.placePhoto) : { ok: false, error: 'none' });
  }
  meters_();
  if (MISSING_IDS_) { fillIds_(); METERS_MEMO_ = null; } // doplnění id jen když nějaké chybí
  const sm = summary_();
  const locks = lockedMap_(meters_().map(m => m.id), who);
  const list = meters_().filter(m => canSee_(P, m)).map(m => {
    const sx = sm[m.id], lastVal = sx && sx.t ? { value: sx.v, time: new Date(sx.t), who: sx.w } : null;
    // fotka bez hodnoty se nepočítá jako splněný odečet, dokud se hodnota nedoplní
    const di = dueInfo_(m, lastVal);
    return { ...m, last: lastVal ? { value: lastVal.value, time: lastVal.time.toISOString(), who: lastVal.who } : null, pending: !!sx && sx.p, lastSup: sx && sx.st ? { value: sx.sv, time: sx.st } : null, due: di.due, term: di.term || '', nextTerm: di.nextTerm || '', overdueDays: di.overdueDays, remainingDays: di.remainingDays, doneThisWeek: !di.due, editable: canEdit_(P, m), lockedBy: locks[m.id] || '' };
  });
  return ok_({ ok: true, ms: Date.now() - T0, who: who, edit: P.edit, editAll: P.editAll, meters: list });
}

// Spusť jednou z editoru: vyvolá povolení externího připojení a ověří klíč Gemini.
function authorize() {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) { Logger.log('Chybí vlastnost GEMINI_KEY (Nastavení projektu -> Vlastnosti skriptu).'); return; }
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=5', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  Logger.log('Gemini odpověděl: HTTP ' + r.getResponseCode() + (r.getResponseCode() === 200 ? ' (klíč funguje)' : ' – ' + r.getContentText().replace(/\s+/g, ' ').slice(0, 200)));
  Logger.log('Vybrané modely: ' + JSON.stringify(geminiModels_()) + (AI_DETAIL ? ' / ' + AI_DETAIL : ''));
}

// ---------- čtení stavu z fotky (Gemini, klíč je jen zde v skriptu) ----------
// Klíč: Nastavení projektu -> Vlastnosti skriptu -> GEMINI_KEY
let AI_DETAIL = '';
// model, který vrátil 404 (už není dostupný), se na 6 hodin přeskakuje
function badModel_(name) {
  try {
    const ca = CacheService.getScriptCache();
    let l = []; try { l = JSON.parse(ca.get('badmodels') || '[]'); } catch (e) {}
    if (l.indexOf(name) < 0) l.push(name);
    ca.put('badmodels', JSON.stringify(l), 21600); ca.remove('models');
  } catch (e) {}
}

// hlavní a záložní modely (další se zkusí, jen když předchozí selže)
const GEMINI_MODELS = ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-3.6-flash', 'gemini-3.7-flash'];
function geminiModels_() { return GEMINI_MODELS.slice(); }

function readWithAi_(meter, b64) {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  if (!meter || !b64) return { ok: false, error: 'data' };
  const models = geminiModels_();
  if (!models.length) return { ok: false, error: 'nomodel', detail: AI_DETAIL || 'žádný model Gemini' };
  const prompt = 'Na fotce je měřidlo: ' + meter.name + ' (druh: ' + meter.typeName + ', zobrazení: ' + meter.type + ', jednotka: ' + meter.unit + (Number(meter.digits) >= 2 ? ', počet číslic před čárkou: ' + meter.digits : '') + '). ' +
    (meter.hint ? meter.hint + ' ' : '') +
    'Přečti aktuální stav měřidla jako jedno číslo s desetinnou tečkou. Odpověz jen JSON {"value": číslo}. Když stav nejde spolehlivě přečíst, dej {"value": null}.';
  const payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0 } });
  const det = [];
  for (const m of models) {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: payload, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) { det.push(m + ' HTTP ' + r.getResponseCode() + ': ' + r.getContentText().replace(/\s+/g, ' ').slice(0, 120)); continue; }
    try {
      const j = JSON.parse(r.getContentText()), t = j.candidates[0].content.parts.map(p => p.text || '').join('');
      const v = JSON.parse(t.replace(/```json|```/g, '')).value;
      if (v !== null && isFinite(Number(v))) return { ok: true, value: Number(v) };
      det.push(m + ': odpověď bez čísla');
    } catch (e) { det.push(m + ': ' + e.message); }
  }
  return { ok: false, error: 'ai', detail: det.join(' | ') };
}

// ---------- příjem odečtu ----------
// tělo (JSON jako text): { token, id, value, note, source, photo: "<base64 jpeg>" }
function doPost(e) {
  MEMO_ON_ = true; METERS_MEMO_ = null;
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return ok_({ ok: false, error: 'json' }); }
  const P = person_(body.token), who = P ? P.name : '';
  if (!who) return ok_({ ok: false, error: 'token' });
  // kdo smí jen prohlížet, nic nezapisuje; ostatní vidí a zapisují jen měřidla svých budov
  const denied = () => ok_({ ok: false, error: 'forbidden' });
  const mine = id => { const mm = meters_().find(m => m.id === id); return !mm || canEdit_(P, mm); };
  if (['read', 'adminCheck', 'placeUpload', 'backfill'].indexOf(body.action) >= 0 && (!P.edit || !mine(body.id))) return denied();
  if (body.action === 'paperUpload' && !P.editAll) return denied(); // arch se týká všech budov
  if (body.action === 'lock') return P.edit && mine(body.id) ? ok_(meterLock_(body, who)) : denied();
  if (body.action === 'editReading') return P.edit ? ok_(editReading_(body, P, who)) : denied();
  if (body.action === 'excludeReading') return P.edit ? ok_(excludeReading_(body, P, who)) : denied();
  if (body.action === 'rereadAi') return P.edit ? ok_(rereadReading_(body, P, who)) : denied();
  if (body.action === 'read') return ok_(readWithAi_(meters_().find(m => m.id === body.id), body.photo));
  if (body.action === 'adminCheck') return ok_(adminCheck_(body.pin));
  if (body.action === 'placeUpload') return ok_(placeUpload_(body));
  if (body.action === 'backfill') return ok_(backfill_(body, who));
  if (body.action === 'paperUpload') return ok_(paperUpload_(body));

  const meter = meters_().find(m => m.id === body.id);
  if (!P.edit || (meter && !canEdit_(P, meter))) return denied();
  if (body.supplier === true && (!meter || !meter.supplier)) return ok_({ ok: false, error: 'nosupplier' });
  const sup = body.supplier === true;
  let noValue = body.noValue === true || body.value === '' || body.value === null || body.value === undefined;
  let value = noValue ? NaN : (meter ? round_(Number(String(body.value).replace(',', '.')), meter.decimals) : Number(String(body.value).replace(',', '.')));
  if (!meter || (!noValue && !isFinite(value)) || (noValue && !body.photo) || (sup && noValue)) return ok_({ ok: false, error: 'data' });
  let source = body.source || '', aiFailed = !!body.aiFailed, aiNote = '';
  // hodnota nebyla zadána a appka chce přečtení na pozadí: přečte ji AI z fotky (mimo zámek, trvá chvíli)
  if (noValue && body.aiRead === true) {
    const ai = readWithAi_(meter, body.photo);
    if (ai.ok) { value = round_(Number(ai.value), meter.decimals); noValue = false; source = 'AI na pozadí'; aiNote = 'AI na pozadí, ověř podle fotky.'; }
    else aiFailed = true;
  }
  // čas pořízení z telefonu (záznam pořízený offline se nezapíše až časem odeslání); nesmí být v budoucnu ani starší než 60 dní
  let when = new Date();
  if (body.takenAt) { const t = new Date(body.takenAt); if (!isNaN(t) && t.getTime() <= Date.now() + 5 * 60000 && t.getTime() > Date.now() - 60 * 86400000) when = t; }
  if (sup && /^\d{4}-\d{2}-\d{2}$/.test(String(body.readingDate || ''))) {
    let d = null; try { d = Utilities.parseDate(body.readingDate + ' 12:00', SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'); } catch (e) {}
    if (!d || d.getFullYear() < 2000) return ok_({ ok: false, error: 'date' });
    if (d.getTime() > Date.now() + 36e5) return ok_({ ok: false, error: 'future' });
    when = d;
  }
  // stejný záznam se při opakovaném odeslání (např. ztracená odpověď) nezapíše dvakrát
  let rid = String(body.clientId || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24);
  if (!rid) rid = Utilities.getUuid().replace(/-/g, '').slice(0, 10);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS);
    const allR = readings_();
    if (body.clientId && allR.some(x => x.rid === rid)) return ok_({ ok: true, dup: true });
    const prev = allR.filter(x => x.id === meter.id && x.has && x.time <= when).sort((a, b) => b.time - a.time)[0];
    const diff = (!noValue && prev) ? Math.round((value - prev.value) * 1000) / 1000 : '';
    let note = (aiNote + ' ' + String(body.note || '')).trim().slice(0, 300);
    if (!noValue && prev && value < prev.value) note = ('POZOR: nižší než minule. ' + note).trim();

    let fileUrl = '', blob = null;
    if (body.photo) {
      const stamp = Utilities.formatDate(when, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd_HH-mm');
      blob = Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', stamp + '_' + meter.id + '.jpg');
      fileUrl = DriveApp.getFolderById(props_().getProperty('FOLDER_ID')).createFile(blob).getUrl();
    }
    // sloupce: Čas, id, Měřidlo, Stav, Předchozí, Přírůstek, Poznámka, Fotka, Zdroj, Kdo, Budova, Chyba, ID záznamu
    const err = noValue ? 'Hodnota nezapsána' + (aiFailed ? ' (OCR ji nepřečetlo)' : '') : '';
    sheet.appendRow([when, meter.id, meter.name, noValue ? '' : value, (!noValue && prev) ? prev.value : '', diff, note, fileUrl,
      noValue ? 'bez hodnoty' : source, who, meter.building, err, rid, '', '', sup ? 'ANO' : '', sup ? new Date() : '']);
    if (sup) { ensureReadingCols_(sheet); sheet.getRange(sheet.getLastRow(), 1, 1, COL_SUP).setBackground(SUP_BG); }
    if (sup && when.getTime() < Date.now() - 36e5) relink_(sheet, meter.id);
    bustCache_(meter.id);

    if (noValue) {
      try { sendMissingMail_(meter, who, rid, blob, aiFailed); } catch (mailErr) { Logger.log('E-mail o chybějící hodnotě se nepodařilo odeslat: ' + mailErr.message); }
      return ok_({ ok: true, pendingValue: true, rid: rid, aiFailed: aiFailed });
    }
    const to = props_().getProperty('NOTIFY');
    if (to && !mailsPaused_()) {
      MailApp.sendEmail(to, envTag_() + (sup ? 'Odečet DODAVATELE: ' : 'Odečet: ') + meter.name + ' = ' + value + ' ' + (meter.unit || ''),
        (meter.buildingLabel ? 'Budova: ' + meter.buildingLabel + '\n' : '') + meter.name + '\nFotil(a): ' + who + '\nStav: ' + value + ' ' + (meter.unit || '') +
        (prev ? '\nPředchozí: ' + prev.value + ' (přírůstek ' + diff + ')' : '') +
        (note ? '\nPoznámka: ' + note : '') + (fileUrl ? '\nFotka: ' + fileUrl : ''));
    }
    return ok_({ ok: true, diff: diff, warn: note.indexOf('POZOR') >= 0, value: value, rid: rid, source: source, supplier: sup, when: when.toISOString() });
  } finally {
    lock.releaseLock();
  }
}

// ---------- chybějící hodnota: e-mail s fotkou a odkazem na formulář ----------
function esc_(x) { return String(x === undefined || x === null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function formSecret_() {
  const p = props_();
  let sec = p.getProperty('FORM_SECRET');
  if (!sec) { sec = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('FORM_SECRET', sec); }
  return sec;
}
function formKey_(rid) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(rid), formSecret_())).slice(0, 22);
}
function formUrl_(rid) {
  return ScriptApp.getService().getUrl() + '?action=form&rid=' + encodeURIComponent(rid) + '&key=' + encodeURIComponent(formKey_(rid));
}

// komu chodí e-maily o chybějící hodnotě: řádky v listu Upozornění s "ano" ve sloupci Chybí hodnota, jinak správce
function missingRecipients_() {
  const t = table_(SHEET_ALERTS), out = [];
  if (t) t.rows.forEach(r => {
    const f = String(r.v['chybí hodnota'] || '').trim().toLowerCase();
    if (['ano', '1', 'x', 'true', 'yes'].indexOf(f) < 0) return;
    String(r.v['komu'] || '').replace(/;/g, ',').split(',').forEach(a => { a = a.trim(); if (a.indexOf('@') > 0 && out.indexOf(a) < 0) out.push(a); });
  });
  if (!out.length) { const n = props_().getProperty('NOTIFY'); if (n) out.push(n); }
  return out;
}

function sendMissingMail_(meter, who, rid, blob, aiFailed) {
  const to = missingRecipients_().join(',');
  if (!to || mailsPaused_()) return;
  const url = formUrl_(rid);
  const title = (meter.buildingLabel ? meter.buildingLabel + ': ' : '') + meter.name + (meter.place ? ' (' + meter.place + ')' : '');
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px"><h3 style="margin:0 0 8px">Chybí hodnota měřidla</h3>' +
    '<p style="margin:0 0 10px"><b>' + esc_(title) + '</b><br>Fotil(a): ' + esc_(who) + (aiFailed ? '<br>Čtení z fotky (OCR) se nepodařilo.' : '') + '</p>' +
    (blob ? '<p><img src="cid:foto" style="max-width:100%;border-radius:8px"></p>' : '') +
    '<p><a href="' + url + '" style="display:inline-block;background:#2563eb;color:#fff;padding:12px 18px;border-radius:10px;text-decoration:none;font-weight:bold">Doplnit hodnotu</a></p>' +
    '<p style="color:#666;font-size:12px">Po klepnutí se otevře formulář s fotkou a polem pro hodnotu. Odeslaná hodnota se zapíše do tabulky.</p></div>';
  const opts = { to: to, subject: envTag_() + 'Fotoapp SVJ: chybí hodnota – ' + title, htmlBody: html, body: 'Chybí hodnota měřidla ' + title + '. Doplnit: ' + url, name: 'Fotoapp SVJ Od Vysoké' };
  if (blob) opts.inlineImages = { foto: blob };
  MailApp.sendEmail(opts);
}

// stránka s formulářem (odkaz z e-mailu): fotka + pole pro hodnotu
function formPage_(rid, key) {
  const page = body => HtmlService.createHtmlOutput(body).setTitle('Doplnění hodnoty').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  const msg = t => page('<div style="font-family:Arial,sans-serif;padding:24px;max-width:520px;margin:auto"><p>' + esc_(t) + '</p></div>');
  rid = String(rid || '');
  if (!rid || String(key || '') !== formKey_(rid)) return msg('Odkaz není platný.');
  const data = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS).getDataRange().getValues();
  const i = data.findIndex((r, k) => k > 0 && String(r[12]) === rid);
  if (i < 0) return msg('Záznam nebyl nalezen.');
  const row = data[i];
  if (row[3] !== '' && row[3] !== null) return msg('Hodnota už je zapsaná: ' + row[3] + '. Děkujeme.');
  let img = '';
  const m = String(row[7]).match(/\/d\/([^/]+)/);
  if (m) { try { const b = DriveApp.getFileById(m[1]).getBlob(); img = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (er) {} }
  let prev = null;
  for (let j = 1; j < data.length; j++) { const r = data[j]; if (j !== i && String(r[1]) === String(row[1]) && r[3] !== '' && new Date(r[0]) < new Date(row[0]) && (!prev || new Date(r[0]) > new Date(prev[0]))) prev = r; }
  const title = (row[10] ? row[10] + ': ' : '') + row[2];
  const js = 'const RID=' + JSON.stringify(rid) + ',KEY=' + JSON.stringify(String(key)) + ';' +
    'function send(){const v=document.getElementById("v").value.trim();if(!v){msg("Zadej hodnotu.",1);return}document.getElementById("b").disabled=true;msg("Ukládám…",0);' +
    'google.script.run.withSuccessHandler(function(r){document.getElementById("f").style.display="none";msg("Hotovo, hodnota "+r.value+" je zapsaná v tabulce. Děkujeme.",0)}).withFailureHandler(function(e){document.getElementById("b").disabled=false;msg(e.message||"Nepodařilo se uložit.",1)}).submitValue(RID,KEY,v,document.getElementById("n").value)}' +
    'function msg(t,err){var m=document.getElementById("m");m.textContent=t;m.style.color=err?"#c2312a":"#1d4fbd"}';
  return page('<div style="font-family:Arial,sans-serif;padding:18px;max-width:520px;margin:auto"><h2 style="margin:0 0 6px">Doplnění hodnoty</h2>' +
    '<p style="margin:0 0 12px"><b>' + esc_(title) + '</b><br>Fotil(a): ' + esc_(row[9]) + (prev ? '<br>Minulý stav: ' + esc_(prev[3]) : '') + '</p>' +
    (img ? '<img src="' + img + '" style="width:100%;border-radius:10px;margin-bottom:12px">' : '<p>Fotka není k dispozici.</p>') +
    '<div id="f"><label style="font-size:13px;color:#555">Stav měřidla</label><input id="v" inputmode="decimal" autocomplete="off" style="width:100%;box-sizing:border-box;font-size:22px;padding:10px;margin:4px 0 10px;border:1px solid #bbb;border-radius:10px">' +
    '<label style="font-size:13px;color:#555">Kdo doplňuje (nepovinné)</label><input id="n" style="width:100%;box-sizing:border-box;font-size:16px;padding:10px;margin:4px 0 12px;border:1px solid #bbb;border-radius:10px">' +
    '<button id="b" onclick="send()" style="width:100%;background:#2563eb;color:#fff;border:0;border-radius:10px;padding:14px;font-size:16px;font-weight:bold">Zapsat do tabulky</button></div>' +
    '<p id="m" style="margin-top:12px"></p><script>' + js + '</' + 'script></div>');
}

// volá se z formuláře: zapíše hodnotu do řádku, ve kterém chyběla
function submitValue(rid, key, value, name) {
  rid = String(rid || '');
  if (!rid || String(key || '') !== formKey_(rid)) throw new Error('Odkaz není platný.');
  let v = Number(String(value).replace(/\s/g, '').replace(',', '.'));
  if (!isFinite(v) || String(value).trim() === '') throw new Error('Zadej prosím číslo.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS), data = sh.getDataRange().getValues();
    const i = data.findIndex((r, k) => k > 0 && String(r[12]) === rid);
    if (i < 0) throw new Error('Záznam nebyl nalezen.');
    const row = data[i];
    if (row[13]) throw new Error('Záznam byl vyřazen.');
    if (row[3] !== '' && row[3] !== null) throw new Error('Hodnota už je zapsaná.');
    const mm = meters_().find(m => m.id === String(row[1]));
    v = round_(v, mm ? mm.decimals : 1);
    let prev = null;
    for (let j = 1; j < data.length; j++) { const r = data[j]; if (j !== i && String(r[1]) === String(row[1]) && r[3] !== '' && new Date(r[0]) < new Date(row[0]) && (!prev || new Date(r[0]) > new Date(prev[0]))) prev = r; }
    const diff = prev ? Math.round((v - Number(prev[3])) * 1000) / 1000 : '';
    let note = String(row[6] || '');
    if (prev && v < Number(prev[3])) note = ('POZOR: nižší než minule. ' + note).trim();
    note = (note + ' Hodnotu doplnil(a) z e-mailu' + (name ? ': ' + String(name).slice(0, 60) : '') + '.').trim();
    const n = i + 1;
    sh.getRange(n, 4).setValue(v);
    sh.getRange(n, 5).setValue(prev ? Number(prev[3]) : '');
    sh.getRange(n, 6).setValue(diff);
    sh.getRange(n, 7).setValue(note);
    sh.getRange(n, 9).setValue('doplněno ručně (e-mail)');
    sh.getRange(n, 12).setValue('');
    bustCache_(String(row[1]));
    return { ok: true, value: v };
  } finally {
    lock.releaseLock();
  }
}

// ---------- správce: PIN a nahrání fotky místa z appky ----------
// PIN je ve vlastnostech skriptu (ADMIN_PIN). Ověřuje se tady, ne v appce. Po 5 špatných pokusech se na 10 minut zamkne.
function adminCheck_(pin) {
  const want = props_().getProperty('ADMIN_PIN');
  if (!want) return { ok: false, error: 'nopin' };
  const cache = CacheService.getScriptCache(), fails = Number(cache.get('pinfail') || 0);
  if (fails >= 5) return { ok: false, error: 'locked' };
  if (String(pin || '') !== String(want)) { cache.put('pinfail', String(fails + 1), 600); return { ok: false, error: 'pin' }; }
  cache.remove('pinfail');
  return { ok: true };
}

// Uloží novou fotku místa do složky Místa měřidel a zapíše její název k měřidlu. Stará fotka na Disku zůstává (nic se nemaže).
function placeUpload_(body) {
  const chk = adminCheck_(body.pin);
  if (!chk.ok) return chk;
  const meter = meters_().find(m => m.id === body.id);
  if (!meter || !body.photo) return { ok: false, error: 'data' };
  const fid = props_().getProperty('PLACES_FOLDER_ID');
  if (!fid) return { ok: false, error: 'nofolder' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const stamp = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyyMMdd-HHmm');
    const name = 'misto-' + (slug_(String(meter.id).replace(/#.*$/, '')) || 'meridlo') + '-' + stamp + '.jpg';
    DriveApp.getFolderById(fid).createFile(Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', name));
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_METERS);
    const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    let col = head.findIndex(x => /^(fotka|plánek)[_ ]?místa/i.test(String(x).trim()));
    if (col < 0) { col = head.findIndex(x => String(x).trim() === ''); if (col < 0) col = head.length; sh.getRange(1, col + 1).setValue('Fotka_místa'); }
    sh.getRange(meter.row, col + 1).setValue(name);
    bustMeters_();
    return { ok: true, name: name };
  } finally {
    lock.releaseLock();
  }
}

// Zpětný záznam bez fotky (stavy z papírových záznamů). Jen správce s PINem.
// Zapíše se k vybranému datu; předchozí stav a přírůstek se dopočítají podle času a upraví se i následující záznam.
// Samotný zápis zpětného záznamu (volající drží zámek). Vrací { ok, diff, warn } nebo { ok: false, error: 'dup' }.
function backfillCore_(meter, when, value, note, who, force, photoUrl, supplier) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS);
  const all = readings_().filter(x => x.id === meter.id && x.has).sort((x, y) => x.time - y.time);
  if (!force && all.some(x => dayNum_(x.time) === dayNum_(when) && x.value === value)) return { ok: false, error: 'dup' };
  const before = all.filter(x => x.time < when), prev = before.length ? before[before.length - 1] : null;
  const next = all.find(x => x.time > when);
  const diff = prev ? Math.round((value - prev.value) * 1000) / 1000 : '';
  let n = String(note || '').slice(0, 200);
  if (prev && value < prev.value) n = ('POZOR: nižší než předchozí záznam. ' + n).trim();
  n = (n + ' Zpětně doplněno správcem.').trim();
  const rid = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sheet.appendRow([when, meter.id, meter.name, value, prev ? prev.value : '', diff, n, photoUrl || '', 'doplněno zpětně (papír)', who, meter.building, '', rid, '', '', supplier && meter.supplier ? 'ANO' : '']);
  if (supplier && meter.supplier) { ensureReadingCols_(sheet); sheet.getRange(sheet.getLastRow(), 1, 1, COL_SUP).setBackground(SUP_BG); }
  // následující záznam má nový předchozí stav
  if (next && next.row) { sheet.getRange(next.row, 5).setValue(value); sheet.getRange(next.row, 6).setValue(Math.round((next.value - value) * 1000) / 1000); }
  bustCache_(meter.id);
  return { ok: true, diff: diff, warn: !!(prev && value < prev.value) };
}

function backfill_(body, who) {
  const chk = adminCheck_(body.pin);
  if (!chk.ok) return chk;
  const meter = meters_().find(m => m.id === body.id);
  const raw = String(body.value === undefined || body.value === null ? '' : body.value).replace(',', '.').trim();
  const value = meter && raw !== '' ? round_(Number(raw), meter.decimals) : NaN;
  if (!meter || !isFinite(value)) return { ok: false, error: 'data' };
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  let when;
  try { when = Utilities.parseDate(String(body.when), tz, "yyyy-MM-dd'T'HH:mm"); } catch (e) { return { ok: false, error: 'date' }; }
  if (!(when.getFullYear() >= 2000)) return { ok: false, error: 'date' };
  if (when.getTime() > Date.now() + 3600000) return { ok: false, error: 'future' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return backfillCore_(meter, when, value, body.note, who, !!body.force, '', body.supplier === true);
  } finally {
    lock.releaseLock();
  }
}

// ---------- fotka / plánek místa měřidla ----------
// V tabulce je název souboru ze složky Místa měřidel (nebo odkaz / id souboru na Disku).
function resolvePlaceId_(ref) {
  ref = String(ref || '').trim();
  if (!ref) return '';
  let m = ref.match(/\/d\/([A-Za-z0-9_-]{20,})/) || ref.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
  if (m) return m[1];
  if (/^[A-Za-z0-9_-]{25,}$/.test(ref)) return ref;
  const fid = props_().getProperty('PLACES_FOLDER_ID');
  if (!fid) return '';
  const it = DriveApp.getFolderById(fid).getFilesByName(ref);
  return it.hasNext() ? it.next().getId() : '';
}

// ---------- ceny za jednotku (list Ceny) ----------
// Typ | Cena za jednotku (Kč, s DPH) | Platná od (prázdné = vždy) | Poznámka. Pro jeden typ může být víc řádků s různým datem.
function ensurePrices_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PRICES);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_PRICES);
  sh.appendRow(['Typ', 'Cena za jednotku (Kč, s DPH)', 'Platná od', 'Poznámka']);
  sh.appendRow(['Elektroměr', 8.17, '', 'Průměr z 10 faktur PRE (19. 7. 2025 – 23. 7. 2026): 111 586 Kč s DPH / 13 656 kWh, včetně pevných plateb']);
  sh.setFrozenRows(1); sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  return sh;
}
function pricesFor_(typeName) {
  try {
    const sh = ensurePrices_(), v = sh.getDataRange().getValues();
    const want = String(typeName || '').normalize('NFC').trim().toLowerCase(), out = [];
    v.slice(1).forEach(r => {
      if (String(r[0]).normalize('NFC').trim().toLowerCase() !== want) return;
      const p = Number(String(r[1]).replace(/\s/g, '').replace(',', '.'));
      if (!isFinite(p) || r[1] === '') return;
      let from = 0;
      if (r[2] !== '' && r[2] !== null) { const d = Object.prototype.toString.call(r[2]) === '[object Date]' ? r[2] : new Date(r[2]); if (!isNaN(d.getTime())) from = d.getTime(); }
      out.push({ from: from, price: p });
    });
    return out.sort((a, b) => a.from - b.from);
  } catch (e) { return []; }
}

// id souboru z odkazu na Disk (…/d/ID/… nebo …?id=ID)
function driveId_(url) {
  const m = String(url || '').match(/\/d\/([\w-]{20,})/) || String(url || '').match(/[?&]id=([\w-]{20,})/);
  return m ? m[1] : '';
}
// fotka k záznamu (arch nebo fotka měřidla) se čte přes skript, takže ji uvidí každý, kdo má aplikaci, bez oprávnění k Disku.
// Povolené jsou jen soubory, které jsou uvedené v listu Odečty.
function recordImage_(fid, P) {
  try {
    fid = String(fid || '');
    if (!/^[\w-]{20,}$/.test(fid)) return { ok: false, error: 'nofile' };
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS), n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false, error: 'nofile' };
    const col = sh.getRange(2, 1, n, 8).getValues();
    const row = col.find(r => String(r[7]).indexOf(fid) >= 0);
    if (!row) return { ok: false, error: 'nofile' };
    const mm = meters_().find(m => m.id === String(row[1]));
    if (P && mm && !canSee_(P, mm)) return { ok: false, error: 'nofile' };
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + fid + '&sz=w1600', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    let blob = null;
    if (r.getResponseCode() === 200) blob = r.getBlob();
    else { const b = DriveApp.getFileById(fid).getBlob(); if (b.getBytes().length <= 2500000) blob = b; }
    if (!blob) return { ok: false, error: 'big' };
    return { ok: true, mime: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
  } catch (e) { return { ok: false, error: 'fail' }; }
}

function placeImage_(ref) {
  try {
    const id = resolvePlaceId_(ref);
    if (!id) return { ok: false, error: 'nofile' };
    // zmenšený náhled z Disku (rychlý přenos); funguje i pro PDF plánky
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1100', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    let blob = null;
    if (r.getResponseCode() === 200) blob = r.getBlob();
    else { const b = DriveApp.getFileById(id).getBlob(); if (b.getBytes().length <= 1200000) blob = b; }
    if (!blob) return { ok: false, error: 'big' };
    return { ok: true, mime: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { ok: false, error: 'fail' };
  }
}

// ---------- uspořádání Disku: SVJ_app / OSTRE nebo TESTOVACI ----------
// Ostrá i testovací verze mají každá svou podsložku; testovací kopie skriptu má ve vlastnostech ENV = test.
function child_(parent, name) {
  const want = String(name).toLowerCase();
  const all = parent.getFolders();
  while (all.hasNext()) { const f = all.next(); if (f.getName().toLowerCase() === want) return f; }
  return parent.createFolder(name);
}
function driveRoot_() {
  const env = props_().getProperty('ENV') === 'test' ? 'TESTOVACI' : 'OSTRE';
  return child_(child_(DriveApp.getRootFolder(), DRIVE_ROOT), env);
}

// Přesune do složky tabulku a složky s fotkami (odkazy, id a nastavení zůstávají v platnosti).
function organizeDrive() {
  const dest = driveRoot_(), p = props_(), done = [];
  const mv = (file, label) => { try { if (file.getParents().hasNext() && file.getParents().next().getId() === dest.getId()) { done.push(label + ': už je na místě'); return; } file.moveTo(dest); done.push(label + ': přesunuto'); } catch (e) { done.push(label + ': NEPODAŘILO SE (' + e.message + ')'); } };
  mv(DriveApp.getFileById(SpreadsheetApp.getActive().getId()), 'Tabulka');
  const f1 = p.getProperty('FOLDER_ID');
  if (f1) { try { const fo = DriveApp.getFolderById(f1); if (fo.getName() === 'Odečty měřidel – fotky') fo.setName(FOLDER_NAME); mv(fo, 'Fotodokumentace odečtů'); } catch (e) { done.push('Fotodokumentace: NEPODAŘILO SE (' + e.message + ')'); } }
  const f2 = p.getProperty('PLACES_FOLDER_ID');
  if (f2) { try { mv(DriveApp.getFolderById(f2), 'Místa měřidel'); } catch (e) { done.push('Místa měřidel: NEPODAŘILO SE (' + e.message + ')'); } }
  const text = 'Složka: ' + DRIVE_ROOT + ' / ' + dest.getName() + String.fromCharCode(10) + dest.getUrl() + String.fromCharCode(10) + String.fromCharCode(10) + done.join(String.fromCharCode(10));
  Logger.log(text);
  try { SpreadsheetApp.getActive().toast(text.replace(/\s+/g, ' ').slice(0, 280), 'Hotovo (podrobnosti v Protokolu spuštění)', 25); } catch (e) {}
}

// ---------- testovací sada ----------
// Spouští se v OSTRÉM skriptu. V SVJ_app/TESTOVACI vytvoří kopii tabulky a obě složky pro fotky.
function createTestSet() {
  if (props_().getProperty('ENV') === 'test') throw new Error('Testovací sadu zakládej z ostré tabulky, ne z testovací.');
  const nl = String.fromCharCode(10);
  const testDir = child_(child_(DriveApp.getRootFolder(), DRIVE_ROOT), 'TESTOVACI');
  const copyName = 'mereni_SVJ_TEST';
  let copy, made = false;
  const ex = testDir.getFilesByName(copyName);
  if (ex.hasNext()) copy = ex.next();
  else { copy = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).makeCopy(copyName, testDir); made = true; }
  child_(testDir, FOLDER_NAME);
  child_(testDir, PLACES_NAME);
  const text = 'Testovací sada je v: ' + DRIVE_ROOT + ' / TESTOVACI' + nl + testDir.getUrl() + nl + nl +
    (made ? 'Vytvořena kopie tabulky: ' : 'Kopie tabulky už existovala: ') + copyName + nl + copy.getUrl() + nl + nl +
    'Dál (v kopii): Rozšíření → Apps Script → Vlastnosti skriptu: ENV = test (a GEMINI_KEY, NOTIFY), potom spustit setup, clearTestData a resetToken.';
  Logger.log(text);
  try { SpreadsheetApp.getActive().toast(text.replace(/\s+/g, ' ').slice(0, 280), 'Hotovo (podrobnosti v Protokolu spuštění)', 25); } catch (e) {}
}

// Spouští se v TESTOVACÍM skriptu (ENV = test): smaže zkopírované odečty a tajné kódy z ostré tabulky.
// V ostré verzi odmítne běžet (ochrana ostrých dat).
function clearTestData() {
  if (props_().getProperty('ENV') !== 'test') throw new Error('Nejdřív nastav vlastnost skriptu ENV = test. V ostré verzi se tato funkce nespouští.');
  const ss = SpreadsheetApp.getActive(), out = [];
  const r = ss.getSheetByName(SHEET_READINGS);
  if (r && r.getLastRow() > 1) { r.getRange(2, 1, r.getLastRow() - 1, Math.max(1, r.getLastColumn())).clearContent(); out.push('Odečty: smazány zkopírované řádky'); }
  const p = ss.getSheetByName(SHEET_PEOPLE);
  if (p && p.getLastRow() > 1) { p.getRange(2, 2, p.getLastRow() - 1, 2).clearContent(); out.push('Lidé: smazány zkopírované tajné kódy a odkazy (vygeneruješ nové)'); }
  props_().deleteProperty('NOTIF'); bustAll_();
  Logger.log(out.join(String.fromCharCode(10)) + String.fromCharCode(10) + 'Teď spusť resetToken a vygeneruj odkazy (menu Měřidla).');
  try { SpreadsheetApp.getActive().toast(out.join(' · ') + ' · Teď spusť resetToken.', 'Hotovo', 25); } catch (e) {}
}

// kontrola listu Měřidla: duplicitní id, neznámý typ, chybějící budova
function checkMeters() {
  const tt = table_(SHEET_TYPES), mt = table_(SHEET_METERS), types = {}, rows = {}, problems = [];
  if (tt) tt.rows.forEach(r => { const k = String(r.v['typ'] || '').trim().toLowerCase(); if (k) types[k] = 1; });
  mt.rows.forEach(r => {
    const id = String(r.v['id'] || '').trim(), type = String(r.v['typ'] || '').trim(), name = String(r.v['název'] || '').trim();
    if (!id) { if (name || type) problems.push('řádek ' + r.row + ': chybí id (doplní se samo při dalším otevření appky)'); return; }
    (rows[id] = rows[id] || []).push(r.row);
    if (!type || !types[type.toLowerCase()]) problems.push('řádek ' + r.row + ': neznámý typ "' + type + '" (musí být v listu Typy)');
    if (!String(r.v['budova'] || '').trim()) problems.push('řádek ' + r.row + ': chybí budova');
  });
  Object.keys(rows).forEach(id => { if (rows[id].length > 1) problems.push('id "' + id + '" je víckrát (řádky ' + rows[id].join(', ') + '). Musí být jedinečné, jinak se slije historie.'); });
  const total = mt.rows.filter(r => String(r.v['id'] || '').trim()).length;
  const keys = Object.keys((mt.rows[0] || { v: {} }).v);
  ['id', 'název', 'typ', 'budova'].forEach(k => { if (keys.indexOf(k) < 0) problems.push('v hlavičce chybí sloupec "' + k + '"'); });
  const sk = keys.find(k => /^(dodavatel|plomb)/.test(k)), isYes = t => /^(ano|a|1|x|true|yes)$/i.test(String(t === undefined || t === null ? '' : t).trim());
  let info = '';
  if (!sk) problems.push('chybí sloupec Dodavatel (založí se při příštím otevření tabulky)');
  else {
    const ids = mt.rows.filter(r => String(r.v['id'] || '').trim()), yes = ids.filter(r => isYes(r.v[sk]));
    const odd = ids.filter(r => String(r.v[sk] === undefined || r.v[sk] === null ? '' : r.v[sk]).trim() !== '' && !isYes(r.v[sk]));
    odd.forEach(r => problems.push('řádek ' + r.row + ': ve sloupci Dodavatel je "' + r.v[sk] + '" (platí jen "ano", jinak nech prázdné)'));
    info = '\n\nPod plombou dodavatele: ' + yes.length + ' z ' + ids.length + ' měřidel, interních: ' + (ids.length - yes.length - odd.length) + '.';
  }
  SpreadsheetApp.getUi().alert((problems.length ? 'Nalezeno ' + problems.length + ' problémů:\n\n' + problems.join('\n') : 'Všechna měřidla jsou v pořádku (' + total + ').') + info);
}

// ---------- upozornění e-mailem, když chybí rozhodující odečet (poslední den v měsíci) ----------
// Pravidla jsou v listu Upozornění: Komu (e-mail, víc adres odděl čárkou), Budova (prázdné = všechny), Dní po termínu (0 = v den termínu).
// Každé pravidlo upozorní na dané měřidlo jen jednou za jeden termín (po novém odečtu se začíná znovu).
// ---------- pozastavení e-mailů a reset záznamů (volá se z menu) ----------
function mailsPaused_() { return props_().getProperty('MAIL_PAUSED') === '1'; }
function toggleMails() {
  const ui = SpreadsheetApp.getUi(), p = props_();
  if (mailsPaused_()) { p.deleteProperty('MAIL_PAUSED'); ui.alert('E-maily jsou znovu zapnuté.'); }
  else { p.setProperty('MAIL_PAUSED', '1'); ui.alert('Odesílání e-mailů je pozastavené (připomínky, odečty i chybějící hodnoty). Stejnou položkou menu ho znovu zapneš.'); }
}
function resetReadings() {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS), n = sh ? sh.getLastRow() - 1 : 0;
  if (n < 1) { ui.alert('List Odečty je prázdný.'); return; }
  if (ui.alert('Resetovat záznamy', 'Opravdu vymazat všech ' + n + ' řádků z listu Odečty (hlavička zůstane)? Záloha se nedělá. Fotky na Disku zůstanou. Vrátit to půjde jen přes Soubor → Historie verzí.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { sh.getRange(2, 1, n, sh.getLastColumn()).clearContent(); props_().deleteProperty('NOTIF'); bustAll_(); } finally { lock.releaseLock(); }
  ui.alert('Hotovo. Vymazáno ' + n + ' řádků.');
}

function resetImportSheet() {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActive().getSheetByName('Import z papíru'), n = sh ? sh.getLastRow() - 1 : 0;
  if (n < 1) { ui.alert('List Import z papíru je prázdný.'); return; }
  if (ui.alert('Vymazat Import z papíru', 'Opravdu vymazat všech ' + n + ' řádků z listu Import z papíru (hlavička zůstane)? Fotky archů na Disku zůstanou. Vrátit to půjde jen přes Soubor → Historie verzí.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.getRange(2, 1, n, sh.getLastColumn()).clearContent();
  ui.alert('Hotovo. Vymazáno ' + n + ' řádků.');
}

function installReminder() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'reminder').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('reminder').timeBased().everyDays(1).atHour(8).create();
}

function rules_() {
  const t = table_(SHEET_ALERTS);
  if (!t) return [];
  return t.rows.map(r => ({
    to: String(r.v['komu'] || '').replace(/;/g, ',').split(',').map(x => x.trim()).filter(x => x.indexOf('@') > 0).join(','),
    building: String(r.v['budova'] || '').trim(),
    after: Math.max(0, Math.round(Number(r.v['dní po termínu']) || 0))
  })).filter(r => r.to);
}

function reminder() {
  if (mailsPaused_()) return;
  const rules = rules_();
  if (!rules.length) return;
  const props = props_(), all = readings_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  let sent = {};
  try { sent = JSON.parse(props.getProperty('NOTIF') || '{}'); } catch (e) {}
  const next = {}, out = {};
  meters_().forEach(m => {
    const last = all.filter(x => x.id === m.id && x.has).sort((x, y) => y.time - x.time)[0];
    const di = dueInfo_(m, last);
    if (!di.due) return;
    rules.forEach((r, ri) => {
      if (r.building && r.building.toLowerCase() !== String(m.building).toLowerCase() && r.building.toLowerCase() !== String(m.parent).toLowerCase()) return;
      if (last ? di.overdueDays < r.after : r.after > 0) return;
      const key = ri + '|' + m.id + '|' + di.key; // jedno upozornění za měřidlo a termín (u intervalu za poslední odečet)
      if (sent[key]) { next[key] = sent[key]; return; }
      (out[ri] = out[ri] || []).push({ m: m, key: key, last: last, di: di });
    });
  });
  const base = props.getProperty('APP_URL') || DEFAULT_APP_URL;
  Object.keys(out).forEach(ri => {
    const items = out[ri].sort((x, y) => String(x.m.building).localeCompare(String(y.m.building), 'cs') || String(x.m.name).localeCompare(String(y.m.name), 'cs'));
    const lines = items.map(it => '- ' + (it.m.buildingLabel ? it.m.buildingLabel + ': ' : '') + it.m.name + (it.m.place ? ' (' + it.m.place + ')' : '') + ' – ' +
      (it.last ? (it.di.overdueDays ? 'po termínu ' + it.di.overdueDays + ' d' : 'termín je dnes') + ', naposledy ' + Utilities.formatDate(it.last.time, tz, 'd. M. yyyy') + (it.last.who ? ' (' + it.last.who + ')' : '') : 'zatím bez odečtu'));
    try {
      MailApp.sendEmail({
        to: rules[ri].to,
        subject: envTag_() + 'Fotoapp SVJ: chybí odečty měřidel (' + items.length + ')',
        body: 'Dobrý den,' + '\n\n' + 'u těchto měřidel chybí odečet za rozhodující termín (poslední den v měsíci, nebo interval měřidla):' + '\n\n' + lines.join('\n') + '\n\n' + 'Otevřít aplikaci: ' + base + '\n\n' + 'Tato zpráva se pro každé měřidlo posílá jen jednou za jeden termín.',
        name: 'Fotoapp SVJ Od Vysoké'
      });
      items.forEach(it => { next[it.key] = new Date().toISOString(); });
    } catch (e) {
      Logger.log('E-mail se nepodařilo odeslat (' + rules[ri].to + '): ' + e.message);
    }
  });
  props.setProperty('NOTIF', JSON.stringify(next));
}
