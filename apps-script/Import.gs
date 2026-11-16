// Jednorázový import stavů z papírového záznamu k 30. 9. 2026 (č. p. 271–275).
// Zapisuje JEN hodnoty k měřidlům, která už jsou v listu Měřidla. Nová měřidla nezakládá, chybějící vypíše.
// Vodoměry se přiřazují podle budovy (v každé je jeden), elektroměry podle čísla. Opakované spuštění nic nezdvojí.
// Po použití se tento soubor ze skriptu odstraní.
const PAPER_WHEN = '2026-09-30 12:00';
const PAPER_NOTE = 'Stav z papírového záznamu k 30. 9. 2026.';
const PAPER_ITEMS = [
  // [budova, druh, číslo elektroměru ('' u vodoměru), stav]
  ['275', 'Vodoměr', '', 219],        ['275', 'Elektroměr', 'W060073', 1900.6], ['275', 'Elektroměr', 'T698532', 535.7],
  ['274', 'Vodoměr', '', 1230],       ['274', 'Elektroměr', 'W033731', 6563.2], ['274', 'Elektroměr', 'W035443', 139.6],
  ['273', 'Vodoměr', '', 287],        ['273', 'Elektroměr', 'W035442', 6480.8], ['273', 'Elektroměr', 'T698239', 1044.2],
  ['272', 'Vodoměr', '', 273],        ['272', 'Elektroměr', 'W035424', 2795.7], ['272', 'Elektroměr', 'W032185', 192.6],
  ['271', 'Vodoměr', '', 1524],       ['271', 'Elektroměr', 'T751277', 2761.9], ['271', 'Elektroměr', 'T583695', 80540.0]
];

const paperKey_ = x => String(x).toLowerCase().replace(/[^a-z0-9]/g, '');

// najde měřidlo v tabulce: elektroměr podle čísla, vodoměr podle budovy
function paperMeter_(it, meters) {
  if (it[2]) return { m: meters.find(m => paperKey_(m.id) === paperKey_(it[2])) || null };
  const c = meters.filter(m => String(m.building).toLowerCase() === String(it[0]).toLowerCase() && String(m.typeName).toLowerCase() === String(it[1]).toLowerCase());
  return c.length === 1 ? { m: c[0] } : { m: null, why: c.length ? 'v budově je víc měřidel tohoto druhu' : '' };
}

function paperRun_(write) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const when = Utilities.parseDate(PAPER_WHEN, tz, 'yyyy-MM-dd HH:mm');
  const meters = meters_(), out = [], missing = [];
  let wrote = 0, dup = 0;
  const lock = LockService.getScriptLock();
  if (write) lock.waitLock(30000);
  try {
    PAPER_ITEMS.forEach(it => {
      const f = paperMeter_(it, meters), label = it[0] + ' ' + it[1] + (it[2] ? ' ' + it[2] : '');
      if (!f.m) { missing.push(label + (f.why ? ' (' + f.why + ')' : '')); out.push('CHYBÍ   ' + label + ' = ' + it[3] + (f.why ? '  [' + f.why + ']' : '')); return; }
      const value = round_(it[3], f.m.decimals);
      if (!write) { out.push('ZAPÍŠE  ' + label + ' → id ' + f.m.id + ' (' + f.m.name + ') = ' + value); return; }
      const r = backfillCore_(f.m, when, value, PAPER_NOTE, 'Správce', false);
      if (r.ok) { wrote++; out.push('ZAPSÁNO ' + label + ' → id ' + f.m.id + ' = ' + value + (r.warn ? ' (nižší než předchozí!)' : '')); }
      else if (r.error === 'dup') { dup++; out.push('PŘESKOČENO (už zapsáno) ' + label + ' → id ' + f.m.id); }
      else { missing.push(label + ' (' + r.error + ')'); out.push('CHYBA   ' + label); }
    });
  } finally {
    if (write) lock.releaseLock();
  }
  const nl = String.fromCharCode(10);
  const text = (write ? 'IMPORT PROVEDEN' : 'NÁHLED (nic se nezapisuje)') + nl + out.join(nl) + nl + nl +
    (write ? 'Zapsáno: ' + wrote + ', už zapsáno dříve: ' + dup + ', chybí nebo chyba: ' + missing.length : 'Bude zapsáno: ' + out.filter(x => x.indexOf('ZAPÍŠE') === 0).length + ', chybí v tabulce: ' + missing.length);
  Logger.log(text);
  try { SpreadsheetApp.getActive().toast(text.split(nl).slice(-1)[0], write ? 'Import hotov' : 'Náhled hotov', 20); } catch (e) {}
}

function importPaper930Preview() { paperRun_(false); }
function importPaper930() { paperRun_(true); }
