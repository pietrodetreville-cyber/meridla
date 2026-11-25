// ---------- tabulka měřidel: samostatný soubor „Měřidla SVJ Od Vysoké“ (jen listy Měřidla a Odečty) ----------
// Skript je připojený k tabulce Administrace (dřív SVJ_app): Lidé, Typy, Budovy, Ceny, Upozornění, Import z papíru, archiv odečtů a menu.
// Doklady, Faktury a Přehled pro výbor jsou v souboru Doklady (Receipts.gs, Overview.gs).
// Přesun se dělá jednou z menu Administrace → Přesunout měřidla do samostatné tabulky. Do té doby se vše čte z Administrace jako dřív.
// ID souboru měřidel je ve vlastnosti skriptu METERS_SS. Původní listy zůstanou v Administraci skryté jako záloha (nic se nemaže).

function envSuffix_() { return props_().getProperty('ENV') === 'test' ? ' – TEST' : ''; }
function metersFileName_() { return 'Měřidla SVJ Od Vysoké' + envSuffix_(); }
function adminFileName_() { return 'Administrace SVJ Od Vysoké' + envSuffix_(); }

let METERS_SS_MEMO_ = null;
// soubor s listy Měřidla a Odečty (před přesunem je to Administrace)
function metersSS_() {
  if (METERS_SS_MEMO_) return METERS_SS_MEMO_;
  const id = props_().getProperty('METERS_SS');
  if (!id) return SpreadsheetApp.getActive();
  return (METERS_SS_MEMO_ = SpreadsheetApp.openById(id)); // když soubor nejde otevřít, ať to skončí chybou, ne zápisem jinam
}
// list podle názvu: Měřidla a Odečty ze souboru měřidel, ostatní z Administrace
function sheet_(name) {
  return (name === SHEET_METERS || name === SHEET_READINGS ? metersSS_() : SpreadsheetApp.getActive()).getSheetByName(name);
}

// menu v souboru měřidel (instalovaný spouštěč při otevření)
function metersOnOpen() {
  ensureSupplierCol_(); formatMeterRows_();
  SpreadsheetApp.getUi().createMenu('Měřidla')
    .addItem('Doplnit chybějící id měřidel', 'fillIds_')
    .addItem('Obarvit měřidla (dodavatel zeleně, interní žlutě)', 'formatMeterRows')
    .addItem('Obnovit data v aplikaci (vymazat mezipaměť)', 'clearAppCache')
    .addToUi();
}
function installMetersMenu_(ss) {
  try {
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'metersOnOpen' && t.getTriggerSourceId() === ss.getId()))
      ScriptApp.newTrigger('metersOnOpen').forSpreadsheet(ss).onOpen().create();
  } catch (e) { Logger.log('Menu pro soubor měřidel se nepodařilo zapnout: ' + e.message); }
}

function openMetersFile() {
  const ss = metersSS_(), sep = !!props_().getProperty('METERS_SS');
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(sep
    ? '<p style="font:14px Arial">Měřidla a odečty jsou v samostatné tabulce:</p><p style="font:14px Arial"><a href="' + ss.getUrl() + '" target="_blank">' + escH_(ss.getName()) + '</a></p>'
    : '<p style="font:14px Arial">Měřidla jsou zatím v této tabulce. Do samostatné je přesuneš z menu Administrace → Přesunout měřidla do samostatné tabulky.</p>').setWidth(440).setHeight(120), 'Tabulka měřidel');
}

// jednorázový přesun listů Měřidla a Odečty do nového souboru
function splitMetersFile() {
  const ui = SpreadsheetApp.getUi(), admin = SpreadsheetApp.getActive(), tz = admin.getSpreadsheetTimeZone();
  if (props_().getProperty('METERS_SS')) { openMetersFile(); return; }
  const names = [SHEET_METERS, SHEET_READINGS], src = names.map(n => admin.getSheetByName(n));
  if (src.some(s => !s)) { ui.alert('V tabulce chybí list Měřidla nebo Odečty, nic nepřesouvám.'); return; }
  // vzorce s odkazem na jiný list by se v novém souboru rozbily: v tom případě nic nepřesouvat
  const bad = [];
  src.forEach(s => s.getDataRange().getFormulas().forEach((r, i) => r.forEach((f, j) => { if (f && f.indexOf('!') >= 0) bad.push(s.getName() + ' ' + s.getRange(i + 1, j + 1).getA1Notation()); })));
  if (bad.length) { ui.alert('Nepřesouvám', 'Tyto buňky mají vzorec s odkazem na jiný list a v novém souboru by se rozbil: ' + bad.slice(0, 15).join(', ') + (bad.length > 15 ? ' …' : '') + '. Napiš mi, co s nimi.', ui.ButtonSet.OK); return; }
  if (ui.alert('Přesunout měřidla do samostatné tabulky',
    'Listy Měřidla a Odečty se zkopírují do nového souboru „' + metersFileName_() + '“ ve stejné složce. Aplikace bude dál fungovat stejně.\n\n' +
    'Tato tabulka se přejmenuje na „' + adminFileName_() + '“ a zůstanou v ní Lidé, Typy, Budovy, Ceny, Upozornění, Import a archiv.\n\n' +
    'Nic se nemaže: původní listy tu zůstanou skryté s názvem „… – záloha“. Pokračovat?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const lock = LockService.getScriptLock(); lock.waitLock(30000); // aplikace během přesunu počká (zápisy odečtů drží stejný zámek)
  let ss;
  try {
    ss = SpreadsheetApp.create(metersFileName_());
    try { const p = DriveApp.getFileById(admin.getId()).getParents(); if (p.hasNext()) DriveApp.getFileById(ss.getId()).moveTo(p.next()); } catch (e) {}
    const blank = ss.getSheets()[0];
    const copies = src.map(s => s.copyTo(ss).setName(s.getName()));
    ss.deleteSheet(blank); // prázdný úvodní list nového souboru
    copies.forEach(c => c.setFrozenRows(1));
    // kontrola: kopie má stejná data jako originál
    src.forEach((s, i) => {
      const a = JSON.stringify(s.getDataRange().getValues()), b = JSON.stringify(copies[i].getDataRange().getValues());
      if (a !== b) throw new Error('Kopie listu ' + s.getName() + ' nesouhlasí s originálem, přesun zastaven (nic se nezměnilo).');
    });
    // výběr typu v listu Měřidla: dřív odkaz na list Typy (je v Administraci), teď seznam názvů
    try {
      const m = copies[0], head = m.getRange(1, 1, 1, m.getLastColumn()).getValues()[0].map(h => String(h).normalize('NFC').trim().toLowerCase());
      const ci = head.indexOf('typ'), tt = table_(SHEET_TYPES);
      const types = tt ? tt.rows.map(r => String(r.v['typ'] || r.v['název'] || r.v['nazev'] || '').trim()).filter(Boolean) : [];
      if (ci >= 0 && types.length) m.getRange(2, ci + 1, Math.max(m.getMaxRows() - 1, 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(types, true).setAllowInvalid(true).build());
    } catch (e) { Logger.log('Výběr typu: ' + e.message); }
    props_().setProperty('METERS_SS', ss.getId()); METERS_SS_MEMO_ = null;
    // originály: přejmenovat a skrýt jako zálohu
    const stamp = Utilities.formatDate(new Date(), tz, 'd. M. yyyy');
    src.forEach(s => { s.setName(s.getName() + ' – záloha ' + stamp); s.hideSheet(); });
    // staré kopie dokladů z doby před samostatným souborem dokladů jen skrýt
    admin.getSheets().forEach(s => { if (/\(přesunuto\)$/.test(s.getName())) { try { s.hideSheet(); } catch (e) {} } });
    try { admin.rename(adminFileName_()); } catch (e) {}
    bustAll_();
  } finally { lock.releaseLock(); }
  installMetersMenu_(ss);
  ensureSupplierCol_(); formatMeterRows_();
  ui.alert('Hotovo', 'Měřidla a odečty jsou v souboru „' + ss.getName() + '“:\n' + ss.getUrl() + '\n\nAdministrace: ' + admin.getName() + '\nDoklady a přehled pro výbor: soubor „' + docsFileName_() + '“.', ui.ButtonSet.OK);
}
