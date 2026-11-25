// ---------- přehled dokladů pro výbor: samostatná tabulka ve složce „dokumenty“ ----------
// Kolegové z výboru nepotřebují přístup do SVJ_app: dostanou právo číst tuto tabulku a složku Doklady (kvůli odkazům na fotky / PDF).
// Přehled se obnoví po uložení dokladu, registraci faktury a odeslání účetní; ručně menu Doklady → Obnovit přehled dokladů.

const DASH_FOLDER = 'dokumenty';
function dashFileName_() { return 'Přehled dokladů SVJ Od Vysoké' + (props_().getProperty('ENV') === 'test' ? ' – TEST' : ''); }
function dashFolder_() {
  const it = DriveApp.getRootFolder().getFoldersByName(DASH_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(DASH_FOLDER);
}
function dashSS_() {
  const id = props_().getProperty('DASH_SS');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(dashFileName_());
  try { DriveApp.getFileById(ss.getId()).moveTo(dashFolder_()); } catch (e) {}
  props_().setProperty('DASH_SS', ss.getId());
  ss.getSheets()[0].setName('Přehled');
  ss.insertSheet('Doklady');
  shareDash_(ss);
  return ss;
}

// obnoví celý přehled z listu Doklady
function dashRefresh_() {
  const dss = docsSSIfExists_(), src = dss && dss.getSheetByName(SHEET_DOCS);
  if (!src) return;
  const ss = dashSS_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const rows = docRows_(src, Math.max(DOCS_HEAD.length, src.getLastColumn())).filter(r => r[9]);
  const isD = x => Object.prototype.toString.call(x) === '[object Date]';
  const sent = r => /^odesl/i.test(String(r[8] || ''));
  const num = r => String(r[COL_DNUM - 1] || '');
  rows.sort((a, b) => num(b).localeCompare(num(a)) || new Date(b[0]) - new Date(a[0]));

  // ---- list Doklady: všechny doklady s odkazem
  const head = ['Číslo dokladu', 'Typ', 'Datum dokladu', 'DUZP', 'Částka (Kč)', 'Dodavatel', 'Účel', 'Budova', 'Číslo faktury', 'Splatnost', 'Zapsal(a)', 'Stav', 'Odesláno účetní', 'Dávka', 'Zpráva pro účetní', 'Doklad'];
  let sh = ss.getSheetByName('Doklady') || ss.insertSheet('Doklady');
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  const data = rows.map(r => [num(r), String(r[COL_DTYPE - 1] || 'paragon'), isD(r[1]) ? r[1] : '', isD(r[COL_DDUZP - 1]) ? r[COL_DDUZP - 1] : '', Number(r[2]) || 0, String(r[3] || ''), String(r[4] || ''), String(r[5] || ''), String(r[COL_DINVNO - 1] || ''), isD(r[COL_DDUE - 1]) ? r[COL_DDUE - 1] : '', String(r[6] || ''), sent(r) ? 'odesláno' : 'nový', isD(r[11]) ? r[11] : '', String(r[COL_DBATCH - 1] || ''), String(r[COL_DMSG - 1] || ''), '']);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#eef2f7');
  sh.setFrozenRows(1);
  if (data.length) {
    sh.getRange(2, 1, data.length, head.length).setValues(data);
    sh.getRange(2, 3, data.length, 2).setNumberFormat('d. M. yyyy'); sh.getRange(2, 10, data.length, 1).setNumberFormat('d. M. yyyy'); sh.getRange(2, 13, data.length, 1).setNumberFormat('d. M. yyyy');
    sh.getRange(2, 5, data.length, 1).setNumberFormat('#,##0.00 "Kč"');
    rows.forEach((r, i) => { const u = String(r[7] || ''); if (/^https?:/.test(u)) setPhotoLink_(sh, i + 2, 16, (String(r[COL_DTYPE - 1]) === 'faktura' ? '📄 ' : '📷 ') + (num(r) || 'otevřít'), u); });
    sh.getRange(1, 1, data.length + 1, head.length).createFilter();
  }
  [110, 70, 95, 95, 110, 200, 220, 70, 110, 95, 140, 75, 105, 120, 220, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ---- list Přehled: součty a měsíce
  let ov = ss.getSheetByName('Přehled') || ss.insertSheet('Přehled', 0);
  ov.clear();
  const sum = a => a.reduce((s, r) => s + (Number(r[2]) || 0), 0);
  const open = rows.filter(r => !sent(r)), months = {};
  rows.forEach(r => { const d = isD(r[COL_DDUZP - 1]) ? r[COL_DDUZP - 1] : isD(r[1]) ? r[1] : null; const k = d ? Utilities.formatDate(d, tz, 'yyyy-MM') : '(bez data)'; const m = months[k] || (months[k] = { n: 0, s: 0, open: 0, inv: 0 }); m.n++; m.s += Number(r[2]) || 0; if (!sent(r)) m.open++; if (String(r[COL_DTYPE - 1]) === 'faktura') m.inv++; });
  const out = [
    ['Přehled dokladů SVJ Od Vysoké', '', '', '', ''],
    ['Aktualizováno ' + Utilities.formatDate(new Date(), tz, 'd. M. yyyy H:mm') + ' · podrobnosti a odkazy na doklady jsou v listu Doklady', '', '', '', ''],
    ['', '', '', '', ''],
    ['Dokladů celkem', rows.length, '', 'Celkem (Kč)', sum(rows)],
    ['Z toho faktur', rows.filter(r => String(r[COL_DTYPE - 1]) === 'faktura').length, '', 'Paragony (Kč)', sum(rows.filter(r => String(r[COL_DTYPE - 1]) !== 'faktura'))],
    ['Neodesláno účetní', open.length, '', 'Neodesláno (Kč)', sum(open)],
    ['', '', '', '', ''],
    ['Měsíc (podle DUZP / data dokladu)', 'Dokladů', 'Z toho faktur', 'Neodesláno', 'Celkem (Kč)']
  ];
  Object.keys(months).sort().reverse().forEach(k => { const m = months[k]; out.push([k, m.n, m.inv, m.open, m.s]); });
  ov.getRange(1, 1, out.length, 5).setValues(out);
  ov.getRange(1, 1).setFontSize(16).setFontWeight('bold'); ov.getRange(2, 1).setFontColor('#6b7280');
  ov.getRange(4, 1, 3, 1).setFontWeight('bold'); ov.getRange(4, 4, 3, 1).setFontWeight('bold');
  ov.getRange(4, 5, 3, 1).setNumberFormat('#,##0.00 "Kč"');
  ov.getRange(8, 1, 1, 5).setFontWeight('bold').setBackground('#eef2f7');
  if (out.length > 8) ov.getRange(9, 5, out.length - 8, 1).setNumberFormat('#,##0.00 "Kč"');
  if (open.length) ov.getRange(6, 1, 1, 5).setBackground('#fff8e1');
  [260, 90, 110, 130, 140].forEach((w, i) => ov.setColumnWidth(i + 1, w));
  ss.setActiveSheet(ov);
}
function dashRefreshSafe_() { try { dashRefresh_(); } catch (e) { Logger.log('Přehled dokladů se nepodařilo obnovit: ' + e.message); } }

// čtenáři přehledu (výbor): právo číst přehled a složku Doklady (kvůli odkazům); jinam do SVJ_app přístup nedostanou
function shareDash_(ss) {
  const list = mailList_(props_().getProperty('DASH_VIEWERS'));
  if (!list.length) return [];
  const done = [];
  try { DriveApp.getFileById((ss || dashSS_()).getId()).addViewers(list); done.push('přehled'); } catch (e) { Logger.log('Sdílení přehledu: ' + e.message); }
  try { child_(driveRoot_(), DOCS_FOLDER).addViewers(list); done.push('složka Doklady'); } catch (e) { Logger.log('Sdílení složky Doklady: ' + e.message); }
  return done;
}

// ---------- menu ----------
function refreshDash() {
  dashRefresh_();
  const ss = dashSS_();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p style="font:14px Arial">Přehled je obnovený:</p><p style="font:14px Arial"><a href="' + ss.getUrl() + '" target="_blank">' + escH_(ss.getName()) + '</a> (složka „' + DASH_FOLDER + '“)</p>').setWidth(440).setHeight(120), 'Přehled dokladů');
}
function setDashViewers() {
  const ui = SpreadsheetApp.getUi(), cur = props_().getProperty('DASH_VIEWERS') || '';
  const r = ui.prompt('Kdo uvidí přehled dokladů', 'E-maily kolegů z výboru (oddělené čárkou). Dostanou právo číst přehled a otevírat odkazy na doklady (složka Doklady). Do SVJ_app ani do tabulky měřidel přístup nedostanou.' + (cur ? ' Teď: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const list = mailList_(r.getResponseText());
  if (!list.length) { ui.alert('Nenašel jsem žádnou platnou adresu. Přístup, který už někdo má, se tímto neodebírá; odebrat ho jde v Disku přes Sdílet.'); return; }
  props_().setProperty('DASH_VIEWERS', list.join(','));
  dashRefresh_();
  const done = shareDash_();
  ui.alert('Uloženo: ' + list.join(', ') + '\n\nSdíleno pro čtení: ' + (done.join(' a ') || 'nic (podívej se do protokolu)') + '. Google jim může poslat e-mail o sdílení.');
}
