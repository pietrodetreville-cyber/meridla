// ---------- přehled dokladů pro výbor: list „Přehled“ v souboru Doklady ----------
// Kolegové z výboru nepotřebují přístup do Administrace ani k měřidlům: dostanou právo číst soubor Doklady a složku Doklady (kvůli odkazům na fotky / PDF).
// Přehled se obnoví po uložení dokladu, registraci faktury a odeslání účetní; ručně menu Doklady → Obnovit přehled dokladů.
// (Dřívější samostatný soubor „Přehled dokladů…“ ve složce dokumenty se už neaktualizuje; skript ho nemaže.)
function dashSS_() { return docsSS_(); }

// obnoví celý přehled z listu Doklady
function dashRefresh_() {
  const dss = docsSSIfExists_(), src = dss && dss.getSheetByName(SHEET_DOCS);
  if (!src) return;
  try { clearSentChecks_(src); } catch (e) {}
  const ss = dashSS_(), tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const rows = docRows_(src).filter(r => r[9]);
  const isD = x => Object.prototype.toString.call(x) === '[object Date]';
  const sent = r => /^odesl/i.test(String(r[8] || ''));
  const num = r => String(r[COL_DNUM - 1] || '');
  const isInv = r => String(r[COL_DTYPE - 1]) === 'faktura';
  const day = r => isD(r[COL_DDUZP - 1]) ? r[COL_DDUZP - 1] : isD(r[1]) ? r[1] : null;
  rows.sort((a, b) => num(b).localeCompare(num(a)) || new Date(b[0]) - new Date(a[0]));
  const BLUE = '#2563eb', INK = '#1f2933', MUTED = '#6b7280', KC = '#,##0.00 "Kč"', KC0 = '#,##0 "Kč"';
  const NEW_BG = '#fff4cc', NEW_FG = '#8a6100', OK_BG = '#dcf3e3', OK_FG = '#166534';
  // vyčistí list včetně grafů, sloučených buněk, pruhování a podmíněného formátu
  const wipe = s => {
    if (s.getFilter()) s.getFilter().remove();
    s.getCharts().forEach(c => s.removeChart(c));
    s.getBandings().forEach(b => b.remove());
    s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).breakApart();
    s.setConditionalFormatRules([]);
    s.clear();
  };
  const band = (rg, headBg) => rg.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false)
    .setHeaderRowColor(headBg).setFirstRowColor('#ffffff').setSecondRowColor('#f4f7fc');

  // ---- list Přehled: pruh s názvem, dlaždice, výdaje po měsících, poslední doklady
  const ov = ss.getSheetByName('Přehled') || ss.insertSheet('Přehled', 0);
  wipe(ov);
  ov.setHiddenGridlines(true);
  const W = 8, sum = a => a.reduce((s, r) => s + (Number(r[2]) || 0), 0);
  const open = rows.filter(r => !sent(r)), invs = rows.filter(isInv), months = {};
  const MES = ['leden', 'únor', 'březen', 'duben', 'květen', 'červen', 'červenec', 'srpen', 'září', 'říjen', 'listopad', 'prosinec'];
  const mName = k => /^\d{4}-\d{2}$/.test(k) ? MES[Number(k.slice(5)) - 1] + ' ' + k.slice(0, 4) : k;
  const nDocs = n => n + (n === 1 ? ' doklad' : n > 1 && n < 5 ? ' doklady' : ' dokladů');
  rows.forEach(r => { const d = day(r); const k = d ? Utilities.formatDate(d, tz, 'yyyy-MM') : '(bez data)'; const m = months[k] || (months[k] = { n: 0, s: 0, open: 0, inv: 0 }); m.n++; m.s += Number(r[2]) || 0; if (!sent(r)) m.open++; if (isInv(r)) m.inv++; });
  const nowK = Utilities.formatDate(new Date(), tz, 'yyyy-MM'), cur = months[nowK] || { n: 0, s: 0 };
  [150, 105, 105, 105, 120, 105, 105, 120].forEach((w, i) => ov.setColumnWidth(i + 1, w));

  // pruh s názvem
  ov.getRange(1, 1, 1, W).merge().setValue('Přehled dokladů SVJ Od Vysoké').setFontSize(18).setFontWeight('bold');
  ov.getRange(2, 1, 1, W).merge().setValue('Aktualizováno ' + Utilities.formatDate(new Date(), tz, 'd. M. yyyy H:mm') + '  ·  všechny doklady s odkazy jsou v listu Doklady').setFontSize(10);
  ov.getRange(1, 1, 2, W).setBackground(BLUE).setFontColor('#ffffff').setVerticalAlignment('middle').setHorizontalAlignment('left');
  ov.getRange(2, 1, 1, W).setFontColor('#dbe7ff');
  ov.setRowHeight(1, 44); ov.setRowHeight(2, 24); ov.setRowHeight(3, 14);

  // dlaždice: popisek, velké číslo, doplněk
  const cards = [
    ['Dokladů celkem', rows.length, '0', 'faktury ' + invs.length + ' · paragony ' + (rows.length - invs.length), '#e8f0fe', '#1d4ed8'],
    ['Celkem', sum(rows), KC0, 'z toho faktury ' + fmtKc_(sum(invs)), '#e6f4ea', OK_FG],
    ['Neodesláno účetní', open.length, '0', open.length ? fmtKc_(sum(open)) : 'vše odesláno ✓', open.length ? NEW_BG : '#e6f4ea', open.length ? NEW_FG : OK_FG],
    ['Tento měsíc', cur.s, KC0, mName(nowK) + ' · ' + nDocs(cur.n), '#f3e8fd', '#7e22ce']
  ];
  cards.forEach((c, i) => {
    const col = 1 + i * 2, box = ov.getRange(4, col, 3, 2);
    box.setBackground(c[4]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ov.getRange(4, col, 1, 2).merge().setValue(c[0].toUpperCase()).setFontSize(9).setFontWeight('bold').setFontColor(c[5]);
    ov.getRange(5, col, 1, 2).merge().setValue(c[1]).setNumberFormat(c[2]).setFontSize(22).setFontWeight('bold').setFontColor(INK);
    ov.getRange(6, col, 1, 2).merge().setValue(c[3]).setFontSize(9).setFontColor(MUTED);
    box.setBorder(true, true, true, true, null, null, '#ffffff', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  ov.setRowHeight(4, 26); ov.setRowHeight(5, 42); ov.setRowHeight(6, 24); ov.setRowHeight(7, 18);

  const section = (row, text) => {
    ov.getRange(row, 1, 1, W).merge().setValue(text).setFontSize(12).setFontWeight('bold').setFontColor(BLUE).setVerticalAlignment('bottom')
      .setBorder(null, null, true, null, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID);
    ov.setRowHeight(row, 30);
  };

  // výdaje po měsících: tabulka vlevo (nejnovější nahoře), graf vpravo (posledních 12 měsíců)
  section(8, 'Výdaje po měsících');
  const keys = Object.keys(months).filter(k => k !== '(bez data)').sort().reverse();
  const dated = keys.length;
  if (months['(bez data)']) keys.push('(bez data)');
  const T = 10, mt = [['Měsíc', 'Dokladů', 'Faktur', 'Neodesláno', 'Celkem']].concat(keys.map(k => { const m = months[k]; return [mName(k), m.n, m.inv, m.open, m.s]; }));
  ov.setRowHeight(9, 8);
  ov.getRange(T, 1, mt.length, 1).setNumberFormat('@');
  ov.getRange(T, 1, mt.length, 5).setValues(mt).setVerticalAlignment('middle').setFontColor(INK);
  ov.getRange(T, 1, 1, 5).setFontWeight('bold');
  ov.getRange(T, 2, mt.length, 3).setHorizontalAlignment('center');
  ov.getRange(T, 5, mt.length, 1).setHorizontalAlignment('right');
  if (mt.length > 1) {
    ov.getRange(T + 1, 5, mt.length - 1, 1).setNumberFormat(KC).setFontWeight('bold');
    ov.getRange(T + 1, 4, mt.length - 1, 1).setFontColor(NEW_FG);
    band(ov.getRange(T, 1, mt.length, 5), '#dbe7ff');
  }
  if (dated) {
    const k = Math.min(12, dated);
    ov.insertChart(ov.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(ov.getRange(T, 1, k + 1, 1)).addRange(ov.getRange(T, 5, k + 1, 1)).setNumHeaders(1)
      .setPosition(T, 6, 10, 0)
      .setOption('title', 'Výdaje po měsících (Kč)').setOption('titleTextStyle', { color: INK, fontSize: 13, bold: true })
      .setOption('legend.position', 'none').setOption('colors', [BLUE])
      .setOption('hAxis.direction', -1).setOption('hAxis.textStyle', { color: MUTED, fontSize: 10 })
      .setOption('vAxis.format', '#,##0').setOption('vAxis.textStyle', { color: MUTED, fontSize: 10 }).setOption('vAxis.gridlines.color', '#eef1f5').setOption('vAxis.minValue', 0)
      .setOption('bar.groupWidth', '65%').setOption('backgroundColor', '#ffffff')
      .setOption('width', 330).setOption('height', 250).build());
  }

  // posledních 8 dokladů s odkazem (graf je vysoký asi 12 řádků)
  const L = Math.max(T + mt.length, T + 12) + 1, last = rows.slice(0, 8);
  section(L, 'Poslední doklady');
  ov.setRowHeight(L + 1, 8);
  const H = L + 2;
  const lt = [['Číslo', 'Datum', 'Dodavatel', '', 'Účel', '', 'Částka', 'Stav']].concat(last.map(r => [num(r), day(r) || '', String(r[3] || ''), '', String(r[4] || ''), '', Number(r[2]) || 0, sent(r) ? 'odesláno' : 'nový']));
  ov.getRange(H, 1, lt.length, 1).setNumberFormat('@');
  ov.getRange(H, 1, lt.length, W).setValues(lt).setVerticalAlignment('middle').setFontColor(INK);
  ov.getRange(H, 3, lt.length, 2).mergeAcross(); ov.getRange(H, 5, lt.length, 2).mergeAcross();
  ov.getRange(H, 1, 1, W).setFontWeight('bold');
  ov.getRange(H, 7, lt.length, 1).setHorizontalAlignment('right');
  ov.getRange(H, 8, lt.length, 1).setHorizontalAlignment('center');
  if (last.length) {
    band(ov.getRange(H, 1, lt.length, W), '#dbe7ff');
    ov.getRange(H + 1, 2, last.length, 1).setNumberFormat('d. M. yyyy');
    ov.getRange(H + 1, 7, last.length, 1).setNumberFormat(KC).setFontWeight('bold');
    ov.getRange(H + 1, 8, last.length, 1).setFontWeight('bold')
      .setBackgrounds(last.map(r => [sent(r) ? OK_BG : NEW_BG])).setFontColors(last.map(r => [sent(r) ? OK_FG : NEW_FG]));
    last.forEach((r, i) => { const u = String(r[7] || ''); if (/^https?:/.test(u) && num(r)) setPhotoLink_(ov, H + 1 + i, 1, (isInv(r) ? '📄 ' : '📷 ') + num(r), u); });
  } else ov.getRange(H + 1, 1, 1, W).merge().setValue('Zatím žádné doklady.').setFontColor(MUTED);
}
function dashRefreshSafe_() { try { dashRefresh_(); } catch (e) { Logger.log('Přehled dokladů se nepodařilo obnovit: ' + e.message); } }

// čtenáři přehledu (výbor): právo číst přehled a složku Doklady (kvůli odkazům); jinam do SVJ_app přístup nedostanou
function shareDash_(ss) {
  const list = mailList_(props_().getProperty('DASH_VIEWERS'));
  if (!list.length) return [];
  const done = [];
  try { DriveApp.getFileById((ss || dashSS_()).getId()).addViewers(list); done.push('soubor Doklady s přehledem'); } catch (e) { Logger.log('Sdílení přehledu: ' + e.message); }
  try { child_(driveRoot_(), DOCS_FOLDER).addViewers(list); done.push('složka Doklady'); } catch (e) { Logger.log('Sdílení složky Doklady: ' + e.message); }
  return done;
}

// ---------- menu ----------
function refreshDash() {
  dashRefresh_();
  const ss = dashSS_();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p style="font:14px Arial">Přehled je obnovený:</p><p style="font:14px Arial"><a href="' + ss.getUrl() + '" target="_blank">' + escH_(ss.getName()) + '</a> (list Přehled)</p>').setWidth(440).setHeight(120), 'Přehled dokladů');
}
function setDashViewers() {
  const ui = SpreadsheetApp.getUi(), cur = props_().getProperty('DASH_VIEWERS') || '';
  const r = ui.prompt('Kdo uvidí přehled dokladů', 'E-maily kolegů z výboru (oddělené čárkou). Dostanou právo číst soubor Doklady (s přehledem) a otevírat odkazy na doklady (složka Doklady). Do Administrace ani do tabulky měřidel přístup nedostanou.' + (cur ? ' Teď: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const list = mailList_(r.getResponseText());
  if (!list.length) { ui.alert('Nenašel jsem žádnou platnou adresu. Přístup, který už někdo má, se tímto neodebírá; odebrat ho jde v Disku přes Sdílet.'); return; }
  props_().setProperty('DASH_VIEWERS', list.join(','));
  dashRefresh_();
  const done = shareDash_();
  ui.alert('Uloženo: ' + list.join(', ') + '\n\nSdíleno pro čtení: ' + (done.join(' a ') || 'nic (podívej se do protokolu)') + '. Google jim může poslat e-mail o sdílení.');
}
