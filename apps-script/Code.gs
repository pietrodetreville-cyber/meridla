/**
 * Příjem odečtů měřidel: fotka -> složka na Disku, hodnota -> tabulka, upozornění e-mailem.
 *
 * Jednorázově: spusť funkci setup() (založí listy, složku a tajný kód).
 * Potom: Nasadit -> Nová nasazení -> Webová aplikace (Spustit jako: já, Přístup: kdokoli).
 *
 * Aplikace musí POST posílat s Content-Type: text/plain (jinak prohlížeč vyžaduje preflight,
 * který Apps Script neumí).
 */

const SHEET_METERS = 'Měřidla';
const SHEET_READINGS = 'Odečty';
const SHEET_PEOPLE = 'Lidé';
const FOLDER_NAME = 'Odečty měřidel – fotky';
const DEFAULT_APP_URL = 'https://pietrodetreville-cyber.github.io/meridla/app/';

// ---------- jednorázové nastavení ----------
function setup() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  let m = ss.getSheetByName(SHEET_METERS);
  if (!m) {
    m = ss.insertSheet(SHEET_METERS);
    m.appendRow(['id', 'Název', 'Typ', 'Číslic před čárkou', 'Nápověda pro čtení', 'Jednotka']);
    m.appendRow(['elektro-kotelna', 'Elektroměr – kotelna', 'digitální', 5, 'Číslice za desetinnou čárkou (červené) ignoruj.', 'kWh']);
    m.appendRow(['voda-sklep', 'Vodoměr – sklep', 'kolečka', 5, 'Červená kolečka jsou desetiny m³.', 'm³']);
    m.setFrozenRows(1);
  }
  let r = ss.getSheetByName(SHEET_READINGS);
  if (!r) {
    r = ss.insertSheet(SHEET_READINGS);
    r.appendRow(['Čas', 'id', 'Měřidlo', 'Stav', 'Předchozí', 'Přírůstek', 'Poznámka', 'Fotka', 'Zdroj']);
    r.setFrozenRows(1);
  }

  let p = ss.getSheetByName(SHEET_PEOPLE);
  if (!p) {
    p = ss.insertSheet(SHEET_PEOPLE);
    p.appendRow(['Jméno', 'Tajný kód (negenerovat ručně)', 'Odkaz pro tuto osobu']);
    p.setFrozenRows(1);
  }
  if (!r.getRange(1, 10).getValue()) r.getRange(1, 10).setValue('Kdo');
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', DriveApp.createFolder(FOLDER_NAME).getId());
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  if (!props.getProperty('NOTIFY')) props.setProperty('NOTIFY', Session.getEffectiveUser().getEmail());

  Logger.log('TAJNÝ KÓD (token): ' + props.getProperty('TOKEN'));
  Logger.log('Upozornění chodí na: ' + props.getProperty('NOTIFY'));
  Logger.log('Složka s fotkami: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Měřidla').addItem('Vygenerovat odkazy pro lidi', 'generateLinks').addToUi();
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
function props_() { return PropertiesService.getScriptProperties(); }
// Vrací jméno člověka podle tajného kódu, nebo '' když kód neplatí.
// Hlavní kód (vlastnost TOKEN) patří "Správci", ostatní kódy jsou v listu Lidé.
function who_(t) {
  if (!t) return '';
  const master = props_().getProperty('TOKEN');
  if (master && t === master) return 'Správce';
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PEOPLE);
  if (!sh || sh.getLastRow() < 2) return '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const hit = rows.find(r => String(r[0]).trim() && r[1] && String(r[1]) === t);
  return hit ? String(hit[0]).trim() : '';
}

function meters_() {
  const v = SpreadsheetApp.getActive().getSheetByName(SHEET_METERS).getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => ({ id: String(r[0]), name: r[1], type: r[2], digits: r[3], hint: r[4], unit: r[5] }));
}

function readings_() {
  const v = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS).getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => ({ time: new Date(r[0]), id: String(r[1]), value: Number(r[3]) }));
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
  const who = who_(e.parameter.token);
  if (!who) return ok_({ ok: false, error: 'token' });
  const ws = weekStart_(), all = readings_();
  const list = meters_().map(m => {
    const mine = all.filter(x => x.id === m.id).sort((a, b) => b.time - a.time);
    const last = mine[0];
    return { ...m, last: last ? { value: last.value, time: last.time.toISOString() } : null, doneThisWeek: !!last && last.time >= ws };
  });
  return ok_({ ok: true, who: who, meters: list });
}

// Spusť jednou z editoru: vyvolá povolení externího připojení a ověří klíč Gemini.
function authorize() {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) { Logger.log('Chybí vlastnost GEMINI_KEY (Nastavení projektu -> Vlastnosti skriptu).'); return; }
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=5', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  Logger.log('Gemini odpověděl: HTTP ' + r.getResponseCode() + (r.getResponseCode() === 200 ? ' (klíč funguje)' : ' – ' + r.getContentText().replace(/s+/g, ' ').slice(0, 200)));
  Logger.log('Vybrané modely: ' + JSON.stringify(geminiModels_()) + (AI_DETAIL ? ' / ' + AI_DETAIL : ''));
}

// ---------- čtení stavu z fotky (Gemini, klíč je jen zde v skriptu) ----------
// Klíč: Nastavení projektu -> Vlastnosti skriptu -> GEMINI_KEY
let AI_DETAIL = '';
function geminiModels_() {
  const cache = CacheService.getScriptCache(), c = cache.get('models');
  if (c) return JSON.parse(c);
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return [];
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) { AI_DETAIL = 'seznam modelů HTTP ' + r.getResponseCode() + ': ' + r.getContentText().replace(/s+/g, ' ').slice(0, 150); return []; }
  const names = (JSON.parse(r.getContentText()).models || [])
    .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
    .map(m => m.name.replace('models/', ''))
    .filter(n => /^gemini-/.test(n) && /flash/.test(n) && !/image|tts|live|audio|embed|robotics|computer|native|exp|thinking|customtools/.test(n));
  const score = n => (/preview/.test(n) ? 0 : 1000) + parseFloat((n.match(/gemini-(d+(?:.d+)?)/) || [0, 0])[1]) * 10 - (/lite/.test(n) ? 5 : 0) + (/latest/.test(n) ? 0.5 : 0);
  names.sort((a, b) => score(b) - score(a));
  const top = names.slice(0, 3);
  if (top.length) cache.put('models', JSON.stringify(top), 21600);
  return top;
}

function readWithAi_(meter, b64) {
  const key = props_().getProperty('GEMINI_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  if (!meter || !b64) return { ok: false, error: 'data' };
  const models = geminiModels_();
  if (!models.length) return { ok: false, error: 'nomodel', detail: AI_DETAIL || 'žádný model Gemini Flash' };
  const prompt = 'Na fotce je měřidlo: ' + meter.name + ' (typ: ' + meter.type + ', jednotka: ' + meter.unit + ', počet číslic před čárkou: ' + meter.digits + '). ' +
    (meter.hint ? meter.hint + ' ' : '') +
    'Přečti aktuální stav měřidla jako jedno číslo s desetinnou tečkou. Odpověz jen JSON {"value": číslo}. Když stav nejde spolehlivě přečíst, dej {"value": null}.';
  const payload = JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0 } });
  const res = UrlFetchApp.fetchAll(models.map(m => ({
    url: 'https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', method: 'post',
    contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: payload, muteHttpExceptions: true })));
  const det = [];
  for (let i = 0; i < res.length; i++) {
    const r = res[i];
    if (r.getResponseCode() !== 200) { det.push(models[i] + ' HTTP ' + r.getResponseCode() + ': ' + r.getContentText().replace(/s+/g, ' ').slice(0, 120)); continue; }
    try {
      const j = JSON.parse(r.getContentText()), t = j.candidates[0].content.parts.map(p => p.text || '').join('');
      const v = JSON.parse(t.replace(/```json|```/g, '')).value;
      if (v !== null && isFinite(Number(v))) return { ok: true, value: Number(v) };
      det.push(models[i] + ': odpověď bez čísla');
    } catch (e) { det.push(models[i] + ': ' + e.message); }
  }
  return { ok: false, error: 'ai', detail: det.join(' | ') };
}

// ---------- příjem odečtu ----------
// tělo (JSON jako text): { token, id, value, note, source, photo: "<base64 jpeg>" }
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return ok_({ ok: false, error: 'json' }); }
  const who = who_(body.token);
  if (!who) return ok_({ ok: false, error: 'token' });
  if (body.action === 'read') return ok_(readWithAi_(meters_().find(m => m.id === body.id), body.photo));

  const meter = meters_().find(m => m.id === body.id);
  const value = Number(body.value);
  if (!meter || !isFinite(value)) return ok_({ ok: false, error: 'data' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_READINGS);
    const prev = readings_().filter(x => x.id === meter.id).sort((a, b) => b.time - a.time)[0];
    const diff = prev ? Math.round((value - prev.value) * 1000) / 1000 : '';
    let note = String(body.note || '').slice(0, 300);
    if (prev && value < prev.value) note = ('POZOR: nižší než minule. ' + note).trim();

    let fileUrl = '';
    if (body.photo) {
      const stamp = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd_HH-mm');
      const blob = Utilities.newBlob(Utilities.base64Decode(body.photo), 'image/jpeg', stamp + '_' + meter.id + '.jpg');
      fileUrl = DriveApp.getFolderById(props_().getProperty('FOLDER_ID')).createFile(blob).getUrl();
    }
    sheet.appendRow([new Date(), meter.id, meter.name, value, prev ? prev.value : '', diff, note, fileUrl, body.source || '', who]);

    const to = props_().getProperty('NOTIFY');
    if (to) {
      MailApp.sendEmail(to, 'Odečet: ' + meter.name + ' = ' + value + ' ' + (meter.unit || ''),
        meter.name + '\nFotil(a): ' + who + '\nStav: ' + value + ' ' + (meter.unit || '') +
        (prev ? '\nPředchozí: ' + prev.value + ' (přírůstek ' + diff + ')' : '') +
        (note ? '\nPoznámka: ' + note : '') + (fileUrl ? '\nFotka: ' + fileUrl : ''));
    }
    return ok_({ ok: true, diff: diff, warn: note.indexOf('POZOR') === 0 });
  } finally {
    lock.releaseLock();
  }
}

// ---------- volitelné: týdenní připomínka ----------
// Spusť installReminder() jednou; každý pátek v 9 h pošle e-mail, co tento týden chybí.
function installReminder() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'reminder').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('reminder').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
}
function reminder() {
  const ws = weekStart_(), all = readings_();
  const missing = meters_().filter(m => !all.some(x => x.id === m.id && x.time >= ws)).map(m => m.name);
  const to = props_().getProperty('NOTIFY');
  if (missing.length && to) MailApp.sendEmail(to, 'Chybí odečty měřidel', 'Tento týden zatím chybí:\n- ' + missing.join('\n- '));
}
