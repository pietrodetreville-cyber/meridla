// ---------- došlé faktury (PDF): výběr ze složky na Disku, přečtení AI, registrace do listu Doklady ----------
// PDF faktury se nahrají do složky SVJ_app/OSTRE (TESTOVACI)/Došlé faktury. Menu Doklady → Zaregistrovat došlou fakturu (PDF)
// otevře okno: seznam PDF ve složce, náhled, AI předvyplní údaje. Po registraci dostane faktura interní číslo rrrr-mm-nnn
// (stejná řada jako paragony), PDF se přejmenuje na číslo a přesune do složky Doklady. Odeslání účetní je stejné jako u paragonů.

const INBOX_FOLDER = 'Došlé faktury';
// faktura může být PDF nebo fotka / sken (obrázek)
const INV_TYPES = ['application/pdf', 'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif'];

function invInbox_() { return child_(driveRoot_(), INBOX_FOLDER); }

function invoiceDialog() {
  invInbox_(); // složka se založí, pokud chybí
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('InvoiceUi').setWidth(1150).setHeight(760), 'Došlá faktura');
}

// seznam faktur (PDF i fotky) ve složce Došlé faktury (nejnovější nahoře)
function invUiList() {
  const folder = invInbox_(), it = folder.getFiles(), out = [], tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  while (it.hasNext() && out.length < 200) {
    const f = it.next();
    if (INV_TYPES.indexOf(String(f.getMimeType()).toLowerCase()) < 0) continue;
    out.push({ id: f.getId(), name: f.getName(), t: f.getDateCreated().getTime(), date: Utilities.formatDate(f.getDateCreated(), tz, 'd. M. yyyy H:mm'), kb: Math.round(f.getSize() / 1024) });
  }
  out.sort((a, b) => b.t - a.t);
  // okno běží i ze souboru dokladů, kde list Měřidla není: budovy si pamatuje vlastnost BUILDINGS (ukládá se při načtení aplikace)
  const sh = docsSheet_();
  let bs = String(props_().getProperty('BUILDINGS') || '').split('|').filter(Boolean);
  if (!bs.length) { try { bs = meters_().map(m => m.building).filter((b, i, a) => b && a.indexOf(b) === i); } catch (e) {} }
  if (!bs.length) bs = ['271', '272', '273', '274', '275', 'Kotelna'];
  return { folderUrl: folder.getUrl(), files: out, buildings: bs.sort((a, b) => String(a).localeCompare(String(b), 'cs', { numeric: true })), docsUrl: sh.getParent().getUrl() };
}

// AI přečte údaje z PDF faktury (jen předvyplnění, člověk je zkontroluje)
function invUiRead(id) {
  try {
    const f = DriveApp.getFileById(String(id || ''));
    if (f.getSize() > 15 * 1024 * 1024) return { ok: false, error: 'Soubor je větší než 15 MB.' };
    return invRead_(Utilities.base64Encode(f.getBlob().getBytes()), f.getMimeType());
  } catch (e) { return { ok: false, error: e.message }; }
}
// společné čtení faktury (okno v tabulce i aplikace v telefonu)
function invRead_(b64, mimeIn) {
  try {
    const key = props_().getProperty('GEMINI_KEY');
    if (!key) return { ok: false, error: 'Chybí klíč AI (GEMINI_KEY).' };
    if (!b64) return { ok: false, error: 'Chybí soubor faktury.' };
    const mime = String(mimeIn || '').toLowerCase();
    if (INV_TYPES.indexOf(mime) < 0) return { ok: false, error: 'Tento typ souboru AI nepřečte (jen PDF nebo fotka JPG/PNG).' };
    const prompt = 'V příloze je došlá faktura (PDF nebo fotka) pro společenství vlastníků. Vyčti z ní: dodavatel (název firmy), ico (IČO dodavatele), cislo (číslo faktury / variabilní symbol faktury), vystaveni (datum vystavení), duzp (datum uskutečnění zdanitelného plnění), splatnost (datum splatnosti), castka (celková fakturovaná částka v Kč včetně DPH, ne zálohy ani přeplatek), predmet (krátce předmět plnění, nejvýš 8 slov). Data ve formátu YYYY-MM-DD. Co nejde přečíst, nech prázdné, nic nevymýšlej.';
    const schema = { type: 'OBJECT', properties: { dodavatel: { type: 'STRING', nullable: true }, ico: { type: 'STRING', nullable: true }, cislo: { type: 'STRING', nullable: true }, vystaveni: { type: 'STRING', nullable: true }, duzp: { type: 'STRING', nullable: true }, splatnost: { type: 'STRING', nullable: true }, castka: { type: 'NUMBER', nullable: true }, predmet: { type: 'STRING', nullable: true } } };
    const think = m => /gemini-2\.5/.test(m) ? { thinkingBudget: 0 } : /gemini-3/.test(m) ? { thinkingLevel: 'minimal' } : null;
    const pay = (m, withThink) => JSON.stringify({ contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mime, data: b64 } }] }], generationConfig: Object.assign({ responseMimeType: 'application/json', responseSchema: schema, temperature: 0.1, maxOutputTokens: 1024 }, withThink && think(m) ? { thinkingConfig: think(m) } : {}) });
    const call = (m, withThink) => UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: pay(m, withThink), muteHttpExceptions: true });
    const det = [], day = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? String(v) : '';
    for (const m of GEMINI_MODELS) {
      let r = call(m, true);
      if (r.getResponseCode() === 400 && think(m)) r = call(m, false);
      const code = r.getResponseCode();
      if (code !== 200) { det.push(m + ': ' + (code === 429 ? 'kvóta AI vyčerpána' : code === 404 ? 'model není dostupný' : code === 503 ? 'model je přetížený' : 'HTTP ' + code)); continue; }
      try {
        const j = JSON.parse(r.getContentText());
        const t = ((j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || []).map(p => p.text || '').join('');
        const o = JSON.parse(String(t).replace(/```json|```/g, ''));
        const amount = isFinite(docNum_(o.castka)) ? Math.round(docNum_(o.castka) * 100) / 100 : null;
        return { ok: true, vendor: String(o.dodavatel || '').trim(), ico: String(o.ico || '').replace(/\s/g, ''), invNo: String(o.cislo || '').trim(), issued: day(o.vystaveni), duzp: day(o.duzp), due: day(o.splatnost), amount: amount, subject: String(o.predmet || '').trim() };
      } catch (e) { det.push(m + ': ' + e.message); }
    }
    return { ok: false, error: 'AI fakturu nepřečetla: ' + det.join(' | ') };
  } catch (e) { return { ok: false, error: e.message }; }
}

// registrace faktury: zápis do listu Doklady (typ faktura), interní číslo, přejmenování a přesun PDF do složky Doklady
function invUiRegister(d) { const r = invUiRegisterCore_(d); if (r.ok) dashRefreshSafe_(); return r; }
// kontrola a převod údajů faktury (společné pro okno v tabulce i aplikaci)
function invParse_(d) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const parse = v => { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v || ''))) return null; try { return Utilities.parseDate(v + ' 12:00', tz, 'yyyy-MM-dd HH:mm'); } catch (e) { return null; } };
  const amount = Math.round(docNum_(d.amount) * 100) / 100, issued = parse(d.issued), duzp = parse(d.duzp), due = parse(d.due);
  if (!issued && !duzp) return { error: 'Zadej datum vystavení nebo DUZP.' };
  if (!isFinite(amount)) return { error: 'Zadej částku jako číslo.' };
  if (!String(d.purpose || '').trim()) return { error: 'Napiš účel (na co faktura je).' };
  const ico = String(d.ico || '').replace(/\s/g, '');
  const vendor = [String(d.vendor || '').trim(), ico ? 'IČO ' + ico : ''].filter(Boolean).join(', ');
  return { amount: amount, issued: issued, duzp: duzp, due: due, vendor: vendor, ico: ico, invNo: String(d.invNo || '').trim() };
}
// stejná faktura (číslo faktury a dodavatel podle IČO, jinak podle názvu) už zaregistrovaná? vrátí text upozornění
function invDup_(rows, d, v) {
  if (!v.invNo) return '';
  const ven = String(d.vendor || '').trim().toLowerCase();
  const same = rows.find(r => { const s = String(r[3] || ''); return String(r[COL_DINVNO - 1] || '').trim() === v.invNo && (v.ico ? s.indexOf('IČO ' + v.ico) >= 0 : !ven || s.toLowerCase().indexOf(ven) === 0); });
  return same ? 'Faktura ' + v.invNo + ' od ' + (d.vendor || 'tohoto dodavatele') + ' už je zaregistrovaná jako ' + (same[COL_DNUM - 1] || same[9]) + '.' : '';
}
// zápis řádku faktury do listu Doklady (soubor už je ve složce Doklady)
function invAppend_(sh, f, v, d, who, rid, num, note) {
  const row = docAppend_(sh, [new Date(), v.issued || v.duzp, v.amount, v.vendor, String(d.purpose).trim().slice(0, 200), String(d.building || '').slice(0, 40), who, f.getUrl(), 'nový', rid, String(note || '').slice(0, 300), '', false, '', '', num, 'faktura', v.invNo, v.duzp || '', v.due || '']);
  sh.getRange(row, dc_(sh, COL_DSEL)).insertCheckboxes();
  setPhotoLink_(sh, row, dc_(sh, 8), '📄 ' + num, f.getUrl());
}

// faktura z aplikace v telefonu: vyfocená nebo vybraná ze souboru (PDF / obrázek), údaje zkontroloval člověk
function invoiceSave_(body, who) {
  const v = invParse_(body);
  if (v.error) return { ok: false, error: v.error };
  const mime = String(body.mime || '').toLowerCase();
  if (!body.file) return { ok: false, error: 'Chybí soubor faktury.' };
  if (INV_TYPES.indexOf(mime) < 0) return { ok: false, error: 'Fakturu jde uložit jen jako PDF nebo fotku.' };
  let rid = String(body.clientId || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24);
  if (!rid) rid = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = docsSheet_(), rows = docRows_(sh);
    if (rows.some(r => String(r[9]) === rid)) return { ok: true, dup: true, rid: rid };
    const same = body.force ? '' : invDup_(rows, body, v);
    if (same) return { ok: false, dup: true, error: same };
    const num = nextDocNum_(sh, v.duzp || v.issued);
    const f = child_(driveRoot_(), DOCS_FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(body.file), mime, num + (isTest_() ? ' TEST' : '') + docExt_(mime)));
    invAppend_(sh, f, v, body, who, rid, num, body.note);
    return { ok: true, rid: rid, num: num };
  } finally { lock.releaseLock(); }
}

function invUiRegisterCore_(d) {
  if (!d.id) return { ok: false, error: 'Vyber fakturu (PDF nebo fotku).' };
  const v = invParse_(d);
  if (v.error) return { ok: false, error: v.error };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = docsSheet_(), rows = docRows_(sh);
    const same = d.force ? '' : invDup_(rows, d, v);
    if (same) return { ok: false, dup: true, error: same };
    let f; try { f = DriveApp.getFileById(String(d.id)); } catch (e) { return { ok: false, error: 'Soubor se nepodařilo najít na Disku.' }; }
    // jen PDF ze složky Došlé faktury tohoto prostředí (test a ostrá nesmí sdílet jeden soubor) a ještě nezaregistrované
    const inbox = invInbox_().getId(), parents = f.getParents(); let inInbox = false;
    while (parents.hasNext()) if (parents.next().getId() === inbox) inInbox = true;
    if (!inInbox) return { ok: false, error: 'Tento soubor není ve složce Došlé faktury ' + (isTest_() ? 'TESTU' : 'OSTRÉ verze') + '. Nahraj fakturu znovu do správné složky (stejný soubor nepoužívej v testu i na ostré).' };
    const fid = f.getId(), used = rows.find(r => String(r[7] || '').indexOf(fid) >= 0);
    if (used) return { ok: false, error: 'Tento soubor už je zaregistrovaný jako ' + (used[COL_DNUM - 1] || used[9]) + '.' };
    const num = nextDocNum_(sh, v.duzp || v.issued), rid = Utilities.getUuid().replace(/-/g, '').slice(0, 10), orig = f.getName();
    f.setName(num + (isTest_() ? ' TEST' : '') + docExt_(f.getMimeType()));
    try { f.moveTo(child_(driveRoot_(), DOCS_FOLDER)); } catch (e) {}
    const who = Session.getActiveUser().getEmail() || 'správce';
    invAppend_(sh, f, v, d, who, rid, num, [String(d.note || '').trim(), 'soubor: ' + orig].filter(Boolean).join(' · '));
    return { ok: true, num: num, name: f.getName() };
  } finally { lock.releaseLock(); }
}
