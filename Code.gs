// Dr. G.W. Williams Music Program - instrument inventory backend
// Lives inside a Google Sheet. Keeps two tabs: "Instruments" and "Log".
// Admin password is NOT in this file: set it in Project Settings > Script Properties (ADMIN_PASSWORD).

// Build tag: open your Web app URL with ?action=about on the end to see which version of this script is live.
const BUILD = 'assignments-locked-2026-11-14a';

// Most instruments one person (same first + last name; grade ignored) can have signed out at once. Admin is exempt.
const MAX_PER_PERSON = 2;
function person_(s) { return String(s == null ? '' : s).toLowerCase().replace(/\s*\(.*\)\s*$/, '').replace(/\s+/g, ' ').trim(); }

const HEAD = ['id', 'section', 'make', 'serial', 'primary', 'grade', 'secondary', 'status', 'signedBy'];
const LOG_HEAD = ['time', 'action', 'id', 'section', 'name'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (name === 'Instruments') s.getRange('A:I').setNumberFormat('@'); // keep serials like 086883 as text
  }
  if (name === 'Instruments' && s.getRange(1, 9).getValue() === '') s.getRange(1, 9).setValue('signedBy').setFontWeight('bold');
  return s;
}

function clean_(v) { return String(v == null ? '' : v).trim().slice(0, 120); }
// For text that students can type (who signed it out): no leading = + - @ so a Sheet or Excel never treats it as a formula.
function safe_(v) { return clean_(v).replace(/^[\s=+\-@]+/, '').slice(0, 80); }

function cleanItem_(i) {
  return { id: clean_(i.id), section: clean_(i.section), make: clean_(i.make), serial: clean_(i.serial),
           primary: clean_(i.primary), grade: clean_(i.grade), secondary: clean_(i.secondary),
           status: ['Assigned', 'Available', 'Out for Repair'].indexOf(i.status) >= 0 ? i.status : 'Available',
           by: i.status === 'Assigned' ? safe_(i.by) : '' };
}

function row_(i) { return [i.id, i.section, i.make, i.serial, i.primary, i.grade, i.secondary, i.status, i.by || '']; }

// Be forgiving if someone types a status by hand in the Sheet.
function status_(v) {
  const t = String(v == null ? '' : v).trim().toLowerCase();
  if (t.indexOf('repair') >= 0) return 'Out for Repair';
  if (t.indexOf('avail') >= 0 || t.indexOf('storage') >= 0 || t === '') return 'Available';
  return 'Assigned';
}

function items_() {
  const s = sheet_('Instruments', HEAD);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 9).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return { id: String(r[0]), section: String(r[1]), make: String(r[2]), serial: String(r[3]),
               primary: String(r[4]), grade: String(r[5]), secondary: String(r[6]), status: status_(r[7]),
               by: status_(r[7]) === 'Assigned' ? String(r[8] || '') : '' };
    });
}

function itemById_(id) { // reads just one row (always fresh, never cached)
  const s = sheet_('Instruments', HEAD);
  const last = s.getLastRow();
  if (last < 2) return null;
  const ids = s.getRange(1, 1, last, 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      const r = s.getRange(i + 1, 1, 1, 9).getValues()[0];
      return { id: String(r[0]), section: String(r[1]), make: String(r[2]), serial: String(r[3]),
               primary: String(r[4]), grade: String(r[5]), secondary: String(r[6]), status: status_(r[7]),
               by: status_(r[7]) === 'Assigned' ? String(r[8] || '') : '' };
    }
  }
  return null;
}
function upsert_(item, oldId) {
  const s = sheet_('Instruments', HEAD);
  const last = s.getLastRow();
  const ids = last > 0 ? s.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const idx = ids.indexOf(String(oldId || item.id));
  if (idx > 0) s.getRange(idx + 1, 1, 1, 9).setValues([row_(item)]);
  else s.appendRow(row_(item));
}

// Sign out / return / repair only ever touch the status and "signed out by" cells (columns H and I).
// The assigned student, grade and shared-with names in columns E-G are never written by these actions.
function setStatus_(id, status, by) {
  const s = sheet_('Instruments', HEAD);
  const last = s.getLastRow();
  const ids = last > 0 ? s.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const idx = ids.indexOf(String(id));
  if (idx > 0) s.getRange(idx + 1, 8, 1, 2).setValues([[status, by]]);
}

function logAdd_(type, item, name) {
  sheet_('Log', LOG_HEAD).appendRow([new Date(), type, item.id, item.section, safe_(name)]);
}

function logRead_() {
  const s = sheet_('Log', LOG_HEAD);
  if (s.getLastRow() < 2) return [];
  const start = Math.max(2, s.getLastRow() - 999); // newest 1000 entries
  return s.getRange(start, 1, s.getLastRow() - start + 1, 5).getValues().map(function (r) {
    return { t: new Date(r[0]).getTime(), a: String(r[1]), id: String(r[2]), section: String(r[3]), name: String(r[4]) };
  });
}

function isAdmin_(key) {
  const p = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  const k = String(key == null ? '' : key).trim();
  if (!p || k === '') return false;
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminFails') || 0);
  if (fails >= 10) return false;                 // too many wrong guesses: locked for 10 minutes
  if (k === p) return true;
  cache.put('adminFails', String(fails + 1), 600);
  return false;
}

// A "last changed" stamp for the Instruments tab. It is a fingerprint of the sheet's contents, so edits typed
// straight into the Sheet are noticed too. Cached for 5 seconds so many phones checking at once stay cheap.
// One cached copy of the instrument list + its version fingerprint. Rebuilt at most every 20 seconds,
// and immediately after any save (bump_). Makes loading and the 20-second polling much faster.
function snapshot_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('snap');
  if (hit) return JSON.parse(hit);
  const items = items_();
  const str = JSON.stringify(items);
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) | 0;
  const snap = { items: items, v: String(h) };
  try { cache.put('snap', JSON.stringify(snap), 20); } catch (err) {}
  return snap;
}
function version_() { return snapshot_().v; }
function bump_() { const c = CacheService.getScriptCache(); c.remove('v'); c.remove('snap'); }

// ---- YRDSB sign-in check ----
// Setup (Project Settings > Script Properties): GOOGLE_CLIENT_ID = the same OAuth client ID that is pasted into the page.
// Optional: ALLOWED_DOMAINS = comma-separated list, default "gapps.yrdsb.ca,yrdsb.ca".
// The page sends a Google sign-in token with every request; we ask Google to verify it and check the email's domain.
function authed_(token) {
  token = String(token == null ? '' : token);
  if (!token) return '';
  const props = PropertiesService.getScriptProperties();
  const cid = props.getProperty('GOOGLE_CLIENT_ID');
  if (!cid) return '';                                   // not set up: nobody gets in (fails closed)
  const cache = CacheService.getScriptCache();
  const key = 'tok' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const hit = cache.get(key);
  if (hit) return hit;
  let info;
  try {
    const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return '';
    info = JSON.parse(res.getContentText());
  } catch (err) { return ''; }
  const now = Math.floor(Date.now() / 1000);
  const iss = String(info.iss || '');
  if (info.aud !== cid || Number(info.exp) <= now) return '';
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') return '';
  if (String(info.email_verified) !== 'true') return '';
  const email = String(info.email || '').toLowerCase();
  const domain = email.split('@')[1] || '';
  const allowed = (props.getProperty('ALLOWED_DOMAINS') || 'gapps.yrdsb.ca,yrdsb.ca')
    .split(',').map(function (d) { return d.trim().toLowerCase(); }).filter(Boolean);
  if (allowed.indexOf(domain) < 0) return '';
  try { cache.put(key, email, Math.max(1, Math.min(600, Number(info.exp) - now))); } catch (err) {}
  return email;
}
function denied_() { return json_({ ok: false, auth: true, error: 'Please sign in with your YRDSB school account (that account is not allowed).' }); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'about') {
    ensureData_();
    const all = items_();
    return json_({ ok: true, build: BUILD, dataVersion: PropertiesService.getScriptProperties().getProperty('DATA_VERSION'),
                   instruments: all.length, withAssignedStudent: all.filter(function (x) { return x.primary; }).length,
                   savesSignedOutBy: true, signOutChangesAssignedStudent: false, returnChangesAssignedStudent: false });
  }
  if (!authed_(e && e.parameter && e.parameter.token)) return denied_();
  ensureData_();
  if (a === 'version') return json_({ ok: true, v: version_() });
  if (a === 'list') { const s = snapshot_(); return json_({ ok: true, items: s.items, v: s.v }); }
  return json_({ ok: false, error: 'Unknown request.' });
}

function doPost(e) {
  let pre;
  try { pre = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  if (!authed_(pre && pre.token)) return denied_();
  ensureData_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);

    if (b.action === 'auth') return json_({ ok: isAdmin_(b.key) });

    if (b.action === 'log') {
      if (!isAdmin_(b.key)) return json_({ ok: false, error: 'Not allowed.' });
      return json_({ ok: true, log: logRead_() });
    }

    if (b.action === 'clearlog') { // admin only: wipe the activity log (keeps the header row)
      if (!isAdmin_(b.key)) return json_({ ok: false, error: 'Not allowed.' });
      const ls = sheet_('Log', LOG_HEAD);
      if (ls.getLastRow() > 1) ls.getRange(2, 1, ls.getLastRow() - 1, ls.getMaxColumns()).clearContent();
      return json_({ ok: true });
    }

    if (b.action === 'seed') { // first run only: fill an empty Sheet with the starting list
      if (items_().length) return json_({ ok: false, error: 'Sheet already has data.' });
      sheet_('Instruments', HEAD);
      (b.items || []).forEach(function (i) { upsert_(cleanItem_(i)); });
      bump_();
      return json_({ ok: true });
    }

    if (b.action === 'event') { // sign out / return are open to everyone; sending to repair is admin only
      const cur = itemById_(clean_(b.item && b.item.id));
      if (!cur) return json_({ ok: false, error: 'Instrument not found.' });
      const adminNow = isAdmin_(b.key);
      if (b.type === 'rep' && !adminNow)
        return json_({ ok: false, error: 'Only a teacher (admin) can send an instrument for repair.' });
      if (cur.status === 'Out for Repair' && !adminNow)
        return json_({ ok: false, error: cur.id + ' is in repair and cannot be signed out or returned. Please see your teacher.' });
      const n = JSON.parse(JSON.stringify(cur));
      if (b.type === 'out') {
        if (cur.status === 'Out for Repair')
          return json_({ ok: false, error: cur.id + ' is in repair and cannot be signed out.' });
        if (cur.status === 'Assigned')
          return json_({ ok: false, error: cur.id + ' is already signed out' + (cur.by ? ' by ' + cur.by : '') + '.' });
        // Assigned student, grade and "shared with" never change here. We only record WHO has it signed out right now.
        const nm = safe_(b.item.by);
        if (!nm) return json_({ ok: false, error: 'Please enter the name of the person signing it out.' });
        if (!adminNow) {
          const who = person_(nm);
          const have = items_().filter(function (x) { return x.status === 'Assigned' && person_(x.by) === who; });
          if (have.length >= MAX_PER_PERSON)
            return json_({ ok: false, error: nm.replace(/\s*\(.*\)\s*$/, '') + ' already has ' + have.length + ' instrument' + (have.length > 1 ? 's' : '') + ' signed out (' + have.map(function (x) { return x.id; }).join(', ') + '). Please return one before signing out another.' });
        }
        n.status = 'Assigned';
        n.by = nm;
        setStatus_(cur.id, 'Assigned', nm); logAdd_('out', n, nm);
      } else if (b.type === 'ret') {
        n.status = 'Available'; n.by = ''; // assigned student stays on the record
        setStatus_(cur.id, 'Available', ''); logAdd_('ret', n, cur.by || cur.primary);
      } else if (b.type === 'rep') {
        n.status = 'Out for Repair'; n.by = ''; // assigned student stays on the record
        setStatus_(cur.id, 'Out for Repair', ''); logAdd_('rep', n, cur.by || cur.primary);
      } else return json_({ ok: false, error: 'Unknown action.' });
      bump_();
      return json_({ ok: true, item: n });
    }

    if (b.action === 'update') { // add / edit a record - admin only
      if (!isAdmin_(b.key)) return json_({ ok: false, error: 'Not allowed.' });
      upsert_(cleanItem_(b.item), b.oldId);
      bump_();
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Unknown request.' });
  } catch (err) {
    return json_({ ok: false, error: 'Something went wrong. Try again.' });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Restore the original list ----------
// One instrument per line: tag|section|make|serial|assigned to|grade|also used by|status (V = in storage, R = in repair, blank = signed out)
const ORIGINAL = `
PIC-01|Piccolo|Yamaha (2020)|086883|Danielle Suen|11|Luna Zhou (Gr. 10)
PIC-02|Piccolo|Yamaha (2023)|168417|Evelyn Sonemanivong|10|Victoria Cheng (Gr. 12)
FLT-01|Flute|Yamaha (2025)|090456|Virika Lad|10|Frances Tang (Gr. 12)
FLT-04|Flute|Yamaha (2023) IB|S50454|Phoebe Lu|11|
FLT-05|Flute|Selmer|7281365|Deniise Ma|9|
FLT-06|Flute|Yamaha|3100005||||V
FLT-07|Flute|Jupiter|UD59539|Amelia Ng|9|
FLT-09|Flute|Philharmonic|E64066|Betty Liao|10|Akili Fu (Gr. 9)
FLT-11|Flute|Yamaha (2025)|087095|Kaitlyn Liu|9|Fiona Chen (Gr. 11)
FLT-12|Flute|Yamaha|016118|Eva Lin|9|
FLT-14|Flute|Yamaha|034121|Caroline Kim|11|Aidan Xie (Gr. 9)
FLT-15|Flute|Yamaha (2023) IB|005781|Victoria Cheng|12|
FLT-16|Flute|Yamaha (2022) IB|S07158|Gracie Wu|11|
FLT-17|Flute|Yamaha|028470|Yalaena Valencia|11|
FLT-18|Flute|Selmer|7281366|Priscilla Ye|9|
FLT-19|Flute|Yamaha (2022) IB|S07163|Katherine Tan|12|
FLT-24|Flute|Buffett|676608|Ashley Fu|9|
FLT-25|Flute|Yamaha|871802|Olivia Barros|9|
FLT-26|Flute|Yamaha|406720A|Karen Jia|10|
FLT-28|Flute|Selmer|7281111|Adele Poon|10|
OBO-01|Oboe|Yamaha|0-11871|Mavis Lee|10|Ankie Lo (Gr. 9)
OBO-02|Oboe|Yamaha|024307|Hillary Chen|10|Reanna Ko (Gr. 9)
OBO-03|Oboe|Yamaha Duet (2022)|0711246|Mehryad Bayat|11|
OBO-04|Oboe|Yamaha Duet (2025)|075951|Jaden Wu|11|
OBO-05|Oboe|Yamaha (2025)|016697|Izabella Pilcz|11|
OBO-TD|Oboe|TDSS Borrowed||Luna Zhu|10|Priscilla Ye (Gr. 9)
BSN-01|Bassoon|Fox|45468|Evan Trinh|11|Ella Guan (10), Linus Cheung (9)
BSN-02|Bassoon|Fox-41 (2024)|73932|Lauren Cassell|12|Callum Ho (Gr. 9)
BSN-EC|Bassoon|Fox (Emily Carr)||Leona Xiong|10|Jenny Ji (11), Wynn Zhao (9)
BSN-TD|Bassoon|TDSS Borrowed||Logan Leung|10|Jennifer He (Gr. 9)
CLA-01|Clarinet|Yamaha (2024)|R31245|Sarah McGuire|12|
CLA-02|Clarinet|Yamaha||Mia Cao|9|
CLA-03|Clarinet|Jupiter|UE62366|Jordan Ji|9|
CLA-04|Clarinet|Yamaha (2025)|R33927|Megan Guo|9|Alison Li (Gr. 10)
CLA-05|Clarinet|Yamaha (2024)|R30889|Sophie Zhang|10|Iris Liu (Gr. 9)
CLA-06|Clarinet|Yamaha (2025)|R32930|Jasmine Zhen|12|Teresa Xu (Gr. 10)
CLA-07|Clarinet|Yamaha (2024)||Liam Gaertner|12|Andre Duong (Gr. 9)
CLA-08|Clarinet|Yamaha (2024)||Archita Aasdhir|10|
CLA-09|Clarinet|Yamaha (2025)|R33923|Maya Mantia|12|Jayden Tang (Gr. 9)
CLA-10|Clarinet|Yamaha (2024)||Emily Snape|12|
CLA-11|Clarinet|Yamaha (2025)|R29147|Nina Garcia|12|Alison Wang (Gr. 9)
CLA-12|Clarinet|Yamaha|123860|Behrad Khorsandi|10|Chloe Zhu (Gr. 9)
CLA-13|Clarinet|Jupiter|UE62362|Ryan Lee|10|
CLA-14|Clarinet|Artley|39 16384|Hana Asgharian|9|
CLA-16|Clarinet|Selmer USA|69180|Rachael Yi|10|
CLA-17|Clarinet|Yamaha (2025)|R33253|Jayden Jung|11|Daphne Lo (Gr. 9)
CLA-18|Clarinet|Buffet|457440|Samuel Wang|9|
CLA-21|Clarinet|Yamaha|B12466|Yoav Mor|9|
CLA-22|Clarinet|Yamaha|028197A|Josie Either|9|
CLA-23|Clarinet|Yamaha|0-55137A|Annie Zhuang|9|
CLA-24|Clarinet|Selmer (2021)|P0168701|Chloe Du|9|
CLA-27|Clarinet|Yamaha|055426A|Nika Hosseini|9|
CLA-28|Clarinet|Yamaha|0-71666|Hylas Lin|9|Luna Zhu (Gr. 10)
CLA-29|Clarinet|Yamaha (2025)|R33402|J Henriques|12|Madison Henriques (Gr. 10)
CLA-30|Clarinet|Vito|50318|Amelia Young|9|
CLA-31|Clarinet|Artley|7100132|Nava Sherzai|9|
CLA-32|Clarinet|Vito|B09250|Oscar Ye|9|
CLA-33|Clarinet|Vito|B43953|Joseph Lee|10|Belinda Lu (Gr. 9)
CLA-34|Clarinet|Buffet|457368|Misha Kushal|11|Emily Reed (Gr. 9)
CLA-35|Clarinet|Buffet|7101961|Reagan Ye|9|
CLA-36|Clarinet|Yamaha (2019)|13637|Carina Meng|10|
CLA-37|Clarinet|Yamaha|7613||||R
CLA-38|Clarinet|Armstrong|219126|Amelia Behrooz|9|
CLA-39|Clarinet|Frankenstein||Ethan Cheng|9|
CLA-40|Clarinet|Selmer||Aiden Fridman|9|
BCL-01|Bass Clarinet|Jupiter|UE66347|Jacob Gold|11|Alex Zhou (Gr. 9)
BCL-02|Bass Clarinet|Jupiter|D80244|Viriginia Xia|11|Annie Li (Gr. 9)
BCL-03|Bass Clarinet|Selmer|67491|Jayden Tang|12|Aishia Gehlot (Gr. 9)
BCL-04|Bass Clarinet|Vito|678OE|Signy Nguyen|10|Matthew Wong (Gr. 9)
BCL-05|Bass Clarinet|Yamaha|002251|Jeffrey Lau|10|Tyler Karakashian (Gr. 9)
BCL-06|Bass Clarinet|Jupiter|L70940|Morgan Dionne|10|Larissa Han (Gr. 9)
BCL-07|Bass Clarinet|Jupiter|VE62578|Najla Najafi|10|Elvin Huang (Gr. 9)
BCL-08|Bass Clarinet|Jupiter|P71978|Ben Zhou|9|
SAX-S01|Soprano Sax|Jupiter|536601||||V
SAX-A01|Alto Sax|Jupiter 700 (2022)|WF01297|Eric Cheng|12|Jason Kwon (Gr. 9)
SAX-A02|Alto Sax|Yamaha|218140|Doris Lin|9|Evelyn Sonemanivong (10, Jazz)
SAX-A04|Alto Sax|Yamaha-61|4972|Katya Krinisyn|10|Charlotte Andary (Gr. 9)
SAX-A05|Alto Sax|Yamaha-61|5005|Sienna Mio|12|Victoria Papadimatos (Gr. 9)
SAX-A06|Alto Sax|Jupiter 700 (2022)|XF02521|Ella Guan|10|Andre Tang (Gr. 9)
SAX-A07|Alto Sax|Jupiter 700 (2022)|TF06737|Bellina Li|10|
SAX-A08|Alto Sax|Jupiter 700 (2022)|TF06620|Elora Rossouw|11|Yoav Mor (9, Jazz)
SAX-A09|Alto Sax|Jupiter 700 (2022)|XF07190|Lyla Garlow|9|
SAX-A10|Alto Sax|Evette|650156|Vanessa Munits|10|
SAX-A11|Alto Sax|Jupiter 700 (2022)|UF03619|Samuel Wang|11|
SAX-A12|Alto Sax|Jupiter 700 (2022)|XF07193|Jireh Andrews|10|Mina Gosse-Gomes (Gr. 9)
SAX-A14|Alto Sax|Yamaha|312577A|Damian Dayaljee|9|
SAX-A15|Alto Sax|Yamaha|030921|Alvin Zou|10|Zoe Zhu (Gr. 9)
SAX-A16|Alto Sax|Jupiter 700 (2022)|XF00635|Jane Du|10|Kimberly Cheung (Gr. 9)
SAX-A17|Alto Sax|Yamaha|201711 A|Deniz Ozcan|12|Audrina Weng (9, Jazz)
SAX-T01|Tenor Sax|Yamaha|0-60153|Wynn Zhao|9|Behrad Poorakbar (11, Jazz)
SAX-T02|Tenor Sax|Yamaha|0-54411|Sunny Feng|10|Jey Srianujan (Gr. 9)
SAX-T03|Tenor Sax|Yamaha YTS-480|R57245|Kosima Keo|10|Signy Nguyen (10, Jazz)
SAX-T04|Tenor Sax|Yamaha YTS-480|R60924|Jasmine Liu|9|Matthew Wong (9, Jazz)
SAX-T05|Tenor Sax|Jupiter|UF06279|Helya Keyhaninezhad|10|Boe Graham (Gr. 11)
SAX-T07|Tenor Sax|Jupiter|TF08168|Ali Meschi|10|Ethan Tran (Gr. 11)
SAX-T08|Tenor Sax|Yamaha YTS-480|R98710|Terrence Gu|12|Justin Lin (9), Aiden Xie (9, Jazz)
SAX-T09|Tenor Sax|Yamaha|032252|Behrad Poorakbar|11|Hannah Yen (Gr. 9)
SAX-T10|Tenor Sax|Yamaha YTS-480|R32965|Sophie Potapov|12|Talia Gharajeh (9, Jazz)
SAX-T11|Tenor Sax|Yamaha|201281|Ian Tang|9|
SAX-T12|Tenor Sax|Yamaha|047692A|Allegra Cristanti|11|Anya Liu (Gr. 10)
SAX-B01|Bari Sax|Yamaha|017522|Evan Robichaud|10|Lucas Jiang (Gr. 9)
SAX-B02|Bari Sax|Jupiter||Jason Huang|9|Kaitlyn Liu (9, Jazz)
SAX-B03|Bari Sax|Yamaha|038935|Brayden Pan|12|Samuel Pan (Gr. 9)
SAX-B04|Bari Sax|Bari Pro||Sarah Feng|10|
SAX-B05|Bari Sax|Yamaha (2024)|012024|Elias Mio|10|
SAX-B06|Bari Sax|Yamaha (2024)||Sathane Vivekananthan|11|Emily Cheng (10, Jazz)
TRP-01|Trumpet|Yamaha|630809|Olivia Wang|10|Kareen Soubra (9, Jazz)
TRP-02|Trumpet|Yamaha 3335 (2023)|175396|Soheil Rajaei Marbini|9|Deniise Ma (9, Jazz)
TRP-03|Trumpet|Yamaha 1335|133163|Lucian Parmigiano|9|
TRP-04|Trumpet|Bach|396125|Ilanah Arenzonova|10|Kaitlyn Ma (Gr. 9)
TRP-05|Trumpet|Yamaha 3335 (2023)|297774|Anthony Mavros|12|Zachary Guo (9, Jazz)
TRP-06|Trumpet|Yamaha 3335 (2023)|294322|Jayden Fu|10|
TRP-07|Trumpet|Yamaha 3335 (2023)|301754|Lorin Watanabe|12|Wayne Wang (9, Jazz)
TRP-08|Trumpet|Yamaha Japan|114206|Mark Kosynskyy|9|
TRP-09|Trumpet|Yamaha|232013A|Joshua Ng|9|
TRP-10|Trumpet|Yamaha 3335 (2023)|301770|Diana Kirazim|10|Austin Xu (Gr. 9)
TRP-11|Trumpet|Besson 609|841957|Ava Shirali|10|Tristan Chan (9, Jazz)
TRP-12|Trumpet|Yamaha 3335 (2023)|307293|Hendrix Nastou|10|Ben Hsu (Gr. 9)
TRP-13|Trumpet|Yamaha 3335 (2023)|306594|Tyler Chieu|9|
TRP-14|Trumpet|Yamaha|462876A|Lia Mui|11|Navya Bhardwaj (Gr. 9)
TRP-15|Trumpet|Yamaha 3335 (2024)|316606|Marcus Wu|9|Ben Wong (10)
TRP-16|Trumpet|Yamaha 3335 (2024)|311588|Liana Khorshid|10|
TRP-17|Trumpet|Yamaha 3335 (2024)|313844|Orion Nastou|10|
TRP-18|Trumpet|Yamaha 3335 (2024)|317636|Emanuel Volanis|10|Victoria Branscombe (11, Jazz)
TRP-19|Trumpet|Yamaha 3335 (2023)|241572|Ben Vertlib|12|Jordan Ji (9, Jazz)
TRP-20|Trumpet|Yamaha 3335 (2024)|317084|Diana Hu|11|Sopiha Chen (9, Jazz)
TRP-21|Trumpet|Yamaha 3335 (2024)|316848|June Heping|11|Samantha Zhao (Gr. 9)
TRP-22|Trumpet|Yamaha 3335 (2024)|317077|Olivia Liu|9|
TRP-23|Trumpet|Yamaha (2023 Don.)|142388|Aram Alipour|12|Alissa Petkoff (9, Jazz)
TRP-24|Trumpet|Yamaha Japan YTR232|159853A|Lewis Gunston|10|
TRP-25|Trumpet|Yamaha|918073|Kevin Gao|9|
FHN-01|Flugelhorn|King|918540||||V
HNN-01|French Horn|Besson (Double)|04445|Wayne Wang|9|
HNN-02|French Horn|Holton (Double)|H179 5-593190|Arwen Leong|11|Olivia Wang (10), Serena Lu (9)
HNN-03|French Horn|Besson (Double)|04434|Victoria Branscombe|11|Matthew Branscombe (Gr. 9)
HNN-04|French Horn|Conn (Double)|OP5157061218|Ella Hao|10|Kareem Soubra (Gr. 9)
HNN-05|French Horn|Yamaha YFR567 (2023)|043416|Blake French|10|Elaine Liu (12), Kaitlyn Zhong (9)
HNN-07|French Horn|Yamaha (2023)|054186|Eric Lee|11|Kimberly Tan (10), Mia Patel (9)
HNN-08|French Horn|Jupiter (Double)|62123|Richard Lee|12|Claire Zhou (Gr. 10)
HNN-10|French Horn|HHSS10||Alvin Cao|12|Chelleen Chun (Gr. 10)
HNN-04B|French Horn|HHSS4||Athena Zhuo|10|Jay Fu (Gr. 9)
HNN-13|French Horn|HHSS13||Tyler Jin|11|Fiona Irgengioro-Wu (10), Terry Wang (9)
TRB-01|Trombone|Jupiter (2021)|416936|Arses Amiri|10|
TRB-02|Trombone|Yamaha|249376|Roman Motylev|9|Jacob Pozdrowski (12, Jazz)
TRB-03|Trombone|Yamaha (2023)|810025|Elena Zhang|11|Caleb Sales (Gr. 9)
TRB-04|Trombone|Yamaha (2023)|813595|Rosha Rahnama|9|
TRB-05|Trombone|Yamaha (2022)|798947|Fernando Guan|9|
TRB-06|Trombone|Yamaha (2022)|798781|Emma Tang|11|Quintus Li (Gr. 9)
TRB-07|Trombone|Yamaha (2022)|797589|Maha Latify|11|
TRB-08|Trombone|King|974885|Matthew Branscombe|9|Mikka Oinonen (Gr. 9)
TRB-09|Trombone|Yamaha (2023)|819304|Isabella Tan|10|Calix Tsang (Gr. 9)
TRB-10|Trombone|Yamaha (2023)|809435|Noel Lam|9|
TRB-11|Trombone|Yamaha (2022)|792437|Allison Miller|10|Zachary Guo (Gr. 9)
TRB-12|Trombone|Yamaha (2022)|796405|Kaylee Pan|10|Tyler Karakashian (9)
TRB-13|Trombone|Yamaha (2022)|798467|Aiden Ng|9|
TRB-14|Trombone|Yamaha (2022)|798453|Hazal Ozcan|10|Patrick Fan (Gr. 9)
TRB-15|Trombone|Yamaha (2022)|794887|Linnah Kimm|10|Matteo Allen (Gr. 12)
TRB-16|Trombone|Yamaha (2022)|796417|Titus Wu|12|Elena Zhang (Tbn Quartet)
TRB-17|Trombone|Yamaha (2022)|796488|Harsheen Anand|11|Evelyn Rossouw (9, Jazz)
TRB-18|Trombone|Jupiter (2021)|416709|Sarah Long|9|Logan Gardhouse (9, Jazz)
TRB-19|Trombone|King|783679|Frankie Huang|10|
BTR-01|Bass Trombone|Yamaha|00-2686|Alastrina Ng|9|Jocelyn Zou (9, Jazz)
BTR-RG|Bass Trombone|Yamaha (RGSS1)||Matteo Allen|11|Jazz Ensemble
VTR-01|Valve Trombone|Mercedes II|684749|Bowen Pan|9|Jazz Ensemble
EUP-01|Euphonium|Yamaha 4-Valve Reg.|00-3007|Juli He|12|Arshan Rajabian (10), Peter Sun (9)
EUP-02|Euphonium|Yamaha 3-Valve Reg.|000-808|Lucas Luo|9|Kavi Bhanugopan (Gr. 10)
EUP-04|Euphonium|Yamaha YEP321 (2023)|474420|Jayden Tan|10|Iliya Masjedi (9), Kathleen Chen (9)
EUP-05|Euphonium|Yamaha YEP321 (2026)|804253|Iliya Emadi|10|Kaitlyn Wu (9), Leo Lu (9)
EUP-06|Euphonium|Besson Large Shank 3V|762-825617|Mustansir Bandukwala|10|Ava Mooy (Gr. 9)
EUP-07|Euphonium|Yamaha YEP321 (2026)|807329|Jacob Pozdrowski|11|Adrian Pau (Gr. 9)
EUP-08|Euphonium|Yamaha (2022) 4V Reg|586113|Queena Lin|11|Logan Lam (9), Colin Xia (9)
EUP-09|Euphonium|Yamaha Large Shank|115730|Silas Wu|11|Yuki Li (10), Bowen Pan (9)
EUP-10|Euphonium|B&H Large Shank 3V|662-683322|Austen Wu|9|Greyson Liu (Gr. 9)
EUP-11|Euphonium|B&H Large Shank 3V|662-684373|David Hou|9|Sophie Pavlova (Gr. 11)
TUB-01|Tuba|Yamaha 4-Valve (2024)||Ryan Zheng|12|Aidan Sull (9), Dennis Xu (9)
TUB-02|Tuba|Yamaha 4-Valve (2022)||Aaron Huang|12|Kandula Abeysignhe (10), Jocelyn Zou (9)
TUB-03|Tuba|Yamaha 4-Valve||Nancy Qui|11|Felix Miao (Gr. 9)
TUB-04|Tuba|Yamaha 4-Valve (2023)|615219|Dhilan Aggarwal|9|Tiffany Phuong (11), Emily Cheng (10), Logan Gardhouse (9)
TUB-05|Tuba|Yamaha 3/4||Matteo Henriques|9|Guliano De Angelis Geraldo (10)
TUB-07|Tuba|Yamaha 4-Valve (2023)|604104|Roy Chen|10|Bryan Chen (11), Shiheng Liu (9)
TUB-EC|Tuba|Emily Carr 02||Tianxin Liang|11|Hayden Kapoor (9), Mustansir Bandukwala (10)
TUB-M1|Tuba|Maple 01||Kevin He|9|Martin Chen (Gr. 11)
TUB-M2|Tuba|Maple 02||Aaron Dong|9|
TUB-SDSS|Tuba|SDSS Borrowed|||||V
`;

// Bump this text whenever the ORIGINAL list above changes. The first time the website talks to a script with a new
// DATA_VERSION, syncAssignments_() runs once: it fills in every instrument's assigned student, grade and
// "also used by" from the teacher's sheet (the ORIGINAL list). Signed out / in storage / repair statuses are NOT changed.
const DATA_VERSION = '2026-10-16-c';

// Runs at the start of every load of the website (list) and every save.
//  1. Once per DATA_VERSION: syncAssignments_() puts EVERY assigned student, grade and shared-with name back from the list above.
//  2. Every time after that: fillMissingAssignments_() puts back any assigned student, grade or shared-with name that is blank
//     in the Sheet. It never touches status or "signed out by", and it never overwrites a name that is already there,
//     so changes an admin makes with the Edit form are kept.
// Instruments whose tag number was wrong. Done ONCE per RENAMES_V directly in the Sheet row, so the
// instrument keeps its student names, status and "signed out by" (it is only the Tag that changes).
const RENAMES_V = '2026-11-13-a';
const RENAMES = { 'HNN-12': 'HNN-13' };   // French Horn HHSS13 is inventory number 13

function runRenames_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('RENAMES_V') === RENAMES_V) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (props.getProperty('RENAMES_V') !== RENAMES_V) {
      const s = sheet_('Instruments', HEAD);
      const last = s.getLastRow();
      if (last >= 2) {
        const rng = s.getRange(2, 1, last - 1, 1);
        const ids = rng.getValues();
        const have = {};
        ids.forEach(function (r) { have[String(r[0])] = true; });
        let changed = false;
        Object.keys(RENAMES).forEach(function (oldId) {
          const newId = RENAMES[oldId];
          if (!have[oldId] || have[newId]) return;
          for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === oldId) { ids[i][0] = newId; changed = true; }
        });
        if (changed) { rng.setNumberFormat('@').setValues(ids); bump_(); }
      }
      props.setProperty('RENAMES_V', RENAMES_V);
    }
  } finally { lock.releaseLock(); }
}

function fixFormats_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FMT_V') === '1') return;
  const s = sheet_('Instruments', HEAD);
  s.getRange(1, 9, Math.max(s.getMaxRows(), 2), 1).setNumberFormat('@');
  const l = sheet_('Log', LOG_HEAD);
  l.getRange(1, 5, Math.max(l.getMaxRows(), 2), 1).setNumberFormat('@');
  props.setProperty('FMT_V', '1');
}

function ensureData_(skipFill) {
  runRenames_();
  fixFormats_();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DATA_VERSION') !== DATA_VERSION) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (props.getProperty('DATA_VERSION') !== DATA_VERSION) {
        syncAssignments_();
        props.setProperty('DATA_VERSION', DATA_VERSION);
      }
    } finally { lock.releaseLock(); }
  }
  if (!skipFill) fillMissingAssignments_();
}

function fillMissingAssignments_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('filled')) return;          // already checked in the last 15 seconds
  cache.put('filled', '1', 15);
  const s = sheet_('Instruments', HEAD);
  const last = s.getLastRow();
  if (last < 2) return;
  const rng = s.getRange(2, 1, last - 1, 7);
  const vals = rng.getValues();
  const byId = {};
  vals.forEach(function (r, i) { byId[String(r[0])] = i; });
  let changed = 0;
  original_().forEach(function (o) {
    const i = byId[o.id];
    if (i === undefined) return;
    const r = vals[i];
    if (String(r[4]).trim() === '' && o.primary !== '') { r[4] = o.primary; r[5] = o.grade; changed++; }
    else if (String(r[5]).trim() === '' && o.primary !== '' && String(r[4]).trim() === o.primary) { r[5] = o.grade; changed++; }
    if (String(r[6]).trim() === '' && o.secondary !== '') { r[6] = o.secondary; changed++; }
  });
  if (changed) { rng.setNumberFormat('@').setValues(vals); bump_(); }
}

function original_() {
  const CODE = { V: 'Available', R: 'Out for Repair' };
  return ORIGINAL.trim().split('\n').map(function (l) {
    const p = l.split('|');
    return { id: p[0], section: p[1], make: p[2], serial: p[3] || '', primary: p[4] || '', grade: p[5] || '',
             secondary: p[6] || '', status: CODE[(p[7] || '').trim()] || 'Assigned' };
  });
}

// Run this ONCE from the Apps Script editor (pick resetInstruments, click Run) to put the Instruments tab
// back to the original list. It overwrites the current list. The Log tab is not touched.
function resetInstruments() {
  const s = sheet_('Instruments', HEAD);
  const items = original_();
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getMaxColumns()).clearContent();
  s.getRange(2, 1, items.length, 9).setNumberFormat('@').setValues(items.map(row_));
  bump_();
  Logger.log('Restored ' + items.length + ' instruments from the original list.');
}

function syncAssignments_() {
  const s = sheet_('Instruments', HEAD);
  const orig = original_();
  const last = s.getLastRow();
  const rng = last > 1 ? s.getRange(2, 1, last - 1, 9) : null;
  const vals = rng ? rng.getValues() : [];
  const byId = {};
  vals.forEach(function (r, i) { byId[String(r[0])] = i; });
  const add = [];
  orig.forEach(function (o) {
    const i = byId[o.id];
    if (i === undefined) { add.push(row_(o)); return; }  // instrument missing from the tab: add it back
    const r = vals[i];
    r[4] = o.primary; r[5] = o.grade; r[6] = o.secondary; // names from the teacher's sheet; status (r[7]) stays
  });
  if (rng) rng.setNumberFormat('@').setValues(vals);
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, 9).setNumberFormat('@').setValues(add);
  bump_();
  Logger.log('Synced assigned students for ' + orig.length + ' instruments.');
}


// Manual helper: in the Apps Script editor pick restoreAllNames and press Run. Puts every assigned student, grade and
// "also used by" back from the teacher's list. Signed out / in storage / repair statuses and "signed out by" are not touched.
function restoreAllNames() { syncAssignments_(); }
