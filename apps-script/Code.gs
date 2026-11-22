/**
 * IVORY ENQUIRIES - back end
 * Goes into the Apps Script project attached to the "Ivory Enquiry Book" sheet
 * (Google Sheet > Extensions > Apps Script). Paste this whole file into Code.gs.
 */

const SHEET_ID = '19STFjeiGbPxJXZzn2uJ-Rf0SKMII07vOPnF2ZoIphlo';

const CFG = {
  PIN: '480048',        // 1) CHANGE THIS to a 6-digit number. Both phones type it once.
  PARTNER_EMAIL: 'anilaelsa@gmail.com',       // 2) Partner's Google email, so reminders also show on that phone.
  REMINDER_HOUR: 10,       // reminder time on a 24h clock (10 = 10:00 AM)
  REMIND_EVERY_DAYS: 2,
  ENQ: 'Enquiries',
  CLI: 'Clients',
  C_STATUS: 6,             // column F
  C_SINCE: 18,             // column R (new): date the status became Costing
  C_LAST: 19               // column S (new): date the last reminder was made
};

const STATUSES = ['Enquiry', 'Costing', 'Blocked', 'Confirmed', 'Lost'];

// Staff checklist sheets (made the day before a wedding). Anyone with the link can tick items, so staff need no sign-in.
// Optional: list staff Google emails here to also give them named edit access.
const STAFF = { FOLDER: 'Ivory Staff Checklists', EMAILS: [], ANYONE_WITH_LINK: true };

// Outside vendors: tracked ONLY in this app (tick = booked, plus name/details). Saved in a "Vendors" tab of the Enquiry Book.
const VENDOR_NAMES = ['Garland and Bouquet', 'Wedding cake', 'Confetti', 'Cold Spark', 'Dry Ice', 'Ambiance lighting',
  'Parking Security', 'Live Band', 'Wedding Car', 'Generator', 'Pagoda', 'Table & Chair covers', 'Anchor / MC', 'LED Wall', 'Tall Men Security'];
const VENDOR_TAB = 'Vendors';
const QUOTE_TAB = 'Quotes';

/* ---------- web app ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ivory Enquiries')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- helpers ---------- */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function iso_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function parse_(iso) { return Utilities.parseDate(iso, tz_(), 'yyyy-MM-dd'); }

function check_(pin) {
  if (CFG.PIN === 'CHANGE-ME') throw new Error('Set a PIN in Code.gs first');
  if (String(pin) !== String(CFG.PIN)) throw new Error('Wrong PIN');
}

function lastRowWith_(sh, col) {
  const v = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function nextId_(ss) {
  let max = 0;
  [CFG.CLI, CFG.ENQ].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    const v = sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues();
    v.forEach(function (r) {
      const m = /^IV-(\d+)$/.exec(String(r[0]).trim());
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  });
  return 'IV-' + String(max + 1).padStart(3, '0');
}

// Copies the look, dropdowns and formulas of one row to another; typed values are not copied.
function cloneRow_(sh, srcRow, dstRow, ncols) {
  const src = sh.getRange(srcRow, 1, 1, ncols);
  const dst = sh.getRange(dstRow, 1, 1, ncols);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  const f = src.getFormulasR1C1()[0];
  f.forEach(function (x, i) {
    if (x) dst.getCell(1, i + 1).setFormulaR1C1(x);
  });
}

/* ---------- calls from the phone screen ---------- */

function api_list(pin) {
  check_(pin);
  const ss = ss_();
  const sh = ss.getSheetByName(CFG.ENQ);
  const last = lastRowWith_(sh, 1);
  const items = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, CFG.C_LAST).getValues();
    vals.forEach(function (r, i) {
      const id = String(r[0]).trim();
      if (!id) return;
      const status = String(r[5]).trim();
      if (status === 'Lost') return;            // Lost enquiries drop off the list
      items.push({
        row: i + 2,
        id: id,
        name: String(r[1]),
        date: iso_(r[2]),
        venue: String(r[4]),
        status: status,
        since: iso_(r[CFG.C_SINCE - 1])
      });
    });
  }
  items.sort(function (a, b) { return (a.date || '0000').localeCompare(b.date || '0000'); });

  const cli = ss.getSheetByName(CFG.CLI);
  const cl = lastRowWith_(cli, 1);
  const seen = {};
  const names = [];
  if (cl > 1) {
    cli.getRange(2, 2, cl - 1, 1).getValues().forEach(function (r) {
      const n = String(r[0]).trim();
      if (n && !seen[n.toLowerCase()]) { seen[n.toLowerCase()] = 1; names.push(n); }
    });
  }
  return { items: items, names: names, today: today_() };
}

function api_add(pin, p) {
  check_(pin);
  const name = String(p.name || '').trim();
  const venue = String(p.venue || '').trim();
  const date = String(p.date || '').trim();
  const status = STATUSES.indexOf(p.status) >= 0 ? p.status : 'Enquiry';
  if (!name || !date) throw new Error('Client name and event date are needed');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Event date is not valid');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const cli = ss.getSheetByName(CFG.CLI);
    const enq = ss.getSheetByName(CFG.ENQ);
    const cLast = lastRowWith_(cli, 1);
    const eLast = lastRowWith_(enq, 1);

    // A returning client keeps the same ID; a new name gets the next one.
    let id = '';
    if (cLast > 1) {
      cli.getRange(2, 1, cLast - 1, 2).getValues().forEach(function (r) {
        if (!id && String(r[1]).trim().toLowerCase() === name.toLowerCase()) id = String(r[0]).trim();
      });
    }

    if (id && eLast > 1) {
      enq.getRange(2, 1, eLast - 1, 6).getValues().forEach(function (r) {
        if (String(r[0]).trim() === id && iso_(r[2]) === date && String(r[5]).trim() !== 'Lost') {
          throw new Error(name + ' is already listed for that date');
        }
      });
    }

    if (!id) {
      id = nextId_(ss);
      const cr = cLast + 1;
      cloneRow_(cli, cLast, cr, 6);
      cli.getRange(cr, 1).setValue(id);
      cli.getRange(cr, 2).setValue(name);
    }

    const r = eLast + 1;
    cloneRow_(enq, eLast, r, 17);
    enq.getRange(r, 1).setValue(id);
    enq.getRange(r, 3).setValue(parse_(date));
    enq.getRange(r, 5).setValue(venue);
    enq.getRange(r, CFG.C_STATUS).setValue(status);
    if (status === 'Costing') {
      enq.getRange(r, CFG.C_SINCE).setValue(parse_(today_())).setNumberFormat('d mmm yyyy');
    }
    SpreadsheetApp.flush();
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function api_status(pin, p) {
  check_(pin);
  if (STATUSES.indexOf(p.status) < 0) throw new Error('Unknown status');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(CFG.ENQ);
    const r = Number(p.row);
    const row = sh.getRange(r, 1, 1, 6).getValues()[0];
    if (String(row[0]).trim() !== p.id || iso_(row[2]) !== p.date) {
      throw new Error('The sheet has changed. Refresh and try again.');
    }
    const old = String(row[5]).trim();
    sh.getRange(r, CFG.C_STATUS).setValue(p.status);
    if (p.status === 'Costing' && old !== 'Costing') {
      sh.getRange(r, CFG.C_SINCE).setValue(parse_(today_())).setNumberFormat('d mmm yyyy');
      sh.getRange(r, CFG.C_LAST).clearContent();
    } else if (p.status !== 'Costing') {
      sh.getRange(r, CFG.C_SINCE, 1, 2).clearContent();
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- quotes sent + outside vendors (opened from the detail sheet) ---------- */

const MONTHS_ = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

// A date cell may hold a real date, "2026-10-22" or "28 Jul 2026". Returns yyyy-MM-dd or ''.
function isoAny_(v) {
  if (v instanceof Date) return iso_(v);
  const t = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  const m = /^(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})$/.exec(t);
  if (!m || !MONTHS_[m[2].toLowerCase()]) return '';
  return m[3] + '-' + String(MONTHS_[m[2].toLowerCase()]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
}

// Same safety check api_status uses: the row must still be the same client and date.
function checkRow_(sh, p) {
  const row = sh.getRange(Number(p.row), 1, 1, 6).getValues()[0];
  if (String(row[0]).trim() !== String(p.id) || iso_(row[2]) !== String(p.date)) {
    throw new Error('The sheet has changed. Refresh and try again.');
  }
  return row;
}

function vendorTab_(ss) {
  const HEAD = ['Key', 'Client', 'Event date', 'Vendor', 'Details', 'Booked', 'Updated', 'Supplier'];
  let sh = ss.getSheetByName(VENDOR_TAB);
  if (!sh) {
    sh = ss.insertSheet(VENDOR_TAB);
    sh.getRange(1, 1, 1, 8).setValues([HEAD]);
    sh.setFrozenRows(1);
    sh.getRange('A:E').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  } else {
    // tab made by the earlier version: column E becomes "Details" and a new "Supplier" column H is added
    const h = sh.getRange(1, 1, 1, 8).getValues()[0];
    if (String(h[4]) !== 'Details' || String(h[7]) !== 'Supplier') {
      sh.getRange(1, 5).setValue('Details');
      sh.getRange(1, 8).setValue('Supplier');
      sh.getRange('H:H').setNumberFormat('@');
    }
  }
  return sh;
}

function api_detail(pin, p) {
  check_(pin);
  const ss = ss_();
  checkRow_(ss.getSheetByName(CFG.ENQ), p);

  // quotes recorded in the Quotes tab (the Google Sites quotes), matched on client ID + event date
  const sites = [];
  const qs = ss.getSheetByName(QUOTE_TAB);
  const ql = qs ? lastRowWith_(qs, 1) : 1;
  if (ql > 1) {
    qs.getRange(2, 1, ql - 1, 9).getValues().forEach(function (r) {
      if (String(r[1]).trim() !== String(p.id) || isoAny_(r[3]) !== String(p.date)) return;
      const sent = isoAny_(r[4]);
      sites.push({
        no: sent ? sent.slice(2, 4) + sent.slice(5, 7) + sent.slice(8, 10) : '',   // quote no = quote date as yymmdd
        sent: sent,
        url: String(r[6]).trim()
      });
    });
  }

  // saved outside-vendor ticks and details for this work
  const vendors = {}, suppliers = [], seenSup = {};
  const vs = ss.getSheetByName(VENDOR_TAB);
  const vl = vs ? lastRowWith_(vs, 1) : 1;
  const key = p.id + '|' + p.date;
  if (vl > 1) {
    vs.getRange(2, 1, vl - 1, 8).getValues().forEach(function (r) {
      const sup = String(r[7] || '').trim();
      if (sup && !seenSup[sup.toLowerCase()]) { seenSup[sup.toLowerCase()] = 1; suppliers.push(sup); }
      if (String(r[0]) !== key) return;
      vendors[String(r[3])] = { supplier: sup, details: String(r[4] || ''), booked: r[5] === true || String(r[5]).toLowerCase() === 'true' };
    });
  }
  suppliers.sort(function (a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
  return { sites: sites, vendors: vendors, vendorList: VENDOR_NAMES, suppliers: suppliers };
}

function api_vendor(pin, p) {
  check_(pin);
  // The 12 standard vendors plus any other outside-vendor item listed in the client's quote (e.g. "Pandhal 60ft x 30ft").
  const vendor = String(p.vendor || '').trim();
  if (!vendor || vendor.length > 150) throw new Error('Unknown vendor');
  const f = p.fields || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);            // two people saving at the same moment simply wait their turn
  try {
    const ss = ss_();
    const row = checkRow_(ss.getSheetByName(CFG.ENQ), p);
    const sh = vendorTab_(ss);
    const key = p.id + '|' + p.date;
    const last = lastRowWith_(sh, 1);
    let r = -1;
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (v, i) {
        if (r < 0 && String(v[0]) === key && String(v[3]) === vendor) r = i + 2;
      });
    }
    if (p.remove === true) {          // the latest quote no longer lists this item: delete its row
      if (r > 0) sh.deleteRow(r);
      SpreadsheetApp.flush();
      return { ok: true, removed: r > 0 };
    }
    if (r < 0) {
      r = last + 1;
      sh.getRange(r, 1, 1, 5).setNumberFormat('@').setValues([[key, String(row[1]), String(p.date), vendor, '']]);
      sh.getRange(r, 6).insertCheckboxes().setValue(false);
    }
    // only the field that changed is written, so one person's tick never wipes the other's note
    if ('details' in f) sh.getRange(r, 5).setNumberFormat('@').setValue(String(f.details));
    if ('supplier' in f) sh.getRange(r, 8).setNumberFormat('@').setValue(String(f.supplier).trim());
    if ('booked' in f) sh.getRange(r, 6).insertCheckboxes().setValue(f.booked === true);
    sh.getRange(r, 7).setValue(new Date()).setNumberFormat('d mmm yyyy h:mm AM/PM');
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Everything booked, for the Vendor list screen. Works that were marked Lost are left out.
function api_vendors(pin) {
  check_(pin);
  const ss = ss_();
  const lost = {}, venues = {};
  const es = ss.getSheetByName(CFG.ENQ), el = lastRowWith_(es, 1);
  if (el > 1) {
    es.getRange(2, 1, el - 1, 6).getValues().forEach(function (r) {
      const k = String(r[0]).trim() + '|' + iso_(r[2]);
      venues[k] = String(r[4] || '').trim();
      if (String(r[5]).trim() === 'Lost') lost[k] = 1;
    });
  }
  const rows = [];
  const vs = ss.getSheetByName(VENDOR_TAB), vl = vs ? lastRowWith_(vs, 1) : 1;
  if (vl > 1) {
    vs.getRange(2, 1, vl - 1, 8).getValues().forEach(function (r) {
      const key = String(r[0]);
      if (!key || lost[key]) return;
      const parts = key.split('|');
      rows.push({
        id: parts[0], date: isoAny_(r[2]) || parts[1] || '', client: String(r[1]), venue: venues[key] || '', vendor: String(r[3]),
        details: String(r[4] || ''), supplier: String(r[7] || '').trim(),
        booked: r[5] === true || String(r[5]).toLowerCase() === 'true'
      });
    });
  }
  return { rows: rows };
}

/* ---------- payment details: published to <Client>/payment.html in the Ivory-Packages repo ---------- */
// The GitHub key (fine-grained, Contents read/write on Ivory-Packages) is kept in Script Properties as GH_KEY,
// so it is typed once in the app and works on both phones.
const PAY_REPO = 'ivoryeventz-ops/Ivory-Packages';
const PAY_BRANCH = 'main';

function api_pay_status(pin) {
  check_(pin);
  return { hasKey: !!PropertiesService.getScriptProperties().getProperty('GH_KEY') };
}

function api_pay_publish(pin, p) {
  check_(pin);
  const props = PropertiesService.getScriptProperties();
  if (p.key) props.setProperty('GH_KEY', String(p.key).trim());
  const key = props.getProperty('GH_KEY');
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  if (!p.html || String(p.html).indexOf('Payment Details') < 0) throw new Error('Nothing to publish');
  const url = 'https://api.github.com/repos/' + PAY_REPO + '/contents/' + encodeURIComponent(folder) + '/payment.html';
  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const bad = function (code) {
    if (code === 401 || code === 403) { props.deleteProperty('GH_KEY'); return new Error('GitHub key not accepted - paste it again'); }
    return new Error('GitHub answered ' + code);
  };
  const g = UrlFetchApp.fetch(url + '?ref=' + PAY_BRANCH, { headers: hd, muteHttpExceptions: true });
  let sha = null;
  if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;
  else if (g.getResponseCode() !== 404) throw bad(g.getResponseCode());
  const body = {
    message: String(p.message || 'Payment details').slice(0, 200),
    content: Utilities.base64Encode(Utilities.newBlob(String(p.html)).getBytes()),
    branch: PAY_BRANCH
  };
  if (sha) body.sha = sha;
  const r = UrlFetchApp.fetch(url, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200 && r.getResponseCode() !== 201) throw bad(r.getResponseCode());
  return { ok: true };
}

/* ---------- edit as new quote: published to <Client>/<YYMMDD>.html in the Ivory-Packages repo ---------- */
// Uses the same GH_KEY as the payment pages. A quote that already exists is only overwritten when the app says replace:true
// (same-day re-issue, ticked by the person); a new stage photo is written first so the page never points at a missing image.
function api_quote_publish(pin, p) {
  check_(pin);
  const props = PropertiesService.getScriptProperties();
  if (p.key) props.setProperty('GH_KEY', String(p.key).trim());
  const key = props.getProperty('GH_KEY');
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  const no = String(p.no || '');
  if (!/^\d{6}$/.test(no)) throw new Error('Bad quote number');
  if (!p.html || String(p.html).indexOf('Decor Quote') < 0) throw new Error('Nothing to publish');
  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const base = 'https://api.github.com/repos/' + PAY_REPO + '/contents/' + encodeURIComponent(folder) + '/';
  const bad = function (code) {
    if (code === 401 || code === 403) { props.deleteProperty('GH_KEY'); return new Error('GitHub key not accepted - paste it again'); }
    return new Error('GitHub answered ' + code);
  };
  const put = function (file, b64, message, overwrite) {
    const url = base + encodeURIComponent(file);
    const g = UrlFetchApp.fetch(url + '?ref=' + PAY_BRANCH, { headers: hd, muteHttpExceptions: true });
    let sha = null;
    if (g.getResponseCode() === 200) {
      if (!overwrite) throw new Error(file + ' already exists - tick replace to overwrite it');
      sha = JSON.parse(g.getContentText()).sha;
    } else if (g.getResponseCode() !== 404) throw bad(g.getResponseCode());
    const body = { message: String(message || 'Quote').slice(0, 200), content: b64, branch: PAY_BRANCH };
    if (sha) body.sha = sha;
    const r = UrlFetchApp.fetch(url, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    if (r.getResponseCode() !== 200 && r.getResponseCode() !== 201) throw bad(r.getResponseCode());
  };
  const msg = String(p.message || ('Quote ' + folder + '/' + no));
  if (p.img) {
    const name = String(p.img.name || '');
    const b64 = String(p.img.b64 || '');
    if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,60}\.(jpg|jpeg|png)$/i.test(name) || !b64 || b64.length > 9000000) throw new Error('Bad stage photo');
    put(name, b64, msg + ' - stage photo', true);
  }
  put(no + '.html', Utilities.base64Encode(Utilities.newBlob(String(p.html)).getBytes()), msg, p.replace === true);
  return { ok: true };
}

/* ---------- wedding timeline: published to <Client>/timeline.html in the Ivory-Wedding-Timeline repo ---------- */
// Uses its own key (TL_KEY) when one was pasted, otherwise the same key as the payment pages (GH_KEY).
// A fine-grained key can cover both repositories: add Ivory-Wedding-Timeline with Contents read/write.
const TL_REPO = 'ivoryeventz-ops/Ivory-Wedding-Timeline';
const TL_BRANCH = 'main';

function api_tl_status(pin) {
  check_(pin);
  const props = PropertiesService.getScriptProperties();
  return { hasKey: !!(props.getProperty('TL_KEY') || props.getProperty('GH_KEY')) };
}

function api_tl_publish(pin, p) {
  check_(pin);
  const props = PropertiesService.getScriptProperties();
  if (p.key) props.setProperty('TL_KEY', String(p.key).trim());
  const key = props.getProperty('TL_KEY') || props.getProperty('GH_KEY');
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  if (!p.html || String(p.html).indexOf('Wedding Day Guide') < 0) throw new Error('Nothing to publish');
  const url = 'https://api.github.com/repos/' + TL_REPO + '/contents/' + encodeURIComponent(folder) + '/timeline.html';
  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const noReach = function (code) {
    if (code === 401 && props.getProperty('TL_KEY')) props.deleteProperty('TL_KEY');   // a rejected pasted key is asked for again
    return new Error('GitHub key does not reach Ivory-Wedding-Timeline (GitHub answered ' + code + ')');
  };
  const g = UrlFetchApp.fetch(url + '?ref=' + TL_BRANCH, { headers: hd, muteHttpExceptions: true });
  let sha = null;
  if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;
  else if (g.getResponseCode() === 401 || g.getResponseCode() === 403) throw noReach(g.getResponseCode());
  else if (g.getResponseCode() !== 404) throw new Error('GitHub answered ' + g.getResponseCode());
  const body = {
    message: String(p.message || 'Wedding timeline').slice(0, 200),
    content: Utilities.base64Encode(Utilities.newBlob(String(p.html)).getBytes()),
    branch: TL_BRANCH
  };
  if (sha) body.sha = sha;
  const r = UrlFetchApp.fetch(url, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  const c = r.getResponseCode();
  if (c === 200 || c === 201) return { ok: true };
  if (c === 401 || c === 403 || c === 404) throw noReach(c);
  throw new Error('GitHub answered ' + c);
}

/* ---------- client site tabs: <Client>/data.json in the Ivory-Events repo (New A site) ---------- */
// The client's one link (ivory-events.pages.dev/New A/<Client>/) reads data.json: quote tabs (newest first), then Timeline, then Payment.
// The app adds ONE item per call, and only after Navin has confirmed that item on screen. Uses SITE_KEY, else TL_KEY, else GH_KEY.
const SITE_REPO = 'ivoryeventz-ops/Ivory-Events';
const SITE_BRANCH = 'main';
const SITE_DIR = 'New A';
// Event types: what the client link shows under Packages and Service Checklist (pages live in Ivory-Events > New A > parts/<folder>).
const SITE_TYPES = { christian: 'Wedding - Christian', pentecost: 'Wedding - Pentecost', engagement: 'Engagement', hindu: 'Reception - Hindu' };

function siteKey_(props) { return props.getProperty('SITE_KEY') || props.getProperty('TL_KEY') || props.getProperty('GH_KEY') || ''; }

function api_site_status(pin) {
  check_(pin);
  return { hasKey: !!siteKey_(PropertiesService.getScriptProperties()) };
}

function api_site_tab(pin, p) {
  check_(pin);
  if (p.confirmed !== true) throw new Error('Not confirmed');            // nothing goes on a client link without the tap on the confirm button
  const props = PropertiesService.getScriptProperties();
  if (p.key) props.setProperty('SITE_KEY', String(p.key).trim());
  const key = siteKey_(props);
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  const kind = String(p.kind || '');
  if (['quote', 'timeline', 'payment'].indexOf(kind) < 0) throw new Error('Unknown tab');
  const tabUrl = String(p.url || '');
  if (!/^https:\/\/[a-z0-9.-]+\.pages\.dev\/[^\s"<>]+$/i.test(tabUrl)) throw new Error('Bad page link');
  if (kind === 'quote' && (!/^\d{6}$/.test(String(p.no || '')) || !/^\d{4}-\d{2}-\d{2}$/.test(String(p.created || '')))) throw new Error('Bad quote number or date');

  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const api = 'https://api.github.com/repos/' + SITE_REPO + '/contents/';
  const path = function (f) { return encodeURIComponent(SITE_DIR) + '/' + (f ? encodeURIComponent(folder) + '/' + encodeURIComponent(f) : encodeURIComponent('client-shell.html')); };
  const noReach = function (code) {
    if (code === 401) props.deleteProperty('SITE_KEY');
    return new Error('GitHub key does not reach Ivory-Events (GitHub answered ' + code + ')');
  };
  const get = function (pth) {
    const r = UrlFetchApp.fetch(api + pth + '?ref=' + SITE_BRANCH, { headers: hd, muteHttpExceptions: true });
    const c = r.getResponseCode();
    if (c === 200) return JSON.parse(r.getContentText());
    if (c === 404) return null;
    if (c === 401 || c === 403) throw noReach(c);
    throw new Error('GitHub answered ' + c);
  };
  const put = function (pth, text, sha, msg) {
    const body = { message: msg, content: Utilities.base64Encode(Utilities.newBlob(text).getBytes()), branch: SITE_BRANCH };
    if (sha) body.sha = sha;
    const r = UrlFetchApp.fetch(api + pth, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    const c = r.getResponseCode();
    if (c === 200 || c === 201) return;
    if (c === 401 || c === 403 || c === 404) throw noReach(c);
    throw new Error('GitHub answered ' + c);
  };
  const esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let cur = get(path('data.json'));
    let data, sha = null;
    if (cur) {
      sha = cur.sha;
      data = JSON.parse(Utilities.newBlob(Utilities.base64Decode(cur.content.replace(/\n/g, ''))).getDataAsString());
    } else {
      // first tab for this client: make the client page from the shell, then its data.json
      const name = String(p.name || '').trim(), venue = String(p.venue || '').trim(), wd = String(p.date || '');
      if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(wd)) throw new Error('Client name and wedding date are needed to start this client link');
      const shell = get(path(''));
      if (!shell) throw new Error('client-shell.html not found in the New A site');
      let html = Utilities.newBlob(Utilities.base64Decode(shell.content.replace(/\n/g, ''))).getDataAsString();
      const longDate = Utilities.formatDate(Utilities.parseDate(wd, tz_(), 'yyyy-MM-dd'), tz_(), 'd MMMM yyyy');
      html = html.split('__NAME__').join(esc(name)).split('__WEDDING_DATE__').join(longDate).split('__VENUE__').join(esc(venue)).split('__FOLDER__').join(folder);
      put(path('index.html'), html, (get(path('index.html')) || {}).sha || null, 'Client page ' + folder);
      data = { folder: folder, name: name, wedding_date: wd, venue: venue, quotes: [] };
      if (SITE_TYPES[p.type]) data.type = String(p.type);
    }
    data.quotes = data.quotes || [];
    if (kind === 'quote') {
      if (data.quotes.some(function (q) { return String(q.id) === String(p.no); })) return { ok: true, already: true };
      data.quotes.push({ id: String(p.no), created: String(p.created), url: tabUrl });
    } else {
      data[kind] = { url: tabUrl, added: today_() };
    }
    put(path('data.json'), JSON.stringify(data, null, 2) + '\n', sha, 'Add ' + kind + ' tab ' + folder);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Takes ONE item (a quote, the timeline or the payment page) off the client's link again, after a confirm in the app.
// Only the tab is removed from data.json; the quote / timeline / payment page itself stays where it is and can be added back.
function api_site_untab(pin, p) {
  check_(pin);
  if (p.confirmed !== true) throw new Error('Not confirmed');
  const props = PropertiesService.getScriptProperties();
  const key = siteKey_(props);
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  const kind = String(p.kind || '');
  if (['quote', 'timeline', 'payment'].indexOf(kind) < 0) throw new Error('Unknown tab');
  if (kind === 'quote' && !/^\d{6}$/.test(String(p.no || ''))) throw new Error('Bad quote number');
  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const url = 'https://api.github.com/repos/' + SITE_REPO + '/contents/' + encodeURIComponent(SITE_DIR) + '/' + encodeURIComponent(folder) + '/data.json';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const g = UrlFetchApp.fetch(url + '?ref=' + SITE_BRANCH, { headers: hd, muteHttpExceptions: true });
    const gc = g.getResponseCode();
    if (gc === 404) return { ok: true, already: true };
    if (gc !== 200) throw new Error('GitHub key does not reach Ivory-Events (GitHub answered ' + gc + ')');
    const cur = JSON.parse(g.getContentText());
    const data = JSON.parse(Utilities.newBlob(Utilities.base64Decode(cur.content.replace(/\n/g, ''))).getDataAsString());
    data.quotes = data.quotes || [];
    let changed = false;
    if (kind === 'quote') {
      const n = data.quotes.length;
      data.quotes = data.quotes.filter(function (q) { return String(q.id) !== String(p.no); });
      changed = data.quotes.length !== n;
    } else if (data[kind]) { delete data[kind]; changed = true; }
    if (!changed) return { ok: true, already: true };
    const body = { message: 'Remove ' + kind + ' tab ' + folder, content: Utilities.base64Encode(Utilities.newBlob(JSON.stringify(data, null, 2) + '\n').getBytes()), sha: cur.sha, branch: SITE_BRANCH };
    const r = UrlFetchApp.fetch(url, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    const rc = r.getResponseCode();
    if (rc !== 200 && rc !== 201) throw new Error('GitHub answered ' + rc);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Makes the client's link as soon as an enquiry is entered (and lets the event type be changed later).
// Creates <Client>/index.html (from client-shell.html) and <Client>/data.json with the event type, which decides the Packages and Checklist pages.
// An existing link keeps its quotes/timeline/payment; only its type is updated.
function api_site_create(pin, p) {
  check_(pin);
  const props = PropertiesService.getScriptProperties();
  if (p.key) props.setProperty('SITE_KEY', String(p.key).trim());
  const key = siteKey_(props);
  if (!key) throw new Error('GitHub key needed');
  const folder = String(p.folder || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,60}$/.test(folder)) throw new Error('Bad client folder name');
  const type = String(p.type || '');
  if (!SITE_TYPES[type]) throw new Error('Choose the event type');
  const hd = { Authorization: 'Bearer ' + key, Accept: 'application/vnd.github+json' };
  const api = 'https://api.github.com/repos/' + SITE_REPO + '/contents/';
  const path = function (f) { return encodeURIComponent(SITE_DIR) + '/' + (f ? encodeURIComponent(folder) + '/' + encodeURIComponent(f) : encodeURIComponent('client-shell.html')); };
  const noReach = function (code) {
    if (code === 401) props.deleteProperty('SITE_KEY');
    return new Error('GitHub key does not reach Ivory-Events (GitHub answered ' + code + ')');
  };
  const get = function (pth) {
    const r = UrlFetchApp.fetch(api + pth + '?ref=' + SITE_BRANCH, { headers: hd, muteHttpExceptions: true });
    const c = r.getResponseCode();
    if (c === 200) return JSON.parse(r.getContentText());
    if (c === 404) return null;
    if (c === 401 || c === 403) throw noReach(c);
    throw new Error('GitHub answered ' + c);
  };
  const put = function (pth, text, sha, msg) {
    const body = { message: msg, content: Utilities.base64Encode(Utilities.newBlob(text).getBytes()), branch: SITE_BRANCH };
    if (sha) body.sha = sha;
    const r = UrlFetchApp.fetch(api + pth, { method: 'put', headers: hd, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
    const c = r.getResponseCode();
    if (c === 200 || c === 201) return;
    if (c === 401 || c === 403 || c === 404) throw noReach(c);
    throw new Error('GitHub answered ' + c);
  };
  const esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const url = 'https://ivory-events.pages.dev/' + encodeURIComponent(SITE_DIR) + '/' + encodeURIComponent(folder) + '/';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cur = get(path('data.json'));
    if (cur) {
      const data = JSON.parse(Utilities.newBlob(Utilities.base64Decode(cur.content.replace(/\n/g, ''))).getDataAsString());
      if (data.type !== type && p.change === true) {
        data.type = type;
        put(path('data.json'), JSON.stringify(data, null, 2) + '\n', cur.sha, 'Event type ' + folder + ': ' + SITE_TYPES[type]);
      }
      return { ok: true, url: url, existed: true, type: type };
    }
    const name = String(p.name || '').trim(), venue = String(p.venue || '').trim(), wd = String(p.date || '');
    if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(wd)) throw new Error('Client name and event date are needed to make the client link');
    const shell = get(path(''));
    if (!shell) throw new Error('client-shell.html not found in the New A site');
    let html = Utilities.newBlob(Utilities.base64Decode(shell.content.replace(/\n/g, ''))).getDataAsString();
    const longDate = Utilities.formatDate(Utilities.parseDate(wd, tz_(), 'yyyy-MM-dd'), tz_(), 'd MMMM yyyy');
    html = html.split('__NAME__').join(esc(name)).split('__WEDDING_DATE__').join(longDate).split('__VENUE__').join(esc(venue)).split('__FOLDER__').join(folder);
    put(path('index.html'), html, (get(path('index.html')) || {}).sha || null, 'Client page ' + folder);
    const data = { folder: folder, name: name, wedding_date: wd, venue: venue, type: type, quotes: [] };
    put(path('data.json'), JSON.stringify(data, null, 2) + '\n', null, 'Client link ' + folder + ' (' + SITE_TYPES[type] + ')');
    return { ok: true, url: url, type: type };
  } finally {
    lock.releaseLock();
  }
}

// File lists of the quotes and timeline repos, read with the saved GitHub key. The phone used to ask GitHub directly with no key,
// and GitHub allows only 60 such requests an hour per network, which shows as "GitHub answered 403".
function api_gh_tree(pin, p) {
  check_(pin);
  const repo = String(p.repo || ''), branch = String(p.branch || 'main');
  if (['ivoryeventz-ops/Ivory-Packages', 'ivoryeventz-ops/Ivory-Wedding-Timeline', 'ivoryeventz-ops/Ivory-Events'].indexOf(repo) < 0) throw new Error('Unknown repo');
  if (!/^[A-Za-z0-9._-]+$/.test(branch)) throw new Error('Bad branch');
  const props = PropertiesService.getScriptProperties();
  const keys = ['GH_KEY', 'TL_KEY', 'SITE_KEY'].map(function (k) { return props.getProperty(k); }).filter(Boolean);
  const names = ['GH_KEY', 'TL_KEY', 'SITE_KEY'].filter(function (k) { return props.getProperty(k); });
  keys.push('');
  const tried = [];
  for (let i = 0; i < keys.length; i++) {
    const hd = { Accept: 'application/vnd.github+json' };
    if (keys[i]) hd.Authorization = 'Bearer ' + keys[i];
    const r = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/git/trees/' + branch + '?recursive=1', { headers: hd, muteHttpExceptions: true });
    const code = r.getResponseCode();
    if (code === 200) return { tree: (JSON.parse(r.getContentText()).tree || []).map(function (x) { return { path: x.path }; }) };
    tried.push((keys[i] ? names[i] : 'no key') + ' ' + code);
  }
  throw new Error('GitHub refused: ' + tried.join(', ') + (keys.length === 1 ? ' (no key saved in the script yet)' : ''));
}

/* ---------- staff checklist: a Google Sheet made the day before the wedding, items from the latest quote ---------- */
const CHK_TAB = 'Checklists';

function chkLogTab_(ss) {
  let sh = ss.getSheetByName(CHK_TAB);
  if (!sh) {
    sh = ss.insertSheet(CHK_TAB);
    sh.getRange(1, 1, 1, 6).setValues([['Key', 'Client', 'Event date', 'Sheet link', 'Created', 'Items']]);
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@');
  }
  return sh;
}

// The last checklist sheet made for this work (shared by both phones, so nobody makes a second one by mistake).
function api_chk_get(pin, p) {
  check_(pin);
  const ss = ss_();
  checkRow_(ss.getSheetByName(CFG.ENQ), p);
  const sh = ss.getSheetByName(CHK_TAB), last = sh ? lastRowWith_(sh, 1) : 1;
  const key = p.id + '|' + p.date;
  let out = null;
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 6).getValues().forEach(function (r) {
      if (String(r[0]) === key) out = { url: String(r[3]), created: r[4] instanceof Date ? iso_(r[4]) : String(r[4]), items: Number(r[5]) || 0 };
    });
  }
  return { last: out };
}

function api_chk_create(pin, p) {
  check_(pin);
  if (p.confirmed !== true) throw new Error('Not confirmed');
  const list = (p.items || []).filter(function (x) { return x && String(x.item || '').trim(); }).slice(0, 300);
  if (!list.length) throw new Error('Nothing to put on the checklist');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const row = checkRow_(ss.getSheetByName(CFG.ENQ), p);
    const clean = function (t, n) { return String(t == null ? '' : t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, n); };
    const client = clean(row[1], 80), dateLabel = Utilities.formatDate(parse_(String(p.date)), tz_(), 'EEE d MMM yyyy');
    const venue = clean(p.venue, 120);
    const fileDate = Utilities.formatDate(parse_(String(p.date)), tz_(), 'd MMM yyyy');
    const book = SpreadsheetApp.create('Ivory Checklist - ' + clean(venue || client, 80) + ' - ' + fileDate);
    const sh = book.getSheets()[0];
    sh.setName('Checklist');
    // Look of the "Wedding Setup Checklist" sheet, simplified for staff: Done | Item | Qty | Details. No sections, no vendor names.
    // Ticking a box turns the row pale green and strikes the item and details through.
    const GREEN = '#0e6b57', LINE = '#d8e0db', FONT = 'Comfortaa';
    const n = list.length;
    sh.getRange(1, 1, Math.max(n + 10, 40), 4).setFontFamily(FONT).setFontSize(11).setVerticalAlignment('middle');
    sh.getRange(1, 1).setValue('Wedding Setup Checklist').setFontSize(16).setFontWeight('bold');
    sh.getRange(2, 1).setValue('Tick the box when a job is done');
    sh.getRange(3, 1).setFormula('=COUNTIF(A6:A,TRUE)&" of "&(COUNTIF(A6:A,TRUE)+COUNTIF(A6:A,FALSE))&" jobs done"').setFontSize(12).setFontWeight('bold');
    sh.getRange(4, 1).setValue(client + '  |  ' + dateLabel + (venue ? '  |  ' + venue : '')).setFontColor('#555555');
    sh.getRange(5, 1, 1, 4).setValues([['Done', 'Item', 'Qty', 'Details']]).setBackground(GREEN).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
    sh.getRange(5, 1).setHorizontalAlignment('center');
    sh.getRange(5, 3).setHorizontalAlignment('center');
    const rows = list.map(function (x) { return [false, clean(x.item, 200), clean(x.qty, 20), clean(x.details, 400)]; });
    const body = sh.getRange(6, 1, n, 4);
    body.setValues(rows).setWrap(true);
    sh.getRange(6, 1, n, 1).insertCheckboxes().setHorizontalAlignment('center');
    sh.getRange(6, 3, n, 1).setNumberFormat('@').setHorizontalAlignment('center');
    sh.getRange(5, 1, n + 1, 4).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
    [64, 380, 54, 280].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setFrozenRows(5);
    const rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A6=TRUE')
      .setBackground('#edf7f2').setFontColor('#727f77').setStrikethrough(true).setRanges([sh.getRange(6, 2, n, 3)]).build();
    sh.setConditionalFormatRules([rule]);
    const itemCount = n;
    SpreadsheetApp.flush();

    const file = DriveApp.getFileById(book.getId());
    let folder;
    const it = DriveApp.getFoldersByName(STAFF.FOLDER);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(STAFF.FOLDER);
    file.moveTo(folder);
    if (STAFF.ANYONE_WITH_LINK) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    (STAFF.EMAILS || []).forEach(function (m) { if (m) file.addEditor(m); });

    const url = book.getUrl();
    const log = chkLogTab_(ss), lr = lastRowWith_(log, 1) + 1;
    log.getRange(lr, 1, 1, 6).setNumberFormat('@').setValues([[p.id + '|' + p.date, client, String(p.date), url, today_(), itemCount]]);
    return { ok: true, url: url, items: itemCount };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- quote flags + per-work action reminders ---------- */

const ACTION_TAB = 'Actions';

function actionsTab_(ss) {
  let sh = ss.getSheetByName(ACTION_TAB);
  if (!sh) {
    sh = ss.insertSheet(ACTION_TAB);
    sh.getRange(1, 1, 1, 5).setValues([['Key', 'Client', 'Task', 'Done', 'Updated']]);
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
  }
  return sh;
}

// For the main list: which works have a Google Sites quote on record (Quotes tab), and which reminder tasks are ticked done.
function api_flags(pin) {
  check_(pin);
  const ss = ss_();
  const quotes = {};
  const qs = ss.getSheetByName(QUOTE_TAB), ql = qs ? lastRowWith_(qs, 1) : 1;
  if (ql > 1) {
    qs.getRange(2, 1, ql - 1, 9).getValues().forEach(function (r) {
      const id = String(r[1]).trim(), d = isoAny_(r[3]);
      if (!id || !d) return;
      const k = id + '|' + d;
      quotes[k] = (quotes[k] || 0) + 1;
    });
  }
  const actions = [];
  const as = ss.getSheetByName(ACTION_TAB), al = as ? lastRowWith_(as, 1) : 1;
  if (al > 1) {
    as.getRange(2, 1, al - 1, 5).getValues().forEach(function (r) {
      if (!r[0]) return;
      actions.push({ key: String(r[0]), task: String(r[2]), done: r[3] === true || String(r[3]).toLowerCase() === 'true' });
    });
  }
  return { quotes: quotes, actions: actions };
}

function api_action(pin, p) {
  check_(pin);
  const task = String(p.task || '').trim();
  if (!task || task.length > 40) throw new Error('Unknown task');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const row = checkRow_(ss.getSheetByName(CFG.ENQ), p);
    const sh = actionsTab_(ss);
    const key = p.id + '|' + p.date;
    const last = lastRowWith_(sh, 1);
    let r = -1;
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (v, i) {
        if (r < 0 && String(v[0]) === key && String(v[2]) === task) r = i + 2;
      });
    }
    if (r < 0) {
      r = last + 1;
      sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[key, String(row[1]), task]]);
    }
    sh.getRange(r, 4).insertCheckboxes().setValue(p.done === true);
    sh.getRange(r, 5).setValue(new Date()).setNumberFormat('d mmm yyyy h:mm AM/PM');
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- reminders ---------- */

// Runs every morning. For each enquiry in Costing it makes a calendar reminder
// on day 2, 4, 6... after the status became Costing, until the status changes.
function sendReminders() {
  const sh = ss_().getSheetByName(CFG.ENQ);
  const last = lastRowWith_(sh, 1);
  if (last < 2) return;
  const todayIso = today_();
  const todayMs = parse_(todayIso).getTime();
  const vals = sh.getRange(2, 1, last - 1, CFG.C_LAST).getValues();
  const due = [];

  vals.forEach(function (r, i) {
    if (String(r[0]).trim() === '' || String(r[5]).trim() !== 'Costing') return;
    const row = i + 2;
    const since = r[CFG.C_SINCE - 1];
    if (!(since instanceof Date)) {
      // Already Costing before the app existed: start counting from today.
      sh.getRange(row, CFG.C_SINCE).setValue(parse_(todayIso)).setNumberFormat('d mmm yyyy');
      return;
    }
    const days = Math.round((todayMs - parse_(iso_(since)).getTime()) / 86400000);
    const lastRem = iso_(r[CFG.C_LAST - 1]);
    if (days > 0 && days % CFG.REMIND_EVERY_DAYS === 0 && lastRem !== todayIso) {
      due.push({ row: row, name: String(r[1]), venue: String(r[4]), date: iso_(r[2]), days: days });
    }
  });
  if (!due.length) return;

  const start = Utilities.parseDate(todayIso + ' ' + String(CFG.REMINDER_HOUR).padStart(2, '0') + ':00', tz_(), 'yyyy-MM-dd HH:mm');
  const end = new Date(start.getTime() + 15 * 60000);
  const link = ScriptApp.getService().getUrl() || '';

  try {
    const cal = CalendarApp.getDefaultCalendar();
    due.forEach(function (d) {
      const opts = { description: 'Costing for ' + d.days + ' days. Follow up with the client.\n' + link };
      if (CFG.PARTNER_EMAIL) { opts.guests = CFG.PARTNER_EMAIL; opts.sendInvites = false; }
      const ev = cal.createEvent('Follow up: ' + d.name + ' - ' + d.date + (d.venue ? ' (' + d.venue + ')' : ''), start, end, opts);
      ev.removeAllReminders();
      ev.addPopupReminder(0);
    });
  } catch (err) {
    // Calendar not available: fall back to one email so the reminder is never lost.
    const to = Session.getEffectiveUser().getEmail() + (CFG.PARTNER_EMAIL ? ',' + CFG.PARTNER_EMAIL : '');
    const lines = due.map(function (d) { return '- ' + d.name + ', ' + d.date + ' ' + d.venue + ' (costing ' + d.days + ' days)'; });
    MailApp.sendEmail(to, 'Ivory: follow up on costing enquiries', lines.join('\n') + '\n\n' + link);
  }

  due.forEach(function (d) {
    sh.getRange(d.row, CFG.C_LAST).setValue(parse_(todayIso)).setNumberFormat('d mmm yyyy');
  });
}

/* ---------- keeps the reminder clock right when status is changed in the sheet itself ---------- */

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.ENQ || e.range.getColumn() !== CFG.C_STATUS || e.range.getRow() < 2) return;
    if (e.range.getNumRows() !== 1) return;
    const r = e.range.getRow();
    if (String(e.value || '').trim() === 'Costing') {
      const tz = e.source.getSpreadsheetTimeZone();
      const day = Utilities.parseDate(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
      sh.getRange(r, CFG.C_SINCE).setValue(day).setNumberFormat('d mmm yyyy');
      sh.getRange(r, CFG.C_LAST).clearContent();
    } else {
      sh.getRange(r, CFG.C_SINCE, 1, 2).clearContent();
    }
  } catch (err) { /* never block an edit */ }
}

/* ---------- run this ONCE from the editor ---------- */

function setupOnce() {
  const sh = ss_().getSheetByName(CFG.ENQ);
  sh.getRange(1, CFG.C_SINCE, 1, 2).setValues([['Costing since', 'Last reminded']]);
  sh.getRange(1, 17).copyTo(sh.getRange(1, CFG.C_SINCE, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(2, CFG.C_SINCE, sh.getMaxRows() - 1, 2).setNumberFormat('d mmm yyyy');

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(6).inTimezone(tz_()).create();

  CalendarApp.getDefaultCalendar(); // makes Google ask for calendar permission now
}
function myFunction() {
  
}

// Quote Builder: one-tap Google Sheet of the ticked items (rows = array of arrays; first row is the heading)
function api_quote_sheet(pin, p) {
  check_(pin);
  const rows = (p.rows || []).slice(0, 400).map(function (r) {
    return (r || []).slice(0, 8).map(function (v) { return v == null ? '' : (typeof v === 'number' ? v : String(v).slice(0, 300)); });
  });
  if (!rows.length) throw new Error('Nothing to put in the sheet');
  const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  rows.forEach(function (r) { while (r.length < width) r.push(''); });
  const book = SpreadsheetApp.create(String(p.title || 'Ivory quote').replace(/[\u0000-\u001F]/g, '').slice(0, 90));
  const sh = book.getSheets()[0];
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, width).setFontWeight('bold');
  sh.autoResizeColumns(1, width);
  return { url: book.getUrl() };
}


// Run this once from the Apps Script editor (select "authorizeAll", press Run, then Review permissions > Allow).
// It asks for Google Drive access (staff checklist sheets) and external web access (GitHub), so the app stops saying "You do not have permission".
function authorizeAll() {
  DriveApp.getRootFolder();
  SpreadsheetApp.openById(SHEET_ID).getName();
  UrlFetchApp.fetch('https://api.github.com/zen', { muteHttpExceptions: true });
  return 'Authorised';
}
