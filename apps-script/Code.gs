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

// Outside vendors: tracked ONLY in this app (tick = booked, plus name/details). Saved in a "Vendors" tab of the Enquiry Book.
const VENDOR_NAMES = ['Garland and Bouquet', 'Wedding cake', 'Confetti', 'Cold Spark', 'Dry Ice', 'Ambiance lighting',
  'Parking Security', 'Live Band', 'Wedding Car', 'Generator', 'Pagoda', 'Table & Chair covers'];
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
