/**
 * ============================================================================
 *  HQA Spring · Staff Time Off  (Code.gs)
 * ============================================================================
 *  Staff request time off on a simple web page. HR / the Principal approve,
 *  ask for more info, or deny from a dashboard that shows who else is out,
 *  how many days each person has used this school year, and whether the
 *  dates land on a busy week (testing, conferences, PD).
 *
 *  Files in this Apps Script project:  Code.gs, Request.html, Review.html, Styles.html
 *
 *  1. Make a new Google Sheet > Extensions > Apps Script. Add the 4 files.
 *  2. Check CONFIG below. TRIAL MODE is on: every email goes only to TEST_EMAIL.
 *  3. Run setup() once and accept the permissions.
 *     Optional: run loadSampleRequests() to fill the dashboard with examples.
 *  4. Deploy > New deployment > Web app
 *       Execute as: Me   ·   Who has access: Anyone within hqaspring.org
 *     Staff:  <web app URL>          HR:  <web app URL>?page=review
 * ============================================================================
 */

const CONFIG = {
  // TRIAL MODE: every email goes ONLY to TEST_EMAIL (with a banner showing who
  // it would have gone to), and TEST_EMAIL can approve. Set to false to go live.
  TEST_MODE: true,
  TEST_EMAIL: 'finance@hqaspring.org',
  // In trial mode, try the dashboard as 'approver' (Principal/HR) or 'viewer' (admin, view only).
  TEST_ROLE: 'approver',

  SCHOOL_NAME: 'HQA Spring',
  APP_NAME: 'Staff Time Off',
  TIMEZONE: 'America/Chicago',

  // Who can approve or deny. Email -> name shown to staff ("Approved by ...").
  APPROVERS: {
    'principal@hqaspring.org': 'Principal',
    'hr@hqaspring.org': 'HR',
  },

  // Can open the HR dashboard and see everything, but can't approve or deny.
  VIEWERS: ['admin@hqaspring.org'],

  // FYI copy of every new request and decision.
  ADMIN_CC: ['admin@hqaspring.org'],

  // Told about approved absences so coverage / subs can be arranged.
  COVERAGE_NOTIFY: ['secretary@hqaspring.org'],

  // Optional staff-absence Google Calendar ID. Approved time off is added automatically.
  CALENDAR_ID: '',

  // Planned leave (not sick or emergencies) should be requested this many days ahead.
  MIN_DAYS_NOTICE: 5,
  NOTICE_EXEMPT: ['Sick', 'Family emergency', 'Bereavement'],

  // Optional yearly allowance in days. 0 = just show days used, no limit.
  ANNUAL_DAYS: 0,
  SCHOOL_YEAR_START_MONTH: 8, // August

  // Hours in a staff day (to convert partial days).
  HOURS_PER_DAY: 7,

  // Days the school is closed. They don't count as time off. (HQA 26–27 calendar)
  NO_SCHOOL: [
    { start: '2026-11-23', end: '2026-11-27', label: 'Thanksgiving Break' },
    { start: '2026-12-21', end: '2027-01-01', label: 'Winter Break' },
  ],

  // Busy weeks: requests that touch these get flagged for the approver.
  BUSY: [
    { start: '2026-12-03', end: '2026-12-17', label: 'MOY testing (3rd grade–HS)' },
    { start: '2026-12-18', end: '2026-12-18', label: 'Last day of first semester' },
    { start: '2027-01-04', end: '2027-01-04', label: 'Staff PD day' },
    { start: '2027-01-05', end: '2027-01-22', label: 'Parent-teacher conferences' },
    { start: '2027-01-11', end: '2027-01-28', label: 'MOY testing (Montessori–2nd)' },
  ],

  // Flag when this many other people are already out on any of the same days.
  OVERLAP_WARN: 2,

  DAILY_DIGEST: true,
  DIGEST_HOUR: 7,
  WEB_APP_URL: '',
};

const STATUS = {
  PENDING: 'Pending', APPROVED: 'Approved', INFO: 'Needs Info', DENIED: 'Denied', CANCELLED: 'Cancelled',
};
const STATUS_STYLE = {
  'Pending':    { bg: '#FEF3C7', fg: '#92400E' },
  'Approved':   { bg: '#DCFCE7', fg: '#166534' },
  'Needs Info': { bg: '#FFEDD5', fg: '#9A3412' },
  'Denied':     { bg: '#FEE2E2', fg: '#991B1B' },
  'Cancelled':  { bg: '#EEF0EE', fg: '#4B5563' },
};
const LEAVE_TYPES = ['Sick', 'Personal', 'Vacation', 'Family emergency', 'Bereavement',
  'Religious observance', 'Professional development', 'Jury duty', 'Unpaid', 'Other'];

const SHEET_NAME = 'Time Off';
const HEADERS = ['Request ID', 'Submitted', 'Status', 'Name', 'Email', 'Role / Dept', 'Leave Type',
  'First Day', 'Last Day', 'Duration', 'Days', 'Coverage', 'Reason', 'Decided By', 'Decided At', 'HR Note', 'Data'];
const COL = HEADERS.reduce((m, h, i) => (m[h] = i + 1, m), {});

/* ---------------------------------------------------------------------------
 *  Setup & routing
 * ------------------------------------------------------------------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  const rows = sh.getMaxRows() - 1;
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontFamily('Montserrat').setFontWeight('bold').setFontSize(10)
    .setFontColor('#E9F6E6').setBackground('#9EAF9A').setVerticalAlignment('middle');
  sh.setFrozenRows(1); sh.setRowHeight(1, 34);
  sh.setHiddenGridlines(true); sh.setTabColor('#9EAF9A');
  sh.getRange(2, 1, rows, HEADERS.length).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top');
  const w = { 'Request ID': 120, 'Submitted': 150, 'Status': 105, 'Name': 150, 'Email': 190, 'Role / Dept': 150,
    'Leave Type': 140, 'First Day': 120, 'Last Day': 120, 'Duration': 150, 'Days': 60, 'Coverage': 220,
    'Reason': 240, 'Decided By': 110, 'Decided At': 150, 'HR Note': 240 };
  Object.keys(w).forEach(h => sh.setColumnWidth(COL[h], w[h]));
  ['Coverage', 'Reason', 'HR Note'].forEach(h => sh.getRange(2, COL[h], rows, 1).setWrap(true));
  ['Submitted', 'Decided At'].forEach(h => sh.getRange(2, COL[h], rows, 1).setNumberFormat('mmm d, yyyy h:mm AM/PM'));
  ['First Day', 'Last Day'].forEach(h => sh.getRange(2, COL[h], rows, 1).setNumberFormat('ddd, mmm d, yyyy'));
  sh.hideColumns(COL['Data']);
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length - 1)
    .setBorder(null, null, null, null, null, true, '#DCE8D8', SpreadsheetApp.BorderStyle.SOLID);
  const st = sh.getRange(2, COL['Status'], rows, 1);
  st.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Object.values(STATUS), true).build());
  sh.setConditionalFormatRules(Object.keys(STATUS_STYLE).map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(s).setBackground(STATUS_STYLE[s].bg).setFontColor(STATUS_STYLE[s].fg).setRanges([st]).build()));

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendPendingDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (CONFIG.DAILY_DIGEST) ScriptApp.newTrigger('sendPendingDigest').timeBased().everyDays(1)
    .atHour(CONFIG.DIGEST_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  Logger.log('Setup done. Now Deploy > New deployment > Web app.');
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const file = page === 'review' ? 'Review' : 'Request';
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle((file === 'Review' ? 'Time Off Approvals · ' : 'Time Off · ') + CONFIG.SCHOOL_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/* ---------------------------------------------------------------------------
 *  Storage
 * ------------------------------------------------------------------------- */
function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Run setup() first.');
  return sh;
}

function newId_() {
  return 'TO-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMM') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
}

function findRow_(id) {
  const sh = sheet_();
  if (!id || sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function rowValues_(r) {
  return [r.id, new Date(r.submittedAt), r.status, r.name, r.email, r.role, r.leaveType,
    ymdToDate_(r.date), ymdToDate_(r.endDate || r.date), durationLabel_(r), r.days,
    [].concat(r.coverage || []).join(', ') + (r.coverageNotes ? '\n' + r.coverageNotes : ''),
    r.reason || '', r.decidedBy || '', r.decidedAt ? new Date(r.decidedAt) : '', r.note || '', JSON.stringify(r)];
}

function save_(r) {
  const row = findRow_(r.id);
  if (row) sheet_().getRange(row, 1, 1, HEADERS.length).setValues([rowValues_(r)]);
  else sheet_().appendRow(rowValues_(r));
}

function parse_(v) {
  let r;
  try { r = JSON.parse(v[COL['Data'] - 1]); } catch (e) { return null; }
  if (!r || !r.id) return null;
  r.status = v[COL['Status'] - 1] || STATUS.PENDING;
  r.note = v[COL['HR Note'] - 1] || '';
  r.decidedBy = v[COL['Decided By'] - 1] || '';
  const at = v[COL['Decided At'] - 1];
  r.decidedAt = at instanceof Date ? at.toISOString() : (r.decidedAt || '');
  return r;
}

function all_() {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues().map(parse_).filter(Boolean);
}

function get_(id) {
  const row = findRow_(id);
  return row ? parse_(sheet_().getRange(row, 1, 1, HEADERS.length).getValues()[0]) : null;
}

/* ---------------------------------------------------------------------------
 *  Dates & days
 * ------------------------------------------------------------------------- */
function ymdToDate_(ymd) { const p = String(ymd).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function ymd_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function hm_(t) { const m = String(t || '').match(/(\d{1,2}):(\d{2})/); return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : ''; }
function fmtTime_(t) { if (!t) return ''; const p = t.split(':').map(Number); return (p[0] % 12 || 12) + ':' + ('0' + p[1]).slice(-2) + (p[0] < 12 ? ' AM' : ' PM'); }
function fmtDate_(ymd, pat) { return ymd ? Utilities.formatDate(ymdToDate_(ymd), CONFIG.TIMEZONE, pat || 'EEE, MMM d, yyyy') : ''; }
function inRange_(ymd, p) { return ymd >= p.start && ymd <= p.end; }
function closedFor_(ymd) { return CONFIG.NO_SCHOOL.find(p => inRange_(ymd, p)); }

/** School days in the request (weekdays, minus no-school days). */
function schoolDays_(r) {
  const out = [];
  const d = ymdToDate_(r.date), end = ymdToDate_(r.endDate || r.date);
  while (d <= end) {
    const y = ymd_(d);
    if (d.getDay() !== 0 && d.getDay() !== 6 && !closedFor_(y)) out.push(y);
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function countDays_(r) {
  if (r.duration === 'Half day') return 0.5;
  if (r.duration === 'Partial day') {
    const a = r.start.split(':').map(Number), b = r.end.split(':').map(Number);
    const hrs = Math.max(0, (b[0] * 60 + b[1] - a[0] * 60 - a[1]) / 60);
    return Math.round(hrs / CONFIG.HOURS_PER_DAY * 100) / 100;
  }
  return schoolDays_(r).length;
}

function durationLabel_(r) {
  if (r.duration === 'Half day') return 'Half day (' + (r.halfPart || 'Morning') + ')';
  if (r.duration === 'Partial day') return 'Partial: ' + fmtTime_(r.start) + ' – ' + fmtTime_(r.end);
  return r.endDate && r.endDate !== r.date ? 'Full days' : 'Full day';
}

function dateLabel_(r) {
  if (r.endDate && r.endDate !== r.date) return fmtDate_(r.date, 'EEE, MMM d') + ' – ' + fmtDate_(r.endDate);
  return fmtDate_(r.date, 'EEEE, MMM d, yyyy');
}

function daysLabel_(n) {
  return (n === 1 ? '1 day' : n + ' days') + (n > 0 && n < 1 ? ' (' + Math.round(n * CONFIG.HOURS_PER_DAY * 10) / 10 + ' hrs)' : '');
}

function schoolYearStart_(ymd) {
  const d = ymdToDate_(ymd);
  const y = d.getMonth() + 1 >= CONFIG.SCHOOL_YEAR_START_MONTH ? d.getFullYear() : d.getFullYear() - 1;
  return ymd_(new Date(y, CONFIG.SCHOOL_YEAR_START_MONTH - 1, 1));
}

/** Approved days this school year for the same person (not counting r itself). */
function daysUsed_(r, list) {
  const from = schoolYearStart_(r.date);
  const to = schoolYearStart_(r.date).replace(/^\d{4}/, y => String(+y + 1));
  return (list || all_()).filter(x => x.id !== r.id && x.status === STATUS.APPROVED &&
      String(x.email).toLowerCase() === String(r.email).toLowerCase() && x.date >= from && x.date < to)
    .reduce((s, x) => s + (Number(x.days) || 0), 0);
}

/** Other people out (approved or pending) on any of the same days. */
function overlaps_(r, list) {
  const mine = schoolDays_(r);
  return (list || all_()).filter(x => x.id !== r.id &&
      (x.status === STATUS.APPROVED || x.status === STATUS.PENDING) &&
      String(x.email).toLowerCase() !== String(r.email).toLowerCase() &&
      schoolDays_(x).some(d => mine.indexOf(d) > -1))
    .map(x => ({ name: x.name, role: x.role, status: x.status, dateLabel: fmtDate_(x.date, 'MMM d') +
      (x.endDate && x.endDate !== x.date ? '–' + fmtDate_(x.endDate, 'MMM d') : ''), duration: durationLabel_(x) }));
}

function checks_(r, list) {
  const c = [];
  const sub = new Date(r.submittedAt);
  const notice = Math.round((ymdToDate_(r.date) - new Date(sub.getFullYear(), sub.getMonth(), sub.getDate())) / 864e5);
  if (CONFIG.NOTICE_EXEMPT.indexOf(r.leaveType) < 0) {
    c.push({ ok: notice >= CONFIG.MIN_DAYS_NOTICE, label: CONFIG.MIN_DAYS_NOTICE + '+ days notice',
      detail: notice < 0 ? 'Submitted after the date.' : notice + (notice === 1 ? ' day' : ' days') + ' ahead.' });
  }
  const days = schoolDays_(r);
  const busy = CONFIG.BUSY.filter(p => days.some(d => inRange_(d, p)));
  c.push({ ok: !busy.length, label: 'Not a busy week', detail: busy.map(p => p.label).join(', ') });
  const ov = overlaps_(r, list);
  c.push({ ok: ov.length < CONFIG.OVERLAP_WARN, label: 'Enough staff on campus',
    detail: ov.length ? ov.length + ' other ' + (ov.length === 1 ? 'person is' : 'people are') + ' out on the same day' : '' });
  if (/teacher/i.test(r.role)) {
    const cov = [].concat(r.coverage || []);
    c.push({ ok: cov.length > 0 && !(cov.length === 1 && cov[0] === 'Not needed'),
      label: 'Class coverage planned', detail: cov.join(', ') || 'No coverage plan given.' });
  }
  if (CONFIG.ANNUAL_DAYS) {
    const used = daysUsed_(r, list) + (Number(r.days) || 0);
    c.push({ ok: used <= CONFIG.ANNUAL_DAYS, label: 'Within ' + CONFIG.ANNUAL_DAYS + ' days this year', detail: used + ' days with this request.' });
  }
  if (!days.length) c.push({ ok: false, label: 'Falls on school days', detail: 'These dates are a weekend or school break.' });
  return c;
}

/* ---------------------------------------------------------------------------
 *  People & links
 * ------------------------------------------------------------------------- */
function me_() { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); }
function uniq_(list) { const s = {}; return list.map(e => String(e || '').trim().toLowerCase()).filter(e => e && !s[e] && (s[e] = true)); }
function approverEmails_() { const l = Object.keys(CONFIG.APPROVERS); return uniq_(CONFIG.TEST_MODE && CONFIG.TEST_ROLE !== 'viewer' ? l.concat([CONFIG.TEST_EMAIL]) : l); }
function viewerEmails_() { const l = CONFIG.VIEWERS || []; return uniq_(CONFIG.TEST_MODE && CONFIG.TEST_ROLE === 'viewer' ? l.concat([CONFIG.TEST_EMAIL]) : l); }
function canView_(e) { return isApprover_(e) || (!!e && viewerEmails_().indexOf(String(e).toLowerCase()) > -1); }
function isApprover_(e) { return !!e && approverEmails_().indexOf(String(e).toLowerCase()) > -1; }
function approverName_(e) {
  const k = Object.keys(CONFIG.APPROVERS).find(x => x.toLowerCase() === String(e).toLowerCase());
  if (k) return CONFIG.APPROVERS[k];
  return CONFIG.TEST_MODE && String(e).toLowerCase() === String(CONFIG.TEST_EMAIL).toLowerCase() ? 'HR (trial)' : e;
}
function url_() { if (CONFIG.WEB_APP_URL) return CONFIG.WEB_APP_URL; try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
function reviewLink_(r) { const u = url_(); return u ? u + '?page=review' + (r ? '&id=' + encodeURIComponent(r.id) : '') : ''; }
function staffLink_(r) { const u = url_(); return u ? u + '?page=mine' + (r ? '&id=' + encodeURIComponent(r.id) : '') : ''; }

function forClient_(r, list) {
  return Object.assign({}, r, {
    dateLabel: dateLabel_(r),
    shortDate: fmtDate_(r.date, 'EEE, MMM d') + (r.endDate && r.endDate !== r.date ? ' – ' + fmtDate_(r.endDate, 'EEE, MMM d') : ''),
    durationLabel: durationLabel_(r),
    daysLabel: daysLabel_(Number(r.days) || 0),
    submittedLabel: Utilities.formatDate(new Date(r.submittedAt), CONFIG.TIMEZONE, 'MMM d, h:mm a'),
    decidedLabel: r.decidedAt ? Utilities.formatDate(new Date(r.decidedAt), CONFIG.TIMEZONE, 'MMM d, h:mm a') : '',
    checks: checks_(r, list),
  });
}

/* ---------------------------------------------------------------------------
 *  Staff API
 * ------------------------------------------------------------------------- */
function getStaffData() {
  const me = me_();
  const list = all_();
  const mine = list.filter(r => String(r.email).toLowerCase() === me)
    .sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)));
  const today = ymd_(new Date());
  return {
    email: me,
    name: mine.length ? mine[0].name : '',
    role: mine.length ? mine[0].role : '',
    isApprover: canView_(me),
    reviewUrl: reviewLink_(),
    testMode: CONFIG.TEST_MODE,
    leaveTypes: LEAVE_TYPES,
    noticeExempt: CONFIG.NOTICE_EXEMPT,
    minNotice: CONFIG.MIN_DAYS_NOTICE,
    noSchool: CONFIG.NO_SCHOOL,
    busy: CONFIG.BUSY,
    annualDays: CONFIG.ANNUAL_DAYS,
    usedThisYear: daysUsed_({ id: '', email: me, date: today }, list),
    requests: mine.map(r => forClient_(r, list)),
  };
}

function submitTimeOff(x) {
  const me = me_();
  if (!me) throw new Error('Please sign in with your school Google account.');
  const s = (v, n) => String(v == null ? '' : v).trim().slice(0, n || 300);
  const ymd = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? String(v) : '';
  const d = {
    name: s(x.name, 120), role: s(x.role, 120), leaveType: s(x.leaveType, 60),
    duration: ['Full day(s)', 'Half day', 'Partial day'].indexOf(x.duration) > -1 ? x.duration : 'Full day(s)',
    date: ymd(x.date), endDate: ymd(x.endDate), halfPart: s(x.halfPart, 20),
    start: hm_(x.start), end: hm_(x.end),
    coverage: (Array.isArray(x.coverage) ? x.coverage : []).map(v => s(v, 80)).slice(0, 6),
    coverageNotes: s(x.coverageNotes, 1500), reason: s(x.reason, 1500),
  };
  if (!d.name || !d.role) throw new Error('Please add your name and role.');
  if (LEAVE_TYPES.indexOf(d.leaveType) < 0) throw new Error('Please choose a type of leave.');
  if (!d.date) throw new Error('Please choose a date.');
  if (d.duration !== 'Full day(s)') d.endDate = '';
  if (d.endDate && d.endDate < d.date) throw new Error('The last day is before the first day.');
  if (d.duration === 'Partial day' && (!d.start || !d.end || d.end <= d.start)) throw new Error('Please give a start and end time.');
  if (d.duration !== 'Partial day') { d.start = ''; d.end = ''; }
  d.days = countDays_(d);
  if (!d.days) throw new Error('Those dates are a weekend or school break, so no time off is needed.');

  let r, re = false;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (x.id) {
      r = get_(x.id);
      if (!r || String(r.email).toLowerCase() !== me) throw new Error('That request was not found.');
      if (r.status !== STATUS.PENDING && r.status !== STATUS.INFO) throw new Error('This request was already ' + r.status.toLowerCase() + '.');
      re = true;
    } else {
      r = { id: newId_(), submittedAt: new Date().toISOString(), email: me, history: [] };
    }
    Object.assign(r, d, { status: STATUS.PENDING, note: '', decidedBy: '', decidedAt: '' });
    r.history = (r.history || []).concat([{ at: new Date().toISOString(), by: d.name, action: re ? 'Resubmitted' : 'Submitted' }]);
    save_(r);
  } finally { lock.releaseLock(); }
  notifyNew_(r, re);
  return forClient_(r);
}

/** Staff can cancel their own request (pending, needs info, or approved and not over yet). */
function cancelTimeOff(id) {
  const me = me_();
  let r;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    r = get_(id);
    if (!r || String(r.email).toLowerCase() !== me) throw new Error('That request was not found.');
    if (r.status === STATUS.DENIED || r.status === STATUS.CANCELLED) throw new Error('This request is already ' + r.status.toLowerCase() + '.');
    if ((r.endDate || r.date) < ymd_(new Date())) throw new Error('This time off is already over.');
    const was = r.status;
    r.status = STATUS.CANCELLED;
    r.history = (r.history || []).concat([{ at: new Date().toISOString(), by: r.name, action: 'Cancelled' }]);
    save_(r);
    syncCalendar_(r);
    send_(approverEmails_(), 'Cancelled: ' + subject_(r), email_({
      accent: '#6B7280', eyebrow: 'Time off cancelled', heading: r.name + ' · ' + r.leaveType, status: STATUS.CANCELLED,
      intro: esc_(r.name) + ' cancelled this request' + (was === STATUS.APPROVED ? ', which had been approved.' : '.'),
      rows: rows_(r) }), was === STATUS.APPROVED ? CONFIG.ADMIN_CC.concat(CONFIG.COVERAGE_NOTIFY) : CONFIG.ADMIN_CC);
  } finally { lock.releaseLock(); }
  return forClient_(r);
}

/* ---------------------------------------------------------------------------
 *  HR API
 * ------------------------------------------------------------------------- */
function getReviewData() {
  const me = me_();
  if (!canView_(me)) return { denied: true, email: me };
  const canDecide = isApprover_(me);
  const list = all_();
  const today = ymd_(new Date());
  const week = ymd_(new Date(Date.now() + 6 * 864e5));
  const outSoon = list.filter(r => r.status === STATUS.APPROVED && (r.endDate || r.date) >= today && r.date <= week)
    .sort((a, b) => a.date.localeCompare(b.date))
    .map(r => ({ id: r.id, name: r.name, role: r.role, leaveType: r.leaveType, dateLabel: fmtDate_(r.date, 'EEE, MMM d') +
      (r.endDate && r.endDate !== r.date ? ' – ' + fmtDate_(r.endDate, 'EEE, MMM d') : ''), duration: durationLabel_(r) }));
  return {
    email: me, name: canDecide ? approverName_(me) : me, canDecide: canDecide, testMode: CONFIG.TEST_MODE,
    outSoon: outSoon,
    requests: list.map(r => Object.assign(forClient_(r, list), {
      usedThisYear: daysUsed_(r, list), overlaps: overlaps_(r, list),
    })).sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt))),
  };
}

function decideTimeOff(id, decision, note, force) {
  const me = me_();
  if (!isApprover_(me)) throw new Error('Only HR or the Principal can make decisions.');
  if ([STATUS.APPROVED, STATUS.INFO, STATUS.DENIED].indexOf(decision) < 0) throw new Error('Unknown decision.');
  note = String(note || '').trim().slice(0, 2000);
  if (decision === STATUS.INFO && !note) throw new Error('Add a note saying what information you need.');
  let r;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    r = get_(id);
    if (!r) throw new Error('Request not found.');
    if (r.status === STATUS.CANCELLED) throw new Error('This request was cancelled by ' + r.name + '.');
    if (r.status !== STATUS.PENDING && !force) return { conflict: true, status: r.status, decidedBy: r.decidedBy };
    Object.assign(r, { status: decision, note: note, decidedBy: approverName_(me), decidedByEmail: me, decidedAt: new Date().toISOString() });
    r.history = (r.history || []).concat([{ at: r.decidedAt, by: r.decidedBy, action: decision, note: note }]);
    save_(r);
  } finally { lock.releaseLock(); }
  syncCalendar_(r);
  notifyDecision_(r);
  const list = all_();
  return { ok: true, request: Object.assign(forClient_(r, list), { usedThisYear: daysUsed_(r, list), overlaps: overlaps_(r, list) }) };
}

/* ---------------------------------------------------------------------------
 *  Calendar (optional)
 * ------------------------------------------------------------------------- */
function syncCalendar_(r) {
  if (!CONFIG.CALENDAR_ID) return;
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!cal) return;
    if (r.status === STATUS.APPROVED && !r.calendarEventId) {
      const title = r.name + ' out (' + durationLabel_(r) + ')';
      const opts = { description: r.leaveType + ' · approved by ' + r.decidedBy + (r.coverage && r.coverage.length ? '\nCoverage: ' + r.coverage.join(', ') : '') };
      let ev;
      if (r.duration === 'Partial day') {
        const at = t => { const d = ymdToDate_(r.date); const p = t.split(':'); d.setHours(+p[0], +p[1]); return d; };
        ev = cal.createEvent(title, at(r.start), at(r.end), opts);
      } else {
        const end = ymdToDate_(r.endDate || r.date); end.setDate(end.getDate() + 1);
        ev = cal.createAllDayEvent(title, ymdToDate_(r.date), end, opts);
      }
      r.calendarEventId = ev.getId();
    } else if (r.status !== STATUS.APPROVED && r.calendarEventId) {
      const ev = cal.getEventById(r.calendarEventId);
      if (ev) ev.deleteEvent();
      r.calendarEventId = '';
    } else return;
    save_(r);
  } catch (err) { console.warn('Calendar sync failed: ' + err); }
}

/* ---------------------------------------------------------------------------
 *  Emails
 * ------------------------------------------------------------------------- */
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function subject_(r) { return r.name + ' · ' + r.leaveType + ' · ' + fmtDate_(r.date, 'EEE, MMM d') + (r.endDate && r.endDate !== r.date ? '–' + fmtDate_(r.endDate, 'MMM d') : ''); }

function rows_(r, privateToo) {
  return [
    ['Request', r.id],
    ['Staff member', r.name + ' (' + r.email + ')'],
    ['Role / dept', r.role],
    ['Type of leave', r.leaveType],
    ['Dates', dateLabel_(r)],
    ['Duration', durationLabel_(r)],
    ['Counts as', daysLabel_(Number(r.days) || 0)],
    ['Coverage', [].concat(r.coverage || []).join(', ')],
    ['Coverage notes', r.coverageNotes],
  ].concat(privateToo ? [['Reason', r.reason]] : []);
}

function email_(o) {
  const accent = o.accent || '#1E459D';
  const rows = (o.rows || []).filter(r => r[1] !== '' && r[1] != null).map(r =>
    '<tr><td style="padding:9px 14px;border-bottom:1px solid #EBEFF6;color:#5B667A;font-size:13px;width:36%;vertical-align:top;">' + esc_(r[0]) +
    '</td><td style="padding:9px 14px;border-bottom:1px solid #EBEFF6;color:#172033;font-size:14px;vertical-align:top;">' + esc_(r[1]).replace(/\n/g, '<br>') + '</td></tr>').join('');
  const checks = (o.checks || []).map(c =>
    '<tr><td style="padding:5px 0;vertical-align:top;width:26px;"><span style="display:inline-block;width:18px;height:18px;line-height:18px;border-radius:9px;text-align:center;font-size:12px;font-weight:bold;color:#fff;background-color:' +
    (c.ok ? '#16A34A' : '#EA580C') + ';">' + (c.ok ? '&#10003;' : '!') + '</span></td><td style="padding:5px 0;font-size:14px;color:#172033;"><b>' + esc_(c.label) + '</b>' +
    (c.detail ? '<span style="color:#5B667A;"> · ' + esc_(c.detail) + '</span>' : '') + '</td></tr>').join('');
  const st = o.status ? '<span style="display:inline-block;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:bold;background-color:' +
    STATUS_STYLE[o.status].bg + ';color:' + STATUS_STYLE[o.status].fg + ';">' + esc_(o.status) + '</span>' : '';
  const note = o.note ? '<div style="margin:18px 0 0;padding:14px 16px;border-left:4px solid ' + accent + ';background-color:#F4F6FB;border-radius:6px;">' +
    '<div style="font-size:12px;color:#5B667A;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;">' + esc_(o.noteLabel || 'Note') + '</div>' +
    '<div style="font-size:15px;color:#172033;">' + esc_(o.note).replace(/\n/g, '<br>') + '</div></div>' : '';
  const btn = o.button && o.button.url ? '<div style="margin:24px 0 4px;"><a href="' + esc_(o.button.url) + '" style="display:inline-block;background-color:' + accent +
    ';color:#fff;text-decoration:none;font-weight:bold;font-size:15px;padding:12px 26px;border-radius:8px;">' + esc_(o.button.label) + '</a></div>' : '';
  return '<!doctype html><html><body bgcolor="#EEF2F8" style="margin:0;padding:0;background-color:#EEF2F8;">' +
    '<table role="presentation" width="100%" bgcolor="#EEF2F8" cellpadding="0" cellspacing="0" style="background-color:#EEF2F8;padding:24px 12px;"><tr><td align="center">' +
    '<table role="presentation" width="100%" bgcolor="#FFFFFF" cellpadding="0" cellspacing="0" style="max-width:600px;background-color:#FFFFFF;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;">' +
    '<tr><td bgcolor="#17377F" align="center" style="background-color:#17377F;padding:30px 26px 24px;text-align:center;">' +
    '<img src="cid:hqalogo" width="72" height="72" alt="HQA Spring" style="display:block;margin:0 auto 12px;border:0;width:72px;height:72px;">' +
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#FFFFFF;font-size:19px;font-weight:bold;letter-spacing:.02em;">Houston Quran Academy Spring</div>' +
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#B9C7EA;font-size:12px;letter-spacing:.16em;text-transform:uppercase;margin-top:6px;">' + esc_(CONFIG.APP_NAME) + '</div></td></tr>' +
    '<tr><td bgcolor="#E03A2C" style="background-color:#E03A2C;height:5px;font-size:0;line-height:0;">&nbsp;</td></tr>' +
    '<tr><td style="padding:26px 26px 8px;">' +
    (o.eyebrow ? '<div style="font-size:12px;color:' + accent + ';font-weight:bold;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">' + esc_(o.eyebrow) + '</div>' : '') +
    '<div style="font-size:22px;line-height:1.3;font-weight:bold;color:#172033;margin:0 0 8px;">' + esc_(o.heading) + '</div>' + st +
    (o.intro ? '<p style="font-size:15px;line-height:1.55;color:#34405A;margin:14px 0 0;">' + o.intro + '</p>' : '') + note + btn + '</td></tr>' +
    (checks ? '<tr><td style="padding:16px 26px 0;"><div style="font-size:12px;color:#5B667A;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Checks</div><table role="presentation" cellpadding="0" cellspacing="0">' + checks + '</table></td></tr>' : '') +
    (rows ? '<tr><td style="padding:18px 26px 6px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #DFE5F0;border-radius:10px;border-collapse:separate;overflow:hidden;">' + rows + '</table></td></tr>' : '') +
    '<tr><td style="padding:18px 26px 26px;font-size:12px;color:#8792A8;">Automated message from the ' + esc_(CONFIG.SCHOOL_NAME) + ' time off system.</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function send_(to, subject, html, cc) {
  to = uniq_([].concat(to || []));
  cc = uniq_([].concat(cc || [])).filter(e => to.indexOf(e) < 0);
  if (!to.length) return;
  if (CONFIG.TEST_MODE) {
    const banner = '<div style="background-color:#FEF3C7;color:#92400E;font:13px Arial,sans-serif;padding:10px 16px;text-align:center;">' +
      '<b>TRIAL MODE</b> · Only you received this. It would have gone to: ' + esc_(to.join(', ')) + (cc.length ? ' · cc: ' + esc_(cc.join(', ')) : '') + '</div>';
    html = html.replace('<body bgcolor="#EEF2F8" style="margin:0;padding:0;background-color:#EEF2F8;">', '$&' + banner);
    subject = '[TRIAL] ' + subject;
    to = [String(CONFIG.TEST_EMAIL).toLowerCase()];
    cc = [];
  }
  const msg = { to: to.join(','), subject: subject, htmlBody: html, name: CONFIG.SCHOOL_NAME + ' Time Off' };
  if (cc.length) msg.cc = cc.join(',');
  if (html.indexOf('cid:hqalogo') > -1) msg.inlineImages = { hqalogo: logoBlob_() };
  MailApp.sendEmail(msg);
}

function notifyNew_(r, re) {
  const checks = checks_(r);
  const flags = checks.filter(c => !c.ok).length;
  const used = daysUsed_(r);
  send_(approverEmails_(), (re ? '[Resubmitted] ' : '[Time off request] ') + subject_(r), email_({
    accent: '#17377F', eyebrow: re ? 'Updated time off request' : 'New time off request',
    heading: r.name + ' · ' + r.leaveType, status: STATUS.PENDING,
    intro: esc_(r.name) + ' (' + esc_(r.role) + ') is asking for <b>' + esc_(daysLabel_(r.days)) + '</b> off: ' + esc_(dateLabel_(r)) + '.' +
      ' They have used <b>' + used + (used === 1 ? ' day' : ' days') + '</b> so far this school year.' +
      (flags ? ' <span style="color:#C2410C;font-weight:bold;">' + flags + (flags > 1 ? ' checks need' : ' check needs') + ' a look.</span>' : ''),
    checks: checks, rows: rows_(r, true), button: { label: 'Review request', url: reviewLink_(r) },
  }), CONFIG.ADMIN_CC);

  send_(r.email, 'Time off request received: ' + fmtDate_(r.date, 'EEE, MMM d'), email_({
    eyebrow: 'Request received', heading: r.leaveType + ' · ' + dateLabel_(r), status: STATUS.PENDING,
    intro: 'Assalamu alaikum ' + esc_(greetName_(r.name)) + ', your request was sent to HR and the Principal. You\'ll get an email as soon as it\'s reviewed.',
    rows: rows_(r, true), button: staffLink_(r) ? { label: 'View my time off', url: staffLink_(r) } : null,
  }));
}

function notifyDecision_(r) {
  const accent = { 'Approved': '#16A34A', 'Needs Info': '#EA580C', 'Denied': '#DC2626' }[r.status];
  const word = { 'Approved': 'approved', 'Needs Info': 'sent back with a question', 'Denied': 'not approved' }[r.status];
  send_(r.email, r.status + ': your time off on ' + fmtDate_(r.date, 'EEE, MMM d'), email_({
    accent: accent, eyebrow: 'Time off', heading: r.leaveType + ' · ' + dateLabel_(r), status: r.status,
    intro: 'Your request was <b>' + word + '</b> by ' + esc_(r.decidedBy) + '.' +
      (r.status === STATUS.INFO ? ' Please answer the note below and resubmit.' : ''),
    note: r.note, noteLabel: 'Note from ' + r.decidedBy, rows: rows_(r, true),
    button: staffLink_(r) ? { label: r.status === STATUS.INFO ? 'Update & resubmit' : 'View my time off', url: staffLink_(r) } : null,
  }));
  const others = approverEmails_().filter(e => e !== String(r.decidedByEmail || '').toLowerCase());
  send_(CONFIG.ADMIN_CC.concat(others), r.status + ' by ' + r.decidedBy + ': ' + subject_(r), email_({
    accent: accent, eyebrow: 'FYI', heading: r.name + ' · ' + r.leaveType, status: r.status,
    intro: esc_(r.decidedBy) + ' marked this request as <b>' + esc_(r.status) + '</b>.',
    note: r.note, noteLabel: 'Note', rows: rows_(r, true), button: { label: 'Open dashboard', url: reviewLink_(r) },
  }));
  if (r.status === STATUS.APPROVED && CONFIG.COVERAGE_NOTIFY.length) {
    send_(CONFIG.COVERAGE_NOTIFY, 'Staff out: ' + subject_(r), email_({
      accent: '#17377F', eyebrow: 'Approved absence', heading: r.name + ' is out ' + dateLabel_(r), status: STATUS.APPROVED,
      intro: 'For coverage planning. ' + esc_(r.name) + ' (' + esc_(r.role) + ') will be out: <b>' + esc_(durationLabel_(r)) + '</b>.',
      rows: rows_(r, false),
    }));
  }
}

function sendPendingDigest() {
  if (Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'u')) > 5) return;
  const pending = all_().filter(r => r.status === STATUS.PENDING).sort((a, b) => a.date.localeCompare(b.date));
  if (!pending.length) return;
  send_(approverEmails_(), pending.length + ' time off request' + (pending.length > 1 ? 's' : '') + ' waiting', email_({
    eyebrow: 'Morning summary', heading: pending.length + ' time off request' + (pending.length > 1 ? 's are' : ' is') + ' waiting',
    intro: 'Soonest dates first.',
    rows: pending.map(r => [fmtDate_(r.date, 'EEE, MMM d'), r.name + ' · ' + r.leaveType + ' · ' + daysLabel_(r.days)]),
    button: { label: 'Open dashboard', url: reviewLink_() },
  }));
}

/* ---------------------------------------------------------------------------
 *  TRIAL: sample requests from made-up staff (no emails sent).
 *  Delete the sample rows from the Time Off tab when you're done testing.
 * ------------------------------------------------------------------------- */
function loadSampleRequests() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const day = n => { const d = new Date(today); d.setDate(d.getDate() + n); while (d.getDay() === 0 || d.getDay() === 6 || closedFor_(ymd_(d))) d.setDate(d.getDate() + 1); return ymd_(d); };
  const now = new Date().toISOString();
  const samples = [
    { name: 'Ms. Hana (sample)', email: 'sample.staff1@hqaspring.org', role: '4th Grade Teacher', leaveType: 'Personal',
      duration: 'Full day(s)', date: day(2), endDate: '', coverage: ['Lesson plans in sub folder', 'Need a substitute'],
      reason: 'Moving apartments.' },
    { name: 'Mr. Omar (sample)', email: 'sample.staff2@hqaspring.org', role: '3rd Grade Teacher', leaveType: 'Vacation',
      duration: 'Full day(s)', date: day(10), endDate: day(12), coverage: ['Co-teacher covering'], coverageNotes: 'Ms. Noor will cover 3B.',
      reason: 'Family wedding out of state.' },
    { name: 'Sr. Maryam (sample)', email: 'sample.staff3@hqaspring.org', role: 'Front office', leaveType: 'Sick',
      duration: 'Half day', halfPart: 'Afternoon', date: day(1), coverage: ['Not needed'], reason: 'Doctor appointment.' },
    { name: 'Ms. Noor (sample)', email: 'sample.staff4@hqaspring.org', role: '5th Grade Teacher', leaveType: 'Professional development',
      duration: 'Full day(s)', date: day(2), coverage: ['Lesson plans in sub folder', 'Need a substitute'], reason: 'Region 4 science workshop.',
      status: STATUS.APPROVED, decidedBy: 'HR (trial)', note: 'Approved. Please leave plans by Tuesday.' },
  ];
  samples.forEach(s => {
    const r = Object.assign({ id: newId_(), submittedAt: now, status: STATUS.PENDING, note: '', decidedBy: '', decidedAt: '',
      start: '', end: '', coverageNotes: '', history: [{ at: now, by: s.name, action: 'Submitted' }] }, s);
    r.days = countDays_(r);
    if (r.status !== STATUS.PENDING) { r.decidedAt = now; r.history.push({ at: now, by: r.decidedBy, action: r.status, note: r.note }); }
    save_(r);
  });
}

/** "Ms. Hana Ali" -> "Ms. Hana", "Hana Ali" -> "Hana" (for greetings). */
function greetName_(name) {
  const w = String(name || '').trim().split(/\s+/);
  return /^(ms|mr|mrs|miss|dr|sr|br|sh|ustadh|ustadha|sister|brother)\.?$/i.test(w[0]) && w[1] ? w[0] + ' ' + w[1] : w[0];
}
