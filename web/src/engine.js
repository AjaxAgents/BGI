/*
 * Calculation engine. Hard rules only: every figure the report shows is either
 * read straight from an input file or produced by one of the formulas below.
 * No estimates, no AI, no network. All money is held in whole cents.
 *
 * Inputs
 *   1. The UNDERBILLING_CASHPOSITION_RPT Excel export (one row per job, plus a
 *      SUBJOB_DETAIL column carrying each sub-job's figures).
 *   2. The month-end Adjusted Underbillings schedule (CSV, one row per line of
 *      the schedule the CFO issues).
 *
 * Rules are documented for people in docs/RULES.md; keep the two in step.
 */
(function (BGI) {
  'use strict';

  var RULES = {
    SMALL_JOB_THRESHOLD_DOLLARS: 10000, // schedule groups jobs under this into one line
    BILLED_OVER_CONTRACT_TOLERANCE: 1.05, // billed may exceed contract by 5% before it is flagged
    IDENTITY_TOLERANCE_CENTS: 2,          // rounding allowance when re-checking the export's own math
    SCHEDULE_MATCH_TOLERANCE_DOLLARS: 1   // schedule is printed in whole dollars
  };

  var REQUIRED_COLUMNS = [
    'BU_CODE', 'JOB_CODE', 'JOB_NAME', 'PX', 'CONTRACT_VALUE', 'START_DATE',
    'COST_TO_DATE', 'BILLED_TO_DATE', 'UNDERBILLED_AMOUNT', 'OVERBILLED_AMOUNT',
    'NET_BILL_AMOUNT', 'AR_PAID_AMOUNT', 'AR_OUTSTANDING_AMOUNT',
    'AP_PAID_AMOUNT', 'PAID_AMOUNT_VOUCHER', 'PAID_AMOUNT_NON_VOUCHER',
    'AP_OUTSTANDING_AMOUNT', 'AP_OPEN_COUNT', 'CASH_POSITION', 'SUBJOB_DETAIL'
  ];

  // Positional layout of each entry in SUBJOB_DETAIL (BigQuery struct order).
  var SUB_FIELDS = ['code', 'name', 'status', 'arStatus', 'contract', 'actStart', 'estStart', 'actCompl', 'estCompl',
    'start', 'end', 'cost', 'billed', 'under', 'netBill', 'cumBill', 'arPaid', 'arOut', 'apPaid', 'apV', 'apNV',
    'apOut', 'apOpen', 'lastJbb', 'lastAr', 'apCalc'];
  var SUB_MONEY = ['contract', 'cost', 'billed', 'under', 'netBill', 'cumBill', 'arPaid', 'arOut', 'apPaid', 'apV', 'apNV', 'apOut'];

  function EngineError(message) { var e = new Error(message); e.name = 'EngineError'; return e; }

  // ---------- primitives ----------

  function cents(v) {
    if (v === null || v === undefined || v === '') return 0;
    var n = typeof v === 'number' ? v : Number(String(v).replace(/[$,\s]/g, ''));
    if (!isFinite(n)) return 0;
    return Math.round(n * 100);
  }
  function text(v) { return v === null || v === undefined ? '' : String(v).trim(); }

  // Excel serial date -> 'YYYY-MM-DD'. Accepts ISO-like strings too.
  function isoDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') {
      if (v < 1 || v > 2958465) return null;
      var ms = Math.round((Math.floor(v) - 25569) * 86400000);
      return new Date(ms).toISOString().slice(0, 10);
    }
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(v));
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(String(v).trim());
    if (m) {
      var y = m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
      return y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    }
    return null;
  }

  // ---------- AP categories ----------

  var CATEGORIES = [
    { id: 'GC', label: 'GCs', long: 'General conditions', requested: true },
    { id: 'GR', label: 'GRs', long: 'General requirements', requested: true },
    { id: 'GCGR', label: 'GCs / GRs combined', long: 'Sub-job that holds general conditions and general requirements together' },
    { id: 'TRADES', label: 'Trades', long: 'Trade contractors / cost of work', requested: true },
    { id: 'INS', label: 'Insurance', long: 'Insurance', requested: true },
    { id: 'PRECON', label: 'Precon', long: 'Preconstruction' },
    { id: 'PAYROLL', label: 'Payroll', long: 'Payroll' },
    { id: 'NR', label: 'Non-reimbursable', long: 'Non-reimbursable cost' },
    { id: 'FEE', label: 'GMP / fee', long: 'GMP and fee sub-jobs' },
    { id: 'ODP', label: 'Owner direct purchase', long: 'Owner direct purchase (ODP)' },
    { id: 'SEC', label: 'Ops / exec', long: 'Operations and executive sub-jobs (98 / 99)' },
    { id: 'EST', label: 'Estimating', long: 'Estimating (ES) job' },
    { id: 'OTHER', label: 'Other sub-jobs', long: 'Sub-jobs whose name does not state a category' },
    { id: 'MAIN', label: 'Main job number', long: 'Posted to the main job number, not to a category sub-job' }
  ];
  var CATEGORY_BY_ID = {};
  CATEGORIES.forEach(function (c, i) { c.order = i; CATEGORY_BY_ID[c.id] = c; });

  // Ordered: first match wins. Tested against the end of the sub-job name.
  var CATEGORY_RULES = [
    ['INS', /insur/i],
    ['NR', /non[\s-]*r(ei|ie)mb|noon\s+reimb|\bNRs?\b/i],
    ['GCGR', /GC'?s?\s*(\/|&|and)?\s*GR'?s?\s*$/i],
    ['GR', /\bGR'?s?\s*$|general\s+req/i],
    ['GC', /\bGC'?s?\s*$|general\s+cond/i],
    ['TRADES', /trades?\s*$|cost\s+of\s+(the\s+)?work|\bCOW\s*$/i],
    ['PRECON', /pre-?\s?con/i],
    ['PAYROLL', /payroll|\bPR\s*$/i],
    ['FEE', /\bGMP\s*$|\bfee\s*$/i],
    ['ODP', /\bODP\s*$/i],
    ['SEC', /\bops\b|\bexec\b|bonus|commission/i]
  ];

  function categorize(subCode, subName, jobCode) {
    if (subCode === jobCode) return 'MAIN';
    var name = text(subName);
    for (var i = 0; i < CATEGORY_RULES.length; i++) if (CATEGORY_RULES[i][1].test(name)) return CATEGORY_RULES[i][0];
    var suffix = subCode.indexOf(jobCode) === 0 ? subCode.slice(jobCode.length) : '';
    if (suffix === '98' || suffix === '99') return 'SEC';
    if (/ES$/i.test(subCode)) return 'EST';
    return 'OTHER';
  }

  // ---------- Excel export ----------

  function parseSubjobs(raw, jobCode, problems) {
    var out = [];
    if (!raw) return out;
    var parsed;
    try { parsed = JSON.parse(raw); } catch (e) { problems.push('Job ' + jobCode + ': SUBJOB_DETAIL could not be read.'); return out; }
    var list = parsed && Array.isArray(parsed.v) ? parsed.v : [];
    for (var i = 0; i < list.length; i++) {
      var f = list[i] && list[i].v && Array.isArray(list[i].v.f) ? list[i].v.f : null;
      if (!f || f.length < SUB_FIELDS.length) { problems.push('Job ' + jobCode + ': a sub-job entry has an unexpected layout.'); continue; }
      var s = {};
      for (var k = 0; k < SUB_FIELDS.length; k++) s[SUB_FIELDS[k]] = f[k] ? f[k].v : null;
      for (var j = 0; j < SUB_MONEY.length; j++) s[SUB_MONEY[j]] = cents(s[SUB_MONEY[j]]);
      s.code = text(s.code);
      s.name = text(s.name);
      s.apOpen = parseInt(s.apOpen, 10) || 0;
      s.category = categorize(s.code, s.name, jobCode);
      out.push(s);
    }
    return out;
  }

  function parseExport(rows) {
    if (!rows || rows.length < 2) throw EngineError('The Excel sheet is empty.');
    var header = rows[0].map(text);
    var col = {};
    header.forEach(function (h, i) { if (h) col[h] = i; });
    var missing = REQUIRED_COLUMNS.filter(function (c) { return !(c in col); });
    if (missing.length) {
      throw EngineError('This does not look like the UNDERBILLING_CASHPOSITION export. Missing column' +
        (missing.length > 1 ? 's' : '') + ': ' + missing.join(', ') + '.');
    }
    function get(r, name) { return name in col ? r[col[name]] : null; }

    var jobs = [];
    var byCode = {};
    var problems = [];
    var apCalcDates = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r || !text(get(r, 'JOB_CODE'))) continue;
      var code = text(get(r, 'JOB_CODE'));
      if (byCode[code]) { problems.push('Job ' + code + ' appears more than once; the first row is used.'); continue; }
      var job = {
        code: code,
        name: text(get(r, 'JOB_NAME')),
        bu: text(get(r, 'BU_CODE')),
        px: text(get(r, 'PX')),
        rd: text(get(r, 'RD')),
        rd2: text(get(r, 'RD_2')),
        arStatus: text(get(r, 'AR_STATUS_DESCRIPTION')),
        startDate: isoDate(get(r, 'START_DATE')),
        lastArPost: isoDate(get(r, 'LAST_AR_POST_DATE')),
        lastBillPost: isoDate(get(r, 'LAST_JBB_POST_DATE')),
        apCalcDate: isoDate(get(r, 'AP_CALCULATION_DATE')),
        hasContract: get(r, 'CONTRACT_VALUE') !== null && get(r, 'CONTRACT_VALUE') !== undefined,
        contract: cents(get(r, 'CONTRACT_VALUE')),
        cost: cents(get(r, 'COST_TO_DATE')),
        billed: cents(get(r, 'BILLED_TO_DATE')),
        underReported: cents(get(r, 'UNDERBILLED_AMOUNT')),
        overReported: cents(get(r, 'OVERBILLED_AMOUNT')),
        position: text(get(r, 'BILLING_POSITION')),
        netBill: cents(get(r, 'NET_BILL_AMOUNT')),
        cumBill: cents(get(r, 'CUM_BILL_AMOUNT')),
        arPaid: cents(get(r, 'AR_PAID_AMOUNT')),
        arOut: cents(get(r, 'AR_OUTSTANDING_AMOUNT')),
        apPaid: cents(get(r, 'AP_PAID_AMOUNT')),
        apV: cents(get(r, 'PAID_AMOUNT_VOUCHER')),
        apNV: cents(get(r, 'PAID_AMOUNT_NON_VOUCHER')),
        apOut: cents(get(r, 'AP_OUTSTANDING_AMOUNT')),
        apOpen: Math.round(Number(get(r, 'AP_OPEN_COUNT')) || 0),
        cashReported: cents(get(r, 'CASH_POSITION')),
        cashPctReported: typeof get(r, 'CASH_POSITION_PERCENT') === 'number' ? get(r, 'CASH_POSITION_PERCENT') : null,
        subCount: Math.round(Number(get(r, 'JOB_AND_SUBJOB_COUNT')) || 0)
      };
      job.subs = parseSubjobs(get(r, 'SUBJOB_DETAIL'), code, problems);
      // Derived by this engine (never trusted from the file):
      job.under = job.cost - job.billed;          // Rule E1
      job.cash = job.arPaid - job.apPaid;         // Rule E3
      job.cashPct = job.contract !== 0 ? job.cash / job.contract : null; // Rule E7
      job.contractSubs = job.subs.reduce(function (a, s) { return a + s.contract; }, 0);
      job.checks = exportChecksForJob(job);
      if (job.apCalcDate) apCalcDates[job.apCalcDate] = (apCalcDates[job.apCalcDate] || 0) + 1;
      jobs.push(job);
      byCode[code] = job;
    }
    if (!jobs.length) throw EngineError('The Excel sheet has headers but no job rows.');
    var asOf = Object.keys(apCalcDates).sort().pop() || null;
    return { jobs: jobs, byCode: byCode, problems: problems, asOf: asOf, apCalcDates: apCalcDates };
  }

  function near(a, b) { return Math.abs(a - b) <= RULES.IDENTITY_TOLERANCE_CENTS; }
  function sumSubs(job, field) { return job.subs.reduce(function (a, s) { return a + s[field]; }, 0); }

  // Identity and plausibility checks on one row of the export.
  function exportChecksForJob(job) {
    var c = {};
    c.E1 = near(job.underReported, job.cost - job.billed);
    c.E2 = near(job.overReported, -job.underReported);
    c.E3 = near(job.cashReported, job.arPaid - job.apPaid);
    c.E4 = near(job.apPaid, job.apV + job.apNV);
    c.E5 = near(job.netBill, job.arPaid + job.arOut);
    c.E6 = !job.subs.length || (near(job.cost, sumSubs(job, 'cost')) && near(job.billed, sumSubs(job, 'billed')) &&
      near(job.netBill, sumSubs(job, 'netBill')) && near(job.arPaid, sumSubs(job, 'arPaid')) && near(job.arOut, sumSubs(job, 'arOut')) &&
      near(job.apPaid, sumSubs(job, 'apPaid')) && near(job.apOut, sumSubs(job, 'apOut')));
    c.E7 = job.cashPctReported === null || job.contract === 0 || Math.abs(job.cashPctReported - job.cash / job.contract) < 1e-6;
    // Plausibility: is the figure believable enough to act on?
    c.P1 = !(job.apPaid > 0 && job.cost < job.apPaid); // cost to date can never be less than what has already been paid out
    c.P2 = !(job.contract > 0 && job.billed > job.contract * RULES.BILLED_OVER_CONTRACT_TOLERANCE);
    return c;
  }

  var CHECK_TEXT = {
    E1: ['Underbilled amount = cost to date − billed to date', 'UNDERBILLED_AMOUNT = COST_TO_DATE − BILLED_TO_DATE'],
    E2: ['Overbilled amount is the mirror of underbilled', 'OVERBILLED_AMOUNT = −UNDERBILLED_AMOUNT'],
    E3: ['Cash position = collected from owner − paid out', 'CASH_POSITION = AR_PAID_AMOUNT − AP_PAID_AMOUNT'],
    E4: ['AP paid = voucher + non-voucher payments', 'AP_PAID_AMOUNT = PAID_AMOUNT_VOUCHER + PAID_AMOUNT_NON_VOUCHER'],
    E5: ['Net billed = collected + still owed by owner', 'NET_BILL_AMOUNT = AR_PAID_AMOUNT + AR_OUTSTANDING_AMOUNT'],
    E6: ['Job row = the sum of its sub-jobs', 'Each money column on the job row equals the sum of SUBJOB_DETAIL'],
    E7: ['Cash position % = cash position ÷ contract value', 'CASH_POSITION_PERCENT = CASH_POSITION ÷ CONTRACT_VALUE'],
    P1: ['Cost to date is at least what has already been paid out', 'COST_TO_DATE is lower than AP_PAID_AMOUNT'],
    P2: ['Billed to date does not exceed the contract value', 'BILLED_TO_DATE is more than 105% of CONTRACT_VALUE']
  };

  function apByCategory(job) {
    var map = {};
    job.subs.forEach(function (s) {
      var g = map[s.category] || (map[s.category] = { id: s.category, apPaid: 0, apV: 0, apNV: 0, apOut: 0, apOpen: 0, subs: [] });
      g.apPaid += s.apPaid; g.apV += s.apV; g.apNV += s.apNV; g.apOut += s.apOut; g.apOpen += s.apOpen; g.subs.push(s);
    });
    var rows = [];
    CATEGORIES.forEach(function (c) {
      var g = map[c.id];
      if (!g && !c.requested) return;                       // the four requested categories always show
      if (!g) g = { id: c.id, apPaid: 0, apV: 0, apNV: 0, apOut: 0, apOpen: 0, subs: [] };
      if (!c.requested && g.apPaid === 0 && g.apOut === 0) return; // hide empty extras
      g.label = c.label; g.long = c.long; g.requested = !!c.requested; g.present = !!map[c.id];
      rows.push(g);
    });
    var total = rows.reduce(function (a, g) { return a + g.apPaid; }, 0);
    rows.forEach(function (g) { g.share = total !== 0 ? g.apPaid / total : 0; });
    return { rows: rows, total: total, totalOut: rows.reduce(function (a, g) { return a + g.apOut; }, 0) };
  }

  // ---------- Schedule CSV ----------

  function parseCsv(textIn) {
    var s = String(textIn).replace(/^﻿/, '');
    var rows = [], row = [], field = '', q = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (q) {
        if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && s[i + 1] === '\n') i++;
        row.push(field); field = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else field += ch;
    }
    row.push(field);
    if (row.length > 1 || row[0] !== '') rows.push(row);
    return rows;
  }

  function wholeDollars(v) {
    var t = text(v);
    if (t === '' || t === '-') return 0;
    var neg = /^\(.*\)$/.test(t);
    var n = Number(t.replace(/[$,()\s]/g, ''));
    if (!isFinite(n)) throw EngineError('The schedule file has an amount that is not a number: "' + t + '".');
    return Math.round(neg ? -n : n);
  }

  // Display order: in process first, then not started, then the rolled-up small-jobs line.
  var SECTIONS = [
    { id: 'IN_PROCESS', title: 'Jobs in process' },
    { id: 'NOT_STARTED', title: 'Jobs not started yet' },
    { id: 'SMALL', title: 'Small jobs' }
  ];

  function parseSchedule(csvText) {
    var rows = parseCsv(csvText);
    if (rows.length < 2) throw EngineError('The schedule file is empty.');
    var header = rows[0].map(function (h) { return text(h).toLowerCase(); });
    var col = {};
    header.forEach(function (h, i) { col[h] = i; });
    var periodCols = header.filter(function (h) { return /^underbilling_\d{4}-\d{2}-\d{2}$/.test(h); }).sort();
    var netCol = header.filter(function (h) { return /^net_billing_\d{4}-\d{2}$/.test(h); })[0];
    var need = ['row_type', 'section', 'job_code', 'job_name'];
    var missing = need.filter(function (n) { return !(n in col); });
    if (missing.length || periodCols.length < 1) {
      throw EngineError('This does not look like an Underbilling Schedule file. It needs the columns row_type, section, job_code, job_name and at least one underbilling_YYYY-MM-DD column.');
    }
    var periods = periodCols.map(function (h) { return h.slice('underbilling_'.length); });
    var lines = [], printed = { subtotals: {}, total: null };
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var g = function (name) { return name in col ? text(r[col[name]]) : ''; };
      var type = g('row_type').toLowerCase();
      var amounts = periodCols.map(function (h) { return wholeDollars(r[col[h]]); });
      var section = g('section').toUpperCase();
      if (type === 'subtotal') { printed.subtotals[section] = amounts; continue; }
      if (type === 'total') { printed.total = amounts; continue; }
      if (type !== 'job' && type !== 'small') throw EngineError('Schedule row ' + (i + 1) + ' has an unknown row_type "' + g('row_type') + '".');
      if (!SECTIONS.some(function (s) { return s.id === section; })) throw EngineError('Schedule row ' + (i + 1) + ' has an unknown section "' + g('section') + '".');
      var pct = g('billed_to_date_pct');
      lines.push({
        id: type === 'small' ? 'SMALL' : g('job_code'),
        type: type, section: section, code: g('job_code'), name: g('job_name'),
        am: g('area_manager'), px: g('project_executive'),
        amounts: amounts,
        constStart: isoDate(g('const_start')),
        netBilling: netCol && g(netCol) !== '' ? wholeDollars(g(netCol)) : null,
        billedPct: pct === '' ? null : Number(pct),
        comment: g('comment')
      });
    }
    if (!lines.length) throw EngineError('The schedule file has no job lines.');
    return { periods: periods, netMonth: netCol ? netCol.slice('net_billing_'.length) : null, lines: lines, printed: printed };
  }

  // When no schedule file is supplied, build the schedule from the export alone.
  // Rule S1: underbilled = cost to date - billed to date, listed when above zero.
  // Rule S2: under $10,000 rolls into the small-jobs line.
  // Rule S4: start date after the export date = not started; otherwise in process.
  function scheduleFromExport(exp) {
    var asOf = exp.asOf || new Date().toISOString().slice(0, 10);
    var lines = [], small = 0, threshold = RULES.SMALL_JOB_THRESHOLD_DOLLARS * 100;
    exp.jobs.forEach(function (j) {
      if (j.under <= 0) return;
      if (j.under < threshold) { small += j.under; return; }
      lines.push({
        id: j.code, type: 'job', section: j.startDate && j.startDate > asOf ? 'NOT_STARTED' : 'IN_PROCESS',
        code: j.code, name: j.name, am: j.rd2 || '', px: j.px, amounts: [Math.round(j.under / 100)],
        constStart: j.startDate, netBilling: null,
        billedPct: j.contract > 0 ? Math.round(j.billed / j.contract * 1000) / 10 : null, comment: ''
      });
    });
    lines.sort(function (a, b) { return b.amounts[0] - a.amounts[0]; });
    lines.push({ id: 'SMALL', type: 'small', section: 'SMALL', code: '', name: 'Jobs with underbillings less than $10,000',
      am: '', px: '', amounts: [Math.round(small / 100)], constStart: null, netBilling: null, billedPct: null, comment: '' });
    return { periods: [asOf], netMonth: null, lines: lines, printed: { subtotals: {}, total: null }, fromExport: true };
  }

  // ---------- Report (schedule + export joined) ----------

  function addArrays(a, b) { return a.map(function (v, i) { return v + b[i]; }); }
  function zeros(n) { var z = []; for (var i = 0; i < n; i++) z.push(0); return z; }

  function buildReport(schedule, exp) {
    var n = schedule.periods.length;
    var cur = n - 1, prev = n - 2;
    var asOf = schedule.periods[cur];
    var sections = SECTIONS.map(function (s) {
      var lines = schedule.lines.filter(function (l) { return l.section === s.id; });
      var computed = lines.reduce(function (a, l) { return addArrays(a, l.amounts); }, zeros(n));
      var printed = schedule.printed.subtotals[s.id] || null;
      // Printed lines are rounded to whole dollars, so a subtotal may differ by up to $0.50 per line.
      var tol = Math.max(1, Math.ceil(lines.length / 2));
      var ok = !printed || printed.every(function (p, i) { return Math.abs(p - computed[i]) <= tol; });
      return { id: s.id, title: s.title, lines: lines, computed: computed, subtotal: printed || computed, printed: !!printed, sumsOk: ok, tolerance: tol };
    }).filter(function (s) { return s.lines.length; });

    var totalFromSubs = sections.reduce(function (a, s) { return addArrays(a, s.subtotal); }, zeros(n));
    var total = schedule.printed.total || totalFromSubs;
    var totalOk = total.every(function (t, i) { return t === totalFromSubs[i]; });

    var matched = 0, jobLines = 0;
    schedule.lines.forEach(function (l) {
      l.current = l.amounts[cur];
      l.previous = prev >= 0 ? l.amounts[prev] : null;
      l.change = prev >= 0 ? l.current - l.previous : null;
      l.shareOfTotal = total[cur] !== 0 ? l.current / total[cur] : 0;
      l.excel = l.type === 'job' && exp ? exp.byCode[l.code] || null : null;
      if (l.type === 'job') { jobLines++; if (l.excel) matched++; }
      // Rule S4: section follows the construction start date.
      if (l.type === 'job' && l.constStart) {
        var expected = l.constStart > asOf ? 'NOT_STARTED' : 'IN_PROCESS';
        l.sectionOk = expected === l.section;
      } else l.sectionOk = null;
      // Rule X1: does the export arrive at the schedule's figure?
      if (l.excel) {
        l.excelUnderDollars = Math.max(0, l.excel.under) / 100;
        l.excelAgrees = Math.abs(l.excelUnderDollars - l.current) <= RULES.SCHEDULE_MATCH_TOLERANCE_DOLLARS;
      } else { l.excelUnderDollars = null; l.excelAgrees = null; }
      // Rule X2: Excel cost and billing are shown only when they are plausible and reproduce the line.
      l.excelUsable = !!(l.excel && l.excelAgrees && l.excel.checks.P1 && l.excel.checks.P2);
    });
    var ranked = schedule.lines.filter(function (l) { return l.type === 'job' && l.current > 0; })
      .sort(function (a, b) { return b.current - a.current; });
    ranked.forEach(function (l, i) { l.rank = i + 1; });

    return {
      asOf: asOf, periods: schedule.periods, netMonth: schedule.netMonth, cur: cur, prev: prev,
      sections: sections, total: total, totalOk: totalOk, lines: schedule.lines,
      jobLines: jobLines, matched: matched, rankedCount: ranked.length,
      fromExport: !!schedule.fromExport,
      exportAsOf: exp ? exp.asOf : null,
      exportStale: exp && exp.asOf ? exp.asOf < asOf : null,
      checks: exp ? summarizeChecks(exp, schedule.lines) : null
    };
  }

  function summarizeChecks(exp, lines) {
    var out = [];
    Object.keys(CHECK_TEXT).forEach(function (id) {
      var failed = exp.jobs.filter(function (j) { return !j.checks[id]; });
      out.push({ id: id, kind: id[0] === 'E' ? 'math' : 'plausible', title: CHECK_TEXT[id][0], formula: CHECK_TEXT[id][1],
        total: exp.jobs.length, failed: failed.length, examples: failed.slice(0, 5).map(function (j) { return j.code; }) });
    });
    var jobLines = lines.filter(function (l) { return l.type === 'job'; });
    var withExcel = jobLines.filter(function (l) { return l.excel; });
    out.push({ id: 'X0', kind: 'match', title: 'Every job on the schedule is found in the Excel export', formula: 'Matched on job number',
      total: jobLines.length, failed: jobLines.length - withExcel.length, examples: jobLines.filter(function (l) { return !l.excel; }).slice(0, 5).map(function (l) { return l.code; }) });
    var disagree = withExcel.filter(function (l) { return !l.excelAgrees; });
    out.push({ id: 'X1', kind: 'match', title: 'Excel underbilling agrees with the schedule (within $1)', formula: 'max(0, COST_TO_DATE − BILLED_TO_DATE) vs. the schedule figure',
      total: withExcel.length, failed: disagree.length, examples: disagree.slice(0, 5).map(function (l) { return l.code; }) });
    var sec = jobLines.filter(function (l) { return l.sectionOk !== null; });
    var secBad = sec.filter(function (l) { return !l.sectionOk; });
    out.push({ id: 'S4', kind: 'schedule', title: 'Each job sits in the section its construction start date calls for', formula: 'Start after the report date = not started; otherwise in process',
      total: sec.length, failed: secBad.length, examples: secBad.slice(0, 5).map(function (l) { return l.code; }) });
    return out;
  }

  // What the export still has to supply. Shown on the Data check page and as "Data required" tags.
  var DATA_REQUIRED = [
    { need: 'Cost to date, per job and sub-job', use: 'Underbilling = cost − billed', cmic: 'Job Cost: posted cost transactions', status: 'wrong in export' },
    { need: 'Billed to date and billing for the month', use: 'Underbilling, net billing, billed %', cmic: 'Job Billing: posted invoices', status: 'wrong in export' },
    { need: 'Month-end snapshots of cost and billed', use: 'Prior-month columns and change', cmic: 'Job Cost and Job Billing by period', status: 'not in export' },
    { need: 'Revised contract value', use: 'Billed %', cmic: 'Job Billing contract plus approved change orders', status: 'partly' },
    { need: 'Area manager and CMiC construction start', use: 'Schedule columns and section', cmic: 'Job master and job roles', status: 'not in export' },
    { need: 'CFO adjustments and comments', use: 'Adjusted figure and comment', cmic: 'Not in CMiC: needs an input sheet', status: 'not in export' },
    { need: 'Cost-code category on each AP line', use: 'AP by category: GCs, GRs, Trades, Insurance', cmic: 'Job Cost code and category on the AP distribution', status: 'sub-job name only' },
    { need: 'AP type on each payment', use: 'AP paid by type: payroll, PO, contracts', cmic: 'AP voucher source: subcontract, purchase order, payroll', status: 'not in export' },
    { need: 'Definition of voucher and non-voucher', use: 'AP paid', cmic: 'Query author', status: 'undefined' },
    { need: 'Current AP with extract date and time', use: 'AP and cash position as of today', cmic: 'AP vouchers and payments', status: 'stale snapshot' }
  ];

  BGI.engine = {
    RULES: RULES, DATA_REQUIRED: DATA_REQUIRED, CATEGORIES: CATEGORIES, CATEGORY_BY_ID: CATEGORY_BY_ID, CHECK_TEXT: CHECK_TEXT, SECTIONS: SECTIONS,
    cents: cents, isoDate: isoDate, categorize: categorize, parseExport: parseExport, apByCategory: apByCategory,
    parseCsv: parseCsv, parseSchedule: parseSchedule, scheduleFromExport: scheduleFromExport, buildReport: buildReport
  };
})(typeof globalThis !== 'undefined' ? (globalThis.BGI = globalThis.BGI || {}) : (window.BGI = window.BGI || {}));
