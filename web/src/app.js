/*
 * Interface. Renders three views from the engine's report object:
 *   Schedule               - mirrors the month-end Adjusted Underbillings schedule
 *   How it was calculated  - one job, every figure traced to its source column
 *   Data check             - the rules, re-run against every row of the inputs
 * All text from input files is escaped before it reaches the page.
 */
(function (BGI) {
  'use strict';
  var E = BGI.engine;
  var state = { exp: null, sched: null, report: null, tab: 'schedule', selected: null, q: '', px: '',
    exportName: null, scheduleName: null, exportModified: null, server: false, token: null, loading: false };
  var root, toastTimer;

  // ---------- formatting ----------
  var ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  function esc(s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) { return ESC[c]; }); }
  function usd(d) { var a = Math.abs(Math.round(d)).toLocaleString('en-US'); return d < 0 ? '($' + a + ')' : '$' + a; }
  function usdCell(d) { return d === 0 ? '<span class="muted">–</span>' : usd(d); }
  function usdC(c) {
    var a = Math.abs(c), s = '$' + Math.floor(a / 100).toLocaleString('en-US') + '.' + ('0' + (a % 100)).slice(-2);
    return c < 0 ? '−' + s : s;
  }
  function usdShort(d) {
    var a = Math.abs(d);
    if (a >= 1e6) return '$' + (a / 1e6).toFixed(2) + 'M';
    if (a >= 1e3) return '$' + Math.round(a / 1e3) + 'K';
    return '$' + Math.round(a);
  }
  function pct(x, digits) { return (x * 100).toFixed(digits === undefined ? 1 : digits) + '%'; }
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function parts(iso) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ''); return m ? { y: +m[1], m: +m[2], d: +m[3] } : null; }
  function dLong(iso) { var p = parts(iso); return p ? MONTHS[p.m - 1] + ' ' + p.d + ', ' + p.y : '—'; }
  function dShort(iso) { var p = parts(iso); return p ? ('0' + p.m).slice(-2) + '/' + ('0' + p.d).slice(-2) + '/' + String(p.y).slice(-2) : '—'; }
  function monthName(iso) { var p = parts(iso); return p ? MONTHS[p.m - 1] : ''; }
  function daysBetween(a, b) { return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000); }
  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : many || one + 's'); }

  function displayName(l) {
    if (l.excel && l.excel.name.length > l.name.length && l.excel.name.indexOf(l.name) === 0) return l.excel.name;
    return l.name;
  }

  // ---------- charts ----------
  function donut(segs, size, centerTop, centerBottom) {
    var r = size / 2 - 16, c = 2 * Math.PI * r, cx = size / 2;
    var pos = segs.filter(function (s) { return s.value > 0; });
    var total = pos.reduce(function (a, s) { return a + s.value; }, 0);
    var out = '<svg width="' + size + '" height="' + size + '" viewBox="0 0 ' + size + ' ' + size + '" role="img" aria-label="' + esc(centerBottom + ': ' + centerTop) + '">';
    out += '<circle cx="' + cx + '" cy="' + cx + '" r="' + r + '" fill="none" stroke="#ebeff5" stroke-width="26"/>';
    var start = 0, gap = pos.length > 1 ? 2 : 0;
    pos.forEach(function (s) {
      var len = s.value / total * c;
      var draw = Math.max(0.5, len - gap);
      out += '<circle cx="' + cx + '" cy="' + cx + '" r="' + r + '" fill="none" stroke="' + s.color + '" stroke-width="26" stroke-dasharray="' +
        draw.toFixed(2) + ' ' + (c - draw).toFixed(2) + '" stroke-dashoffset="' + (-start).toFixed(2) + '" transform="rotate(-90 ' + cx + ' ' + cx + ')">' +
        '<title>' + esc(s.label + ': ' + s.text + ' (' + pct(s.value / total) + ')') + '</title></circle>';
      start += len;
    });
    out += '<text x="' + cx + '" y="' + (cx - 2) + '" text-anchor="middle" font-family="Bahnschrift, Segoe UI, sans-serif" font-weight="700" font-size="' + (size > 200 ? 24 : 19) + '" fill="#101b33">' + esc(centerTop) + '</text>';
    out += '<text x="' + cx + '" y="' + (cx + 18) + '" text-anchor="middle" font-size="12" fill="#5a6478">' + esc(centerBottom) + '</text></svg>';
    return out;
  }

  // ---------- shared bits ----------
  function changeCell(l) {
    if (l.change === null) return '';
    if (l.change === 0) return '<span class="chg flat">no change</span>';
    var up = l.change > 0;
    var tag = l.previous === 0 ? ' <span class="tag tag-new">new</span>' : (l.current === 0 ? ' <span class="tag tag-clear">cleared</span>' : '');
    return '<span class="chg ' + (up ? 'up' : 'down') + '"><span aria-hidden="true">' + (up ? '▲' : '▼') + '</span>' +
      '<span class="sr">' + (up ? 'up ' : 'down ') + '</span>' + usd(Math.abs(l.change)) + '</span>' + tag;
  }

  var SECTION_COLOR = { IN_PROCESS: 'var(--s1)', NOT_STARTED: 'var(--s2)', SMALL: 'var(--s3)' };
  var SECTION_HEX = { IN_PROCESS: '#2a78d6', NOT_STARTED: '#eb6834', SMALL: '#1baf7a' };

  // ---------- view: schedule ----------
  function matchesFilter(l) {
    if (state.px && l.px !== state.px) return false;
    if (!state.q) return true;
    var q = state.q.toLowerCase();
    return (displayName(l) + ' ' + l.code + ' ' + l.px + ' ' + l.am).toLowerCase().indexOf(q) >= 0;
  }

  function sourceLine(rep) {
    var h = '<p class="source-line">';
    if (rep.fromExport) {
      h += '<span><span class="pill pill-warn">Excel only</span> No schedule file is loaded, so this page is calculated from the Excel export alone.</span>';
    } else {
      h += '<span>Schedule figures from <b>' + esc(state.scheduleName || 'the schedule file') + '</b></span>';
    }
    if (state.exp) {
      h += '<span>Job detail from <b>' + esc(state.exportName || 'the Excel export') + '</b>, data as of ' + dLong(rep.exportAsOf) + '</span>';
      if (!rep.fromExport && rep.exportStale) h += '<span class="pill pill-warn">Excel is ' + plural(daysBetween(rep.exportAsOf, rep.asOf), 'day') + ' older than the schedule</span>';
      var x1 = rep.checks.filter(function (c) { return c.id === 'X1'; })[0];
      if (!rep.fromExport && x1 && x1.failed) {
        h += '<button class="linkish" data-action="tab" data-tab="check">Excel supports ' + (x1.total - x1.failed) + ' of ' + x1.total + ' schedule figures. See why</button>';
      }
    } else {
      h += '<span class="pill pill-warn">No Excel export loaded</span><span>Drop the UNDERBILLING_CASHPOSITION export on this window to see how each line was calculated.</span>';
    }
    return h + '</p>';
  }

  function hero(rep) {
    var cur = rep.cur, prev = rep.prev, total = rep.total[cur];
    var h = '<section class="hero" aria-label="Summary">';
    // headline
    h += '<div class="panel headline"><div class="asof">As of ' + dLong(rep.asOf) + '</div>' +
      '<div class="big num">' + usd(total) + '</div><div class="what">' + (rep.fromExport ? 'total underbillings in the Excel export' : 'total adjusted underbillings') + '</div>';
    if (prev >= 0) {
      var d = total - rep.total[prev], base = rep.total[prev];
      h += '<div class="delta">' + (d === 0 ? 'No change from ' + monthName(rep.periods[prev]) + '.' :
        '<b>' + (d > 0 ? 'Up ' : 'Down ') + usd(Math.abs(d)) + '</b>' + (base ? ' (' + pct(Math.abs(d) / base) + ')' : '') + ' from ' + monthName(rep.periods[prev]) + '.') + '</div>';
      var max = Math.max.apply(null, rep.total.map(Math.abs)) || 1;
      h += '<div class="trend" aria-label="Total by month">';
      rep.periods.forEach(function (p, i) {
        h += '<div class="col' + (i === cur ? ' now' : '') + '"><div class="val">' + usd(rep.total[i]) + '</div><div class="barwrap"><div class="b" style="height:' +
          Math.max(4, Math.round(Math.abs(rep.total[i]) / max * 100)) + '%"></div></div><div class="lab">' + dShort(p) + '</div></div>';
      });
      h += '</div>';
    }
    h += '</div>';
    // donut
    var order = ['IN_PROCESS', 'NOT_STARTED', 'SMALL'];
    var secs = order.map(function (id) { return rep.sections.filter(function (s) { return s.id === id; })[0]; }).filter(Boolean);
    h += '<div class="panel"><h2>Where the underbilling sits</h2><div class="donut-wrap">' +
      donut(secs.map(function (s) { return { label: s.title, value: s.subtotal[cur], text: usd(s.subtotal[cur]), color: SECTION_HEX[s.id] }; }), 190, usdShort(total), 'total') +
      '<ul class="legend">';
    secs.forEach(function (s) {
      var n = s.id === 'SMALL' ? 'one rolled-up line' : plural(s.lines.filter(function (l) { return l.current > 0; }).length, 'job') + ' with a balance';
      h += '<li><button class="rowbtn" data-action="jump" data-id="sec-' + s.id + '"><span class="sw" style="background:' + SECTION_COLOR[s.id] + '"></span>' +
        '<span class="nm">' + esc(s.title) + '<small>' + n + '</small></span><span class="v">' + usd(s.subtotal[cur]) +
        '<small>' + (total ? pct(s.subtotal[cur] / total) : '') + '</small></span></button></li>';
    });
    h += '</ul></div></div>';
    // movers / top jobs
    var jobs = rep.lines.filter(function (l) { return l.type === 'job'; });
    h += '<div class="panel">';
    if (prev >= 0) {
      var ups = jobs.filter(function (l) { return l.change > 0; }).sort(function (a, b) { return b.change - a.change; }).slice(0, 4);
      var downs = jobs.filter(function (l) { return l.change < 0; }).sort(function (a, b) { return a.change - b.change; }).slice(0, 4);
      h += '<h2>What changed since ' + monthName(rep.periods[prev]) + '</h2><div class="movers">';
      [['Went up the most', ups, 'up', '▲'], ['Came down the most', downs, 'down', '▼']].forEach(function (g) {
        h += '<div><h3>' + g[0] + '</h3><ol>';
        g[1].forEach(function (l) {
          h += '<li><button data-action="open" data-id="' + esc(l.id) + '" title="See how it was calculated"><span class="n">' + esc(displayName(l)) +
            '</span><span class="d ' + g[2] + '"><span aria-hidden="true">' + g[3] + '</span> ' + usd(Math.abs(l.change)) + '</span></button></li>';
        });
        if (!g[1].length) h += '<li class="none">None</li>';
        h += '</ol></div>';
      });
      h += '</div>';
    } else {
      var top = jobs.slice().sort(function (a, b) { return b.current - a.current; }).slice(0, 6);
      h += '<h2>Largest underbillings</h2><div class="movers" style="grid-template-columns:1fr"><div><ol>';
      top.forEach(function (l) {
        h += '<li><button data-action="open" data-id="' + esc(l.id) + '"><span class="n">' + esc(displayName(l)) + '</span><span class="d">' + usd(l.current) + '</span></button></li>';
      });
      if (!top.length) h += '<li class="none">No job is underbilled by $10,000 or more in this export.</li>';
      h += '</ol></div></div>';
    }
    h += '</div></section>';
    return h;
  }

  function scheduleView(rep) {
    var cur = rep.cur, prev = rep.prev, hasPrev = prev >= 0;
    var h = sourceLine(rep) + hero(rep);
    // filters
    var pxs = {};
    rep.lines.forEach(function (l) { if (l.type === 'job' && l.px) pxs[l.px] = (pxs[l.px] || 0) + 1; });
    h += '<div class="filters"><label class="sr" for="q">Find a job</label><input type="search" id="q" placeholder="Find a job by name or number" value="' + esc(state.q) + '" autocomplete="off">' +
      '<label class="sr" for="px">Project executive</label><select id="px"><option value="">All project executives</option>';
    Object.keys(pxs).sort().forEach(function (p) { h += '<option value="' + esc(p) + '"' + (p === state.px ? ' selected' : '') + '>' + esc(p) + ' (' + pxs[p] + ')</option>'; });
    h += '</select>';
    if (state.q || state.px) h += '<button class="btn" data-action="clear-filter">Show all jobs</button>';
    h += '<span class="muted">Click any line to see how it was calculated.</span></div>';
    var filtering = !!(state.q || state.px);
    if (filtering) {
      var vis = rep.lines.filter(function (l) { return l.type === 'job' && matchesFilter(l); });
      var sum = vis.reduce(function (a, l) { return a + l.current; }, 0);
      h += '<p class="filter-note"><b>' + plural(vis.length, 'job') + (state.px ? ' for ' + esc(state.px) : ' found') + ':</b> ' + usd(sum) + ' underbilled, ' +
        (rep.total[cur] ? pct(sum / rep.total[cur]) : '0%') + ' of the ' + usd(rep.total[cur]) + ' company total.' +
        (hasPrev ? ' ' + (function () { var ch = vis.reduce(function (a, l) { return a + l.change; }, 0); return ch === 0 ? 'No change' : (ch > 0 ? 'Up ' : 'Down ') + usd(Math.abs(ch)); })() + ' since ' + monthName(rep.periods[prev]) + '.' : '') + '</p>';
    }
    // sections
    rep.sections.forEach(function (s) {
      var lines = s.lines.filter(function (l) { return filtering ? l.type === 'job' && matchesFilter(l) : true; });
      if (filtering && !lines.length) return;
      var jobCount = s.lines.filter(function (l) { return l.type === 'job'; }).length;
      var shown = filtering ? lines.reduce(function (a, l) { return a + l.current; }, 0) : s.subtotal[cur];
      h += '<section class="section" id="sec-' + s.id + '"><div class="section-head"><h2>' + esc(s.title) + '</h2>' +
        (jobCount ? '<span class="count">' + (filtering ? lines.length + ' of ' : '') + plural(jobCount, 'job') + '</span>' : '') + '<span class="blurb">' + esc(s.blurb) + '</span>' +
        '<span class="sub num">' + usd(shown) + '</span></div><div class="tablewrap"><table class="sched"><thead><tr>' +
        '<th>Job</th><th>Area manager</th><th>Project executive</th>';
      rep.periods.forEach(function (p, i) { h += '<th class="r' + (i === cur ? ' now' : '') + '">' + dShort(p) + '</th>'; });
      if (hasPrev) h += '<th class="r">Change since ' + monthName(rep.periods[prev]).slice(0, 3) + '</th>';
      h += '<th>Const. start</th>' + (rep.netMonth ? '<th class="r">Net ' + MONTHS[+rep.netMonth.slice(5) - 1].slice(0, 3) + ' billing</th>' : '') +
        '<th class="r">Billed to date</th><th>Comment</th><th class="seecol"><span class="sr">Open</span></th></tr></thead><tbody>';
      lines.forEach(function (l) {
        h += '<tr class="line" tabindex="0" role="button" data-action="open" data-id="' + esc(l.id) + '" aria-label="' + esc(displayName(l)) + ', see how it was calculated">';
        h += '<td class="job"><b>' + esc(displayName(l)) + '</b>' + (l.code ? '<small>' + esc(l.code) + '</small>' : '') + '</td>';
        h += '<td>' + esc(l.am) + '</td><td>' + esc(l.px) + '</td>';
        l.amounts.forEach(function (a, i) { h += '<td class="r num ' + (i === cur ? 'now' : 'prior') + '">' + usdCell(a) + '</td>'; });
        if (hasPrev) h += '<td class="r">' + changeCell(l) + '</td>';
        h += '<td class="num">' + (l.constStart ? dShort(l.constStart) : '') + '</td>';
        if (rep.netMonth) h += '<td class="r num">' + (l.netBilling === null ? '' : usdCell(l.netBilling)) + '</td>';
        h += '<td class="r">' + (l.billedPct === null ? '' : '<span class="pct"><span class="track"><span class="fill" style="width:' + Math.max(0, Math.min(100, l.billedPct)) + '%"></span></span><span class="num">' + l.billedPct.toFixed(1) + '%</span></span>') + '</td>';
        h += '<td class="comment">' + esc(l.comment) + '</td><td class="seecol"><span class="see">See the math ›</span></td></tr>';
      });
      if (!filtering) {
        h += '<tr class="subtotal"><td colspan="3">Subtotal</td>';
        s.subtotal.forEach(function (a, i) { h += '<td class="r num' + (i === cur ? ' now' : '') + '">' + usd(a) + '</td>'; });
        h += '<td colspan="' + ((hasPrev ? 1 : 0) + (rep.netMonth ? 1 : 0) + 4) + '">' + (hasPrev ? changeCell({ change: s.subtotal[cur] - s.subtotal[prev], previous: 1, current: 1 }) : '') + '</td></tr>';
      }
      h += '</tbody></table></div></section>';
    });
    // grand total
    h += '<div class="grand"><span class="t">' + (rep.fromExport ? 'Total underbillings' : 'Total adjusted underbillings') + '</span>';
    rep.periods.forEach(function (p, i) { h += '<span class="p' + (i === cur ? ' now' : '') + '"><small>' + dShort(p) + '</small><b>' + usd(rep.total[i]) + '</b></span>'; });
    h += '</div>';
    return h;
  }

  // ---------- view: how it was calculated ----------
  function term(label, value, column, cls) {
    return '<div class="term ' + (cls || '') + '"><span class="l">' + esc(label) + '</span><span class="v">' + value + '</span>' + (column ? '<span class="c">' + esc(column) + '</span>' : '') + '</div>';
  }
  function op(s) { return '<span class="op" aria-hidden="true">' + s + '</span>'; }
  function eq(title, body, note) { return '<div class="eq"><div class="eq-title">' + esc(title) + '</div><div class="eq-row">' + body + '</div>' + (note ? '<div class="eq-note">' + note + '</div>' : '') + '</div>'; }

  var AP_COLOR = { GC: '#2a78d6', GR: '#eb6834', GCGR: '#1baf7a', TRADES: '#eda100', INS: '#e87ba4', OTHERS: '#6250d6', MAIN: '#9aa3b2' };
  function apColorKey(id) { return AP_COLOR[id] ? id : 'OTHERS'; }

  function picker(rep) {
    var ids = rep.lines.map(function (l) { return l.id; });
    var i = ids.indexOf(state.selected);
    var h = '<div class="picker"><button class="btn" data-action="tab" data-tab="schedule">‹ Back to the schedule</button>' +
      '<label class="sr" for="jobpick">Job</label><select id="jobpick">';
    rep.sections.forEach(function (s) {
      h += '<optgroup label="' + esc(s.title) + '">';
      s.lines.forEach(function (l) { h += '<option value="' + esc(l.id) + '"' + (l.id === state.selected ? ' selected' : '') + '>' + esc(displayName(l)) + (l.code ? ' (' + esc(l.code) + ')' : '') + ' ' + usd(l.current) + '</option>'; });
      h += '</optgroup>';
    });
    h += '</select><button class="btn" data-action="open" data-id="' + esc(ids[(i - 1 + ids.length) % ids.length]) + '">‹ Previous job</button>' +
      '<button class="btn" data-action="open" data-id="' + esc(ids[(i + 1) % ids.length]) + '">Next job ›</button></div>';
    return h;
  }

  function scheduleBlock(rep, l) {
    var cur = rep.cur, prev = rep.prev, mon = monthName(rep.asOf);
    var h = '<div class="panel"><div class="step"><h2>1. The number on the schedule</h2><span class="pill pill-src src">' +
      (rep.fromExport ? 'Calculated from the Excel export' : 'From the ' + dShort(rep.asOf) + ' schedule') + '</span></div><div class="walk">';
    rep.periods.forEach(function (p, i) { h += '<div class="cell' + (i === cur ? ' now' : '') + '"><small>' + dLong(p) + '</small><b>' + usd(l.amounts[i]) + '</b></div>'; });
    h += '</div><ul class="plain">';
    if (prev >= 0) {
      var pm = monthName(rep.periods[prev]);
      if (l.change === 0) h += '<li>No change since ' + pm + '.</li>';
      else if (l.previous === 0) h += '<li><b class="up">New this month.</b> There was no underbilling in ' + pm + '.</li>';
      else if (l.current === 0) h += '<li><b class="down">Cleared.</b> The ' + usd(l.previous) + ' from ' + pm + ' is gone.</li>';
      else h += '<li>Underbilling went <b class="' + (l.change > 0 ? 'up">up ' : 'down">down ') + usd(Math.abs(l.change)) + '</b> (' + pct(Math.abs(l.change) / Math.abs(l.previous)) + ') since ' + pm + '.</li>';
    }
    if (l.netBilling !== null) {
      h += l.netBilling === 0 ? '<li>Nothing was billed in ' + mon + '.</li>' :
        (l.netBilling > 0 ? '<li>' + usd(l.netBilling) + ' was billed in ' + mon + '.</li>' : '<li>Billing was a net credit of ' + usd(Math.abs(l.netBilling)) + ' in ' + mon + '.</li>');
    }
    if (l.billedPct !== null) h += '<li>' + l.billedPct.toFixed(1) + '% of the contract has been billed to date.</li>';
    if (l.current > 0 && l.rank) h += '<li>This is ' + pct(l.shareOfTotal) + ' of the company total, number ' + l.rank + ' of ' + rep.rankedCount + ' jobs with a balance.</li>';
    if (l.constStart) {
      h += '<li>Listed under <b>' + esc(sectionTitle(l.section)) + '</b> because construction ' + (l.constStart > rep.asOf ? 'starts ' : 'started ') + dShort(l.constStart) + ', ' +
        (l.constStart > rep.asOf ? 'after' : 'on or before') + ' the ' + dShort(rep.asOf) + ' report date.</li>';
    }
    if (l.comment) h += '<li>Note on the schedule: “' + esc(l.comment) + '”</li>';
    h += '</ul>';
    if (prev >= 0) {
      h += eq('Change since ' + monthName(rep.periods[prev]), term(dShort(rep.asOf), usd(l.current)) + op('−') + term(dShort(rep.periods[prev]), usd(l.previous)) + op('=') +
        term('Change', (l.change > 0 ? '+' : '') + usd(l.change), null, 'result'));
    }
    return h + '</div>';
  }

  function sectionTitle(id) { var s = E.SECTIONS.filter(function (x) { return x.id === id; })[0]; return s ? s.title : id; }

  function excelBlock(rep, l) {
    var j = l.excel;
    var h = '<div class="panel"><div class="step"><h2>2. What the Excel export calculates</h2><span class="pill pill-src src">Excel, as of ' + dShort(rep.exportAsOf) + '</span></div>';
    if (!rep.fromExport) {
      if (l.excelAgrees) {
        h += '<div class="callout callout-ok"><b>The Excel export agrees with the schedule.</b>Cost to date minus billed to date comes to ' + usd(l.excelUnderDollars) + ', the same as the schedule.</div>';
      } else {
        var why = [];
        if (!j.checks.P1) why.push('Cost to date is ' + usdC(j.cost) + ' in the export, but ' + usdC(j.apPaid) + ' has already been paid out on this job. Cost is missing from the export.');
        if (!j.checks.P2) why.push('Billed to date is ' + usdC(j.billed) + ' against a ' + usdC(j.contract) + ' contract (' + (j.billed / j.contract).toFixed(1) + ' times the contract). Billing is overstated in the export.');
        if (rep.exportStale) why.push('The export is as of ' + dLong(rep.exportAsOf) + ', ' + plural(daysBetween(rep.exportAsOf, rep.asOf), 'day') + ' before the report date.');
        if (!why.length || (j.checks.P1 && j.checks.P2)) why.push('The schedule may also carry manual adjustments that are not in the export.');
        h += '<div class="callout callout-warn"><b>The Excel export does not support the schedule figure yet.</b>The schedule says ' + usd(l.current) + '. The export works out to ' +
          (j.under > 0 ? usdC(j.under) + ' underbilled' : (j.under < 0 ? usdC(-j.under) + ' overbilled' : 'even')) + '.<ul>' + why.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul></div>';
      }
    }
    h += eq('Underbilled amount', term('Cost to date', usdC(j.cost), 'COST_TO_DATE') + op('−') + term('Billed to date', usdC(j.billed), 'BILLED_TO_DATE') + op('=') +
      term(j.under >= 0 ? 'Underbilled' : 'Overbilled', usdC(j.under), 'UNDERBILLED_AMOUNT', 'result'),
      'A positive result means cost is ahead of billing (underbilled). Negative means billing is ahead of cost (overbilled).' +
      (j.checks.E1 ? '' : ' <span class="off">The export\'s own UNDERBILLED_AMOUNT column says ' + usdC(j.underReported) + '.</span>'));
    h += '<ul class="facts"><li>Contract value <b>' + (j.hasContract ? usdC(j.contract) : 'not in export') + '</b></li><li>Excel project executive <b>' + esc(j.px || '—') + '</b></li>' +
      '<li>Excel start date <b>' + dShort(j.startDate) + '</b></li><li>Last billing posted <b>' + dShort(j.lastBillPost) + '</b></li></ul>';
    return h + '</div>';
  }

  function cashBlock(l) {
    var j = l.excel;
    var h = '<div class="panel"><div class="step"><h2>3. Cash position</h2><span class="pill pill-src src">Excel</span></div>';
    h += eq('Cash position', term('Collected from owner', usdC(j.arPaid), 'AR_PAID_AMOUNT') + op('−') + term('Paid out', usdC(j.apPaid), 'AP_PAID_AMOUNT') + op('=') +
      term(j.cash >= 0 ? 'Cash ahead' : 'Cash behind', usdC(j.cash), 'CASH_POSITION', 'result'),
      j.cash >= 0 ? 'More has been collected from the owner than paid out.' : 'More has been paid out than collected from the owner.');
    if (j.cashPct !== null) {
      h += eq('Cash position as a share of the contract', term('Cash position', usdC(j.cash), 'CASH_POSITION') + op('÷') + term('Contract value', usdC(j.contract), 'CONTRACT_VALUE') + op('=') +
        term('Share of contract', pct(j.cashPct), 'CASH_POSITION_PERCENT', 'result'));
    }
    h += eq('Billing and collections', term('Collected', usdC(j.arPaid), 'AR_PAID_AMOUNT') + op('+') + term('Still owed by owner', usdC(j.arOut), 'AR_OUTSTANDING_AMOUNT') + op('=') +
      term('Net billed', usdC(j.arPaid + j.arOut), 'NET_BILL_AMOUNT', 'result'),
      (j.arStatus ? 'Owner payment status: ' + esc(j.arStatus) + '. ' : '') + 'Last payment posted ' + dShort(j.lastArPost) + '.' +
      (j.checks.E5 ? '' : ' <span class="off">The export\'s NET_BILL_AMOUNT column says ' + usdC(j.netBill) + ', a difference of ' + usdC(Math.abs(j.netBill - j.arPaid - j.arOut)) + '.</span>'));
    return h + '</div>';
  }

  function apBlock(l) {
    var j = l.excel, ap = E.apByCategory(j);
    var h = '<div class="panel wide"><div class="step"><h2>5. AP by category</h2><span class="pill pill-src src">Excel sub-jobs</span></div>';
    var buckets = {};
    ap.rows.forEach(function (g) { var k = apColorKey(g.id); buckets[k] = (buckets[k] || 0) + Math.max(0, g.apPaid); });
    var labels = { GC: 'GCs', GR: 'GRs', GCGR: 'GCs / GRs combined', TRADES: 'Trades', INS: 'Insurance', OTHERS: 'Other sub-jobs', MAIN: 'Main job number' };
    var segs = ['GC', 'GR', 'GCGR', 'TRADES', 'INS', 'OTHERS', 'MAIN'].filter(function (k) { return buckets[k] > 0; })
      .map(function (k) { return { label: labels[k], value: buckets[k], text: usdC(buckets[k]), color: AP_COLOR[k] }; });
    h += '<div class="donut-wrap" style="align-items:flex-start">';
    if (segs.length) h += donut(segs, 220, usdShort(ap.total / 100), 'AP paid');
    h += '<div style="flex:1;min-width:0"><div class="tablewrap"><table class="mini"><thead><tr><th>Category</th><th class="r">AP paid</th><th class="r">Share</th><th class="r">By voucher</th><th class="r">Not by voucher</th><th class="r">AP outstanding</th><th class="r">Open invoices</th></tr></thead><tbody>';
    ap.rows.forEach(function (g) {
      var empty = !g.present;
      h += '<tr class="' + (g.requested ? 'asked' : '') + '"><td title="' + esc(g.long) + '"><span class="sw-i" style="background:' + AP_COLOR[apColorKey(g.id)] + '"></span>' + esc(g.label) +
        (empty ? ' <span class="none">no sub-job for this in the export</span>' : '') + '</td>' +
        '<td class="r num">' + (empty ? '–' : usdC(g.apPaid)) + '</td><td class="r num">' + (empty || !ap.total ? '–' : pct(g.share)) + '</td>' +
        '<td class="r num">' + (empty ? '–' : usdC(g.apV)) + '</td><td class="r num">' + (empty ? '–' : usdC(g.apNV)) + '</td>' +
        '<td class="r num">' + (empty ? '–' : usdC(g.apOut)) + '</td><td class="r num">' + (empty ? '–' : g.apOpen) + '</td></tr>';
    });
    h += '<tr class="total"><td>Total</td><td class="r num">' + usdC(ap.total) + '</td><td class="r num">' + (ap.total ? '100.0%' : '–') + '</td><td class="r num">' + usdC(j.apV) +
      '</td><td class="r num">' + usdC(j.apNV) + '</td><td class="r num">' + usdC(ap.totalOut) + '</td><td class="r num">' + j.apOpen + '</td></tr></tbody></table></div>';
    var main = ap.rows.filter(function (g) { return g.id === 'MAIN'; })[0];
    if (main && ap.total > 0 && main.apPaid / ap.total >= 0.5) {
      h += '<div class="callout callout-warn" style="margin-top:12px"><b>' + pct(main.apPaid / ap.total, 0) + ' of this job\'s AP is posted to the main job number.</b>The export only tells categories apart by sub-job, so that money cannot be placed in GCs, GRs, Trades or Insurance from this file.</div>';
    }
    h += '</div></div>';
    h += '<details class="subs"><summary>Show all ' + plural(j.subs.length, 'sub-job') + ' behind these totals</summary><div class="tablewrap"><table class="mini"><thead><tr><th>Sub-job</th><th>Name</th><th>Category</th><th class="r">AP paid</th><th class="r">AP outstanding</th><th class="r">Net billed</th><th class="r">Collected</th><th class="r">Cost to date</th></tr></thead><tbody>';
    j.subs.forEach(function (s) {
      h += '<tr><td class="num">' + esc(s.code) + '</td><td>' + esc(s.name) + '</td><td>' + esc(E.CATEGORY_BY_ID[s.category].label) + '</td><td class="r num">' + usdC(s.apPaid) +
        '</td><td class="r num">' + usdC(s.apOut) + '</td><td class="r num">' + usdC(s.netBill) + '</td><td class="r num">' + usdC(s.arPaid) + '</td><td class="r num">' + usdC(s.cost) + '</td></tr>';
    });
    h += '</tbody></table></div></details></div>';
    return h;
  }

  function calcView(rep) {
    var l = rep.lines.filter(function (x) { return x.id === state.selected; })[0];
    if (!l) {
      var top = rep.lines.filter(function (x) { return x.type === 'job'; }).sort(function (a, b) { return b.current - a.current; }).slice(0, 8);
      var h0 = '<div class="panel"><h2>Pick a job to see how its number was calculated</h2><p class="muted">Click any line on the Schedule page, or start with one of the largest:</p><table class="mini"><tbody>';
      top.forEach(function (x) { h0 += '<tr class="clickable" tabindex="0" role="button" data-action="open" data-id="' + esc(x.id) + '"><td>' + esc(displayName(x)) + ' <span class="muted num">' + esc(x.code) + '</span></td><td class="r num"><b>' + usd(x.current) + '</b></td></tr>'; });
      return h0 + '</tbody></table></div>';
    }
    var h = picker(rep);
    h += '<div class="jobhead"><div><h1>' + esc(displayName(l)) + '</h1><ul class="facts">' + (l.code ? '<li>Job <b class="num">' + esc(l.code) + '</b></li>' : '') +
      (l.am ? '<li>Area manager <b>' + esc(l.am) + '</b></li>' : '') + (l.px ? '<li>Project executive <b>' + esc(l.px) + '</b></li>' : '') +
      (l.excel && l.excel.name !== displayName(l) ? '<li>Name in Excel <b>' + esc(l.excel.name) + '</b></li>' : '') + '</ul></div></div>';
    h += '<div class="calc-grid">';
    if (l.type === 'small') {
      h += '<div class="wide">' + scheduleBlock(rep, l) + '</div><div class="panel wide"><h2>What is in this line</h2><p class="plain">' +
        (rep.fromExport ? 'Every job in the Excel export that is underbilled by less than $10,000, added together.' :
          'The schedule rolls every job underbilled by less than $10,000 into this one line. It does not list the jobs behind it, so there is nothing further to open.') + '</p></div>';
      return h + '</div>';
    }
    h += scheduleBlock(rep, l);
    if (!state.exp) h += '<div class="panel"><h2>2. What the Excel export calculates</h2><p class="plain">No Excel export is loaded. Drop the UNDERBILLING_CASHPOSITION export on this window to fill in this page.</p></div>';
    else if (!l.excel) h += '<div class="panel"><h2>2. What the Excel export calculates</h2><div class="callout callout-warn"><b>Job ' + esc(l.code) + ' is not in the Excel export.</b>There is no cost, billing, cash or AP detail to show for it.</div></div>';
    else h += excelBlock(rep, l) + cashBlock(l) + '<div class="panel"><div class="step"><h2>4. AP paid</h2><span class="pill pill-src src">Excel</span></div>' +
      eq('How AP was paid', term('By voucher', usdC(l.excel.apV), 'PAID_AMOUNT_VOUCHER') + op('+') + term('Not by voucher', usdC(l.excel.apNV), 'PAID_AMOUNT_NON_VOUCHER') + op('=') + term('AP paid', usdC(l.excel.apV + l.excel.apNV), 'AP_PAID_AMOUNT', 'result')) +
      eq('Still to pay', term('AP outstanding', usdC(l.excel.apOut), 'AP_OUTSTANDING_AMOUNT') + op('across') + term('Open invoices', String(l.excel.apOpen), 'AP_OPEN_COUNT')) + '</div>' + apBlock(l);
    return h + '</div>';
  }

  // ---------- view: data check ----------
  function checkView(rep) {
    var h = '<div class="cards"><div class="panel"><h2>Files in this report</h2>';
    h += '<p><b>Schedule:</b> ' + (rep.fromExport ? 'none loaded, page 1 is calculated from the Excel export' : esc(state.scheduleName) + ', ' + plural(rep.lines.length, 'line') + ', as of ' + dLong(rep.asOf)) + '</p>';
    h += '<p><b>Excel export:</b> ' + (state.exp ? esc(state.exportName) + ', ' + plural(state.exp.jobs.length, 'job') + ', data as of ' + dLong(rep.exportAsOf) : 'none loaded') + '</p>';
    if (state.exp && state.exp.problems.length) h += '<p class="up">' + esc(state.exp.problems.slice(0, 3).join(' ')) + '</p>';
    h += '</div><div class="panel"><h2>Does the schedule add up?</h2>';
    rep.sections.forEach(function (s) {
      h += '<p><span class="status ' + (s.sumsOk ? 'pass">✓ Yes' : 'fail">✕ No') + '</span> ' + esc(s.title) + ': lines add to ' + usd(s.computed[rep.cur]) + ', subtotal ' + usd(s.subtotal[rep.cur]) +
        (s.printed && s.computed[rep.cur] !== s.subtotal[rep.cur] ? ' <span class="muted">(whole-dollar rounding)</span>' : '') + '</p>';
    });
    h += '<p><span class="status ' + (rep.totalOk ? 'pass">✓ Yes' : 'fail">✕ No') + '</span> Subtotals add to the ' + usd(rep.total[rep.cur]) + ' total</p></div></div>';
    if (!rep.checks) return h + '<div class="panel"><p class="plain">Load the Excel export to run the checks on it.</p></div>';
    h += '<section class="section"><div class="section-head"><h2>Rules re-run on every row</h2><span class="blurb">The report recalculates each formula itself and compares the answer with the file.</span></div><div class="tablewrap"><table class="sched"><thead><tr><th>Rule</th><th>Formula</th><th class="r">Rows checked</th><th class="r">Rows that fail</th><th>Result</th><th>Examples</th></tr></thead><tbody>';
    rep.checks.forEach(function (c) {
      h += '<tr><td><b>' + esc(c.title) + '</b></td><td><code class="f">' + esc(c.formula) + '</code></td><td class="r num">' + c.total + '</td><td class="r num">' + c.failed + '</td>' +
        '<td><span class="status ' + (c.failed ? 'fail">✕ ' + (c.kind === 'math' ? 'Does not add up' : 'Needs attention') : 'pass">✓ Holds') + '</span></td><td class="num muted">' + esc(c.examples.join(', ')) + '</td></tr>';
    });
    h += '</tbody></table></div></section>';
    if (!rep.fromExport) {
      var rows = rep.lines.filter(function (l) { return l.type === 'job'; });
      h += '<section class="section"><div class="section-head"><h2>Schedule against Excel, job by job</h2><span class="blurb">Click a job to open its calculation.</span></div><div class="tablewrap"><table class="sched"><thead><tr><th>Job</th><th class="r now">Schedule ' + dShort(rep.asOf) +
        '</th><th class="r">Excel cost to date</th><th class="r">Excel billed to date</th><th class="r">Excel underbilled</th><th>Agrees?</th><th>What is wrong in the export</th></tr></thead><tbody>';
      rows.forEach(function (l) {
        var j = l.excel, flags = [];
        if (j) { if (!j.checks.P1) flags.push('cost lower than AP paid'); if (!j.checks.P2) flags.push('billing above contract'); }
        h += '<tr class="line" tabindex="0" role="button" data-action="open" data-id="' + esc(l.id) + '"><td class="job"><b>' + esc(displayName(l)) + '</b><small>' + esc(l.code) + '</small></td><td class="r num now">' + usd(l.current) + '</td>' +
          (j ? '<td class="r num">' + usdC(j.cost) + '</td><td class="r num">' + usdC(j.billed) + '</td><td class="r num">' + (j.under > 0 ? usdC(j.under) : '<span class="muted">none</span>') + '</td>' +
            '<td><span class="status ' + (l.excelAgrees ? 'pass">✓ Yes' : 'fail">✕ No') + '</span></td><td class="muted">' + esc(flags.join(', ')) + '</td>'
            : '<td colspan="5" class="muted">Not in the Excel export</td>') + '</tr>';
      });
      h += '</tbody></table></div></section>';
    }
    return h;
  }

  // ---------- shell ----------
  function emptyView() {
    return '<div class="drop"><h1>Drop this month\'s files here</h1><p>The report builds itself from two files. Drag them onto this window together, or one at a time.</p>' +
      '<ul class="need"><li class="' + (state.exp ? 'have' : '') + '">The Excel export, UNDERBILLING_CASHPOSITION_RPT.xlsx</li><li class="' + (state.sched ? 'have' : '') + '">The month-end schedule file, for example 10-31-2026 Underbilling Schedule.csv</li></ul>' +
      '<div><button class="btn btn-primary" data-action="browse">Choose files</button></div><p style="margin-top:18px;font-size:.86rem">Nothing leaves this computer. The files are read here and never uploaded.</p></div>';
  }

  function render() {
    var rep = state.report;
    var h = '<header class="bar"><div class="bar-in"><div class="brand"><b>AJAX BUILDING COMPANY</b><span>' + (rep && rep.fromExport ? 'Underbillings' : 'Adjusted underbillings') + (rep ? ', ' + dLong(rep.asOf) : '') + '</span></div>';
    if (rep) {
      h += '<nav class="tabs" role="tablist">';
      [['schedule', 'Schedule'], ['calc', 'How it was calculated'], ['check', 'Data check']].forEach(function (t) {
        h += '<button class="tab" role="tab" aria-selected="' + (state.tab === t[0]) + '" data-action="tab" data-tab="' + t[0] + '">' + t[1] + '</button>';
      });
      h += '</nav>';
    }
    h += '<div class="bar-actions">' + (rep ? '<button class="btn" data-action="print">Print or save as PDF</button>' : '') + '<button class="btn" data-action="browse">Load files</button></div></div></header><main id="main">';
    if (state.loading) h += '<div class="drop"><h1>Reading the files…</h1></div>';
    else if (!rep) h += emptyView();
    else if (state.tab === 'calc') h += calcView(rep);
    else if (state.tab === 'check') h += checkView(rep);
    else h += scheduleView(rep);
    h += '</main><footer class="foot">Every figure is read from the input files or produced by a fixed formula. Nothing is estimated and nothing leaves this computer. BGI Underbilling Report ' + esc(BGI.version || 'dev') + '</footer>';
    var focusId = document.activeElement && document.activeElement.id;
    var caret = focusId === 'q' ? document.activeElement.selectionStart : null;
    root.innerHTML = h;
    if (focusId === 'q') { var q = document.getElementById('q'); if (q) { q.focus(); try { q.setSelectionRange(caret, caret); } catch (e) {} } }
  }

  function toast(msg, isErr) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (isErr ? ' err' : ''); t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, isErr ? 9000 : 4000);
  }

  function rebuild() {
    var sched = state.sched || (state.exp ? E.scheduleFromExport(state.exp) : null);
    state.report = sched ? E.buildReport(sched, state.exp) : null;
    if (state.report && state.selected && !state.report.lines.some(function (l) { return l.id === state.selected; })) state.selected = null;
  }

  // ---------- loading files ----------
  async function takeExport(buffer, name) {
    var wb = await BGI.xlsx.readWorkbook(buffer, 'UNDERBILLING_DATA');
    state.exp = E.parseExport(wb.rows);
    state.exportName = name;
  }
  function takeSchedule(textIn, name) { state.sched = E.parseSchedule(textIn); state.scheduleName = name; }

  async function handleFiles(files) {
    var list = Array.prototype.slice.call(files || []);
    if (!list.length) return;
    var done = [], errors = [];
    state.loading = !state.report; render();
    for (var i = 0; i < list.length; i++) {
      var f = list[i], lower = f.name.toLowerCase();
      try {
        if (/\.xlsx$/.test(lower)) { await takeExport(await f.arrayBuffer(), f.name); done.push(f.name); }
        else if (/\.csv$/.test(lower)) { takeSchedule(await f.text(), f.name); done.push(f.name); }
        else if (/\.pdf$/.test(lower)) errors.push(f.name + ' is a PDF. This report reads the Excel export (.xlsx) and the schedule file (.csv).');
        else if (/\.xls$/.test(lower)) errors.push(f.name + ' is the old Excel format. Save it as .xlsx and drop it again.');
        else errors.push(f.name + ' is not a file this report can read. Use the .xlsx export and the .csv schedule.');
      } catch (e) { errors.push(f.name + ': ' + (e && e.message ? e.message : 'could not be read.')); }
    }
    state.loading = false;
    rebuild(); render();
    if (errors.length) toast(errors.join(' '), true);
    else toast('Loaded ' + done.join(' and ') + '.');
  }

  function api(path) { return fetch('api/' + path, { headers: { 'X-BGI-Token': state.token }, cache: 'no-store' }); }

  async function loadFromFolder() {
    try {
      var res = await api('files');
      if (!res.ok) return;
      var info = await res.json();
      state.loading = true; render();
      var errors = [];
      if (info.export) {
        try { var r = await api('file?name=' + encodeURIComponent(info.export.name)); if (!r.ok) throw new Error('could not be opened.'); await takeExport(await r.arrayBuffer(), info.export.name); }
        catch (e) { errors.push(info.export.name + ': ' + e.message); }
      }
      if (info.schedule) {
        try { var r2 = await api('file?name=' + encodeURIComponent(info.schedule.name)); if (!r2.ok) throw new Error('could not be opened.'); takeSchedule(await r2.text(), info.schedule.name); }
        catch (e2) { errors.push(info.schedule.name + ': ' + e2.message); }
      }
      state.loading = false; rebuild(); render();
      if (errors.length) toast(errors.join(' '), true);
      else if (state.report) toast('Loaded the newest files from the report folder.');
    } catch (e) { state.loading = false; render(); }
  }

  // ---------- events ----------
  function onAction(el) {
    var a = el.getAttribute('data-action');
    if (a === 'tab') { state.tab = el.getAttribute('data-tab'); render(); window.scrollTo(0, 0); }
    else if (a === 'open') { state.selected = el.getAttribute('data-id'); state.tab = 'calc'; render(); window.scrollTo(0, 0); }
    else if (a === 'jump') { var t = document.getElementById(el.getAttribute('data-id')); if (t) t.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    else if (a === 'print') window.print();
    else if (a === 'browse') document.getElementById('file').click();
    else if (a === 'clear-filter') { state.q = ''; state.px = ''; render(); }
  }

  function start() {
    root = document.getElementById('app');
    var m = /[?&]t=([A-Za-z0-9_-]+)/.exec(location.search);
    state.token = m ? m[1] : null;
    state.server = /^https?:$/.test(location.protocol) && !!state.token;

    document.addEventListener('click', function (e) {
      var el = e.target.closest ? e.target.closest('[data-action]') : null;
      if (el) onAction(el);
    });
    document.addEventListener('keydown', function (e) {
      if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('tr[data-action]')) { e.preventDefault(); onAction(e.target); }
    });
    document.addEventListener('input', function (e) { if (e.target.id === 'q') { state.q = e.target.value; render(); } });
    document.addEventListener('change', function (e) {
      if (e.target.id === 'px') { state.px = e.target.value; render(); }
      else if (e.target.id === 'jobpick') { state.selected = e.target.value; render(); }
      else if (e.target.id === 'file') { handleFiles(e.target.files); e.target.value = ''; }
    });
    var depth = 0;
    window.addEventListener('dragenter', function (e) { e.preventDefault(); depth++; document.body.classList.add('dragging'); });
    window.addEventListener('dragleave', function () { depth = Math.max(0, depth - 1); if (!depth) document.body.classList.remove('dragging'); });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); depth = 0; document.body.classList.remove('dragging'); handleFiles(e.dataTransfer && e.dataTransfer.files); });

    render();
    if (state.server) {
      loadFromFolder();
      setInterval(function () { api('ping').catch(function () {}); }, 5000);
      window.addEventListener('pagehide', function () { try { navigator.sendBeacon('api/bye?t=' + encodeURIComponent(state.token)); } catch (e) {} });
    }
  }

  BGI.app = { start: start, _state: state, _handleFiles: handleFiles };
  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  }
})(typeof globalThis !== 'undefined' ? (globalThis.BGI = globalThis.BGI || {}) : (window.BGI = window.BGI || {}));
