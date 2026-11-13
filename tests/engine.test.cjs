// Run with: node --test tests/
// All fixtures are synthetic. No company data lives in this repository.
const test = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('node:zlib');
require('../web/src/xlsx.js');
require('../web/src/engine.js');
const { xlsx, engine } = globalThis.BGI;

// ---------- helpers ----------
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function zip(files, deflate) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, text] of Object.entries(files)) {
    const data = Buffer.from(text, 'utf8');
    const body = deflate ? zlib.deflateRawSync(data) : data;
    const nameBuf = Buffer.from(name, 'utf8');
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(deflate ? 8 : 0, 8);
    local.writeUInt32LE(crc32(data), 14); local.writeUInt32LE(body.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBuf.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(deflate ? 8 : 0, 10);
    central.writeUInt32LE(crc32(data), 16); central.writeUInt32LE(body.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, body); centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + body.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(files).length, 8); end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...locals, cd, end]));
}
function workbook(deflate) {
  return zip({
    'xl/workbook.xml': '<workbook xmlns:r="x"><sheets><sheet name="Other" sheetId="1" r:id="rId1"/><sheet name="UNDERBILLING_DATA" sheetId="2" r:id="rId2"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>JOB_CODE</t></si><si><t>Smith &amp; Sons &lt;HQ&gt;</t></si><si><r><t>rich </t></r><r><t>text</t></r></si></sst>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
    'xl/worksheets/sheet2.xml': '<worksheet><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>NAME</t></is></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="C2"><v>1234.56</v></c><c r="D2" t="s"><v>2</v></c><c r="E2" t="b"><v>1</v></c></row>' +
      '</sheetData></worksheet>'
  }, deflate);
}

const HEADER = ['BU_CODE', 'JOB_CODE', 'JOB_NAME', 'PX', 'RD', 'RD_2', 'CONTRACT_VALUE', 'START_DATE', 'COST_TO_DATE', 'BILLED_TO_DATE',
  'UNDERBILLED_AMOUNT', 'OVERBILLED_AMOUNT', 'BILLING_POSITION', 'NET_BILL_AMOUNT', 'AR_PAID_AMOUNT', 'AR_OUTSTANDING_AMOUNT', 'AP_PAID_AMOUNT',
  'PAID_AMOUNT_VOUCHER', 'PAID_AMOUNT_NON_VOUCHER', 'AP_OUTSTANDING_AMOUNT', 'AP_OPEN_COUNT', 'CASH_POSITION', 'CASH_POSITION_PERCENT',
  'AP_CALCULATION_DATE', 'JOB_AND_SUBJOB_COUNT', 'SUBJOB_DETAIL'];
function sub(code, name, o = {}) {
  const v = [code, name, 'I', null, o.contract || 0, null, null, null, null, null, null, o.cost || 0, o.billed || 0, (o.cost || 0) - (o.billed || 0),
    o.netBill || 0, 0, o.arPaid || 0, o.arOut || 0, o.apPaid || 0, o.apV || 0, o.apNV || 0, o.apOut || 0, o.apOpen || 0, null, null, null];
  return { v: { f: v.map((x) => ({ v: x === null ? null : String(x) })) } };
}
function jobRow(code, name, px, start, subs, over = {}) {
  const s = (f) => subs.reduce((a, x) => a + Number(x.v.f[f].v), 0);
  const cost = s(11), billed = s(12), arPaid = s(16), arOut = s(17), apPaid = s(18);
  const row = { BU_CODE: '500', JOB_CODE: code, JOB_NAME: name, PX: px, RD: 'R', RD_2: null, CONTRACT_VALUE: s(4), START_DATE: start,
    COST_TO_DATE: cost, BILLED_TO_DATE: billed, UNDERBILLED_AMOUNT: cost - billed, OVERBILLED_AMOUNT: billed - cost, BILLING_POSITION: 'X',
    NET_BILL_AMOUNT: s(14), AR_PAID_AMOUNT: arPaid, AR_OUTSTANDING_AMOUNT: arOut, AP_PAID_AMOUNT: apPaid, PAID_AMOUNT_VOUCHER: s(19),
    PAID_AMOUNT_NON_VOUCHER: s(20), AP_OUTSTANDING_AMOUNT: s(21), AP_OPEN_COUNT: s(22), CASH_POSITION: arPaid - apPaid,
    CASH_POSITION_PERCENT: s(4) ? (arPaid - apPaid) / s(4) : null, AP_CALCULATION_DATE: 46326 /* 2026-10-31 */, JOB_AND_SUBJOB_COUNT: subs.length,
    SUBJOB_DETAIL: JSON.stringify({ v: subs }), ...over };
  return HEADER.map((h) => row[h]);
}
const D = (iso) => Math.round(Date.parse(iso + 'T00:00:00Z') / 86400000) + 25569;
function sampleRows() {
  return [HEADER,
    jobRow('1000', 'Alpha Tower', 'Pat Lee', D('2025-01-15'), [
      sub('1000', 'Alpha Tower', { contract: 1000000, cost: 600000, billed: 450000, netBill: 450000, arPaid: 400000, arOut: 50000, apPaid: 300000, apV: 300000 }),
      sub('100002', "Alpha Tower - GC's", { apPaid: 50000, apNV: 50000 }),
      sub('100003', 'Alpha Tower - Trades', { apPaid: 150000, apV: 150000, apOut: 20000, apOpen: 3 }),
      sub('100006', 'Alpha Tower - Insurances', { apPaid: 10000, apNV: 10000 })]),
    jobRow('2000', 'Beta School', 'Sam Roy', D('2027-02-01'), [sub('2000', 'Beta School', { contract: 500000, cost: 40000, billed: 0 })]),
    jobRow('3000', 'Gamma Clinic', 'Pat Lee', D('2026-03-01'), [sub('3000', 'Gamma Clinic', { contract: 200000, cost: 5000.5, billed: 1000.25 })]),
    jobRow('4000', 'Delta Lab', 'Sam Roy', D('2026-03-01'), [sub('4000', 'Delta Lab', { contract: 100000, cost: 0, billed: 900000, apPaid: 70000, apV: 70000 })])
  ];
}

// ---------- xlsx reader ----------
for (const deflate of [false, true]) {
  test('xlsx reader: ' + (deflate ? 'deflated' : 'stored') + ' workbook, preferred sheet, cell types', async () => {
    const wb = await xlsx.readWorkbook(workbook(deflate), 'UNDERBILLING_DATA');
    assert.equal(wb.sheetName, 'UNDERBILLING_DATA');
    assert.deepEqual(wb.sheetNames, ['Other', 'UNDERBILLING_DATA']);
    assert.deepEqual(wb.rows[0], ['JOB_CODE', 'NAME']);
    assert.equal(wb.rows[1][0], 'Smith & Sons <HQ>');
    assert.equal(wb.rows[1][1], undefined);
    assert.equal(wb.rows[1][2], 1234.56);
    assert.equal(wb.rows[1][3], 'rich text');
    assert.equal(wb.rows[1][4], true);
  });
}
test('xlsx reader: falls back to the first sheet and rejects non-Excel input', async () => {
  const wb = await xlsx.readWorkbook(workbook(true), 'NOPE');
  assert.equal(wb.sheetName, 'Other');
  await assert.rejects(() => xlsx.readWorkbook(new Uint8Array(Buffer.from('%PDF-1.7 not a zip'))), /not an Excel/);
  await assert.rejects(() => xlsx.readWorkbook(zip({ 'hello.txt': 'hi' }, false)), /not an Excel workbook/);
});

// ---------- primitives ----------
test('money is held in whole cents', () => {
  assert.equal(engine.cents(853117.23), 85311723);
  assert.equal(engine.cents('1,019,890.76'), 101989076);
  assert.equal(engine.cents(0.1 + 0.2), 30);
  assert.equal(engine.cents(null), 0);
  assert.equal(engine.cents(-0.005), -0);
});
test('dates: Excel serials and text', () => {
  assert.equal(engine.isoDate(46326), '2026-10-31');
  assert.equal(engine.isoDate(45000.75), '2023-03-15');
  assert.equal(engine.isoDate('2026-10-15 00:00:00'), '2026-10-15');
  assert.equal(engine.isoDate('12/14/26'), '2026-12-14');
  assert.equal(engine.isoDate(''), null);
});
test('sub-jobs are sorted into AP categories by name', () => {
  const c = (code, name) => engine.categorize(code, name, '5000');
  assert.equal(c('5000', 'Job'), 'MAIN');
  assert.equal(c('500003', "Job-GC's"), 'GC');
  assert.equal(c('500002', 'Job - GCs'), 'GC');
  assert.equal(c('500003', 'Job General Conditions'), 'GC');
  assert.equal(c('500003', 'Job - GRs'), 'GR');
  assert.equal(c('500003', 'Job General Req'), 'GR');
  assert.equal(c('500003', "Job-GC's/GR's"), 'GCGR');
  assert.equal(c('500003', "Job - GC's GRs"), 'GCGR');
  assert.equal(c('500003', 'Job - GC/GR'), 'GCGR');
  assert.equal(c('500004', 'Job-Trades'), 'TRADES');
  assert.equal(c('500004', 'Job Cost of Work'), 'TRADES');
  assert.equal(c('500006', 'Job - Insurances, etc'), 'INS');
  assert.equal(c('500004', 'Job - Non Reimb'), 'NR');
  assert.equal(c('500006', 'Job-NR'), 'NR');
  assert.equal(c('500001', 'Job-Precon'), 'PRECON');
  assert.equal(c('500002', 'Job-Payroll'), 'PAYROLL');
  assert.equal(c('500002', 'Job - PR'), 'PAYROLL');
  assert.equal(c('500005', 'Job-GMP'), 'FEE');
  assert.equal(c('500070', 'Job-ODP'), 'ODP');
  assert.equal(c('500098', 'Job-Ops Sec'), 'SEC');
  assert.equal(c('500099', 'Job'), 'SEC');
  assert.equal(c('5000ES', 'Job'), 'EST');
  assert.equal(c('500009', 'Job phase nine'), 'OTHER');
});

// ---------- export ----------
test('export: derives underbilling and cash position itself, and checks the file against them', () => {
  const exp = engine.parseExport(sampleRows());
  assert.equal(exp.jobs.length, 4);
  assert.equal(exp.asOf, '2026-10-31');
  const a = exp.byCode['1000'];
  assert.equal(a.under, 15000000);
  assert.equal(a.cash, 400000_00 - 510000_00);
  assert.equal(a.cashPct, -0.11);
  assert.equal(a.startDate, '2025-01-15');
  assert.ok(Object.values(a.checks).every(Boolean));
  const d = exp.byCode['4000'];
  assert.equal(d.checks.P1, false, 'cost below AP paid is flagged');
  assert.equal(d.checks.P2, false, 'billing above contract is flagged');
});
test('export: a wrong figure in the file is caught, not trusted', () => {
  const rows = sampleRows();
  rows[1][HEADER.indexOf('UNDERBILLED_AMOUNT')] = 999;
  rows[1][HEADER.indexOf('CASH_POSITION')] = 1;
  const a = engine.parseExport(rows).byCode['1000'];
  assert.equal(a.under, 15000000);
  assert.equal(a.checks.E1, false);
  assert.equal(a.checks.E3, false);
});
test('export: missing columns and empty sheets give a plain error', () => {
  assert.throws(() => engine.parseExport([['JOB_CODE', 'JOB_NAME'], ['1', 'x']]), /Missing columns: .*COST_TO_DATE/);
  assert.throws(() => engine.parseExport([HEADER]), /empty/);
});
test('AP by category: the four requested categories always show, totals tie out', () => {
  const ap = engine.apByCategory(engine.parseExport(sampleRows()).byCode['1000']);
  const by = Object.fromEntries(ap.rows.map((r) => [r.id, r]));
  assert.deepEqual(ap.rows.map((r) => r.id), ['GC', 'GR', 'TRADES', 'INS', 'MAIN']);
  assert.equal(by.GC.apPaid, 5000000);
  assert.equal(by.GR.present, false);
  assert.equal(by.TRADES.apOut, 2000000);
  assert.equal(by.TRADES.apOpen, 3);
  assert.equal(by.INS.apPaid, 1000000);
  assert.equal(ap.total, 51000000);
  assert.ok(Math.abs(ap.rows.reduce((a, r) => a + r.share, 0) - 1) < 1e-9);
});

// ---------- schedule ----------
const CSV = [
  'row_type,section,job_code,job_name,area_manager,project_executive,underbilling_2026-09-30,underbilling_2026-10-31,const_start,net_billing_2026-10,billed_to_date_pct,comment',
  'job,NOT_STARTED,2000,Beta School,Ann Wu,Sam Roy,"30,000","40,000",2027-02-01,0,0.0,',
  'subtotal,NOT_STARTED,,Subtotal,,,30000,40000,,,,',
  'small,SMALL,,Jobs under $10K,,,5000,4000,,,,',
  'subtotal,SMALL,,Subtotal,,,5000,4000,,,,',
  'job,IN_PROCESS,1000,Alpha Tower,Ann Wu,Pat Lee,100000,150000,2025-01-15,"(7,133)",45.0,"On hold, per owner"',
  'job,IN_PROCESS,9999,Not In Export,Ann Wu,Pat Lee,12000,-,2026-01-01,0,10.0,',
  'subtotal,IN_PROCESS,,Subtotal,,,112000,150000,,,,',
  'total,,,Total,,,147000,194000,,,,'
].join('\r\n');

test('schedule: parses amounts, dates, quoted commas and printed totals', () => {
  const s = engine.parseSchedule(CSV);
  assert.deepEqual(s.periods, ['2026-09-30', '2026-10-31']);
  assert.equal(s.netMonth, '2026-10');
  assert.equal(s.lines.length, 4);
  assert.deepEqual(s.lines[0].amounts, [30000, 40000]);
  assert.equal(s.lines[2].netBilling, -7133);
  assert.equal(s.lines[2].comment, 'On hold, per owner');
  assert.deepEqual(s.lines[3].amounts, [12000, 0]);
  assert.deepEqual(s.printed.total, [147000, 194000]);
  assert.throws(() => engine.parseSchedule('a,b\n1,2'), /does not look like an Underbilling Schedule/);
});
test('report: totals, change, ranking, section rule and the join to Excel', () => {
  const exp = engine.parseExport(sampleRows());
  const rep = engine.buildReport(engine.parseSchedule(CSV), exp);
  assert.equal(rep.asOf, '2026-10-31');
  assert.deepEqual(rep.total, [147000, 194000]);
  assert.ok(rep.totalOk);
  assert.ok(rep.sections.every((s) => s.sumsOk));
  assert.equal(rep.matched, 2);
  assert.equal(rep.jobLines, 3);
  const alpha = rep.lines.find((l) => l.code === '1000');
  assert.equal(alpha.change, 50000);
  assert.equal(alpha.rank, 1);
  assert.equal(alpha.sectionOk, true);
  assert.equal(alpha.excelUnderDollars, 150000);
  assert.equal(alpha.excelAgrees, true);
  const beta = rep.lines.find((l) => l.code === '2000');
  assert.equal(beta.sectionOk, true);
  assert.equal(beta.excelAgrees, true);
  assert.equal(rep.lines.find((l) => l.code === '9999').excel, null);
  const x0 = rep.checks.find((c) => c.id === 'X0');
  assert.equal(x0.failed, 1);
  assert.equal(rep.exportStale, false);
  assert.deepEqual(rep.sections.map((s) => s.id), ['IN_PROCESS', 'NOT_STARTED', 'SMALL'], 'jobs in process come first');
  assert.equal(alpha.excelUsable, true, 'plausible and matching: Excel figures are shown');
  assert.equal(rep.lines.find((l) => l.code === '9999').excelUsable, false, 'not in the export: data required');
});
test('data required: an export figure that cannot be right is not shown as the math', () => {
  const rows = sampleRows();
  const exp = engine.parseExport(rows);
  const csv = CSV.replace('job,IN_PROCESS,9999,Not In Export,Ann Wu,Pat Lee,12000,-', 'job,IN_PROCESS,4000,Delta Lab,Ann Wu,Pat Lee,12000,-');
  const rep = engine.buildReport(engine.parseSchedule(csv), exp);
  const delta = rep.lines.find((l) => l.code === '4000');
  assert.equal(delta.excelAgrees, true, 'both say zero underbilling');
  assert.equal(delta.excelUsable, false, 'but cost is below AP paid and billing is above contract');
  assert.ok(engine.DATA_REQUIRED.length >= 8);
});
test('report: a subtotal that does not add up is reported', () => {
  const rep = engine.buildReport(engine.parseSchedule(CSV.replace('subtotal,IN_PROCESS,,Subtotal,,,112000,150000', 'subtotal,IN_PROCESS,,Subtotal,,,112000,175000')), null);
  assert.equal(rep.sections.find((s) => s.id === 'IN_PROCESS').sumsOk, false);
  assert.equal(rep.totalOk, false);
  assert.equal(rep.checks, null);
});
test('Excel-only schedule: built by rule from the export', () => {
  const exp = engine.parseExport(sampleRows());
  const s = engine.scheduleFromExport(exp);
  const rep = engine.buildReport(s, exp);
  assert.ok(rep.fromExport);
  const by = Object.fromEntries(rep.lines.map((l) => [l.id, l]));
  assert.equal(by['1000'].section, 'IN_PROCESS');
  assert.equal(by['1000'].current, 150000);
  assert.equal(by['2000'].section, 'NOT_STARTED');
  assert.equal(by['2000'].current, 40000);
  assert.equal(by.SMALL.current, 4000, 'Gamma (3,999.25 underbilled) rolls into small jobs, rounded to whole dollars');
  assert.equal(by['4000'], undefined, 'overbilled jobs are not listed');
  assert.deepEqual(rep.total, [194000]);
  assert.ok(rep.lines.filter((l) => l.excel).every((l) => l.excelAgrees));
});
