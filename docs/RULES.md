# The rules

Every number in the report is either read from an input file or produced by one of these rules. The code for each is in `web/src/engine.js`; the rule IDs match the IDs there and on the Data check page.

Money is held in whole cents. The report recalculates each formula itself and never trusts a calculated column in the file; it then compares its answer with the file and reports any row that differs by more than two cents.

## Formulas on the Excel export

| ID | Rule | Formula |
|----|------|---------|
| E1 | Underbilled amount | `COST_TO_DATE − BILLED_TO_DATE`. Positive is underbilled, negative is overbilled. |
| E2 | Overbilled amount | `−UNDERBILLED_AMOUNT` |
| E3 | Cash position | `AR_PAID_AMOUNT − AP_PAID_AMOUNT` |
| E4 | AP paid | `PAID_AMOUNT_VOUCHER + PAID_AMOUNT_NON_VOUCHER` |
| E5 | Net billed | `AR_PAID_AMOUNT + AR_OUTSTANDING_AMOUNT` |
| E6 | Job row | Each money column on the job row equals the sum of the same figure across `SUBJOB_DETAIL`. |
| E7 | Cash position % | `CASH_POSITION ÷ CONTRACT_VALUE` (blank when the contract value is zero) |

## Plausibility checks on the Excel export

These do not test arithmetic. They test whether a figure can be right.

| ID | Rule | Flagged when |
|----|------|--------------|
| P1 | Cost to date is at least what has already been paid out | `COST_TO_DATE < AP_PAID_AMOUNT` while AP paid is above zero |
| P2 | Billed to date does not exceed the contract value | `BILLED_TO_DATE > 105% × CONTRACT_VALUE` while the contract is above zero |

## Schedule rules

| ID | Rule |
|----|------|
| S1 | A job is listed when its underbilled amount is above zero. |
| S2 | Jobs underbilled less than $10,000 are rolled into one "small jobs" line. |
| S3 | Section subtotal = sum of its lines. Total = sum of the subtotals. Lines are printed in whole dollars, so a printed subtotal may differ from the sum of its lines by up to $0.50 per line; anything more is reported. |
| S4 | A job whose construction start date is after the report date is "not started"; otherwise it is "in process". |
| S5 | Change since last month = current month − prior month. A line that was zero last month is marked new; a line that is zero this month is marked cleared. |
| S6 | Share of total = line ÷ total. Rank is by current-month amount among job lines above zero. |

When a schedule file is loaded, S1 to S4 are checks on that file. When only the Excel export is loaded, the report builds the schedule itself with S1, S2 and S4.

## Schedule against Excel

| ID | Rule |
|----|------|
| X0 | Every job on the schedule is matched to the export by job number. |
| X1 | The export agrees with a schedule line when `max(0, COST_TO_DATE − BILLED_TO_DATE)` is within $1 of the schedule figure. |

The report also flags an export whose data date is earlier than the schedule's report date.

## AP by category

The export separates categories only by sub-job, so the category is read from the end of each sub-job's name. First match wins.

| Category | Sub-job name contains |
|----------|-----------------------|
| Insurance | "insur" |
| Non-reimbursable | "non reimb" (any spelling), "NR" |
| GCs / GRs combined | "GC" followed by "GR" at the end |
| GRs | ends in "GR(s)", or "general req" |
| GCs | ends in "GC(s)", or "general cond" |
| Trades | ends in "trades", "cost of work", "COW" |
| Precon | "precon" |
| Payroll | "payroll", ends in "PR" |
| GMP / fee | ends in "GMP" or "fee" |
| Owner direct purchase | ends in "ODP" |
| Ops / exec | "ops", "exec", "bonus", "commission", or suffix 98 / 99 |
| Estimating | job number ends in "ES" |
| Main job number | the entry whose number equals the job number |
| Other sub-jobs | anything else |

GCs, GRs, Trades and Insurance always appear in the table, with a dash when the job has no sub-job for them. Money posted to the main job number cannot be placed in a category from this export, and the report says so when it is half or more of a job's AP.
