# Business findings reporting specification

Owner: Muhammad Akram. Captured 2026-10-07. Applies to any tabular business
dataset (CSV, TSV, XLSX, JSON). Produce charts, insights, recommendations and PDF
and/or Excel. Use `example_findings_report.md`, when present, for structure, tone
and detail only. Never copy its numbers or column names. Discover the dataset first.

## Golden rules
1. Never invent data, columns, periods or numbers. State what is missing.
2. Every reported number comes from code run on the actual data.
3. Never modify raw files. Work on a copy.
4. Do not claim a cause the data cannot prove. Use association language.
5. State assumptions. Ask before building if a material assumption is unresolved.
6. Verify before calling the work done.

## Project layout and execution
`data/raw/` contains read-only originals; `data/processed/` contains cleaned copies;
`outputs/charts/` contains 200-dpi PNGs (dataset subfolders allowed); `outputs/`
contains final PDF/XLSX files. Python modules: `src/load.py`, `src/analysis.py`,
`src/charts.py`, `src/build_report.py`. Pin versions in `requirements.txt`.
README rerun instructions must be five lines or fewer.
One command: `python src/build_report.py --input data/raw/<file> --format pdf|xlsx`.
Defaults: Python 3.10+, pandas, matplotlib Agg, openpyxl or xlsxwriter, reportlab.
Seed randomness; avoid heavy/unusual packages. An explicitly documented higher
Python minimum is acceptable when pinned dependencies require it.

## 1. Load safely
- Detect extension and verify content, rejecting mismatches.
- CSV/TSV: use csv.Sniffer or test comma, semicolon, tab and pipe. A single header
  containing separators indicates a bad delimiter; retry. Try utf-8, utf-8-sig,
  latin-1. Preserve identifiers and leading zeros until classification.
- Excel: list every sheet, choose the sheet with most data or ask. Detect headers
  below row one, merged cells, blank top rows, bottom total rows. Exclude totals
  from analysis only when verified; record every removed row.
- JSON: flatten records only as needed. Ask which array if ambiguous. Do not
  explode nested lists into invented row grain.
- Print and save shape, names, dtypes, first five rows. Compare physical line count
  and parsed logical records allowing quoting/header differences. Stop on bad loads.

## 2. Profile
For each column: dtype, null count/percentage, unique count, min/max or top values.
Classify ID (unique/near-unique identity), date/time, numeric measure, category,
free text/URL, or constant. Do not confuse a unique continuous measure with an ID.
Exclude constants from analysis, mentioning them once. Report exact duplicate rows,
duplicate IDs, whitespace/case inconsistencies, negative quantities/prices and
future dates. Flag potential impossibilities; confirm whether negatives represent
returns and whether dates are forecasts before excluding them.

## 3. Capability check
Record in the plan what is allowed and skipped:
- Date/time: trends, seasonality or equal-length period comparisons when coverage
  supports them. Without date: state 'no date field', no trends or YoY.
- Revenue/sales amount: totals/shares/rankings. Price plus quantity may derive
  `revenue_est = price * quantity`, labeled 'estimated revenue', with its assumption.
  Otherwise use units or counts. Validate aggregation units and currency.
- Profit/cost: profitability only when definitions and denominator are compatible.
  Without them: no profit/margin figures or claims about raising profit.
- Customer ID: repeat/cohort/concentration only with known transactional grain;
  skip if absent or if repeated rows do not establish repeat purchases.
- Geography/store/region: regional comparison; manufacturing origin is not sales region.
- Binary flags: group comparisons and within-subsegment consistency checks.
- Two numeric measures: correlations or banded comparisons, with adequate observations.
Choose business questions from actual fields and values, never fabricate dimensions.

## 4. Plan
Show 4-6 business questions before building in interactive work, identifying columns,
chart type, assumptions, and skipped analyses with reasons. If fewer are defensible,
explain the shortfall rather than padding with invented or repeated analyses.

## 5. Clean
Log every change: step, rows before, rows after, reason (also affected-cell counts).
Remove exact duplicates. Deduplicate by ID only if confirmed as the row key; never
blindly deduplicate URLs, names or descriptions. Preserve legitimate repeated IDs.
Parse dates explicitly and count failures. Choose and state keep/drop/fill for every
column's nulls; retain unused free-text nulls. Strip and standardize category case
before grouping. Do not remove large outliers without an impossibility reason.
Report final rows and removal percentage; flag removal above 10% as a limitation.
Call out synthetic/repeated data, few unique items, uniform distributions or
surprisingly identical group means. Do not infer synthetic origin with certainty.

## 6. Analyze
Define each metric once and reuse it. Calculate ratios from totals, never means of
row ratios. State units, currency and time period (or unavailable). Every percentage
has an explicit denominator. Do not draw conclusions from groups under 30 rows.
For group differences check subsegment consistency where possible and state missing
checks. Report effect sizes in plain language. Distinguish a flat result using an
explicitly stated descriptive threshold; do not mistake it for significance testing.

## 7. Charts
Target 4-6, each with one message and finding-led title. Label axes and units.
Use lines for time, bars for categories, horizontal bars for long labels, scatter or
banded bars for numeric relationships. No 3D, dual axes, or pie with over five slices.
Bars start at zero (negative data may extend below zero). Sort by value except
natural order. Use a color-blind-safe palette, fonts at least 10 pt, no overlap.
Save PNG at 200 dpi, approximately eight inches wide.

## 8. Findings
Target 4-6 ordered by business relevance. Each has a computed metric and number,
comparison, implication, and figure reference. Each chart gets two or three
interpretive sentences. Use only computed actual values, never example statistics.

## 9. Recommendations
Target 3-5 with linked finding, specific action, affected process/group, priority
High/Medium/Low, and impact method with assumptions. Express impact only in supported
measures. Never promise profit without profit data. Missing costs/margins means
controlled tests instead of broad rollouts. A conditional scenario is not a forecast
or causal estimate. If effects are not identifiable, state 'not estimable' and
quantify measurable review/test coverage instead. No vague advice.

## Report order
1. Title, dataset, covered period (or unavailable), generation date.
2. Executive summary, maximum five lines.
3. Data overview and methodology: rows, used columns, cleaning, assumptions, metrics.
4. Key findings: chart plus interpretation.
5. Recommendations table: priority, action, linked finding, expected impact.
6. Limitations and next steps: missing fields, assumptions, quality and what cannot
   be concluded, always included.
7. Appendix: cleaning log, metric definitions, column dictionary.

## PDF
ReportLab; A4/Letter, 2 cm margins, consistent styles, page numbers. Scale charts to
page width and keep captions with them. Wrap table cells, fit widths, repeat headers.
Main report max eight pages; appendix may follow. Render at least two pages and
inspect for clipping/overlap before marking visual verification done.

## Excel
Exact sheet order: Summary, Findings, Charts, Recommendations, Data_Summary,
Cleaned_Data, Notes. Real typed values and currency/percentage (one decimal)/thousands
formats. Data_Summary totals and ratios use live Excel formulas referencing source
cells, with cached calculated values when possible. Freeze headers, bold headers,
filter Cleaned_Data, readable widths, no merged data cells or internal blank rows.
Keep IDs. Omit long free text/URLs only if needed for size and note every omission.
Reopen with openpyxl and check sheet names. Inspect each sheet visually.
Narrative and embedded images are snapshots; disclose that changes require rerun.

## Common failures to prevent
Wrong delimiter/encoding; invented dates/profit/region; relabeled manufacturing origin;
wrong deduplication key; charted constants/IDs/free text; averaged percentages;
unequal period comparisons; hardcoded report numbers; causal claims; findings with
no comparison/implication; impact without method; leftover examples/placeholders;
failure to open and inspect exports.

## Verification checklist
- Loaded correctly: plausible shape/columns, not a merged column.
- Raw file unchanged (SHA-256 before/after).
- Capability check recorded and respected.
- Cleaning log reconciles row counts.
- Three numbers independently checked against pandas.
- Category parts reconcile to total.
- Charts have finding titles, unit axes, and text references.
- Findings have number, comparison and implication.
- Recommendations link to findings with impact methods.
- Limitations list missing fields and assumptions.
- No unresolved placeholders or copied example numbers.
- Reports reopen; PDF pages and every workbook sheet visually inspected.
- README provides one-command rebuild.

## Final response
Provide created paths; three leading findings (one line each); assumptions/data issues
for user review; every checklist item marked done or not done. Do not claim full
completion while verification is pending. Keep generated demonstrations explicitly
fictional. User-supplied datasets are not automatically authentic business records;
state unresolved provenance and grain without assuming synthetic origin.

## Per-run audit and skill improvement (user addition, 2026-10-07)
Audit every report, retain dated evidence and track unresolved/resolved issues.
Correct verified skill defects, add regression coverage, rebuild and re-audit.
Automated checks and visual review must be separate. Do not claim a real CSV was
tested until the user supplies one. Use the actual next CSV as an acceptance test.
