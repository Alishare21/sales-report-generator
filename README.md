# Sales Report Generator

Sales Report Generator turns tabular business data into a findings report with
charts, written insights, and practical recommendations. It inspects and profiles
the source, cleans a working copy, selects analyses that the available fields
support, then exports a PDF, an Excel workbook, or both.

The project is designed to work with an AI agent. `AGENTS.md` routes report tasks
to the reusable skill in `skills/business-findings/`. The agent reads
`reporting_rules.md`, examines the dataset and its meaning, and prepares a
capability based analysis plan before building the report. Each run records an
audit and verification evidence. Review the final report before using it for
business decisions.

## Features

- Reads CSV, TSV, XLSX, and JSON files. Delimited text loading checks common
  separators and encodings; Excel files are inspected by sheet.
- Profiles column types, missing values, unique values, ranges or common values,
  duplicates, and possible data quality concerns.
- Classifies fields as IDs, dates and times, numeric measures, categories, free
  text or URLs, and constants. These are discovered from actual values and can
  be reviewed or clarified in a JSON config.
- Selects supported comparisons, time summaries, concentration analysis, and
  numeric relationships based on fields present in the dataset.
- Creates labeled, color blind friendly PNG charts at 200 dpi and includes them
  in the report.
- Produces a PDF with an executive summary, method, key findings, charts,
  recommendations, limitations, and an appendix with the cleaning log, metric
  definitions, and column dictionary.
- Optionally creates an Excel workbook with `Summary`, `Findings`, `Charts`,
  `Recommendations`, `Data_Summary`, `Cleaned_Data`, and `Notes` sheets. Summary
  formulas recalculate from workbook data; narratives and charts are snapshots.
- Records the analysis plan, data profile, cleaning log, evidence, verification,
  and an append only audit for each report run.

## Data labels and business roles

| Label | Typical use |
| --- | --- |
| ID | Row identity; not a measure or chart category. |
| Date/time | Time analysis when dates parse and represent the period being studied. |
| Numeric measure | Candidate quantities, amounts, prices, costs, profits, or scores. |
| Category | Group comparisons and breakdowns, subject to group size. |
| Free text / URL | Context retained from the source; not grouped or charted by default. |
| Constant | Excluded from comparisons and charts; mentioned as a limitation. |

The agent can map roles such as `quantity`, `price`, `revenue`, `profit`, `cost`,
`customer`, `region`, and `flag`. Configs can also specify field units, date
formats, Excel sheet or header row, row key, assumptions, and selected comparison
or segment fields. Only map fields that exist and mean what the analysis expects.

Example config:

```json
{
  "roles": {
    "date": "order_date",
    "quantity": "units_sold",
    "price": "unit_price",
    "customer": "customer_id",
    "region": "sales_region",
    "flag": "promoted"
  },
  "units": {
    "units_sold": "units",
    "unit_price": "USD per unit"
  },
  "assumptions": ["Each row represents one sales line."]
}
```

For uncertain row grain, set `"primary_measure": "records"`, disable impact
projections, and explain the scope in `scope_note`. For example, a field named
`origin` is not treated as sales geography without evidence; product labels such
as `Winter` do not make a dataset time series.

## Setup

Requires Python 3.11 or newer.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

On macOS or Linux, activate the environment with
`source .venv/bin/activate`, then install with `python -m pip install -r
requirements.txt`.

## Run a report

Place your input in `data/raw/`. Inspect its profile and analysis plan first:

```powershell
python src/build_report.py --input data/raw/your_data.csv --format pdf --plan-only
```

After reviewing the plan and resolving any field meanings, build the PDF:

```powershell
python src/build_report.py --input data/raw/your_data.csv --format pdf
```

Use `--format xlsx` for Excel or `--format both` for both exports. Use
`--config data/raw/your_data.config.json` for role mappings or assumptions.
Reports go to `outputs/<dataset>/`, charts to `outputs/charts/<dataset>/`, and a
cleaned copy to `data/processed/`.

## Example report

Create the seeded **fictional demonstration dataset** and its config:

```powershell
python scripts/make_example.py
```

Generate its PDF and Excel example:

```powershell
python src/build_report.py --input data/raw/example_retail.csv --config data/raw/example_retail.config.json --format both
```

The example demonstrates available report sections and calculations. It contains
no findings about a real business. Review generated files in
`outputs/example_retail/`.

## Report contents and limitations

The PDF gives the data period when available, an executive summary, data
overview and method, key findings with figures, recommendations, limitations,
next steps, and an appendix. The main report is limited to eight pages where
the content allows; the appendix may add pages.

The Excel workbook keeps summary formulas live when inputs change. Written
findings and embedded chart images are snapshots, so rerun the report after
changing the data. Other outputs include `analysis_plan.md`, `profile.json`,
`cleaning_log.json`, `findings_report.md`, `evidence.json`, and
`verification.json`. Run history is in `audits/business-findings/`.

The report does not invent missing fields or results. No date means no trend;
no cost or profit means no margin or profit claim. Correlations and group
differences do not prove causes. Unknown row grain prevents certifying record
sums as business wide sales. Data profiling and automated checks cannot prove
the source is complete, representative, or authentic. Inspect the exported
report visually before relying on it.

Raw and processed datasets, reports, charts, and audits are ignored by Git by
default. Keep confidential data out of public repositories and inspect staged
files before committing; `.gitignore` is not a substitute for that check.

## Tests

```powershell
python -m unittest discover -s tests -v
```

## Project files

```text
AGENTS.md                       Agent workflow and safety rules
skills/business-findings/       Reusable reporting skill
reporting_rules.md              Full reporting specification
src/load.py                     Safe loading and profiling
src/analysis.py                 Cleaning, capability checks and analysis
src/charts.py                   Chart generation
src/build_report.py             PDF and Excel exports
src/audit.py                    Per run audit records
tests/                          Regression tests
scripts/make_example.py         Fictional sample data generator
```
