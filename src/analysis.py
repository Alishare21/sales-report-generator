"""Evidence objects shared by prose, charts, PDF and Excel; no example metrics."""
from __future__ import annotations

import itertools
import re

import numpy as np
import pandas as pd

from load import NeedsInput, key

ALIASES = {
    "quantity": {"quantity", "units", "units_sold", "sales_quantity", "qty"},
    "price": {"price", "unit_price", "unit_price_inr", "unit_price_usd"},
    "revenue": {"revenue", "sales_amount", "net_sales", "revenue_inr", "revenue_usd"},
    "profit": {"profit", "gross_profit", "net_profit"},
    "cost": {"total_cost", "cost_amount"},
    "customer": {"customer_id", "client_id"},
    "region": {"sales_region", "store", "store_id", "region"},
    "flag": {"promo", "promoted", "promotion", "campaign_flag", "ab_group"},
}


def discover(frame, classes, config):
    roles = {}
    for role, aliases in ALIASES.items():
        matches = [c for c in frame if key(c) in aliases]
        if len(matches) == 1:
            roles[role] = matches[0]
    roles.update(config.get("roles", {}))
    for role, col in roles.items():
        if col not in frame:
            raise NeedsInput(f"Role {role} refers to missing column {col}.")
        if role in {"quantity", "price", "revenue", "profit", "cost"} and classes[col] != "Numeric measure" and not (classes[col] == "Constant" and pd.to_numeric(frame[col].dropna(), errors='coerce').notna().all()):
            raise NeedsInput(f"{col} is not a varying numeric measure. Confirm its meaning/class.")
    dates = [c for c, k in classes.items() if k == "Date/time"]
    if "date" not in roles and len(dates) == 1:
        roles["date"] = dates[0]
    return roles


def clean(frame, classes, formats, roles, config, as_of):
    data = frame.copy(deep=True)
    log, limitations, blockers = [], [], []

    def record(step, before, reason, affected=0):
        log.append({"step": step, "rows_before": before, "rows_after": len(data),
                    "affected_cells": int(affected), "reason": reason})

    before = len(data)
    data = data.dropna(how="all").copy()
    record("Blank rows", before, "Remove wholly empty records.")
    # Only trailing rows labeled total in their first nonempty field qualify.
    before = len(data)
    totals = 0
    while len(data):
        values = data.iloc[-1].dropna()
        if len(values) and str(values.iloc[0]).strip().casefold() in {"total", "grand total", "subtotal"}:
            if config.get("remove_total_rows", False):
                data = data.iloc[:-1].copy()
                totals += 1
            else:
                blockers.append("Trailing total/subtotal row detected. Confirm remove_total_rows=true.")
                break
        else:
            break
    record("Trailing totals", before, "Remove confirmed summary rows only.", totals)
    before = len(data)
    data = data.drop_duplicates().copy()
    record("Exact duplicates", before, "Remove identical full rows; do not deduplicate by item/name/URL.")
    for col, kind in classes.items():
        before = len(data)
        if kind == "Category":
            old = data[col].copy()
            data[col] = data[col].astype("string").str.strip().str.casefold().replace("", pd.NA)
            affected = (old.astype("string").fillna("") != data[col].fillna("")).sum()
            record(f"Normalize {col}", before, "Strip and casefold before grouping; retain missing category as its own group.", affected)
        elif kind == "Numeric measure" or (kind == "Constant" and col in [roles.get(r) for r in ('quantity','price','revenue','cost','profit')]):
            old = data[col].notna()
            data[col] = pd.to_numeric(data[col], errors="coerce").astype(float)
            invalid = ~np.isfinite(data[col]) & data[col].notna()
            data.loc[invalid, col] = np.nan
            failures = int((old & data[col].isna()).sum())
            record(f"Parse {col}", before, "Invalid/nonfinite numeric cells become missing; rows remain for other analyses.", failures)
            if failures:
                limitations.append(f"{col}: {failures:,} invalid numeric values became missing.")
        elif kind == "Date/time":
            if col not in formats:
                blockers.append(f"Date format for {col} is ambiguous. Set date_formats using an explicit strptime format.")
                continue
            old = data[col].notna()
            data[col] = pd.to_datetime(data[col], format=formats[col], errors="coerce", utc=True).dt.tz_localize(None)
            failed = int((old & data[col].isna()).sum())
            record(f"Parse {col}", before, f"Explicit date format {formats[col]}; failures excluded only from date analyses.", failed)
            if failed:
                limitations.append(f"{col}: {failed:,} dates failed parsing and remain missing.")
            future = int((data[col].dt.date > as_of).sum())
            if future:
                limitations.append(f"{col}: {future:,} future-dated records relative to {as_of}.")
                if not config.get("allow_future_dates", False):
                    blockers.append(f"{col} contains future dates; confirm forecast/planned data with allow_future_dates=true or correct the source.")
        count = int(data[col].isna().sum())
        record(f"Null policy {col}", len(data), f"Keep {count:,} nulls; available-case analysis with explicit denominators, no imputation.")
    for role in ("quantity", "price"):
        col = roles.get(role)
        if col and (data[col] < 0).any() and col not in config.get("nonnegative", []):
            blockers.append(f"Negative {col}: confirm returns are valid, or declare this column nonnegative.")
    if config.get("allow_negative", []):
        blockers = [x for x in blockers if not any(x.startswith(f"Negative {col}:") for col in config["allow_negative"])]
    for col in config.get("nonnegative", []):
        if col not in data or classes[col] != "Numeric measure":
            raise NeedsInput(f"nonnegative references a missing or non-numeric field: {col}")
        before = len(data)
        data = data.loc[~(data[col] < 0)].copy()
        record(f"Impossible {col}", before, "Exclude negative values only under confirmed nonnegative domain rule.")
    row_key = config.get("row_key")
    if row_key:
        if row_key not in data:
            raise NeedsInput(f"Unknown row_key: {row_key}")
        if data[row_key].isna().any() or data[row_key].duplicated().any():
            blockers.append(f"Confirmed row key {row_key} has missing/conflicting values; resolve rather than arbitrarily delete rows.")
    removed = (len(frame) - len(data)) / len(frame) * 100
    if removed > 10:
        limitations.append(f"{removed:.1f}% of source rows were removed, above the 10% quality limit.")
    if data.empty:
        blockers.append("No records remain after cleaning.")
    if config.get("synthetic"):
        limitations.append("SYNTHETIC DEMONSTRATION: generated fictional records. Findings do not describe any real business.")
    constants = [c for c, k in classes.items() if k == "Constant"]
    if constants:
        limitations.append("Constant columns excluded from analysis: " + ", ".join(constants) + ".")
    for col, kind in classes.items():
        if kind == "ID" and data[col].nunique() < len(data) * .1:
            limitations.append(f"{col}: only {data[col].nunique():,} unique identifiers across {len(data):,} rows; repeated items need grain validation.")
        if kind == "Category":
            counts = data[col].value_counts()
            if len(counts) > 1 and counts.max() == counts.min():
                limitations.append(f"{col} has exactly uniform group counts; check sampling or synthetic structure.")
        if kind == "Free text / URL" and data[col].dropna().astype(str).str.match(r'^https?://').mean() > .5:
            unique = int(data[col].nunique())
            if unique < len(data) * .1:
                limitations.append(f"{col}: {unique:,} distinct URLs across {len(data):,} records. Repeated URLs were retained because observations differ; independent product counts and additive sales totals are unverified.")
    return data.reset_index(drop=True), log, limitations, list(dict.fromkeys(blockers))


def capabilities(data, classes, roles, config):
    currency = config.get("currency")
    currency_cols = [c for c in data if key(c) in {"currency", "currency_code"}]
    currency_problem = None
    if currency_cols:
        vals = data[currency_cols[0]].dropna().astype(str).str.strip().str.upper().unique()
        if len(vals) > 1:
            currency_problem = "Multiple currencies: split by currency or supply a verified conversion before monetary aggregation."
        elif len(vals) == 1:
            currency = currency or vals[0]
        if data[currency_cols[0]].isna().any() and not config.get("currency"):
            currency_problem = "Missing currency values require a confirmed common currency or separate reports."
    skips, assumptions = [], list(config.get("assumptions", []))
    assumptions.append("Exact duplicate rows represent duplicate exports. Other repeated identifiers are retained.")
    assumptions.append("Missing measurements are kept; totals and means use only observed values with stated coverage.")
    assumptions.append("A group mean difference below 5% is descriptively small, not a statistical equivalence test.")
    primary = config.get('primary_measure','auto')
    if primary not in {'auto','records','quantity'}:
        raise NeedsInput('primary_measure must be auto, records or quantity.')
    if primary == 'records':
        metric, unit = None, 'records'
        skips.append('Revenue totals and revenue projections are intentionally not reported; record-level comparisons do not establish additive business sales.')
    elif primary == 'quantity':
        if 'quantity' not in roles:
            raise NeedsInput('primary_measure=quantity requires a mapped quantity field.')
        metric, unit = roles['quantity'], config.get('units',{}).get(roles['quantity'],'units')
        skips.append('Monetary totals are not selected for this report; quantity is the primary measure.')
    elif "revenue" in roles:
        metric, unit = roles["revenue"], currency or "currency unspecified"
    elif "quantity" in roles and "price" in roles:
        metric, unit = "revenue_est", currency or "currency unspecified"
        if metric in data:
            raise NeedsInput("Input already contains revenue_est; map it explicitly or rename the source column.")
        data[metric] = data[roles["quantity"]] * data[roles["price"]]
        assumptions.append(f"Estimated revenue = {roles['price']} * {roles['quantity']}; assumes per-unit prices, no unrecorded discounts, taxes or returns.")
    elif "quantity" in roles:
        metric, unit = roles["quantity"], config.get("units", {}).get(roles["quantity"], "units")
        skips.append("Revenue unavailable: no revenue amount or price-plus-quantity pair.")
    else:
        metric, unit = None, "records"
        skips.append("No identified additive revenue/quantity measure; use record counts, not arbitrary numeric totals.")
    if unit == "currency unspecified":
        assumptions.append("Currency was not supplied; amounts are reported as currency unspecified, with no currency symbol.")
    if "date" not in roles:
        skips.append("No date field selected: no trend, seasonality, period comparison or YoY analysis.")
    else:
        skips.append("No seasonality or YoY claim: coverage/repeated cycles are not verified. Trend uses observed days only.")
    if not ("profit" in roles or "cost" in roles):
        skips.append("No cost or profit field: margin, profitability and profit impact are unavailable.")
    else:
        skips.append("Profit/cost fields present; profitability is deferred until compatible row grain and amount definitions are verified.")
    if "customer" not in roles:
        skips.append("No customer ID: customer retention, cohorts and concentration are unavailable.")
    else:
        skips.append("Customer IDs present; repeat/cohort analysis is deferred until transaction grain and observation window are verified.")
    if "region" not in roles:
        skips.append("No identified sales geography/store: no regional interpretation of unrelated columns.")
    flags = [c for c, k in classes.items() if k == "Category" and data[c].dropna().nunique() == 2
             and set(data[c].dropna().astype(str).str.casefold()).issubset({'yes','no','true','false','0','1'})]
    if roles.get("flag") in data and data[roles["flag"]].dropna().nunique() == 2:
        flags = [roles["flag"]] + [x for x in flags if x != roles["flag"]]
    nums = [c for c, k in classes.items() if k == "Numeric measure"]
    categories = [c for c, k in classes.items() if k == "Category" and c not in flags[:1]]
    if roles.get("region") in categories:
        categories = [c for c in categories if c != roles["region"]] + [roles["region"]]
    selected_categories = config.get('category_columns', categories[:2])
    comparisons = config.get('comparison_columns',flags[:1])
    segments = config.get('segment_columns',categories[:2])
    for column in set(selected_categories + comparisons + segments):
        if column not in data or classes[column] != 'Category':
            raise NeedsInput(f'Configured comparison/category field is missing, constant or unsuitable: {column}')
    display = "estimated revenue" if metric == "revenue_est" else metric or "record count"
    plan = []
    for col in selected_categories:
        plan.append({"question": f"How concentrated is {display} across {col}?", "columns": [col] + ([metric] if metric else []), "chart": "horizontal bar"})
    if roles.get('quantity') or metric:
        for column in comparisons:
            plan.append({'question':f'How do observed outcomes differ by {column}, including within other groups?',
                         'columns':[column,roles.get('quantity') or metric]+[c for c in segments if c!=column],
                         'chart':'bar of observed group means; flag descriptively small differences'})
    if roles.get("date"):
        plan.append({"question": f"How does daily {display} vary across observed dates?", "columns": [roles['date']] + ([metric] if metric else []), "chart": "line; no unequal-period comparison"})
    if len(nums) >= 2:
        pair = [roles["price"], roles["quantity"]] if roles.get("price") and roles.get("quantity") else nums[:2]
        plan.append({"question": f"Are {pair[0]} and {pair[1]} associated?", "columns": pair, "chart": "scatter"})
    if len(plan) < 4:
        skips.append(f"Only {len(plan)} candidate analyses found; report fewer than four charts if more cannot be supported.")
    if len(plan)>6:
        raise NeedsInput('More than six analyses selected; reduce categories/comparisons to keep the report focused.')
    return {"metric": metric, "metric_label": display, "unit": unit, "currency": currency,
            "currency_problem": currency_problem if metric and (metric == 'revenue_est' or metric == roles.get('revenue')) else None,
            "roles": roles, "questions": plan, "skipped": skips, "assumptions": assumptions,
            "categories": selected_categories, "flags": flags, "numeric": nums,
            "comparisons": comparisons, "segments":segments,
            "scope_note":config.get('scope_note',''),
            "allow_impact_projection":config.get('allow_impact_projection',True)}


def num(x, places=1):
    return f"{x:,.{places}f}"


def analyze(data, cap, classes, config):
    metric, label, unit, roles = cap['metric'], cap['metric_label'], cap['unit'], cap['roles']
    measure = data[metric] if metric else pd.Series(1., index=data.index)
    valid = measure.notna()
    total = float(measure.sum(min_count=1))
    facts = {"rows": len(data), "valid_metric_rows": int(valid.sum()), "total": total,
             "mean": float(measure.mean()), "min": float(measure.min()), "max": float(measure.max()),
             "period": "not available", "unit": unit}
    if roles.get('quantity'):
        outcome_values=data[roles['quantity']].dropna()
        facts.update(outcome_column=roles['quantity'],outcome_mean=float(outcome_values.mean()),
                     outcome_total=float(outcome_values.sum()),outcome_count=int(len(outcome_values)))
    findings, groups, limitations = [], {}, []

    def add(kind, title, text, columns, series, **extra):
        ident = f"F{len(findings) + 1}"
        findings.append(dict(id=ident, figure=len(findings)+1, kind=kind, title=title,
                             text=text, columns=columns, series=series, **extra))

    for col in cap['categories']:
        missing_label = "(missing)"
        while missing_label in set(data[col].dropna().astype(str)):
            missing_label = "_" + missing_label
        work = pd.DataFrame({'group': data[col].fillna(missing_label), 'value': measure})
        grouped = work.groupby('group', dropna=False)['value'].agg(['count', 'sum']).sort_values('sum', ascending=False)
        groups[col] = grouped.reset_index().to_dict('records')
        # No conclusions from low-n groups, even if their apparent share is large.
        if len(grouped) < 2 or (grouped['count'] < 30).any() or grouped['sum'].min() < 0 or not total > 0:
            limitations.append(f"{col}: group conclusions skipped (requires at least two groups, 30 observed outcomes each, and nonnegative share totals).")
            continue
        top, bottom = grouped.iloc[0], grouped.iloc[-1]
        top_name, bottom_name = str(grouped.index[0]), str(grouped.index[-1])
        share = float(top['sum'] / total)
        title = f"{top_name} accounts for {share:.1%} of {label}"
        text = (f"{top_name} contributes {num(top['sum'])} {unit}, versus {num(bottom['sum'])} for {bottom_name}. "
                f"Its {share:.1%} share uses all {num(total)} observed {unit} as the base ({int(top['count']):,} contributing records); "
                f"prioritize monitoring this exposure, without assuming that higher totals mean greater efficiency.")
        add('barh', title, text, [col] + ([metric] if metric else []),
            {'labels': list(map(str, grouped.index)), 'values': grouped['sum'].tolist()},
            x_label=f"{label} ({unit})", y_label=col, group=col, top_name=top_name,
            top_total=float(top['sum']), top_n=int(top['count']), share=share)

    outcome = roles.get('quantity') or metric
    for flag in (cap.get('comparisons',cap['flags'][:1]) if outcome else []):
        usable = data.dropna(subset=[flag, outcome])
        grouped = usable.groupby(flag)[outcome].agg(['count', 'sum', 'mean']).sort_index()
        if 2 <= len(grouped) <= 12 and grouped['count'].min() >= 30:
            lo, hi = grouped.sort_values('mean').iloc[0], grouped.sort_values('mean').iloc[-1]
            ordered = grouped.sort_values('mean',kind='stable')
            lo_name, hi_name = str(ordered.index[0]), str(ordered.index[-1])
            diff = float(hi['mean'] - lo['mean'])
            pct = diff / abs(float(lo['mean'])) if lo['mean'] != 0 else None
            checks = []
            total_segments = 0
            for segment in [c for c in cap.get('segments',cap['categories'][:2]) if c!=flag]:
                for name, section in usable.groupby(segment, dropna=False):
                    total_segments += 1
                    sub = section.groupby(flag)[outcome].agg(['count', 'mean'])
                    if hi_name in sub.index and lo_name in sub.index and sub['count'].min() >= 30:
                        checks.append({'segment': segment, 'value': str(name),
                                       'difference': float(sub.loc[hi_name, 'mean'] - sub.loc[lo_name, 'mean'])})
            same = sum(x['difference'] > 0 for x in checks)
            consistency = (f"The direction holds in {same} of {len(checks)} eligible subsegments; "
                           f"{total_segments - len(checks)} segments lack both groups with 30 outcomes.") if checks else "No eligible subsegment consistency check is available."
            small = pct is not None and pct < .05
            title = (f'{flag}: only a {pct:.1%} spread in mean {outcome}' if small else f"{flag} groups differ by {pct:.1%} in mean {outcome}") if pct is not None else f"{flag} groups differ by {num(diff)} {outcome} per record"
            text = (f"{hi_name} averages {num(hi['mean'],2)} versus {num(lo['mean'],2)} {outcome} per record for {lo_name} "
                    f"(n={int(hi['count']):,} and {int(lo['count']):,}; difference {num(diff,2)}). "
                    f"{consistency} " + ("The difference is below the stated 5% descriptive threshold; do not prioritize a rollout on this signal." if small else "This is an association; use a controlled test before changing policy."))
            grouped = grouped.sort_values('mean',ascending=False,kind='stable')
            add('bar', title, text, [flag, outcome], {'labels': list(map(str,grouped.index)), 'values': grouped['mean'].tolist()},
                x_label=flag, y_label=f"Mean {outcome} per record", difference=diff,
                lower_mean=float(lo['mean']), higher_mean=float(hi['mean']), eligible_n=int(lo['count']),
                outcome=outcome, subsegments=checks, small=small, group=flag,
                higher_group=hi_name,lower_group=lo_name,
                group_stats=[{'label':str(k),'count':int(v['count']),'total':float(v['sum']),'mean':float(v['mean'])} for k,v in grouped.iterrows()])
            if small:
                limitations.append(f'{flag}: similar observed group means (largest spread {pct:.1%}). This is a descriptive flat result, not proof of equivalence or authentic data provenance.')
        else:
            limitations.append(f"{flag}: comparison skipped; require two to twelve groups, each with at least 30 observed outcomes.")

    date = roles.get('date')
    if date and pd.api.types.is_datetime64_any_dtype(data[date]):
        usable = pd.DataFrame({'date': data[date].dt.normalize(), 'value': measure}).dropna()
        if len(usable):
            facts['period'] = f"{usable['date'].min():%Y-%m-%d} to {usable['date'].max():%Y-%m-%d}"
        grouped = usable.groupby('date')['value'].agg(['count','sum'])
        if len(grouped) >= 2 and grouped['count'].min() >= 30:
            high_day, low_day = grouped['sum'].idxmax(), grouped['sum'].idxmin()
            high, low = float(grouped.loc[high_day,'sum']), float(grouped.loc[low_day,'sum'])
            title = f"Daily {label} ranges from {num(low)} to {num(high)}"
            text = (f"{high_day:%Y-%m-%d} records {num(high)} {unit}, compared with {num(low)} on {low_day:%Y-%m-%d}. "
                    f"These are equal one-day observations across {len(grouped)} recorded days; investigate workload variation before claiming growth or seasonality. "
                    "Unobserved dates are gaps, not assumed zeros.")
            add('line', title, text, [date] + ([metric] if metric else []),
                {'labels':[str(x.date()) for x in grouped.index], 'values':grouped['sum'].tolist()},
                x_label=date, y_label=f"Daily {label} ({unit})")
        else:
            limitations.append(f"{date}: daily trend skipped because each shown day needs at least 30 observed outcomes and two dates.")

    nums = cap['numeric']
    if len(nums) >= 2:
        x, y = [roles['price'], roles['quantity']] if roles.get('price') and roles.get('quantity') else nums[:2]
        paired = data[[x,y]].dropna()
        if len(paired) >= 30 and paired[x].nunique() > 1 and paired[y].nunique() > 1:
            corr = float(paired[x].corr(paired[y]))
            magnitude = "weak" if abs(corr) < .3 else "moderate" if abs(corr) < .7 else "strong"
            title = f"{x} and {y}: {magnitude} association (r={corr:.2f})"
            text = (f"Pearson correlation is {corr:.3f} across {len(paired):,} complete pairs, versus zero for no linear association. "
                    f"This {magnitude} observed relationship does not establish causation; compare like-for-like segments before using it to set targets.")
            add('scatter', title, text, [x,y], {'x':paired[x].tolist(),'y':paired[y].tolist()},
                x_label=f"{x} ({config.get('units',{}).get(x, 'source units, unspecified')})",
                y_label=f"{y} ({config.get('units',{}).get(y, 'source units, unspecified')})", correlation=corr)
    if len(findings) < 4:
        limitations.append(f"Only {len(findings)} findings meet the evidence requirements; the requested four-to-six range is not forced.")
    if len(data) < 30:
        limitations.append("Fewer than 30 cleaned records: no business group conclusions are supported.")
    priorities=config.get('finding_order',[])
    if priorities:
        findings.sort(key=lambda f: priorities.index(f['columns'][0]) if f['columns'][0] in priorities else len(priorities))
        for i,finding in enumerate(findings,1):
            finding.update(id=f'F{i}',figure=i)
    recommendations = []
    flag_findings = [f for f in findings if f['kind'] == 'bar']
    if flag_findings:
        f = flag_findings[0]
        recommendations.append({'priority':'High', 'action':f"After validating the data source and experimental unit, have the data owner design a controlled test of {f['columns'][0]}, balancing the available categories.",
            'finding':f['id'], 'impact':"No business impact is projected from this observational comparison. Set test size using the independent experimental unit, outcome variability, and a prespecified minimum detectable effect; measure outcomes prospectively. Cost and profit effects are unavailable."})
    small_findings = [f for f in flag_findings if f['small']]
    if small_findings:
        f=small_findings[0]
        recommendations.append({'priority':'Medium','action':f"Avoid a broad change to {f['group']} on this evidence alone. Ask the operator to collect exposure and availability alongside {f['outcome']} in a matched pilot.",
            'finding':f['id'],'impact':f"Current highest-to-lowest mean gap = {num(f['higher_mean'],2)} - {num(f['lower_mean'],2)} = {num(f['difference'],2)} recorded volume per record. Incremental impact is not estimable; coverage is {sum(g['count'] for g in f['group_stats']):,} source records. The 5% descriptive threshold is an analysis assumption, not a statistical test."})
    for f in [x for x in findings if x['kind']=='barh'][:2]:
        recommendations.append({'priority':'Medium', 'action':f"Have the data owner validate identifiers, repeated items and exposure for {f['group']} = {f['top_name']} before using its apparent concentration to change allocation.",
            'finding':f['id'], 'impact':f"Review coverage: {f['top_n']:,} records ({f['share']:.1%} = group count / {num(total)} total records). Incremental impact is not estimable; validate with a controlled test."})
    if len(recommendations) < 3:
        for f in findings:
            if len(recommendations) >= 3:
                break
            if any(r['finding'] == f['id'] for r in recommendations):
                continue
            recommendations.append({'priority':'Medium','action':f"Validate the {', '.join(f['columns'])} source definitions with the data owner before a controlled operational test.",
                'finding':f['id'], 'impact':f"Coverage = {facts['valid_metric_rows']:,} observed records / {facts['rows']:,} retained records. Incremental business effect is not estimable from observational data."})
    if not findings:
        limitations.append("No defensible findings or finding-linked recommendations; supply a richer dataset or resolve semantic mappings.")
    return facts, findings[:6], recommendations[:5], groups, limitations
