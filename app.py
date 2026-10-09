"""Streamlit web interface for the existing business-report pipeline."""
from __future__ import annotations

import hashlib
import json

import pandas as pd
import streamlit as st

from src.web_runner import MAX_UPLOAD_BYTES, SUPPORTED_SUFFIXES, run_uploaded_report


st.set_page_config(page_title="Business Findings Report", page_icon="📊", layout="wide")
st.markdown("""
<style>
html, body, [class*="css"] { font-family: system-ui, sans-serif; }
.block-container { max-width: 1120px; padding-top: 2.25rem; padding-bottom: 4rem; }
.hero { padding: 2rem 2.1rem; border-radius: 22px; color: #f7fafc; background: linear-gradient(120deg,#10263c,#194a53 68%,#36776c); margin-bottom: 1.6rem; }
.hero h1 { font-family: system-ui,sans-serif; letter-spacing: -0.04em; font-size: clamp(2rem,4vw,3.35rem); margin: 0 0 .65rem; }
.hero p { color: #d8e7e9; font-size: 1.05rem; max-width: 720px; margin: 0; }
.eyebrow { text-transform: uppercase; letter-spacing: .14em; font-size: .72rem; font-weight: 700; color: #9fe2c2; margin-bottom: .8rem; }
.privacy { border-left: 4px solid #dfa94b; padding: .8rem 1rem; background: #fff8e8; border-radius: 8px; color: #51411f; }
.small-note { color: #64748b; font-size: .9rem; }
div[data-testid="stMetric"] { background: #f4f7f8; border: 1px solid #e4eaed; padding: 1rem; border-radius: 14px; }
</style>
<section class="hero">
  <div class="eyebrow">Evidence-led business analysis</div>
  <h1>Turn a business dataset into clear next steps.</h1>
  <p>Inspect the data and analysis plan first, then generate a downloadable PDF, Excel workbook, or both using the reporting pipeline.</p>
</section>
""", unsafe_allow_html=True)

st.markdown("""
<div class="privacy"><strong>Data handling:</strong> uploaded files are sent to this hosting service for analysis. This app runs each job in a temporary workspace and does not save source files to a persistent app folder or database. Avoid uploading confidential or regulated data unless you are authorized to use this hosted service. Download your results before leaving the session.</div>
""", unsafe_allow_html=True)
st.write("")

left, right = st.columns([1.08, .92], gap="large")
with left:
    st.subheader("1 · Choose your dataset")
    uploaded = st.file_uploader(
        "Upload a tabular business file",
        type=[suffix[1:] for suffix in sorted(SUPPORTED_SUFFIXES)],
        help="Supported: CSV, TSV, XLSX and JSON. Hosted upload limit: 25 MB.",
    )
    if uploaded:
        st.caption(f"{uploaded.name} · {uploaded.size:,} bytes")
    output_format = st.radio("Report format", ["PDF", "Excel workbook", "Both"], horizontal=True)
    format_value = {"PDF": "pdf", "Excel workbook": "xlsx", "Both": "both"}[output_format]

with right:
    st.subheader("2 · Optional field guidance")
    config_text = st.text_area(
        "JSON config",
        value="{}",
        height=154,
        help="Use this for confirmed role mappings, units, date formats, Excel sheet/header choices, or assumptions. Leave {} if no clarification is needed.",
    )
    with st.expander("Config example"):
        st.code('''{
  "roles": {"date": "order_date", "quantity": "units_sold", "price": "unit_price"},
  "units": {"units_sold": "units", "unit_price": "USD per unit"},
  "assumptions": ["Each row represents one sales line."]
}''', language="json")

st.markdown("#### 3 · Inspect first, then build")
st.caption("The plan profiles the actual columns and lists supported and skipped analyses. Review it before asking for report exports.")
file_bytes = uploaded.getvalue() if uploaded else None
file_hash = hashlib.sha256(file_bytes).hexdigest() if file_bytes is not None else None
config_hash = hashlib.sha256(config_text.encode("utf-8")).hexdigest()
plan_key = (file_hash, uploaded.name if uploaded else None, format_value, config_hash)

plan_col, build_col = st.columns([1, 1])
with plan_col:
    plan_clicked = st.button("Inspect dataset & show plan", type="primary", use_container_width=True, disabled=not uploaded)
with build_col:
    build_clicked = st.button("Generate report", use_container_width=True, disabled=not uploaded)

if plan_clicked and uploaded and file_bytes is not None:
    try:
        with st.spinner("Profiling the dataset in a temporary workspace…"):
            planned = run_uploaded_report(uploaded.name, file_bytes, format_value, config_text, plan_only=True)
        st.session_state["business_findings_plan"] = {"key": plan_key, "result": planned}
        st.session_state.pop("business_findings_report", None)
    except Exception as exc:
        st.error(f"Could not inspect this file: {exc}")

saved_plan = st.session_state.get("business_findings_plan")
if saved_plan and saved_plan.get("key") == plan_key:
    planned = saved_plan["result"]
    st.divider()
    st.subheader("Dataset profile and analysis plan")
    if planned.get("profile"):
        load = planned["profile"].get("load", {})
        shape = load.get("shape", [])
        p1, p2, p3 = st.columns(3)
        p1.metric("Rows", f"{shape[0]:,}" if shape else "—")
        p2.metric("Columns", f"{shape[1]:,}" if len(shape) > 1 else "—")
        p3.metric("Input format", str(load.get("format", "—")).upper().lstrip("."))
        st.write("**Fields detected:** " + ", ".join(map(str, load.get("columns", []))))
        st.dataframe(pd.DataFrame(planned["profile"].get("profile", {}).get("columns", [])), use_container_width=True, hide_index=True)
        first_five = load.get("first_five", [])
        if first_five:
            with st.expander("Preview first five rows"):
                st.dataframe(pd.DataFrame(first_five), use_container_width=True, hide_index=True)
    if planned.get("plan"):
        st.markdown(planned["plan"])
    if planned.get("stderr"):
        with st.expander("Processing notes"):
            st.code(planned["stderr"])
    if planned.get("audit_markdown"):
        st.download_button("Download this plan's audit", planned["audit_markdown"], file_name="business-findings-plan-audit.md", mime="text/markdown")
    if "blockers: 0" not in planned.get("stdout", ""):
        st.warning("The plan reports a decision or mapping to resolve. Update the optional JSON config, then inspect the plan again before generating.")

if build_clicked and uploaded and file_bytes is not None:
    if not saved_plan or saved_plan.get("key") != plan_key:
        st.warning("Inspect the dataset and review its plan first. If you changed the file, format, or config, rerun the plan.")
    elif "blockers: 0" not in saved_plan["result"].get("stdout", ""):
        st.error("Resolve the plan's decision items before building. Edit the config and rerun the plan.")
    else:
        try:
            with st.spinner("Building the report, charts, and audit evidence…"):
                generated = run_uploaded_report(uploaded.name, file_bytes, format_value, config_text, plan_only=False)
            st.session_state["business_findings_report"] = {"key": plan_key, "result": generated}
        except Exception as exc:
            st.error(f"Report build failed: {exc}")

saved_report = st.session_state.get("business_findings_report")
if saved_report and saved_report.get("key") == plan_key:
    result = saved_report["result"]
    st.divider()
    if not result.get("ok"):
        st.error("The report was not produced. Review the message below, correct the config or dataset if needed, and inspect it again.")
        if result.get("stderr"):
            st.code(result["stderr"])
    else:
        report = result.get("report") or {}
        st.success("Report build completed. Automated checks and audit evidence are included with the downloads; the visual-review status remains pending until a person inspects the exports.")
        summary = report.get("summary", [])
        st.subheader("Findings preview")
        for line in summary[:5]:
            st.write(f"• {line}")
        for finding in report.get("findings", []):
            st.markdown(f"#### {finding.get('id', 'Finding')}: {finding.get('title', 'Business finding')}")
            st.write(finding.get("text", ""))
            figure = finding.get("figure", finding.get("id"))
            if figure in result.get("charts", {}):
                st.image(result["charts"][figure], caption=f"Figure {figure}", use_container_width=True)
        st.subheader("Download deliverables")
        downloads = st.columns(4)
        for col, suffix, label, mime in zip(
            downloads,
            ("pdf", "xlsx"),
            ("Download PDF", "Download Excel"),
            ("application/pdf", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ):
            if suffix in result.get("exports", {}):
                with col:
                    col.download_button(label, result["exports"][suffix], file_name=f"{uploaded.name.rsplit('.', 1)[0]}_findings_report.{suffix}", mime=mime, use_container_width=True)
        if result.get("markdown"):
            downloads[2].download_button("Download report text", result["markdown"], file_name="findings_report.md", mime="text/markdown", use_container_width=True)
        if result.get("audit_markdown"):
            downloads[3].download_button("Download run audit", result["audit_markdown"], file_name="business-findings-audit.md", mime="text/markdown", use_container_width=True)
        if result.get("audit_json"):
            with st.expander("Automated verification details"):
                verification = result.get("verification") or {}
                st.json(verification)
                st.download_button("Download audit JSON", result["audit_json"], file_name="business-findings-audit.json", mime="application/json")
        if report.get("limitations"):
            with st.expander("Limitations and assumptions"):
                for note in report["limitations"]:
                    st.write(f"- {note}")

st.markdown("""
<hr><p class="small-note">The application adapts to the uploaded fields. It does not infer missing dimensions or claim that observed associations prove causation. Hosted demo upload limit: 25 MB.</p>
""", unsafe_allow_html=True)
