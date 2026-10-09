"""Isolated execution adapter for the hosted Streamlit interface."""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SUPPORTED_SUFFIXES = {".csv", ".tsv", ".xlsx", ".json"}
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _slug(filename: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", Path(filename).stem).strip("_") or "dataset"


def run_uploaded_report(
    filename: str,
    content: bytes,
    output_format: str,
    config_text: str = "{}",
    plan_only: bool = False,
) -> dict:
    """Run one job in a disposable project copy; return only report artifacts in memory."""
    safe_name = Path(filename).name
    if Path(safe_name).suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ValueError("Upload a CSV, TSV, XLSX or JSON file.")
    if not content:
        raise ValueError("The selected file is empty.")
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("This hosted demo accepts files up to 25 MB.")
    if output_format not in {"pdf", "xlsx", "both"}:
        raise ValueError("Choose PDF, Excel, or both output formats.")
    try:
        config = json.loads(config_text or "{}")
    except json.JSONDecodeError as exc:
        raise ValueError(f"Config must be valid JSON (line {exc.lineno}, column {exc.colno}).") from exc
    if not isinstance(config, dict):
        raise ValueError("The optional config must be a JSON object.")

    with tempfile.TemporaryDirectory(prefix="business-findings-") as temp:
        root = Path(temp)
        shutil.copytree(PROJECT_ROOT / "src", root / "src")
        reference = PROJECT_ROOT / "example_findings_report.md"
        if reference.exists():
            shutil.copy2(reference, root / reference.name)
        raw_dir = root / "data" / "raw"
        raw_dir.mkdir(parents=True)
        source = raw_dir / safe_name
        source.write_bytes(content)
        config_path = root / "upload.config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")

        command = [sys.executable, str(root / "src" / "build_report.py"),
                   "--input", str(source), "--format", output_format,
                   "--config", str(config_path)]
        if plan_only:
            command.append("--plan-only")
        completed = subprocess.run(command, cwd=root, capture_output=True, text=True, check=False)
        out = root / "outputs" / _slug(safe_name)
        result = {
            "ok": completed.returncode == 0,
            "returncode": completed.returncode,
            "stdout": completed.stdout[-12000:],
            "stderr": completed.stderr[-12000:],
            "plan": (out / "analysis_plan.md").read_text(encoding="utf-8") if (out / "analysis_plan.md").exists() else None,
            "profile": json.loads((out / "profile.json").read_text(encoding="utf-8")) if (out / "profile.json").exists() else None,
            "audit_json": None,
            "audit_markdown": None,
        }
        audit_dir = root / "audits" / "business-findings" / _slug(safe_name)
        audit_jsons = sorted(audit_dir.glob("*.json"))
        audit_mds = sorted(audit_dir.glob("*.md"))
        if audit_jsons:
            result["audit_json"] = audit_jsons[-1].read_bytes()
        if audit_mds:
            result["audit_markdown"] = audit_mds[-1].read_bytes()

        if completed.returncode == 0 and not plan_only:
            evidence_path = out / "evidence.json"
            result["report"] = json.loads(evidence_path.read_text(encoding="utf-8")) if evidence_path.exists() else None
            result["markdown"] = (out / "findings_report.md").read_bytes() if (out / "findings_report.md").exists() else None
            result["exports"] = {
                suffix: (out / f"findings_report.{suffix}").read_bytes()
                for suffix in ("pdf", "xlsx")
                if (out / f"findings_report.{suffix}").exists()
            }
            charts = {}
            if result["report"]:
                for finding in result["report"].get("findings", []):
                    chart = Path(finding.get("chart", ""))
                    if chart.is_file():
                        charts[finding.get("figure", finding.get("id", chart.stem))] = chart.read_bytes()
            result["charts"] = charts
            verification = out / "verification.json"
            result["verification"] = json.loads(verification.read_text(encoding="utf-8")) if verification.exists() else None
        return result
