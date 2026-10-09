import unittest
import csv
import io
import json

from src.web_runner import run_uploaded_report


class WebRunnerTests(unittest.TestCase):
    def test_rejects_unsupported_uploads_and_invalid_config(self):
        with self.assertRaisesRegex(ValueError, "CSV, TSV, XLSX or JSON"):
            run_uploaded_report("report.pdf", b"x", "pdf")
        with self.assertRaisesRegex(ValueError, "valid JSON"):
            run_uploaded_report("report.csv", b"x,y\n1,2\n", "pdf", "{")

    def test_plan_runs_in_isolated_job_and_returns_profile_audit(self):
        csv_data = (
            "order_date,product,quantity,revenue\n"
            "2025-01-01,Alpha,2,20\n"
            "2025-01-02,Beta,3,33\n"
            "2025-01-03,Alpha,1,10\n"
        ).encode("utf-8")
        result = run_uploaded_report("sample.csv", csv_data, "pdf", plan_only=True)
        self.assertTrue(result["ok"], result["stderr"])
        self.assertIn("## Capability check", result["plan"])
        self.assertEqual(result["profile"]["load"]["shape"], [3, 4])
        self.assertIsNotNone(result["audit_json"])
        self.assertIsNotNone(result["audit_markdown"])
        self.assertTrue(any("order_date" in c["column"] for c in result["profile"]["profile"]["columns"]))

    def test_full_run_returns_both_downloadable_exports_and_audit(self):
        buffer = io.StringIO(newline="")
        writer = csv.writer(buffer)
        writer.writerow(["order_date", "product", "quantity", "revenue", "promo", "store"])
        for i in range(60):
            writer.writerow([
                f"2025-03-{(i % 30) + 1:02d}",
                "Alpha" if i < 30 else "Beta",
                (i % 5) + 1,
                ((i % 5) + 1) * (12 if i < 30 else 15),
                "yes" if i % 2 else "no",
                "East" if i % 2 else "West",
            ])
        config = {
            "roles": {"date": "order_date", "quantity": "quantity", "revenue": "revenue", "flag": "promo", "region": "store"},
            "currency": "USD",
            "units": {"quantity": "units"},
            "assumptions": ["Each row is one observed sales record."],
        }
        result = run_uploaded_report("integration.csv", buffer.getvalue().encode(), "both", json.dumps(config), plan_only=False)
        self.assertTrue(result["ok"], result["stderr"])
        self.assertEqual(set(result["exports"]), {"pdf", "xlsx"})
        self.assertTrue(result["exports"]["pdf"].startswith(b"%PDF"))
        self.assertTrue(result["exports"]["xlsx"].startswith(b"PK"))
        self.assertIsNotNone(result["report"])
        self.assertIsNotNone(result["verification"])
        self.assertIsNotNone(result["audit_json"])


if __name__ == "__main__":
    unittest.main()
