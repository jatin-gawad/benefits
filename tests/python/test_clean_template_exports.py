"""Synthetic export checks for the sanitized source handoff's blank template."""
from pathlib import Path
import tempfile
import unittest
import zipfile
from openpyxl import load_workbook
from proposal_engine import DEFAULT_TEMPLATE_PATH
from proposal_engine.excel_writer import write_to_excel, TARGET_SHEETS
from proposal_engine.quote_test_extractor import _empty_pdf_comparison_model, _finalize_pdf_comparison_model, _write_quote_workbook


class CleanTemplateExportsTests(unittest.TestCase):
    def test_template_has_no_hidden_data_comments_or_external_links(self):
        with zipfile.ZipFile(DEFAULT_TEMPLATE_PATH) as archive:
            forbidden = ("externalLinks/", "threadedComments/", "persons/", "comments", "customXml/", "printerSettings/")
            self.assertFalse(any(part in name for name in archive.namelist() for part in forbidden))
        workbook = load_workbook(DEFAULT_TEMPLATE_PATH, data_only=False)
        try:
            self.assertTrue(set(TARGET_SHEETS).issubset(workbook.sheetnames))
            for sheet in workbook:
                self.assertEqual("visible", sheet.sheet_state)
                self.assertEqual(sheet.title, sheet["A2"].value)
                for row in sheet:
                    for cell in row:
                        self.assertIn(cell.value, (None, sheet.title, "Awaiting imported data."))
                        self.assertIsNone(cell.comment)
        finally:
            workbook.close()

    def test_standard_export_populates_synthetic_carrier_and_premium(self):
        model = {
            "proposal_summaries": [{"carrier_name": "Synthetic Carrier"}],
            "benefit_lines": [{"carrier_name": "Synthetic Carrier", "benefit_line": "Employee Life",
                               "class_name": "Class A", "rate": 0.1, "volume": 100000,
                               "monthly_premium": 10, "status": "MATCH"}],
        }
        with tempfile.TemporaryDirectory() as directory:
            result = write_to_excel(model, Path(directory)/"synthetic-standard.xlsx")
            workbook = load_workbook(result, data_only=False)
            try:
                values = [cell.value for row in workbook["Total Cost Comparison"] for cell in row]
                self.assertIn("Synthetic Carrier", values)
                self.assertIn(10, values)
                self.assertGreater(model["_market_survey_stats"]["total_cost_rows"], 0)
            finally:
                workbook.close()

    def test_quote_export_creates_report_tabs_with_no_client_inputs(self):
        model = _empty_pdf_comparison_model([])
        _finalize_pdf_comparison_model(model)
        with tempfile.TemporaryDirectory() as directory:
            result = _write_quote_workbook(model, Path(directory)/"synthetic-quote.xlsx")
            workbook = load_workbook(result, data_only=False)
            try:
                for name in ["Total Cost Comparison", "Gross Cost Comparison", "Plan Design", "Extraction Audit"]:
                    self.assertIn(name, workbook.sheetnames)
                self.assertTrue(all(sheet.sheet_state == "visible" for sheet in workbook))
                self.assertEqual("No rate data available", workbook["Rate Reconciliation"]["A2"].value)
                self.assertIsNone(workbook["Rate Reconciliation"]["F2"].value)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
