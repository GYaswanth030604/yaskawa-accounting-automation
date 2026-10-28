"""
Automated Test Suite for YASKAWA Import Accounting & Tally Excel Automation
Run with: .\\py312\\python.exe -m pytest test_automation.py -v
"""

from decimal import Decimal
from pathlib import Path
import pytest
import openpyxl

from main import (
    BoEParser,
    InvoiceParser,
    DeliveryNoteParser,
    MasterManager,
    AccountingEngine,
    ExcelGenerator,
    run_pipeline,
    INPUT_FOLDER,
    OUTPUT_FOLDER,
    BASE_EXCEL,
    MASTER_EXCEL,
    OUTPUT_FILENAME,
    ROUNDING_TOLERANCE
)

INPUT_DIR = Path(INPUT_FOLDER)
BOE_PATH = INPUT_DIR / "434214010102026INBLR4BE1131020261740.pdf"
INVOICE_PATH = INPUT_DIR / "Commercial Invoice.pdf"
MASTER_PATH = INPUT_DIR / MASTER_EXCEL
BASE_PATH = INPUT_DIR / BASE_EXCEL


class TestBoEParser:
    """Tests for Bill of Entry PDF parsing."""

    def test_boe_file_exists(self):
        assert BOE_PATH.exists(), f"BoE PDF file not found at {BOE_PATH}"

    def test_boe_header_and_items(self):
        parser = BoEParser(BOE_PATH)
        data = parser.parse()
        header = data["header"]
        items = data["items"]

        # Verify Header Fields
        assert header["boe_no"] == "4342140"
        assert header["exchange_rate"] == Decimal("110.50")
        assert header["total_bcd"] == Decimal("686294.20")
        assert header["total_sws"] == Decimal("68629.60")

        # Verify Items Count and Duties
        assert len(items) == 48, f"Expected 48 BoE items, got {len(items)}"

        sum_bcd = sum(it["bcd_amt"] for it in items)
        sum_sws = sum(it["sws_amt"] for it in items)

        assert sum_bcd == Decimal("686294.20"), f"Sum of BCD {sum_bcd} != 686294.20"
        assert sum_sws == Decimal("68629.60"), f"Sum of SWS {sum_sws} != 68629.60"

    def test_boe_first_and_last_item(self):
        parser = BoEParser(BOE_PATH)
        data = parser.parse()
        items = data["items"]

        first = items[0]
        assert first["itemsn"] == 1
        assert "H71-71A41" in first["part_no"]
        assert first["bcd_rate"] == Decimal("7.5")
        assert first["bcd_amt"] == Decimal("2572.2")
        assert first["sws_amt"] == Decimal("257.2")

        last = items[-1]
        assert last["itemsn"] == 48
        assert "972-0DP30" in last["part_no"]
        assert last["bcd_rate"] == Decimal("10.0")


class TestInvoiceParser:
    """Tests for Commercial Invoice PDF parsing and offline OCR."""

    def test_invoice_file_exists(self):
        assert INVOICE_PATH.exists(), f"Invoice PDF file not found at {INVOICE_PATH}"

    def test_invoice_ocr_extraction(self):
        parser = InvoiceParser(INVOICE_PATH)
        data = parser.parse()
        header = data["header"]
        lines = data["lines"]

        assert header["invoice_no"] == "980012194"
        assert header["currency"] == "EUR"
        assert header["total_eur"] == Decimal("75458.98")
        assert len(lines) == 48, f"Expected 48 invoice lines, got {len(lines)}"

        # Validate math on every line: Qty * Unit Price == Total Value within tolerance
        for line in lines:
            if line["unit_price_eur"] and line["total_eur"]:
                expected_total = (line["qty"] * line["unit_price_eur"]).quantize(Decimal("0.01"))
                assert abs(expected_total - line["total_eur"]) <= Decimal("1.00"), (
                    f"Line math mismatch for part {line['part_no']}: {line['qty']} * {line['unit_price_eur']} "
                    f"= {expected_total} != {line['total_eur']}"
                )


class TestDeliveryNoteParser:
    """Tests for Delivery Note range and serial expansion."""

    def test_range_expansion_simple(self):
        text = "114460 - 114462, 114467 - 114470"
        serials = DeliveryNoteParser.expand_serial_ranges(text)
        expected = ["114460", "114461", "114462", "114467", "114468", "114469", "114470"]
        assert serials == expected

    def test_range_expansion_leading_zeros(self):
        text = "00010 - 00013, 00025"
        serials = DeliveryNoteParser.expand_serial_ranges(text)
        expected = ["00010", "00011", "00012", "00013", "00025"]
        assert serials == expected


class TestMasterManager:
    """Tests for Product Master management, 7-column schema, and matching."""

    def test_master_file_exists(self):
        assert MASTER_PATH.exists(), f"Master Excel not found at {MASTER_PATH}"

    def test_master_schema_and_batch_types(self):
        master = MasterManager(MASTER_PATH)
        assert len(master.items) > 0

        for part_no, rec in master.items.items():
            assert rec["batch_type"] in ("Batch No", "Primary Batch"), (
                f"Invalid batch type '{rec['batch_type']}' for {part_no}"
            )
            assert rec["tally_name"] != "", f"Empty Tally Item Name for {part_no}"

    def test_matching_with_normalization(self):
        master = MasterManager(MASTER_PATH)
        # Should match H71-71A41 against H71-71A41-0
        m1 = master.match("H71-71A41")
        assert m1["part_no"] == "H71-71A41-0"
        assert m1["batch_type"] == "Batch No"

        # Should match 972-0DP10 against 972-0DP10
        m2 = master.match("972-0DP10")
        assert "972-0DP10" in m2["part_no"]
        assert m2["batch_type"] == "Primary Batch"

    def test_uncataloged_part_handling(self, tmp_path):
        master = MasterManager(MASTER_PATH)
        initial_new = len(master.new_items)
        rec = master.match("UNKNOWN-PART-TEST", desc="Test Part", hs_code="85371090", bcd_rate=Decimal("7.5"))
        assert rec["notes"] == "NEW"
        assert rec["tally_name"] == ""
        assert rec["batch_type"] == ""
        assert len(master.new_items) == initial_new + 1

        review_path = tmp_path / "master_review.xlsx"
        master.save_reviewed_master(review_path)
        assert review_path.exists()


class TestAccountingEngine:
    """Tests for Decimal accounting calculations, duty allocation, and residual rounding."""

    def test_batch_no_unit_expansion_and_exact_residual(self):
        # Verify that for any Batch No item, unit allocations sum exactly to line totals
        b_p = BoEParser(BOE_PATH)
        boe_data = b_p.parse()
        inv_p = InvoiceParser(INVOICE_PATH)
        inv_data = inv_p.parse()
        master = MasterManager(MASTER_PATH)

        engine = AccountingEngine(boe_data, inv_data, master)
        items = engine.process()

        # Check total items expanded
        assert len(items) == 1009, f"Expected 1009 expanded items, got {len(items)}"

        # Sum of total_inr matches invoice total * exchange rate within tolerance
        tot_inr = sum(it["total_inr"] for it in items)
        expected_inr = (inv_data["header"]["total_eur"] * boe_data["header"]["exchange_rate"]).quantize(Decimal("0.01"))
        assert abs(tot_inr - expected_inr) <= ROUNDING_TOLERANCE

        # Sum of BCD matches printed BoE duty within tolerance
        tot_bcd = sum(it["unit_bcd"] for it in items)
        assert abs(tot_bcd - boe_data["header"]["total_bcd"]) <= Decimal("1.00")

        # Sum of SWS matches printed BoE SWS within tolerance
        tot_sws = sum(it["unit_sws"] for it in items)
        assert abs(tot_sws - boe_data["header"]["total_sws"]) <= Decimal("1.00")


class TestExcelOutput:
    """Tests for generated Tally-ready Excel workbook and review copy."""

    def test_output_workbooks_and_formulas(self):
        out_file = Path(OUTPUT_FOLDER) / OUTPUT_FILENAME
        out_rev = Path(OUTPUT_FOLDER) / OUTPUT_FILENAME.replace(".xlsx", "_REVIEW.xlsx")

        assert out_file.exists(), f"Output workbook not found at {out_file}"
        assert out_rev.exists(), f"Review workbook not found at {out_rev}"

        wb = openpyxl.load_workbook(out_file, data_only=False)
        data_sheets = [s for s in wb.sheetnames if s not in ["Custom Duty 1 ", "Checks", "Serial List"]]
        assert len(data_sheets) > 0, "No invoice sheet found in generated workbook"
        assert "Custom Duty 1 " in wb.sheetnames
        assert "Checks" in wb.sheetnames
        assert "Serial List" in wb.sheetnames

        ws_p = wb[data_sheets[0]]
        # Formulas present in purchase sheet (dynamically placed below item rows)
        sum_r = ws_p.max_row - 3
        assert "=SUM(" in str(ws_p[f"I{sum_r}"].value).upper()
        assert "=SUM(" in str(ws_p[f"L{sum_r}"].value).upper()
        assert "=SUM(" in str(ws_p[f"M{sum_r}"].value).upper()
        assert f"+L{sum_r}-L{sum_r+2}" in str(ws_p[f"L{sum_r+3}"].value)

        ws_d = wb["Custom Duty 1 "]
        # Formulas present in duty sheet (dynamically placed below duty rows)
        d_sum_r = ws_d.max_row - 4
        assert "=SUM(" in str(ws_d[f"J{d_sum_r}"].value).upper()
        assert any(k in str(ws_d[f"J{d_sum_r+2}"].value) for k in ["686294.2", "1033268.9"])
        assert "=ROUND(" in str(ws_d[f"J{d_sum_r+4}"].value).upper()

        # Checks sheet status all PASS
        ws_c = wb["Checks"]
        for r in range(6, 12):
            status = ws_c.cell(r, 5).value
            assert status == "PASS", f"Check in row {r} failed: {ws_c.cell(r, 1).value} has status {status}"
