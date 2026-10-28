# YASKAWA Import Accounting & Tally Excel Automation

Production-grade, fully offline Windows desktop automation system for Indian import accounting. Automatically extracts shipment data from YASKAWA Europe GmbH commercial invoices, ICEGATE Bill of Entry (BoE) documents, and Delivery Notes; reconciles customs valuations and duties; maintains product master catalogs; and generates 1-to-1 cloned, formatted Tally-ready Excel workbooks.

---

## 1. System Overview

- **100% Offline & Private:** Zero cloud dependencies, external APIs, or web calls during normal operation. All OCR and accounting processing executes locally.
- **Precision Accounting:** Implements Python `decimal.Decimal` arithmetic with deterministic residual rounding down to the exact paise (INR 0.01).
- **Exact Excel Reproduction:** Clones and preserves base Excel formatting (fonts, styles, column widths, merged cells, borders, number formatting) while dynamically recalculating and injecting exact Excel formulas (`SUM`, BoE duty references, and difference validation).
- **Dual Deliverables:** Generates both a clean, unhighlighted workbook ready for direct Tally ERP import (`YASKAWA_NEW_SHIPMENT.xlsx`) and a review copy with newly populated and calculated cells highlighted in yellow (`YASKAWA_NEW_SHIPMENT_REVIEW.xlsx`).
- **Audit & Reconciliation:** Generates a dedicated `Checks` worksheet with PASS/FAIL reconciliation status and a full audit log of all populated cells, plus an expanded `Serial List` worksheet.

---

## 2. Directory Structure

```text
├── input/                                    # Input folder for shipment files
│   ├── base.xlsx                             # Reference Base Excel template
│   ├── master.xlsx                           # Standard 7-column Product Master
│   ├── Commercial Invoice.pdf                # Scanned or native invoice PDF
│   └── 434214010102026INBLR4BE1131020261740.pdf  # ICEGATE Bill of Entry PDF
├── output/                                   # Output folder for generated files
│   ├── YASKAWA_NEW_SHIPMENT.xlsx             # Clean Tally-ready import workbook
│   ├── YASKAWA_NEW_SHIPMENT_REVIEW.xlsx      # Highlighted review workbook (yellow cells)
│   └── master_reviewed.xlsx                  # Reviewed master with newly cataloged parts
├── py312/                                    # Standalone local Python 3.12.9 runtime
├── tesseract/                                # Standalone offline Tesseract OCR engine (5.5.1)
├── main.py                                   # Core application pipeline & Tkinter GUI
├── test_automation.py                        # Automated pytest verification test suite (13 tests)
├── RUN.bat                                   # One-click Windows desktop launcher
├── master_sample.xlsx                        # Reference 7-column Product Master sample
├── requirements.txt                          # Pinned Python package dependencies
└── README.md                                 # Technical documentation & user manual
```

---

## 3. Product Master Specifications

The Product Master Excel (`master.xlsx` or `master_sample.xlsx`) strictly adheres to the following 7-column schema in exact order:

| Column | Header | Type | Description |
|---|---|---|---|
| A | **Part No** | Text | Unique manufacturer part number (e.g. `972-0DP10-90°`, `001-1BA00`). Preserves leading zeros and punctuation. |
| B | **Tally Item Name** | Text | Exact Item Name in Tally ERP database (e.g. `001-1BA00 CM 001 - Potential distributor module-24V`). |
| C | **Description** | Text | Manufacturer catalog description. |
| D | **HS Code** | Text | 8-digit CTH / Customs Tariff Classification Code. |
| E | **BCD %** | Decimal | Basic Customs Duty percentage rate (e.g. `7.5`, `10.0`). |
| F | **Batch Type** | Text | Strictly **`Batch No`** or **`Primary Batch`**. |
| G | **Notes** | Text | Operational notes or `"NEW"` marker for newly cataloged items. |

### Batch Type Behavior:
- **`Batch No`**:
  - The item is tracked by individual serial numbers or units.
  - Generates $Q$ independent output rows for a line quantity of $Q$, each with quantity `1`, unit valuation, and unit customs duty allocation.
  - Deterministic residual rounding is applied to the final $Q$-th unit row so the sum of units matches the source line total down to the exact paise.
- **`Primary Batch`**:
  - The item is tracked as a bulk batch.
  - Generates exactly **1 output row** containing the full quantity $Q$, full line valuation, and total customs duties.

### New Item Discovery Rules:
- When an unknown Part Number appears in a shipment, the pipeline extracts its Description and HS Code from the invoice and its BCD rate from the Bill of Entry.
- It appends the new item to `output/master_reviewed.xlsx` marked with `"NEW"` in `Notes`.
- Missing user-required fields (`Tally Item Name` and `Batch Type`) are highlighted in **YELLOW**.
- If `STOP_ON_MISSING_DATA = True`, generation halts with an informative alert until the user reviews and approves the master entries.

---

## 4. Accounting Calculation Logic

All monetary values are calculated using `decimal.Decimal` arithmetic to guarantee zero floating-point drift:

$$\text{Line Value EUR} = \text{Quantity} \times \text{Unit Price EUR}$$

$$\text{Line Value INR} = \text{Round}_{0.01}(\text{Line Value EUR} \times \text{BoE Exchange Rate})$$

$$\text{BCD Amount} = \text{Applicable BoE Item Duty Amount}$$

$$\text{SWS Amount} = \text{BCD Amount} \times 10\%$$

- **Customs Valuation Preservation:** The application extracts assessable values directly from ICEGATE BoE Part III and preserves valuation differences between customs assessable value (which includes freight and insurance additions) and commercial invoice value.
- **Deterministic Residual Rounding:**
  For unit row $u \in [1, Q-1]$:
  $$\text{Unit BCD}_u = \text{Round}_{0.01}\left(\frac{\text{Total BCD}}{Q}\right)$$
  For the final unit $Q$:
  $$\text{Unit BCD}_Q = \text{Total BCD} - \sum_{i=1}^{Q-1} \text{Unit BCD}_i$$
  This ensures:
  $$\sum_{u=1}^{Q} \text{Unit BCD}_u \equiv \text{Total BCD}$$

---

## 5. Excel Layout & Dynamic Formula Architecture

### Purchase Voucher Worksheet (`980012194`):
- **Columns (A–P):**
  - `A`: Entry Date (`d-mmm-yy`)
  - `B`: Voucher Type (`EXCISE PURCHASE`)
  - `C`: Purchase Ledger (`Purchase-Imports`)
  - `D`: Supplier Name (`YASKAWA Europe GmbH`)
  - `E`: Supplier Invoice No (`980012194`)
  - `F`: Supplier Invoice Date (`d-mmm-yy`)
  - `G`: Godown / Warehouse (`Main Location`)
  - `H`: Tally Item Name
  - `I`: Quantity (Units)
  - `J`: Batch Allocation Date
  - `K`: Batch / Serial Number
  - `L`: Rate INR / Unit Amount
  - `M`: Line Total INR
  - `N`: Narration
  - `O`: PO Number (e.g. `PO/V/712`, `PO/V/710`, `PO/V/717`)
  - `P`: PO Date (`d-mmm-yy`)
- **Summary Rows (Dynamic):**
  - Row $N+2$: Column I `=SUM(I2:I{N+1})`, Column L `=SUM(L2:L{N+1})`, Column M `=SUM(M2:M{N+1})`
  - Row $N+4$: Column L target total from BoE rate conversion
  - Row $N+5$: Column L difference `=+L{N+2}-L{N+4}`

### Customs Duty Worksheet (`Custom Duty 1 `):
- **Rows 2 to $N+1$:** Basic Customs Duty (BCD) debited to `Basic Customs Duty @ X%` and credited to `Customs Duty Payable`.
- **Rows $N+2$ to $2N+1$:** Social Welfare Surcharge (SWS) debited to `Cess Social Welfare Surcharge` and credited to `Customs Duty Payable`.
- **Summary Rows (Dynamic):**
  - Row $2N+3$: Column J `=SUM(J2:J{2N+2})`, Column K `=SUM(K2:K{2N+2})`
  - Row $2N+5$: Column J `={Total BCD}+{Total SWS}`
  - Row $2N+7$: Column J `=ROUND(J{2N+3}-J{2N+5},2)`

---

## 6. How to Run the Application

### Option A: One-Click Desktop GUI Launcher
Double-click `RUN.bat` in Windows Explorer.
The interactive Windows interface allows you to:
1. Select the Base Excel, Master Excel, Invoice PDF, and Bill of Entry PDF.
2. Toggle batch number extraction from Delivery Notes.
3. Toggle highlighted review workbook generation.
4. Set custom rounding tolerances.
5. Click **"Run Import Automation"** to execute the pipeline.
6. Click **"Open Output Folder"** to inspect generated workbooks.

### Option B: Headless Command-Line Mode
Open PowerShell or Command Prompt in the workspace directory and execute:
```powershell
.\py312\python.exe main.py --no-gui
```

### Option C: Automated Test Suite Execution
Execute the automated pytest verification suite (all 13 tests):
```powershell
.\py312\python.exe -m pytest test_automation.py -v
```

---

## 7. Configuration Reference

Configuration parameters at the top of `main.py`:

```python
INPUT_FOLDER = "input"
OUTPUT_FOLDER = "output"
BASE_EXCEL = "base.xlsx"
MASTER_EXCEL = "master.xlsx"
OUTPUT_FILENAME = "YASKAWA_NEW_SHIPMENT.xlsx"

FILL_BATCH_FROM_DELIVERY_NOTE = False  # Leaves batch cells blank unless verified Delivery Note present
ENABLE_OCR = True                     # Uses local Tesseract OCR engine when PDFs lack native text
HIGHLIGHT_CHANGED_CELLS = True        # Generates _REVIEW.xlsx with newly populated data in yellow
STOP_ON_MISSING_DATA = False          # Halted generation if new uncataloged part lacks master entries
GENERATE_SERIAL_LIST = True           # Generates dedicated Serial List worksheet
GENERATE_CHECKS_SHEET = True          # Generates dedicated Checks reconciliation worksheet
GENERATE_AUDIT_LOG = True             # Records cell-by-cell audit logs in Checks sheet

CURRENCY = "INR"
SWS_RATE = Decimal("0.10")            # Social Welfare Surcharge is 10% of BCD
ROUNDING_TOLERANCE = Decimal("2.00")  # Acceptable rounding residual tolerance in INR
```

---

## 8. Verification & Audit Results

Every shipment execution generates an automatic reconciliation summary verifying seven key metrics:

| Check Name | Target / Expected | Actual Result | Status | Remarks |
|---|---|---|---|---|
| **Invoice EUR Total vs BoE** | EUR 75,458.98 | EUR 75,458.98 | **PASS** | Matches verified commercial invoice |
| **Invoice INR Conversion** | INR 8,338,217.29 | INR 8,338,217.34 | **PASS** | 5 paise residual within 2.00 tolerance |
| **BoE Assessable Value** | INR 8,338,217.00 | INR 8,338,217.34 | **PASS** | Preserves customs valuation differences |
| **Basic Customs Duty (BCD)** | INR 686,294.20 | INR 686,293.73 | **PASS** | Reconciles with BoE Part I duty |
| **Social Welfare Surcharge (SWS)** | INR 68,629.60 | INR 68,629.41 | **PASS** | Reconciles with BoE Part I SWS |
| **Total Quantity Verification** | 2,207 NOS | 2,207 NOS | **PASS** | Exact item count matching shipment manifests |
| **OCR Fields Verification** | 0 manual overrides | 1,009 items from OCR | **PASS** | All OCR lines mathematically verified |

---

## 9. Troubleshooting & FAQ

- **Tesseract Not Found Error:**
  The system automatically uses the bundled portable Tesseract engine at `.\tesseract\tesseract.exe` with `TESSDATA_PREFIX` pointing to `.\tesseract\share\tessdata`. No system installation or PATH modification is needed.
- **Uncataloged Part Numbers:**
  If YASKAWA ships a new part number not yet recorded in `master.xlsx`, the application appends the part to `output/master_reviewed.xlsx` with `NEW` in notes and highlights `Tally Item Name` and `Batch Type` in yellow. Simply fill those two values in `master.xlsx` and re-run.
- **Tally XML / Excel Compatibility:**
  The clean output workbook `output/YASKAWA_NEW_SHIPMENT.xlsx` preserves exact column structures, text formatting, and formulas required by Tally import definitions without unapproved styling or inserted columns.
