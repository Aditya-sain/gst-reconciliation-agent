# GST Reconciliation AI Agent (Finance Department)

Automated **GSTR-2B vs Tally / ERP Books** Reconciliation Engine designed for Indian Finance & Accounts departments.

---

## 🌟 Key Features

1. **Auto-Format Detection**:
   - Accepts both **Excel (.xlsx, .xls)** and **CSV** files.
   - Automatically finds the B2B inward supplies table in GSTR-2B (skips metadata/header rows).
   - Automatically identifies Tally ledger columns (`Date`, `Particulars`, `Vch Type`, `Vch No`, `Debit`, `Credit`).

2. **Multi-Tax Mode Support**:
   - **IGST Mode**: For Interstate purchases (matches GSTR-2B Integrated Tax vs Tally `Input IGST` Ledger).
   - **CGST + SGST Mode**: For Intrastate purchases (matches GSTR-2B Central + State Tax vs Tally `Input CGST/SGST` Ledgers).
   - **Total Tax Mode**: Matches overall GST liability / ITC across all tax heads.

3. **Smart Party & Invoice Matching**:
   - Normalizes party names (removes legal suffixes like "Pvt Ltd", state tags like "- PB", prefixes like "M/s").
   - Resolves banking aliases (e.g. `HSBC BANK` ⟷ `The Hongkong and Shanghai Banking Corporation Ltd`).
   - Matches by clean Invoice Number or by Party + Exact Tax Amount.

4. **Actionable Discrepancy Categorization**:
   - 🟢 **MATCHED**: Tax matches in both sheets (within ₹1 tolerance).
   - 🟡 **TAX_MISMATCH**: Party present in both sheets, but tax amount is different.
   - 🔴 **MISSING_IN_GSTR2B**: Tax booked in Tally, but missing in GSTR-2B (Vendor GSTR-1 not filed/late; ITC cannot be claimed).
   - 🔵 **MISSING_IN_TALLY**: Tax appears in GSTR-2B, but missing in Tally (Bill unbooked; accounts team needs to enter bill).
   - 🟣 **CREDIT_REVERSAL**: Credit entries in Tally ledger (tax reversals, credit notes, or sale entries).

5. **Dual Interface**:
   - **Interactive Web App** with Tailwind CSS UI, KPI dashboard, filter tabs, and one-click demo data.
   - **Python CLI Tool** for automated scripts / scheduled pipelines.
   - **Multi-Tab Excel Report** (`GST_Reconciliation_Report.xlsx`) with corporate formatting.

---

## 🚀 Quick Start

### 1. Launch Interactive Web App
Double-click `run_agent.bat` or run in terminal:
```bash
python app.py
```
Open your browser at `http://localhost:5000`.
- Click **"Load October 2026 Sample Data"** to immediately test with the sample sheets.
- Or upload your own GSTR-2B and Tally files and click **"Reconcile Now"**.
- Click **"Download Excel Report"** to export formatted spreadsheets.

### 2. Run via Command Line (CLI)
```bash
# IGST Reconciliation
python gst_recon_agent.py --gstr2b sample_data/gstr2b_oct_2026.csv --tally sample_data/tally_oct_2026.csv --mode igst --output GST_Reconciliation_Report.xlsx

# CGST + SGST Reconciliation
python gst_recon_agent.py --gstr2b path/to/2b.xlsx --tally path/to/tally.xlsx --mode cgst_sgst

# Total Tax Reconciliation
python gst_recon_agent.py --gstr2b path/to/2b.xlsx --tally path/to/tally.xlsx --mode total
```

---

## 📊 Summary of October 2026 Sample Data Analysis

| Metric | Amount (₹) |
|---|---|
| **Total Tax in GSTR-2B (Portal)** | ₹ 5,75,607.89 |
| **Total Input Tax Claimed in Tally** | ₹ 11,66,286.09 |
| **Tally Credit / Reversal Adjustments** | ₹ 33,302.16 |
| **Net Tally Tax Booked** | ₹ 11,32,983.93 |
| **Net Difference (2B - Tally)** | **-₹ 5,57,376.04** |

### Key Findings & Action Items:
1. **Orane Consulting Private Limited**: Tally booked 2 vouchers of ₹4,95,288 each (`OCTI26270334` on 05-Oct and `OCTI26270433` on 27-Oct). However, GSTR-2B only contains `OCTI26270433`. Follow up with vendor regarding invoice `OCTI26270334` or check for duplicate booking in Tally.
2. **Care Health Insurance Limited**: Tally booked ₹1,01,926.78 under voucher `C4778049-1`, but GSTR-2B only shows invoice `C4778049-3` with tax ₹3,203.78. Difference is ₹98,723.00.
3. **Associated Printers**: 4 bills in GSTR-2B (Total ₹4,602.02). Tally only booked 2 bills (`AP/365` for ₹1,800 and `AP/399` for ₹1,404). Bills `AP/376` (₹1,297.22) and `AP/429` (₹100.80) are unbooked in Tally!
4. **Google India Pvt Ltd & Google Cloud India**: Bills present in GSTR-2B (₹432.00 and ₹1,513.97) but not booked under these vendor names in Tally.
5. **Lagozon Technologies Private Limited**: Credit entry of ₹33,302.16 in Tally Input IGST ledger (Voucher `PB/26-27/0063`). This is an output / sale credit adjustment, not inward ITC.
6. **Matched Vendors**:
   - `Madhurima Gupta`: ₹21,357.00 (Matched)
   - `Netplus Broadband Services`: ₹3,074.82 (Matched across 2 invoices)
   - `Rudrarya International LLP`: ₹43,746.06 (Matched across 2 invoices)
