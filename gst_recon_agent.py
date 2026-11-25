"""
GST Reconciliation Agent for Finance Departments
Matches GSTR-2B (GST Portal) with Books / Tally Ledgers.
Supports IGST, CGST+SGST, and Total Tax reconciliation.
Handles Excel (.xlsx, .xls) and CSV files.
"""

import os
import re
import csv
import difflib
import argparse
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common aliases in Indian corporate accounting
KNOWN_BANK_ALIASES = {
    'THE HONGKONG AND SHANGHAI BANKING CORPORATION': 'HSBC',
    'THE HONGKONG AND SHANGHAI BANKING CORPORATION LTD': 'HSBC',
    'HONGKONG AND SHANGHAI BANKING CORPORATION': 'HSBC',
    'HSBC BANK': 'HSBC',
    'CREDIT CARD HSBC BANK': 'HSBC',
    'CREDIT CARD HSBC': 'HSBC',
    'HDFC BANK LIMITED': 'HDFC',
    'HDFC BANK': 'HDFC',
    'ICICI BANK LIMITED': 'ICICI',
    'ICICI BANK': 'ICICI',
    'STATE BANK OF INDIA': 'SBI',
    'PUNJAB NATIONAL BANK': 'PNB',
    'AXIS BANK LIMITED': 'AXIS',
    'AXIS BANK': 'AXIS',
    'KOTAK MAHINDRA BANK LIMITED': 'KOTAK',
    'KOTAK MAHINDRA BANK': 'KOTAK',
    'BANK OF BARODA': 'BOB',
}

def normalize_party_name(name):
    """
    Normalizes party/vendor names for fuzzy matching.
    Removes common business entity terms, prefixes, state tags, and punctuation.
    """
    if not isinstance(name, str):
        return ""
    
    s = name.upper().strip()
    
    # Check known aliases first
    for alias_key, alias_val in KNOWN_BANK_ALIASES.items():
        if alias_key in s:
            return alias_val
            
    # Remove common prefixes
    s = re.sub(r'^(M/S\.?|MS\.?|MR\.?|MRS\.?|SH\.?|SHRI\.?)\s+', '', s)
    
    # Remove state or branch suffixes like "- PUNJAB", "- PB", "- HRY", "(PB)"
    s = re.sub(r'[\-\–]\s*(PUNJAB|PB|HARYANA|HRY|DELHI|DL|CHANDIGARH|CHD|MOHALI)\b', '', s)
    s = re.sub(r'\((PUNJAB|PB|HARYANA|HRY|DELHI|DL|CHANDIGARH|CHD|MOHALI)\)', '', s)
    
    # Remove common corporate suffixes
    terms_to_remove = [
        r'\bPRIVATE\s+LIMITED\b', r'\bPVT\.?\s*LTD\.?\b', r'\bLIMITED\b', r'\bLTD\.?\b',
        r'\bLLP\b', r'\bINC\.?\b', r'\bCORP\.?\b', r'\bCORPORATION\b', r'\bCO\.?\b',
        r'\bSERVICES\b', r'\bSOLUTIONS\b', r'\bENTERPRISES\b', r'\bTRADERS\b'
    ]
    for term in terms_to_remove:
        s = re.sub(term, '', s)
        
    # Remove non-alphanumeric characters (keep single spaces)
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_invoice_no(inv):
    """
    Cleans invoice number by stripping leading zeros, spaces, slashes, dashes.
    """
    if not isinstance(inv, str):
        return ""
    return re.sub(r'[^A-Z0-9]', '', inv.upper())

class GSTReconciliationAgent:
    def __init__(self, mode='igst', tolerance=1.0):
        """
        mode: 'igst', 'cgst_sgst', or 'total'
        tolerance: amount tolerance in Rupees (default ₹1.00 for rounding)
        """
        self.mode = mode.lower()
        self.tolerance = tolerance
        
    def parse_gstr2b(self, file_path):
        """
        Parses GSTR-2B file (CSV or Excel).
        Auto-locates B2B inward supplies table.
        """
        if file_path.endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(file_path)
            b2b_sheet = None
            for sheet in xl.sheet_names:
                if 'b2b' in sheet.lower() and 'cdnr' not in sheet.lower() and 'a' not in sheet.lower():
                    b2b_sheet = sheet
                    break
            if not b2b_sheet:
                b2b_sheet = xl.sheet_names[0]
            raw_df = pd.read_excel(file_path, sheet_name=b2b_sheet, header=None)
            rows = raw_df.values.tolist()
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                rows = list(csv.reader(f))
                
        # Find header index containing 'GSTIN of supplier'
        header_idx = -1
        for i, row in enumerate(rows):
            row_str = " ".join([str(c) for c in row if pd.notna(c)]).lower()
            if 'gstin of supplier' in row_str:
                header_idx = i
                break
                
        if header_idx == -1:
            raise ValueError(f"Could not find GSTR-2B B2B table header in {file_path}")
            
        records = []
        for r in rows[header_idx + 2:]:
            if not r or len(r) < 10:
                continue
            first_cell = str(r[0]).strip() if pd.notna(r[0]) else ""
            if any(term in first_cell.lower() for term in ['amendments', 'goods and services tax', 'table', 'summary']):
                break
            if not first_cell:
                continue
                
            # Valid GSTIN check: 15 chars or has digit in first 2 characters
            if len(first_cell) != 15 and not (len(first_cell) >= 2 and first_cell[:2].isdigit()):
                continue
                
            def get_float(val):
                if pd.isna(val):
                    return 0.0
                val_str = str(val).replace(',', '').strip()
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0

            gstin = str(r[0]).strip()
            trade_name = str(r[1]).strip() if pd.notna(r[1]) else ""
            inv_no = str(r[2]).strip() if pd.notna(r[2]) else ""
            inv_type = str(r[3]).strip() if len(r) > 3 and pd.notna(r[3]) else "Regular"
            inv_date = str(r[4]).strip() if len(r) > 4 and pd.notna(r[4]) else ""
            inv_val = get_float(r[5]) if len(r) > 5 else 0.0
            pos = str(r[6]).strip() if len(r) > 6 and pd.notna(r[6]) else ""
            rcm = str(r[7]).strip() if len(r) > 7 and pd.notna(r[7]) else "No"
            taxable_val = get_float(r[8]) if len(r) > 8 else 0.0
            igst = get_float(r[9]) if len(r) > 9 else 0.0
            cgst = get_float(r[10]) if len(r) > 10 else 0.0
            sgst = get_float(r[11]) if len(r) > 11 else 0.0
            cess = get_float(r[12]) if len(r) > 12 else 0.0
            
            # Select target tax based on reconciliation mode
            if self.mode == 'igst':
                target_tax = igst
            elif self.mode == 'cgst_sgst':
                target_tax = cgst + sgst
            else:
                target_tax = igst + cgst + sgst
                
            records.append({
                'source': 'GSTR-2B',
                'gstin': gstin,
                'party_name': trade_name,
                'party_norm': normalize_party_name(trade_name),
                'invoice_no': inv_no,
                'invoice_no_clean': clean_invoice_no(inv_no),
                'invoice_date': inv_date,
                'invoice_type': inv_type,
                'taxable_value': taxable_val,
                'igst': igst,
                'cgst': cgst,
                'sgst': sgst,
                'cess': cess,
                'total_tax': igst + cgst + sgst,
                'target_tax': target_tax,
                'invoice_value': inv_val,
                'rcm': rcm
            })
        return pd.DataFrame(records)

    def parse_tally(self, file_path):
        """
        Parses Tally Ledger file (CSV or Excel).
        Extracts Date, Particulars (Party Name), Vch Type, Vch No, Debit, Credit.
        """
        if file_path.endswith(('.xlsx', '.xls')):
            raw_df = pd.read_excel(file_path, header=None)
            rows = raw_df.values.tolist()
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                rows = list(csv.reader(f))
                
        header_idx = -1
        for i, row in enumerate(rows):
            row_str = " ".join([str(c) for c in row if pd.notna(c)]).lower()
            if 'particulars' in row_str and ('debit' in row_str or 'vch' in row_str):
                header_idx = i
                break
                
        if header_idx == -1:
            raise ValueError(f"Could not find Tally ledger table header in {file_path}")
            
        header = [str(c).strip().lower() if pd.notna(c) else "" for c in rows[header_idx]]
        
        col_map = {}
        for j, h in enumerate(header):
            if 'date' in h:
                col_map['date'] = j
            elif 'particular' in h:
                col_map['particulars'] = j
            elif 'vch type' in h:
                col_map['vch_type'] = j
            elif 'vch no' in h:
                col_map['vch_no'] = j
            elif 'debit' in h:
                col_map['debit'] = j
            elif 'credit' in h:
                col_map['credit'] = j
                
        def get_val(r, key, default=""):
            idx = col_map.get(key)
            if idx is not None and idx < len(r) and pd.notna(r[idx]):
                return r[idx]
            return default

        records = []
        for r in rows[header_idx + 1:]:
            if not r or len(r) < 4:
                continue
            row_text = " ".join([str(c) for c in r if pd.notna(c)]).lower()
            if 'opening balance' in row_text or 'closing balance' in row_text:
                continue
                
            date = str(get_val(r, 'date', '')).strip()
            p_idx = col_map.get('particulars', 1)
            party = str(r[p_idx]).strip() if p_idx < len(r) and pd.notna(r[p_idx]) else ""
            
            # If sub-column with Dr/Cr exists before party name
            if party in ['Dr', 'Cr'] and p_idx + 1 < len(r) and pd.notna(r[p_idx + 1]):
                party = str(r[p_idx + 1]).strip()
                
            if not party or party.lower() in ['total', 'grand total']:
                continue
                
            vch_type = str(get_val(r, 'vch_type', '')).strip()
            vch_no = str(get_val(r, 'vch_no', '')).strip()
            
            def parse_num(val):
                if pd.isna(val):
                    return 0.0
                s = str(val).replace(',', '').strip()
                try:
                    return float(s)
                except ValueError:
                    return 0.0

            debit = parse_num(get_val(r, 'debit', 0.0))
            credit = parse_num(get_val(r, 'credit', 0.0))
            
            # In Input Tax Ledgers:
            # Debit = Input Tax claimed on purchase
            # Credit = Tax reversal, credit note, or output tax adjustment
            is_credit_reversal = (credit > 0 and debit == 0)
            tax_amount = debit if debit > 0 else -credit
            
            records.append({
                'source': 'Tally',
                'date': date,
                'party_name': party,
                'party_norm': normalize_party_name(party),
                'vch_type': vch_type,
                'vch_no': vch_no,
                'vch_no_clean': clean_invoice_no(vch_no),
                'debit': debit,
                'credit': credit,
                'tax_amount': tax_amount,
                'is_credit_reversal': is_credit_reversal
            })
            
        return pd.DataFrame(records)

    def reconcile(self, gstr2b_file, tally_file):
        """
        Executes full reconciliation between GSTR-2B and Tally.
        Returns:
            - summary: high-level dashboard metrics
            - party_summary: party-level reconciliation table
            - invoice_details: invoice-level matching details
        """
        df_2b = self.parse_gstr2b(gstr2b_file)
        df_tally = self.parse_tally(tally_file)
        
        # Filter GSTR-2B based on target tax > 0
        df_2b_filtered = df_2b[df_2b['target_tax'] > 0].copy()
        
        # Aggregate GSTR-2B by Party
        p2b = df_2b_filtered.groupby(['party_norm']).agg(
            party_name=('party_name', 'first'),
            gstin=('gstin', lambda x: ", ".join(sorted(set(x)))),
            total_tax_2b=('target_tax', 'sum'),
            invoice_count_2b=('invoice_no', 'count'),
            invoices_2b=('invoice_no', lambda x: ", ".join(list(x)))
        ).reset_index()
        
        # Separate normal Tally debits from credit reversals
        df_tally_debit = df_tally[~df_tally['is_credit_reversal']].copy()
        df_tally_credit = df_tally[df_tally['is_credit_reversal']].copy()
        
        ptally = df_tally_debit.groupby(['party_norm']).agg(
            party_name=('party_name', lambda x: " / ".join(sorted(set(x)))),
            total_tax_tally=('debit', 'sum'),
            voucher_count_tally=('vch_no', 'count'),
            vouchers_tally=('vch_no', lambda x: ", ".join(list(x)))
        ).reset_index()
        
        p2b_list = p2b.to_dict('records')
        ptally_list = ptally.to_dict('records')
        
        matched_2b_indices = set()
        matched_tally_indices = set()
        party_recon_rows = []
        
        # Pass 1: Exact normalized name / alias match
        for i, r2b in enumerate(p2b_list):
            for j, rt in enumerate(ptally_list):
                if j in matched_tally_indices:
                    continue
                if r2b['party_norm'] and rt['party_norm'] and r2b['party_norm'] == rt['party_norm']:
                    matched_2b_indices.add(i)
                    matched_tally_indices.add(j)
                    party_recon_rows.append(self._build_party_row(r2b, rt))
                    break

        # Pass 2: Fuzzy match on party name (similarity >= 0.70)
        for i, r2b in enumerate(p2b_list):
            if i in matched_2b_indices:
                continue
            best_score = 0.0
            best_j = -1
            for j, rt in enumerate(ptally_list):
                if j in matched_tally_indices:
                    continue
                score = difflib.SequenceMatcher(None, r2b['party_norm'], rt['party_norm']).ratio()
                if score > best_score:
                    best_score = score
                    best_j = j
                    
            if best_score >= 0.70 and best_j != -1:
                matched_2b_indices.add(i)
                matched_tally_indices.add(best_j)
                party_recon_rows.append(self._build_party_row(r2b, ptally_list[best_j], fuzzy_score=best_score))

        # Pass 3: Remaining in GSTR-2B only (Missing in Tally)
        for i, r2b in enumerate(p2b_list):
            if i not in matched_2b_indices:
                party_recon_rows.append({
                    'party_name_2b': r2b['party_name'],
                    'party_name_tally': '--- Not Found in Tally ---',
                    'gstin': r2b['gstin'],
                    'tax_gstr2b': r2b['total_tax_2b'],
                    'tax_tally': 0.0,
                    'difference': r2b['total_tax_2b'],
                    'status': 'MISSING_IN_TALLY',
                    'status_label': 'Missing in Tally (Bill Unrecorded)',
                    'reason': f"Tax of ₹{r2b['total_tax_2b']:,.2f} is available in GSTR-2B but unrecorded in Tally books. (Action: Accounts team should record invoice to claim ITC).",
                    'invoices_2b': r2b['invoices_2b'],
                    'vouchers_tally': ''
                })

        # Pass 4: Remaining in Tally only (Missing in GSTR-2B)
        for j, rt in enumerate(ptally_list):
            if j not in matched_tally_indices:
                party_recon_rows.append({
                    'party_name_2b': '--- Not Found in GSTR-2B ---',
                    'party_name_tally': rt['party_name'],
                    'gstin': '---',
                    'tax_gstr2b': 0.0,
                    'tax_tally': rt['total_tax_tally'],
                    'difference': -rt['total_tax_tally'],
                    'status': 'MISSING_IN_GSTR2B',
                    'status_label': 'Missing in GSTR-2B (Vendor GSTR-1 Pending)',
                    'reason': f"Tax of ₹{rt['total_tax_tally']:,.2f} is booked in Tally but not reflected in GSTR-2B. (Action: Supplier has not filed GSTR-1 or filed with incorrect GSTIN; follow up with vendor).",
                    'invoices_2b': '',
                    'vouchers_tally': rt['vouchers_tally']
                })
                
        # Pass 5: Credit Reversals / Sale Entries in Tally ledger
        for idx, rc in df_tally_credit.iterrows():
            party_recon_rows.append({
                'party_name_2b': '---',
                'party_name_tally': rc['party_name'],
                'gstin': '---',
                'tax_gstr2b': 0.0,
                'tax_tally': -rc['credit'],
                'difference': rc['credit'],
                'status': 'CREDIT_REVERSAL',
                'status_label': 'Credit / Tax Reversal in Tally',
                'reason': f"Credit entry of ₹{rc['credit']:,.2f} in Tally ({rc['vch_type']} - Vch {rc['vch_no']}). Represents a tax reversal or sales adjustment, not an inward purchase supply.",
                'invoices_2b': '',
                'vouchers_tally': rc['vch_no']
            })

        df_party_summary = pd.DataFrame(party_recon_rows)
        
        # Sort priority
        priority_map = {'MISSING_IN_GSTR2B': 1, 'MISSING_IN_TALLY': 2, 'TAX_MISMATCH': 3, 'CREDIT_REVERSAL': 4, 'MATCHED': 5}
        df_party_summary['priority'] = df_party_summary['status'].map(priority_map).fillna(9)
        df_party_summary = df_party_summary.sort_values(by=['priority', 'difference'], ascending=[True, False]).drop(columns=['priority'])
        
        # Calculate summary metrics
        total_2b_tax = df_2b_filtered['target_tax'].sum()
        total_tally_tax = df_tally_debit['debit'].sum()
        total_tally_credit = df_tally_credit['credit'].sum()
        net_difference = total_2b_tax - (total_tally_tax - total_tally_credit)
        
        matched_count = len(df_party_summary[df_party_summary['status'] == 'MATCHED'])
        mismatch_count = len(df_party_summary[df_party_summary['status'] == 'TAX_MISMATCH'])
        missing_2b_count = len(df_party_summary[df_party_summary['status'] == 'MISSING_IN_GSTR2B'])
        missing_tally_count = len(df_party_summary[df_party_summary['status'] == 'MISSING_IN_TALLY'])
        
        summary = {
            'reconciliation_mode': self.mode.upper(),
            'total_gstr2b_tax': round(total_2b_tax, 2),
            'total_tally_debit_tax': round(total_tally_tax, 2),
            'total_tally_credit_tax': round(total_tally_credit, 2),
            'net_tally_tax': round(total_tally_tax - total_tally_credit, 2),
            'net_difference': round(net_difference, 2),
            'total_parties_analyzed': len(df_party_summary),
            'count_matched': matched_count,
            'count_mismatch': mismatch_count,
            'count_missing_in_gstr2b': missing_2b_count,
            'count_missing_in_tally': missing_tally_count
        }
        
        invoice_details = self._reconcile_invoices(df_2b_filtered, df_tally)
        
        return {
            'summary': summary,
            'party_summary': df_party_summary,
            'invoice_details': invoice_details,
            'df_2b_raw': df_2b_filtered,
            'df_tally_raw': df_tally
        }

    def _build_party_row(self, r2b, rt, fuzzy_score=1.0):
        tax_2b = r2b['total_tax_2b']
        tax_tally = rt['total_tax_tally']
        diff = round(tax_2b - tax_tally, 2)
        
        if abs(diff) <= self.tolerance:
            status = 'MATCHED'
            status_label = 'Matched (Exact / Within ₹1)'
            reason = f"Tax amounts match exactly across both records (₹{tax_2b:,.2f})."
        else:
            status = 'TAX_MISMATCH'
            status_label = 'Tax Amount Mismatch'
            if diff > 0:
                reason = f"GSTR-2B tax (₹{tax_2b:,.2f}) exceeds Tally books (₹{tax_tally:,.2f}). Unbooked ITC variance: ₹{diff:,.2f}."
            else:
                reason = f"Tally tax claimed (₹{tax_tally:,.2f}) exceeds GSTR-2B reflection (₹{tax_2b:,.2f}). Variance: ₹{abs(diff):,.2f}. (Action: Reconcile with supplier or adjust ITC in GSTR-3B)."
                
        return {
            'party_name_2b': r2b['party_name'],
            'party_name_tally': rt['party_name'],
            'gstin': r2b['gstin'],
            'tax_gstr2b': tax_2b,
            'tax_tally': tax_tally,
            'difference': diff,
            'status': status,
            'status_label': status_label,
            'reason': reason,
            'invoices_2b': r2b['invoices_2b'],
            'vouchers_tally': rt['vouchers_tally']
        }

    def _reconcile_invoices(self, df_2b, df_tally):
        """
        Invoice-by-invoice detailed comparison
        """
        rows = []
        tally_used = set()
        
        for _, r2 in df_2b.iterrows():
            inv_clean = r2['invoice_no_clean']
            amt_2b = r2['target_tax']
            party_norm = r2['party_norm']
            
            match_found = False
            for j, rt in df_tally.iterrows():
                if j in tally_used:
                    continue
                vch_clean = rt['vch_no_clean']
                amt_tally = rt['debit']
                
                # Condition A: exact invoice number match
                if inv_clean and vch_clean and inv_clean == vch_clean:
                    diff = round(amt_2b - amt_tally, 2)
                    status = 'MATCHED' if abs(diff) <= self.tolerance else 'AMOUNT_MISMATCH'
                    rows.append({
                        'party_2b': r2['party_name'],
                        'party_tally': rt['party_name'],
                        'invoice_2b': r2['invoice_no'],
                        'vch_no_tally': rt['vch_no'],
                        'tax_2b': amt_2b,
                        'tax_tally': amt_tally,
                        'difference': diff,
                        'match_type': 'Invoice Number Match',
                        'status': status
                    })
                    tally_used.add(j)
                    match_found = True
                    break
                    
            if not match_found:
                # Condition B: match by Party + Exact Tax Amount
                for j, rt in df_tally.iterrows():
                    if j in tally_used:
                        continue
                    if abs(r2['target_tax'] - rt['debit']) <= self.tolerance and (party_norm == rt['party_norm']):
                        rows.append({
                            'party_2b': r2['party_name'],
                            'party_tally': rt['party_name'],
                            'invoice_2b': r2['invoice_no'],
                            'vch_no_tally': rt['vch_no'],
                            'tax_2b': amt_2b,
                            'tax_tally': rt['debit'],
                            'difference': 0.0,
                            'match_type': 'Party + Amount Match (Vch No differs)',
                            'status': 'MATCHED'
                        })
                        tally_used.add(j)
                        match_found = True
                        break
                        
            if not match_found:
                rows.append({
                    'party_2b': r2['party_name'],
                    'party_tally': '---',
                    'invoice_2b': r2['invoice_no'],
                    'vch_no_tally': '---',
                    'tax_2b': amt_2b,
                    'tax_tally': 0.0,
                    'difference': amt_2b,
                    'match_type': 'Not in Tally',
                    'status': 'MISSING_IN_TALLY'
                })
                
        # Remaining Tally entries
        for j, rt in df_tally.iterrows():
            if j not in tally_used:
                rows.append({
                    'party_2b': '---',
                    'party_tally': rt['party_name'],
                    'invoice_2b': '---',
                    'vch_no_tally': rt['vch_no'],
                    'tax_2b': 0.0,
                    'tax_tally': rt['debit'] if rt['debit'] > 0 else -rt['credit'],
                    'difference': -rt['debit'] if rt['debit'] > 0 else rt['credit'],
                    'match_type': 'Not in GSTR-2B',
                    'status': 'CREDIT_REVERSAL' if rt['credit'] > 0 else 'MISSING_IN_GSTR2B'
                })
                
        return pd.DataFrame(rows)

    def export_excel_report(self, recon_result, output_path='GST_Reconciliation_Report.xlsx'):
        """
        Exports a multi-tab, beautifully styled Excel Reconciliation Report with corporate formatting.
        """
        summary = recon_result['summary']
        df_party = recon_result['party_summary']
        df_inv = recon_result['invoice_details']
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        navy_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        title_font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
        sub_font = Font(name="Calibri", size=10, italic=True, color="475569")
        bold_font = Font(name="Calibri", size=11, bold=True, color="000000")
        regular_font = Font(name="Calibri", size=11, color="000000")
        thin_border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )
        
        status_styles = {
            'MATCHED': (PatternFill(start_color="DCFCE7", fill_type="solid"), Font(color="166534", bold=True)),
            'TAX_MISMATCH': (PatternFill(start_color="FEF3C7", fill_type="solid"), Font(color="92400E", bold=True)),
            'MISSING_IN_GSTR2B': (PatternFill(start_color="FFE4E6", fill_type="solid"), Font(color="9F1239", bold=True)),
            'MISSING_IN_TALLY': (PatternFill(start_color="DBEAFE", fill_type="solid"), Font(color="1E40AF", bold=True)),
            'CREDIT_REVERSAL': (PatternFill(start_color="F3E8FF", fill_type="solid"), Font(color="6B21A8", bold=True))
        }

        # ----------------------------------------------------
        # SHEET 1: DASHBOARD SUMMARY
        # ----------------------------------------------------
        ws1 = wb.create_sheet(title="Dashboard Summary")
        ws1.views.sheetView[0].showGridLines = True
        
        ws1.merge_cells("A1:E1")
        ws1["A1"] = "GST RECONCILIATION EXECUTIVE REPORT"
        ws1["A1"].font = title_font
        
        ws1.merge_cells("A2:E2")
        ws1["A2"] = f"Generated on: {datetime.now().strftime('%d %B %Y, %I:%M %p')} | Mode: {summary['reconciliation_mode']}"
        ws1["A2"].font = sub_font
        
        summary_rows = [
            ("Reconciliation Mode", summary['reconciliation_mode'], ""),
            ("Total Tax in GSTR-2B (GST Portal)", summary['total_gstr2b_tax'], "₹ #,##0.00"),
            ("Total Input Tax Claimed in Tally (Purchases)", summary['total_tally_debit_tax'], "₹ #,##0.00"),
            ("Tally Credit Reversals / Sale Adjustments", summary['total_tally_credit_tax'], "₹ #,##0.00"),
            ("Net Tax Claimed in Books / Tally", summary['net_tally_tax'], "₹ #,##0.00"),
            ("Net Difference (GSTR-2B minus Tally)", summary['net_difference'], "₹ #,##0.00"),
            ("Total Parties Analyzed", summary['total_parties_analyzed'], "#,##0"),
            ("Parties Matched (Exact / Within ₹1)", summary['count_matched'], "#,##0"),
            ("Parties with Tax Amount Mismatch", summary['count_mismatch'], "#,##0"),
            ("Parties Missing in GSTR-2B (Vendor GSTR-1 Pending)", summary['count_missing_in_gstr2b'], "#,##0"),
            ("Parties Missing in Tally (Unrecorded Invoices)", summary['count_missing_in_tally'], "#,##0"),
        ]
        
        start_row = 4
        ws1.cell(row=start_row, column=1, value="Metric / Category").font = white_bold
        ws1.cell(row=start_row, column=1).fill = navy_header
        ws1.cell(row=start_row, column=2, value="Value / Amount").font = white_bold
        ws1.cell(row=start_row, column=2).fill = navy_header
        
        for idx, (label, val, fmt) in enumerate(summary_rows, start=start_row + 1):
            c1 = ws1.cell(row=idx, column=1, value=label)
            c2 = ws1.cell(row=idx, column=2, value=val)
            c1.font = bold_font if "Net Difference" in label else regular_font
            c2.font = bold_font if "Net Difference" in label else regular_font
            c1.border = thin_border
            c2.border = thin_border
            if fmt:
                c2.number_format = fmt
                
        # ----------------------------------------------------
        # SHEET 2: PARTY RECONCILIATION
        # ----------------------------------------------------
        ws2 = wb.create_sheet(title="Party Reconciliation")
        ws2.views.sheetView[0].showGridLines = True
        
        headers2 = [
            "Status", "Party Name (Tally)", "Party Name (GSTR-2B)", "Supplier GSTIN",
            "Tax in GSTR-2B (₹)", "Tax in Tally (₹)", "Difference (₹)", 
            "Explanation & Action Required", "Invoices in GSTR-2B", "Vouchers in Tally"
        ]
        
        for col_idx, h in enumerate(headers2, start=1):
            c = ws2.cell(row=1, column=col_idx, value=h)
            c.font = white_bold
            c.fill = navy_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        for row_idx, r in df_party.iterrows():
            curr_row = row_idx + 2
            st_key = r['status']
            st_fill, st_font = status_styles.get(st_key, (None, None))
            
            c_status = ws2.cell(row=curr_row, column=1, value=r['status_label'])
            if st_fill:
                c_status.fill = st_fill
                c_status.font = st_font
                
            ws2.cell(row=curr_row, column=2, value=r['party_name_tally']).font = regular_font
            ws2.cell(row=curr_row, column=3, value=r['party_name_2b']).font = regular_font
            ws2.cell(row=curr_row, column=4, value=r['gstin']).font = regular_font
            
            c_2b = ws2.cell(row=curr_row, column=5, value=r['tax_gstr2b'])
            c_2b.number_format = "₹ #,##0.00"
            c_2b.font = regular_font
            
            c_tally = ws2.cell(row=curr_row, column=6, value=r['tax_tally'])
            c_tally.number_format = "₹ #,##0.00"
            c_tally.font = regular_font
            
            c_diff = ws2.cell(row=curr_row, column=7, value=r['difference'])
            c_diff.number_format = "₹ #,##0.00"
            c_diff.font = bold_font if abs(r['difference']) > self.tolerance else regular_font
            
            ws2.cell(row=curr_row, column=8, value=r['reason']).font = regular_font
            ws2.cell(row=curr_row, column=9, value=r['invoices_2b']).font = regular_font
            ws2.cell(row=curr_row, column=10, value=r['vouchers_tally']).font = regular_font
            
            for col_c in range(1, 11):
                ws2.cell(row=curr_row, column=col_c).border = thin_border

        # ----------------------------------------------------
        # SHEET 3: INVOICE DRILLDOWN
        # ----------------------------------------------------
        ws3 = wb.create_sheet(title="Invoice Drilldown")
        ws3.views.sheetView[0].showGridLines = True
        
        headers3 = [
            "Status", "Party (Tally)", "Party (GSTR-2B)", "Voucher No (Tally)", "Invoice No (2B)",
            "Tax in 2B (₹)", "Tax in Tally (₹)", "Difference (₹)", "Match Logic"
        ]
        
        for col_idx, h in enumerate(headers3, start=1):
            c = ws3.cell(row=1, column=col_idx, value=h)
            c.font = white_bold
            c.fill = navy_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        for row_idx, r in df_inv.iterrows():
            curr_row = row_idx + 2
            st_key = r['status']
            st_fill, st_font = status_styles.get(st_key, (None, None))
            
            c_status = ws3.cell(row=curr_row, column=1, value=st_key)
            if st_fill:
                c_status.fill = st_fill
                c_status.font = st_font
                
            ws3.cell(row=curr_row, column=2, value=r['party_tally']).font = regular_font
            ws3.cell(row=curr_row, column=3, value=r['party_2b']).font = regular_font
            ws3.cell(row=curr_row, column=4, value=r['vch_no_tally']).font = regular_font
            ws3.cell(row=curr_row, column=5, value=r['invoice_2b']).font = regular_font
            
            c_2b = ws3.cell(row=curr_row, column=6, value=r['tax_2b'])
            c_2b.number_format = "₹ #,##0.00"
            c_2b.font = regular_font
            
            c_tally = ws3.cell(row=curr_row, column=7, value=r['tax_tally'])
            c_tally.number_format = "₹ #,##0.00"
            c_tally.font = regular_font
            
            c_diff = ws3.cell(row=curr_row, column=8, value=r['difference'])
            c_diff.number_format = "₹ #,##0.00"
            c_diff.font = bold_font if abs(r['difference']) > self.tolerance else regular_font
            
            ws3.cell(row=curr_row, column=9, value=r['match_type']).font = regular_font
            
            for col_c in range(1, 10):
                ws3.cell(row=curr_row, column=col_c).border = thin_border

        # Adjust column widths
        for ws in [ws1, ws2, ws3]:
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        val_str = str(cell.value)
                        if len(val_str) > max_len and len(val_str) < 60:
                            max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        wb.save(output_path)
        print(f"Corporate Excel report saved to: {output_path}")
        return output_path

def main():
    parser = argparse.ArgumentParser(description="GST Reconciliation Agent: Match GSTR-2B with Tally Ledger")
    parser.add_argument("--gstr2b", default="sample_data/gstr2b_oct_2026.csv", help="Path to GSTR-2B file (xlsx or csv)")
    parser.add_argument("--tally", default="sample_data/tally_oct_2026.csv", help="Path to Tally ledger file (xlsx or csv)")
    parser.add_argument("--mode", default="igst", choices=["igst", "cgst_sgst", "total"], help="Reconciliation tax mode")
    parser.add_argument("--output", default="GST_Reconciliation_Report.xlsx", help="Output Excel report path")
    
    args = parser.parse_args()
    
    agent = GSTReconciliationAgent(mode=args.mode)
    results = agent.reconcile(args.gstr2b, args.tally)
    agent.export_excel_report(results, args.output)
    
    print("\n" + "="*80)
    print("                 GST RECONCILIATION SUMMARY DASHBOARD")
    print("="*80)
    for k, v in results['summary'].items():
        print(f"  {k.replace('_', ' ').title():<40}: {v}")
    print("="*80)

if __name__ == '__main__':
    main()
