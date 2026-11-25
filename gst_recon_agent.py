"""
GST Reconciliation Agent for Finance Departments
Exact invoice/voucher level reconciliation between GSTR-2B and Books / Tally Ledgers.
Supports automatic tax head detection (CGST, SGST, IGST), invoice prefix tolerance, and clean discrepancy reporting.
"""

import os
import re
import csv
import argparse
import pandas as pd
import numpy as np
from datetime import datetime
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    if not isinstance(name, str):
        return ""
    s = name.upper().strip()
    for alias_key, alias_val in KNOWN_BANK_ALIASES.items():
        if alias_key in s:
            return alias_val
    s = re.sub(r'^(M/S\.?|MS\.?|MR\.?|MRS\.?|SH\.?|SHRI\.?)\s+', '', s)
    s = re.sub(r'[\-\–]\s*(PUNJAB|PB|HARYANA|HRY|DELHI|DL|CHANDIGARH|CHD|MOHALI)\b', '', s)
    s = re.sub(r'\((PUNJAB|PB|HARYANA|HRY|DELHI|DL|CHANDIGARH|CHD|MOHALI)\)', '', s)
    terms_to_remove = [
        r'\bPRIVATE\s+LIMITED\b', r'\bPVT\.?\s*LTD\.?\b', r'\bLIMITED\b', r'\bLTD\.?\b',
        r'\bLLP\b', r'\bINC\.?\b', r'\bCORP\.?\b', r'\bCORPORATION\b', r'\bCO\.?\b',
        r'\bSERVICES\b', r'\bSOLUTIONS\b', r'\bENTERPRISES\b', r'\bTRADERS\b'
    ]
    for term in terms_to_remove:
        s = re.sub(term, '', s)
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def clean_invoice_no(inv):
    if not isinstance(inv, str):
        return ""
    s = re.sub(r'[^A-Z0-9]', '', inv.upper())
    # Strip common ERP prefixes like FB on FBADS...
    if s.startswith('FBADS'):
        s = s[2:]
    return s

def identify_file_type(file_path):
    """Detects whether a file is GSTR-2B or Tally / ERP ledger."""
    if file_path.endswith(('.xlsx', '.xls')):
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sheet in wb.sheetnames:
                s_lower = sheet.lower()
                if s_lower == 'b2b' or (s_lower.startswith('b2b') and 'cdnr' not in s_lower):
                    return 'GSTR2B'
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for r in range(1, min(ws.max_row + 1, 30)):
                    row_str = " ".join([str(ws.cell(r, c).value or '').lower() for c in range(1, min(ws.max_column + 1, 15))])
                    if 'gstin of supplier' in row_str:
                        return 'GSTR2B'
                    if 'particular' in row_str and ('debit' in row_str or 'vch' in row_str or 'credit' in row_str):
                        return 'TALLY'
                    if 'input cgst' in row_str or 'input sgst' in row_str or 'input igst' in row_str:
                        return 'TALLY'
        except Exception:
            pass
    return 'UNKNOWN'

def auto_detect_tax_mode(tally_path):
    """Automatically detects tax head (CGST, SGST, IGST) from Tally ledger."""
    try:
        wb = openpyxl.load_workbook(tally_path, data_only=True)
        all_text = " ".join(wb.sheetnames).lower()
        ws = wb.active
        for r in range(1, min(ws.max_row + 1, 20)):
            row_str = " ".join([str(ws.cell(r, c).value or '').lower() for c in range(1, min(ws.max_column + 1, 15))])
            all_text += " " + row_str

        if 'input cgst' in all_text or 'cgst 9%' in all_text or ('cgst' in all_text and 'sgst' not in all_text and 'igst' not in all_text):
            return 'cgst'
        if 'input sgst' in all_text or 'sgst 9%' in all_text or ('sgst' in all_text and 'cgst' not in all_text and 'igst' not in all_text):
            return 'sgst'
        if 'input igst' in all_text or 'igst 18%' in all_text or ('igst' in all_text and 'cgst' not in all_text and 'sgst' not in all_text):
            return 'igst'
        if 'cgst' in all_text and 'sgst' in all_text:
            return 'cgst_sgst'
    except Exception:
        pass
    return 'igst'

class GSTReconciliationAgent:
    def __init__(self, mode='auto', tolerance=1.0):
        self.mode = mode.lower()
        self.tolerance = tolerance
        self.detected_mode = None

    def parse_gstr2b(self, file_path, effective_mode=None):
        mode = effective_mode or self.mode
        if file_path.endswith(('.xlsx', '.xls')):
            xl = pd.ExcelFile(file_path)
            b2b_sheet = None
            for sheet in xl.sheet_names:
                s_lower = sheet.lower()
                if s_lower == 'b2b' or (s_lower.startswith('b2b') and 'cdnr' not in s_lower):
                    b2b_sheet = sheet
                    break
            if not b2b_sheet:
                b2b_sheet = xl.sheet_names[0]
            raw_df = pd.read_excel(file_path, sheet_name=b2b_sheet, header=None)
            rows = raw_df.values.tolist()
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                rows = list(csv.reader(f))

        header_idx = -1
        for i, row in enumerate(rows):
            row_str = " ".join([str(c) for c in row if pd.notna(c)]).lower()
            if 'gstin of supplier' in row_str or ('gstin' in row_str and 'trade' in row_str):
                header_idx = i
                break

        if header_idx == -1:
            raise ValueError(f"Could not find GSTR-2B B2B table header in {file_path}")

        records = []
        for i, r in enumerate(rows[header_idx + 1:], start=header_idx + 1):
            if not r or len(r) < 10:
                continue
            first_cell = str(r[0]).strip() if pd.notna(r[0]) else ""
            if any(term in first_cell.lower() for term in ['amendments', 'goods and services tax', 'table', 'summary']):
                break
            if not first_cell:
                continue
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
            inv_date = str(r[4]).strip() if len(r) > 4 and pd.notna(r[4]) else ""
            igst = get_float(r[9]) if len(r) > 9 else 0.0
            cgst = get_float(r[10]) if len(r) > 10 else 0.0
            sgst = get_float(r[11]) if len(r) > 11 else 0.0

            if mode == 'cgst':
                target_tax = cgst
            elif mode == 'sgst':
                target_tax = sgst
            elif mode == 'igst':
                target_tax = igst
            elif mode == 'cgst_sgst':
                target_tax = cgst + sgst
            else:
                target_tax = igst + cgst + sgst

            if target_tax != 0:
                records.append({
                    'id': i,
                    'gstin': gstin,
                    'company': trade_name,
                    'party_norm': normalize_party_name(trade_name),
                    'invoice_no': inv_no,
                    'invoice_no_clean': clean_invoice_no(inv_no),
                    'date': inv_date,
                    'tax': round(target_tax, 2)
                })
        return records

    def parse_tally(self, file_path):
        if file_path.endswith(('.xlsx', '.xls')):
            raw_df = pd.read_excel(file_path, header=None)
            rows = raw_df.values.tolist()
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                rows = list(csv.reader(f))

        header_idx = -1
        for i, row in enumerate(rows):
            row_str = " ".join([str(c) for c in row if pd.notna(c)]).lower()
            if 'particular' in row_str and ('debit' in row_str or 'vch' in row_str or 'credit' in row_str):
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
        for i, r in enumerate(rows[header_idx + 1:], start=header_idx + 1):
            if not r or len(r) < 4:
                continue
            row_text = " ".join([str(c) for c in r if pd.notna(c)]).lower()
            if any(t in row_text for t in ['opening balance', 'closing balance', 'grand total']):
                continue

            date = str(get_val(r, 'date', '')).strip()
            p_idx = col_map.get('particulars', 1)
            party = str(r[p_idx]).strip() if p_idx < len(r) and pd.notna(r[p_idx]) else ""
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
            is_credit = (credit > 0 and debit == 0)

            records.append({
                'id': i,
                'date': date[:10] if date else '',
                'company': party,
                'party_norm': normalize_party_name(party),
                'vch_type': vch_type,
                'invoice_no': vch_no,
                'invoice_no_clean': clean_invoice_no(vch_no),
                'debit': debit,
                'credit': credit,
                'tax': debit if debit > 0 else -credit,
                'is_credit': is_credit
            })
        return records

    def reconcile(self, file_a, file_b):
        # Auto-detect which file is GSTR-2B and which is Tally
        type_a = identify_file_type(file_a)
        type_b = identify_file_type(file_b)

        if type_a == 'GSTR2B' and type_b == 'TALLY':
            gstr2b_file, tally_file = file_a, file_b
        elif type_a == 'TALLY' and type_b == 'GSTR2B':
            gstr2b_file, tally_file = file_b, file_a
        elif type_a == 'GSTR2B':
            gstr2b_file, tally_file = file_a, file_b
        elif type_b == 'GSTR2B':
            gstr2b_file, tally_file = file_b, file_a
        else:
            gstr2b_file, tally_file = file_a, file_b

        # Resolve effective tax mode
        if self.mode == 'auto':
            effective_mode = auto_detect_tax_mode(tally_file)
        else:
            effective_mode = self.mode
        self.detected_mode = effective_mode

        df_2b = self.parse_gstr2b(gstr2b_file, effective_mode=effective_mode)
        df_tally = self.parse_tally(tally_file)

        tally_matched = set()
        matched_2b = set()
        all_rows = []

        # Pass 1: Match by Invoice Number (Exact or Suffix) + Tax Amount
        for r2 in df_2b:
            inv_clean = r2['invoice_no_clean']
            if not inv_clean:
                continue
            for j, rt in enumerate(df_tally):
                if j in tally_matched or rt['tax'] <= 0:
                    continue
                rt_clean = rt['invoice_no_clean']
                is_inv_match = (inv_clean == rt_clean) or (
                    len(inv_clean) >= 6 and len(rt_clean) >= 6 and (
                        inv_clean.endswith(rt_clean) or rt_clean.endswith(inv_clean)
                    )
                )
                if is_inv_match and abs(r2['tax'] - rt['debit']) <= self.tolerance:
                    tally_matched.add(j)
                    matched_2b.add(r2['id'])
                    all_rows.append({
                        'company': r2['company'],
                        'invoice_no': r2['invoice_no'],
                        'date': r2['date'] or rt['date'],
                        'tax_2b': r2['tax'],
                        'tax_books': rt['debit'],
                        'difference': 0.0,
                        'reason': 'ok',
                        'is_matched': True
                    })
                    break

        # Pass 2: Match by Party + Tax Amount ONLY IF amount is unique for that party
        unmatched_2b = [r2 for r2 in df_2b if r2['id'] not in matched_2b]
        unmatched_tally = [rt for j, rt in enumerate(df_tally) if j not in tally_matched and rt['tax'] > 0]

        g_counts = Counter((g['party_norm'], g['tax']) for g in unmatched_2b)
        b_counts = Counter((b['party_norm'], b['tax']) for b in unmatched_tally)

        for r2 in unmatched_2b:
            key = (r2['party_norm'], r2['tax'])
            # Only allow party-amount matching if unique on both sides
            if g_counts[key] == 1 and b_counts[key] == 1:
                for j, rt in enumerate(df_tally):
                    if j in tally_matched or rt['tax'] <= 0:
                        continue
                    if (rt['party_norm'], rt['debit']) == key:
                        tally_matched.add(j)
                        matched_2b.add(r2['id'])
                        all_rows.append({
                            'company': r2['company'],
                            'invoice_no': r2['invoice_no'] or rt['invoice_no'],
                            'date': r2['date'] or rt['date'],
                            'tax_2b': r2['tax'],
                            'tax_books': rt['debit'],
                            'difference': 0.0,
                            'reason': 'ok',
                            'is_matched': True
                        })
                        break

        # Pass 3: Unmatched in GSTR-2B (Not Recorded in Books)
        for r2 in df_2b:
            if r2['id'] not in matched_2b:
                all_rows.append({
                    'company': r2['company'],
                    'invoice_no': r2['invoice_no'],
                    'date': r2['date'],
                    'tax_2b': r2['tax'],
                    'tax_books': 0.0,
                    'difference': round(r2['tax'], 2),
                    'reason': 'Not Recorded in Books',
                    'is_matched': False
                })

        # Pass 4: Unmatched in Books (Not Reflected in GSTR-2B or Credit Entry)
        for j, rt in enumerate(df_tally):
            if j not in tally_matched:
                if rt['is_credit']:
                    all_rows.append({
                        'company': rt['company'],
                        'invoice_no': rt['invoice_no'],
                        'date': rt['date'],
                        'tax_2b': 0.0,
                        'tax_books': -rt['credit'],
                        'difference': round(rt['credit'], 2),
                        'reason': 'Credit / Adjustment in Books',
                        'is_matched': False
                    })
                else:
                    all_rows.append({
                        'company': rt['company'],
                        'invoice_no': rt['invoice_no'],
                        'date': rt['date'],
                        'tax_2b': 0.0,
                        'tax_books': rt['debit'],
                        'difference': round(-rt['debit'], 2),
                        'reason': 'Not Reflected in GSTR-2B',
                        'is_matched': False
                    })

        all_rows.sort(key=lambda x: (x['is_matched'], -abs(x['difference'])))

        df_all = pd.DataFrame(all_rows)
        df_disc = df_all[~df_all['is_matched']].copy()
        df_matched = df_all[df_all['is_matched']].copy()
        net_diff = round(df_all['difference'].sum(), 2)

        return {
            'all_rows': df_all,
            'discrepancies': df_disc,
            'matched': df_matched,
            'net_diff': net_diff,
            'count_discrepancies': len(df_disc),
            'count_matched': len(df_matched),
            'effective_mode': effective_mode
        }

    def export_excel_report(self, recon_result, output_path='GST_Reconciliation_Discrepancies.xlsx'):
        df = recon_result['discrepancies']
        export_df = df[['company', 'invoice_no', 'date', 'tax_2b', 'tax_books', 'difference', 'reason']].rename(columns={
            'company': 'Company Name',
            'invoice_no': 'Invoice / Voucher No',
            'date': 'Date',
            'tax_2b': 'Tax in GSTR-2B (₹)',
            'tax_books': 'Tax in Books (₹)',
            'difference': 'Difference (₹)',
            'reason': 'Reason / Action Required'
        })
        export_df.to_excel(output_path, index=False)
        print(f"Report saved to {output_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Reconcile GSTR-2B against Tally Books")
    parser.add_argument('gstr2b', help="Path to GSTR-2B Excel or CSV")
    parser.add_argument('tally', help="Path to Tally Ledger Excel or CSV")
    parser.add_argument('--mode', default='auto', choices=['auto', 'cgst', 'sgst', 'igst', 'cgst_sgst', 'total'], help="Tax mode")
    parser.add_argument('--out', default='GST_Reconciliation_Discrepancies.xlsx', help="Output report path")
    args = parser.parse_args()

    agent = GSTReconciliationAgent(mode=args.mode)
    res = agent.reconcile(args.gstr2b, args.tally)
    print(f"\nReconciliation Mode: {res['effective_mode'].upper()}")
    print(f"Matched Records: {res['count_matched']}")
    print(f"Discrepancies: {res['count_discrepancies']}")
    print(f"Net Tax Difference: ₹{res['net_diff']:,.2f}")
    print("\n--- Discrepancies Summary ---")
    print(res['discrepancies'][['company', 'invoice_no', 'tax_2b', 'tax_books', 'difference', 'reason']].to_string(index=False))
    agent.export_excel_report(res, args.out)
