"""
GST Reconciliation Agent for Finance Departments
Exact invoice/voucher level reconciliation between GSTR-2B and Books / Tally Ledgers.
"""

import os
import re
import csv
import argparse
import pandas as pd
import numpy as np
from datetime import datetime
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
    return re.sub(r'[^A-Z0-9]', '', inv.upper())

class GSTReconciliationAgent:
    def __init__(self, mode='igst', tolerance=1.0):
        self.mode = mode.lower()
        self.tolerance = tolerance

    def parse_gstr2b(self, file_path):
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

        header_idx = -1
        for i, row in enumerate(rows):
            row_str = " ".join([str(c) for c in row if pd.notna(c)]).lower()
            if 'gstin of supplier' in row_str:
                header_idx = i
                break

        if header_idx == -1:
            raise ValueError(f"Could not find GSTR-2B B2B table header in {file_path}")

        records = []
        for i, r in enumerate(rows[header_idx + 2:], start=header_idx + 2):
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

            if self.mode == 'igst':
                target_tax = igst
            elif self.mode == 'cgst_sgst':
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
                    'tax': target_tax
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
        for i, r in enumerate(rows[header_idx + 1:], start=header_idx + 1):
            if not r or len(r) < 4:
                continue
            row_text = " ".join([str(c) for c in r if pd.notna(c)]).lower()
            if 'opening balance' in row_text or 'closing balance' in row_text:
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
                'date': date,
                'company': party,
                'party_norm': normalize_party_name(party),
                'vch_type': vch_type,
                'invoice_no': vch_no,
                'invoice_no_clean': clean_invoice_no(vch_no),
                'debit': debit,
                'credit': credit,
                'is_credit': is_credit
            })
        return records

    def reconcile(self, gstr2b_file, tally_file):
        df_2b = self.parse_gstr2b(gstr2b_file)
        df_tally = self.parse_tally(tally_file)

        tally_matched = set()
        matched_2b = set()
        all_rows = []

        # Pass 1: Match by Invoice Number + Amount
        for r2 in df_2b:
            inv_clean = r2['invoice_no_clean']
            if not inv_clean:
                continue
            for j, rt in enumerate(df_tally):
                if j in tally_matched:
                    continue
                if rt['invoice_no_clean'] and inv_clean == rt['invoice_no_clean'] and abs(r2['tax'] - rt['debit']) <= self.tolerance:
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

        # Pass 2: Match by Party + Amount
        for r2 in df_2b:
            if r2['id'] in matched_2b:
                continue
            for j, rt in enumerate(df_tally):
                if j in tally_matched:
                    continue
                if r2['party_norm'] == rt['party_norm'] and abs(r2['tax'] - rt['debit']) <= self.tolerance:
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
            'count_matched': len(df_matched)
        }

    def export_excel_report(self, recon_result, output_path='GST_Reconciliation_Discrepancies.xlsx'):
        df = recon_result['all_rows']
        export_df = df[['company', 'invoice_no', 'date', 'tax_2b', 'tax_books', 'difference', 'reason']].rename(columns={
            'company': 'Company Name',
            'invoice_no': 'Invoice / Voucher No',
            'date': 'Date',
            'tax_2b': 'Tax in GSTR-2B (₹)',
            'tax_books': 'Tax in Books (₹)',
            'difference': 'Difference (₹)',
            'reason': 'Reason / Remarks'
        })
        export_df.to_excel(output_path, index=False)
        print(f"Report saved to {output_path}")
        return output_path

if __name__ == '__main__':
    agent = GSTReconciliationAgent()
    res = agent.reconcile('sample_data/gstr2b_oct_2026.csv', 'sample_data/tally_oct_2026.csv')
    print("\n" + "="*80)
    print("                 GST DISCREPANCY RECONCILIATION REPORT")
    print("="*80)
    print(res['discrepancies'][['company', 'invoice_no', 'tax_2b', 'tax_books', 'difference', 'reason']].to_string())
    print("\nTotal Net Difference:", res['net_diff'])
