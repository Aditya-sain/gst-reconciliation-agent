import pandas as pd
import numpy as np
import re
import csv

def parse_gstr2b(file_path):
    """
    Parses GSTR-2B file (CSV or Excel).
    Detects the 'Taxable inward supplies received from registered persons' table or 'B2B' sheet.
    """
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        reader = list(csv.reader(f))
    
    # Locate header row containing 'GSTIN of supplier'
    header_idx = -1
    for i, row in enumerate(reader):
        row_str = " ".join(row).lower()
        if 'gstin of supplier' in row_str:
            header_idx = i
            break
            
    if header_idx == -1:
        raise ValueError("Could not find GSTR-2B B2B table header.")
        
    # Standard GSTR-2B has 2 header rows
    row1 = reader[header_idx]
    row2 = reader[header_idx + 1] if header_idx + 1 < len(reader) else []
    
    # Extract records
    records = []
    for r in reader[header_idx + 2:]:
        if not r or len(r) < 10:
            continue
        # Stop if we hit another section (like Amendments, summary, etc.)
        first_cell = r[0].strip()
        if 'amendments' in first_cell.lower() or 'goods and services tax' in first_cell.lower() or 'table' in first_cell.lower():
            break
        if not first_cell: # empty line or summary
            continue
        
        # Valid GSTIN is 15 chars alphanumeric (e.g. 03AADCB0848H1Z2)
        if len(first_cell) != 15 and not any(c.isdigit() for c in first_cell[:2]):
            continue
            
        gstin = r[0].strip()
        trade_name = r[1].strip()
        inv_no = r[2].strip()
        inv_type = r[3].strip() if len(r) > 3 else ""
        inv_date = r[4].strip() if len(r) > 4 else ""
        inv_val = float(r[5].replace(',', '')) if len(r) > 5 and r[5].strip() else 0.0
        pos = r[6].strip() if len(r) > 6 else ""
        rcm = r[7].strip() if len(r) > 7 else "No"
        taxable_val = float(r[8].replace(',', '')) if len(r) > 8 and r[8].strip() else 0.0
        igst = float(r[9].replace(',', '')) if len(r) > 9 and r[9].strip() else 0.0
        cgst = float(r[10].replace(',', '')) if len(r) > 10 and r[10].strip() else 0.0
        sgst = float(r[11].replace(',', '')) if len(r) > 11 and r[11].strip() else 0.0
        cess = float(r[12].replace(',', '')) if len(r) > 12 and r[12].strip() else 0.0
        
        records.append({
            'source': 'GSTR-2B',
            'gstin': gstin,
            'party_name': trade_name,
            'invoice_no': inv_no,
            'invoice_date': inv_date,
            'taxable_value': taxable_val,
            'igst': igst,
            'cgst': cgst,
            'sgst': sgst,
            'total_tax': igst + cgst + sgst,
            'inv_val': inv_val,
            'rcm': rcm
        })
    return pd.DataFrame(records)

def parse_tally(file_path):
    """
    Parses Tally Ledger file (CSV or Excel).
    Looks for Date, Particulars, Vch Type, Vch No, Debit, Credit.
    """
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        reader = list(csv.reader(f))
        
    header_idx = -1
    for i, row in enumerate(reader):
        row_str = " ".join(row).lower()
        if 'particulars' in row_str and ('debit' in row_str or 'vch' in row_str):
            header_idx = i
            break
            
    if header_idx == -1:
        raise ValueError("Could not find Tally ledger table header.")
        
    header = reader[header_idx]
    # Identify column indices
    col_map = {}
    for j, h in enumerate(header):
        hl = h.strip().lower()
        if 'date' in hl:
            col_map['date'] = j
        elif 'particular' in hl:
            col_map['particulars'] = j
        elif 'vch type' in hl:
            col_map['vch_type'] = j
        elif 'vch no' in hl:
            col_map['vch_no'] = j
        elif 'debit' in hl:
            col_map['debit'] = j
        elif 'credit' in hl:
            col_map['credit'] = j
            
    records = []
    for r in reader[header_idx + 1:]:
        if not r or len(r) < 4:
            continue
        
        # Check if row is opening balance, closing balance, or total
        row_text = " ".join(r).lower()
        if 'opening balance' in row_text or 'closing balance' in row_text:
            continue
            
        date = r[col_map.get('date', 0)].strip() if 'date' in col_map else ""
        # Sometimes Particulars is shifted by 1 column if Dr/Cr is in between
        p_idx = col_map.get('particulars', 1)
        party = r[p_idx].strip() if len(r) > p_idx else ""
        if party in ['Dr', 'Cr'] and len(r) > p_idx + 1:
            party = r[p_idx + 1].strip()
            
        if not party or party.lower() in ['total', 'grand total']:
            continue
            
        vch_type = r[col_map['vch_type']].strip() if 'vch_type' in col_map and len(r) > col_map['vch_type'] else ""
        vch_no = r[col_map['vch_no']].strip() if 'vch_no' in col_map and len(r) > col_map['vch_no'] else ""
        
        debit_str = r[col_map['debit']].strip().replace(',', '') if 'debit' in col_map and len(r) > col_map['debit'] else ""
        credit_str = r[col_map['credit']].strip().replace(',', '') if 'credit' in col_map and len(r) > col_map['credit'] else ""
        
        debit = float(debit_str) if debit_str else 0.0
        credit = float(credit_str) if credit_str else 0.0
        
        # In Input tax ledger, Debit is tax claimed on purchases. Credit is reversal or sale tax.
        tax_amount = debit if debit > 0 else -credit
        
        records.append({
            'source': 'Tally',
            'date': date,
            'party_name': party,
            'vch_type': vch_type,
            'vch_no': vch_no,
            'tax_amount': tax_amount,
            'debit': debit,
            'credit': credit
        })
    return pd.DataFrame(records)

df_2b = parse_gstr2b('sample_data/gstr2b_oct_2026.csv')
df_tally = parse_tally('sample_data/tally_oct_2026.csv')

print(f"GSTR-2B records: {len(df_2b)}")
print(df_2b[['gstin', 'party_name', 'invoice_no', 'igst', 'cgst', 'sgst']])

print("\n" + "="*50 + "\n")

print(f"Tally records: {len(df_tally)}")
print(df_tally[['date', 'party_name', 'vch_type', 'vch_no', 'tax_amount']])
