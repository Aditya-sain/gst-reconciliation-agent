import pandas as pd
from test_reconcile import parse_gstr2b, parse_tally

df_2b = parse_gstr2b('sample_data/gstr2b_oct_2026.csv')
df_tally = parse_tally('sample_data/tally_oct_2026.csv')

# Filter 2B for IGST > 0
df_2b_igst = df_2b[df_2b['igst'] > 0].copy()

print("=== GSTR-2B IGST TRANSACTIONS ===")
for idx, r in df_2b_igst.iterrows():
    print(f"2B: {r['party_name']} | Inv: {r['invoice_no']} | Date: {r['invoice_date']} | IGST: {r['igst']}")

print("\n=== TALLY LEDGER TRANSACTIONS (Debit / Credit) ===")
for idx, r in df_tally.iterrows():
    print(f"Tally: {r['party_name']} | Vch: {r['vch_no']} | Date: {r['date']} | Debit: {r['debit']} | Credit: {r['credit']} | Tax: {r['tax_amount']}")
