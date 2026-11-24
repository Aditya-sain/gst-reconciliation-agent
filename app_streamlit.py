"""
Streamlit Web App for GST Reconciliation
Can be deployed to Streamlit Community Cloud (https://share.streamlit.io) in 1-click.
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from gst_recon_agent import GSTReconciliationAgent

st.set_page_config(
    page_title="GST Reconciliation AI Agent",
    page_icon="💼",
    layout="wide"
)

st.title("💼 GST Reconciliation AI Agent")
st.caption("Automated Reconciliation: GSTR-2B (GST Portal) ⟷ Tally / ERP Books")

# Sidebar
with st.sidebar:
    st.header("⚙️ Reconciliation Settings")
    tax_mode = st.radio(
        "Tax Mode",
        options=["igst", "cgst_sgst", "total"],
        format_func=lambda x: "IGST (Inter-State)" if x == "igst" else ("CGST + SGST (Intra-State)" if x == "cgst_sgst" else "Total Tax"),
        index=0
    )
    tolerance = st.number_input("Rounding Tolerance (₹)", min_value=0.0, max_value=10.0, value=1.0, step=0.5)
    
    st.markdown("---")
    st.info("💡 **GSTR-2B:** Upload B2B Sheet\n\n💡 **Tally:** Upload Input Tax Ledger")

col1, col2 = st.columns(2)
with col1:
    gstr2b_file = st.file_uploader("1. Upload GSTR-2B File (Excel or CSV)", type=["xlsx", "xls", "csv"])
with col2:
    tally_file = st.file_uploader("2. Upload Tally Ledger File (Excel or CSV)", type=["xlsx", "xls", "csv"])

# Quick demo button
use_demo = st.button("✨ Load October 2026 Sample Data")

if use_demo or (gstr2b_file and tally_file):
    with st.spinner("Analyzing vouchers and matching parties..."):
        agent = GSTReconciliationAgent(mode=tax_mode, tolerance=tolerance)
        
        if use_demo:
            file_2b_path = "sample_data/gstr2b_oct_2026.csv"
            file_tally_path = "sample_data/tally_oct_2026.csv"
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(gstr2b_file.name)[1]) as f1:
                f1.write(gstr2b_file.read())
                file_2b_path = f1.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(tally_file.name)[1]) as f2:
                f2.write(tally_file.read())
                file_tally_path = f2.name
                
        results = agent.reconcile(file_2b_path, file_tally_path)
        s = results['summary']
        df_party = results['party_summary']
        df_inv = results['invoice_details']
        
        # Display KPI Cards
        st.markdown("### 📊 Executive Summary")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("GSTR-2B Tax", f"₹ {s['total_gstr2b_tax']:,.2f}")
        k2.metric("Tally Tax Claimed", f"₹ {s['total_tally_debit_tax']:,.2f}")
        k3.metric("Net Difference", f"₹ {s['net_difference']:,.2f}")
        k4.metric("🟢 Matched", s['count_matched'])
        k5.metric("🟡 Mismatch", s['count_mismatch'])
        k6.metric("🔴 Missing in 2B", s['count_missing_in_gstr2b'])
        
        # Party Table
        st.markdown("### 📋 Party-Wise Reconciliation")
        
        filter_status = st.selectbox(
            "Filter by Status",
            options=["ALL", "TAX_MISMATCH", "MISSING_IN_GSTR2B", "MISSING_IN_TALLY", "CREDIT_REVERSAL", "MATCHED"]
        )
        
        filtered_df = df_party if filter_status == "ALL" else df_party[df_party['status'] == filter_status]
        
        display_cols = [
            'status_label', 'party_name_tally', 'party_name_2b', 'gstin', 
            'tax_gstr2b', 'tax_tally', 'difference', 'reason'
        ]
        
        st.dataframe(
            filtered_df[display_cols].rename(columns={
                'status_label': 'Status',
                'party_name_tally': 'Party Name (Tally)',
                'party_name_2b': 'Party Name (2B)',
                'gstin': 'GSTIN',
                'tax_gstr2b': 'Tax 2B (₹)',
                'tax_tally': 'Tax Tally (₹)',
                'difference': 'Diff (₹)',
                'reason': 'Explanation / Reason'
            }),
            use_container_width=True
        )
        
        # Excel Download
        output_excel = "GST_Reconciliation_Report.xlsx"
        agent.export_excel_report(results, output_excel)
        with open(output_excel, "rb") as f:
            st.download_button(
                label="📥 Download Excel Report (.xlsx)",
                data=f.read(),
                file_name="GST_Reconciliation_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Invoice drilldown
        with st.expander("🔍 View Detailed Invoice-Level Matching"):
            st.dataframe(df_inv, use_container_width=True)
