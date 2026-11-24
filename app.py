"""
GST Reconciliation Web App & Server
Provides a modern browser interface for the Finance Department.
Runs natively using Python's built-in HTTP server (zero external web dependencies).
"""

import os
import sys
import json
import base64
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn
import traceback

from gst_recon_agent import GSTReconciliationAgent

PORT = 5000

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>GST Reconciliation Agent | Finance Department</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; }
    .badge-matched { background-color: #DCFCE7; color: #166534; border: 1px solid #86EFAC; }
    .badge-mismatch { background-color: #FEF3C7; color: #92400E; border: 1px solid #FCD34D; }
    .badge-missing-2b { background-color: #FFE4E6; color: #9F1239; border: 1px solid #FDA4AF; }
    .badge-missing-tally { background-color: #DBEAFE; color: #1E40AF; border: 1px solid #93C5FD; }
    .badge-reversal { background-color: #F3E8FF; color: #6B21A8; border: 1px solid #D8B4FE; }
  </style>
</head>
<body class="bg-slate-50 min-h-screen text-slate-800">

  <!-- TOP HEADER -->
  <header class="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 text-white shadow-lg sticky top-0 z-50">
    <div class="max-w-7xl mx-auto px-4 py-4 flex flex-col md:flex-row justify-between items-center gap-3">
      <div class="flex items-center gap-3">
        <div class="w-10 h-10 rounded-xl bg-blue-500/20 border border-blue-400 flex items-center justify-center text-blue-300 text-xl font-bold">
          <i class="fa-solid fa-scale-balanced"></i>
        </div>
        <div>
          <h1 class="text-xl font-bold tracking-tight">GST Reconciliation AI Agent</h1>
          <p class="text-xs text-blue-200">GSTR-2B (Portal) ⟷ Tally Ledger Automated Matcher</p>
        </div>
      </div>
      <div class="flex items-center gap-2">
        <button onclick="loadSampleDemo()" class="bg-amber-500 hover:bg-amber-600 text-slate-950 font-semibold px-4 py-2 rounded-lg text-sm shadow transition flex items-center gap-2">
          <i class="fa-solid fa-wand-magic-sparkles"></i> Load October 2026 Sample Data
        </button>
        <button id="downloadBtn" onclick="downloadReport()" class="hidden bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-4 py-2 rounded-lg text-sm shadow transition flex items-center gap-2">
          <i class="fa-solid fa-file-excel"></i> Download Excel Report
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 py-6 space-y-6">

    <!-- UPLOAD & CONFIG CARD -->
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
      <h2 class="text-lg font-bold text-slate-900 mb-4 flex items-center gap-2">
        <i class="fa-solid fa-cloud-arrow-up text-indigo-600"></i> Step 1: Upload Sheets & Select Tax Mode
      </h2>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        <!-- File 1: GSTR-2B -->
        <div class="border-2 border-dashed border-slate-300 hover:border-indigo-500 rounded-xl p-5 text-center transition bg-slate-50/50">
          <div class="w-12 h-12 mx-auto mb-2 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center text-xl">
            <i class="fa-solid fa-file-invoice"></i>
          </div>
          <h3 class="font-bold text-slate-800">1. GSTR-2B File (GST Portal)</h3>
          <p class="text-xs text-slate-500 mb-3">Upload B2B Excel (.xlsx) or CSV downloaded from GST Portal</p>
          <input type="file" id="gstr2bFile" accept=".csv, .xlsx, .xls" class="block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer">
        </div>

        <!-- File 2: Tally Ledger -->
        <div class="border-2 border-dashed border-slate-300 hover:border-indigo-500 rounded-xl p-5 text-center transition bg-slate-50/50">
          <div class="w-12 h-12 mx-auto mb-2 rounded-full bg-amber-100 text-amber-600 flex items-center justify-center text-xl">
            <i class="fa-solid fa-book-bookmark"></i>
          </div>
          <h3 class="font-bold text-slate-800">2. Tally Ledger / Books File</h3>
          <p class="text-xs text-slate-500 mb-3">Upload Tally Input Tax Ledger (e.g. Input IGST 18% - PB)</p>
          <input type="file" id="tallyFile" accept=".csv, .xlsx, .xls" class="block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-amber-50 file:text-amber-700 hover:file:bg-amber-100 cursor-pointer">
        </div>
      </div>

      <!-- Controls Row -->
      <div class="mt-6 pt-5 border-t border-slate-100 flex flex-col md:flex-row items-center justify-between gap-4">
        <div class="flex items-center gap-4">
          <span class="text-sm font-semibold text-slate-700">Tax Mode:</span>
          <label class="inline-flex items-center text-sm gap-2 cursor-pointer">
            <input type="radio" name="taxMode" value="igst" checked class="text-indigo-600 focus:ring-indigo-500">
            <span class="font-medium">IGST (Inter-State)</span>
          </label>
          <label class="inline-flex items-center text-sm gap-2 cursor-pointer">
            <input type="radio" name="taxMode" value="cgst_sgst" class="text-indigo-600 focus:ring-indigo-500">
            <span class="font-medium">CGST + SGST (Intra-State)</span>
          </label>
          <label class="inline-flex items-center text-sm gap-2 cursor-pointer">
            <input type="radio" name="taxMode" value="total" class="text-indigo-600 focus:ring-indigo-500">
            <span class="font-medium">Total Tax</span>
          </label>
        </div>

        <button onclick="startReconciliation()" id="reconcileBtn" class="w-full md:w-auto bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-6 py-2.5 rounded-xl shadow-md hover:shadow-lg transition flex items-center justify-center gap-2">
          <i class="fa-solid fa-play"></i> Reconcile Now (Match Karein)
        </button>
      </div>
    </div>

    <!-- LOADING SPINNER -->
    <div id="loadingBox" class="hidden text-center py-12">
      <div class="inline-block animate-spin rounded-full h-10 w-10 border-4 border-indigo-600 border-t-transparent"></div>
      <p class="mt-3 text-sm font-semibold text-slate-600">Analyzing GSTR-2B & Tally vouchers with AI...</p>
    </div>

    <!-- RESULTS SECTION -->
    <div id="resultsSection" class="hidden space-y-6">

      <!-- KPI METRIC CARDS -->
      <div class="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">GSTR-2B Tax</div>
          <div id="kpi2bTax" class="text-lg font-bold text-blue-700 mt-1">₹ 0.00</div>
          <div class="text-[11px] text-slate-400 mt-0.5">Portal Inward Supplies</div>
        </div>

        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">Tally Claimed Tax</div>
          <div id="kpiTallyTax" class="text-lg font-bold text-indigo-700 mt-1">₹ 0.00</div>
          <div class="text-[11px] text-slate-400 mt-0.5">Purchase Vouchers Debit</div>
        </div>

        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">Net Difference</div>
          <div id="kpiDiff" class="text-lg font-bold text-rose-600 mt-1">₹ 0.00</div>
          <div class="text-[11px] text-slate-400 mt-0.5">GSTR-2B minus Tally</div>
        </div>

        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">Matched Parties</div>
          <div id="kpiMatched" class="text-lg font-bold text-emerald-600 mt-1">0</div>
          <div class="text-[11px] text-emerald-600 mt-0.5">🟢 100% Verified</div>
        </div>

        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">Tax Mismatches</div>
          <div id="kpiMismatch" class="text-lg font-bold text-amber-600 mt-1">0</div>
          <div class="text-[11px] text-amber-600 mt-0.5">🟡 Value Differs</div>
        </div>

        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
          <div class="text-xs font-semibold text-slate-500 uppercase">Missing In 2B</div>
          <div id="kpiMissing2b" class="text-lg font-bold text-rose-600 mt-1">0</div>
          <div class="text-[11px] text-rose-600 mt-0.5">🔴 Vendor GSTR-1 Pending</div>
        </div>
      </div>

      <!-- FILTER TABS & SEARCH -->
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5 space-y-4">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div class="flex flex-wrap gap-2" id="filterTabs">
            <button onclick="setFilter('ALL')" class="filter-btn active px-3.5 py-1.5 rounded-lg text-xs font-bold bg-slate-900 text-white transition">All Parties (<span id="countAll">0</span>)</button>
            <button onclick="setFilter('TAX_MISMATCH')" class="filter-btn px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-amber-100 transition">Tax Mismatch (<span id="countMismatch">0</span>)</button>
            <button onclick="setFilter('MISSING_IN_GSTR2B')" class="filter-btn px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-rose-100 transition">Missing in 2B (<span id="countMissing2b">0</span>)</button>
            <button onclick="setFilter('MISSING_IN_TALLY')" class="filter-btn px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-blue-100 transition">Missing in Tally (<span id="countMissingTally">0</span>)</button>
            <button onclick="setFilter('CREDIT_REVERSAL')" class="filter-btn px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-purple-100 transition">Reversals/Credit (<span id="countReversal">0</span>)</button>
            <button onclick="setFilter('MATCHED')" class="filter-btn px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-emerald-100 transition">Matched (<span id="countMatchedTab">0</span>)</button>
          </div>

          <div class="w-full md:w-72 relative">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
            <input type="text" id="searchInput" oninput="renderTable()" placeholder="Search Party or GSTIN..." class="w-full pl-9 pr-3 py-1.5 text-xs rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-indigo-500">
          </div>
        </div>

        <!-- PARTY DETAILS TABLE -->
        <div class="overflow-x-auto rounded-xl border border-slate-200">
          <table class="w-full text-left border-collapse text-xs">
            <thead class="bg-slate-100 text-slate-700 uppercase font-semibold text-[11px] border-b border-slate-200">
              <tr>
                <th class="py-3 px-3">Status</th>
                <th class="py-3 px-3">Party Name (Tally / 2B)</th>
                <th class="py-3 px-3">Supplier GSTIN</th>
                <th class="py-3 px-3 text-right">Tax in GSTR-2B</th>
                <th class="py-3 px-3 text-right">Tax in Tally</th>
                <th class="py-3 px-3 text-right">Difference</th>
                <th class="py-3 px-4">Explanation & Action Required (Reason)</th>
                <th class="py-3 px-3 text-center">Invoices</th>
              </tr>
            </thead>
            <tbody id="partyTableBody" class="divide-y divide-slate-100 font-medium">
              <!-- Dynamically Populated -->
            </tbody>
          </table>
        </div>
      </div>

      <!-- INVOICE LEVEL MODAL / DRAWER -->
      <div id="invoiceModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl max-w-4xl w-full max-h-[85vh] flex flex-col shadow-2xl overflow-hidden border border-slate-200">
          <div class="p-4 bg-slate-900 text-white flex justify-between items-center">
            <h3 class="font-bold text-sm" id="modalTitle">Invoice Drilldown</h3>
            <button onclick="closeModal()" class="text-slate-400 hover:text-white text-lg"><i class="fa-solid fa-xmark"></i></button>
          </div>
          <div class="p-4 overflow-y-auto flex-1">
            <table class="w-full text-left text-xs">
              <thead class="bg-slate-100 text-slate-600 font-semibold uppercase text-[10px]">
                <tr>
                  <th class="p-2">Status</th>
                  <th class="p-2">Invoice No (2B)</th>
                  <th class="p-2">Voucher No (Tally)</th>
                  <th class="p-2 text-right">Tax in 2B</th>
                  <th class="p-2 text-right">Tax in Tally</th>
                  <th class="p-2 text-right">Diff</th>
                  <th class="p-2">Match Logic</th>
                </tr>
              </thead>
              <tbody id="modalTableBody" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>
      </div>

    </div>

  </main>

  <script>
    let globalData = null;
    let currentFilter = 'ALL';

    const formatCurrency = (val) => {
      const num = Number(val) || 0;
      return '₹ ' + num.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    function setFilter(filter) {
      currentFilter = filter;
      document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.classList.remove('bg-slate-900', 'text-white');
        btn.classList.add('bg-slate-100', 'text-slate-700');
      });
      event.target.classList.remove('bg-slate-100', 'text-slate-700');
      event.target.classList.add('bg-slate-900', 'text-white');
      renderTable();
    }

    async function loadSampleDemo() {
      document.getElementById('loadingBox').classList.remove('hidden');
      document.getElementById('resultsSection').classList.add('hidden');

      try {
        const mode = document.querySelector('input[name="taxMode"]:checked').value;
        const res = await fetch(`/api/reconcile-sample?mode=${mode}`);
        const data = await res.json();
        if (data.error) {
          alert("Error: " + data.error);
          return;
        }
        globalData = data;
        displayResults(data);
      } catch (err) {
        alert("Failed to run sample reconciliation: " + err);
      } finally {
        document.getElementById('loadingBox').classList.add('hidden');
      }
    }

    function readFileAsBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const res = reader.result;
          const base64 = res.split(',')[1];
          resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    async function startReconciliation() {
      const gstr2bFile = document.getElementById('gstr2bFile').files[0];
      const tallyFile = document.getElementById('tallyFile').files[0];

      if (!gstr2bFile || !tallyFile) {
        alert("Kripya dono sheets (GSTR-2B aur Tally) upload karein ya upar 'Load October 2026 Sample Data' pe click karein.");
        return;
      }

      document.getElementById('loadingBox').classList.remove('hidden');
      document.getElementById('resultsSection').classList.add('hidden');

      try {
        const gstr2bB64 = await readFileAsBase64(gstr2bFile);
        const tallyB64 = await readFileAsBase64(tallyFile);
        const mode = document.querySelector('input[name="taxMode"]:checked').value;

        const payload = {
          mode: mode,
          gstr2b_data: gstr2bB64,
          gstr2b_name: gstr2bFile.name,
          tally_data: tallyB64,
          tally_name: tallyFile.name
        };

        const res = await fetch('/api/reconcile-upload', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const data = await res.json();
        if (data.error) {
          alert("Error: " + data.error);
          return;
        }
        globalData = data;
        displayResults(data);
      } catch (err) {
        alert("Failed to reconcile files: " + err);
      } finally {
        document.getElementById('loadingBox').classList.add('hidden');
      }
    }

    function displayResults(data) {
      document.getElementById('resultsSection').classList.remove('hidden');
      document.getElementById('downloadBtn').classList.remove('hidden');

      const s = data.summary;
      document.getElementById('kpi2bTax').innerText = formatCurrency(s.total_gstr2b_tax);
      document.getElementById('kpiTallyTax').innerText = formatCurrency(s.total_tally_debit_tax);
      document.getElementById('kpiDiff').innerText = formatCurrency(s.net_difference);
      document.getElementById('kpiMatched').innerText = s.count_matched;
      document.getElementById('kpiMismatch').innerText = s.count_mismatch;
      document.getElementById('kpiMissing2b').innerText = s.count_missing_in_gstr2b;

      // Update Tab Counts
      const parties = data.party_summary;
      document.getElementById('countAll').innerText = parties.length;
      document.getElementById('countMismatch').innerText = parties.filter(p => p.status === 'TAX_MISMATCH').length;
      document.getElementById('countMissing2b').innerText = parties.filter(p => p.status === 'MISSING_IN_GSTR2B').length;
      document.getElementById('countMissingTally').innerText = parties.filter(p => p.status === 'MISSING_IN_TALLY').length;
      document.getElementById('countReversal').innerText = parties.filter(p => p.status === 'CREDIT_REVERSAL').length;
      document.getElementById('countMatchedTab').innerText = parties.filter(p => p.status === 'MATCHED').length;

      renderTable();
    }

    function renderTable() {
      if (!globalData) return;
      const tbody = document.getElementById('partyTableBody');
      tbody.innerHTML = '';

      const query = document.getElementById('searchInput').value.toLowerCase().trim();

      const filtered = globalData.party_summary.filter(r => {
        const matchesFilter = (currentFilter === 'ALL') || (r.status === currentFilter);
        const searchStr = (r.party_name_tally + " " + r.party_name_2b + " " + (r.gstin || "")).toLowerCase();
        const matchesSearch = !query || searchStr.includes(query);
        return matchesFilter && matchesSearch;
      });

      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" class="text-center py-8 text-slate-400">No parties match the selected filter.</td></tr>`;
        return;
      }

      filtered.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80 transition";

        let badgeClass = 'badge-matched';
        if (r.status === 'TAX_MISMATCH') badgeClass = 'badge-mismatch';
        else if (r.status === 'MISSING_IN_GSTR2B') badgeClass = 'badge-missing-2b';
        else if (r.status === 'MISSING_IN_TALLY') badgeClass = 'badge-missing-tally';
        else if (r.status === 'CREDIT_REVERSAL') badgeClass = 'badge-reversal';

        const displayName = r.party_name_tally !== '--- Not Found in Tally ---' ? r.party_name_tally : r.party_name_2b;
        const subName = (r.party_name_2b && r.party_name_2b !== displayName) ? `<div class="text-[10px] text-slate-400">2B: ${r.party_name_2b}</div>` : '';

        tr.innerHTML = `
          <td class="py-3 px-3 whitespace-nowrap">
            <span class="inline-block px-2.5 py-1 rounded-full text-[10px] font-bold ${badgeClass}">
              ${r.status_label}
            </span>
          </td>
          <td class="py-3 px-3">
            <div class="font-bold text-slate-800">${displayName}</div>
            ${subName}
          </td>
          <td class="py-3 px-3 font-mono text-[11px] text-slate-600">${r.gstin || '---'}</td>
          <td class="py-3 px-3 text-right font-semibold text-blue-700">${formatCurrency(r.tax_gstr2b)}</td>
          <td class="py-3 px-3 text-right font-semibold text-indigo-700">${formatCurrency(r.tax_tally)}</td>
          <td class="py-3 px-3 text-right font-bold ${Math.abs(r.difference) > 1 ? 'text-rose-600' : 'text-slate-500'}">
            ${formatCurrency(r.difference)}
          </td>
          <td class="py-3 px-4 text-slate-600 max-w-md">
            ${r.reason}
          </td>
          <td class="py-3 px-3 text-center whitespace-nowrap">
            <button onclick="viewInvoices('${escape(displayName)}')" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-medium">
              <i class="fa-solid fa-list-check"></i> View
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function viewInvoices(escapedName) {
      const name = unescape(escapedName);
      const invoices = globalData.invoice_details.filter(inv => 
        (inv.party_tally && inv.party_tally.includes(name)) || 
        (inv.party_2b && inv.party_2b.includes(name))
      );

      document.getElementById('modalTitle').innerText = `Invoice Drilldown: ${name} (${invoices.length} entries)`;
      const modalBody = document.getElementById('modalTableBody');
      modalBody.innerHTML = '';

      if (invoices.length === 0) {
        modalBody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">No invoice level lines found.</td></tr>`;
      } else {
        invoices.forEach(inv => {
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td class="p-2 font-bold ${inv.status === 'MATCHED' ? 'text-emerald-600' : 'text-rose-600'}">${inv.status}</td>
            <td class="p-2 font-mono">${inv.invoice_2b || '---'}</td>
            <td class="p-2 font-mono">${inv.vch_no_tally || '---'}</td>
            <td class="p-2 text-right font-semibold text-blue-700">${formatCurrency(inv.tax_2b)}</td>
            <td class="p-2 text-right font-semibold text-indigo-700">${formatCurrency(inv.tax_tally)}</td>
            <td class="p-2 text-right font-bold text-slate-700">${formatCurrency(inv.difference)}</td>
            <td class="p-2 text-slate-500 text-[11px]">${inv.match_type}</td>
          `;
          modalBody.appendChild(tr);
        });
      }

      document.getElementById('invoiceModal').classList.remove('hidden');
    }

    function closeModal() {
      document.getElementById('invoiceModal').classList.add('hidden');
    }

    function downloadReport() {
      window.location.href = '/api/download-report';
    }
  </script>
</body>
</html>
"""

class ReconHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path == '/' or parsed.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode('utf-8'))
            return

        elif parsed.path == '/api/reconcile-sample':
            query = urllib.parse.parse_qs(parsed.query)
            mode = query.get('mode', ['igst'])[0]
            try:
                agent = GSTReconciliationAgent(mode=mode)
                res = agent.reconcile('sample_data/gstr2b_oct_2026.csv', 'sample_data/tally_oct_2026.csv')
                agent.export_excel_report(res, 'GST_Reconciliation_Report.xlsx')
                
                output = {
                    'summary': res['summary'],
                    'party_summary': res['party_summary'].to_dict(orient='records'),
                    'invoice_details': res['invoice_details'].to_dict(orient='records')
                }
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(output).encode('utf-8'))
            except Exception as e:
                traceback.print_exc()
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))
            return

        elif parsed.path == '/api/download-report':
            report_file = 'GST_Reconciliation_Report.xlsx'
            if os.path.exists(report_file):
                with open(report_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="GST_Reconciliation_Report.xlsx"')
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_response(404)
                self.end_headers()
            return
            
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/api/reconcile-upload':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(body)
                
                mode = data.get('mode', 'igst')
                gstr2b_b64 = data.get('gstr2b_data', '')
                tally_b64 = data.get('tally_data', '')
                gstr2b_name = data.get('gstr2b_name', '2b.tmp')
                tally_name = data.get('tally_name', 'tally.tmp')
                
                gstr2b_ext = os.path.splitext(gstr2b_name)[1] or '.tmp'
                tally_ext = os.path.splitext(tally_name)[1] or '.tmp'
                
                os.makedirs('temp_uploads', exist_ok=True)
                gstr2b_path = os.path.join('temp_uploads', f'uploaded_gstr2b{gstr2b_ext}')
                tally_path = os.path.join('temp_uploads', f'uploaded_tally{tally_ext}')
                
                with open(gstr2b_path, 'wb') as f:
                    f.write(base64.b64decode(gstr2b_b64))
                with open(tally_path, 'wb') as f:
                    f.write(base64.b64decode(tally_b64))
                    
                agent = GSTReconciliationAgent(mode=mode)
                res = agent.reconcile(gstr2b_path, tally_path)
                agent.export_excel_report(res, 'GST_Reconciliation_Report.xlsx')
                
                output = {
                    'summary': res['summary'],
                    'party_summary': res['party_summary'].to_dict(orient='records'),
                    'invoice_details': res['invoice_details'].to_dict(orient='records')
                }
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(output).encode('utf-8'))
                return
            except Exception as e:
                traceback.print_exc()
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))
                return

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

def run_server(port=PORT):
    server = ThreadedHTTPServer(('0.0.0.0', port), ReconHandler)
    print(f"============================================================")
    print(f"  GST Reconciliation Agent Web Server running at:")
    print(f"  http://localhost:{port}")
    print(f"============================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        server.server_close()

if __name__ == '__main__':
    run_server(PORT)
