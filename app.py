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

# HTML_PAGE is loaded from index.html directly for perfect consistency
def get_html_page():
    if os.path.exists('index.html'):
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    return "<h1>GST Reconciliation Agent</h1>"

class ReconHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path == '/' or parsed.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(get_html_page().encode('utf-8'))
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
