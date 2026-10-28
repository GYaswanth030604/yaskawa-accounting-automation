"""
YASKAWA Import Accounting Automation - Comprehensive Web GUI Server
Zero-dependency, offline desktop interface running on localhost.
Allows uploading all shipment documents (Commercial Invoices, Bills of Entry,
Delivery Notes, Packing Lists, Air Waybills, Masters, and Supporting Files),
viewing and inline-editing the 7-Column Product Master (master.xlsx),
running the accounting pipeline, inspecting live audit checks, and downloading workbooks.
"""

import sys
import os
import json
import shutil
import logging
import hashlib
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote
import openpyxl

# Import main pipeline
from main import (
    run_pipeline,
    INPUT_FOLDER,
    OUTPUT_FOLDER,
    BASE_EXCEL,
    MASTER_EXCEL,
    OUTPUT_FILENAME,
    ROUNDING_TOLERANCE,
    SWS_RATE,
    FILL_BATCH_FROM_DELIVERY_NOTE,
    STOP_ON_MISSING_DATA
)

PORT = int(os.environ.get("PORT", 8765))
HOST = "0.0.0.0" if os.environ.get("PORT") or os.environ.get("RENDER") else "127.0.0.1"
WORKSPACE_ROOT = Path(__file__).resolve().parent

# Global execution state
pipeline_state = {
    "running": False,
    "logs": [],
    "result": None,
    "error": None
}

class PipelineLogHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        pipeline_state["logs"].append(msg)

# Attach custom log handler to root logger
log_handler = PipelineLogHandler()
log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logging.getLogger().addHandler(log_handler)


def classify_file(filename: str) -> str:
    """Classifies a document based on its filename and extension."""
    f = filename.lower()
    if f.endswith(".xlsx") or f.endswith(".xls"):
        if "master" in f:
            return "Product Master"
        elif "base" in f or "vipa" in f or "purchase" in f:
            return "Base Excel Template"
        return "Excel Data File"
    elif f.endswith(".pdf"):
        if "invoice" in f or "rechnung" in f or "facture" in f or "380" in f:
            return "Commercial Invoice"
        elif "be" in f or "bill" in f or "4342" in f or "4687" in f or "icegate" in f:
            return "Bill of Entry"
        elif "delivery" in f or "dn" in f or "lieferschein" in f:
            return "Delivery Note"
        elif "packing" in f or "pkg" in f:
            return "Packing List"
        elif "awb" in f or "airway" in f or "bl" in f:
            return "Air Waybill / BL"
        elif "challan" in f or "duty" in f or "receipt" in f:
            return "Customs Duty Receipt"
        return "Supporting Document (PDF)"
    return "Supporting Document"


def get_files_list() -> list:
    in_dir = WORKSPACE_ROOT / INPUT_FOLDER
    in_dir.mkdir(parents=True, exist_ok=True)
    file_list = []
    for p in sorted(in_dir.iterdir()):
        if p.is_file():
            stat = p.stat()
            size_kb = stat.st_size / 1024
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
            dt_str = datetime.fromtimestamp(stat.st_mtime).strftime("%d-%b-%Y %H:%M")
            file_list.append({
                "name": p.name,
                "size": size_str,
                "bytes": stat.st_size,
                "date": dt_str,
                "category": classify_file(p.name)
            })
    return file_list


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>YASKAWA Import Accounting & Automation Hub</title>
  <style>
    :root {
      --bg: #090d16;
      --card-bg: #131c2e;
      --card-sub: #1a253c;
      --border: #26354f;
      --text: #f8fafc;
      --muted: #94a3b8;
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --success: #10b981;
      --warning: #f59e0b;
      --danger: #ef4444;
      --purple: #8b5cf6;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding: 24px;
    }
    .container { max-width: 1180px; margin: 0 auto; }
    .header {
      background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #4f46e5 100%);
      padding: 26px 30px;
      border-radius: 16px;
      margin-bottom: 24px;
      box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.35);
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 16px;
    }
    .header h1 { font-size: 24px; font-weight: 700; color: #fff; }
    .header p { font-size: 13px; color: #bfdbfe; margin-top: 4px; }
    .badge {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 9999px;
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      background: rgba(16, 185, 129, 0.2);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
      margin-bottom: 6px;
    }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; margin-bottom: 20px; }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .card-title {
      font-size: 13px;
      text-transform: uppercase;
      font-weight: 700;
      letter-spacing: 0.5px;
      color: var(--muted);
      margin-bottom: 14px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    /* Universal Upload Box */
    .universal-dropzone {
      border: 2px dashed #3b82f6;
      border-radius: 12px;
      padding: 24px 20px;
      text-align: center;
      transition: all 0.2s ease;
      cursor: pointer;
      background: rgba(37, 99, 235, 0.05);
      position: relative;
      margin-bottom: 16px;
    }
    .universal-dropzone:hover, .universal-dropzone.dragover {
      border-color: #60a5fa;
      background: rgba(37, 99, 235, 0.12);
      transform: translateY(-1px);
    }
    .universal-dropzone input[type="file"] {
      position: absolute;
      top: 0; left: 0; width: 100%; height: 100%;
      opacity: 0;
      cursor: pointer;
    }
    .dz-icon {
      width: 40px; height: 40px; margin: 0 auto 8px;
      background: rgba(37, 99, 235, 0.2);
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      color: #60a5fa;
    }
    .dz-title { font-size: 14px; font-weight: 600; color: #fff; }
    .dz-desc { font-size: 12px; color: var(--muted); margin-top: 4px; }
    
    /* Tables */
    .table-container {
      overflow-x: auto;
      border-radius: 10px;
      border: 1px solid var(--border);
      max-height: 440px;
      overflow-y: auto;
    }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    th {
      text-align: left; padding: 10px 12px;
      background: var(--card-sub);
      color: var(--muted);
      border-bottom: 1px solid var(--border);
      font-weight: 600;
      position: sticky; top: 0; z-index: 2;
    }
    td { padding: 8px 12px; border-bottom: 1px solid rgba(255, 255, 255, 0.04); font-family: monospace; }
    tr:hover td { background: rgba(255, 255, 255, 0.02); }
    
    /* Badges */
    .cat-badge {
      display: inline-block; padding: 2px 8px; border-radius: 6px;
      font-size: 11px; font-weight: 600; font-family: -apple-system, sans-serif;
    }
    .cat-invoice { background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
    .cat-boe { background: rgba(59, 130, 246, 0.15); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }
    .cat-dn { background: rgba(168, 85, 247, 0.15); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.3); }
    .cat-base { background: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
    .cat-master { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
    .cat-support { background: rgba(148, 163, 184, 0.15); color: #cbd5e1; border: 1px solid rgba(148, 163, 184, 0.3); }

    .btn {
      display: inline-flex; align-items: center; justify-content: center;
      gap: 8px; padding: 9px 16px; border-radius: 8px;
      font-size: 12px; font-weight: 600; cursor: pointer;
      border: none; transition: all 0.15s ease; text-decoration: none;
    }
    .btn-primary {
      background: var(--primary); color: white;
      box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
    .btn-primary:hover:not(:disabled) { background: var(--primary-hover); }
    .btn-primary:disabled { opacity: 0.6; cursor: not-allowed; }
    .btn-secondary {
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--border);
      color: var(--text);
    }
    .btn-secondary:hover { background: rgba(255, 255, 255, 0.1); }
    .btn-purple {
      background: rgba(139, 92, 246, 0.15);
      border: 1px solid rgba(139, 92, 246, 0.3);
      color: #c084fc;
    }
    .btn-purple:hover { background: rgba(139, 92, 246, 0.3); color: #fff; }
    .btn-danger {
      background: rgba(239, 68, 68, 0.15);
      border: 1px solid rgba(239, 68, 68, 0.3);
      color: #f87171; padding: 4px 10px; font-size: 11px;
    }
    .btn-danger:hover { background: rgba(239, 68, 68, 0.3); color: #fff; }
    .btn-success { background: #059669; color: white; }
    .btn-success:hover { background: #047857; }

    .actions-bar {
      display: flex; gap: 12px; align-items: center;
      justify-content: space-between; margin-bottom: 20px;
      background: var(--card-bg); border: 1px solid var(--border);
      border-radius: 14px; padding: 16px 20px;
      flex-wrap: wrap;
    }
    .log-box {
      background: #060911; border: 1px solid var(--border);
      border-radius: 10px; padding: 14px;
      font-family: "Consolas", "Courier New", monospace;
      font-size: 12px; color: #cbd5e1; height: 210px;
      overflow-y: auto; white-space: pre-wrap;
    }
    .status-pass { color: var(--success); font-weight: bold; background: rgba(16, 185, 129, 0.1); padding: 2px 8px; border-radius: 4px; }
    .status-warning { color: var(--warning); font-weight: bold; background: rgba(245, 158, 11, 0.1); padding: 2px 8px; border-radius: 4px; }
    .status-fail { color: var(--danger); font-weight: bold; background: rgba(239, 68, 68, 0.1); padding: 2px 8px; border-radius: 4px; }

    /* Editable Input in Table */
    .cell-input {
      width: 100%;
      background: rgba(0, 0, 0, 0.25);
      border: 1px solid var(--border);
      color: #fff;
      padding: 5px 8px;
      border-radius: 6px;
      font-size: 12px;
      font-family: monospace;
    }
    .cell-input:focus { outline: none; border-color: #3b82f6; background: rgba(0, 0, 0, 0.5); }
    .cell-select {
      background: #0f172a;
      border: 1px solid var(--border);
      color: #fff;
      padding: 5px 8px;
      border-radius: 6px;
      font-size: 12px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div>
        <div class="badge">● 100% Offline Shipment Hub</div>
        <h1>YASKAWA Import Accounting & Automation</h1>
        <p>Incorporate all shipment files: Commercial Invoices, Bill of Entry, Delivery Notes, and Editable Product Master.</p>
      </div>
      <div style="text-align: right;">
        <span style="font-size: 12px; color: #bfdbfe;">Exchange Rate Support</span>
        <div style="font-size: 20px; font-weight: bold; font-family: monospace;">Universal ICEGATE BoE</div>
        <span style="font-size: 11px; color: #6ee7b7;">Dynamic Duty Allocations • Zero Float Drift</span>
      </div>
    </div>

    <!-- Section 1: Multi-File Ingestion -->
    <div class="card">
      <div class="card-title">
        <span>1. Shipment Document Manager</span>
        <button class="btn btn-secondary" onclick="openExplorer('input')" style="padding: 5px 12px; font-size: 11px;">
          Open input/ in Explorer
        </button>
      </div>

      <!-- Universal Multi-File Dropzone -->
      <div class="universal-dropzone" id="universal-dz">
        <input type="file" id="multi-file-input" multiple onchange="uploadMultipleFiles(this)">
        <div class="dz-icon">
          <svg width="22" height="22" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
        </div>
        <div class="dz-title">Drag & Drop Any Files Here or Click to Browse</div>
        <div class="dz-desc">Commercial Invoice (*.pdf), Bill of Entry (*.pdf), Delivery Notes, Base Template & Master (*.xlsx)</div>
      </div>

      <!-- Live File Inventory Table -->
      <div class="table-container" style="max-height: 220px;">
        <table>
          <thead>
            <tr>
              <th>File Name</th>
              <th>Detected Category / Role</th>
              <th>Size</th>
              <th>Modified Date</th>
              <th style="text-align: center;">Action</th>
            </tr>
          </thead>
          <tbody id="files-tbody">
            <tr><td colspan="5" style="text-align: center; color: var(--muted); padding: 14px;">Loading files from input/ folder...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Section 2: Editable Product Master -->
    <div class="card">
      <div class="card-title">
        <span>2. Product Master Database (master.xlsx - Editable)</span>
        <div style="display: flex; gap: 8px;">
          <button class="btn btn-secondary" onclick="openMasterExcel()" style="padding: 5px 12px; font-size: 11px;">
            Open in Excel
          </button>
          <button class="btn btn-purple" onclick="mergeReviewedMaster()" style="padding: 5px 12px; font-size: 11px;">
            Import New Parts from Review
          </button>
          <button class="btn btn-secondary" onclick="addNewMasterRow()" style="padding: 5px 12px; font-size: 11px;">
            + Add Part
          </button>
          <button class="btn btn-success" onclick="saveMasterData()" id="btn-save-master" style="padding: 5px 14px; font-size: 11px;">
            Save to master.xlsx
          </button>
        </div>
      </div>

      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <input type="text" id="master-search" oninput="filterMasterTable()" placeholder="Search Part No, Description, or Tally Name..." style="width: 320px; padding: 7px 12px; background: #0a0f1d; border: 1px solid var(--border); border-radius: 8px; color: #fff; font-size: 12px;">
        <span id="master-count" style="font-size: 11px; color: var(--muted); font-mono;">Loading items...</span>
      </div>

      <div class="table-container" style="max-height: 380px;">
        <table id="master-table">
          <thead>
            <tr>
              <th style="width: 140px;">Part No</th>
              <th style="width: 260px;">Tally Item Name</th>
              <th>Description</th>
              <th style="width: 110px;">HS Code</th>
              <th style="width: 80px;">BCD %</th>
              <th style="width: 130px;">Batch Type</th>
              <th style="width: 160px;">Notes</th>
              <th style="text-align: center; width: 60px;">Del</th>
            </tr>
          </thead>
          <tbody id="master-tbody">
            <tr><td colspan="8" style="text-align: center; color: var(--muted); padding: 16px;">Loading Product Master...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Section 3: Processing Options & Execution -->
    <div class="card" style="padding: 16px 20px;">
      <div class="card-title">3. Processing Options</div>
      <div style="display: flex; gap: 24px; flex-wrap: wrap; font-size: 12px;">
        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
          <input type="checkbox" id="chk-batch" style="accent-color: #2563eb;">
          <span>Fill Batch & Serial No from Delivery Note (if present)</span>
        </label>
        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
          <input type="checkbox" id="chk-review" checked style="accent-color: #2563eb;">
          <span>Generate Highlighted Audit Copy (_REVIEW.xlsx)</span>
        </label>
        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
          <input type="checkbox" id="chk-stop" style="accent-color: #2563eb;">
          <span>Stop if Master fields are uncataloged</span>
        </label>
      </div>
    </div>

    <!-- Execution Bar -->
    <div class="actions-bar">
      <div>
        <div style="font-size: 14px; font-weight: 600;">Run Automated Extraction & Allocation</div>
        <div style="font-size: 12px; color: var(--muted);">Executes OCR, BoE spatial extraction, batch expansion, and Excel formatting</div>
      </div>
      <div style="display: flex; gap: 10px;">
        <button class="btn btn-secondary" onclick="openExplorer('output')">Open Output Folder</button>
        <button class="btn btn-primary" id="btn-run" onclick="startPipeline()">
          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
          Run Pipeline Now
        </button>
      </div>
    </div>

    <!-- Logs & Outputs -->
    <div class="grid">
      <!-- Live Terminal Output -->
      <div class="card" style="margin-bottom: 0;">
        <div class="card-title">Live Execution Logs</div>
        <div class="log-box" id="log-box">Waiting to start pipeline...</div>
      </div>

      <!-- Generated Workbooks -->
      <div class="card" style="margin-bottom: 0;">
        <div class="card-title">Generated Workbooks (16-Column Purchase Format)</div>
        <div style="display: flex; flex-direction: column; gap: 10px;">
          <a class="btn btn-success" id="link-clean" href="/download/YASKAWA_NEW_SHIPMENT.xlsx" style="justify-content: space-between;">
            <span>Clean Tally Excel (Import-Ready)</span>
            <span>Download</span>
          </a>
          <a class="btn btn-secondary" id="link-review" href="/download/YASKAWA_NEW_SHIPMENT_REVIEW.xlsx" style="justify-content: space-between;">
            <span>Highlighted Audit Copy (_REVIEW.xlsx)</span>
            <span>Download</span>
          </a>
          <a class="btn btn-secondary" id="link-master" href="/download/master_reviewed.xlsx" style="justify-content: space-between;">
            <span>Reviewed Master Copy</span>
            <span>Download</span>
          </a>
        </div>
      </div>
    </div>

    <!-- Reconciliation Checks Table -->
    <div class="card" id="reconciliation-card" style="margin-top: 20px;">
      <div class="card-title">
        <span>Reconciliation Checks (Audit Verification)</span>
        <span id="rec-summary" class="status-pass">All Checks Passed</span>
      </div>
      <table id="checks-table">
        <thead>
          <tr>
            <th>Verification Check</th>
            <th>Expected / Document</th>
            <th>Calculated / Actual</th>
            <th>Difference</th>
            <th style="text-align: center;">Status</th>
          </tr>
        </thead>
        <tbody id="checks-tbody">
          <!-- Populated dynamically -->
        </tbody>
      </table>
    </div>

  </div>

  <script>
    let isRunning = false;
    let pollInterval = null;
    let masterRows = [];

    // --- File Inventory ---
    async function loadFiles() {
      try {
        const res = await fetch('/api/files');
        const files = await res.json();
        const tbody = document.getElementById('files-tbody');
        if (!files || files.length === 0) {
          tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; color: var(--muted); padding: 14px;">No files currently in input/ folder. Drag and drop files above to add them.</td></tr>';
          return;
        }

        tbody.innerHTML = '';
        files.forEach(f => {
          const tr = document.createElement('tr');
          let catClass = 'cat-support';
          if (f.category.includes('Invoice')) catClass = 'cat-invoice';
          else if (f.category.includes('Bill of Entry')) catClass = 'cat-boe';
          else if (f.category.includes('Delivery Note')) catClass = 'cat-dn';
          else if (f.category.includes('Base')) catClass = 'cat-base';
          else if (f.category.includes('Master')) catClass = 'cat-master';

          tr.innerHTML = `
            <td style="font-weight: 600;">${f.name}</td>
            <td><span class="cat-badge ${catClass}">${f.category}</span></td>
            <td>${f.size}</td>
            <td style="color: var(--muted);">${f.date}</td>
            <td style="text-align: center;">
              ${f.name.toLowerCase() === 'base.xlsx' 
                ? '<span style="color: #60a5fa; font-size: 11px; font-weight: 600;">🔒 Protected</span>' 
                : `<button class="btn btn-danger" onclick="deleteFile('${encodeURIComponent(f.name)}')">Delete</button>`}
            </td>
          `;
          tbody.appendChild(tr);
        });
      } catch (e) {
        console.error("Failed loading files", e);
      }
    }

    async function uploadMultipleFiles(input) {
      if (!input.files || input.files.length === 0) return;
      const files = Array.from(input.files);
      const logBox = document.getElementById('log-box');
      logBox.textContent = `Uploading ${files.length} file(s)...\\n`;

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        logBox.textContent += `[${i+1}/${files.length}] Uploading ${file.name}...\\n`;
        logBox.scrollTop = logBox.scrollHeight;
        try {
          await fetch('/api/upload?filename=' + encodeURIComponent(file.name), {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: file
          });
          logBox.textContent += `  ✓ Uploaded ${file.name}\\n`;
        } catch (err) {
          logBox.textContent += `  ✗ Failed ${file.name}: ${err.message}\\n`;
        }
      }
      input.value = '';
      loadFiles();
      updateStatus();
    }

    async function deleteFile(encodedName) {
      if (!confirm(`Are you sure you want to remove ${decodeURIComponent(encodedName)} from input/ folder?`)) return;
      try {
        await fetch('/api/files?filename=' + encodedName, { method: 'DELETE' });
        loadFiles();
        updateStatus();
      } catch (err) {
        alert("Failed to delete file: " + err.message);
      }
    }

    // --- Master Excel Management (Editable) ---
    async function loadMasterData() {
      try {
        const res = await fetch('/api/master');
        masterRows = await res.json();
        renderMasterTable(masterRows);
      } catch (e) {
        console.error("Failed to load master", e);
      }
    }

    function renderMasterTable(rows) {
      const tbody = document.getElementById('master-tbody');
      tbody.innerHTML = '';
      document.getElementById('master-count').textContent = `${rows.length} Items Configured`;

      rows.forEach((it, idx) => {
        const tr = document.createElement('tr');
        tr.dataset.index = idx;
        tr.innerHTML = `
          <td><input class="cell-input" value="${it.part_no || ''}" onchange="updateMasterCell(${idx}, 'part_no', this.value)"></td>
          <td><input class="cell-input" value="${it.tally_name || ''}" onchange="updateMasterCell(${idx}, 'tally_name', this.value)"></td>
          <td><input class="cell-input" value="${it.description || ''}" onchange="updateMasterCell(${idx}, 'description', this.value)"></td>
          <td><input class="cell-input" value="${it.hs_code || ''}" onchange="updateMasterCell(${idx}, 'hs_code', this.value)"></td>
          <td><input class="cell-input" value="${it.bcd_rate || ''}" onchange="updateMasterCell(${idx}, 'bcd_rate', this.value)"></td>
          <td>
            <select class="cell-select" onchange="updateMasterCell(${idx}, 'batch_type', this.value)">
              <option value="Batch No" ${it.batch_type === 'Batch No' ? 'selected' : ''}>Batch No</option>
              <option value="Primary Batch" ${it.batch_type === 'Primary Batch' ? 'selected' : ''}>Primary Batch</option>
            </select>
          </td>
          <td><input class="cell-input" value="${it.notes || ''}" onchange="updateMasterCell(${idx}, 'notes', this.value)"></td>
          <td style="text-align: center;">
            <button class="btn btn-danger" onclick="deleteMasterRow(${idx})" style="padding: 2px 6px;">✕</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function updateMasterCell(idx, field, val) {
      if (masterRows[idx]) {
        masterRows[idx][field] = val;
      }
    }

    function addNewMasterRow() {
      masterRows.unshift({
        part_no: "NEW-PART",
        tally_name: "",
        description: "",
        hs_code: "",
        bcd_rate: "10.0",
        batch_type: "Batch No",
        notes: "Manual Entry"
      });
      renderMasterTable(masterRows);
    }

    function deleteMasterRow(idx) {
      masterRows.splice(idx, 1);
      renderMasterTable(masterRows);
    }

    function filterMasterTable() {
      const q = document.getElementById('master-search').value.toLowerCase();
      const filtered = masterRows.filter(r => 
        (r.part_no && r.part_no.toLowerCase().includes(q)) ||
        (r.tally_name && r.tally_name.toLowerCase().includes(q)) ||
        (r.description && r.description.toLowerCase().includes(q))
      );
      renderMasterTable(filtered);
    }

    async function saveMasterData() {
      const btn = document.getElementById('btn-save-master');
      btn.disabled = true;
      btn.textContent = 'Saving...';
      try {
        const res = await fetch('/api/master', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(masterRows)
        });
        const d = await res.json();
        alert(`Saved ${d.count} items successfully to input/master.xlsx!`);
      } catch (e) {
        alert("Failed to save master: " + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = 'Save to master.xlsx';
      }
    }

    async function mergeReviewedMaster() {
      if (!confirm("Import all new uncataloged parts from output/master_reviewed.xlsx into master.xlsx?")) return;
      try {
        const res = await fetch('/api/master/merge', { method: 'POST' });
        const d = await res.json();
        alert(`Merged ${d.rows} items into master.xlsx!`);
        loadMasterData();
      } catch (e) {
        alert("Merge failed: " + e.message);
      }
    }

    function openMasterExcel() {
      fetch('/api/open_file?file=master');
    }

    // --- Pipeline Status & Controls ---
    async function updateStatus() {
      try {
        const res = await fetch('/api/status');
        const data = await res.json();
        
        // Update logs
        if (data.logs && data.logs.length > 0) {
          const logBox = document.getElementById('log-box');
          logBox.textContent = data.logs.join('\\n');
          logBox.scrollTop = logBox.scrollHeight;
        }

        // Update checks table if available
        if (data.result && data.result.checks) {
          renderChecks(data.result.checks);
        }

        const btnRun = document.getElementById('btn-run');
        if (data.running) {
          isRunning = true;
          btnRun.disabled = true;
          btnRun.textContent = 'Processing Shipment... (OCR Running)';
        } else {
          isRunning = false;
          btnRun.disabled = false;
          btnRun.innerHTML = '<svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg> Run Pipeline Now';
        }
      } catch (e) {
        console.error("Status fetch failed", e);
      }
    }

    function renderChecks(checks) {
      const tbody = document.getElementById('checks-tbody');
      tbody.innerHTML = '';
      checks.forEach(chk => {
        const tr = document.createElement('tr');
        const cls = chk.status === 'PASS' ? 'status-pass' : (chk.status === 'WARNING' ? 'status-warning' : 'status-fail');
        tr.innerHTML = `
          <td style="font-family: inherit; font-weight: 500;">${chk.name}</td>
          <td>${chk.expected}</td>
          <td style="color: #34d399;">${chk.actual}</td>
          <td style="color: #94a3b8;">${chk.diff}</td>
          <td style="text-align: center;"><span class="${cls}">${chk.status}</span></td>
        `;
        tbody.appendChild(tr);
      });
    }

    async function startPipeline() {
      if (isRunning) return;
      document.getElementById('btn-run').disabled = true;
      document.getElementById('log-box').textContent = 'Starting pipeline execution...\\n';
      
      const payload = {
        fill_batch: document.getElementById('chk-batch').checked,
        stop_missing: document.getElementById('chk-stop').checked
      };

      await fetch('/api/run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (!pollInterval) {
        pollInterval = setInterval(updateStatus, 1000);
      }
    }

    function openExplorer(type) {
      fetch(`/api/open_folder?type=${type}`);
    }

    // Drag-and-drop highlight
    const dz = document.getElementById('universal-dz');
    ['dragenter', 'dragover'].forEach(name => {
      dz.addEventListener(name, (e) => { e.preventDefault(); dz.classList.add('dragover'); }, false);
    });
    ['dragleave', 'drop'].forEach(name => {
      dz.addEventListener(name, (e) => { e.preventDefault(); dz.classList.remove('dragover'); }, false);
    });

    // Initial load
    loadFiles();
    loadMasterData();
    updateStatus();
    setInterval(updateStatus, 2500);
    setInterval(loadFiles, 6000);
  </script>
</body>
</html>
"""


class AutomationHttpServer(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode("utf-8"))

        elif path in ("/health", "/healthz"):
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"OK")

        elif path == "/api/files":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            files_data = get_files_list()
            self.wfile.write(json.dumps(files_data).encode("utf-8"))

        elif path == "/api/master":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            master_file = WORKSPACE_ROOT / INPUT_FOLDER / MASTER_EXCEL
            rows = []
            if master_file.exists():
                wb = openpyxl.load_workbook(master_file, data_only=True)
                ws = wb.active
                for r in range(2, ws.max_row + 1):
                    p_no = str(ws.cell(r, 1).value or "").strip()
                    if not p_no:
                        continue
                    rows.append({
                        "part_no": p_no,
                        "tally_name": str(ws.cell(r, 2).value or "").strip(),
                        "description": str(ws.cell(r, 3).value or "").strip(),
                        "hs_code": str(ws.cell(r, 4).value or "").strip(),
                        "bcd_rate": str(ws.cell(r, 5).value or "").strip(),
                        "batch_type": str(ws.cell(r, 6).value or "Batch No").strip(),
                        "notes": str(ws.cell(r, 7).value or "").strip()
                    })
            self.wfile.write(json.dumps(rows).encode("utf-8"))

        elif path == "/api/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()

            status_payload = {
                "running": pipeline_state["running"],
                "logs": pipeline_state["logs"][-30:],  # last 30 lines
                "result": pipeline_state["result"],
                "error": pipeline_state["error"]
            }
            self.wfile.write(json.dumps(status_payload).encode("utf-8"))

        elif path == "/api/open_folder":
            qs = parse_qs(parsed.query)
            target = qs.get("type", ["output"])[0]
            dir_to_open = (WORKSPACE_ROOT / INPUT_FOLDER) if target == "input" else (WORKSPACE_ROOT / OUTPUT_FOLDER)
            try:
                os.startfile(str(dir_to_open))
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"OK")
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode())

        elif path == "/api/open_file":
            qs = parse_qs(parsed.query)
            target = qs.get("file", ["master"])[0]
            file_to_open = (WORKSPACE_ROOT / INPUT_FOLDER / MASTER_EXCEL) if target == "master" else (WORKSPACE_ROOT / OUTPUT_FOLDER / OUTPUT_FILENAME)
            try:
                os.startfile(str(file_to_open))
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"OK")
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(str(e).encode())

        elif path.startswith("/download/"):
            filename = unquote(path.replace("/download/", ""))
            file_path = WORKSPACE_ROOT / OUTPUT_FOLDER / filename
            if file_path.exists():
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f"attachment; filename=\"{filename}\"")
                self.send_header("Content-Length", str(file_path.stat().st_size))
                self.end_headers()
                with open(file_path, "rb") as f:
                    shutil.copyfileobj(f, self.wfile)
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"File not found.")

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/upload":
            qs = parse_qs(parsed.query)
            filename = qs.get("filename", ["unnamed_file"])[0]
            filename = Path(filename).name  # sanitize

            in_dir = WORKSPACE_ROOT / INPUT_FOLDER
            in_dir.mkdir(parents=True, exist_ok=True)
            dest = in_dir / filename

            content_length = int(self.headers.get("Content-Length", 0))
            bytes_written = 0
            chunk_size = 64 * 1024
            with open(dest, "wb") as f_out:
                while bytes_written < content_length:
                    to_read = min(chunk_size, content_length - bytes_written)
                    chunk = self.rfile.read(to_read)
                    if not chunk:
                        break
                    f_out.write(chunk)
                    bytes_written += len(chunk)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"status": "OK", "filename": filename, "bytes": bytes_written}).encode("utf-8"))

        elif path == "/api/master":
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            rows = json.loads(body.decode("utf-8"))

            master_file = WORKSPACE_ROOT / INPUT_FOLDER / MASTER_EXCEL
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Product Master"
            headers = ["Part No", "Tally Item Name", "Description", "HS Code", "BCD %", "Batch Type", "Notes"]
            ws.append(headers)

            for item in rows:
                p_no = str(item.get("part_no", "")).strip()
                if not p_no:
                    continue
                bcd_val = str(item.get("bcd_rate", "")).strip()
                try:
                    bcd_num = float(bcd_val) if bcd_val else None
                except Exception:
                    bcd_num = None

                ws.append([
                    p_no,
                    str(item.get("tally_name", "")).strip(),
                    str(item.get("description", "")).strip(),
                    str(item.get("hs_code", "")).strip(),
                    bcd_num,
                    str(item.get("batch_type", "Batch No")).strip(),
                    str(item.get("notes", "")).strip()
                ])

            wb.save(master_file)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"status": "SAVED", "count": len(rows)}).encode("utf-8"))

        elif path == "/api/master/merge":
            rev_file = WORKSPACE_ROOT / OUTPUT_FOLDER / "master_reviewed.xlsx"
            master_file = WORKSPACE_ROOT / INPUT_FOLDER / MASTER_EXCEL
            if rev_file.exists():
                wb_rev = openpyxl.load_workbook(rev_file, data_only=True)
                ws_rev = wb_rev.active
                wb_rev.save(master_file)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "MERGED", "rows": ws_rev.max_row - 1}).encode("utf-8"))
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"No reviewed master file found.")

        elif path == "/api/run":
            if pipeline_state["running"]:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Already running")
                return

            def _worker():
                pipeline_state["running"] = True
                pipeline_state["logs"] = []
                pipeline_state["result"] = None
                pipeline_state["error"] = None
                try:
                    res = run_pipeline(
                        input_folder=str(WORKSPACE_ROOT / INPUT_FOLDER),
                        output_folder=str(WORKSPACE_ROOT / OUTPUT_FOLDER),
                        base_excel=BASE_EXCEL,
                        master_excel=MASTER_EXCEL,
                        output_filename=OUTPUT_FILENAME
                    )
                    pipeline_state["result"] = res
                except Exception as ex:
                    pipeline_state["error"] = str(ex)
                    logging.exception("Pipeline failed")
                finally:
                    pipeline_state["running"] = False

            t = threading.Thread(target=_worker, daemon=True)
            t.start()

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"status": "STARTED"}).encode("utf-8"))

        else:
            self.send_response(404)
            self.end_headers()

    def do_DELETE(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/files":
            qs = parse_qs(parsed.query)
            filename = qs.get("filename", [""])[0]
            if not filename:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Missing filename")
                return

            target = WORKSPACE_ROOT / INPUT_FOLDER / Path(filename).name
            if target.name.lower() == "base.xlsx":
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"base.xlsx is the core template and cannot be deleted.")
                return

            if target.exists() and target.is_file():
                try:
                    target.unlink()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "DELETED", "filename": target.name}).encode("utf-8"))
                except Exception as e:
                    self.send_response(500)
                    self.end_headers()
                    self.wfile.write(str(e).encode())
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"File not found")
        else:
            self.send_response(404)
            self.end_headers()


def launch_server():
    server = HTTPServer((HOST, PORT), AutomationHttpServer)
    url = f"http://{HOST}:{PORT}"
    print(f"\n=======================================================")
    print(f" YASKAWA AUTOMATION WEB HUB READY")
    print(f" Web Interface: {url}")
    print(f" Supports multi-file ingestion & editable Master Excel.")
    print(f"=======================================================\n")
    if HOST == "127.0.0.1":
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down GUI server.")
        server.server_close()


if __name__ == "__main__":
    launch_server()
