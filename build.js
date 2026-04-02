#!/usr/bin/env node
/* ==========================================================================
   build.js – Bundles the app into a single self-contained HTML file.
   Output: ImportHRO.html (works offline, from SharePoint, file://, etc.)
   Run:    node build.js
   ========================================================================== */

const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const OUT = path.join(ROOT, "ImportHRO.html");

function read(rel) {
  return fs.readFileSync(path.join(ROOT, rel), "utf8");
}

console.log("[build] Reading source files…");

const css = read("styles.css");
const pdfJs = read("lib/pdf.min.js");
const xlsxJs = read("lib/xlsx.full.min.js");
const jszipJs = read("lib/jszip.min.js");
const appJs = read("app.js");

console.log("[build] Building single-file bundle…");

const html = `<!DOCTYPE html>
<html lang="ro">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Import DVI / MRN – Extractor</title>
  <style>
${css}
  </style>
</head>
<body>

<div class="container">

  <!-- HEADER -->
  <header>
    <h1>📦 Import DVI / MRN – Extractor</h1>
    <p>Încarcă declarații vamale PDF → extrage automat câmpurile → exportă Excel</p>
  </header>

  <!-- UPLOAD -->
  <div class="card">
    <h2>📂 Încărcare fișiere PDF</h2>
    <div class="drop-zone" id="drop-zone">
      <input type="file" id="file-input" accept=".pdf" multiple>
      <div class="icon">📄</div>
      <p><strong>Drag & Drop</strong> fișiere PDF aici sau <strong>click</strong> pentru a selecta</p>
      <p class="hint">Declarații vamale cu text selectabil (fără OCR)</p>
    </div>

    <!-- FILE LIST -->
    <ul class="file-list" id="file-list"></ul>

    <!-- PROGRESS -->
    <div class="progress-wrap" id="progress-wrap">
      <div class="progress-bar-bg">
        <div class="progress-bar-fill" id="progress-fill"></div>
      </div>
      <div class="progress-text" id="progress-text"></div>
    </div>

    <!-- BUTTONS -->
    <div class="btn-row">
      <button class="btn btn-success" id="btn-export" disabled>📥 Export XLSX</button>
      <button class="btn btn-primary" id="btn-zip" disabled>📦 Export ZIP (MRN rename)</button>
      <button class="btn btn-outline" id="btn-clear">🗑️ Clear</button>
      <button class="btn btn-danger" id="btn-clear-aggressive">🧹 Curățare agresivă</button>
    </div>
  </div>

  <!-- PREVIEW TABLE -->
  <div class="card" id="preview-section" style="display:none;">
    <h2>📋 Preview date extrase</h2>
    <div class="table-wrap">
      <table class="preview-table">
        <thead>
        </thead>
        <tbody id="preview-tbody"></tbody>
      </table>
    </div>
  </div>

  <!-- WARNINGS -->
  <div class="card">
    <h2>⚠️ Warnings</h2>
    <div id="warnings-container">
      <p style="color:var(--text-muted); font-size:.85rem;">Niciun fișier procesat încă.</p>
    </div>
  </div>

</div>

<!-- pdf.js (inlined) -->
<script>
${pdfJs}
</script>

<!-- SheetJS (inlined) -->
<script>
${xlsxJs}
</script>

<!-- JSZip (inlined) -->
<script>
${jszipJs}
</script>

<!-- App (inlined) -->
<script>
${appJs}
</script>

</body>
</html>`;

fs.writeFileSync(OUT, html, "utf8");

const sizeMB = (Buffer.byteLength(html, "utf8") / (1024 * 1024)).toFixed(1);
console.log("[build] ✅ Created: " + OUT + " (" + sizeMB + " MB)");
console.log("[build] This single file works offline — no external dependencies.");
