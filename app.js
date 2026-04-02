/* ==========================================================================
   app.js – Import DVI / MRN PDF extractor  (100 % client-side)
   Dependencies: pdf.js, SheetJS/xlsx, JSZip (self-hosted in lib/)
   ========================================================================== */

// ─── HELPER: Romanian char class (handle both Unicode variants) ─────────────
// ș can be U+0219 or U+015F; ț can be U+021B or U+0163
function ro(pattern, flags) {
  const p = pattern
    .replace(/ș/g, "[șş]")
    .replace(/Ș/g, "[ȘŞșş]")
    .replace(/ț/g, "[țţ]")
    .replace(/Ț/g, "[ȚŢțţ]");
  return new RegExp(p, flags || "i");
}

// ─── REGEX PATTERNS (single source of truth) ────────────────────────────────
const patterns = {
  mrn: /MRN\s+(\d{2}[A-Z]{2}[A-Z0-9]{10,20})/i,
  dataMRN: /MRN\s+\S+\s+(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/i,
  awbSection: ro("Documentul\\s+de\\s+transport\\s*-\\s*\\[12\\s*05\\]"),
  awbLine: /N(?:74[01]|705|730|787)\s*\/\s*([A-Z0-9][A-Z0-9\-. ]{2,})/i,
  exportatorSection: ro("Exportatorul\\s*-\\s*\\[13\\s*01\\]"),
  moneda: ro("Moneda\\s+de\\s+facturare\\s*-\\s*\\[14\\s*05\\]\\s+([A-Z]{3})"),
  valoare: ro(
    "Cuantumul\\s+total\\s+facturat\\s*-\\s*\\[14\\s*06\\]\\s+([\\d]+(?:[.,]\\d+)?)",
  ),
  dataLiberVama: ro(
    "Dat[aăâ]\\s+liber\\s+vam[aăâ]\\s+(\\d{1,2})[\\/.\\-](\\d{1,2})[\\/.\\-](\\d{4})",
  ),
  dataAcceptarii: ro(
    "Data\\s+accept[aăâ]rii\\s*-\\s*\\[15\\s*09\\]\\s*(\\d{1,2})[\\/.\\-](\\d{1,2})[\\/.\\-](\\d{4})",
  ),
  docJustificativ: ro("Document\\s+justificativ\\s*-\\s*\\[12\\s*03\\]"),
  codTaric: /Cod\s+TARIC\s+unificat\s+(\d{6,10})/i,
  importatorSection: ro("Importatorul\\s*-\\s*\\[13\\s*04\\]"),
  taraExpediere: ro(
    "Țara\\s+de\\s+expedi(?:ere|ție)\\s*-?\\s*(?:\\[16\\s*06\\])?\\s+([A-Z]{2})",
  ),
  regimUnificat: /Regim\s+unificat\s+(\d{4})/i,
  preferinte: ro("Preferințe\\s*-\\s*\\[14\\s*11\\]\\s+(\\d+)"),
  declarantSection: ro("Declarantul\\/?\\s*Reprezentantul\\s*-\\s*\\[13\\s*14\\]"),
};

// ─── CUI FILTER ─────────────────────────────────────────────────────────────
const EXPECTED_CUI = "RO16297090";

const MONTHS_RO = [
  "Ianuarie",
  "Februarie",
  "Martie",
  "Aprilie",
  "Mai",
  "Iunie",
  "Iulie",
  "August",
  "Septembrie",
  "Octombrie",
  "Noiembrie",
  "Decembrie",
];

// ─── STATE ──────────────────────────────────────────────────────────────────
let fileRows = [];
let skippedRows = [];
let processingInProgress = false;
let abortProcessing = false;
let queuedFiles = [];
let seenFileFingerprints = new Set();

function isPdfFile(file) {
  return (
    file &&
    (file.type === "application/pdf" ||
      (file.name || "").toLowerCase().endsWith(".pdf"))
  );
}

function fileFingerprint(file) {
  return [file.name || "", file.size || 0, file.lastModified || 0].join("::");
}

function enqueueFiles(files) {
  const pdfFiles = (files || []).filter(isPdfFile);
  if (!pdfFiles.length) return;

  const freshFiles = pdfFiles.filter((file) => {
    const fp = fileFingerprint(file);
    if (seenFileFingerprints.has(fp)) return false;
    seenFileFingerprints.add(fp);
    return true;
  });

  if (!freshFiles.length) {
    console.log("[app.js] enqueueFiles: duplicate batch skipped");
    return;
  }

  queuedFiles = queuedFiles.concat(freshFiles);
  if (!processingInProgress) {
    drainFileQueue();
  }
}

async function drainFileQueue() {
  if (processingInProgress) return;
  processingInProgress = true;
  abortProcessing = false;

  try {
    while (queuedFiles.length && !abortProcessing) {
      const batch = queuedFiles.splice(0, queuedFiles.length);
      await readFiles(batch);
    }
  } finally {
    processingInProgress = false;
    abortProcessing = false;
  }
}

// ─── INIT ───────────────────────────────────────────────────────────────────
const APP_VERSION = "1.4.0";

document.addEventListener("DOMContentLoaded", () => {
  console.log("[app.js] v" + APP_VERSION + " loaded");

  // ── Library availability check ───────────────────────────────────────
  var missing = [];
  if (typeof pdfjsLib === "undefined") missing.push("pdf.js");
  if (typeof XLSX === "undefined") missing.push("SheetJS/xlsx");
  if (typeof JSZip === "undefined") missing.push("JSZip");

  if (missing.length) {
    console.error("[app.js] Missing libraries:", missing.join(", "));
    var warn = document.getElementById("warnings-container");
    if (warn) {
      warn.innerHTML =
        '<p style="color:var(--error);font-weight:700;">⛔ Librăriile nu s-au încărcat: ' +
        missing.join(", ") +
        ". Încearcă Ctrl+Shift+R (hard refresh) sau contactează administratorul de rețea.</p>";
    }
  }

  if (typeof pdfjsLib !== "undefined") {
    // Disable Worker entirely — customs declarations are 1-5 pages so
    // main-thread parsing is <100 ms and avoids ALL corporate proxy /
    // Tracking Prevention / CSP issues with Web Workers.
    // IMPORTANT: pdf.js v3 throws "No GlobalWorkerOptions.workerSrc specified"
    // if workerSrc is empty/falsy, even when disableWorker:true is passed.
    // Create a no-op worker blob URL to satisfy the check while keeping
    // everything on the main thread via disableWorker:true in getDocument().
    try {
      var noopWorkerBlob = new Blob(["// no-op worker"], { type: "application/javascript" });
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(noopWorkerBlob);
    } catch (_) {
      // Blob/URL unavailable (very old browser) — set a dummy non-empty string
      pdfjsLib.GlobalWorkerOptions.workerSrc = "noop-worker.js";
    }
    console.log("[app.js] pdfjsLib version:", pdfjsLib.version, "(no-worker mode, workerSrc:", pdfjsLib.GlobalWorkerOptions.workerSrc.substring(0, 50) + ")");
  }
  setupDropZone();
  setupButtons();

  // Show version in page
  var header = document.querySelector("header p");
  if (header) header.textContent += " — v" + APP_VERSION;
});

// ─── DROP ZONE & FILE INPUT ─────────────────────────────────────────────────
function setupDropZone() {
  const zone = document.getElementById("drop-zone");
  const input = document.getElementById("file-input");
  if (!zone || !input) return;

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    const files = [...e.dataTransfer.files];
    if (files.length) enqueueFiles(files);
  });

  input.addEventListener("change", () => {
    const files = [...input.files];
    if (files.length) enqueueFiles(files);
    input.value = "";
  });
}

function setupButtons() {
  const btnExport = document.getElementById("btn-export");
  const btnClear = document.getElementById("btn-clear");
  const btnClearAggressive = document.getElementById("btn-clear-aggressive");
  const btnZip = document.getElementById("btn-zip");
  if (btnExport)
    btnExport.addEventListener("click", () => exportXlsx(fileRows));
  if (btnClear) btnClear.addEventListener("click", clearAll);
  if (btnClearAggressive) btnClearAggressive.addEventListener("click", aggressiveClear);
  if (btnZip) btnZip.addEventListener("click", () => exportZip(fileRows));
}

// ─── READ FILES ─────────────────────────────────────────────────────────────
async function readFiles(files) {
  console.log("[app.js] readFiles called with", files.length, "files");
  const listEl = document.getElementById("file-list");
  const progressWrap = document.getElementById("progress-wrap");
  const progressFill = document.getElementById("progress-fill");
  const progressText = document.getElementById("progress-text");

  const newItems = files.map((f) => ({
    file: f,
    fileName: f.name,
    fields: null,
    warnings: [],
    status: "pending",
  }));

  newItems.forEach((item) => {
    const safeId = item.fileName.replace(/[^a-zA-Z0-9]/g, "_");
    const li = document.createElement("li");
    li.id = "fli-" + safeId;
    li.innerHTML = `
      <span class="status status-pending" id="st-fli-${safeId}">●</span>
      <span class="name">${escHtml(item.fileName)}</span>
      <span class="badge badge-pending" id="bd-fli-${safeId}">pending</span>`;
    listEl.appendChild(li);
  });

  progressWrap.classList.add("visible");

  const total = newItems.length;
  const batchT0 = performance.now();
  for (let i = 0; i < total; i++) {
    if (abortProcessing) break;
    const item = newItems[i];
    const safeId = "fli-" + item.fileName.replace(/[^a-zA-Z0-9]/g, "_");
    let text = "";
    setFileStatus(safeId, "processing");
    var elapsed = ((performance.now() - batchT0) / 1000).toFixed(1);
    progressText.textContent = `Processing file ${i + 1}/${total} — ${item.fileName} (${elapsed}s)`;
    progressFill.style.width = `${(i / total) * 100}%`;

    try {
      var itemT0 = performance.now();
      text = await extractPdfText(item.file, function (pageNo, totalPages) {
        var elapsed = ((performance.now() - batchT0) / 1000).toFixed(1);
        progressText.textContent =
          `Processing file ${i + 1}/${total} — ${item.fileName}` +
          ` (page ${pageNo}/${totalPages}) — ${elapsed}s`;
      });
      console.log(
        "[app.js] Extracted text length:",
        text.length,
        "in",
        Math.round(performance.now() - itemT0),
        "ms —",
        item.fileName,
      );
      const result = parseFields(text);
      item.fields = result.fields;
      item.warnings = result.warnings;

      // ── CUI validation: skip files not belonging to expected company ──
      var cui = extractCUI(normalizeText(text));
      if (cui && cui !== EXPECTED_CUI) {
        item.status = "skipped";
        item.skipped = true;
        item.warnings = ["CUI declarat: " + cui + " (expected " + EXPECTED_CUI + ") — fișier ignorat"];
        diagLog("SKIP: " + item.fileName + " — CUI " + cui + " ≠ " + EXPECTED_CUI);
      } else if (!cui) {
        item.status = result.warnings.length ? "warning" : "done";
        item.warnings.push("CUI declarant nedetectat — se procesează oricum");
      } else {
        item.status = result.warnings.length ? "warning" : "done";
      }
    } catch (err) {
      console.error("[app.js] Error processing", item.fileName, err);
      diagLog("ERROR: " + item.fileName + " — " + err.message);
      item.status = "error";
      item.warnings = ["Eroare la procesare: " + err.message];
      item.fields = emptyFields();
    } finally {
      // Release large transient string after parsing to help GC on big documents.
      text = "";
    }
    setFileStatus(safeId, item.status);
    if (!item.skipped) {
      fileRows.push(item);
    } else {
      skippedRows.push(item);
    }
    progressFill.style.width = `${((i + 1) / total) * 100}%`;
  }

  var totalSec = ((performance.now() - batchT0) / 1000).toFixed(1);
  var skippedCount = newItems.filter(function (it) { return it.skipped; }).length;
  var processedCount = total - skippedCount;
  progressText.textContent = `Done — ${processedCount} file(s) processed` +
    (skippedCount ? `, ${skippedCount} skipped (CUI differs)` : "") +
    ` in ${totalSec}s.`;
  renderPreview(fileRows);
  renderWarnings(fileRows.concat(skippedRows));
  document.getElementById("btn-export").disabled = false;
  var zipBtn = document.getElementById("btn-zip");
  if (zipBtn) zipBtn.disabled = false;
}

function setFileStatus(liId, status) {
  const st = document.getElementById("st-" + liId);
  const bd = document.getElementById("bd-" + liId);
  if (st) st.className = "status status-" + status;
  if (bd) {
    bd.className = "badge badge-" + status;
    bd.textContent = status;
  }
}

// ─── EXTRACT PDF TEXT ───────────────────────────────────────────────────────
function buildPageText(tc) {
  const items = tc.items.filter((it) => it.str && it.str.trim().length > 0);
  if (items.length === 0) return "";

  const lines = {};
  items.forEach((it) => {
    if (!it.transform) return;
    const y = Math.round(it.transform[5] / 2) * 2;
    const x = it.transform[4];
    if (!lines[y]) lines[y] = [];
    lines[y].push({ x, str: it.str });
  });

  const sortedYs = Object.keys(lines)
    .map(Number)
    .sort((a, b) => b - a);
  return sortedYs
    .map((y) => {
      return lines[y]
        .sort((a, b) => a.x - b.x)
        .map((it) => it.str)
        .join(" ");
    })
    .join("\n");
}

// Helper: read file as ArrayBuffer with FileReader fallback
function readFileAsArrayBuffer(file) {
  // Prefer modern API, fall back to FileReader for older browsers
  if (typeof file.arrayBuffer === "function") {
    return file.arrayBuffer();
  }
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function () { resolve(reader.result); };
    reader.onerror = function () { reject(new Error("FileReader failed: " + (reader.error || "unknown"))); };
    reader.readAsArrayBuffer(file);
  });
}

// Helper: race a promise against a timeout
function withTimeout(promise, ms, label) {
  return new Promise(function (resolve, reject) {
    var timer = setTimeout(function () {
      reject(new Error("TIMEOUT after " + (ms / 1000) + "s — " + label));
    }, ms);
    promise.then(
      function (v) { clearTimeout(timer); resolve(v); },
      function (e) { clearTimeout(timer); reject(e); }
    );
  });
}

// Diagnostic log visible in UI (appended to warnings container)
function diagLog(msg) {
  console.log("[diag]", msg);
  var el = document.getElementById("diag-log");
  if (!el) {
    var container = document.getElementById("warnings-container");
    if (container) {
      el = document.createElement("pre");
      el.id = "diag-log";
      el.style.cssText = "font-size:0.75rem;color:#666;max-height:200px;overflow:auto;white-space:pre-wrap;margin-top:8px;";
      container.appendChild(el);
    }
  }
  if (el) {
    el.textContent += "[" + new Date().toLocaleTimeString() + "] " + msg + "\n";
    el.scrollTop = el.scrollHeight;
  }
}

async function extractPdfText(file, onPageProcessed) {
  var fileT0 = performance.now();
  let arrayBuf = null;
  let pdf = null;

  try {
    // ── Step 1: Read file into memory ──────────────────────────────────
    diagLog("Step 1/3: Reading " + file.name + " (" + Math.round((file.size || 0) / 1024) + " KB)...");
    arrayBuf = await withTimeout(readFileAsArrayBuffer(file), 15000, "readFileAsArrayBuffer");
    diagLog("Step 1/3: OK — read " + arrayBuf.byteLength + " bytes in " + Math.round(performance.now() - fileT0) + " ms");

    // ── Step 2: Open PDF with pdf.js ───────────────────────────────────
    // IMPORTANT: pass Uint8Array (not raw ArrayBuffer) for max compatibility
    // with disableWorker mode in all pdf.js versions
    var pdfData = new Uint8Array(arrayBuf);
    arrayBuf = null; // free original buffer immediately

    diagLog("Step 2/3: Opening PDF with pdf.js (no-worker mode)...");
    var loadT0 = performance.now();
    var loadingTask = pdfjsLib.getDocument({
      data: pdfData,
      disableWorker: true,
    });

    pdf = await withTimeout(loadingTask.promise, 30000, "pdfjsLib.getDocument");
    diagLog("Step 2/3: OK — " + pdf.numPages + " pages, opened in " + Math.round(performance.now() - loadT0) + " ms");

    // Free pdfData after getDocument has consumed it
    pdfData = null;

    // ── Step 3: Extract text from pages ────────────────────────────────
    const numPages = pdf.numPages;
    const pages = new Array(numPages).fill("");
    const batchSize = 4;

    diagLog("Step 3/3: Extracting text from " + numPages + " pages...");

    for (let start = 1; start <= numPages; start += batchSize) {
      const end = Math.min(numPages, start + batchSize - 1);
      const batchPromises = [];

      for (let p = start; p <= end; p++) {
        batchPromises.push(
          pdf
            .getPage(p)
            .then(function (page) { return page.getTextContent(); })
            .then(function (tc) {
              pages[p - 1] = buildPageText(tc);
            })
            .catch(function (err) {
              console.warn("[app.js] Page extraction failed", file.name, "page", p, err);
              pages[p - 1] = "";
            })
            .finally(function () {
              if (typeof onPageProcessed === "function") {
                onPageProcessed(p, numPages);
              }
            }),
        );
      }

      await withTimeout(Promise.all(batchPromises), 15000, "page batch " + start + "-" + end);

      // Yield to UI thread between page batches so the app feels responsive.
      await new Promise(function (resolve) { setTimeout(resolve, 0); });
    }

    var totalMs = Math.round(performance.now() - fileT0);
    diagLog("Step 3/3: OK — all " + numPages + " pages extracted in " + totalMs + " ms total");

    return pages.join("\n\n");
  } finally {
    arrayBuf = null;

    if (pdf) {
      try {
        await pdf.cleanup();
      } catch (_) {
        // noop
      }
      try {
        await pdf.destroy();
      } catch (_) {
        // noop
      }
    }
  }
}

// ─── NORMALIZE ──────────────────────────────────────────────────────────────
function normalizeText(raw) {
  return raw
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/ ?\n ?/g, "\n")
    .replace(/\n{3,}/g, "\n\n");
}

// ─── EMPTY FIELDS ───────────────────────────────────────────────────────────
function emptyFields() {
  return {
    dvi: "",
    dataMRN: "",
    awb: "",
    exportator: "",
    taraExp: "",
    moneda: "",
    valoare: "",
    awbLunaAn: "",
    nrFactura: "",
    codTaric: "",
    regimUnificat: "",
    locatie: "",
    preferinte: "",
    gratis: "NU",
  };
}

// ─── PARSE FIELDS ───────────────────────────────────────────────────────────
function parseFields(rawText) {
  const text = normalizeText(rawText);
  const warnings = [];
  const f = emptyFields();

  // 1) DVI (MRN)
  const mrnMatch = text.match(patterns.mrn);
  if (mrnMatch) f.dvi = mrnMatch[1].trim();
  else warnings.push("MRN missing");

  // 2) Data MRN
  const dataMrnMatch = text.match(patterns.dataMRN);
  if (dataMrnMatch) {
    f.dataMRN =
      dataMrnMatch[1].padStart(2, "0") +
      "/" +
      dataMrnMatch[2].padStart(2, "0") +
      "/" +
      dataMrnMatch[3];
  } else warnings.push("Data MRN missing");

  // 3) AWB
  f.awb = extractAWB(text);
  if (!f.awb) warnings.push("AWB missing");

  // 4) EXPORTATOR
  f.exportator = extractExportator(text);
  if (!f.exportator) warnings.push("Exportator missing");

  // 5) Tara exportator
  f.taraExp = extractTaraExp(text);
  if (!f.taraExp) warnings.push("Țara exportator missing");

  // 6) Moneda + Valoare
  const monedaMatch = text.match(patterns.moneda);
  if (monedaMatch) f.moneda = monedaMatch[1].trim();
  else warnings.push("Moneda missing");

  const valMatch = text.match(patterns.valoare);
  if (valMatch) f.valoare = parseFloat(valMatch[1].replace(",", "."));
  else warnings.push("Valoare marfă missing");

  // 7) AWB - Luna An
  f.awbLunaAn = extractAwbLunaAn(text);
  if (!f.awbLunaAn) warnings.push("Data liber vamă / Data acceptării missing");

  // 8) Nr. Factură
  f.nrFactura = extractFactura(text);
  if (!f.nrFactura) warnings.push("Factura missing");

  // 9) Cod TARIC
  f.codTaric = extractCodTaric(text);
  if (!f.codTaric) warnings.push("Cod TARIC missing");

  // 10) Regim unificat
  f.regimUnificat = extractRegimUnificat(text);
  if (!f.regimUnificat) warnings.push("Regim unificat missing");

  // 11) Locatie
  f.locatie = extractLocatie(text);
  if (!f.locatie) warnings.push("Locație missing");

  // 12) Preferințe [14 11]
  const prefMatch = text.match(patterns.preferinte);
  if (prefMatch) f.preferinte = prefMatch[1].trim();
  else warnings.push("Preferințe missing");

  // 13) Gratis
  f.gratis = "NU";

  return { fields: f, warnings };
}

// ─── EXTRACTION HELPERS ─────────────────────────────────────────────────────

function extractAWB(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const docIdx = searchText.search(patterns.awbSection);
  if (docIdx < 0) {
    const m = searchText.match(patterns.awbLine);
    return m ? cleanAWB(m[1]) : "";
  }
  // Search the window after "Documentul de transport - [12 05]"
  const win = searchText.substring(docIdx, docIdx + 500);
  const lineRe = /N(74[01]|705|730|787)\s*\/\s*([A-Z0-9][A-Z0-9\-. ]{2,})/gi;
  const candidates = [];
  let match;
  while ((match = lineRe.exec(win)) !== null) {
    candidates.push({ code: match[1], value: cleanAWB(match[2]) });
  }
  if (candidates.length === 0) return "";
  // Priority: N740/N741 > N705 > N787 > N730
  // N730 is generic (CMR), prefer specific codes
  const priority = { 740: 1, 741: 1, 705: 2, 787: 3, 730: 4 };
  candidates.sort((a, b) => (priority[a.code] || 9) - (priority[b.code] || 9));
  return candidates[0].value;
}

function cleanAWB(raw) {
  // Trim and remove trailing noise (slashes, excess spaces)
  let v = raw
    .replace(/\s*\/\s*$/, "")
    .replace(/\s{2,}/g, " ")
    .trim();
  // pdf.js merges adjacent PDF columns, so AWB values may be followed by
  // section label fragments like "Antrepozit -", "Tipul", "Identificatorul".
  // Strip trailing Title-case words that are common field labels.
  v = v.replace(/(?:\s+[A-Z][a-z]\w*)+\s*-?\s*$/, "").trim();
  return v;
}

function extractExportator(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const expIdx = searchText.search(patterns.exportatorSection);
  if (expIdx < 0) return "";
  const win = searchText.substring(expIdx, expIdx + 600);
  const nameMatch = win.match(/Numele\s+([A-Z][A-Z0-9\s\-&.,()\/]{2,})/);
  if (nameMatch) {
    let name = nameMatch[1].trim();
    // Remove trailing partial words from next field labels
    name = name.replace(/\s+[A-Z]?\s*$/, "").trim(); // trailing single char
    name = name
      .replace(/\s+(Adresa|Strada|Ora|Codul|Nr[:\.]|Numele|V[aâ]nz).*$/i, "")
      .trim();
    // Final cleanup: if name ends with single letter that looks like start of next field
    name = name.replace(/\s+[A-Z]$/, "").trim();
    return name;
  }
  return "";
}

function extractTaraExp(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const expIdx = searchText.search(patterns.exportatorSection);
  if (expIdx >= 0) {
    const win = searchText.substring(expIdx, expIdx + 800);
    const taraRe = ro("[ȚT]ara\\s+([A-Z]{2})\\b", "gi");
    const matches = [...win.matchAll(taraRe)];
    for (const m of matches) {
      const code = m[1].toUpperCase();
      if (code.length === 2) return code;
    }
  }
  const fallback = text.match(patterns.taraExpediere);
  if (fallback) return fallback[1].toUpperCase();
  return "";
}

function extractAwbLunaAn(text) {
  let month, year;
  const dlv = text.match(patterns.dataLiberVama);
  if (dlv) {
    month = parseInt(dlv[2]);
    year = parseInt(dlv[3]);
  } else {
    const mrnDate = text.match(patterns.dataMRN);
    if (mrnDate) {
      month = parseInt(mrnDate[2]);
      year = parseInt(mrnDate[3]);
    } else {
      const da = text.match(patterns.dataAcceptarii);
      if (da) {
        month = parseInt(da[2]);
        year = parseInt(da[3]);
      }
    }
  }
  if (!month || !year) return "";
  return "AWB - " + (MONTHS_RO[month - 1] || "") + " " + year;
}

function extractFactura(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const djIdx = searchText.search(patterns.docJustificativ);
  if (djIdx < 0) return "";
  const afterDJ = searchText.substring(djIdx);
  const nextRe = ro(
    "Documentul\\s+de\\s+transport|Exportatorul|Loca[țt]ia\\s+m[aă]rfurilor",
  );
  const nextIdx = afterDJ.search(nextRe);
  const win =
    nextIdx > 0 ? afterDJ.substring(0, nextIdx) : afterDJ.substring(0, 1500);
  // Capture everything between "N380 / " and " / /" (the next double-slash delimiter)
  const n380 = win.match(/N380\s*\/\s*(.*?)\s*\/\s*\//i);
  if (n380) {
    let val = n380[1].trim();
    // Strip trailing date in format /DD.MM.YYYY or /YYYY-MM-DD... (attached to last invoice)
    val = val.replace(/\/\s*\d{2}\.\d{2}\.\d{4}\s*$/, "").trim();
    val = val.replace(/\/\s*\d{4}-\d{2}-\d{2}[\s\d:.]*$/, "").trim();
    if (val) return val;
  }
  const n325 = win.match(/N325\s*\/\s*(.*?)\s*\/\s*\//i);
  if (n325) {
    let val = n325[1].trim();
    val = val.replace(/\/\s*\d{2}\.\d{2}\.\d{4}\s*$/, "").trim();
    val = val.replace(/\/\s*\d{4}-\d{2}-\d{2}[\s\d:.]*$/, "").trim();
    if (val) return val;
  }
  return "";
}

function extractCodTaric(text) {
  // Extract every TARIC/HS code occurrence, keep insertion order and drop duplicates.
  const taricRe = /Cod\s+TARIC\s+unificat\s+([0-9][0-9\s]{5,19})/gi;
  const uniqueCodes = [];
  const seen = new Set();

  for (const match of text.matchAll(taricRe)) {
    const code = (match[1] || "").replace(/\s+/g, "").trim();
    if (!/^\d{6,10}$/.test(code)) continue;
    if (seen.has(code)) continue;
    seen.add(code);
    uniqueCodes.push(code);
  }

  return uniqueCodes.join("; ");
}

function extractRegimUnificat(text) {
  const m = text.match(patterns.regimUnificat);
  return m ? m[1].trim() : "";
}

function extractCUI(text) {
  // Look in Declarantul/Reprezentantul section [13 14]
  var declIdx = text.search(patterns.declarantSection);
  if (declIdx >= 0) {
    var win = text.substring(declIdx, declIdx + 600);
    // CUI format: RO followed by digits (e.g. RO1629090)
    var cuiMatch = win.match(/\bRO\s*(\d{5,10})\b/);
    if (cuiMatch) return "RO" + cuiMatch[1];
  }
  // Fallback: also check Importatorul section [13 04]
  var impIdx = text.search(patterns.importatorSection);
  if (impIdx >= 0) {
    var win2 = text.substring(impIdx, impIdx + 600);
    var cuiMatch2 = win2.match(/\bRO\s*(\d{5,10})\b/);
    if (cuiMatch2) return "RO" + cuiMatch2[1];
  }
  return "";
}

function extractLocatie(text) {
  // Find Importatorul section and extract Orașul (city)
  // Note: don't use ro() here because Ora[sș]ul already contains ș in a char class
  // ro() would nest it as [s[șş]] which is invalid. Use direct regex instead.
  const orasRe = /Ora[sșş]ul\s+(\S+(?:\s+\S+){0,4})/i;
  const impIdx = text.search(patterns.importatorSection);
  if (impIdx >= 0) {
    const win = text.substring(impIdx, impIdx + 800);
    const m = win.match(orasRe);
    if (m) {
      let city = m[1].trim();
      // Remove trailing noise from adjacent fields ([\s\S]* to cross newlines)
      city = city
        .replace(
          /\s+(Codul|Total|MRN|LRN|Dat[aă]|Biroul|po[sșş]tal)[\s\S]*$/i,
          "",
        )
        .trim();
      if (city.length >= 2) return city;
    }
  }
  return "";
}

// ─── COLUMN DEFINITIONS ─────────────────────────────────────────────────────
const COLUMNS = [
  { key: "awb", label: "AWB" },
  { key: "exportator", label: "Exportator" },
  { key: "taraExp", label: "Țara Exp." },
  { key: "nrFactura", label: "Nr. Factură" },
  { key: "valoare", label: "Valoare Marfă" },
  { key: "moneda", label: "Moneda" },
  { key: "dvi", label: "MRN" },
  { key: "dataMRN", label: "Data MRN" },
  { key: "codTaric", label: "Cod TARIC" },
  { key: "regimUnificat", label: "Regim unificat" },
  { key: "preferinte", label: "Preferințe" },
  { key: "locatie", label: "Locație" },
  { key: "gratis", label: "Gratis", type: "select", options: ["NU", "DA"] },
];

const XLSX_EXPORT_COLUMNS = [
  { key: "awb", label: "AWB" },
  { key: "exportator", label: "Exportator" },
  { key: "taraExp", label: "Țara Exp." },
  { key: "nrFactura", label: "Nr. Factură" },
  { key: "valoare", label: "Valoare Marfă" },
  { key: "moneda", label: "Moneda" },
  { key: "dvi", label: "MRN" },
  { key: "dataMRN", label: "Data MRN" },
  { key: "codTaric", label: "Cod TARIC" },
  { key: "regimUnificat", label: "Regim unificat" },
  { key: "preferinte", label: "Preferințe" },
  { key: "fileName", label: "Fișier" },
  { key: "locatie", label: "Locație" },
  { key: "gratis", label: "Gratis" },
];

function buildXlsxExportData(rows) {
  var headers = ["#"];
  XLSX_EXPORT_COLUMNS.forEach(function (col) {
    headers.push(col.label);
  });

  var data = rows.map(function (row, i) {
    var f = row.fields || emptyFields();
    var rowData = [i + 1];

    XLSX_EXPORT_COLUMNS.forEach(function (col) {
      var value = col.key === "fileName" ? row.fileName : f[col.key];
      rowData.push(value != null ? value : "");
    });

    return rowData;
  });

  return { headers: headers, data: data };
}

// ─── RENDER PREVIEW TABLE ───────────────────────────────────────────────────
function renderPreview(rows) {
  const section = document.getElementById("preview-section");
  const thead = document.querySelector("#preview-section thead");
  const tbody = document.getElementById("preview-tbody");

  // Build header row
  thead.innerHTML = "";
  const headerRow = document.createElement("tr");
  headerRow.innerHTML =
    "<th>#</th><th>Fișier</th>" +
    COLUMNS.map((c) => "<th>" + c.label + "</th>").join("");
  thead.appendChild(headerRow);

  // Build filter row
  const filterRow = document.createElement("tr");
  filterRow.className = "filter-row";
  filterRow.innerHTML =
    '<th></th><th><input type="text" class="col-filter" data-col="fileName" placeholder="🔍"></th>' +
    COLUMNS.map(
      (c) =>
        '<th><input type="text" class="col-filter" data-col="' +
        c.key +
        '" placeholder="🔍"></th>',
    ).join("");
  thead.appendChild(filterRow);

  renderTableRows(rows, tbody);

  // Bind filter events
  thead.querySelectorAll(".col-filter").forEach((input) => {
    input.addEventListener("input", function () {
      applyFilters(rows, tbody, thead);
    });
  });

  section.style.display = "block";
}

function renderTableRows(rows, tbody, visibleIndices) {
  tbody.innerHTML = "";
  const indices = visibleIndices || rows.map((_, i) => i);

  indices.forEach((idx) => {
    const row = rows[idx];
    const f = row.fields || emptyFields();
    const tr = document.createElement("tr");
    let html =
      "<td>" + (idx + 1) + "</td><td>" + escHtml(row.fileName) + "</td>";

    COLUMNS.forEach((col) => {
      const val = f[col.key] != null ? String(f[col.key]) : "";
      if (col.type === "select") {
        const opts = col.options
          .map(
            (o) =>
              '<option value="' +
              o +
              '"' +
              (val === o ? " selected" : "") +
              ">" +
              o +
              "</option>",
          )
          .join("");
        html +=
          '<td><select class="editable" data-idx="' +
          idx +
          '" data-field="' +
          col.key +
          '">' +
          opts +
          "</select></td>";
      } else {
        html +=
          '<td><input class="editable" type="text" value="' +
          escAttr(val) +
          '" data-idx="' +
          idx +
          '" data-field="' +
          col.key +
          '"></td>';
      }
    });

    tr.innerHTML = html;
    tbody.appendChild(tr);
  });

  // Bind editable events
  bindEditableEvents(tbody);
}

function bindEditableEvents(tbody) {
  tbody.querySelectorAll(".editable").forEach((el) => {
    el.addEventListener("change", function (e) {
      const idx = parseInt(e.target.dataset.idx);
      const field = e.target.dataset.field;
      var val = e.target.value;
      if (field === "valoare") {
        var n = parseFloat(val.replace(",", "."));
        if (!isNaN(n)) val = n;
      }
      fileRows[idx].fields[field] = val;
    });
    if (el.tagName === "INPUT") {
      el.addEventListener("input", function (e) {
        const idx = parseInt(e.target.dataset.idx);
        const field = e.target.dataset.field;
        if (field !== "valoare") {
          fileRows[idx].fields[field] = e.target.value;
        }
      });
    }
  });
}

function applyFilters(rows, tbody, thead) {
  const filters = {};
  thead.querySelectorAll(".col-filter").forEach((input) => {
    const col = input.dataset.col;
    const val = input.value.trim().toLowerCase();
    if (val) filters[col] = val;
  });

  if (Object.keys(filters).length === 0) {
    renderTableRows(rows, tbody);
    return;
  }

  const visibleIndices = [];
  rows.forEach((row, idx) => {
    const f = row.fields || emptyFields();
    let match = true;
    for (const [col, query] of Object.entries(filters)) {
      let cellVal = "";
      if (col === "fileName") {
        cellVal = (row.fileName || "").toLowerCase();
      } else {
        cellVal = (f[col] != null ? String(f[col]) : "").toLowerCase();
      }
      if (cellVal.indexOf(query) < 0) {
        match = false;
        break;
      }
    }
    if (match) visibleIndices.push(idx);
  });

  renderTableRows(rows, tbody, visibleIndices);
}

function escHtml(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escAttr(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ─── RENDER WARNINGS ────────────────────────────────────────────────────────
function renderWarnings(rows) {
  var container = document.getElementById("warnings-container");
  container.innerHTML = "";
  var withWarnings = rows.filter(function (r) {
    return r.warnings && r.warnings.length > 0;
  });
  if (withWarnings.length === 0) {
    container.innerHTML =
      '<p style="color:var(--success);font-size:.85rem;">✓ No warnings.</p>';
    return;
  }
  withWarnings.forEach(function (row) {
    var div = document.createElement("div");
    div.className = "warning-block";
    div.innerHTML =
      '<div class="file-name">⚠ ' +
      escHtml(row.fileName) +
      "</div><ul>" +
      row.warnings
        .map(function (w) {
          return "<li>" + escHtml(w) + "</li>";
        })
        .join("") +
      "</ul>";
    container.appendChild(div);
  });
}

// ─── EXPORT XLSX ────────────────────────────────────────────────────────────
function exportXlsx(rows) {
  if (!rows.length) return;

  // Visual feedback — disable button and show generating label
  var btn = document.getElementById("btn-export");
  var origLabel = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ Se generează...";
  }

  // Defer heavy XLSX work so the browser can repaint the button first
  setTimeout(function () {
    try {
      var exportData = buildXlsxExportData(rows);
      var ws = XLSX.utils.aoa_to_sheet(
        [exportData.headers].concat(exportData.data),
      );
      ws["!cols"] = [
        { wch: 4 },
        { wch: 24 },
        { wch: 34 },
        { wch: 10 },
        { wch: 24 },
        { wch: 14 },
        { wch: 10 },
        { wch: 24 },
        { wch: 12 },
        { wch: 28 },
        { wch: 14 },
        { wch: 14 },
        { wch: 35 },
        { wch: 20 },
        { wch: 8 },
      ];

      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Import DVI");
      XLSX.writeFile(wb, "Extract_DVI_AWB.xlsx");
      console.log("[app.js] Excel exported successfully");
    } catch (err) {
      console.error("[app.js] Export error:", err);
      alert("Eroare la export Excel: " + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = origLabel;
      }
    }
  }, 50);
}

// ─── EXPORT ZIP (PDFs renamed by MRN) ───────────────────────────────────────
async function exportZip(rows) {
  if (!rows.length) return;
  if (typeof JSZip === "undefined") {
    alert("JSZip library not loaded. Cannot create ZIP.");
    return;
  }

  var btn = document.getElementById("btn-zip");
  var origLabel = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ Se creează ZIP...";
  }

  try {
    var zip = new JSZip();
    var usedNames = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var mrn = row.fields && row.fields.dvi ? row.fields.dvi : "";
      var baseName = mrn || "Unknown_MRN_" + (i + 1);

      // Handle duplicate names
      if (usedNames[baseName]) {
        usedNames[baseName]++;
        baseName = baseName + "_" + usedNames[baseName];
      } else {
        usedNames[baseName] = 1;
      }

      var fileName = baseName + ".pdf";

      if (row.file) {
        var arrayBuf = await row.file.arrayBuffer();
        zip.file(fileName, arrayBuf);
        // Release loop-local reference quickly when processing many files.
        arrayBuf = null;
      }
    }

    var content = await zip.generateAsync({ type: "blob" });
    var link = document.createElement("a");
    link.href = URL.createObjectURL(content);
    link.download = "DVI_PDFs_Renamed.zip";
    link.click();
    URL.revokeObjectURL(link.href);
    console.log("[app.js] ZIP exported successfully");
  } catch (err) {
    console.error("[app.js] ZIP export error:", err);
    alert("Eroare la export ZIP: " + err.message);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = origLabel;
    }
  }
}

// ─── CLEAR ──────────────────────────────────────────────────────────────────
function clearAll() {
  // Signal any in-progress processing to stop.
  abortProcessing = true;
  // Explicitly break references to large objects (File/Blob) before reset.
  fileRows.forEach(function (row) {
    row.file = null;
    row.fields = null;
    row.warnings = null;
  });
  skippedRows.forEach(function (row) {
    row.file = null;
    row.fields = null;
    row.warnings = null;
  });
  fileRows = [];
  skippedRows = [];
  queuedFiles = [];
  seenFileFingerprints = new Set();
  document.getElementById("file-list").innerHTML = "";
  document.getElementById("preview-tbody").innerHTML = "";
  document.getElementById("preview-section").style.display = "none";
  document.getElementById("warnings-container").innerHTML = "";
  document.getElementById("progress-wrap").classList.remove("visible");
  document.getElementById("progress-fill").style.width = "0%";
  document.getElementById("progress-text").textContent = "";
  document.getElementById("btn-export").disabled = true;
  var zipBtnClear = document.getElementById("btn-zip");
  if (zipBtnClear) zipBtnClear.disabled = true;
}

function aggressiveClear() {
  var hasWork = fileRows.length > 0 || queuedFiles.length > 0 || processingInProgress;
  var msg = hasWork
    ? "Curățare agresivă: toate datele din sesiunea curentă vor fi șterse și pagina va fi reîncărcată. Continui?"
    : "Pagina va fi reîncărcată pentru reset complet de memorie. Continui?";

  if (!window.confirm(msg)) return;

  clearAll();

  // Some browsers expose manual GC in special modes; call it if available.
  if (typeof window.gc === "function") {
    try {
      window.gc();
    } catch (_) {
      // noop
    }
  }

  // Hard reset the tab context to release all runtime references quickly.
  setTimeout(function () {
    window.location.reload();
  }, 80);
}

// ─── EXPORTS FOR TESTING ────────────────────────────────────────────────────
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    normalizeText,
    parseFields,
    emptyFields,
    patterns,
    ro,
    extractAWB,
    extractExportator,
    extractTaraExp,
    extractAwbLunaAn,
    extractFactura,
    extractCodTaric,
    extractLocatie,
    extractRegimUnificat,
    extractCUI,
    EXPECTED_CUI,
    COLUMNS,
    XLSX_EXPORT_COLUMNS,
    buildXlsxExportData,
    extractPreferinte: function (text) {
      var m = text.match(patterns.preferinte);
      return m ? m[1].trim() : "";
    },
    isPdfFile,
    fileFingerprint,
    escHtml,
    escAttr,
    buildPageText,
    APP_VERSION,
  };
}
