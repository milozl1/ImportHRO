/* ==========================================================================
   app.js – Import DVI / MRN PDF extractor  (100 % client-side)
   Dependencies: pdf.js (CDN), SheetJS/xlsx (CDN)
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
};

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

// ─── INIT ───────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  if (typeof pdfjsLib !== "undefined") {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    console.log("[app.js] Initialized. pdfjsLib version:", pdfjsLib.version);
  }
  setupDropZone();
  setupButtons();
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
    const files = [...e.dataTransfer.files].filter(
      (f) =>
        f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf"),
    );
    if (files.length) readFiles(files);
  });

  input.addEventListener("change", () => {
    const files = [...input.files];
    if (files.length) readFiles(files);
    input.value = "";
  });
}

function setupButtons() {
  const btnExport = document.getElementById("btn-export");
  const btnClear = document.getElementById("btn-clear");
  const btnZip = document.getElementById("btn-zip");
  if (btnExport)
    btnExport.addEventListener("click", () => exportXlsx(fileRows));
  if (btnClear) btnClear.addEventListener("click", clearAll);
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
  for (let i = 0; i < total; i++) {
    const item = newItems[i];
    const safeId = "fli-" + item.fileName.replace(/[^a-zA-Z0-9]/g, "_");
    setFileStatus(safeId, "processing");
    progressText.textContent = `Processing file ${i + 1}/${total} — ${item.fileName}`;
    progressFill.style.width = `${(i / total) * 100}%`;

    try {
      const text = await extractPdfText(item.file);
      console.log(
        "[app.js] Extracted text length:",
        text.length,
        "first 400 chars:",
        text.substring(0, 400),
      );
      const result = parseFields(text);
      item.fields = result.fields;
      item.warnings = result.warnings;
      item.status = result.warnings.length ? "warning" : "done";
    } catch (err) {
      console.error("[app.js] Error processing", item.fileName, err);
      item.status = "error";
      item.warnings = ["Eroare la procesare: " + err.message];
      item.fields = emptyFields();
    }
    setFileStatus(safeId, item.status);
    fileRows.push(item);
    progressFill.style.width = `${((i + 1) / total) * 100}%`;
  }

  progressText.textContent = `Done — ${total} file(s) processed.`;
  renderPreview(fileRows);
  renderWarnings(fileRows);
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
async function extractPdfText(file) {
  const arrayBuf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuf }).promise;
  const numPages = pdf.numPages;

  const pagePromises = [];
  for (let p = 1; p <= numPages; p++) {
    pagePromises.push(
      pdf
        .getPage(p)
        .then((page) => page.getTextContent())
        .then((tc) => {
          const items = tc.items.filter(
            (it) => it.str && it.str.trim().length > 0,
          );
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
        }),
    );
  }

  const pages = await Promise.all(pagePromises);
  pdf.destroy();
  return pages.join("\n\n");
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
  { key: "dvi", label: "DVI (MRN)" },
  { key: "dataMRN", label: "Data MRN" },
  { key: "awb", label: "AWB" },
  { key: "exportator", label: "Exportator" },
  { key: "taraExp", label: "Țara Exp." },
  { key: "moneda", label: "Moneda" },
  { key: "valoare", label: "Valoare Marfă" },
  { key: "awbLunaAn", label: "AWB - Luna An" },
  { key: "nrFactura", label: "Nr. Factură" },
  { key: "codTaric", label: "Cod TARIC" },
  { key: "regimUnificat", label: "Regim unificat" },
  { key: "locatie", label: "Locație" },
  { key: "preferinte", label: "Preferințe" },
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
  fileRows = [];
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
    COLUMNS,
    XLSX_EXPORT_COLUMNS,
    buildXlsxExportData,
    extractPreferinte: function (text) {
      var m = text.match(patterns.preferinte);
      return m ? m[1].trim() : "";
    },
  };
}
