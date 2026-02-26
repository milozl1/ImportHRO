/* ==========================================================================
   app.js – Import DVI / MRN PDF extractor  (100 % client-side)
   Dependencies: pdf.js (CDN), SheetJS/xlsx (CDN)
   ========================================================================== */

// ─── HELPER: Romanian char class (handle both Unicode variants) ─────────────
// ș can be U+0219 or U+015F; ț can be U+021B or U+0163
function ro(pattern, flags) {
  const p = pattern
    .replace(/ș/g, '[șş]').replace(/Ș/g, '[ȘŞșş]')
    .replace(/ț/g, '[țţ]').replace(/Ț/g, '[ȚŢțţ]');
  return new RegExp(p, flags || 'i');
}

// ─── REGEX PATTERNS (single source of truth) ────────────────────────────────
const patterns = {
  mrn: /MRN\s+(\d{2}[A-Z]{2}[A-Z0-9]{10,20})/i,
  dataMRN: /MRN\s+\S+\s+(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/i,
  awbSection: ro('Documentul\\s+de\\s+transport\\s*-\\s*\\[12\\s*05\\]'),
  awbLine: /N74[01]\s*\/\s*([A-Z0-9][A-Z0-9\-]{2,})/i,
  exportatorSection: ro('Exportatorul\\s*-\\s*\\[13\\s*01\\]'),
  moneda: ro('Moneda\\s+de\\s+facturare\\s*-\\s*\\[14\\s*05\\]\\s+([A-Z]{3})'),
  valoare: ro('Cuantumul\\s+total\\s+facturat\\s*-\\s*\\[14\\s*06\\]\\s+([\\d]+(?:[.,]\\d+)?)'),
  dataLiberVama: ro('Dat[aăâ]\\s+liber\\s+vam[aăâ]\\s+(\\d{1,2})[\\/.\\-](\\d{1,2})[\\/.\\-](\\d{4})'),
  dataAcceptarii: ro('Data\\s+accept[aăâ]rii\\s*-\\s*\\[15\\s*09\\]\\s*(\\d{1,2})[\\/.\\-](\\d{1,2})[\\/.\\-](\\d{4})'),
  docJustificativ: ro('Document\\s+justificativ\\s*-\\s*\\[12\\s*03\\]'),
  codTaric: /Cod\s+TARIC\s+unificat\s+(\d{6,10})/i,
  importatorSection: ro('Importatorul\\s*-\\s*\\[13\\s*04\\]'),
  taraExpediere: ro('Țara\\s+de\\s+expedi(?:ere|ție)\\s*-?\\s*(?:\\[16\\s*06\\])?\\s+([A-Z]{2})'),
};

const MONTHS_RO = [
  'Ianuarie', 'Februarie', 'Martie', 'Aprilie', 'Mai', 'Iunie',
  'Iulie', 'August', 'Septembrie', 'Octombrie', 'Noiembrie', 'Decembrie'
];

// ─── STATE ──────────────────────────────────────────────────────────────────
let fileRows = [];

// ─── INIT ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    console.log('[app.js] Initialized. pdfjsLib version:', pdfjsLib.version);
  }
  setupDropZone();
  setupButtons();
});

// ─── DROP ZONE & FILE INPUT ─────────────────────────────────────────────────
function setupDropZone() {
  const zone = document.getElementById('drop-zone');
  const input = document.getElementById('file-input');
  if (!zone || !input) return;

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    const files = [...e.dataTransfer.files].filter(f =>
      f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));
    if (files.length) readFiles(files);
  });

  input.addEventListener('change', () => {
    const files = [...input.files];
    if (files.length) readFiles(files);
    input.value = '';
  });
}

function setupButtons() {
  const btnExport = document.getElementById('btn-export');
  const btnClear = document.getElementById('btn-clear');
  if (btnExport) btnExport.addEventListener('click', () => exportXlsx(fileRows));
  if (btnClear) btnClear.addEventListener('click', clearAll);
}

// ─── READ FILES ─────────────────────────────────────────────────────────────
async function readFiles(files) {
  console.log('[app.js] readFiles called with', files.length, 'files');
  const listEl = document.getElementById('file-list');
  const progressWrap = document.getElementById('progress-wrap');
  const progressFill = document.getElementById('progress-fill');
  const progressText = document.getElementById('progress-text');

  const newItems = files.map(f => ({
    file: f, fileName: f.name, fields: null, warnings: [], status: 'pending'
  }));

  newItems.forEach(item => {
    const safeId = item.fileName.replace(/[^a-zA-Z0-9]/g, '_');
    const li = document.createElement('li');
    li.id = 'fli-' + safeId;
    li.innerHTML = `
      <span class="status status-pending" id="st-fli-${safeId}">●</span>
      <span class="name">${escHtml(item.fileName)}</span>
      <span class="badge badge-pending" id="bd-fli-${safeId}">pending</span>`;
    listEl.appendChild(li);
  });

  progressWrap.classList.add('visible');

  const total = newItems.length;
  for (let i = 0; i < total; i++) {
    const item = newItems[i];
    const safeId = 'fli-' + item.fileName.replace(/[^a-zA-Z0-9]/g, '_');
    setFileStatus(safeId, 'processing');
    progressText.textContent = `Processing file ${i + 1}/${total} — ${item.fileName}`;
    progressFill.style.width = `${((i) / total) * 100}%`;

    try {
      const text = await extractPdfText(item.file);
      console.log('[app.js] Extracted text length:', text.length, 'first 400 chars:', text.substring(0, 400));
      const result = parseFields(text);
      item.fields = result.fields;
      item.warnings = result.warnings;
      item.status = result.warnings.length ? 'warning' : 'done';
    } catch (err) {
      console.error('[app.js] Error processing', item.fileName, err);
      item.status = 'error';
      item.warnings = ['Eroare la procesare: ' + err.message];
      item.fields = emptyFields();
    }
    setFileStatus(safeId, item.status);
    fileRows.push(item);
    progressFill.style.width = `${((i + 1) / total) * 100}%`;
  }

  progressText.textContent = `Done — ${total} file(s) processed.`;
  renderPreview(fileRows);
  renderWarnings(fileRows);
  document.getElementById('btn-export').disabled = false;
}

function setFileStatus(liId, status) {
  const st = document.getElementById('st-' + liId);
  const bd = document.getElementById('bd-' + liId);
  if (st) st.className = 'status status-' + status;
  if (bd) { bd.className = 'badge badge-' + status; bd.textContent = status; }
}

// ─── EXTRACT PDF TEXT ───────────────────────────────────────────────────────
async function extractPdfText(file) {
  const arrayBuf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(arrayBuf) }).promise;
  const numPages = pdf.numPages;

  const pagePromises = [];
  for (let p = 1; p <= numPages; p++) {
    pagePromises.push(
      pdf.getPage(p).then(page => page.getTextContent()).then(tc => {
        const items = tc.items.filter(it => it.str && it.str.trim().length > 0);
        if (items.length === 0) return '';

        const lines = {};
        items.forEach(it => {
          const y = Math.round(it.transform[5] / 2) * 2;
          const x = it.transform[4];
          if (!lines[y]) lines[y] = [];
          lines[y].push({ x, str: it.str });
        });

        const sortedYs = Object.keys(lines).map(Number).sort((a, b) => b - a);
        return sortedYs.map(y => {
          return lines[y].sort((a, b) => a.x - b.x).map(it => it.str).join(' ');
        }).join('\n');
      })
    );
  }

  const pages = await Promise.all(pagePromises);
  return pages.join('\n\n');
}

// ─── NORMALIZE ──────────────────────────────────────────────────────────────
function normalizeText(raw) {
  return raw
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/ ?\n ?/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}

// ─── EMPTY FIELDS ───────────────────────────────────────────────────────────
function emptyFields() {
  return {
    dvi: '', dataMRN: '', awb: '', exportator: '', taraExp: '',
    moneda: '', valoare: '', awbLunaAn: '', nrFactura: '',
    codTaric: '', locatie: '', gratis: 'NU'
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
  else warnings.push('MRN missing');

  // 2) Data MRN
  const dataMrnMatch = text.match(patterns.dataMRN);
  if (dataMrnMatch) {
    f.dataMRN = dataMrnMatch[1].padStart(2,'0') + '/' + dataMrnMatch[2].padStart(2,'0') + '/' + dataMrnMatch[3];
  } else warnings.push('Data MRN missing');

  // 3) AWB
  f.awb = extractAWB(text);
  if (!f.awb) warnings.push('AWB missing');

  // 4) EXPORTATOR
  f.exportator = extractExportator(text);
  if (!f.exportator) warnings.push('Exportator missing');

  // 5) Tara exportator
  f.taraExp = extractTaraExp(text);
  if (!f.taraExp) warnings.push('Țara exportator missing');

  // 6) Moneda + Valoare
  const monedaMatch = text.match(patterns.moneda);
  if (monedaMatch) f.moneda = monedaMatch[1].trim();
  else warnings.push('Moneda missing');

  const valMatch = text.match(patterns.valoare);
  if (valMatch) f.valoare = parseFloat(valMatch[1].replace(',', '.'));
  else warnings.push('Valoare marfă missing');

  // 7) AWB - Luna An
  f.awbLunaAn = extractAwbLunaAn(text);
  if (!f.awbLunaAn) warnings.push('Data liber vamă / Data acceptării missing');

  // 8) Nr. Factură
  f.nrFactura = extractFactura(text);
  if (!f.nrFactura) warnings.push('Factura missing');

  // 9) Cod TARIC
  const taricMatch = text.match(patterns.codTaric);
  if (taricMatch) f.codTaric = taricMatch[1].trim();
  else warnings.push('Cod TARIC missing');

  // 10) Locatie
  f.locatie = extractLocatie(text);
  if (!f.locatie) warnings.push('Locație missing');

  // 11) Gratis
  f.gratis = 'NU';

  return { fields: f, warnings };
}

// ─── EXTRACTION HELPERS ─────────────────────────────────────────────────────

function extractAWB(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  if (transportIdx < 0) {
    const m = text.match(patterns.awbLine);
    return m ? m[1].trim() : '';
  }
  const transportText = text.substring(transportIdx);
  const docIdx = transportText.search(patterns.awbSection);
  if (docIdx < 0) return '';
  const win = transportText.substring(docIdx, docIdx + 500);
  const m = win.match(patterns.awbLine);
  return m ? m[1].trim() : '';
}

function extractExportator(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const expIdx = searchText.search(patterns.exportatorSection);
  if (expIdx < 0) return '';
  const win = searchText.substring(expIdx, expIdx + 600);
  const nameMatch = win.match(/Numele\s+([A-Z][A-Z0-9\s\-&.,()\/]{2,})/);
  if (nameMatch) {
    let name = nameMatch[1].trim();
    // Remove trailing partial words from next field labels
    name = name.replace(/\s+[A-Z]?\s*$/,'').trim();  // trailing single char
    name = name.replace(/\s+(Adresa|Strada|Ora|Codul|Nr[:\.]|Numele|V[aâ]nz).*$/i, '').trim();
    // Final cleanup: if name ends with single letter that looks like start of next field
    name = name.replace(/\s+[A-Z]$/,'').trim();
    return name;
  }
  return '';
}

function extractTaraExp(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const expIdx = searchText.search(patterns.exportatorSection);
  if (expIdx >= 0) {
    const win = searchText.substring(expIdx, expIdx + 800);
    const taraRe = ro('[ȚT]ara\\s+([A-Z]{2})\\b', 'gi');
    const matches = [...win.matchAll(taraRe)];
    for (const m of matches) {
      const code = m[1].toUpperCase();
      if (code.length === 2) return code;
    }
  }
  const fallback = text.match(patterns.taraExpediere);
  if (fallback) return fallback[1].toUpperCase();
  return '';
}

function extractAwbLunaAn(text) {
  let month, year;
  const dlv = text.match(patterns.dataLiberVama);
  if (dlv) { month = parseInt(dlv[2]); year = parseInt(dlv[3]); }
  else {
    const mrnDate = text.match(patterns.dataMRN);
    if (mrnDate) { month = parseInt(mrnDate[2]); year = parseInt(mrnDate[3]); }
    else {
      const da = text.match(patterns.dataAcceptarii);
      if (da) { month = parseInt(da[2]); year = parseInt(da[3]); }
    }
  }
  if (!month || !year) return '';
  return 'AWB - ' + (MONTHS_RO[month - 1] || '') + ' ' + year;
}

function extractFactura(text) {
  const transportIdx = text.search(/SEGMENT\s+TRANSPORT/i);
  const searchText = transportIdx >= 0 ? text.substring(transportIdx) : text;
  const djIdx = searchText.search(patterns.docJustificativ);
  if (djIdx < 0) return '';
  const afterDJ = searchText.substring(djIdx);
  const nextRe = ro('Documentul\\s+de\\s+transport|Exportatorul|Loca[țt]ia\\s+m[aă]rfurilor');
  const nextIdx = afterDJ.search(nextRe);
  const win = nextIdx > 0 ? afterDJ.substring(0, nextIdx) : afterDJ.substring(0, 1500);
  const n380 = win.match(/N380\s*\/\s*([A-Z0-9][A-Z0-9\-\/.]*?)(?:\s*\/|\s{2,}|\n|$)/i);
  if (n380) return n380[1].trim();
  const n325 = win.match(/N325\s*\/\s*(?:F\.?\s*TR\.?\s*)?([A-Z0-9][A-Z0-9\-\/.]*?)(?:\s*\/|\s{2,}|\n|$)/i);
  if (n325) return n325[1].trim();
  return '';
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
      city = city.replace(/\s+(Codul|Total|MRN|LRN|Dat[aă]|Biroul|po[sșş]tal)[\s\S]*$/i, '').trim();
      if (city.length >= 2) return city;
    }
  }
  return '';
}

// ─── COLUMN DEFINITIONS ─────────────────────────────────────────────────────
const COLUMNS = [
  { key: 'dvi',        label: 'DVI (MRN)' },
  { key: 'dataMRN',    label: 'Data MRN' },
  { key: 'awb',        label: 'AWB' },
  { key: 'exportator', label: 'Exportator' },
  { key: 'taraExp',    label: 'Țara Exp.' },
  { key: 'moneda',     label: 'Moneda' },
  { key: 'valoare',    label: 'Valoare Marfă' },
  { key: 'awbLunaAn',  label: 'AWB - Luna An' },
  { key: 'nrFactura',  label: 'Nr. Factură' },
  { key: 'codTaric',   label: 'Cod TARIC' },
  { key: 'locatie',    label: 'Locație' },
  { key: 'gratis',     label: 'Gratis', type: 'select', options: ['NU', 'DA'] },
];

// ─── RENDER PREVIEW TABLE ───────────────────────────────────────────────────
function renderPreview(rows) {
  const section = document.getElementById('preview-section');
  const thead = document.querySelector('#preview-section thead tr');
  const tbody = document.getElementById('preview-tbody');

  // Build header dynamically
  thead.innerHTML = '<th>#</th><th>Fișier</th>' + COLUMNS.map(c => '<th>' + c.label + '</th>').join('');
  tbody.innerHTML = '';

  rows.forEach((row, idx) => {
    const f = row.fields || emptyFields();
    const tr = document.createElement('tr');
    let html = '<td>' + (idx + 1) + '</td><td>' + escHtml(row.fileName) + '</td>';

    COLUMNS.forEach(col => {
      const val = f[col.key] != null ? String(f[col.key]) : '';
      if (col.type === 'select') {
        const opts = col.options.map(o =>
          '<option value="' + o + '"' + (val === o ? ' selected' : '') + '>' + o + '</option>'
        ).join('');
        html += '<td><select class="editable" data-idx="' + idx + '" data-field="' + col.key + '">' + opts + '</select></td>';
      } else {
        html += '<td><input class="editable" type="text" value="' + escAttr(val) + '" data-idx="' + idx + '" data-field="' + col.key + '"></td>';
      }
    });

    tr.innerHTML = html;
    tbody.appendChild(tr);
  });

  // Bind editable events
  tbody.querySelectorAll('.editable').forEach(el => {
    // 'change' fires on blur/enter — safe for numeric parse
    el.addEventListener('change', function(e) {
      const idx = parseInt(e.target.dataset.idx);
      const field = e.target.dataset.field;
      var val = e.target.value;
      if (field === 'valoare') {
        var n = parseFloat(val.replace(',', '.'));
        if (!isNaN(n)) val = n;
      }
      fileRows[idx].fields[field] = val;
    });
    // 'input' fires on every keystroke — update text fields live (skip valoare to avoid NaN)
    if (el.tagName === 'INPUT') {
      el.addEventListener('input', function(e) {
        const idx = parseInt(e.target.dataset.idx);
        const field = e.target.dataset.field;
        if (field !== 'valoare') {
          fileRows[idx].fields[field] = e.target.value;
        }
      });
    }
  });

  section.style.display = 'block';
}

function escHtml(str) {
  return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escAttr(str) {
  return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ─── RENDER WARNINGS ────────────────────────────────────────────────────────
function renderWarnings(rows) {
  var container = document.getElementById('warnings-container');
  container.innerHTML = '';
  var withWarnings = rows.filter(function(r) { return r.warnings && r.warnings.length > 0; });
  if (withWarnings.length === 0) {
    container.innerHTML = '<p style="color:var(--success);font-size:.85rem;">✓ No warnings.</p>';
    return;
  }
  withWarnings.forEach(function(row) {
    var div = document.createElement('div');
    div.className = 'warning-block';
    div.innerHTML = '<div class="file-name">⚠ ' + escHtml(row.fileName) + '</div><ul>' +
      row.warnings.map(function(w) { return '<li>' + escHtml(w) + '</li>'; }).join('') + '</ul>';
    container.appendChild(div);
  });
}

// ─── EXPORT XLSX ────────────────────────────────────────────────────────────
function exportXlsx(rows) {
  if (!rows.length) return;

  // Visual feedback — disable button and show generating label
  var btn = document.getElementById('btn-export');
  var origLabel = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Se generează...'; }

  // Defer heavy XLSX work so the browser can repaint the button first
  setTimeout(function() {
    try {
      var headers = ['#', 'Fișier'];
      COLUMNS.forEach(function(c) { headers.push(c.label); });

      var data = rows.map(function(row, i) {
        var f = row.fields || emptyFields();
        var rowData = [i + 1, row.fileName];
        COLUMNS.forEach(function(c) {
          var v = f[c.key];
          rowData.push(v != null ? v : '');
        });
        return rowData;
      });

      var ws = XLSX.utils.aoa_to_sheet([headers].concat(data));
      ws['!cols'] = [
        { wch: 4 }, { wch: 35 }, { wch: 24 }, { wch: 12 }, { wch: 14 },
        { wch: 30 }, { wch: 8 }, { wch: 8 }, { wch: 14 },
        { wch: 22 }, { wch: 22 }, { wch: 14 }, { wch: 20 }, { wch: 8 }
      ];

      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Import DVI');
      XLSX.writeFile(wb, 'Extract_DVI_AWB.xlsx');
      console.log('[app.js] Excel exported successfully');
    } catch (err) {
      console.error('[app.js] Export error:', err);
      alert('Eroare la export Excel: ' + err.message);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = origLabel; }
    }
  }, 50);
}

// ─── CLEAR ──────────────────────────────────────────────────────────────────
function clearAll() {
  fileRows = [];
  document.getElementById('file-list').innerHTML = '';
  document.getElementById('preview-tbody').innerHTML = '';
  document.getElementById('preview-section').style.display = 'none';
  document.getElementById('warnings-container').innerHTML = '';
  document.getElementById('progress-wrap').classList.remove('visible');
  document.getElementById('progress-fill').style.width = '0%';
  document.getElementById('progress-text').textContent = '';
  document.getElementById('btn-export').disabled = true;
}

// ─── EXPORTS FOR TESTING ────────────────────────────────────────────────────
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { normalizeText, parseFields, emptyFields, patterns, ro,
    extractAWB, extractExportator, extractTaraExp, extractAwbLunaAn,
    extractFactura, extractLocatie, COLUMNS };
}
