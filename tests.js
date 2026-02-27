/* ==========================================================================
   tests.js – Automated tests for app.js extraction logic
   Run with:  node tests.js
   ========================================================================== */

// Minimal DOM stub so app.js can load without errors
const doc = {
  listeners: {},
  elements: {},
  createElement: function() { return { textContent: '', innerHTML: '', get innerHTML() { return this._ih || ''; }, set innerHTML(v) { this._ih = v; } }; },
  getElementById: function() { return null; },
  querySelector: function() { return null; },
  addEventListener: function(ev, fn) { doc.listeners[ev] = fn; },
};
global.document = doc;

// Load app.js
const {
  normalizeText, parseFields, emptyFields, patterns, ro,
  extractAWB, extractExportator, extractTaraExp, extractAwbLunaAn,
  extractFactura, extractLocatie, extractRegimUnificat, COLUMNS,
  extractPreferinte
} = require('./app.js');

// ─── Test infrastructure ────────────────────────────────────────────────────
let passed = 0, failed = 0, errors = [];
function assert(condition, msg) {
  if (condition) { passed++; }
  else { failed++; errors.push('FAIL: ' + msg); }
}
function assertEqual(actual, expected, msg) {
  if (actual === expected) { passed++; }
  else { failed++; errors.push('FAIL: ' + msg + ' | expected: "' + expected + '" | got: "' + actual + '"'); }
}

// ─── Test data: simulates text extracted from PDFs ──────────────────────────
const PDF1_TEXT = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Tip dec. IM Biroul vamal ROTM8730 BV Aeroport Timisoara
Numele HELLA ROMANIA SRL
Tip dec. sup A Punct vamal
Adresa - [13 04 018 000]
Categ. H1 Biroul vamal de sup. [17 10]
Strada și Numărul HELLA NR. 3
Total articole 1 Biroul vamal de prez. [17 09]
Orașul GHIRODA
Total colete 1 MRN 26ROTM87300008A1R5 17/02/2026
Codul poștal 307200
LRN 26RO1590678DHLVOZ947 17/02/2026
Țara RO
Dată liber vamă 17/02/2026

LISTA SEGMENTELOR DE TRANSPORT

SEGMENT TRANSPORT Nr. 1
Documentul precedent - [12 01]
1. NMNS / 125 / 108 / 16.02.2026
Document justificativ - [12 03]
1. N380 / 90022227 / / 2026-02-17 00:00:00.0 /
2. 1049 / TRADUCERE / / 2026-02-17 00:00:00.0 /
3. 1111 / 1612 / / 2026-02-17 00:00:00.0 /
4. N864 / DECL PREF EXP AUT / / /
Documentul de transport - [12 05]
1. N740 / 6646529444
Exportatorul - [13 01] Nr:
Numele PRECI-DIP SA
Adresa - [13 01 018 000]
Strada și Numărul RUE ST-HENRI 11
Orașul DELEMONT
Codul poștal 2800
Țara CH
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 8802.5
Țara de expediere - [16 06] CH

Cod TARIC unificat 8536693000
Preferințe - [14 11] 300

SEGMENT MARFURI
Nr. art. 1
Regimul solicitat - [11 09] 40
Regimul anterior - [11 09] 00
Regim unificat 4000`;

const PDF2_TEXT = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Tip dec. IM Biroul vamal ROTM8730 BV Aeroport Timisoara
Numele HELLA ROMANIA SRL
Tip dec. sup A Punct vamal
Adresa - [13 04 018 000]
Categ. H1 Biroul vamal de sup. [17 10]
Strada și Numărul HELLA NR. 3
Total articole 1 Biroul vamal de prez. [17 09]
Orașul GHIRODA
Total colete 1 MRN 26ROTM87300008BJR7 17/02/2026
Codul poștal 307200
LRN 26RO1592989FEDEX225Z 17/02/2026
Țara RO
Dată liber vamă 17/02/2026

LISTA SEGMENTELOR DE TRANSPORT

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. 1113 / DA/17.02.2026 / / /
2. 1049 / TRADUCERE/12.02.2026 / / /
3. N380 / CI020226969A/12.02.2026 / / /
4. 1111 / NR. 1612/30.12.2026 / / /
5. N325 / F.TR. 2026667/12.02.2026 / / /
Documentul de transport - [12 05]
1. N741 / 344501879
Exportatorul - [13 01] Nr:
Numele INDIUM CORPORATION EUROPEAN
Adresa - [13 01 018 000]
Strada și Numărul 7 NEWMARKET COURT
Orașul MILTON KEYNES
Codul poștal MK10 0AG
Țara GB
Moneda de facturare - [14 05] USD
Cuantumul total facturat - [14 06] 32205.6
Țara de expediere - [16 06] GB

Cod TARIC unificat 3810100000
Preferințe - [14 11] 300

SEGMENT MARFURI
Nr. art. 1
Regimul solicitat - [11 09] 40
Regimul anterior - [11 09] 00
Regim unificat 4000`;

// Text with cedilla variants (ş/ţ instead of ș/ț)
const PDF3_CEDILLA = `DECLARAŢIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Biroul vamal ROTM8730 BV Aeroport Timişoara
Oraşul TIMIŞOARA
MRN 26ROTM87300009XYZ3 25/03/2026
Dată liber vamă 25/03/2026

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. N380 / INV-2026-001 / / /
Documentul de transport - [12 05]
1. N740 / 9998887776
Exportatorul - [13 01] Nr:
Numele TEST COMPANY GMBH
Ţara DE
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 1234.56
Ţara de expediere - [16 06] DE

Cod TARIC unificat 8471300000
Preferințe - [14 11] 100

SEGMENT MARFURI
Nr. art. 1
Regimul solicitat - [11 09] 40
Regimul anterior - [11 09] 71
Regim unificat 4071`;

// ─── PDF4: N705 transport document ─────────────────────────────────────────
const PDF4_N705 = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Orașul GHIRODA
MRN 26ROTM87200009RYR3 26/02/2026
Dată liber vamă 26/02/2026

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. N380 / IKD-HELLA-ROM-S014/13.12.2025 / / /
Documentul de transport - [12 05]
1. N730 / CMR
2. N705 / NGB25SE14687
Exportatorul - [13 01] Nr:
Numele DAICEL SAFETY SYSTEMS EUROPE
Țara HU
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 55000.00
Țara de expediere - [16 06] HU

Cod TARIC unificat 8708950000
Preferințe - [14 11] 200

SEGMENT MARFURI
Nr. art. 1
Regimul solicitat - [11 09] 40
Regimul anterior - [11 09] 00
Regim unificat 4000`;

// ─── PDF5: N787 transport document (picking list) ───────────────────────────
const PDF5_N787 = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Orașul GHIRODA
MRN 26ROTM87200010ABC5 26/02/2026
Dată liber vamă 26/02/2026

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. N380 / FACT-2026-100 / / /
Documentul de transport - [12 05]
1. N787 / 296. Picking list Hella 26.02.2026 -AR
Exportatorul - [13 01] Nr:
Numele AUTOMOTIVE PARTS SRL
Țara DE
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 12500.00
Țara de expediere - [16 06] DE

Cod TARIC unificat 8708299000
Preferințe - [14 11] 100

SEGMENT MARFURI
Nr. art. 1
Regimul solicitat - [11 09] 40
Regimul anterior - [11 09] 71
Regim unificat 4071`;

// PDF6: Multi-invoice N380 with comma-separated values (DHL_import_1463107122 style)
const PDF6_MULTI_COMMA = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Orașul GHIRODA
MRN 26ROTM87300008ETR0 17/02/2026
Dată liber vamă 17/02/2026

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. N380 / NL26020303, -304 / / 2026-02-17 00:00:00.0 /
Documentul de transport - [12 05]
1. N740 / 1463107122
Exportatorul - [13 01] Nr:
Numele RAYBEN TECHNOLOGIES HK LIMITED
Țara HK
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 2834.91
Țara de expediere - [16 06] HK

Cod TARIC unificat 8517140000
Preferințe - [14 11] 100
Regim unificat 4000`;

// PDF7: Multi-invoice N380 with semicolon-separated values (user screenshot style)
const PDF7_MULTI_SEMICOLON = `DECLARAȚIE VAMALĂ DE IMPORT
SEGMENT GENERAL
Importatorul - [13 04] Nr: RO16297090
Orașul TIMISOARA
MRN 26ROTM87300009XYZ1 12/02/2026
Dată liber vamă 12/02/2026

SEGMENT TRANSPORT Nr. 1
Document justificativ - [12 03]
1. N380 / 20199698; 20200760/12.02.2026 / / /
Documentul de transport - [12 05]
1. N740 / 9988776655
Exportatorul - [13 01] Nr:
Numele EXAMPLE EXPORT CO
Țara CN
Moneda de facturare - [14 05] EUR
Cuantumul total facturat - [14 06] 5500.00
Țara de expediere - [16 06] CN

Cod TARIC unificat 3926909790
Preferințe - [14 11] 200
Regim unificat 4000`;

console.log('═══════════════════════════════════════════════');
console.log('  Running automated tests for app.js');
console.log('═══════════════════════════════════════════════\n');

// ─── TEST GROUP: normalizeText ──────────────────────────────────────────────
console.log('▸ normalizeText');
assertEqual(normalizeText('hello   world'), 'hello world', 'collapse spaces');
assertEqual(normalizeText('a\r\nb'), 'a\nb', 'CRLF to LF');
assertEqual(normalizeText('a\n\n\n\nb'), 'a\n\nb', 'max 2 newlines');
assertEqual(normalizeText(' \n hello \n '), '\nhello\n', 'trim spaces around newlines');

// ─── TEST GROUP: emptyFields ────────────────────────────────────────────────
console.log('▸ emptyFields');
const ef = emptyFields();
assertEqual(ef.gratis, 'NU', 'gratis default is NU');
assertEqual(ef.dvi, '', 'dvi default is empty');
assertEqual(ef.codTaric, '', 'codTaric default is empty');
assertEqual(ef.preferinte, '', 'preferinte default is empty');
assert(Object.keys(ef).length === 14, 'emptyFields has 14 keys, got ' + Object.keys(ef).length);
assert(ef.hasOwnProperty('regimUnificat'), 'emptyFields has regimUnificat');
assert(ef.hasOwnProperty('preferinte'), 'emptyFields has preferinte');

// ─── TEST GROUP: PDF1 parsing ───────────────────────────────────────────────
console.log('▸ PDF1 – DHL / PRECI-DIP');
const r1 = parseFields(PDF1_TEXT);
const f1 = r1.fields;

assertEqual(f1.dvi, '26ROTM87300008A1R5', 'PDF1 MRN');
assertEqual(f1.dataMRN, '17/02/2026', 'PDF1 Data MRN');
assertEqual(f1.awb, '6646529444', 'PDF1 AWB');
assertEqual(f1.exportator, 'PRECI-DIP SA', 'PDF1 Exportator');
assertEqual(f1.taraExp, 'CH', 'PDF1 Țara Exp');
assertEqual(f1.moneda, 'EUR', 'PDF1 Moneda');
assertEqual(f1.valoare, 8802.5, 'PDF1 Valoare');
assertEqual(f1.awbLunaAn, 'AWB - Februarie 2026', 'PDF1 AWB Luna An');
assertEqual(f1.nrFactura, '90022227', 'PDF1 Nr. Factură');
assertEqual(f1.codTaric, '8536693000', 'PDF1 Cod TARIC');
assertEqual(f1.regimUnificat, '4000', 'PDF1 Regim unificat');
assertEqual(f1.locatie, 'GHIRODA', 'PDF1 Locație');
assertEqual(f1.preferinte, '300', 'PDF1 Preferințe');
assertEqual(f1.gratis, 'NU', 'PDF1 Gratis');

// ─── TEST GROUP: PDF2 parsing ───────────────────────────────────────────────
console.log('▸ PDF2 – FedEx / INDIUM');
const r2 = parseFields(PDF2_TEXT);
const f2 = r2.fields;

assertEqual(f2.dvi, '26ROTM87300008BJR7', 'PDF2 MRN');
assertEqual(f2.dataMRN, '17/02/2026', 'PDF2 Data MRN');
assertEqual(f2.awb, '344501879', 'PDF2 AWB');
assertEqual(f2.exportator, 'INDIUM CORPORATION EUROPEAN', 'PDF2 Exportator');
assertEqual(f2.taraExp, 'GB', 'PDF2 Țara Exp');
assertEqual(f2.moneda, 'USD', 'PDF2 Moneda');
assertEqual(f2.valoare, 32205.6, 'PDF2 Valoare');
assertEqual(f2.awbLunaAn, 'AWB - Februarie 2026', 'PDF2 AWB Luna An');
assertEqual(f2.nrFactura, 'CI020226969A', 'PDF2 Nr. Factură (N380 priority)');
assertEqual(f2.codTaric, '3810100000', 'PDF2 Cod TARIC');
assertEqual(f2.regimUnificat, '4000', 'PDF2 Regim unificat');
assertEqual(f2.locatie, 'GHIRODA', 'PDF2 Locație');
assertEqual(f2.preferinte, '300', 'PDF2 Preferințe');
assertEqual(f2.gratis, 'NU', 'PDF2 Gratis');

// ─── TEST GROUP: Cedilla variants (ş/ţ) ────────────────────────────────────
console.log('▸ PDF3 – Cedilla variants (ş/ţ)');
const r3 = parseFields(PDF3_CEDILLA);
const f3 = r3.fields;

assertEqual(f3.dvi, '26ROTM87300009XYZ3', 'PDF3 MRN with cedilla');
assertEqual(f3.awb, '9998887776', 'PDF3 AWB');
assertEqual(f3.exportator, 'TEST COMPANY GMBH', 'PDF3 Exportator');
assertEqual(f3.taraExp, 'DE', 'PDF3 Țara Exp (cedilla Ţara)');
assertEqual(f3.moneda, 'EUR', 'PDF3 Moneda');
assertEqual(f3.valoare, 1234.56, 'PDF3 Valoare');
assertEqual(f3.nrFactura, 'INV-2026-001', 'PDF3 Factura');
assertEqual(f3.codTaric, '8471300000', 'PDF3 Cod TARIC');
assertEqual(f3.regimUnificat, '4071', 'PDF3 Regim unificat (cedilla text)');
assert(f3.locatie.length > 0, 'PDF3 Locație extracted (cedilla Oraşul): "' + f3.locatie + '"');
assertEqual(f3.dataMRN, '25/03/2026', 'PDF3 Data MRN');
assertEqual(f3.awbLunaAn, 'AWB - Martie 2026', 'PDF3 AWB Luna An');
assertEqual(f3.preferinte, '100', 'PDF3 Preferințe');

// ─── TEST GROUP: ro() helper ────────────────────────────────────────────────
console.log('▸ ro() helper');
const testRe1 = ro('Orașul\\s+(\\S+)');
assert(testRe1.test('Orașul GHIRODA'), 'ro() matches s-comma-below');
assert(testRe1.test('Oraşul GHIRODA'), 'ro() matches s-cedilla');
const testRe2 = ro('Țara\\s+([A-Z]{2})');
assert(testRe2.test('Țara CH'), 'ro() matches t-comma-below');
assert(testRe2.test('Ţara CH'), 'ro() matches t-cedilla');
// ro() uppercase variants
const testRe3 = ro('Ș');
assert(testRe3.test('Ș'), 'ro() uppercase S-comma');
assert(testRe3.test('Ş'), 'ro() uppercase S-cedilla');
const testRe4 = ro('Ț');
assert(testRe4.test('Ț'), 'ro() uppercase T-comma');
assert(testRe4.test('Ţ'), 'ro() uppercase T-cedilla');
// ro() passthrough – non-Romanian chars stay untouched
const testRe5 = ro('hello\\s+world');
assert(testRe5.test('hello  world'), 'ro() plain pattern works');
assert(!testRe5.test('helloworld'), 'ro() plain pattern rejects mismatch');

// ─── TEST GROUP: PDF4 – N705 transport ───────────────────────────────────────
console.log('▸ PDF4 – N705 transport');
const r4 = parseFields(PDF4_N705);
const f4 = r4.fields;

assertEqual(f4.dvi, '26ROTM87200009RYR3', 'PDF4 MRN');
assertEqual(f4.dataMRN, '26/02/2026', 'PDF4 Data MRN');
assertEqual(f4.awb, 'NGB25SE14687', 'PDF4 AWB from N705');
assertEqual(f4.exportator, 'DAICEL SAFETY SYSTEMS EUROPE', 'PDF4 Exportator');
assertEqual(f4.taraExp, 'HU', 'PDF4 Țara Exp');
assertEqual(f4.moneda, 'EUR', 'PDF4 Moneda');
assertEqual(f4.valoare, 55000, 'PDF4 Valoare');
assertEqual(f4.nrFactura, 'IKD-HELLA-ROM-S014', 'PDF4 Nr. Factură');
assertEqual(f4.codTaric, '8708950000', 'PDF4 Cod TARIC');
assertEqual(f4.regimUnificat, '4000', 'PDF4 Regim unificat');
assertEqual(f4.locatie, 'GHIRODA', 'PDF4 Locație');
assertEqual(f4.preferinte, '200', 'PDF4 Preferințe');
assertEqual(f4.gratis, 'NU', 'PDF4 Gratis');

// ─── TEST GROUP: PDF5 – N787 transport (picking list) ────────────────────────
console.log('▸ PDF5 – N787 transport (picking list)');
const r5 = parseFields(PDF5_N787);
const f5 = r5.fields;

assertEqual(f5.dvi, '26ROTM87200010ABC5', 'PDF5 MRN');
assertEqual(f5.dataMRN, '26/02/2026', 'PDF5 Data MRN');
assertEqual(f5.awb, '296. Picking list Hella 26.02.2026 -AR', 'PDF5 AWB from N787');
assertEqual(f5.exportator, 'AUTOMOTIVE PARTS SRL', 'PDF5 Exportator');
assertEqual(f5.taraExp, 'DE', 'PDF5 Țara Exp');
assertEqual(f5.moneda, 'EUR', 'PDF5 Moneda');
assertEqual(f5.valoare, 12500, 'PDF5 Valoare');
assertEqual(f5.nrFactura, 'FACT-2026-100', 'PDF5 Nr. Factură');
assertEqual(f5.codTaric, '8708299000', 'PDF5 Cod TARIC');
assertEqual(f5.regimUnificat, '4071', 'PDF5 Regim unificat 4071');
assertEqual(f5.locatie, 'GHIRODA', 'PDF5 Locație');
assertEqual(f5.preferinte, '100', 'PDF5 Preferințe');
assertEqual(f5.gratis, 'NU', 'PDF5 Gratis');

// ─── TEST GROUP: PDF6 – Multi-invoice comma-separated ────────────────────────
console.log('▸ PDF6 – Multi-invoice comma-separated');
const r6 = parseFields(PDF6_MULTI_COMMA);
const f6 = r6.fields;

assertEqual(f6.dvi, '26ROTM87300008ETR0', 'PDF6 MRN');
assertEqual(f6.dataMRN, '17/02/2026', 'PDF6 Data MRN');
assertEqual(f6.awb, '1463107122', 'PDF6 AWB');
assertEqual(f6.exportator, 'RAYBEN TECHNOLOGIES HK LIMITED', 'PDF6 Exportator');
assertEqual(f6.taraExp, 'HK', 'PDF6 Țara Exp');
assertEqual(f6.moneda, 'EUR', 'PDF6 Moneda');
assertEqual(f6.valoare, 2834.91, 'PDF6 Valoare');
assertEqual(f6.nrFactura, 'NL26020303, -304', 'PDF6 Nr. Factură (comma-separated)');
assertEqual(f6.codTaric, '8517140000', 'PDF6 Cod TARIC');
assertEqual(f6.regimUnificat, '4000', 'PDF6 Regim unificat');
assertEqual(f6.locatie, 'GHIRODA', 'PDF6 Locație');
assertEqual(f6.preferinte, '100', 'PDF6 Preferințe');
assertEqual(f6.gratis, 'NU', 'PDF6 Gratis');

// ─── TEST GROUP: PDF7 – Multi-invoice semicolon-separated ────────────────────
console.log('▸ PDF7 – Multi-invoice semicolon-separated');
const r7 = parseFields(PDF7_MULTI_SEMICOLON);
const f7 = r7.fields;

assertEqual(f7.dvi, '26ROTM87300009XYZ1', 'PDF7 MRN');
assertEqual(f7.dataMRN, '12/02/2026', 'PDF7 Data MRN');
assertEqual(f7.awb, '9988776655', 'PDF7 AWB');
assertEqual(f7.exportator, 'EXAMPLE EXPORT CO', 'PDF7 Exportator');
assertEqual(f7.taraExp, 'CN', 'PDF7 Țara Exp');
assertEqual(f7.moneda, 'EUR', 'PDF7 Moneda');
assertEqual(f7.valoare, 5500, 'PDF7 Valoare');
assertEqual(f7.nrFactura, '20199698; 20200760', 'PDF7 Nr. Factură (semicolon-separated)');
assertEqual(f7.codTaric, '3926909790', 'PDF7 Cod TARIC');
assertEqual(f7.regimUnificat, '4000', 'PDF7 Regim unificat');
assertEqual(f7.locatie, 'TIMISOARA', 'PDF7 Locație');
assertEqual(f7.preferinte, '200', 'PDF7 Preferințe');
assertEqual(f7.gratis, 'NU', 'PDF7 Gratis');

// ─── TEST GROUP: Individual extractors ──────────────────────────────────────
console.log('▸ extractAWB');
assertEqual(extractAWB(PDF1_TEXT), '6646529444', 'extractAWB PDF1 (N740)');
assertEqual(extractAWB(PDF2_TEXT), '344501879', 'extractAWB PDF2 (N741)');
assertEqual(extractAWB(PDF3_CEDILLA), '9998887776', 'extractAWB PDF3 (N740)');
assertEqual(extractAWB('no transport section here'), '', 'extractAWB empty on no match');
assertEqual(extractAWB(PDF4_N705), 'NGB25SE14687', 'extractAWB PDF4 (N705)');
assertEqual(extractAWB(PDF5_N787), '296. Picking list Hella 26.02.2026 -AR', 'extractAWB PDF5 (N787)');
// N730 CMR should be skipped if N705 is first match by position — in PDF4 N730 comes first but N705 second
// The function should pick the first one in document order (N730/CMR first)
// But in actual data N705 has the real AWB code, so let's verify with standalone test
const textN705Only = `SEGMENT TRANSPORT Nr. 1\nDocumentul de transport - [12 05]\n1. N705 / ABC123456\n`;
assertEqual(extractAWB(textN705Only), 'ABC123456', 'extractAWB standalone N705');
const textN730Only = `SEGMENT TRANSPORT Nr. 1\nDocumentul de transport - [12 05]\n1. N730 / CMR-SHIPMENT-99\n`;
assertEqual(extractAWB(textN730Only), 'CMR-SHIPMENT-99', 'extractAWB standalone N730');
// Regression: pdf.js merges columns so AWB line may have "Antrepozit - [12 11]" appended
const textAWBNoise = `SEGMENT TRANSPORT Nr. 1\nDocumentul de transport - [12 05]\n1. N740 / 6646529444 Antrepozit - [12 11]\n`;
assertEqual(extractAWB(textAWBNoise), '6646529444', 'extractAWB strips trailing Antrepozit label');
// Regression: "Tipul" appended after AWB value
const textAWBTipul = `SEGMENT TRANSPORT Nr. 1\nDocumentul de transport - [12 05]\n2. N705 / NGB25SE14687 Tipul\n`;
assertEqual(extractAWB(textAWBTipul), 'NGB25SE14687', 'extractAWB strips trailing Tipul');

console.log('▸ extractExportator');
assertEqual(extractExportator(PDF1_TEXT), 'PRECI-DIP SA', 'extractExportator PDF1');
assertEqual(extractExportator(PDF2_TEXT), 'INDIUM CORPORATION EUROPEAN', 'extractExportator PDF2');
assertEqual(extractExportator(PDF3_CEDILLA), 'TEST COMPANY GMBH', 'extractExportator PDF3');
assertEqual(extractExportator('no exporter here'), '', 'extractExportator empty on no match');
// Should not include trailing single char from next field
const textTrailingA = `SEGMENT TRANSPORT Nr. 1
Exportatorul - [13 01] Nr:
Numele BIG CORP A
Adresa - [13 01 018 000]`;
assertEqual(extractExportator(textTrailingA), 'BIG CORP', 'extractExportator strips trailing single char');

console.log('▸ extractTaraExp');
assertEqual(extractTaraExp(PDF1_TEXT), 'CH', 'extractTaraExp PDF1');
assertEqual(extractTaraExp(PDF2_TEXT), 'GB', 'extractTaraExp PDF2');
assertEqual(extractTaraExp(PDF3_CEDILLA), 'DE', 'extractTaraExp PDF3 cedilla');
assertEqual(extractTaraExp('no country here'), '', 'extractTaraExp empty on no match');

console.log('▸ extractAwbLunaAn');
assertEqual(extractAwbLunaAn(PDF1_TEXT), 'AWB - Februarie 2026', 'extractAwbLunaAn PDF1');
assertEqual(extractAwbLunaAn(PDF2_TEXT), 'AWB - Februarie 2026', 'extractAwbLunaAn PDF2');
assertEqual(extractAwbLunaAn(PDF3_CEDILLA), 'AWB - Martie 2026', 'extractAwbLunaAn PDF3');
assertEqual(extractAwbLunaAn('no dates here'), '', 'extractAwbLunaAn empty on no match');

console.log('▸ extractFactura');
assertEqual(extractFactura(PDF1_TEXT), '90022227', 'extractFactura PDF1 (N380)');
assertEqual(extractFactura(PDF2_TEXT), 'CI020226969A', 'extractFactura PDF2 (N380 before N325)');
assertEqual(extractFactura(PDF3_CEDILLA), 'INV-2026-001', 'extractFactura PDF3');
assertEqual(extractFactura(PDF6_MULTI_COMMA), 'NL26020303, -304', 'extractFactura PDF6 (comma-separated)');
assertEqual(extractFactura(PDF7_MULTI_SEMICOLON), '20199698; 20200760', 'extractFactura PDF7 (semicolon-separated)');
assertEqual(extractFactura('no invoices here'), '', 'extractFactura empty on no match');

console.log('▸ extractLocatie');
assertEqual(extractLocatie(PDF1_TEXT), 'GHIRODA', 'extractLocatie PDF1');
assertEqual(extractLocatie(PDF2_TEXT), 'GHIRODA', 'extractLocatie PDF2');
assertEqual(extractLocatie(PDF3_CEDILLA), 'TIMIŞOARA', 'extractLocatie PDF3 cedilla');
assertEqual(extractLocatie('no importator section'), '', 'extractLocatie empty on no match');
// City name with multiple words
const textMultiWordCity = `Importatorul - [13 04] Nr: RO123
Orașul BAIA MARE
Codul poștal 430000`;
assertEqual(extractLocatie(textMultiWordCity), 'BAIA MARE', 'extractLocatie multi-word city');

console.log('▸ extractRegimUnificat');
assertEqual(extractRegimUnificat(PDF1_TEXT), '4000', 'extractRegimUnificat PDF1');
assertEqual(extractRegimUnificat(PDF2_TEXT), '4000', 'extractRegimUnificat PDF2');
assertEqual(extractRegimUnificat(PDF3_CEDILLA), '4071', 'extractRegimUnificat PDF3');
assertEqual(extractRegimUnificat(PDF4_N705), '4000', 'extractRegimUnificat PDF4');
assertEqual(extractRegimUnificat(PDF5_N787), '4071', 'extractRegimUnificat PDF5');
assertEqual(extractRegimUnificat('no regim here'), '', 'extractRegimUnificat empty on no match');
assertEqual(extractRegimUnificat('Regim unificat 9999'), '9999', 'extractRegimUnificat generic 4-digit');

// ─── TEST GROUP: Warnings correctness ───────────────────────────────────────
console.log('▸ Warnings');
// Full PDFs should have zero or minimal warnings
assert(r1.warnings.length === 0, 'PDF1 has no warnings, got: ' + JSON.stringify(r1.warnings));
assert(r2.warnings.length === 0, 'PDF2 has no warnings, got: ' + JSON.stringify(r2.warnings));
// Empty text should warn about all fields
const emptyWarnings = parseFields('').warnings;
assert(emptyWarnings.length >= 10, 'Empty text produces ≥10 warnings, got ' + emptyWarnings.length);
assert(emptyWarnings.some(w => /MRN/i.test(w)), 'Empty text warns about MRN');
assert(emptyWarnings.some(w => /AWB/i.test(w)), 'Empty text warns about AWB');
assert(emptyWarnings.some(w => /Exportator/i.test(w)), 'Empty text warns about Exportator');
assert(emptyWarnings.some(w => /Regim/i.test(w)), 'Empty text warns about Regim unificat');

// ─── TEST GROUP: normalizeText edge cases ───────────────────────────────────
console.log('▸ normalizeText edge cases');
assertEqual(normalizeText(''), '', 'normalizeText empty string');
assertEqual(normalizeText('  '), ' ', 'normalizeText only spaces');
assertEqual(normalizeText('a\tb'), 'a b', 'normalizeText tab to space');
assertEqual(normalizeText('line1\n \n \n \nline2'), 'line1\n\nline2', 'normalizeText multi blank lines');

// ─── TEST GROUP: Edge cases ─────────────────────────────────────────────────
console.log('▸ Edge cases');

// Empty text
const rEmpty = parseFields('');
assert(rEmpty.warnings.length > 0, 'Empty text produces warnings');
assertEqual(rEmpty.fields.gratis, 'NU', 'Empty text gratis is NU');

// Text with only MRN
const rPartial = parseFields('MRN 26ROTM87300008A1R5 17/02/2026');
assertEqual(rPartial.fields.dvi, '26ROTM87300008A1R5', 'Partial text extracts MRN');
assertEqual(rPartial.fields.dataMRN, '17/02/2026', 'Partial text extracts Data MRN');
assert(rPartial.warnings.length > 0, 'Partial text has warnings for missing fields');

// Valoare with comma
const rComma = parseFields('Cuantumul total facturat - [14 06] 1234,56\nMRN 26ROTM87300008A1R5 01/01/2026');
assertEqual(rComma.fields.valoare, 1234.56, 'Valoare with comma is parsed');

// AWB with N741 prefix
const rN741 = parseFields('SEGMENT TRANSPORT Nr. 1\nDocumentul de transport - [12 05]\n1. N741 / 555666777\n');
assertEqual(rN741.fields.awb, '555666777', 'AWB with N741 prefix');

// Large valoare
const rBigVal = parseFields('Cuantumul total facturat - [14 06] 999999.99\nMRN 26ROTM87300008A1R5 01/01/2026');
assertEqual(rBigVal.fields.valoare, 999999.99, 'Large valoare');

// MRN variant format
const rMRN2 = parseFields('MRN 26ROTM87300099ABCD 05/12/2025');
assertEqual(rMRN2.fields.dvi, '26ROTM87300099ABCD', 'MRN different format');
assertEqual(rMRN2.fields.dataMRN, '05/12/2025', 'Data MRN different format');
assertEqual(rMRN2.fields.awbLunaAn, 'AWB - Decembrie 2025', 'AWB Luna An from MRN date');

// Multiple months
const rJan = parseFields('Dată liber vamă 15/01/2026\nMRN 26ROTM87300008A1R5 15/01/2026');
assertEqual(rJan.fields.awbLunaAn, 'AWB - Ianuarie 2026', 'January extraction');
const rDec = parseFields('Dată liber vamă 25/12/2025\nMRN 26ROTM87300008A1R5 25/12/2025');
assertEqual(rDec.fields.awbLunaAn, 'AWB - Decembrie 2025', 'December extraction');

// ─── TEST GROUP: COLUMNS config ─────────────────────────────────────────────
console.log('▸ COLUMNS config');
assert(COLUMNS.length === 14, 'COLUMNS has 14 entries, got ' + COLUMNS.length);
assertEqual(COLUMNS[COLUMNS.length - 1].key, 'gratis', 'Last column is gratis');
assertEqual(COLUMNS[COLUMNS.length - 1].type, 'select', 'Gratis column is select type');
assert(COLUMNS.find(c => c.key === 'codTaric'), 'COLUMNS contains codTaric');
assert(COLUMNS.find(c => c.key === 'dataMRN'), 'COLUMNS contains dataMRN');
assert(COLUMNS.find(c => c.key === 'locatie'), 'COLUMNS contains locatie');
assert(COLUMNS.find(c => c.key === 'regimUnificat'), 'COLUMNS contains regimUnificat');
assert(COLUMNS.find(c => c.key === 'preferinte'), 'COLUMNS contains preferinte');
// Verify column order matches emptyFields keys
const colKeys = COLUMNS.map(c => c.key);
assert(colKeys.indexOf('dvi') < colKeys.indexOf('dataMRN'), 'dvi before dataMRN');
assert(colKeys.indexOf('dataMRN') < colKeys.indexOf('awb'), 'dataMRN before awb');
assert(colKeys.indexOf('codTaric') < colKeys.indexOf('locatie'), 'codTaric before locatie');
assert(colKeys.indexOf('locatie') < colKeys.indexOf('gratis'), 'locatie before gratis');
// Every COLUMN key must exist in emptyFields
const efKeys = Object.keys(emptyFields());
COLUMNS.forEach(function(c) {
  assert(efKeys.includes(c.key), 'emptyFields contains key: ' + c.key);
});

// ─── TEST GROUP: Cross-field consistency ────────────────────────────────────
console.log('▸ Cross-field consistency');
// PDF1: data MRN should match the date from AWB Luna An month
assertEqual(f1.dataMRN.split('/')[1], '02', 'PDF1 MRN month matches Februarie');
assertEqual(f1.awbLunaAn, 'AWB - Februarie 2026', 'PDF1 AWB Luna An consistent with Data MRN');
// PDF2: same checks
assertEqual(f2.dataMRN.split('/')[1], '02', 'PDF2 MRN month matches Februarie');
assertEqual(f2.awbLunaAn, 'AWB - Februarie 2026', 'PDF2 AWB Luna An consistent with Data MRN');
// All PDFs should have gratis = NU by default
assertEqual(f1.gratis, 'NU', 'PDF1 gratis default');
assertEqual(f2.gratis, 'NU', 'PDF2 gratis default');
assertEqual(f3.gratis, 'NU', 'PDF3 gratis default');

// ─── RESULTS ────────────────────────────────────────────────────────────────
console.log('\n═══════════════════════════════════════════════');
if (failed === 0) {
  console.log('  ✅ ALL ' + passed + ' TESTS PASSED');
} else {
  console.log('  ❌ ' + failed + ' FAILED / ' + passed + ' passed');
  errors.forEach(e => console.log('  ' + e));
}
console.log('═══════════════════════════════════════════════');
process.exit(failed > 0 ? 1 : 0);
