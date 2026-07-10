/* =====================================================================
   1. ZIP READER (xlsx/xlsm = zip). Native DecompressionStream untuk
      inflate (deflate-raw) -> tidak butuh library eksternal apapun,
      sehingga file ini tetap 100% portable & bisa jalan offline.
   ===================================================================== */
function readZipEntries(buf){
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for(let i = buf.length - 22; i >= 0; i--){
    if(view.getUint32(i, true) === 0x06054b50){ eocd = i; break; }
  }
  if(eocd === -1) throw new Error('Bukan file zip (xlsx/xlsm) yang valid.');
  const cdOffset = view.getUint32(eocd + 16, true);
  const cdCount = view.getUint16(eocd + 10, true);
  let p = cdOffset;
  const entries = [];
  for(let i = 0; i < cdCount; i++){
    const sig = view.getUint32(p, true);
    if(sig !== 0x02014b50) throw new Error('Struktur zip tidak terbaca.');
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localHeaderOffset = view.getUint32(p + 42, true);
    const name = new TextDecoder('utf-8').decode(buf.slice(p + 46, p + 46 + nameLen));
    entries.push({ name, method, compSize, localHeaderOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

async function extractEntry(buf, entry){
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const p = entry.localHeaderOffset;
  const nameLen = view.getUint16(p + 26, true);
  const extraLen = view.getUint16(p + 28, true);
  const dataStart = p + 30 + nameLen + extraLen;
  const compData = buf.slice(dataStart, dataStart + entry.compSize);
  if(entry.method === 0) return compData;
  if(entry.method === 8){
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    writer.write(compData);
    writer.close();
    const chunks = [];
    const reader = ds.readable.getReader();
    while(true){
      const { done, value } = await reader.read();
      if(done) break;
      chunks.push(value);
    }
    let total = 0; chunks.forEach(c => total += c.length);
    const out = new Uint8Array(total);
    let off = 0; chunks.forEach(c => { out.set(c, off); off += c.length; });
    return out;
  }
  throw new Error('Metode kompresi zip tidak didukung (' + entry.method + ').');
}

async function loadZipText(buf, entries, name){
  const map = {};
  entries.forEach(e => map[e.name] = e);
  if(!map[name]) return null;
  const bytes = await extractEntry(buf, map[name]);
  return new TextDecoder('utf-8').decode(bytes);
}

/* =====================================================================
   2. XLSX PARSING (pakai DOMParser bawaan browser)
   ===================================================================== */
function parseXml(text){
  return new DOMParser().parseFromString(text, 'application/xml');
}

function parseSharedStrings(doc){
  if(!doc) return [];
  const sis = doc.getElementsByTagName('si');
  const out = [];
  for(const si of sis) out.push(si.textContent);
  return out;
}

function getCellValue(cEl, sharedStrings){
  const t = cEl.getAttribute('t');
  if(t === 's'){
    const v = cEl.getElementsByTagName('v')[0];
    if(!v) return '';
    const idx = parseInt(v.textContent, 10);
    return sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
  }
  if(t === 'inlineStr'){
    const is = cEl.getElementsByTagName('is')[0];
    return is ? is.textContent : '';
  }
  if(t === 'b'){
    const v = cEl.getElementsByTagName('v')[0];
    return v && v.textContent === '1' ? 'TRUE' : 'FALSE';
  }
  const v = cEl.getElementsByTagName('v')[0];
  return v ? v.textContent : '';
}

function parseSheetToRows(doc, sharedStrings){
  const data = {};
  let lastRow = 0;
  const dim = doc.getElementsByTagName('dimension')[0];
  if(dim){
    const ref = dim.getAttribute('ref') || '';
    const m = ref.match(/:[A-Z]+(\d+)/);
    if(m) lastRow = parseInt(m[1], 10);
  }
  const rowEls = doc.getElementsByTagName('row');
  for(const rowEl of rowEls){
    const rNum = parseInt(rowEl.getAttribute('r'), 10);
    if(rNum > lastRow) lastRow = rNum;
    const rowData = {};
    const cellEls = rowEl.getElementsByTagName('c');
    for(const cEl of cellEls){
      const ref = cEl.getAttribute('r') || '';
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if(!m) continue;
      rowData[m[1]] = getCellValue(cEl, sharedStrings);
    }
    data[rNum] = rowData;
  }
  return { data, lastRow };
}

function cellStr(data, row, col){
  const r = data[row];
  if(!r) return '';
  const v = r[col];
  return (v === undefined || v === null) ? '' : String(v);
}

async function extractContractLines(arrayBuffer){
  const buf = new Uint8Array(arrayBuffer);
  const entries = readZipEntries(buf);

  const wbXmlText = await loadZipText(buf, entries, 'xl/workbook.xml');
  if(!wbXmlText) throw new Error('xl/workbook.xml tidak ditemukan di dalam file.');
  const wbDoc = parseXml(wbXmlText);
  const sheetEls = wbDoc.getElementsByTagName('sheet');
  let targetRid = null;
  for(const s of sheetEls){
    if(s.getAttribute('name') === 'Contract Lines'){
      targetRid = s.getAttribute('r:id');
      break;
    }
  }
  if(!targetRid) throw new Error('Sheet "Contract Lines" tidak ditemukan di dalam workbook.');

  const relsText = await loadZipText(buf, entries, 'xl/_rels/workbook.xml.rels');
  const relsDoc = parseXml(relsText);
  const relEls = relsDoc.getElementsByTagName('Relationship');
  let target = null;
  for(const r of relEls){
    if(r.getAttribute('Id') === targetRid){ target = r.getAttribute('Target'); break; }
  }
  if(!target) throw new Error('Relasi sheet "Contract Lines" tidak ditemukan.');
  const sheetPath = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\.?\//, '');

  const sstText = await loadZipText(buf, entries, 'xl/sharedStrings.xml');
  const sharedStrings = sstText ? parseSharedStrings(parseXml(sstText)) : [];

  const sheetText = await loadZipText(buf, entries, sheetPath);
  if(!sheetText) throw new Error('File sheet "Contract Lines" (' + sheetPath + ') tidak ditemukan.');
  const sheetDoc = parseXml(sheetText);

  return parseSheetToRows(sheetDoc, sharedStrings);
}

/* =====================================================================
   3. HELPER ANGKA & TANGGAL (mengikuti locale Indonesia: titik = pemisah
      ribuan, koma = desimal -- sesuai cara macro VBA aslinya membaca
      angka pada workbook ini)
   ===================================================================== */
function toNumberID(raw){
  if(raw === null || raw === undefined) return NaN;
  let s = String(raw).trim();
  if(s === '') return NaN;
  if(/^-?\d+$/.test(s)) return parseFloat(s);
  if(/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) return parseFloat(s.replace(/\./g,'').replace(',', '.'));
  if(/^-?\d+,\d+$/.test(s)) return parseFloat(s.replace(',', '.'));
  if(/^-?\d+\.\d+$/.test(s)) return parseFloat(s);
  return NaN;
}
function isNumericID(raw){ return !isNaN(toNumberID(raw)); }
function isZeroID(raw){ return toNumberID(raw) === 0; }

function parseDatePositional(s){
  if(!s || s.length !== 10) return null;
  const day = parseInt(s.substr(0,2), 10);
  const month = parseInt(s.substr(3,2), 10);
  const year = parseInt(s.substr(6,4), 10);
  if(isNaN(day) || isNaN(month) || isNaN(year)) return null;
  return year * 10000 + month * 100 + day;
}

/* =====================================================================
   4. VALIDATION ENGINE -- porting 1:1 dari Module1.bas (CMIIWmodule)
   ===================================================================== */
function runValidation(data, lastRow){
  const logs = [];
  const add = (errorId, kolom, type, msg, line) => logs.push({ errorId, kolom, line, type, msg });
  const cell = (row, col) => cellStr(data, row, col);
  const trim = (s) => s.trim();

  /* ---- VALIDASI GLOBAL ---- */
  // Point.1 - Validate_Sequence B & F
  for(let i = 2; i <= lastRow; i++){
    if(toNumberID(cell(i,'B')) !== i - 1) add(1, 'B','Error','Isian kolom B harus berurutan (1, 2, 3, dst)', i);
  }
  for(let i = 2; i <= lastRow; i++){
    if(toNumberID(cell(i,'F')) !== i - 1) add(2, 'F','Error','Isian kolom F harus berurutan (1, 2, 3, dst)', i);
  }

  // Point.23a - Validate_AI_Material
  {
    let isAllMaterial = true, isAllAJEmpty = true;
    for(let i = 2; i <= lastRow; i++){
      if(cell(i,'C') !== 'Material') isAllMaterial = false;
      if(trim(cell(i,'AJ')) !== '') isAllAJEmpty = false;
    }
    if(isAllMaterial && isAllAJEmpty){
      for(let i = 2; i <= lastRow; i++){
        if(toNumberID(cell(i,'AI')) !== i - 1) add(3, 'AI','Error','Isian pada kolom AI harus berurutan (1, 2, 3, dst)', i);
      }
    }
  }

  // Point.23b - Validate_AI_Service
  {
    let isAllService = true, isAllAJEmpty = true;
    for(let i = 2; i <= lastRow; i++){
      if(!cell(i,'C').includes('Service')) isAllService = false;
      if(trim(cell(i,'AJ')) !== '') isAllAJEmpty = false;
    }
    if(isAllService && isAllAJEmpty){
      let expected = 100010001;
      for(let i = 2; i <= lastRow; i++){
        if(toNumberID(cell(i,'AI')) !== expected) add(4, 'AI','Error','Isian pada kolom AI harus berurutan (100010001, 100020001, 100030001, dst)', i);
        expected += 10000;
      }
    }
  }

  // Point 3b - Validate_Service_E_Warning
  {
    let isAllService = true, isAllEEmpty = true, isAllAJFilled = true;
    for(let i = 2; i <= lastRow; i++){
      const valC = cell(i,'C').toLowerCase().trim();
      const valE = trim(cell(i,'E'));
      const valAJ = trim(cell(i,'AJ'));
      if(!valC.includes('service')) isAllService = false;
      if(valE !== '') isAllEEmpty = false;
      if(valAJ === '') isAllAJFilled = false;
    }
    if(isAllService && isAllEEmpty && isAllAJFilled){
      for(let i = 2; i <= lastRow; i++){
        add(5, 'E & AI','Warning','Pastikan Kolom E dan AI mengikuti PR SAP(Jika menggunakan Service Master atau Service Number, maka kolom E Tidak boleh Kosong)', i);
      }
    }
  }

  // Point.12 - Validate_AJ_Consistency
  {
    let hasEmpty = false, hasValue = false;
    for(let i = 2; i <= lastRow; i++){
      if(trim(cell(i,'AJ')) === '') hasEmpty = true; else hasValue = true;
    }
    if(hasEmpty && hasValue){
      for(let i = 2; i <= lastRow; i++){
        add(6, 'AJ','Error','Isian pada kolom AJ tidak boleh campur (ada PR dan tidak ada PR, harus ada PR semua atau harus tidak ada PR semua)', i);
      }
    }
  }

  /* ---- VALIDASI PER BARIS ---- */
  for(let i = 2; i <= lastRow; i++){
    const valC = trim(cell(i,'C'));

    // Point.13
    const valH = cell(i,'H');
    if(valH.length >= 1000 || valH.slice(-2) === '||'){
      add(7, 'H','Error','Isian kolom H untuk Item Text PO dan OA tidak boleh lebih dari 1000 karakter atau terdapat Double Piping/|| pada karakter terakhir', i);
    }

    // Point NEW - G max 40 karakter
    if(cell(i,'G').length >= 41){
      add(8, 'G','Error','Isian kolom G untuk Short Name tidak boleh lebih dari 40 karakter', i);
    }

    // Point NEW - AH tidak boleh ada tanda baca titik
    /* ---- if(cell(i,'AH').includes('.')){
      add(9, 'AH','Error','Isian pada kolom AH tidak boleh ada tanda baca titik', i);
    } ---- */

    // Point NEW - Jika A = "Create", maka I harus dikosongkan
    if(trim(cell(i,'A')) === 'Create' && trim(cell(i,'I')) !== ''){
      add(10, 'I','Error','Jika kolom A Create, maka kolom I harus dikosongkan', i);
    }

    // Point.8d
    if(!isNumericID(cell(i,'O')) || isZeroID(cell(i,'O'))){
      add(11, 'O','Error','Kolom O tidak boleh kosong atau 0', i);
    }
    // Point.8a
    if(isNumericID(cell(i,'O')) && toNumberID(cell(i,'O')) > 99999999999){
      add(12, 'O','Error','Kolom O tidak boleh lebih besar 99.999.999.999', i);
    }

    // Point.14 - kolom yang harus dikosongkan
    ['Q','R','S','T','U','V','W','X','Y','Z','AA','AE','AF','AG'].forEach(col => {
      if(trim(cell(i,col)) !== ''){
        add(13, col,'Error','Untuk kolom ' + col + ' tidak perlu diisi (dikosongkan saja)', i);
      }
    });

    // Point NEW - Mandatory checks
    if(trim(cell(i,'D')) === '') add(14, 'D','Error','Untuk kolom D Mandatory (harus diisi 100 / 160 / 180)', i);
    if(trim(cell(i,'B')) === '') add(15, 'B','Error','Untuk kolom B wajib diisi (Mandatory)', i);
    if(!['Fixed Service','Variable Service','Material'].includes(cell(i,'C'))){
      add(16, 'C','Error','Untuk kolom C Mandatory (biasanya terisi Fixed Service / Variable Service / Material)', i);
    }
    if(trim(cell(i,'F')) === '') add(17, 'F','Error','Untuk kolom F wajib diisi (Mandatory)', i);
    if(trim(cell(i,'G')) === '') add(18, 'G','Error','Untuk kolom G wajib diisi (Mandatory)', i);
    if(trim(cell(i,'H')) === '') add(19, 'H','Error','Untuk kolom H wajib diisi (Mandatory)', i);
    if(trim(cell(i,'K')) === '') add(20, 'K','Error','Untuk kolom K wajib diisi (Mandatory)', i);
    if(trim(cell(i,'N')) === '') add(21, 'N','Error','Untuk kolom N wajib diisi (Mandatory)', i);
    if(trim(cell(i,'P')) === '') add(22, 'P','Error','Untuk kolom P wajib diisi (Mandatory)', i);
    if(trim(cell(i,'AC')) === '') add(23, 'AC','Error','Untuk kolom AC wajib diisi (Mandatory)', i);
    if(trim(cell(i,'AD')) === '') add(24, 'AD','Error','Untuk kolom AD wajib diisi (Mandatory)', i);
    if(trim(cell(i,'AI')) === '') add(25, 'AI','Error','Untuk kolom AI wajib diisi (Mandatory)', i);

    // Point.10
    const valAI = cell(i,'AI');
    if(valAI.length > 0 && valAI.charAt(0) === '0'){
      add(26, 'AI','Error','Isian pada kolom AI tidak boleh ada angka 0 di depannya', i);
    }

    // Point.11
    if(cell(i,'AJ').includes('-')){
      add(27, 'AJ','Error','Isian pada kolom AJ tidak boleh ada 100- / 160- / 180-', i);
    }

    // Point NEW - spasi di kolom M
    const valM = cell(i,'M');
    if(trim(valM) !== '' && valM.includes(' ')){
      add(28, 'M','Error','Isian kolom M tidak boleh ada Spasi', i);
    }

    // ===== MATERIAL =====
    if(valC === 'Material'){
      const valE = cell(i,'E');
      const lenE = valE.length;
      const valD = cell(i,'D');

      // Point.5d
      if((lenE === 12 || lenE === 14) && trim(cell(i,'J')) === ''){
        add(29, 'J','Error','Pastikan Kolom J mengikuti Material Group yang ada di PR SAP (dapat berdiskusi dengan Tim Logistic Cataloguer SHU atau Regional terkait)', i);
      }
      // Point.5a
      if((valD === '100' || valD === '160') && lenE === 10 && trim(cell(i,'J')) === ''){
        add(30, 'J','Error','Isian kolom J  harus mengikuti Material Group di PR SAP (Rumusnya Digit Pertama dari Kolom E)', i);
      }
      // Point.5b
      if(valD === '180' && lenE === 10 && trim(cell(i,'J')) === ''){
        add(31, 'J','Error','Isian kolom J harus mengikuti Material Group di PR SAP (Rumusnya K0 + dua Digit Pertama dari Kolom E)', i);
      }
      // Point.3a
      if(trim(valE) === '' || isZeroID(valE)){
        add(32, 'E','Error','Isian kolom E tidak boleh kosong atau blank (harus mengikuti Kimap pada PR SAP)', i);
      }
      // Point.4a
      if(trim(cell(i,'M')) === '' || isZeroID(cell(i,'M'))){
        add(33, 'M','Error','Isian kolom M tidak boleh 0 atau blank (biasanya diisi angka 30)', i);
      }
      // Point.9c
      if(trim(cell(i,'AH')) === ''){
        add(34, 'AH','Error','Isian pada kolom AH tidak boleh kosong', i);
      }
    }

    // ===== SERVICE =====
    if(valC === 'Fixed Service' || valC === 'Variable Service'){
      // Point.5c
      if(trim(cell(i,'J')) === ''){
        add(35, 'J','Error','Isian kolom J harus mengikuti Material Group yang ada di PR SAP (Biasanya diisi dengan S99)', i);
      }
      if(trim(cell(i,'J')) !== 'S99' && trim(cell(i,'J')) !== ''){
        add(36, 'J','Warning','Pastikan kolom J mengikuti Material Group yang ada di PR SAP (Biasanya diisi dengan S99)', i);
      }
      // Point.4b
      if(trim(cell(i,'M')) === ''){
        add(37, 'M','Error','Isian pada kolom M tidak boleh kosong atau blank (biasanya diisi angka 0)', i);
      }
      // Point.9b
      if(trim(cell(i,'AH')) === '' || isZeroID(cell(i,'AH'))){
        add(38, 'AH','Error','Isian pada kolom AH tidak boleh 0 atau kosong', i);
      }
      // Point.2b
      if(trim(cell(i,'AI')) !== '' && cell(i,'AI').length !== 9){
        add(39, 'AI','Error','Isian pada kolom AI harus 9 Digit', i);
      }
    }

    // ===== ERROR/WARNING L =====
    const valD2 = cell(i,'D');
    // Point.6a
    if(valD2 === '100' || valD2 === '160'){
      if(cell(i,'L') === ''){
        add(40, 'L','Error','Kolom L tidak boleh kosong. Untuk SAP Client 100 atau 160, isian hanya mencakup V0 / SA / YA / YC', i);
      }
    }
    // Point.6b
    if(valD2 === '180'){
      if(cell(i,'L') === ''){
        add(41, 'L','Error','Kolom L tidak boleh kosong. Untuk SAP Client 180, isian hanya mencakup V0 / TB / TI / TH / YA', i);
      }
    }
    // Point.6a warning
    if(valD2 === '100' || valD2 === '160'){
      if(!['V0','SA','YA','YC'].includes(cell(i,'L')) && cell(i,'L') !== ''){
        add(42, 'L','Warning','Isian kolom L, untuk SAP Client 100 atau 160 hanya mencakup V0 / SA / YA / YC', i);
      }
    }
    // Point.6b warning
    if(valD2 === '180'){
      if(!['V0','TB','TI','TH','YA'].includes(cell(i,'L')) && cell(i,'L') !== ''){
        add(43, 'L','Warning','Isian kolom L, Untuk SAP Client 180 hanya mencakup V0 / TB / TI / TH / YA', i);
      }
    }

    // ===== CURRENCY =====
    if(isNumericID(cell(i,'O'))){
      const valO = toNumberID(cell(i,'O'));
      // Point.7a
      if(cell(i,'N') === 'IDR'){
        if(Math.abs(valO - Math.round(valO)) > 1e-9){
          add(44, 'O','Error','Isian pada kolom O tidak boleh Desimal', i);
        }
      }
      // Point.7b
      if(cell(i,'N') === 'USD'){
        const rounded = Math.round(valO * 100) / 100;
        if(Math.abs(valO - rounded) > 1e-9){
          add(45, 'O','Error','Jumlah digit desimal pada kolom O tidak boleh lebih dari 2 digit', i);
        }
      }
    }

    // Point.8b
    if(isNumericID(cell(i,'O')) && isNumericID(cell(i,'AH'))){
      if(toNumberID(cell(i,'O')) * toNumberID(cell(i,'AH')) > 99999999999){
        add(46, 'O dan AH','Error','Hasil perkalian antara Kolom O dengan AH tidak boleh lebih dari 99.999.999.999', i);
      }
    }

    // Point NEW - AI vs AJ tertukar
    if(trim(cell(i,'AJ')) !== ''){
      if(isNumericID(cell(i,'AI')) && isNumericID(cell(i,'AJ'))){
        if(toNumberID(cell(i,'AI')) > toNumberID(cell(i,'AJ'))){
          add(47, 'AI dan AJ','Error','Value kolom AI dan AJ tertukar', i);
        }
      }
    }

    // Point NEW - AC vs AD tertukar
    if(trim(cell(i,'AC')) !== '' && trim(cell(i,'AD')) !== ''){
      const dAC = parseDatePositional(cell(i,'AC'));
      const dAD = parseDatePositional(cell(i,'AD'));
      if(dAC !== null && dAD !== null && dAC > dAD){
        add(48, 'AC dan AD','Error','Value kolom AC dan AD tertukar', i);
      }
	  else if(dAD !== null && dAC !== null && dAD > dAC){
        add(53, 'AC dan AD','Warning','Pastikan Start Date dan End Date dalam range Effective Date dan Expiry Date di Terms Smart GEP', i);
      }
    }
  }

  // Point.25 - Validate_Column_C_Service_Order
  {
    let hasMaterial = false, hasService = false;
    for(let i = 2; i <= lastRow; i++){
      const v = cell(i,'C').toLowerCase().trim();
      if(v === 'material') hasMaterial = true;
      if(v === 'fixed service' || v === 'variable service') hasService = true;
    }
    if(hasMaterial && hasService){
      const v2 = cell(2,'C').toLowerCase().trim();
      if(!(v2 === 'fixed service' || v2 === 'variable service')){
        add(49, 'C','Error','Untuk OA ZKC3 (Material & Service), Line Pertama Kolom C harus diisi dengan isian Service', 2);
      }
    }
  }

  // Point NEW - Validate_Column_D
  {
    const distinct = new Set();
    for(let i = 2; i <= lastRow; i++){
      const v = trim(cell(i,'D'));
      if(v !== '') distinct.add(v);
    }
    if(distinct.size > 1){
      for(let i = 2; i <= lastRow; i++){
        if(trim(cell(i,'D')) !== ''){
          add(50, 'D','Error','Isian pada kolom D tidak boleh campur (harus 100 semua atau 160 semua atau 180 semua)', i);
        }
      }
    } else if(distinct.size === 1){
      const only = [...distinct][0];
      if(!['100','160','180'].includes(only)){
        for(let i = 2; i <= lastRow; i++){
          if(trim(cell(i,'D')) !== ''){
            add(51, 'D','Error','Isian pada kolom D tidak boleh campur (harus 100 semua atau 160 semua atau 180 semua)', i);
          }
        }
      }
    }
  }

  // Point 16 - Validate_B_to_J_Error (scan log yang sudah terkumpul)
  {
    const watched = new Set(['B','C','D','E','F','G','H','I','J']);
    const found = logs.some(l => l.type === 'Error' && watched.has(l.kolom));
    if(found){
      add(52, '','Error','Untuk perubahan/fixing pada Kolom B sampai J, maka Isian Kolom A harus dipilih Delete terlebih dahulu > Upload sampai Success > Kemudian Kolom A dipilih Create kembali', '');
    }
  }

  // Sort ascending by Kolom (mengikuti Range.Sort Key1:=A2 pada macro asli)
  logs.sort((a, b) => {
    if(a.kolom === '' && b.kolom === '') return 0;
    if(a.kolom === '') return 1;
    if(b.kolom === '') return -1;
    return a.kolom < b.kolom ? -1 : (a.kolom > b.kolom ? 1 : 0);
  });

  return logs;
}

/* =====================================================================
   5. UI WIRING
   ===================================================================== */
const fileInput = document.getElementById('fileInput');
const chooseFileBtn = document.getElementById('chooseFileBtn');
const fileNameLabel = document.getElementById('fileNameLabel');
const validateBtn = document.getElementById('validateBtn');
const dropZone = document.getElementById('dropZone');
const statusBanner = document.getElementById('statusBanner');
const summaryCard = document.getElementById('summaryCard');
const clearFilterBtn = document.getElementById('clearFilterBtn');
clearFilterBtn.addEventListener('click', clearAllFilters);
const tableCard = document.getElementById('tableCard');
const resultBody = document.getElementById('resultBody');
const badgeTotal = document.getElementById('badgeTotal');
const badgeError = document.getElementById('badgeError');
const badgeWarning = document.getElementById('badgeWarning');

let currentFile = null;
let currentArrayBuffer = null;
let currentData = null;
let currentLastRow = 0;
let autoFixedRows = new Set();

const badgeLineFix = document.getElementById('badgeLineFix');
const downloadCard = document.getElementById('downloadCard');
const downloadBtn = document.getElementById('downloadBtn');

function setFile(file){
  currentFile = file;
  fileNameLabel.textContent = file ? file.name : 'Belum ada file dipilih';
  validateBtn.disabled = !file;
  statusBanner.innerHTML = '';
  tableCard.hidden = true;
  closeFilterPanel();
  activeFilters = { errorId: null, kolom: null, line: null, type: null, msg: null };
  clearFilterBtn.disabled = true;
  autoFixedRows = new Set();
  currentArrayBuffer = null; currentData = null; currentLastRow = 0;
  badgeLineFix.hidden = true;
  badgeLineFix.textContent = 'Line Fixing : 0';
  downloadCard.hidden = true;
  downloadBtn.disabled = true;
  if(file){
    badgeTotal.textContent = 'Total Checking Line : 0';
    badgeError.textContent = 'Error : 0';
    badgeWarning.textContent = 'Warning : 0';
    summaryCard.hidden = false;
  } else {
    summaryCard.hidden = true;
  }
}

fileInput.addEventListener('change', () => {
  setFile(fileInput.files[0] || null);
});

chooseFileBtn.addEventListener('click', () => {
  fileInput.click();
});

['dragenter','dragover'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });
});
['dragleave','drop'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
  });
});
dropZone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if(file){
    fileInput.files = e.dataTransfer.files;
    setFile(file);
  }
});

function showBanner(type, text){
  statusBanner.innerHTML = '<div class="status-banner ' + type + '">' + text + '</div>';
}

let currentLogs = [];
let activeFilters = { errorId: null, kolom: null, line: null, type: null, msg: null };
let activePanel = null;
let activeAnchorTh = null;

function closeFilterPanel(){
  if(activePanel){ activePanel.remove(); activePanel = null; activeAnchorTh = null; }
}

function repositionFilterPanel(){
  if(!activePanel || !activeAnchorTh) return;
  const rect = activeAnchorTh.getBoundingClientRect();
  activePanel.style.top = (rect.bottom + 4) + 'px';
  let left = rect.left;
  const maxLeft = document.documentElement.clientWidth - 270;
  if(left > maxLeft) left = Math.max(8, maxLeft);
  activePanel.style.left = left + 'px';
}
window.addEventListener('scroll', repositionFilterPanel, true);
window.addEventListener('resize', repositionFilterPanel);

function passesFilters(l, excludeKey){
  for(const key of Object.keys(activeFilters)){
    if(key === excludeKey) continue;
    const f = activeFilters[key];
    if(f && !f.has(String(l[key]))) return false;
  }
  return true;
}

function getUniqueValuesExcluding(colKey){
  return [...new Set(currentLogs.filter(l => passesFilters(l, colKey)).map(l => String(l[colKey])))]
    .sort((a,b) => a.localeCompare(b, undefined, { numeric:true }));
}

function getFilteredLogs(){
  return currentLogs.filter(l => passesFilters(l, null));
}

function updateCaretStates(){
  document.querySelectorAll('.caret-btn').forEach(btn => {
    btn.classList.toggle('filtered', !!activeFilters[btn.dataset.col]);
  });
  const anyActive = Object.values(activeFilters).some(f => !!f);
  clearFilterBtn.disabled = !anyActive;
}

function clearAllFilters(){
  activeFilters = { errorId: null, kolom: null, line: null, type: null, msg: null };
  updateCaretStates();
  renderTable();
}

function openFilterPanel(colKey, anchorTh){
  if(activePanel && activePanel.dataset.col === colKey){ closeFilterPanel(); return; }
  closeFilterPanel();

  const values = getUniqueValuesExcluding(colKey);
  const activeSet = activeFilters[colKey];

  const panel = document.createElement('div');
  panel.className = 'filter-panel';
  panel.dataset.col = colKey;

  let html = '<div class="filter-search"><input type="text" id="fp-search" placeholder="Cari..." autocomplete="off">' +
    '<span class="search-icon">&#128269;</span></div>' +
    '<label class="filter-row filter-all"><input type="checkbox" id="fp-all"><span>(Pilih Semua)</span></label>' +
    '<div class="filter-list">';
  values.forEach((v, idx) => {
    const checked = !activeSet || activeSet.has(v);
    const label = v === '' ? '(Kosong)' : v;
    html += '<label class="filter-row" data-label="' + escapeHtml(label.toLowerCase()) + '">' +
      '<input type="checkbox" class="fp-item" data-idx="' + idx + '"' + (checked ? ' checked' : '') + '>' +
      '<span>' + escapeHtml(label) + '</span></label>';
  });
  html += '</div><div class="filter-actions"><button class="btn btn-sm" id="fp-cancel" type="button">Batal</button>' +
    '<button class="btn btn-sm primary" id="fp-ok" type="button">OK</button></div>';
  panel.innerHTML = html;
  document.body.appendChild(panel);

  activePanel = panel;
  activeAnchorTh = anchorTh;
  repositionFilterPanel();

  const allCb = panel.querySelector('#fp-all');
  const itemCbs = [...panel.querySelectorAll('.fp-item')];
  const searchInput = panel.querySelector('#fp-search');

  function visibleCbs(){
    return itemCbs.filter(cb => cb.closest('.filter-row').style.display !== 'none');
  }
  function updateSelectAllState(){
    const vis = visibleCbs();
    const checkedCount = vis.filter(cb => cb.checked).length;
    allCb.checked = vis.length > 0 && checkedCount === vis.length;
    allCb.indeterminate = checkedCount > 0 && checkedCount < vis.length;
  }

  searchInput.addEventListener('input', () => {
    const term = searchInput.value.trim().toLowerCase();
    itemCbs.forEach(cb => {
      const row = cb.closest('.filter-row');
      row.style.display = row.dataset.label.includes(term) ? '' : 'none';
    });
    updateSelectAllState();
  });

  allCb.addEventListener('change', () => {
    visibleCbs().forEach(cb => cb.checked = allCb.checked);
    allCb.indeterminate = false;
  });
  itemCbs.forEach(cb => cb.addEventListener('change', updateSelectAllState));
  updateSelectAllState();
  searchInput.focus();

  panel.querySelector('#fp-cancel').addEventListener('click', () => closeFilterPanel());
  panel.querySelector('#fp-ok').addEventListener('click', () => {
    const checkedVals = itemCbs.filter(cb => cb.checked).map(cb => values[parseInt(cb.dataset.idx, 10)]);
    activeFilters[colKey] = (checkedVals.length === values.length) ? null : new Set(checkedVals);
    closeFilterPanel();
    updateCaretStates();
    renderTable();
  });
}

document.querySelectorAll('.caret-btn').forEach(btn => {
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    openFilterPanel(btn.dataset.col, btn.closest('th'));
  });
});

function renderTable(){
  const rows = getFilteredLogs();
  resultBody.innerHTML = '';
  if(rows.length === 0){
    resultBody.innerHTML = '<tr><td colspan="5" class="empty-state">Tidak ada baris yang cocok dengan filter saat ini.</td></tr>';
    return;
  }
  for(const l of rows){
    const tr = document.createElement('tr');
    const isFixed = l.line !== '' && autoFixedRows.has(l.line) && FIXABLE_IDS.has(l.errorId);
    tr.className = isFixed ? 'row-fixed' : (l.type === 'Error' ? 'row-error' : 'row-warning');
    tr.innerHTML =
      '<td class="col-errid">' + escapeHtml(String(l.errorId)) + '</td>' +
      '<td class="col-kolom">' + escapeHtml(l.kolom) + '</td>' +
      '<td class="col-line">' + escapeHtml(String(l.line)) + '</td>' +
      '<td class="col-type">' + escapeHtml(l.type) + '</td>' +
      '<td>' + escapeHtml(l.msg) + '</td>';
    resultBody.appendChild(tr);
  }
}

/* =====================================================================
   6. MESSAGE BOX (mirip MsgBox pada macro asli, muncul setelah validasi)
   ===================================================================== */
const modalOverlay = document.getElementById('modalOverlay');
const modalIcon = document.getElementById('modalIcon');
const modalMessage = document.getElementById('modalMessage');
const modalOkBtn = document.getElementById('modalOkBtn');

const MODAL_ICONS = {
  error: '<svg width="56" height="56" viewBox="0 0 24 24"><circle cx="12" cy="12" r="11" fill="#e54848"/>' +
    '<path d="M8 8L16 16M16 8L8 16" stroke="#fff" stroke-width="2.2" stroke-linecap="round"/></svg>',
  warning: '<svg width="56" height="56" viewBox="0 0 24 24"><path d="M12 2.3L22.8 21H1.2L12 2.3Z" fill="#f5b400"/>' +
    '<rect x="11" y="9" width="2" height="6.2" rx="1" fill="#3a2c00"/>' +
    '<rect x="11" y="16.6" width="2" height="2" rx="1" fill="#3a2c00"/></svg>',
  success: '<svg width="56" height="56" viewBox="0 0 24 24"><circle cx="12" cy="12" r="11" fill="#4caf50"/>' +
    '<path d="M7 12.5L10.5 16L17 8.5" stroke="#fff" stroke-width="2.3" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>'
};

function showMessageBox(type, message){
  modalIcon.innerHTML = MODAL_ICONS[type];
  modalMessage.textContent = message;
  modalOverlay.hidden = false;
  modalOkBtn.focus();
}
function closeMessageBox(){ modalOverlay.hidden = true; }
modalOkBtn.addEventListener('click', closeMessageBox);
modalOverlay.addEventListener('click', (e) => { if(e.target === modalOverlay) closeMessageBox(); });
document.addEventListener('keydown', (e) => { if(e.key === 'Escape' && !modalOverlay.hidden) closeMessageBox(); });

function renderResults(logs){
  currentLogs = logs;
  activeFilters = { errorId: null, kolom: null, line: null, type: null, msg: null };
  autoFixedRows = new Set();
  closeFilterPanel();

  const errorCount = logs.filter(l => l.type === 'Error').length;
  const warningCount = logs.filter(l => l.type === 'Warning').length;

  // Line Fixing badge: count unique data rows that have at least one fixable errorId
  const fixingLineCount = new Set(logs.filter(l => FIXABLE_IDS.has(l.errorId) && l.line !== '').map(l => l.line)).size;
  badgeLineFix.textContent = 'Line Fixing : ' + fixingLineCount;
  badgeLineFix.hidden = fixingLineCount === 0;

  badgeTotal.textContent = 'Total Checking Line : ' + logs.length;
  badgeError.textContent = 'Error : ' + errorCount;
  badgeWarning.textContent = 'Warning : ' + warningCount;
  summaryCard.hidden = false;

  // Enable download button
  downloadCard.hidden = false;
  downloadBtn.disabled = false;

  if(logs.length === 0){
    tableCard.hidden = true;
    updateCaretStates();
    showBanner('ok', 'Checking Contract Line selesai tanpa issue.');
    showMessageBox('success', 'Proses Checking telah selesai. Silahkan coba upload Data anda ke Smart by GEP');
    return;
  }
  tableCard.hidden = false;
  updateCaretStates();
  renderTable();

  if(errorCount > 0){
    showBanner('error', 'Terdapat Issue pada data Contract Line anda. Silahkan check Issue Logs di bawah !!!');
    showMessageBox('error', 'Proses Checking telah selesai. dan terdapat Error yang harus di Fixing !!!');
  } else {
    showBanner('warn', 'Terdapat Catatan pada data Contract Line anda. Mohon pastikan kembali isian Contract Line anda sebelum upload ke Smart by GEP !!!');
    showMessageBox('warning', 'Proses Checking telah selesai. Mohon pastikan kembali jika Data anda sudah benar !');
  }
}

function escapeHtml(s){
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

/* =====================================================================
   7. AUTO-FIX & DOWNLOAD ENGINE
   ===================================================================== */

const FIXABLE_IDS = new Set([1,2,3,4,10,13,26,27,28,30,31,33,35,37,44,45,47,48]);

// CRC32
const CRC_TABLE=(()=>{const t=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[i]=c;}return t;})();
function crc32(data){let c=0xFFFFFFFF;for(let i=0;i<data.length;i++)c=CRC_TABLE[(c^data[i])&0xFF]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}

// Deflate
async function deflateRaw(data){
  const cs=new CompressionStream('deflate-raw');
  const w=cs.writable.getWriter();w.write(data);w.close();
  const chunks=[];const r=cs.readable.getReader();
  while(true){const{done,value}=await r.read();if(done)break;chunks.push(value);}
  let total=0;chunks.forEach(c=>total+=c.length);
  const out=new Uint8Array(total);let off=0;chunks.forEach(c=>{out.set(c,off);off+=c.length;});
  return out;
}

// ZIP Writer
class ZipWriter{
  constructor(){this.localParts=[];this.centralDir=[];this.offset=0;}
  async addFile(name,data){
    const nameBytes=new TextEncoder().encode(name);
    const raw=typeof data==='string'?new TextEncoder().encode(data):data;
    const comp=await deflateRaw(raw);
    const crc=crc32(raw);
    const lh=new Uint8Array(30+nameBytes.length);
    const lv=new DataView(lh.buffer);
    lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint16(6,0,true);lv.setUint16(8,8,true);
    lv.setUint16(10,0,true);lv.setUint16(12,0,true);
    lv.setUint32(14,crc,true);lv.setUint32(18,comp.length,true);lv.setUint32(22,raw.length,true);
    lv.setUint16(26,nameBytes.length,true);lv.setUint16(28,0,true);
    lh.set(nameBytes,30);
    const cd=new Uint8Array(46+nameBytes.length);
    const cv=new DataView(cd.buffer);
    cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);
    cv.setUint16(8,0,true);cv.setUint16(10,8,true);cv.setUint16(12,0,true);cv.setUint16(14,0,true);
    cv.setUint32(16,crc,true);cv.setUint32(20,comp.length,true);cv.setUint32(24,raw.length,true);
    cv.setUint16(28,nameBytes.length,true);cv.setUint16(30,0,true);cv.setUint16(32,0,true);
    cv.setUint16(34,0,true);cv.setUint16(36,0,true);cv.setUint32(38,0,true);
    cv.setUint32(42,this.offset,true);
    cd.set(nameBytes,46);
    this.localParts.push(lh,comp);this.centralDir.push(cd);
    this.offset+=lh.length+comp.length;
  }
  build(){
    let cdSize=0;this.centralDir.forEach(c=>cdSize+=c.length);
    const eocd=new Uint8Array(22);const ev=new DataView(eocd.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,this.centralDir.length,true);ev.setUint16(10,this.centralDir.length,true);
    ev.setUint32(12,cdSize,true);ev.setUint32(16,this.offset,true);
    const all=[...this.localParts,...this.centralDir,eocd];
    let total=0;all.forEach(p=>total+=p.length);
    const out=new Uint8Array(total);let off=0;all.forEach(p=>{out.set(p,off);off+=p.length;});
    return out;
  }
}

// Column number to letter
function colNumToLetter(n){
  let s='';while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;
}

// Apply auto-fixes → {fixedData, fixedCells}
function applyAutoFixes(data, lastRow, logs){
  const fd={};
  for(const r in data){fd[r]={};for(const c in data[r])fd[r][c]=data[r][c];}
  const fc=new Map();
  const mark=(row,col)=>{if(!fc.has(row))fc.set(row,new Set());fc.get(row).add(col);};
  const cell=(r,c)=>fd[r]?String(fd[r][c]||''):'';
  const set=(r,c,v)=>{if(!fd[r])fd[r]={};fd[r][c]=String(v);mark(r,c);};
  const fixable=logs.filter(l=>FIXABLE_IDS.has(l.errorId));
  const hasId=(id)=>fixable.some(l=>l.errorId===id);
  if(hasId(1)){for(let i=2;i<=lastRow;i++)set(i,'B',i-1);}
  if(hasId(2)){for(let i=2;i<=lastRow;i++)set(i,'F',i-1);}
  if(hasId(3)){for(let i=2;i<=lastRow;i++)set(i,'AI',i-1);}
  if(hasId(4)){for(let i=2;i<=lastRow;i++)set(i,'AI',100010001+(i-2)*10000);}
  fixable.filter(l=>l.errorId===10&&l.line!=='').forEach(l=>set(l.line,'I',''));
  const cols13=['Q','R','S','T','U','V','W','X','Y','Z','AA','AE','AF','AG'];
  fixable.filter(l=>l.errorId===13&&l.line!=='').forEach(l=>{if(cols13.includes(l.kolom))set(l.line,l.kolom,'');});
  fixable.filter(l=>l.errorId===26&&l.line!=='').forEach(l=>set(l.line,'AI',cell(l.line,'AI').replace(/^0+/,'')));
  fixable.filter(l=>l.errorId===27&&l.line!=='').forEach(l=>{const v=cell(l.line,'AJ');const idx=v.indexOf('-');if(idx!==-1)set(l.line,'AJ',v.slice(idx+1));});
  fixable.filter(l=>l.errorId===28&&l.line!=='').forEach(l=>set(l.line,'M',cell(l.line,'M').replace(/\s+/g,'')));
  fixable.filter(l=>l.errorId===30&&l.line!=='').forEach(l=>{const e=cell(l.line,'E');set(l.line,'J',e.length>0?e.charAt(0):'');});
  fixable.filter(l=>l.errorId===31&&l.line!=='').forEach(l=>{const e=cell(l.line,'E');set(l.line,'J','K0'+e.slice(0,2));});
  fixable.filter(l=>l.errorId===33&&l.line!=='').forEach(l=>set(l.line,'M','30'));
  fixable.filter(l=>l.errorId===35&&l.line!=='').forEach(l=>set(l.line,'J','S99'));
  fixable.filter(l=>l.errorId===37&&l.line!=='').forEach(l=>set(l.line,'M','0'));
  fixable.filter(l=>l.errorId===44&&l.line!=='').forEach(l=>{const v=toNumberID(cell(l.line,'O'));if(!isNaN(v))set(l.line,'O',Math.floor(v));});
  fixable.filter(l=>l.errorId===45&&l.line!=='').forEach(l=>{const v=toNumberID(cell(l.line,'O'));if(!isNaN(v))set(l.line,'O',Math.floor(v*100)/100);});
  fixable.filter(l=>l.errorId===47&&l.line!=='').forEach(l=>{const ai=cell(l.line,'AI'),aj=cell(l.line,'AJ');set(l.line,'AI',aj);set(l.line,'AJ',ai);});
  fixable.filter(l=>l.errorId===48&&l.line!=='').forEach(l=>{const ac=cell(l.line,'AC'),ad=cell(l.line,'AD');set(l.line,'AC',ad);set(l.line,'AD',ac);});
  return {fixedData:fd, fixedCells:fc};
}

// Build modified styles.xml → {xml, xfError, xfWarn, xfGreen, xfHeader}
function buildModifiedStyles(stylesXml){
  let xml=stylesXml;
  const fillCount=parseInt((xml.match(/<fills count="(\d+)"/)||[,2])[1]);
  const fontCount=parseInt((xml.match(/<fonts count="(\d+)"/)||[,1])[1]);
  const xfCount=parseInt((xml.match(/<cellXfs count="(\d+)"/)||[,1])[1]);
  const newFills=
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>'+
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>'+
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>'+
    '<fill><patternFill patternType="solid"><fgColor rgb="FF0070C0"/></patternFill></fill>';
  xml=xml.replace('</fills>',newFills+'</fills>');
  xml=xml.replace(/<fills count="(\d+)"/,`<fills count="${fillCount+4}"`);
  const newFonts=
    '<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>'+
    '<font><sz val="11"/><color rgb="FF9C6500"/><name val="Calibri"/></font>'+
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>';
  xml=xml.replace('</fonts>',newFonts+'</fonts>');
  xml=xml.replace(/<fonts count="(\d+)"/,`<fonts count="${fontCount+3}"`);
  const newXfs=
    `<xf numFmtId="0" fontId="${fontCount+0}" fillId="${fillCount+0}" borderId="0" applyFont="1" applyFill="1"/>`+
    `<xf numFmtId="0" fontId="${fontCount+1}" fillId="${fillCount+1}" borderId="0" applyFont="1" applyFill="1"/>`+
    `<xf numFmtId="0" fontId="0" fillId="${fillCount+2}" borderId="0" applyFill="1"/>`+
    `<xf numFmtId="0" fontId="${fontCount+2}" fillId="${fillCount+3}" borderId="0" applyFont="1" applyFill="1"/>`;
  xml=xml.replace('</cellXfs>',newXfs+'</cellXfs>');
  xml=xml.replace(/<cellXfs count="(\d+)"/,`<cellXfs count="${xfCount+4}"`);
  return {xml, xfError:xfCount+0, xfWarn:xfCount+1, xfGreen:xfCount+2, xfHeader:xfCount+3};
}

// Build Checking Log sheet XML
function buildCheckingLogSheet(logs, fixedLogRows, xfError, xfWarn, xfHeader){
  const enc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const headers=['Error ID','Kolom','Line','Type','Logs'];
  let rowsXml='<row r="1">';
  headers.forEach((h,ci)=>{
    rowsXml+=`<c r="${colNumToLetter(ci+1)}1" t="inlineStr" s="${xfHeader}"><is><t>${enc(h)}</t></is></c>`;
  });
  rowsXml+='</row>';
  logs.forEach((l,idx)=>{
    const rn=idx+2;
    const isFixed=l.line!==''&&fixedLogRows.has(l.line)&&FIXABLE_IDS.has(l.errorId);
    const xf=isFixed?0:(l.type==='Error'?xfError:xfWarn);
    [String(l.errorId),String(l.kolom),String(l.line),String(l.type),String(l.msg)].forEach((v,ci)=>{
      const sA=xf>0?` s="${xf}"`:'';
      rowsXml+=`<c r="${colNumToLetter(ci+1)+rn}" t="inlineStr"${sA}><is><t>${enc(v)}</t></is></c>`;
    });
    rowsXml=rowsXml.replace(`</row><c r="${colNumToLetter(0+1)+rn}`,`</row>`.slice(0,-6)+`<row r="${rn}"><c r="${colNumToLetter(0+1)+rn}`);
  });
  // Fix row wrapping (build properly)
  let properRows='<row r="1">';
  headers.forEach((h,ci)=>{
    properRows+=`<c r="${colNumToLetter(ci+1)}1" t="inlineStr" s="${xfHeader}"><is><t>${enc(h)}</t></is></c>`;
  });
  properRows+='</row>';
  logs.forEach((l,idx)=>{
    const rn=idx+2;
    const isFixed=l.line!==''&&fixedLogRows.has(l.line)&&FIXABLE_IDS.has(l.errorId);
    const xf=isFixed?0:(l.type==='Error'?xfError:xfWarn);
    properRows+=`<row r="${rn}">`;
    [String(l.errorId),String(l.kolom),String(l.line),String(l.type),String(l.msg)].forEach((v,ci)=>{
      const sA=xf>0?` s="${xf}"`:'';
      properRows+=`<c r="${colNumToLetter(ci+1)+rn}" t="inlineStr"${sA}><is><t>${enc(v)}</t></is></c>`;
    });
    properRows+='</row>';
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="10" customWidth="1"/><col min="3" max="3" width="8" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/><col min="5" max="5" width="80" customWidth="1"/></cols><sheetData>${properRows}</sheetData></worksheet>`;
}

// Modify Contract Lines sheet XML with fixes + green highlights
function modifyContractLinesXml(sheetXml, fixedData, fixedCells, greenXfIdx){
  // Text-based approach: avoids XMLSerializer namespace re-declaration issues.
  let xml = sheetXml;
  const enc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  // Convert column letter(s) to number: A=1, B=2, ..., Z=26, AA=27, AI=35 etc.
  const colToNum = c => {let n=0; for(let i=0;i<c.length;i++) n=n*26+(c.charCodeAt(i)-64); return n;};

  const buildCell = (ref, val, sIdx) => {
    const v = (val===null||val===undefined) ? '' : String(val);
    if(v==='') return `<c r="${ref}" s="${sIdx}"/>`;
    if(/^-?\d+(\.\d+)?$/.test(v)) return `<c r="${ref}" s="${sIdx}"><v>${v}</v></c>`;
    return `<c r="${ref}" t="inlineStr" s="${sIdx}"><is><t>${enc(v)}</t></is></c>`;
  };

  const findRowTagEnd = (xmlStr, rowNum) => {
    const tag = `<row r="${rowNum}"`;
    let pos = xmlStr.indexOf(tag);
    while(pos !== -1){
      const ch = xmlStr[pos + tag.length];
      if(ch===' '||ch==='>'||ch==='\n'||ch==='\t'||ch==='\r') break;
      pos = xmlStr.indexOf(tag, pos+1);
    }
    if(pos===-1) return -1;
    let end = pos + tag.length;
    while(end < xmlStr.length && xmlStr[end] !== '>') end++;
    return end;
  };

  const replaceCell = (xmlStr, ref, newCellXml) => {
    const rAttr = `r="${ref}"`;
    let rPos = xmlStr.indexOf(rAttr);
    while(rPos !== -1){
      const ch = xmlStr[rPos + rAttr.length];
      if(ch===' '||ch==='>'||ch==='/'||ch==='\n'||ch==='\t'){
        let start = rPos - 1;
        while(start > 0 && xmlStr[start] !== '<') start--;
        if(xmlStr.slice(start, start+2) === '<c'){
          let tagEnd = rPos;
          while(tagEnd < xmlStr.length && xmlStr[tagEnd] !== '>') tagEnd++;
          let cellEnd;
          if(xmlStr[tagEnd-1]==='/'){
            cellEnd = tagEnd+1;
          } else {
            const ci = xmlStr.indexOf('</c>', tagEnd);
            cellEnd = ci !== -1 ? ci+4 : tagEnd+1;
          }
          return xmlStr.slice(0, start) + newCellXml + xmlStr.slice(cellEnd);
        }
      }
      rPos = xmlStr.indexOf(rAttr, rPos + rAttr.length);
    }
    return null;
  };

  // Insert a new cell in the CORRECT column-order position within the row.
  // E.g. inserting B2 goes BEFORE C2, inserting AI3 goes AFTER AH3 (last existing cell).
  // Without this, out-of-order cells trigger Excel's auto-repair which discards cells.
  const insertCellOrdered = (xmlStr, rowNum, ref, newCellXml) => {
    const col = ref.replace(/\d+$/, '');
    const colNum = colToNum(col);
    const rowTagEnd = findRowTagEnd(xmlStr, rowNum);
    if(rowTagEnd === -1) return xmlStr;
    const rowClose = xmlStr.indexOf('</row>', rowTagEnd);
    if(rowClose === -1) return xmlStr;
    const rowContent = xmlStr.slice(rowTagEnd+1, rowClose);
    // Find the first existing cell whose column number is GREATER than ours
    const cellPat = /<c r="([A-Z]+)\d+"/g;
    let insertOffset = rowContent.length; // default: append at the end (before </row>)
    let m;
    while((m = cellPat.exec(rowContent)) !== null){
      if(colToNum(m[1]) > colNum){ insertOffset = m.index; break; }
    }
    const insertPos = rowTagEnd+1+insertOffset;
    return xmlStr.slice(0, insertPos) + newCellXml + xmlStr.slice(insertPos);
  };

  for(const [rowNum, cols] of fixedCells){
    const rowKey = String(rowNum);
    const rowData = fixedData[rowKey] || fixedData[rowNum] || {};
    for(const col of cols){
      const ref = col + rowNum;
      const val = rowData[col] ?? '';
      const newCell = buildCell(ref, val, greenXfIdx);
      const replaced = replaceCell(xml, ref, newCell);
      if(replaced !== null){
        xml = replaced;
      } else {
        // Cell not present in original — insert in correct column-order position
        xml = insertCellOrdered(xml, rowNum, ref, newCell);
      }
    }
  }
  return xml;
}

// Main Excel output builder
async function buildExcelOutput(doAutoFix, fixedData, fixedCells, fixedLogRows){
  const buf=new Uint8Array(currentArrayBuffer);
  const zipEntries=readZipEntries(buf);
  const entryMap={};zipEntries.forEach(e=>entryMap[e.name]=e);
  const allBytes={};
  for(const entry of zipEntries)allBytes[entry.name]=await extractEntry(buf,entry);
  const getText=name=>allBytes[name]?new TextDecoder('utf-8').decode(allBytes[name]):null;
  const wbXml=getText('xl/workbook.xml');
  const wbRelsXml=getText('xl/_rels/workbook.xml.rels');
  const stylesXml=getText('xl/styles.xml');
  // Find Contract Lines sheet
  const wbDoc=parseXml(wbXml);
  const sheetEls=wbDoc.getElementsByTagName('sheet');
  let contractRid=null;
  for(const s of sheetEls){if(s.getAttribute('name')==='Contract Lines'){contractRid=s.getAttribute('r:id');break;}}
  const relsDoc=parseXml(wbRelsXml);
  let contractTarget=null;
  for(const r of relsDoc.getElementsByTagName('Relationship')){if(r.getAttribute('Id')===contractRid){contractTarget=r.getAttribute('Target');break;}}
  const contractPath='xl/'+contractTarget.replace(/^\/?xl\//,'').replace(/^\.?\//,'');
  // Modify styles
  const {xml:newStylesXml,xfError,xfWarn,xfGreen,xfHeader}=buildModifiedStyles(stylesXml);
  // Build Checking Log sheet
  const checkingLogXml=buildCheckingLogSheet(currentLogs,fixedLogRows,xfError,xfWarn,xfHeader);
  // New sheet numbering
  const sheetCount=zipEntries.filter(e=>/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name)).length;
  const newSheetNum=sheetCount+1;
  const newSheetPath=`xl/worksheets/sheet${newSheetNum}.xml`;
  const newRid=`rIdCL${Date.now()}`;
  let maxSheetId=0;
  for(const s of sheetEls){const id=parseInt(s.getAttribute('sheetId')||'0');if(id>maxSheetId)maxSheetId=id;}
  // Update workbook.xml
  const newWbXml=wbXml.replace('</sheets>',`<sheet name="Checking Log" sheetId="${maxSheetId+1}" r:id="${newRid}"/></sheets>`);
  // Update workbook.xml.rels
  const newWbRelsXml=wbRelsXml
    .replace(/<Relationship[^>]*vbaProject[^>]*\/>/g,'')
    .replace('</Relationships>',`<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${newSheetNum}.xml"/></Relationships>`);
  // Update [Content_Types].xml
  const newContentTypes=(getText('[Content_Types].xml')||'')
    .replace('application/vnd.ms-excel.sheet.macroEnabled.main+xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    .replace(/<Override[^>]*vbaProject[^>]*\/>/g,'')
    .replace('</Types>',`<Override PartName="/${newSheetPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
  // Build zip
  const zip=new ZipWriter();
  const skip=new Set(['xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml','xl/styles.xml',contractPath,'xl/vbaProject.bin']);
  for(const entry of zipEntries){
    if(skip.has(entry.name))continue;
    await zip.addFile(entry.name,allBytes[entry.name]);
  }
  await zip.addFile('xl/workbook.xml',newWbXml);
  await zip.addFile('xl/_rels/workbook.xml.rels',newWbRelsXml);
  await zip.addFile('[Content_Types].xml',newContentTypes);
  await zip.addFile('xl/styles.xml',newStylesXml);
  if(doAutoFix&&fixedCells&&fixedCells.size>0){
    const modifiedXml=modifyContractLinesXml(getText(contractPath),fixedData,fixedCells,xfGreen);
    await zip.addFile(contractPath,modifiedXml);
  } else {
    await zip.addFile(contractPath,allBytes[contractPath]);
  }
  await zip.addFile(newSheetPath,checkingLogXml);
  return zip.build();
}

// Save file dialog
async function saveFileAs(buffer, suggestedName){
  if(window.showSaveFilePicker){
    try{
      const handle=await window.showSaveFilePicker({
        suggestedName,
        types:[{description:'Excel Files',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]
      });
      const writable=await handle.createWritable();
      await writable.write(buffer);await writable.close();
      return true;
    }catch(e){if(e.name==='AbortError')return false;}
  }
  const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download=suggestedName;
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return true;
}

// Download flow
const dlModal=document.getElementById('dlModal');
const dlCancel=document.getElementById('dlCancel');
const dlNo=document.getElementById('dlNo');
const dlYes=document.getElementById('dlYes');

function closeDlModal(){dlModal.hidden=true;}

async function doDownload(doAutoFix){
  closeDlModal();
  downloadBtn.disabled=true;
  downloadBtn.textContent='Memproses...';
  try{
    const fixableInLogs=currentLogs.filter(l=>FIXABLE_IDS.has(l.errorId));
    let fixedData=null,fixedCells=null;
    let fixedLogRows=new Set();
    if(doAutoFix&&fixableInLogs.length>0){
      const result=applyAutoFixes(currentData,currentLastRow,currentLogs);
      fixedData=result.fixedData;fixedCells=result.fixedCells;
      fixedLogRows=new Set(fixableInLogs.filter(l=>l.line!=='').map(l=>l.line));
    }
    const buf=await buildExcelOutput(doAutoFix&&fixableInLogs.length>0,fixedData,fixedCells,fixedLogRows);
    const baseName=(currentFile.name||'CMIIW').replace(/\.[^.]+$/,'').replace(/-(AutoFixed|Checked)$/,'');
    const suffix=(doAutoFix&&fixableInLogs.length>0)?'-AutoFixed':'-Checked';
    const saved=await saveFileAs(buf,baseName+suffix+'.xlsx');
    if(saved&&doAutoFix&&fixableInLogs.length>0){
      autoFixedRows=fixedLogRows;
      renderTable();
      const fixedCount=fixedLogRows.size;
      badgeLineFix.textContent='Line Fixing : '+fixedCount;
      badgeLineFix.hidden=false;
    }
  }catch(e){
    alert('Gagal membuat file: '+e.message);console.error(e);
  }finally{
    downloadBtn.disabled=false;downloadBtn.textContent='⬇ Download';
  }
}

dlCancel.addEventListener('click',closeDlModal);
dlNo.addEventListener('click',()=>doDownload(false));
dlYes.addEventListener('click',()=>doDownload(true));
dlModal.addEventListener('click',(e)=>{if(e.target===dlModal)closeDlModal();});

downloadBtn.addEventListener('click',()=>{
  if(!currentArrayBuffer||!currentLogs.length){doDownload(false);return;}
  const fixableExists=currentLogs.some(l=>FIXABLE_IDS.has(l.errorId));
  if(fixableExists){dlModal.hidden=false;}
  else{doDownload(false);}
});

validateBtn.addEventListener('click', async () => {
  if(!currentFile) return;
  closeFilterPanel();
  validateBtn.disabled = true;
  validateBtn.textContent = 'Memproses...';
  statusBanner.innerHTML = '';
  summaryCard.hidden = true;
  tableCard.hidden = true;
  try{
    const arrayBuffer = await currentFile.arrayBuffer();
    currentArrayBuffer = arrayBuffer;
    const { data, lastRow } = await extractContractLines(arrayBuffer);
    currentData = data;
    currentLastRow = lastRow;
    if(lastRow < 2){
      showBanner('error', 'Sheet "Contract Lines" tidak memiliki data untuk dicek.');
    } else {
      const logs = runValidation(data, lastRow);
      renderResults(logs);
    }
  } catch(err){
    showBanner('error', 'Gagal memproses file: ' + err.message);
    console.error(err);
  } finally {
    validateBtn.disabled = false;
    validateBtn.textContent = 'Validate';
  }
});