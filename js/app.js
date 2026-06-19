/* =================================================================
   STATE
   ================================================================= */
const S = {
  step:          1,
  appLinkName:   null,
  selectedForm:  null,
  workbook:      null,
  sheetName:     null,
  excelHeaders:  [],
  rawRows:       [],
  mapping:       {},     // { excelHeader -> fieldLinkName | '' }
  tableData:     [],     // [{ excelHeader: value, ... }]
  rowErrors:     {},     // { rowIdx: { excelHeader: errMsg } }
  rowStatus:     {},     // { rowIdx: 'success'|'failed' }
  selectedRows:  new Set(),
  importResults: null,
  isImporting:   false,
};

/* =================================================================
   INIT
   ================================================================= */
document.addEventListener('DOMContentLoaded', () => {
  populateFormSelector();
  setupUploadZone();
  initZohoSDK();
});

function initZohoSDK() {
  // V2 SDK auto-initializes via postMessage with the Creator parent iframe.
  // ZOHO.CREATOR.init() does NOT exist in V2 — calling it would throw.
  // ZOHO.CREATOR.DATA.addRecords() waits internally for the SDK Load event.
  if (typeof ZOHO === 'undefined' || !ZOHO.CREATOR || !ZOHO.CREATOR.DATA) {
    console.warn('[BulkImport] ZOHO V2 SDK not found — demo mode');
    return;
  }
  console.log('[BulkImport] ZOHO V2 SDK loaded | in Creator iframe:', isInsideCreator());
}

// Creator injects a `serviceOrigin` query param when embedding the widget.
// That is the reliable signal that we are running inside a live Creator app.
function isInsideCreator() {
  try {
    if (typeof ZOHO === 'undefined' || !ZOHO.CREATOR || !ZOHO.CREATOR.DATA) return false;
    const params = new URLSearchParams(window.location.search);
    // serviceOrigin = Creator injects this; also check iframe context as a fallback
    return params.has('serviceOrigin') || window.self !== window.top;
  } catch (e) {
    return false;
  }
}

function populateFormSelector() {
  const sel = document.getElementById('formSelect');
  FORM_CONFIG.forEach(f => {
    const opt = document.createElement('option');
    opt.value = f.id;
    opt.textContent = f.name;
    sel.appendChild(opt);
  });
}

/* =================================================================
   STEP 1 — FORM SELECTION
   ================================================================= */
function onFormSelect(formId) {
  S.selectedForm = FORM_CONFIG.find(f => f.id === formId) || null;
  const infoBox = document.getElementById('formInfo');
  if (!S.selectedForm) { infoBox.style.display = 'none'; return; }

  document.getElementById('templateFormName').textContent = S.selectedForm.name;

  const req = S.selectedForm.fields.filter(f => f.required);
  const opt = S.selectedForm.fields.filter(f => !f.required);

  infoBox.innerHTML = `
    <h3>${S.selectedForm.name} &mdash; ${S.selectedForm.fields.length} fields</h3>
    ${S.selectedForm.description ? `<p style="font-size:12px;color:#5f6368;margin-bottom:8px">${S.selectedForm.description}</p>` : ''}
    <div class="fields-grid">
      ${req.map(f => `<span class="field-chip required"><span class="req-dot"></span>${f.label}</span>`).join('')}
      ${opt.map(f => `<span class="field-chip">${f.label}</span>`).join('')}
    </div>
    <p class="legend">Red = required &nbsp;&bull;&nbsp; * must be filled in the Excel file</p>
  `;
  infoBox.style.display = 'block';
}

/* =================================================================
   STEP 2 — UPLOAD
   ================================================================= */
function setupUploadZone() {
  const zone = document.getElementById('uploadZone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  zone.addEventListener('click', e => {
    if (e.target.classList.contains('link-btn') || e.target === zone ||
        e.target.closest('.upload-icon') || e.target.classList.contains('upload-title') ||
        e.target.classList.contains('upload-sub') || e.target.classList.contains('upload-hint')) {
      document.getElementById('fileInput').click();
    }
  });
}

function handleFileSelect(event) {
  const file = event.target.files[0];
  if (file) handleFile(file);
}

function handleFile(file) {
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
    toast('Only .xlsx, .xls, and .csv files are supported.', 'err'); return;
  }
  if (file.size > 10 * 1024 * 1024) {
    toast('File size exceeds the 10 MB limit.', 'err'); return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      S.workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      populateSheetSelector();
      loadSheet(S.workbook.SheetNames[0]);
    } catch (err) {
      toast('Could not parse the file. Please check the format.', 'err');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
  showFilePreview(file);
}

function populateSheetSelector() {
  const sheets = S.workbook.SheetNames;
  const wrapper = document.getElementById('sheetSelectorWrapper');
  const sel = document.getElementById('sheetSelect');
  sel.innerHTML = '';
  sheets.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    sel.appendChild(opt);
  });
  wrapper.style.display = sheets.length > 1 ? 'flex' : 'none';
}

function onSheetSelect(sheetName) { loadSheet(sheetName); }

function loadSheet(sheetName) {
  S.sheetName = sheetName;
  const ws  = S.workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  if (raw.length < 2) {
    toast('Sheet must have a header row and at least one data row.', 'err'); return;
  }

  S.excelHeaders = raw[0].map(h => String(h).trim()).filter(h => h);
  S.rawRows = raw.slice(1).filter(row => row.some(c => String(c).trim() !== ''));

  const preview = document.getElementById('filePreview');
  const meta = preview.querySelector('.file-meta');
  if (meta) meta.textContent = `${S.rawRows.length} records • Sheet: ${sheetName}`;

  toast(`Loaded ${S.rawRows.length} records from "${sheetName}"`, 'ok');
}

function showFilePreview(file) {
  const preview = document.getElementById('filePreview');
  const ext = file.name.split('.').pop().toUpperCase();
  preview.innerHTML = `
    <span class="file-icon">${ext === 'CSV' ? '📄' : '📊'}</span>
    <div class="file-details">
      <div class="file-name">${esc(file.name)}</div>
      <div class="file-meta">Loading…</div>
    </div>
    <button class="remove-file" onclick="removeFile()" title="Remove file">&#10005;</button>
  `;
  preview.style.display = 'flex';
}

function removeFile() {
  S.workbook = null; S.excelHeaders = []; S.rawRows = [];
  document.getElementById('filePreview').style.display = 'none';
  document.getElementById('sheetSelectorWrapper').style.display = 'none';
  document.getElementById('fileInput').value = '';
}

/* =================================================================
   AUTO MAPPING (silent — no UI step)
   ================================================================= */
function buildAutoMapping() {
  S.mapping = {};
  S.excelHeaders.forEach(h => {
    const match = autoMatch(h, S.selectedForm.fields);
    S.mapping[h] = match ? match.linkName : '';
  });
}

function autoMatch(header, fields) {
  const norm = s => s.toLowerCase().replace(/[\s_\-./]+/g, '');
  const h = norm(header);
  return (
    fields.find(f => norm(f.linkName) === h) ||
    fields.find(f => norm(f.label) === h) ||
    fields.find(f => norm(f.linkName).includes(h) || h.includes(norm(f.linkName))) ||
    fields.find(f => norm(f.label).includes(h) || h.includes(norm(f.label))) ||
    null
  );
}

/* =================================================================
   STEP 3 — PREVIEW & EDIT
   ================================================================= */
function buildTableData() {
  S.tableData = S.rawRows.map(row => {
    const record = {};
    S.excelHeaders.forEach((h, i) => { record[h] = String(row[i] ?? '').trim(); });
    return record;
  });
  S.rowErrors   = {};
  S.rowStatus   = {};
  S.selectedRows = new Set();
}

/* ---- Header rename ---- */
function onHeaderChange(colIdx, newName) {
  newName = newName.trim();
  const oldName = S.excelHeaders[colIdx];
  if (!newName || newName === oldName) return;

  S.excelHeaders[colIdx] = newName;

  // Rename key in tableData
  S.tableData.forEach(record => {
    if (oldName in record) { record[newName] = record[oldName]; delete record[oldName]; }
  });

  // Re-map the renamed column
  delete S.mapping[oldName];
  const match = autoMatch(newName, S.selectedForm.fields);
  S.mapping[newName] = match ? match.linkName : '';

  // Rename key in rowErrors
  Object.values(S.rowErrors).forEach(errs => {
    if (oldName in errs) { errs[newName] = errs[oldName]; delete errs[oldName]; }
  });

  validateAll();
  renderImportWarnings();
}

/* ---- Table render ---- */
function renderDataTable() {
  const headers = S.excelHeaders;
  const colSpan = headers.length + 3; // check + # + status + data cols

  // Head
  document.getElementById('dataHead').innerHTML = `<tr>
    <th class="check-col" style="z-index:15">
      <input type="checkbox" id="checkAll" onchange="toggleAll(this.checked)">
    </th>
    <th class="row-num" style="z-index:15">#</th>
    <th class="status-col">&#10003;</th>
    ${headers.map((h, ci) => headerThHtml(h, ci)).join('')}
  </tr>`;

  // Body
  let html = '';
  S.tableData.forEach((record, ri) => {
    const errors = S.rowErrors[ri] || {};
    const hasErr = Object.keys(errors).length > 0;
    const st     = S.rowStatus[ri];
    const rowCls = [
      hasErr ? 'row-has-error' : '',
      S.selectedRows.has(ri) ? 'row-selected' : '',
      st === 'success' ? 'row-success' : '',
      st === 'failed'  ? 'row-failed'  : '',
    ].filter(Boolean).join(' ');

    html += `<tr id="row-${ri}" class="${rowCls}">
      <td class="check-col">
        <input type="checkbox" class="row-cb" data-ri="${ri}"
          onchange="toggleRow(${ri}, this.checked)" ${S.selectedRows.has(ri) ? 'checked' : ''}>
      </td>
      <td class="row-num">${ri + 1}</td>
      <td class="status-col" id="rstat-${ri}">${rowStatusIcon(ri)}</td>
      ${headers.map(h => cellHtml(ri, h, record[h] ?? '')).join('')}
    </tr>`;

    // Inline warning row — shown only when there are errors
    if (hasErr) {
      const tags = Object.entries(errors)
        .map(([h, msg]) => `<span class="row-warn-tag"><strong>${esc(h)}:</strong> ${esc(msg)}</span>`)
        .join('');
      html += `<tr id="row-warn-${ri}" class="row-warn-row">
        <td class="check-col"></td>
        <td class="row-num row-warn-num">&#9888;</td>
        <td colspan="${headers.length + 1}" class="row-warn-cell">
          <div class="row-warn-content">${tags}</div>
        </td>
      </tr>`;
    }
  });

  document.getElementById('dataBody').innerHTML = html;
  updateRecordCount();
  updateValBadge();
}

function headerThHtml(h, colIdx) {
  const lnk    = S.mapping[h];
  const field  = lnk ? S.selectedForm.fields.find(f => f.linkName === lnk) : null;
  const unmapped = !lnk;
  const reqCls   = field?.required ? ' req' : '';
  const tooltip  = unmapped
    ? 'No matching form field — rename this column to match, or it will be skipped'
    : `Maps to: ${field?.label || lnk}${field?.required ? ' (required)' : ''}`;

  return `<th class="${unmapped ? 'col-unmapped' : reqCls}">
    <input class="header-input${unmapped ? ' header-unmapped' : ''}"
      value="${escAttr(h)}"
      title="${escAttr(tooltip)}"
      onchange="onHeaderChange(${colIdx}, this.value)">
    ${unmapped ? '<span class="badge-skip-col">skip</span>' : ''}
  </th>`;
}

function cellHtml(ri, header, value) {
  const lnk = S.mapping[header];

  if (!lnk) {
    // Unmapped — read-only, grayed
    return `<td class="cell-unmapped"><span class="cell-skip-val">${esc(value)}</span></td>`;
  }

  const err = (S.rowErrors[ri] || {})[header];
  return `<td>
    <input class="cell-input${err ? ' cell-error' : ''}"
      type="text"
      value="${escAttr(value)}"
      ${err ? `title="${escAttr(err)}"` : ''}
      onchange="onCellChange(${ri}, ${JSON.stringify(header)}, this)"
      oninput="onCellInput(${ri}, ${JSON.stringify(header)}, this)">
  </td>`;
}

/* ---- Cell editing ---- */
function onCellInput(ri, header, input) {
  // Clear error highlight immediately as the user starts typing
  if (input.classList.contains('cell-error')) {
    input.classList.remove('cell-error');
    input.removeAttribute('title');
    if (S.rowErrors[ri]) delete S.rowErrors[ri][header];
    updateRowWarning(ri);
  }
}

function onCellChange(ri, header, input) {
  const value = input.value.trim();
  S.tableData[ri][header] = value;

  const lnk  = S.mapping[header];
  const field = lnk ? S.selectedForm.fields.find(f => f.linkName === lnk) : null;
  const err   = field ? fieldError(field, value) : null;

  if (!S.rowErrors[ri]) S.rowErrors[ri] = {};

  if (err) {
    S.rowErrors[ri][header] = err;
    input.classList.add('cell-error');
    input.title = err;
  } else {
    delete S.rowErrors[ri][header];
    input.classList.remove('cell-error');
    input.removeAttribute('title');
  }

  updateRowWarning(ri);
  updateValBadge();
}

/* ---- Row state helpers ---- */
function updateRowWarning(ri) {
  const tr = document.getElementById(`row-${ri}`);
  if (!tr) return;

  const errors = S.rowErrors[ri] || {};
  const hasErr = Object.keys(errors).length > 0;
  tr.classList.toggle('row-has-error', hasErr);

  const statCell = document.getElementById(`rstat-${ri}`);
  if (statCell) statCell.innerHTML = rowStatusIcon(ri);

  // Insert, update, or remove the inline warning row
  const existing = document.getElementById(`row-warn-${ri}`);

  if (hasErr) {
    const tags = Object.entries(errors)
      .map(([h, msg]) => `<span class="row-warn-tag"><strong>${esc(h)}:</strong> ${esc(msg)}</span>`)
      .join('');
    const warnHtml = `<tr id="row-warn-${ri}" class="row-warn-row">
      <td class="check-col"></td>
      <td class="row-num row-warn-num">&#9888;</td>
      <td colspan="${S.excelHeaders.length + 1}" class="row-warn-cell">
        <div class="row-warn-content">${tags}</div>
      </td>
    </tr>`;

    if (existing) {
      existing.outerHTML = warnHtml;
    } else {
      tr.insertAdjacentHTML('afterend', warnHtml);
    }
  } else if (existing) {
    existing.remove();
  }
}

function rowStatusIcon(ri) {
  const st = S.rowStatus[ri];
  if (st === 'success') return '<span style="color:#137333" title="Imported">&#10003;</span>';
  if (st === 'failed')  return '<span style="color:#c5221f" title="Failed">&#10007;</span>';
  const hasErr = S.rowErrors[ri] && Object.keys(S.rowErrors[ri]).length > 0;
  if (hasErr)  return '<span style="color:#f9ab00" title="Has errors">!</span>';
  return '<span style="color:#dadce0">&#8212;</span>';
}

/* =================================================================
   IMPORT WARNINGS PANEL (column-level)
   ================================================================= */
function renderImportWarnings() {
  const el = document.getElementById('importWarnings');
  if (!el) return;

  const parts = [];

  const missingRequired = S.selectedForm.fields
    .filter(f => f.required && !Object.values(S.mapping).includes(f.linkName))
    .map(f => f.label);

  if (missingRequired.length) {
    parts.push(`<div class="warn-item warn-error">
      <strong>&#9888; Missing required fields:</strong> ${esc(missingRequired.join(', '))}
      — not found in your file and will be empty.
    </div>`);
  }

  const unmapped = S.excelHeaders.filter(h => !S.mapping[h]);
  if (unmapped.length) {
    parts.push(`<div class="warn-item warn-warning">
      <strong>&#8505; Skipped columns (${unmapped.length}):</strong> ${esc(unmapped.join(', '))}
      — no matching form field; these will not be imported.
    </div>`);
  }

  el.innerHTML  = parts.join('');
  el.style.display = parts.length ? 'block' : 'none';
}

/* =================================================================
   VALIDATION
   ================================================================= */
function validateAll() {
  S.rowErrors = {};

  S.tableData.forEach((record, ri) => {
    S.excelHeaders.forEach(header => {
      const lnk  = S.mapping[header];
      if (!lnk) return;
      const field = S.selectedForm.fields.find(f => f.linkName === lnk);
      if (!field) return;
      const err = fieldError(field, record[header] ?? '');
      if (err) {
        if (!S.rowErrors[ri]) S.rowErrors[ri] = {};
        S.rowErrors[ri][header] = err;
      }
    });
  });

  renderDataTable();
  updateValBadge();
}

function clearAllErrors() {
  S.rowErrors = {};
  renderDataTable();
  updateValBadge();
}

function fieldError(field, value) {
  const v = String(value ?? '').trim();
  if (field.required && !v) return `${field.label} is required`;
  if (!v) return null;

  switch (field.type) {
    case 'email':
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return 'Invalid email address';
      break;
    case 'number':
      if (isNaN(Number(v.replace(/,/g, '')))) return 'Must be a valid number';
      break;
    case 'date':
      if (isNaN(Date.parse(v))) return 'Invalid date (use YYYY-MM-DD or MM/DD/YYYY)';
      break;
    case 'url':
      try { new URL(v); } catch { return 'Invalid URL (include https://)'; }
      break;
    case 'phone':
      if (!/^[+\d\s\-().]{6,20}$/.test(v)) return 'Invalid phone number';
      break;
  }
  return null;
}

function updateValBadge() {
  const badge    = document.getElementById('validationBadge');
  const errCount = Object.values(S.rowErrors).filter(e => Object.keys(e).length > 0).length;
  if (errCount === 0) {
    badge.textContent = `All ${S.tableData.length} records valid`;
    badge.className   = 'val-badge ok';
  } else {
    badge.textContent = `${errCount} row${errCount > 1 ? 's' : ''} with errors`;
    badge.className   = 'val-badge errors';
  }
}

function updateRecordCount() {
  document.getElementById('recordCount').textContent = `${S.tableData.length} records`;
}

/* =================================================================
   ROW SELECTION
   ================================================================= */
function toggleAll(checked) {
  document.querySelectorAll('.row-cb').forEach(cb => {
    cb.checked = checked;
    toggleRow(parseInt(cb.dataset.ri), checked);
  });
}

function toggleRow(ri, checked) {
  if (checked) S.selectedRows.add(ri);
  else         S.selectedRows.delete(ri);
  const row = document.getElementById(`row-${ri}`);
  if (row) row.classList.toggle('row-selected', checked);
  document.getElementById('btnDeleteRows').disabled = S.selectedRows.size === 0;
}

function deleteSelectedRows() {
  if (!S.selectedRows.size) return;
  const toRemove = [...S.selectedRows].sort((a, b) => b - a);
  toRemove.forEach(ri => { S.tableData.splice(ri, 1); });
  S.rowErrors   = {};
  S.rowStatus   = {};
  S.selectedRows = new Set();
  validateAll();
  toast(`Deleted ${toRemove.length} row${toRemove.length > 1 ? 's' : ''}`, 'ok');
}

/* =================================================================
   IMPORT
   ================================================================= */
async function startImport() {
  const errorRows = Object.entries(S.rowErrors)
    .filter(([, e]) => Object.keys(e).length > 0).length;

  if (errorRows > 0) {
    const skip = confirm(
      `${errorRows} row(s) have validation errors.\n\n` +
      `Click OK to skip those rows and import the valid records only.\n` +
      `Click Cancel to go back and fix the errors first.`
    );
    if (!skip) return;
  }

  // Build Zoho-field-keyed records from Excel-header-keyed tableData
  const toImport = S.tableData
    .map((rawRecord, ri) => {
      const zohoRecord = {};
      S.excelHeaders.forEach(h => {
        const lnk = S.mapping[h];
        if (lnk) zohoRecord[lnk] = rawRecord[h] ?? '';
      });
      return { zohoRecord, rawRecord, ri };
    })
    .filter(({ ri }) => !S.rowErrors[ri] || Object.keys(S.rowErrors[ri]).length === 0);

  if (!toImport.length) {
    toast('No valid records to import.', 'err'); return;
  }

  goToStep(4);

  S.isImporting = true;
  document.getElementById('btnImport').disabled = true;

  const total = toImport.length;
  let success = 0, failed = 0;
  const failedList    = [];
  const insertedIds   = [];

  for (let i = 0; i < toImport.length; i++) {
    const { zohoRecord, rawRecord, ri } = toImport[i];
    updateProgress(i + 1, total, `Inserting record ${i + 1} of ${total}…`, success, failed);

    try {
      const result = await callAddRecord(zohoRecord);
      success++;
      S.rowStatus[ri] = 'success';
      // V2 response: { code: 3000, data: { ID: "..." }, message: "success" }
      const newId = result?.data?.ID || result?.data?.id || null;
      insertedIds.push({ ri, rowNum: ri + 1, id: newId });
    } catch (err) {
      failed++;
      S.rowStatus[ri] = 'failed';
      failedList.push({
        rowNum: ri + 1,
        rawRecord,
        error: String(err?.message || err || 'Unknown error'),
      });
    }

    // Respect Creator API rate limits (~10 req/sec safe)
    await sleep(120);
  }

  S.isImporting   = false;
  S.importResults = { total, success, failed, failedList, insertedIds };
  showResults();
}

/* -----------------------------------------------------------------
   callAddRecord — Zoho Creator JS API v2
   SDK: ZOHO.CREATOR.DATA.addRecords(config)
   Config: { form_name, payload: { data: { Field_Link: value } } }
   workspace_name and app_name are auto-filled by the V2 SDK from
   initParams it receives from the Creator parent iframe.
   Success response: { code: 3000, data: { ID: "..." }, message: "success" }
   ----------------------------------------------------------------- */
async function callAddRecord(zohoRecord) {
  if (!isInsideCreator()) {
    // Running on localhost / GitHub Pages — simulate the API
    await sleep(60 + Math.random() * 80);
    if (Math.random() < 0.04) throw new Error('Simulated API error (demo mode)');
    return { code: 3000, data: { ID: 'DEMO-' + Math.random().toString(36).slice(2, 9) }, message: 'success' };
  }

  // workspace_name and app_name intentionally omitted —
  // the V2 SDK fills them automatically from the Creator iframe context
  const config = {
    form_name: S.selectedForm.linkName,
    payload:   { data: zohoRecord },
  };

  let response;
  try {
    response = await ZOHO.CREATOR.DATA.addRecords(config);
  } catch (sdkErr) {
    throw new Error(sdkErr?.message || 'SDK error during insert');
  }

  // V2: code 3000 = success; anything else = API-level error
  if (!response || response.code !== 3000) {
    throw new Error(response?.message || `Insert failed (code: ${response?.code ?? 'unknown'})`);
  }

  return response;
}

function updateProgress(current, total, label, success, failed) {
  const pct = Math.round((current / total) * 100);
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressLabel').textContent = label;
  document.getElementById('progressPct').textContent   = pct + '%';
  document.getElementById('progressCounts').innerHTML  =
    `<span style="color:#137333">&#10003; ${success} imported</span>` +
    `<span style="color:#c5221f">&#10007; ${failed} failed</span>`;
}

function showResults() {
  const { total, success, failed, failedList, insertedIds } = S.importResults;
  const isDemo = !isInsideCreator();

  document.getElementById('importProgress').style.display = 'none';
  document.getElementById('importResults').style.display  = 'block';

  document.getElementById('resultsGrid').innerHTML = `
    <div class="result-card r-total">
      <div class="result-num">${total}</div>
      <div class="result-label">Attempted</div>
    </div>
    <div class="result-card r-success">
      <div class="result-num">${success}</div>
      <div class="result-label">Inserted</div>
    </div>
    <div class="result-card r-failed">
      <div class="result-num">${failed}</div>
      <div class="result-label">Failed</div>
    </div>
    ${isDemo ? `<div class="result-card r-demo">
      <div class="result-num" style="font-size:18px">Demo</div>
      <div class="result-label">No real records created</div>
    </div>` : ''}
  `;

  // Show inserted IDs table when we have real IDs
  const hasIds = insertedIds.length > 0 && insertedIds.some(r => r.id && !r.id.startsWith('DEMO-'));
  if (hasIds) {
    document.getElementById('failedSection').insertAdjacentHTML('beforebegin', `
      <div class="inserted-section">
        <div class="failed-header">
          <h3 style="color:#137333">&#10003; Inserted Records — Creator IDs</h3>
        </div>
        <div class="table-scroll">
          <table class="failed-table" style="border-color:#ceead6">
            <thead>
              <tr style="background:#e6f4ea">
                <th style="border-color:#ceead6">Row #</th>
                <th style="border-color:#ceead6">Zoho Creator Record ID</th>
              </tr>
            </thead>
            <tbody>
              ${insertedIds.filter(r => r.id).map(({ rowNum, id }) =>
                `<tr><td>${rowNum}</td><td style="font-family:monospace;color:#137333">${esc(id)}</td></tr>`
              ).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `);
  }

  // Failed records table
  if (failedList.length > 0) {
    document.getElementById('failedSection').style.display = 'block';

    document.getElementById('failedHead').innerHTML = `<tr>
      <th>Row #</th>
      <th>Error</th>
      ${S.excelHeaders.map(h => `<th>${esc(h)}</th>`).join('')}
    </tr>`;

    document.getElementById('failedBody').innerHTML = failedList
      .map(({ rowNum, rawRecord, error }) => `<tr>
        <td>${rowNum}</td>
        <td style="color:#c5221f;min-width:180px">${esc(error)}</td>
        ${S.excelHeaders.map(h => `<td>${esc(rawRecord[h] ?? '')}</td>`).join('')}
      </tr>`)
      .join('');
  }

  document.getElementById('btnNext').style.display      = 'none';
  document.getElementById('btnImport').style.display    = 'none';
  document.getElementById('btnNewImport').style.display = 'inline-block';

  const msg = success > 0
    ? `${success} record${success > 1 ? 's' : ''} inserted into ${S.selectedForm.name}` + (failed ? `, ${failed} failed` : '')
    : `All ${total} records failed to insert`;
  toast(msg, success > 0 ? 'ok' : 'err');
}

/* =================================================================
   TEMPLATE DOWNLOAD
   ================================================================= */
function downloadTemplate() {
  if (!S.selectedForm) { toast('Select a form first.', 'warn'); return; }

  const fields = S.selectedForm.fields;
  const labels = fields.map(f => f.linkName);
  const hints  = fields.map(f => `${f.label}${f.required ? ' (required)' : ''} [${f.type}]`);
  const sample = fields.map(f => {
    switch (f.type) {
      case 'email':  return 'user@example.com';
      case 'number': return '100';
      case 'date':   return new Date().toISOString().split('T')[0];
      case 'phone':  return '+1 555-000-0000';
      case 'url':    return 'https://example.com';
      default:       return `Sample ${f.label}`;
    }
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([labels, hints, sample]);
  ws['!cols'] = fields.map(() => ({ wch: 22 }));
  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 18 };

  XLSX.utils.book_append_sheet(wb, ws, S.selectedForm.name);
  XLSX.writeFile(wb, `${S.selectedForm.linkName}_import_template.xlsx`);
  toast('Template downloaded', 'ok');
}

/* =================================================================
   DOWNLOAD FAILED RECORDS
   ================================================================= */
function downloadFailedRecords() {
  const { failedList } = S.importResults;
  if (!failedList.length) return;

  const headers = ['Row Number', 'Error', ...S.excelHeaders];
  const rows    = failedList.map(({ rowNum, rawRecord, error }) =>
    [rowNum, error, ...S.excelHeaders.map(h => rawRecord[h] ?? '')]
  );

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = headers.map(() => ({ wch: 22 }));
  XLSX.utils.book_append_sheet(wb, ws, 'Failed Records');
  XLSX.writeFile(wb, `failed_records_${Date.now()}.xlsx`);
  toast('Failed records downloaded', 'ok');
}

/* =================================================================
   STEP NAVIGATION
   ================================================================= */
function nextStep() {
  if (!canAdvance(S.step)) return;
  goToStep(S.step + 1);
}

function prevStep() {
  if (S.step > 1) goToStep(S.step - 1);
}

function goToStep(n) {
  const prev = document.getElementById(`panel${S.step}`);
  if (prev) prev.classList.remove('active');
  S.step = n;
  const next = document.getElementById(`panel${n}`);
  if (next) next.classList.add('active');
  onEnter(n);
  refreshStepNav();
  refreshFooter();
}

function onEnter(n) {
  if (n === 3) {
    buildAutoMapping();
    buildTableData();
    validateAll();          // validates then renders table
    renderImportWarnings(); // column-level banner
  }
  if (n === 4) {
    // Clean up any inserted-section left over from a previous run
    document.querySelectorAll('.inserted-section').forEach(el => el.remove());

    document.getElementById('importProgress').style.display = 'block';
    document.getElementById('importResults').style.display  = 'none';
    document.getElementById('failedSection').style.display  = 'none';
    document.getElementById('progressFill').style.width     = '0%';
    document.getElementById('progressLabel').textContent    = 'Ready to import…';
    document.getElementById('progressPct').textContent      = '0%';
    document.getElementById('progressCounts').innerHTML     = '';
  }
}

function canAdvance(step) {
  switch (step) {
    case 1:
      if (!S.selectedForm) { toast('Please select a form to continue.', 'warn'); return false; }
      return true;
    case 2:
      if (!S.rawRows.length) { toast('Please upload a file to continue.', 'warn'); return false; }
      return true;
    default:
      return true;
  }
}

function refreshStepNav() {
  document.querySelectorAll('.step').forEach(el => {
    const s = parseInt(el.dataset.step);
    el.classList.toggle('active',    s === S.step);
    el.classList.toggle('completed', s < S.step);
  });
  document.querySelectorAll('.step-line').forEach((line, i) => {
    line.classList.toggle('completed', i + 1 < S.step);
  });
}

function refreshFooter() {
  const back   = document.getElementById('btnBack');
  const next   = document.getElementById('btnNext');
  const imp    = document.getElementById('btnImport');
  const newImp = document.getElementById('btnNewImport');

  back.style.display   = S.step > 1 && S.step < 4 ? 'inline-block' : 'none';
  next.style.display   = S.step < 3               ? 'inline-block' : 'none';
  imp.style.display    = S.step === 3             ? 'inline-block' : 'none';
  newImp.style.display = 'none';

  if (S.step === 3) imp.disabled = false;
}

/* =================================================================
   RESET
   ================================================================= */
function resetWidget() {
  Object.assign(S, {
    step: 1, selectedForm: null, workbook: null, sheetName: null,
    excelHeaders: [], rawRows: [], mapping: {}, tableData: [],
    rowErrors: {}, rowStatus: {}, selectedRows: new Set(),
    importResults: null, isImporting: false,
  });

  document.getElementById('formSelect').value                      = '';
  document.getElementById('formInfo').style.display               = 'none';
  document.getElementById('filePreview').style.display            = 'none';
  document.getElementById('sheetSelectorWrapper').style.display   = 'none';
  document.getElementById('fileInput').value                      = '';

  document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel1').classList.add('active');

  refreshStepNav();
  refreshFooter();

  document.getElementById('btnNext').style.display      = 'inline-block';
  document.getElementById('btnNewImport').style.display = 'none';
}

/* =================================================================
   UTILITIES
   ================================================================= */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escAttr(s) { return String(s ?? '').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }

let _toastTimer;
function toast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = `toast show${type ? ' ' + type : ''}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
