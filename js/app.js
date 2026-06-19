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
  page:          0,
  pageSize:      200,
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
  S.page         = 0;
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

/* ---- Table render (paginated) ---- */
function renderDataTable() {
  const headers  = S.excelHeaders;
  const total    = S.tableData.length;
  const maxPg    = Math.max(0, Math.ceil(total / S.pageSize) - 1);
  if (S.page > maxPg) S.page = maxPg; // guard if rows were deleted

  const start    = S.page * S.pageSize;
  const end      = Math.min(start + S.pageSize, total);
  const colSpan  = headers.length + 3;

  // Head — no inline z-index; CSS handles sticky corners
  document.getElementById('dataHead').innerHTML = `<tr>
    <th class="check-col">
      <input type="checkbox" id="checkAll" onchange="toggleAll(this.checked)">
    </th>
    <th class="row-num">#</th>
    <th class="status-col">&#10003;</th>
    ${headers.map((h, ci) => headerThHtml(h, ci)).join('')}
  </tr>`;

  // Body — only render current page slice
  let html = '';
  for (let localIdx = 0; localIdx < end - start; localIdx++) {
    const ri     = start + localIdx;
    const record = S.tableData[ri];
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
  }

  document.getElementById('dataBody').innerHTML = html;
  renderPagination(start, end, total);
  updateRecordCount();
  updateValBadge();
  setupScrollShadow();
}

function renderPagination(start, end, total) {
  const pg = document.getElementById('tablePagination');
  if (!pg) return;
  const totalPgs = Math.ceil(total / S.pageSize);
  if (totalPgs <= 1) { pg.style.display = 'none'; return; }

  const canPrev = S.page > 0;
  const canNext = S.page < totalPgs - 1;

  pg.style.display = 'flex';
  pg.innerHTML = `
    <button class="btn-secondary btn-sm" onclick="prevPage()" ${canPrev ? '' : 'disabled'}>&#8592; Prev</button>
    <span class="page-info">Rows ${start + 1}–${end} of ${total}
      <span class="page-counter">Page ${S.page + 1} / ${totalPgs}</span>
    </span>
    <button class="btn-secondary btn-sm" onclick="nextPage()" ${canNext ? '' : 'disabled'}>Next &#8594;</button>
  `;
}

function prevPage() {
  if (S.page > 0) { S.page--; renderDataTable(); document.querySelector('.data-table-container').scrollTop = 0; }
}
function nextPage() {
  if (S.page < Math.ceil(S.tableData.length / S.pageSize) - 1) {
    S.page++;
    renderDataTable();
    document.querySelector('.data-table-container').scrollTop = 0;
  }
}

function setupScrollShadow() {
  const wrap = document.querySelector('.data-table-wrap');
  const c    = document.querySelector('.data-table-container');
  if (!wrap || !c) return;
  const update = () => {
    const atRight = c.scrollLeft >= c.scrollWidth - c.clientWidth - 2;
    const atLeft  = c.scrollLeft <= 2;
    wrap.classList.toggle('shadow-right', !atRight && c.scrollWidth > c.clientWidth + 4);
    wrap.classList.toggle('shadow-left',  !atLeft);
  };
  if (c._scrollShadow) c.removeEventListener('scroll', c._scrollShadow);
  c._scrollShadow = update;
  c.addEventListener('scroll', update, { passive: true });
  // re-check after paint (column widths settle after render)
  requestAnimationFrame(update);
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
      data-ri="${ri}"
      data-header="${escAttr(header)}"
      ${err ? `title="${escAttr(err)}"` : ''}
      onchange="onCellChange(${ri}, ${JSON.stringify(header)}, this)"
      oninput="onCellInput(${ri}, ${JSON.stringify(header)}, this)">
  </td>`;
}

/* ---- Cell editing ---- */
function onCellInput(ri, header, input) {
  // Sync to S.tableData on EVERY keystroke — do not wait for blur (onchange).
  // Without this, clicking Import or Validate All before blurring a cell
  // would read the old value from state instead of what's visible on screen.
  S.tableData[ri][header] = input.value.trim();

  const lnk  = S.mapping[header];
  const field = lnk ? S.selectedForm.fields.find(f => f.linkName === lnk) : null;

  if (field?.required && !input.value.trim()) {
    // Required field was just cleared — show error immediately, no need to blur
    if (!S.rowErrors[ri]) S.rowErrors[ri] = {};
    S.rowErrors[ri][header] = `${field.label} is required`;
    input.classList.add('cell-error');
    input.title = S.rowErrors[ri][header];
    updateRowWarning(ri);
    updateValBadge();
    updateImportButton();
  } else if (input.classList.contains('cell-error')) {
    // Value is being filled in — clear error immediately
    input.classList.remove('cell-error');
    input.removeAttribute('title');
    if (S.rowErrors[ri]) delete S.rowErrors[ri][header];
    updateRowWarning(ri);
    updateValBadge();
    updateImportButton();
  }
}

function onCellChange(ri, header, input) {
  const value = input.value.trim();
  if (input.value !== value) input.value = value; // strip leading/trailing whitespace visually
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
    // Auto-convert date to DD-MMM-YYYY on blur
    if (field?.type === 'date' && value) {
      const norm = parseAndFormatDate(value);
      if (norm && norm !== value) { input.value = norm; S.tableData[ri][header] = norm; }
    }
  }

  updateRowWarning(ri);
  updateValBadge();
  updateImportButton();
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

  el.innerHTML     = parts.join('');
  el.style.display = parts.length ? 'block' : 'none';
  updateImportButton();
}

/* =================================================================
   VALIDATION
   ================================================================= */
function flushDOMEdits() {
  // Sync every visible cell input → S.tableData before reading state.
  // Without this, if the user edits a cell and immediately clicks Import
  // (or Validate All), the browser destroys the inputs inside renderDataTable()
  // before onchange (blur) fires, leaving the old value in S.tableData.
  document.querySelectorAll('#dataBody .cell-input').forEach(input => {
    const ri     = Number(input.dataset.ri);
    const header = input.dataset.header;
    if (Number.isFinite(ri) && header && S.tableData[ri] !== undefined) {
      S.tableData[ri][header] = input.value.trim();
    }
  });
}

function validateAll() {
  flushDOMEdits(); // capture any in-progress edit before re-reading state

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

/* =================================================================
   DATE HELPERS
   ================================================================= */
const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function fmtDate(dt) {
  return `${String(dt.getDate()).padStart(2,'0')}-${MONTH_ABBR[dt.getMonth()]}-${dt.getFullYear()}`;
}

// Accepts any common date format; returns "DD-MMM-YYYY" or null if unparseable.
function parseAndFormatDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;

  // Already DD-MMM-YYYY (case-insensitive month)
  const already = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/i.exec(s);
  if (already) {
    const mi = MONTH_ABBR.findIndex(m => m.toLowerCase() === already[2].toLowerCase());
    if (mi >= 0) return `${String(+already[1]).padStart(2,'0')}-${MONTH_ABBR[mi]}-${already[3]}`;
  }

  let dt;

  // ISO YYYY-MM-DD — local constructor avoids UTC offset shifting the day
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (iso) {
    dt = new Date(+iso[1], +iso[2] - 1, +iso[3]);
    if (!isNaN(dt) && dt.getMonth() === +iso[2] - 1) return fmtDate(dt);
  }

  // DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY (European / Indian format)
  // Also catches MM/DD/YYYY by falling back when month-slot > 12
  const dmy = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(s);
  if (dmy) {
    const [, a, b, y] = dmy;
    if (+a > 12) {
      // First part must be day (DD/MM/YYYY)
      dt = new Date(+y, +b - 1, +a);
      if (!isNaN(dt) && dt.getMonth() === +b - 1) return fmtDate(dt);
    } else if (+b > 12) {
      // Second part must be day → this is MM/DD/YYYY
      dt = new Date(+y, +a - 1, +b);
      if (!isNaN(dt) && dt.getMonth() === +a - 1) return fmtDate(dt);
    } else {
      // Ambiguous — default to DD/MM/YYYY (international)
      dt = new Date(+y, +b - 1, +a);
      if (!isNaN(dt) && dt.getMonth() === +b - 1) return fmtDate(dt);
    }
  }

  // Fallback: let the browser parse ("Jun 15, 2024", "June 15 2024", JS Date.toString(), etc.)
  const ms = Date.parse(s);
  if (!isNaN(ms)) return fmtDate(new Date(ms));

  return null;
}

// After auto-mapping, convert every date field column to DD-MMM-YYYY in tableData
function normalizeDateFields() {
  if (!S.selectedForm) return;
  S.excelHeaders.forEach(header => {
    const lnk = S.mapping[header];
    if (!lnk) return;
    const field = S.selectedForm.fields.find(f => f.linkName === lnk);
    if (!field || field.type !== 'date') return;
    S.tableData.forEach(record => {
      const norm = parseAndFormatDate(record[header] ?? '');
      if (norm) record[header] = norm;
    });
  });
}

/* =================================================================
   FIELD VALIDATION
   ================================================================= */
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
      if (!parseAndFormatDate(v)) return 'Invalid date — enter as DD-MMM-YYYY (e.g. 15-Jun-2024)';
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
  const total    = S.tableData.length;
  const errCount = Object.values(S.rowErrors).filter(e => Object.keys(e).length > 0).length;
  const valid    = total - errCount;
  if (errCount === 0) {
    badge.textContent = `All ${total} record${total !== 1 ? 's' : ''} valid — ready to import`;
    badge.className   = 'val-badge ok';
  } else if (valid > 0) {
    badge.textContent = `${errCount} row${errCount > 1 ? 's' : ''} with errors — ${valid} will import`;
    badge.className   = 'val-badge errors';
  } else {
    badge.textContent = `All ${total} rows have errors — fix before importing`;
    badge.className   = 'val-badge errors';
  }
}

function updateImportButton() {
  const btn = document.getElementById('btnImport');
  if (!btn || S.step !== 3) return;

  const mappedCount = Object.values(S.mapping).filter(Boolean).length;
  if (mappedCount === 0) {
    btn.disabled = true;
    btn.title    = 'No columns are mapped to form fields — rename column headers to match field names';
    return;
  }

  const missingRequired = S.selectedForm
    ? S.selectedForm.fields
        .filter(f => f.required && !Object.values(S.mapping).includes(f.linkName))
        .map(f => f.label)
    : [];

  if (missingRequired.length > 0) {
    btn.disabled = true;
    btn.title    = `Cannot import — required fields not in your file: ${missingRequired.join(', ')}`;
  } else {
    btn.disabled = false;
    btn.title    = '';
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
  S.page         = 0;
  validateAll();
  toast(`Deleted ${toRemove.length} row${toRemove.length > 1 ? 's' : ''}`, 'ok');
}

/* =================================================================
   IMPORT
   ================================================================= */
async function startImport() {
  // Always run a fresh full validation pass before import
  validateAll();

  // Guard: required columns must be mapped (button should already be disabled, this is a safety net)
  const missingRequired = S.selectedForm.fields
    .filter(f => f.required && !Object.values(S.mapping).includes(f.linkName))
    .map(f => f.label);
  if (missingRequired.length) {
    toast(`Cannot import — required fields missing from file: ${missingRequired.join(', ')}`, 'err');
    return;
  }

  // Count rows with validation errors
  const errorRowIndices = new Set(
    Object.entries(S.rowErrors)
      .filter(([, e]) => Object.keys(e).length > 0)
      .map(([ri]) => Number(ri))
  );

  // Build Zoho-field-keyed records, skipping any row with validation errors
  const toImport = S.tableData
    .map((rawRecord, ri) => {
      const zohoRecord = {};
      S.excelHeaders.forEach(h => {
        const lnk = S.mapping[h];
        if (!lnk) return;
        const field = S.selectedForm.fields.find(f => f.linkName === lnk);
        let val = rawRecord[h] ?? '';
        // Ensure date fields are always in DD-MMM-YYYY when sent to Creator API
        if (field?.type === 'date' && val) val = parseAndFormatDate(val) || val;
        zohoRecord[lnk] = val;
      });
      return { zohoRecord, rawRecord, ri };
    })
    .filter(({ ri }) => !errorRowIndices.has(ri));

  if (!toImport.length) {
    toast(
      errorRowIndices.size > 0
        ? `All ${S.tableData.length} rows have validation errors — fix them before importing.`
        : 'No data to import.',
      'err'
    );
    return;
  }

  if (errorRowIndices.size > 0) {
    toast(
      `${errorRowIndices.size} row${errorRowIndices.size > 1 ? 's' : ''} with errors will be skipped — importing ${toImport.length} valid record${toImport.length > 1 ? 's' : ''}.`,
      'warn'
    );
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
      insertedIds.push({ ri, rowNum: ri + 1, id: newId, rawRecord });
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

  // First 3 mapped columns as row-level context in both tables
  const previewCols = S.excelHeaders.filter(h => S.mapping[h]).slice(0, 3);

  /* ── Summary cards ── */
  const cards = `
    <div class="result-card r-total">
      <div class="result-num">${total}</div>
      <div class="result-label">Attempted</div>
    </div>
    <div class="result-card r-success">
      <div class="result-num">${success}</div>
      <div class="result-label">&#10003; Inserted</div>
    </div>
    <div class="result-card r-failed">
      <div class="result-num">${failed}</div>
      <div class="result-label">&#10007; Failed</div>
    </div>`;

  /* ── Status banner ── */
  let banner;
  if (isDemo) {
    banner = `<div class="result-banner banner-demo">
      <strong>&#9432; Demo mode</strong> — no real records were created.
      Embed this widget inside Zoho Creator to perform a live import.
    </div>`;
  } else if (failed === 0) {
    banner = `<div class="result-banner banner-success">
      <strong>&#10003; All ${total} record${total > 1 ? 's' : ''} inserted</strong>
      into <em>${esc(S.selectedForm.name)}</em> successfully.
    </div>`;
  } else if (success === 0) {
    banner = `<div class="result-banner banner-error">
      <strong>&#10007; Import failed</strong> — all ${total} records could not be inserted.
      Fix the errors below and re-import.
    </div>`;
  } else {
    banner = `<div class="result-banner banner-partial">
      <strong>&#9888; Partial import</strong> — ${success} inserted, ${failed} failed.
      Fix the failed rows below and re-import them.
    </div>`;
  }

  /* ── Inserted records table ── */
  let insertedSection = '';
  if (insertedIds.length > 0) {
    const colHdrs = previewCols.map(h => `<th>${esc(h)}</th>`).join('');
    const rows = insertedIds.map(({ rowNum, id, rawRecord }) => `
      <tr class="res-ins-row">
        <td class="res-row-num">${rowNum}</td>
        <td class="res-id">${esc(id || '—')}</td>
        ${previewCols.map(h => `<td class="res-data" title="${escAttr(rawRecord[h] ?? '')}">${esc(truncate(rawRecord[h] ?? '', 28))}</td>`).join('')}
      </tr>`).join('');

    insertedSection = `
      <div class="res-section res-ins-section">
        <div class="res-section-hdr">
          <span class="res-section-title res-title-ins">
            &#10003; Inserted
            <span class="res-badge res-badge-ins">${success}</span>
          </span>
          <button class="btn-secondary btn-sm" onclick="downloadInsertedRecords()">&#8595; Download with IDs</button>
        </div>
        <div class="res-table-scroll">
          <table class="results-status-table res-ins-table">
            <thead><tr>
              <th>Row #</th>
              <th>Creator Record ID</th>
              ${colHdrs}
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  }

  /* ── Failed records table ── */
  let failedSection = '';
  if (failedList.length > 0) {
    const colHdrs = previewCols.map(h => `<th>${esc(h)}</th>`).join('');
    const rows = failedList.map(({ rowNum, rawRecord, error }) => `
      <tr class="res-fail-row">
        <td class="res-row-num">${rowNum}</td>
        <td class="res-error">${esc(error)}</td>
        ${previewCols.map(h => `<td class="res-data" title="${escAttr(rawRecord[h] ?? '')}">${esc(truncate(rawRecord[h] ?? '', 28))}</td>`).join('')}
      </tr>`).join('');

    failedSection = `
      <div class="res-section res-fail-section">
        <div class="res-section-hdr">
          <span class="res-section-title res-title-fail">
            &#10007; Failed
            <span class="res-badge res-badge-fail">${failed}</span>
          </span>
          <button class="btn-secondary btn-sm" onclick="downloadFailedRecords()">&#8595; Download Errors</button>
        </div>
        <p class="failed-hint">Fix these rows and re-import.</p>
        <div class="res-table-scroll">
          <table class="results-status-table res-fail-table">
            <thead><tr>
              <th>Row #</th>
              <th>Error</th>
              ${colHdrs}
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  }

  /* ── Render ── */
  const el = document.getElementById('importResults');
  el.innerHTML = `<div class="results-grid">${cards}</div>${banner}${insertedSection}${failedSection}`;
  el.style.display = 'block';

  document.getElementById('importProgress').style.display = 'none';
  document.getElementById('btnNext').style.display        = 'none';
  document.getElementById('btnImport').style.display      = 'none';
  document.getElementById('btnNewImport').style.display   = 'inline-block';

  const msg = failed === 0
    ? `${success} record${success > 1 ? 's' : ''} inserted into ${S.selectedForm.name}`
    : `${success} inserted, ${failed} failed — see results`;
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
      case 'date':   return fmtDate(new Date());
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
function downloadInsertedRecords() {
  const { insertedIds } = S.importResults;
  if (!insertedIds.length) return;

  const headers = ['Row Number', 'Creator Record ID', ...S.excelHeaders];
  const rows    = insertedIds.map(({ rowNum, id, rawRecord }) =>
    [rowNum, id || '', ...S.excelHeaders.map(h => rawRecord[h] ?? '')]
  );

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = headers.map(() => ({ wch: 22 }));
  XLSX.utils.book_append_sheet(wb, ws, 'Inserted Records');
  XLSX.writeFile(wb, `inserted_records_${Date.now()}.xlsx`);
  toast('Inserted records downloaded', 'ok');
}

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
    buildTableData();
    buildAutoMapping();
    normalizeDateFields();  // convert all date columns to DD-MMM-YYYY before display
    validateAll();
    renderImportWarnings();
  }
  if (n === 4) {
    document.getElementById('importProgress').style.display = 'block';
    document.getElementById('importResults').style.display  = 'none';
    document.getElementById('importResults').innerHTML      = '';
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

  if (S.step === 3) updateImportButton();
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
function truncate(s, n) { s = String(s ?? ''); return s.length > n ? s.slice(0, n) + '…' : s; }

let _toastTimer;
function toast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = `toast show${type ? ' ' + type : ''}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
