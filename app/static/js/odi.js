// ── ODI · Clasificador Multi-Base de Datos (Multi-Colaborador & UX/UI) ─────────

let odiFiles = [];
let odiUnifiedRucMap = {};
let odiAllSheets = [];
const odiDisabledFileIds = new Set();
let odiActiveScope = 'all'; // 'all' | 'my'
let odiManageFilter = 'all'; // 'all' | 'mine'
let odiCurrentUserId = null;
let odiLastResults = [];

// Elementos del DOM
const odiPage               = document.getElementById('odiPage');
const odiStatFilesVal       = document.getElementById('odiStatFilesVal');
const odiStatRucsVal        = document.getElementById('odiStatRucsVal');
const odiManagerCount       = document.getElementById('odiManagerCount');
const odiRucList            = document.getElementById('odiRucList');
const odiRucCount           = document.getElementById('odiRucCount');
const odiBtnClassify        = document.getElementById('odiBtnClassify');
const odiBasesSwitchList    = document.getElementById('odiBasesSwitchList');
const odiScopeHint          = document.getElementById('odiScopeHint');
const odiToggleAllBasesBtn  = document.getElementById('odiToggleAllBasesBtn');
const odiScopeAllBtn        = document.getElementById('odiScopeAllBtn');
const odiScopeMyBtn         = document.getElementById('odiScopeMyBtn');
const odiResultsStats       = document.getElementById('odiResultsStats');
const odiResultsTable       = document.getElementById('odiResultsTable');
const odiResultsEmpty       = document.getElementById('odiResultsEmpty');
const odiExportGroup        = document.getElementById('odiExportGroup');
const odiResultFilterWrap   = document.getElementById('odiResultFilterWrap');
const odiResultFilterInput  = document.getElementById('odiResultFilterInput');
const odiSheetsEl           = document.getElementById('odiSheets');
const odiSheetsTotal        = document.getElementById('odiSheetsTotal');
const odiSheetFilter        = document.getElementById('odiSheetFilter');

// Modales
const odiUploadModal        = document.getElementById('odiUploadModal');
const odiManageModal        = document.getElementById('odiManageModal');
const odiDropzone           = document.getElementById('odiDropzone');
const odiFileInput          = document.getElementById('odiFileInput');
const odiSpinner            = document.getElementById('odiSpinner');
const odiManageTableBody    = document.getElementById('odiManageTableBody');
const odiManageSearchInput  = document.getElementById('odiManageSearchInput');

function odiStart() {
  const page = document.getElementById('odiPage');
  if (page && page.dataset.userId) {
    odiCurrentUserId = page.dataset.userId;
  }
  odiInitData();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', odiStart);
} else {
  odiStart();
}

// Cargar datos iniciales desde el backend
async function odiInitData() {
  try {
    const r = await fetch('/odi/api/data');
    if (!r.ok) {
      console.warn("API /odi/api/data status:", r.status);
      odiApplyDataset({ files: [], unified_ruc_map: {}, all_sheets: [] });
      return;
    }
    const contentType = r.headers.get('content-type') || '';
    if (!contentType.includes('application/json')) {
      console.warn("Respuesta no es JSON:", contentType);
      odiApplyDataset({ files: [], unified_ruc_map: {}, all_sheets: [] });
      return;
    }
    const res = await r.json();
    odiApplyDataset(res);
  } catch (e) {
    console.error("Error al obtener datos almacenados de ODI:", e);
    odiApplyDataset({ files: [], unified_ruc_map: {}, all_sheets: [] });
  }
}

function odiApplyDataset(res) {
  odiFiles = res.files || [];
  odiUnifiedRucMap = res.unified_ruc_map || {};
  odiAllSheets = res.all_sheets || [];
  if (res.current_user_id) {
    odiCurrentUserId = res.current_user_id;
  }

  // Limpiar IDs deshabilitados que ya no existan
  const validIds = new Set(odiFiles.map(f => f.id));
  odiDisabledFileIds.forEach(id => {
    if (!validIds.has(id)) odiDisabledFileIds.delete(id);
  });

  // Actualizar métricas de la barra superior
  if (odiStatFilesVal) odiStatFilesVal.textContent = odiFiles.length;
  if (odiStatRucsVal) odiStatRucsVal.textContent = (res.total_rucs || 0).toLocaleString();
  if (odiManagerCount) odiManagerCount.textContent = odiFiles.length;

  // Renderizar selector de bases activas con switches
  odiRenderBasesSwitchList();

  // Renderizar hojas en panel lateral derecho
  if (odiSheetsTotal) odiSheetsTotal.textContent = odiAllSheets.length;
  if (odiSheetFilter) odiSheetFilter.value = '';
  odiRenderSheets(odiAllSheets);

  // Renderizar tabla del modal gestor si está abierto
  odiRenderManageList();

  // Si ya hay RUCs ingresados, reclasificar con los datos actualizados
  if (odiRucList && odiRucList.value.trim().length > 0) {
    odiClassifyList();
  }
}

// ── Control de Alcance (Scope: Todas vs Mis Bases) ───────────────────────────
function odiSetScope(scope) {
  odiActiveScope = scope;
  if (odiScopeAllBtn && odiScopeMyBtn) {
    odiScopeAllBtn.classList.toggle('active', scope === 'all');
    odiScopeMyBtn.classList.toggle('active', scope === 'my');
  }

  if (odiScopeHint) {
    odiScopeHint.textContent = scope === 'my'
      ? "Buscando solo en tus archivos"
      : "Buscando en todos los archivos";
  }

  odiRenderBasesSwitchList();

  // Reclasificar inmediatamente si hay texto
  if (odiRucList && odiRucList.value.trim().length > 0) {
    odiClassifyList();
  }
}

// Renderizar la lista de bases en la Columna 1 con switches
function odiRenderBasesSwitchList() {
  if (!odiBasesSwitchList) return;

  const filteredFiles = odiFiles.filter(f => {
    if (odiActiveScope === 'my') {
      return Boolean(f.is_mine);
    }
    return true;
  });

  if (!filteredFiles.length) {
    odiBasesSwitchList.innerHTML = `
      <div class="odi-empty-bases-hint">
        <i class="fa-solid fa-folder-open"></i>
        <span>${odiActiveScope === 'my' ? 'No has subido archivos aún.' : 'No hay archivos cargados.'}</span>
        <button type="button" class="odi-btn-link-action" onclick="odiOpenUploadModal()">+ Subir Excel ahora</button>
      </div>`;
    return;
  }

  odiBasesSwitchList.innerHTML = filteredFiles.map(f => {
    const isEnabled = !odiDisabledFileIds.has(f.id);
    const authorPill = f.is_mine
      ? `<span class="odi-author-badge mine"><i class="fa-solid fa-user-check"></i> Tuya</span>`
      : `<span class="odi-author-badge other" title="Subido por: ${odiEsc(f.uploaded_by)}"><i class="fa-regular fa-user"></i> ${odiEsc(f.uploaded_by)}</span>`;

    return `
      <div class="odi-base-switch-item ${isEnabled ? 'active' : 'inactive'}">
        <label class="odi-switch-control">
          <input type="checkbox" ${isEnabled ? 'checked' : ''} onchange="odiToggleBase('${f.id}')">
          <span class="odi-slider"></span>
        </label>
        <div class="odi-base-info-block" title="${odiEsc(f.filename)}&#10;Subido el: ${odiEsc(f.uploaded_at)}&#10;${f.sheets_count} hojas | ${f.total_rucs} RUCs">
          <div class="odi-base-name-row">
            <span class="odi-base-name">${odiEsc(f.filename)}</span>
            ${authorPill}
          </div>
          <div class="odi-base-metrics-row">
            <span><i class="fa-solid fa-id-card"></i> <b>${f.total_rucs.toLocaleString()}</b></span>
            <span><i class="fa-regular fa-file-lines"></i> ${f.sheets_count} hojas</span>
          </div>
        </div>
      </div>`;
  }).join('');

  // Actualizar texto del botón activar/desactivar todas
  if (odiToggleAllBasesBtn) {
    const anyEnabled = filteredFiles.some(f => !odiDisabledFileIds.has(f.id));
    odiToggleAllBasesBtn.textContent = anyEnabled ? "Desactivar todas" : "Activar todas";
  }
}

// Activar o desactivar una base de datos individualmente
function odiToggleBase(fileId) {
  if (odiDisabledFileIds.has(fileId)) {
    odiDisabledFileIds.delete(fileId);
  } else {
    odiDisabledFileIds.add(fileId);
  }

  odiRenderBasesSwitchList();

  // Reclasificar inmediatamente en vivo
  if (odiRucList && odiRucList.value.trim().length > 0) {
    odiClassifyList();
  }
}

// Alternar entre activar o desactivar todas las bases del alcance actual
function odiToggleAllActiveBases() {
  const currentFiles = odiFiles.filter(f => {
    if (odiActiveScope === 'my') return Boolean(f.is_mine);
    return true;
  });

  const anyEnabled = currentFiles.some(f => !odiDisabledFileIds.has(f.id));

  if (anyEnabled) {
    // Desactivar todas las actuales
    currentFiles.forEach(f => odiDisabledFileIds.add(f.id));
  } else {
    // Activar todas las actuales
    currentFiles.forEach(f => odiDisabledFileIds.delete(f.id));
  }

  odiRenderBasesSwitchList();

  if (odiRucList && odiRucList.value.trim().length > 0) {
    odiClassifyList();
  }
}

// Limpiar RUCs del textarea
function odiClearRucs() {
  if (odiRucList) {
    odiRucList.value = '';
    odiClassifyList();
  }
}

// ── Clasificación Multi-Base (RUC, DNI, Nombre o Código) ─────────────────────
function odiSetText(el, text) {
  if (el) el.textContent = text;
}

function odiSetDisplay(el, value) {
  if (el) el.style.display = value;
}

function odiShowEmptyResults() {
  if (odiResultsTable) odiResultsTable.innerHTML = '';
  odiSetDisplay(odiResultsEmpty, '');
  odiSetText(odiResultsStats, 'Esperando consultas...');
  odiSetDisplay(odiExportGroup, 'none');
  if (odiResultFilterWrap) odiResultFilterWrap.classList.add('hidden');
  odiLastResults = [];
}

function odiClassifyList() {
  if (!odiRucList) return;

  const queries = odiRucList.value
    .split(/[\r\n,;\t]+/)
    .map(v => v.trim())
    .filter(v => v.length > 0);

  const unique = [...new Set(queries)];
  odiSetText(odiRucCount, unique.length ? `${unique.length} elemento(s)` : '0 elementos');

  if (!unique.length) {
    odiShowEmptyResults();
    return;
  }

  odiSetDisplay(odiResultsEmpty, 'none');
  if (odiResultFilterWrap) odiResultFilterWrap.classList.remove('hidden');

  let foundCount = 0;
  odiLastResults = [];

  const rowsHtml = unique.map(q => {
    const row = odiBuildResultRow(q);
    if (row.isMatch) foundCount++;

    odiLastResults.push({
      ruc: q,
      status: row.isMatch ? 'ENCONTRADO' : 'NO ENCONTRADO',
      matches_count: (row.matches || []).length,
      files: row.files || [],
      sheets: row.sheets || [],
      matches: row.matches || []
    });

    return row.html;
  }).join('');

  if (odiResultsTable) odiResultsTable.innerHTML = rowsHtml;
  odiSetText(odiResultsStats, `${foundCount} de ${unique.length} encontrados`);
  odiSetDisplay(odiExportGroup, 'inline-flex');
}

function odiIsFileMine(fileId) {
  const f = odiFiles.find(item => item.id === fileId);
  return f ? Boolean(f.is_mine) : false;
}

function odiBuildResultRow(query) {
  const cleanQ = query.trim();
  if (cleanQ.length < 2) {
    return {
      isMatch: false,
      files: [], sheets: [], matches: [],
      html: `
        <div class="odi-result-card warn" data-ruc="${odiEsc(cleanQ)}">
          <div class="odi-card-head">
            <span class="odi-card-ruc">${odiEsc(cleanQ)}</span>
            <span class="odi-badge-status warn"><i class="fa-solid fa-triangle-exclamation"></i> Consulta muy corta</span>
          </div>
          <p class="odi-card-note">Ingresa al menos 2 caracteres (RUC, DNI, nombre o código).</p>
        </div>`
    };
  }

  let allEntries = odiUnifiedRucMap[cleanQ] || odiUnifiedRucMap[cleanQ.toUpperCase()] || [];

  if (!allEntries.length) {
    const upperQ = cleanQ.toUpperCase();
    const cleanDigits = cleanQ.replace(/\D/g, '');

    const matchedKeys = new Set();
    Object.keys(odiUnifiedRucMap).forEach(key => {
      const upperKey = key.toUpperCase();
      const matchesText = upperKey === upperQ || (upperQ.length >= 3 && upperKey.includes(upperQ));
      if (matchesText || (cleanDigits.length >= 7 && key.includes(cleanDigits))) {
        matchedKeys.add(key);
      }
    });

    const combined = [];
    matchedKeys.forEach(k => {
      if (odiUnifiedRucMap[k]) combined.push(...odiUnifiedRucMap[k]);
    });
    allEntries = combined;
  }

  // Filtrar según bases activas
  const activeEntries = allEntries.filter(e => {
    if (odiDisabledFileIds.has(e.file_id)) return false;
    if (odiActiveScope === 'my' && !odiIsFileMine(e.file_id)) return false;
    return true;
  });

  if (activeEntries && activeEntries.length) {
    const byFile = {};
    const fileSet = new Set();
    const sheetSet = new Set();

    activeEntries.forEach(e => {
      const fname = e.file_name || 'Base de Datos';
      if (!byFile[fname]) byFile[fname] = [];
      byFile[fname].push(e);
      fileSet.add(fname);
      if (e.sheet) sheetSet.add(e.sheet);
    });

    const fileCount = Object.keys(byFile).length;
    const fileGroupsHtml = Object.entries(byFile).map(([fname, sheetEntries]) => {
      const sheetsHtml = sheetEntries.map(e => odiBuildSheetAccordion(cleanQ, e, false)).join('');
      const uploaderText = sheetEntries[0]?.uploader_name ? ` · Subido por ${odiEsc(sheetEntries[0].uploader_name)}` : '';

      return `
        <div class="odi-file-match-box">
          <div class="odi-file-match-bar">
            <div class="odi-file-match-meta">
              <i class="fa-solid fa-file-excel"></i>
              <span class="odi-file-match-name">${odiEsc(fname)}</span>
              <span class="odi-file-match-uploader">${uploaderText}</span>
            </div>
            <span class="odi-file-match-badge">${sheetEntries.length} registro${sheetEntries.length > 1 ? 's' : ''}</span>
          </div>
          <div class="odi-sheet-items-list">${sheetsHtml}</div>
        </div>`;
    }).join('');

    return {
      isMatch: true,
      files: Array.from(fileSet),
      sheets: Array.from(sheetSet),
      matches: activeEntries,
      html: `
        <div class="odi-result-card ok" data-ruc="${odiEsc(cleanQ)}">
          <div class="odi-card-head">
            <span class="odi-card-ruc">${odiEsc(cleanQ)}</span>
            <span class="odi-badge-status ok">
              <i class="fa-solid fa-circle-check"></i> Encontrado en ${fileCount} base${fileCount > 1 ? 's' : ''} (${activeEntries.length} coincidencia${activeEntries.length > 1 ? 's' : ''})
            </span>
          </div>
          <div class="odi-match-groups-list">${fileGroupsHtml}</div>
        </div>`
    };
  }

  return {
    isMatch: false,
    files: [], sheets: [], matches: [],
    html: `
      <div class="odi-result-card miss" data-ruc="${odiEsc(cleanQ)}">
        <div class="odi-card-head">
          <span class="odi-card-ruc">${odiEsc(cleanQ)}</span>
          <span class="odi-badge-status miss"><i class="fa-solid fa-circle-xmark"></i> No encontrado</span>
        </div>
      </div>`
  };
}


function odiBuildRowKv(row) {
  const keys = Object.keys(row || {}).filter(k => String(row[k] ?? '').trim() !== '');
  return keys.length
    ? keys.map(k => `
        <div class="odi-kv-item">
          <span class="odi-kv-label">${odiEsc(k)}</span>
          <span class="odi-kv-val">${odiEsc(row[k])}</span>
        </div>`).join('')
    : `<p class="odi-card-note">Sin campos adicionales en este registro.</p>`;
}

function odiBuildSheetAccordion(ruc, entry, open) {
  const rows = entry.rows || [];
  const body = rows.length > 1
    ? rows.map((row, i) => `
        <div class="odi-accordion-subrow">
          <div class="odi-subrow-tag">Registro ${i + 1} de ${rows.length}</div>
          <div class="odi-kv-grid">${odiBuildRowKv(row)}</div>
        </div>`).join('')
    : `<div class="odi-kv-grid">${odiBuildRowKv(rows[0])}</div>`;

  const countBadge = rows.length > 1 ? `<span class="odi-sheet-count-tag">×${rows.length}</span>` : '';

  return `
    <div class="odi-sheet-accordion ${open ? 'open' : ''}">
      <div class="odi-accordion-header" onclick="odiToggleAccordion(this)">
        <div class="odi-acc-title-side">
          <i class="fa-regular fa-file-lines"></i>
          <span class="odi-acc-sheet-name">${odiEsc(entry.sheet)}</span>
          ${countBadge}
        </div>
        <button type="button" class="odi-acc-btn">
          <span class="odi-acc-btn-text">${open ? 'Ocultar campos' : 'Ver todos los campos'}</span>
          <i class="fa-solid fa-chevron-down odi-acc-chevron"></i>
        </button>
      </div>
      <div class="odi-accordion-content">${body}</div>
    </div>`;
}

function odiToggleAccordion(headerEl) {
  const acc = headerEl.closest('.odi-sheet-accordion');
  if (!acc) return;
  const isOpen = acc.classList.toggle('open');
  const txt = headerEl.querySelector('.odi-acc-btn-text');
  if (txt) {
    txt.textContent = isOpen ? 'Ocultar campos' : 'Ver todos los campos';
  }
}

// Filtrar tarjetas de resultado en vivo
function odiFilterResultCards(query) {
  const q = String(query || '').trim().toLowerCase();
  const cards = document.querySelectorAll('.odi-result-card');
  cards.forEach(c => {
    if (!q) {
      c.style.display = '';
      return;
    }
    const ruc = c.dataset.ruc ? c.dataset.ruc.toLowerCase() : '';
    const text = c.textContent.toLowerCase();
    if (ruc.includes(q) || text.includes(q)) {
      c.style.display = '';
    } else {
      c.style.display = 'none';
    }
  });
}

// ── Lista de Hojas en Columna 3 ─────────────────────────────────────────────
function odiRenderSheets(list) {
  if (!odiSheetsEl) return;
  if (!list.length) {
    odiSheetsEl.innerHTML = `<div class="odi-empty-bases-hint">No hay hojas disponibles</div>`;
    return;
  }
  odiSheetsEl.innerHTML = list.map(s => `
    <div class="odi-sheet-card-mini" title="${odiEsc(s.file_name)} · ${odiEsc(s.name)}">
      <i class="fa-regular fa-file-lines"></i>
      <div class="odi-sheet-mini-info">
        <span class="odi-sheet-mini-name">${odiEsc(s.name)}</span>
        <span class="odi-sheet-mini-file">${odiEsc(s.file_name)}</span>
      </div>
      <span class="odi-sheet-mini-count">${s.ruc_count}</span>
    </div>
  `).join('');
}

function odiFilterSheets(val) {
  const q = String(val || '').trim().toLowerCase();
  const filtered = q
    ? odiAllSheets.filter(s => s.name.toLowerCase().includes(q) || (s.file_name && s.file_name.toLowerCase().includes(q)))
    : odiAllSheets;
  odiRenderSheets(filtered);
}

// ── Modales: Subir Bases y Gestor de Repositorio ─────────────────────────────
function odiOpenManageModal() {
  if (odiManageModal) {
    odiManageModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
    odiRenderManageList();
  }
}

function odiCloseManageModal() {
  if (odiManageModal) {
    odiManageModal.classList.add('hidden');
    document.body.style.overflow = '';
  }
}

// Filtro dentro del modal de gestión (Todas vs Mis Bases)
function odiSetManageFilter(filter) {
  odiManageFilter = filter;
  const btnAll = document.getElementById('odiManageFilterAll');
  const btnMine = document.getElementById('odiManageFilterMine');
  if (btnAll) btnAll.classList.toggle('active', filter === 'all');
  if (btnMine) btnMine.classList.toggle('active', filter === 'mine');
  odiRenderManageList();
}

function odiRenderManageList() {
  if (!odiManageTableBody) return;

  const searchQ = odiManageSearchInput ? odiManageSearchInput.value.trim().toLowerCase() : '';

  const filtered = odiFiles.filter(f => {
    if (odiManageFilter === 'mine' && !f.is_mine) return false;
    if (searchQ) {
      const name = (f.filename || '').toLowerCase();
      const author = (f.uploaded_by || '').toLowerCase();
      if (!name.includes(searchQ) && !author.includes(searchQ)) return false;
    }
    return true;
  });

  if (!filtered.length) {
    odiManageTableBody.innerHTML = `
      <tr>
        <td colspan="7" class="odi-table-empty">
          <i class="fa-solid fa-folder-open"></i>
          <span>No se encontraron bases de datos en este filtro.</span>
        </td>
      </tr>`;
    return;
  }

  odiManageTableBody.innerHTML = filtered.map(f => {
    const authorBadge = f.is_mine
      ? `<span class="odi-author-badge mine"><i class="fa-solid fa-user-check"></i> Tú</span>`
      : `<span class="odi-author-badge other">${odiEsc(f.uploaded_by)}</span>`;

    const deleteBtn = f.can_delete ? `
      <button type="button" class="odi-btn-table-action del" title="Eliminar esta base de datos" data-id="${odiEsc(f.id)}" data-name="${odiEsc(f.filename)}" onclick="odiDeleteFile(this.dataset.id, this.dataset.name)">
        <i class="fa-solid fa-trash-can"></i>
      </button>` : `
      <span class="odi-btn-table-action disabled" title="Solo el autor o un administrador pueden eliminar esta base">
        <i class="fa-solid fa-lock"></i>
      </span>`;

    return `
      <tr>
        <td>
          <div class="odi-file-cell">
            <i class="fa-solid fa-file-excel"></i>
            <span class="odi-file-cell-name" title="${odiEsc(f.filename)}">${odiEsc(f.filename)}</span>
          </div>
        </td>
        <td>${authorBadge}</td>
        <td><span class="odi-date-cell">${odiEsc(f.uploaded_at)}</span></td>
        <td><b>${f.total_rucs.toLocaleString()}</b></td>
        <td>${f.sheets_count} hojas</td>
        <td>${f.size_kb} KB</td>
        <td style="text-align: right;">
          <div class="odi-actions-cell">
            <a href="/odi/api/files/${f.id}/download" class="odi-btn-table-action dl" title="Descargar archivo Excel original" download="${odiEsc(f.filename)}">
              <i class="fa-solid fa-download"></i>
            </a>
            ${deleteBtn}
          </div>
        </td>
      </tr>`;
  }).join('');
}

let odiPendingUploadFiles = [];

// POPUP 1: Selección / Arrastrar Archivo
function odiOpenUploadModal() {
  const uploadModal = document.getElementById('odiUploadModal');
  if (odiFileInput) odiFileInput.value = '';
  if (uploadModal) {
    uploadModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }
}

function odiCloseUploadModal() {
  const uploadModal = document.getElementById('odiUploadModal');
  if (uploadModal) {
    uploadModal.classList.add('hidden');
    document.body.style.overflow = '';
  }
  if (odiFileInput) odiFileInput.value = '';
}

// Al seleccionar o arrastrar archivo -> Cierra Popup 1 y abre Popup 2 (Confirmación)
function odiToast(message, type) {
  if (window.showToast) showToast(message, type);
}

function odiRejectionReason(f, maxBytes) {
  if (!f.name.toLowerCase().endsWith('.xlsx')) return `"${f.name}" no es .xlsx`;
  if (f.size === 0) return `"${f.name}" está vacío`;
  if (f.size > maxBytes) return `"${f.name}" supera los 25 MB`;
  return '';
}

function odiShowConfirmModal(validFiles) {
  const totalKb = validFiles.reduce((a, f) => a + f.size, 0) / 1024;
  const confirmName = document.getElementById('odiConfirmFileName');
  const confirmSize = document.getElementById('odiConfirmFileSize');
  if (confirmName) {
    confirmName.textContent = validFiles.length > 1
      ? `${validFiles.length} archivos (${validFiles.map(f => f.name).join(', ')})`
      : validFiles[0].name;
  }
  if (confirmSize) confirmSize.textContent = `${totalKb.toFixed(1)} KB`;

  // Abrir Popup 2 (Confirmación)
  const confirmModal = document.getElementById('odiConfirmModal');
  if (confirmModal) {
    confirmModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }
}

function odiHandleFiles(fileList) {
  if (!fileList || !fileList.length) return;

  const validFiles = [];
  const MAX_BYTES = 25 * 1024 * 1024;
  let rejected = '';
  for (const f of Array.from(fileList)) {
    const reason = odiRejectionReason(f, MAX_BYTES);
    if (reason) {
      rejected = rejected || reason;
    } else {
      validFiles.push(f);
    }
  }

  if (!validFiles.length) {
    odiToast(rejected || 'Solo se admiten archivos .xlsx', 'error');
    if (odiFileInput) odiFileInput.value = '';
    return;
  }
  if (rejected) odiToast(`${rejected}; se omitió`, 'warning');

  odiPendingUploadFiles = validFiles;

  // Cerrar Popup 1
  odiCloseUploadModal();
  odiShowConfirmModal(validFiles);
}

// POPUP 2: Cancelar Confirmación
function odiCancelConfirmModal() {
  const confirmModal = document.getElementById('odiConfirmModal');
  if (confirmModal) {
    confirmModal.classList.add('hidden');
    document.body.style.overflow = '';
  }
  odiPendingUploadFiles = [];
  if (odiFileInput) odiFileInput.value = '';
}

// POPUP 2 -> POPUP 3: Ejecutar subida y abrir Éxito
function odiSessionExpired() {
  odiToast('Tu sesión ha expirado. Por favor inicia sesión nuevamente.', 'error');
  setTimeout(() => { window.location.href = '/login'; }, 1500);
}

// Devuelve el resultado del servidor, o null si ya se avisó del problema al usuario.
async function odiReadUploadResponse(r) {
  if (!r.ok) {
    if (r.status === 401 || r.redirected) {
      odiSessionExpired();
      return null;
    }
    const errRes = await r.json().catch(() => ({}));
    odiToast(errRes.error || `Error (${r.status}) al subir archivo`, 'error');
    return null;
  }

  if (r.redirected || !(r.headers.get('content-type') || '').includes('application/json')) {
    odiSessionExpired();
    return null;
  }

  const res = await r.json();
  if (!res.success) {
    odiToast(res.error || 'Error al subir archivo', 'error');
    return null;
  }
  if (!res.processed_count) {
    odiToast(res.warning || 'No se pudo procesar ningún archivo', 'error');
    return null;
  }
  return res;
}

function odiShowSuccessModal(uploadedFileName, res) {
  const successName = document.getElementById('odiSuccessFileName');
  const successRucs = document.getElementById('odiSuccessRucCount');

  if (successName) successName.textContent = uploadedFileName;
  if (successRucs) {
    const uploaded = (res.files || []).find(f => f.filename === uploadedFileName && f.is_mine);
    successRucs.textContent = (uploaded ? uploaded.total_rucs : (res.total_rucs || 0)).toLocaleString();
  }

  const successModal = document.getElementById('odiSuccessModal');
  if (successModal) {
    successModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }
}

async function odiExecuteUpload() {
  if (!odiPendingUploadFiles.length) return;

  const uploadedFileName = odiPendingUploadFiles[0] ? odiPendingUploadFiles[0].name : 'Archivo Excel';
  const btnConfirm = document.getElementById('odiBtnConfirmUpload');
  const originalText = btnConfirm ? btnConfirm.innerHTML : 'Subir Archivo';

  if (btnConfirm) {
    btnConfirm.disabled = true;
    btnConfirm.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Subiendo...`;
  }

  const fd = new FormData();
  odiPendingUploadFiles.forEach(f => fd.append('files', f));

  try {
    const r = await fetch('/odi/api/upload', { method: 'POST', body: fd });
    const res = await odiReadUploadResponse(r);
    if (!res) return;

    odiApplyDataset(res);
    if (res.warning) odiToast(res.warning, 'warning');

    // Cerrar Popup 2 y abrir Popup 3 (Éxito)
    odiCancelConfirmModal();
    odiShowSuccessModal(uploadedFileName, res);
  } catch (e) {
    odiToast('Error al subir: ' + e.message, 'error');
  } finally {
    if (btnConfirm) {
      btnConfirm.disabled = false;
      btnConfirm.innerHTML = originalText;
    }
    odiPendingUploadFiles = [];
    if (odiFileInput) odiFileInput.value = '';
  }
}

// POPUP 3: Cerrar Éxito
function odiCloseSuccessModal() {
  const successModal = document.getElementById('odiSuccessModal');
  if (successModal) {
    successModal.classList.add('hidden');
    document.body.style.overflow = '';
  }
  if (odiRucList) {
    odiRucList.focus();
  }
}

// Eliminar archivo de base de datos
async function odiDeleteFile(fileId, fileName) {
  if (!confirm(`¿Deseas eliminar el archivo "${fileName}"?`)) {
    return;
  }

  try {
    const r = await fetch(`/odi/api/files/${encodeURIComponent(fileId)}`, { method: 'DELETE' });
    const res = await r.json().catch(() => ({}));

    if (!res.success) {
      if (window.showToast) showToast(res.error || 'No se pudo eliminar la base de datos', 'error');
      return;
    }

    odiApplyDataset(res);
    if (window.showToast) showToast(`Archivo "${fileName}" eliminado`, 'info');
  } catch (e) {
    if (window.showToast) showToast('Error al eliminar: ' + e.message, 'error');
  }
}

// Contador en vivo y búsqueda automática al escribir/pegar
let odiTypingTimer = null;
if (odiRucList) {
  odiRucList.addEventListener('input', () => {
    const n = new Set(odiRucList.value.split(/[\r\n,;\t]+/).map(v => v.trim()).filter(Boolean)).size;
    if (odiRucCount) odiRucCount.textContent = n ? `${n} elemento(s)` : '0 elementos';
    clearTimeout(odiTypingTimer);
    odiTypingTimer = setTimeout(odiClassifyList, 400);
  });
}

// Cerrar modales con tecla Escape
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    odiCloseUploadModal();
    odiCancelConfirmModal();
    odiCloseSuccessModal();
    odiCloseManageModal();
  }
});


// Drag & drop listeners para el modal de subida
if (odiDropzone) {
  odiDropzone.addEventListener('dragover', e => {
    e.preventDefault();
    odiDropzone.classList.add('odi-drag-over');
  });
  odiDropzone.addEventListener('dragleave', () => odiDropzone.classList.remove('odi-drag-over'));
  odiDropzone.addEventListener('drop', e => {
    e.preventDefault();
    odiDropzone.classList.remove('odi-drag-over');
    const files = e.dataTransfer.files;
    if (files && files.length) odiHandleFiles(files);
  });
}

// ── Exportación a Excel y PDF con todos los campos ──────────────────────────
function odiExportFilename(res, format) {
  const disp = res.headers.get('Content-Disposition');
  const match = disp ? disp.match(/filename="?([^"]+)"?/) : null;
  if (match && match[1]) return match[1];
  const ext = format === 'excel' ? 'xlsx' : 'pdf';
  return `Clasificacion_ODI.${ext}`;
}

function odiDownloadBlob(blob, filename) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
  a.remove();
}

async function odiExportResults(format) {
  if (!odiLastResults || !odiLastResults.length) {
    odiToast('No hay resultados clasificados para exportar', 'warning');
    return;
  }

  const btn = document.getElementById(format === 'excel' ? 'odiBtnExportExcel' : 'odiBtnExportPdf');
  const originalHtml = btn ? btn.innerHTML : '';

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Exportando...`;
  }

  try {
    const res = await fetch(`/odi/api/export/${format}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ results: odiLastResults })
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Error en el servidor (${res.status})`);
    }

    const filename = odiExportFilename(res, format);
    odiDownloadBlob(await res.blob(), filename);
    odiToast(`Reporte ${filename} descargado exitosamente ✓`, 'success');
  } catch (e) {
    console.error(`Error al exportar ${format}:`, e);
    odiToast(`Error al exportar a ${format.toUpperCase()}: ${e.message}`, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalHtml;
    }
  }
}

function odiEsc(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
