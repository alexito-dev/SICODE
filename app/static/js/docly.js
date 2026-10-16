/**
 * DOCLY MODULE - JAVASCRIPT CONTROLLER (UX/UI 2.0)
 */

let doclyCurrentFile = null;
let doclyTotalPages = 0;

document.addEventListener('DOMContentLoaded', () => {
    setupDoclyEvents();
});

function setupDoclyEvents() {
    const dz = document.getElementById('doclyDropzone');
    if (!dz) return;

    ['dragenter', 'dragover'].forEach(name => {
        dz.addEventListener(name, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dz.classList.add('drag-over');
        });
    });

    ['dragleave', 'drop'].forEach(name => {
        dz.addEventListener(name, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dz.classList.remove('drag-over');
        });
    });

    dz.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            doclyHandleFile(files[0]);
        }
    });
}

function doclyTriggerFileInput(e) {
    if (e.target.id !== 'doclyFileInput') {
        document.getElementById('doclyFileInput').click();
    }
}

async function doclyHandleFile(file) {
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!['pdf', 'docx', 'doc'].includes(ext)) {
        showDoclyToast('Formato no permitido. Sube un archivo PDF (.pdf) o Word (.docx)', true);
        return;
    }

    doclySetLoading(true);

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/docly/api/doc-info', {
            method: 'POST',
            body: formData
        });

        // Verificar autenticación o redirección de sesión
        if (response.status === 401 || response.redirected) {
            showDoclyToast('Tu sesión ha expirado. Redirigiendo al login...', true);
            setTimeout(() => window.location.href = '/login', 1500);
            return;
        }

        const data = await response.json();
        doclySetLoading(false);

        if (!data.success) {
            showDoclyToast(data.error || 'Error al procesar el archivo', true);
            return;
        }

        doclyCurrentFile = file;
        doclyTotalPages = data.total_pages;

        // Actualizar tarjeta del documento
        document.getElementById('doclyFileName').textContent = data.filename;
        document.getElementById('doclyPageCount').innerHTML = `<i class="fa-solid fa-layer-group"></i> ${data.total_pages} página(s)`;
        document.getElementById('doclyFileSize').innerHTML = `<i class="fa-solid fa-hard-drive"></i> ${data.size_kb} KB`;

        const badge = document.getElementById('doclyFileBadge');
        badge.textContent = ext.toUpperCase();
        if (ext === 'docx' || ext === 'doc') {
            badge.classList.add('docx');
        } else {
            badge.classList.remove('docx');
        }

        document.getElementById('doclyFileCard').classList.remove('hidden');
        document.getElementById('doclyConfigCard').classList.remove('hidden');
        document.getElementById('doclyPreviewCard').classList.remove('hidden');
        document.getElementById('doclyEmptyRight').classList.add('hidden');

        // Nombre base por defecto
        const baseName = data.filename.replace(/\.[^/.]+$/, "") + "_part_";
        document.getElementById('nombres').value = baseName;

        doclyUpdatePreview();
        showDoclyToast(`Documento listo: ${data.total_pages} página(s) detectadas ✓`);

        // Genera automáticamente los documentos, sin esperar un clic extra.
        doclyAutoProcess();

    } catch (err) {
        doclySetLoading(false);
        console.error('Error doclyHandleFile:', err);
        showDoclyToast('Error al leer el documento: ' + err.message, true);
    }
}

function doclySetLoading(isLoading) {
    const iconGroup = document.getElementById('doclyDzIcon');
    const spinner = document.getElementById('doclySpinner');

    if (isLoading) {
        if (iconGroup) iconGroup.style.display = 'none';
        if (spinner) spinner.style.display = 'block';
    } else {
        if (iconGroup) iconGroup.style.display = 'flex';
        if (spinner) spinner.style.display = 'none';
    }
}

function doclyClearFile() {
    doclyCurrentFile = null;
    doclyTotalPages = 0;
    doclyResetResults();
    document.getElementById('doclyFileInput').value = '';
    document.getElementById('doclyFileCard').classList.add('hidden');
    document.getElementById('doclyConfigCard').classList.add('hidden');
    document.getElementById('doclyPreviewCard').classList.add('hidden');
    document.getElementById('doclyEmptyRight').classList.remove('hidden');
}

function doclyAdjustPages(delta) {
    const input = document.getElementById('numPaginas');
    let val = (parseInt(input.value) || 1) + delta;
    if (val < 1) val = 1;
    input.value = val;
    doclyOnConfigChange();
}

function doclySetPreset(pagesPerDoc) {
    document.getElementById('numPaginas').value = pagesPerDoc;

    document.querySelectorAll('.docly-chip').forEach(btn => {
        btn.classList.toggle('active', btn.textContent.trim().startsWith(`${pagesPerDoc} Pág`));
    });

    doclyOnConfigChange();
}

// Se llama cada vez que el usuario toca un parámetro: actualiza el plan al
// instante y, tras una pequeña pausa, genera los documentos de verdad.
let doclyAutoTimer = null;

function doclyOnConfigChange() {
    doclyUpdatePreview();
    clearTimeout(doclyAutoTimer);
    doclyAutoTimer = setTimeout(doclyAutoProcess, 550);
}

function doclyEsc(s) {
    return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// Resultados ya generados y descargables (se invalidan si el usuario cambia los parámetros)
let doclyLastResults = null;

function doclyResetResults() {
    doclyLastResults = null;
    document.getElementById('doclyBtnDownloadAll').classList.add('hidden');
    document.getElementById('doclyBtnSaveFolder').classList.add('hidden');
    document.getElementById('previewCount').classList.remove('hidden');
    const title = document.getElementById('doclyPreviewTitle');
    title.classList.remove('generating');
    title.innerHTML = '<i class="fa-solid fa-list-check"></i> Plan de División';
}

function doclyUpdatePreview() {
    if (!doclyCurrentFile || doclyTotalPages <= 0) return;

    doclyResetResults();

    const numPaginas = parseInt(document.getElementById('numPaginas').value) || 1;
    const numDoc = parseInt(document.getElementById('numDoc').value) || 0;
    const nombres = document.getElementById('nombres').value || 'Documento_';
    const pagInicial = parseInt(document.getElementById('pagInicial').value) || 1;

    const grid = document.getElementById('doclyPreviewGrid');
    grid.innerHTML = '';

    let currentPage = Math.max(1, pagInicial);
    let docCount = 1;
    let totalGenerated = 0;

    while (currentPage <= doclyTotalPages) {
        if (numDoc > 0 && docCount > numDoc) break;

        const paginaFinal = Math.min(currentPage + numPaginas - 1, doclyTotalPages);
        const fileName = `${nombres}${docCount}.pdf`;
        const pagesCount = (paginaFinal - currentPage + 1);

        const item = document.createElement('div');
        item.className = 'docly-preview-item';
        item.innerHTML = `
            <span class="docly-item-title"><i class="fa-solid fa-file-pdf"></i> ${doclyEsc(fileName)}</span>
            <span class="docly-item-range">Páginas ${currentPage} - ${paginaFinal} (${pagesCount} pág${pagesCount > 1 ? 's' : ''})</span>
        `;
        grid.appendChild(item);

        docCount++;
        totalGenerated++;
        currentPage = paginaFinal + 1;
    }

    document.getElementById('previewCount').textContent = `${totalGenerated} PDF(s) a crear`;
}

/** Tarjetas "esqueleto" con efecto shimmer mientras se generan los documentos de verdad. */
function doclyRenderSkeleton(count) {
    const n = Math.max(1, Math.min(count || 3, 8));
    const grid = document.getElementById('doclyPreviewGrid');
    grid.innerHTML = Array.from({ length: n }).map(() => `
        <div class="docly-skeleton-card">
            <div class="docly-skeleton-icon"></div>
            <div class="docly-skeleton-lines">
                <div class="docly-skeleton-line w70"></div>
                <div class="docly-skeleton-line w40"></div>
            </div>
        </div>
    `).join('');
}

function doclyBase64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

function doclyTriggerDownload(base64, mime, filename) {
    const a = document.createElement('a');
    a.href = `data:${mime};base64,${base64}`;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
}

function doclyDownloadAll() {
    if (!doclyLastResults) return;
    doclyTriggerDownload(doclyLastResults.zip_base64, 'application/zip', doclyLastResults.zip_filename);
}

async function doclySaveFilesToFolder(dirHandle, files) {
    for (const f of files) {
        const bytes = doclyBase64ToBytes(f.base64);
        const fileHandle = await dirHandle.getFileHandle(f.filename, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(bytes);
        await writable.close();
    }
}

async function doclySaveAllToFolder() {
    if (!doclyLastResults) return;

    if (!('showDirectoryPicker' in window)) {
        showDoclyToast('Tu navegador no soporta selección directa de carpetas. Usa "Descargar todo (ZIP)".', true);
        return;
    }

    let dirHandle;
    try {
        dirHandle = await window.showDirectoryPicker();
    } catch (err) {
        if (err.name !== 'AbortError') {
            showDoclyToast('No se pudo acceder a la carpeta seleccionada', true);
        }
        return; // Cancelado por el usuario
    }

    try {
        await doclySaveFilesToFolder(dirHandle, doclyLastResults.files);
        showDoclyToast(`¡${doclyLastResults.files.length} archivo(s) guardado(s) en tu carpeta! 📁`);
    } catch (err) {
        console.error('Error doclySaveAllToFolder:', err);
        showDoclyToast('Error al guardar los archivos en la carpeta', true);
    }
}

// Un botón de descarga por tarjeta (delegado, evita inyectar nombres de archivo en atributos onclick)
document.getElementById('doclyPreviewGrid').addEventListener('click', (e) => {
    const btn = e.target.closest('.docly-btn-download-item');
    if (!btn || !doclyLastResults) return;
    const idx = parseInt(btn.dataset.idx, 10);
    const f = doclyLastResults.files[idx];
    if (f) doclyTriggerDownload(f.base64, 'application/pdf', f.filename);
});

function doclyRenderResults(data) {
    doclyLastResults = data;

    const grid = document.getElementById('doclyPreviewGrid');
    grid.innerHTML = data.files.map((f, i) => `
        <div class="docly-preview-item docly-result-item">
            <div>
                <span class="docly-item-title"><i class="fa-solid fa-file-pdf"></i> ${doclyEsc(f.filename)}</span>
                <span class="docly-item-range">Páginas ${f.page_start} - ${f.page_end} (${f.pages} pág${f.pages > 1 ? 's' : ''}) · ${f.size_kb} KB</span>
            </div>
            <button type="button" class="docly-btn-download-item" data-idx="${i}" title="Descargar ${doclyEsc(f.filename)}">
                <i class="fa-solid fa-download"></i>
            </button>
        </div>
    `).join('');

    const countEl = document.getElementById('previewCount');
    countEl.classList.remove('hidden');
    countEl.textContent = `${data.files.length} PDF(s) generado(s)`;

    const title = document.getElementById('doclyPreviewTitle');
    title.classList.remove('generating');
    title.innerHTML = '<i class="fa-solid fa-circle-check"></i> Documentos Generados';

    document.getElementById('doclyBtnDownloadAll').classList.remove('hidden');
    document.getElementById('doclyBtnSaveFolder').classList.toggle('hidden', !('showDirectoryPicker' in window));
}

/** Llama al backend y genera los sub-documentos con los parámetros actuales. */
async function doclyRunSubdivide() {
    const numPaginas = parseInt(document.getElementById('numPaginas').value) || 1;
    const numDoc = parseInt(document.getElementById('numDoc').value) || 0;
    const nombres = document.getElementById('nombres').value || 'Documento_';
    const pagInicial = parseInt(document.getElementById('pagInicial').value) || 1;

    const formData = new FormData();
    formData.append('file', doclyCurrentFile);
    formData.append('num_paginas', numPaginas);
    formData.append('num_doc', numDoc);
    formData.append('nombres', nombres);
    formData.append('pag_inicial', pagInicial);

    const response = await fetch('/docly/api/subdividir', { method: 'POST', body: formData });

    if (response.status === 401 || response.redirected) {
        showDoclyToast('Tu sesión ha expirado. Redirigiendo al login...', true);
        setTimeout(() => window.location.href = '/login', 1500);
        throw new Error('Sesión expirada');
    }

    const data = await response.json();
    if (!data.success) {
        throw new Error(data.error || 'Error al subdividir el documento');
    }
    return data;
}

let doclyProcessing = false;

/** Genera los documentos solo, sin ningún clic: al subir el archivo o cambiar un parámetro. */
async function doclyAutoProcess() {
    if (!doclyCurrentFile || doclyProcessing) return;
    doclyProcessing = true;

    document.getElementById('previewCount').classList.add('hidden');
    const title = document.getElementById('doclyPreviewTitle');
    title.classList.add('generating');
    title.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Generando documentos...';

    // Muestra tarjetas "esqueleto" (con el mismo número que el plan actual)
    // en vez de dejar la grilla en blanco o con contenido desactualizado.
    doclyRenderSkeleton(document.getElementById('doclyPreviewGrid').children.length);

    try {
        const data = await doclyRunSubdivide();
        doclyRenderResults(data);
    } catch (err) {
        console.error('Error doclyAutoProcess:', err);
        showDoclyToast(err.message || 'Error al generar los documentos', true);
        doclyUpdatePreview();
    } finally {
        doclyProcessing = false;
    }
}

function showDoclyToast(msg, isError = false) {
    const toast = document.getElementById('sicodeToast');
    if (!toast) return;

    toast.textContent = msg;
    toast.style.background = isError ? '#ef4444' : '#1e2939';
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 2800);
}
