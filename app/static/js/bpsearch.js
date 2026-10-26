// BPSearch - Logic

document.addEventListener('DOMContentLoaded', () => {
    // --- Elements ---
    // We toggle display of this container
    const folderDisplay = document.getElementById('folder-display');
    const folderPathSpan = document.getElementById('folder-path');

    // progress details hidden by default but updated
    const pdfCountSpan = document.getElementById('pdf-count') || { textContent: '' }; // Fallback mainly

    const searchInput = document.getElementById('search-input');
    const btnSearch = document.getElementById('btn-search');

    // Pills logic replaced chkSubcarpetas
    const pillSubcarpetas = document.getElementById('pill-subcarpetas');

    const chkIgnorarHistorial = document.getElementById('chk-ignorar-historial');

    const progressContainer = document.getElementById('progressContainer');
    const progressText = document.getElementById('progressText');
    const progressPercent = document.getElementById('progressPercent');
    const progressFill = document.getElementById('progressFill');
    const progressFiles = document.getElementById('progress-files');
    const progressTime = document.getElementById('progress-time');

    const btnPause = document.getElementById('btn-pause');
    const btnCancel = document.getElementById('btn-cancel');

    // Results logic (List Container)
    const resultsCount = document.getElementById('results-count');
    const resultsList = document.getElementById('results-list');

    const exportSection = document.getElementById('export-section');
    const btnExportPdf = document.getElementById('btn-export-pdf');
    const btnExportExcel = document.getElementById('btn-export-excel');
    const btnExportTxt = document.getElementById('btn-export-txt');

    // --- State ---
    let currentFolder = "";
    let isSearching = false;
    let isPaused = false;
    let eventSource = null;

    // ===== TOAST NOTIFICATIONS (Copied from DocCheck) =====
    function showToast(type, message) {
        if (window.showToast) {
            window.showToast(message, type);
        }
    }


    // --- Client-Side Folder Upload Logic ---
    const btnSelectFolderNative = document.getElementById('btn-select-folder-native');
    const nativeDirInput = document.getElementById('native-dir-input');

    if (btnSelectFolderNative && nativeDirInput) {
        btnSelectFolderNative.addEventListener('click', () => {
            nativeDirInput.click();
        });

        nativeDirInput.addEventListener('change', async (e) => {
            const files = e.target.files;
            if (!files || files.length === 0) return;

            // Filtrar solo los PDF (ignorando mayúsculas/minúsculas)
            const pdfFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.pdf'));

            if (pdfFiles.length === 0) {
                showToast('warning', 'La carpeta seleccionada no contiene archivos PDF.');
                e.target.value = ""; // reset
                return;
            }

            const originalText = btnSelectFolderNative.innerHTML;
            btnSelectFolderNative.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Subiendo al servidor...';
            btnSelectFolderNative.disabled = true;
            
            // Subir usando FormData
            const formData = new FormData();
            pdfFiles.forEach(file => {
                formData.append('files', file);
            });

            try {
                // Al subir muchos archivos pesados puede tardar
                showToast('info', `Subiendo ${pdfFiles.length} PDFs, por favor espera...`);
                
                const res = await fetch('/bpsearch/upload-pdfs', {
                    method: 'POST',
                    body: formData
                });
                const data = await res.json();

                if (data.status === 'success') {
                    // Guardamos la ruta temporal donde el servidor guardó los PDFs
                    currentFolder = data.session_path;
                    
                    folderPathSpan.textContent = `${data.pdfs_total} Archivos Listos`;
                    
                    if (pdfCountSpan) pdfCountSpan.textContent = `${data.pdfs_carpeta} PDFs`;

                    btnSelectFolderNative.style.display = 'none';
                    folderDisplay.style.display = 'block';
                    
                    showToast('success', `${data.pdfs_total} PDFs cargados y listos para la búsqueda.`);
                } else {
                    showToast('error', data.message || 'Error al cargar los archivos');
                }
            } catch (err) {
                console.error("Error upload:", err);
                showToast('error', 'Ocurrió un error de red al subir los archivos.');
            } finally {
                btnSelectFolderNative.innerHTML = originalText;
                btnSelectFolderNative.disabled = false;
                e.target.value = ""; // Limpiar input para permitir seleccionar de nuevo la misma
            }
        });
    }

    // Modal elements no longer used in this approach, safely ignored
    if (document.getElementById('modal-close')) {
         document.getElementById('modal-close').addEventListener('click', () => {
             document.getElementById('folder-modal').style.display = 'none';
         });
    }

    // --- Search Logic ---
    btnSearch.addEventListener('click', startSearch);
    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') startSearch();
    });

    async function startSearch() {
        if (isSearching) return;

        const palabra = searchInput.value.trim();
        if (!palabra) {
            showToast('warning', 'Ingresa una palabra para buscar');
            return;
        }

        if (!currentFolder) {
            showToast('warning', 'Selecciona una carpeta primero');
            return;
        }

        // Determine mode from pills
        const incluirSubcarpetas = pillSubcarpetas.classList.contains('active');

        // Reset UI
        resultsList.innerHTML = '';
        resultsCount.textContent = '0';
        progressContainer.classList.remove('completed', 'error');
        progressContainer.classList.add('processing');
        progressFill.style.width = '0%';
        progressText.textContent = 'Buscando con cuidado...';
        progressPercent.textContent = '0%';
        progressFiles.textContent = '0/0 archivos';
        progressTime.textContent = '0s transcurridos';
        
        // Enable controls instead of showing
        btnPause.disabled = false;
        btnCancel.disabled = false;
        btnPause.innerHTML = '<i class="fa-solid fa-pause"></i> Pausar';

        isSearching = true;
        isPaused = false;

        try {
            // Trigger search start
            const res = await fetch('/bpsearch/search', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    carpeta: currentFolder,
                    palabra: palabra,
                    incluir_subcarpetas: incluirSubcarpetas,
                    ignorar_historial: chkIgnorarHistorial.checked
                })
            });

            const data = await res.json();
            if (data.status !== 'started') {
                showToast('error', data.message || 'Error iniciando búsqueda');
                isSearching = false;
                return;
            }

            // Connect to SSE
            if (eventSource) eventSource.close();
            eventSource = new EventSource('/bpsearch/search-stream');

            eventSource.onmessage = (e) => {
                const msg = JSON.parse(e.data);
                handleSearchEvent(msg);
            };

            eventSource.onerror = () => {
                if (isSearching) {
                    // Don't toast on clean close, but here likely error
                    // showToast('Conexión perdida con el servidor', 'error');
                    // stopSearchUI(false);
                }
                eventSource.close();
            };

        } catch (e) {
            showToast('error', 'Error iniciando búsqueda: ' + e.message);
            isSearching = false;
        }
    }

    function handleSearchEvent(msg) {
        if (msg.type === 'progress') {
            updateProgress(msg);
        } else if (msg.type === 'found') {
            addResult(msg);
        } else if (msg.type === 'complete') {
            finishSearch(msg);
        } else if (msg.type === 'cancelled') {
            cancelSearchUI();
        } else if (msg.type === 'error') {
            console.error(msg.error);
        }
    }

    function updateProgress(data) {
        const pct = data.percent;
        progressFill.style.width = `${pct}%`;
        progressPercent.textContent = `${pct}%`;
        progressText.textContent = `Procesando: ${data.archivo}`;
        progressFiles.textContent = `${data.current}/${data.total} archivos`;
        progressTime.textContent = `${data.tiempo_transcurrido}s transcurridos`;
    }

    function addResult(data) {
        const div = document.createElement('div');
        div.className = 'result-item'; // CSS handles this

        div.innerHTML = `
            <div class="result-item-title">${data.archivo}</div>
            <div class="result-item-details">
                <span class="result-highlight">${data.veces} coincidencias</span> 
                | Pág: ${data.paginas.join(', ')}
            </div>
        `;
        resultsList.appendChild(div);

        // Update count
        const currentCount = parseInt(resultsCount.textContent) || 0;
        resultsCount.textContent = `${currentCount + 1}`;
    }

    function finishSearch(data) {
        stopSearchUI(true);
        progressText.textContent = data.message;
        progressPercent.textContent = "100%";
        progressFill.style.width = "100%";
        showToast('success', `Búsqueda completada. ${data.encontrados} encontrados.`);

        if (data.encontrados > 0) {
            exportSection.style.display = 'block';
        }
    }

    function cancelSearchUI() {
        stopSearchUI(false);
        progressText.textContent = "Cancelado por el usuario";
        progressContainer.classList.add('error');
        showToast('info', 'Búsqueda cancelada');
    }

    function stopSearchUI(completed) {
        isSearching = false;
        if (eventSource) {
            eventSource.close();
            eventSource = null;
        }
        
        // Disable controls instead of hiding
        btnPause.disabled = true;
        btnCancel.disabled = true;
        
        progressContainer.classList.remove('processing');
        if (completed) progressContainer.classList.add('completed');
    }

    // --- Controls ---
    btnPause.addEventListener('click', async () => {
        if (!isSearching) return;

        const url = isPaused ? '/bpsearch/resume' : '/bpsearch/pause';
        await fetch(url, { method: 'POST' });

        isPaused = !isPaused;
        if (isPaused) {
            btnPause.innerHTML = '<i class="fa-solid fa-play"></i> Reanudar';
            progressText.textContent = "Búsqueda pausada";
            progressContainer.classList.remove('processing');
            showToast('info', 'Búsqueda pausada');
        } else {
            btnPause.innerHTML = '<i class="fa-solid fa-pause"></i> Pausar';
            progressText.textContent = "Continuando búsqueda...";
            progressContainer.classList.add('processing');
            showToast('info', 'Búsqueda reanudada');
        }
    });

    btnCancel.addEventListener('click', async () => {
        if (!isSearching) return;
        await fetch('/bpsearch/cancel', { method: 'POST' });
    });

    // --- Exports ---
    const getSearchWord = () => searchInput.value.trim() || "busqueda";

    btnExportPdf.addEventListener('click', () => {
        window.location.href = `/bpsearch/export/pdf?palabra=${encodeURIComponent(getSearchWord())}`;
        showToast('info', 'Exportando a PDF...');
    });

    btnExportExcel.addEventListener('click', () => {
        window.location.href = `/bpsearch/export/excel?palabra=${encodeURIComponent(getSearchWord())}`;
        showToast('info', 'Exportando a Excel...');
    });

    btnExportTxt.addEventListener('click', () => {
        window.location.href = `/bpsearch/export/txt?palabra=${encodeURIComponent(getSearchWord())}`;
        showToast('info', 'Exportando a TXT...');
    });

    // Fix history clear for central folder reference
    if (document.getElementById('btn-clear-history')) {
        document.getElementById('btn-clear-history').addEventListener('click', async (e) => {
            e.preventDefault();
            const palabra = searchInput.value.trim();
            if (!palabra) {
                showToast('warning', 'Ingresa una palabra para borrar su historial');
                return;
            }

            if (!currentFolder) {
                showToast('warning', 'Selecciona una carpeta primero');
                return;
            }

            if (confirm(`¿Seguro que deseas borrar el historial de búsqueda para "${palabra}" en ${currentFolder}?`)) {
                try {
                    const res = await fetch('/bpsearch/clear-history', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ palabra: palabra })
                    });
                    const data = await res.json();
                    showToast('success', data.message || 'Historial borrado');
                } catch (e) {
                    console.error('Error al borrar historial:', e);
                    showToast('error', 'Error al borrar historial');
                }
            }
        });
    }
});
