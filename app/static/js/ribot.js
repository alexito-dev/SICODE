/**
 * ribot.js - Lógica limpia, intuitiva y robusta para el Módulo RIbot (SICODE)
 * Carga directa de las 6 plantillas oficiales Word desde D:\SICODE\APPS\RIBOT
 */

(function () {
    let selectedFiles = []; // Lista de objetos { id, file, modelFilename, texto, editando }
    let availableModels = []; // Lista de objetos { filename, name, size }
    let selectedGlobalModel = '';
    let statusInterval = null;

    // Mensajes de animación del loader
    const loaderMessages = [
        "Iniciando motor de redacción SUNAT...",
        "Abriendo Fichas de Valor PDF...",
        "Extrayendo RUC, Razón Social e Intendencia...",
        "Buscando Orden de Pago, Periodo y Tributos...",
        "Cargando plantilla Word (.docx) seleccionada...",
        "Sustituyendo marcadores {{RUC}}, {{RAZON_SOCIAL}}, {{NRO_RES}}...",
        "Escribiendo los considerandos normativos...",
        "Reconstruyendo tablas oficiales de tributos...",
        "Insertando el SE RESUELVE institucional...",
        "Guardando archivos de Resolución Word (.docx)...",
        "Empaquetando resultados..."
    ];

    // Elementos DOM
    let dropzone, fileInput, fileCard, fileList, queueCount, processBtn;
    let advancedToggle, advancedBox, destPathInput;
    let emptyRight, resultCard, resultAlert, resultTitleText, resultStats;
    let downloadIcon, downloadName, downloadDesc, downloadBtn;
    let overlay, loaderStatus;
    let evalBox, evalToggle, evalSub, evalBody, evalText, evalInfo, evalClear;

    // Elementos DOM de Modelos
    let modelsGrid, syncBtn, openFolderBtn;

    document.addEventListener("DOMContentLoaded", () => {
        initElements();
        initEvents();
        loadAvailableModels();
    });

    function initElements() {
        dropzone = document.getElementById('ribotDropzone');
        fileInput = document.getElementById('ribotFileInput');
        fileCard = document.getElementById('ribotFileCard');
        fileList = document.getElementById('ribotFileList');
        queueCount = document.getElementById('ribotQueueCount');
        processBtn = document.getElementById('ribotProcessBtn');

        advancedToggle = document.getElementById('ribotAdvancedToggle');
        advancedBox = document.getElementById('ribotAdvancedBox');
        destPathInput = document.getElementById('ribotDestPath');

        emptyRight = document.getElementById('ribotEmptyRight');
        resultCard = document.getElementById('ribotResultCard');
        resultAlert = document.getElementById('ribotResultAlert');
        resultTitleText = document.getElementById('ribotResultTitleText');
        resultStats = document.getElementById('ribotResultStats');

        downloadIcon = document.getElementById('ribotDownloadIcon');
        downloadName = document.getElementById('ribotDownloadName');
        downloadDesc = document.getElementById('ribotDownloadDesc');
        downloadBtn = document.getElementById('ribotDownloadBtn');

        overlay = document.getElementById('ribotOverlay');
        loaderStatus = document.getElementById('ribotLoaderStatus');

        evalBox = document.getElementById('ribotEval');
        evalToggle = document.getElementById('ribotEvalToggle');
        evalSub = document.getElementById('ribotEvalSub');
        evalBody = document.getElementById('ribotEvalBody');
        evalText = document.getElementById('ribotEvalText');
        evalInfo = document.getElementById('ribotEvalInfo');
        evalClear = document.getElementById('ribotEvalClear');

        // Botones y Grid de Modelos
        modelsGrid = document.getElementById('ribotModelsGrid');
        syncBtn = document.getElementById('ribotSyncModelsBtn');
        openFolderBtn = document.getElementById('ribotOpenFolderBtn');
    }

    function initEvents() {
        if (!dropzone) return;

        // Abrir selector de PDFs al presionar la dropzone
        dropzone.addEventListener('click', () => fileInput.click());

        // Eventos Drag & Drop de PDFs
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });

        dropzone.addEventListener('dragleave', () => {
            dropzone.classList.remove('dragover');
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
        });

        // Redacción de la evaluación: interruptor y texto general
        if (evalToggle) {
            evalToggle.addEventListener('change', () => {
                renderEvalState();
                renderFileList();
                if (evalToggle.checked && evalText) evalText.focus();
            });
            evalText.addEventListener('input', renderEvalState);
            evalClear.addEventListener('click', () => {
                evalText.value = '';
                renderEvalState();
                evalText.focus();
            });
        }

        // Alternar opciones avanzadas
        if (advancedToggle) {
            advancedToggle.addEventListener('click', () => {
                advancedBox.classList.toggle('active');
                const isExpanded = advancedBox.classList.contains('active');
                advancedToggle.innerHTML = isExpanded ?
                    '<i class="fa-solid fa-chevron-up"></i> Ocultar Opciones Avanzadas' :
                    '<i class="fa-solid fa-sliders"></i> Opciones Avanzadas (Guardar en Carpeta Específica)';
            });
        }

        // Botón Sincronizar
        if (syncBtn) {
            syncBtn.addEventListener('click', () => {
                const icon = syncBtn.querySelector('i');
                if (icon) icon.classList.add('fa-spin');
                loadAvailableModels(true).finally(() => {
                    if (icon) setTimeout(() => icon.classList.remove('fa-spin'), 500);
                });
            });
        }

        // Botón Abrir Carpeta Física
        if (openFolderBtn) {
            openFolderBtn.addEventListener('click', async () => {
                try {
                    const r = await fetch('/ribot/api/open-templates-folder', { method: 'POST' });
                    const res = await r.json();
                    if (res.status === 'success') {
                        if (window.showToast) window.showToast('Carpeta de plantillas abierta en tu PC', 'info');
                    } else {
                        throw new Error(res.message || 'No se pudo abrir la carpeta');
                    }
                } catch (e) {
                    if (window.showToast) window.showToast('Error al abrir carpeta: ' + e.message, 'error');
                }
            });
        }
    }

    // ── GESTIÓN DE PLANTILLAS / MODELOS WORD ──────────────────────────────
    async function loadAvailableModels(showNotification = false) {
        if (!modelsGrid) return;
        try {
            const res = await fetch('/ribot/api/models');
            if (!res.ok) throw new Error('Error cargando modelos.');
            const data = await res.json();

            if (data.status === 'success' && Array.isArray(data.models)) {
                availableModels = data.models;

                // Seleccionar primer modelo por defecto si aún no hay uno elegido
                if (availableModels.length > 0 && !selectedGlobalModel) {
                    selectedGlobalModel = availableModels[0].filename;
                }

                renderModelsGrid();
                updateAllFileModelSelectors();

                if (showNotification && window.showToast) {
                    window.showToast(`Sincronizado: ${availableModels.length} plantilla(s) lista(s)`, 'success');
                }
            }
        } catch (e) {
            console.error('Error cargando plantillas:', e);
            modelsGrid.innerHTML = `<div class="ribot-model-error"><i class="fa-solid fa-triangle-exclamation"></i> Error al cargar las plantillas de D:\\SICODE\\APPS\\RIBOT</div>`;
        }
    }

    function renderModelsGrid() {
        if (!modelsGrid) return;
        modelsGrid.innerHTML = '';

        if (availableModels.length === 0) {
            modelsGrid.innerHTML = `<div class="ribot-model-empty"><i class="fa-solid fa-folder-open"></i> No se encontraron plantillas .docx en D:\\SICODE\\APPS\\RIBOT</div>`;
            return;
        }

        availableModels.forEach(m => {
            const isSelected = m.filename === selectedGlobalModel;
            const card = document.createElement('div');
            card.className = `ribot-model-card ${isSelected ? 'active' : ''}`;

            card.innerHTML = `
                <div class="ribot-model-card-icon"><i class="fa-solid fa-file-word"></i></div>
                <span class="ribot-model-card-name" title="${escapeHtml(m.name)}">${escapeHtml(m.name)}</span>
                <div class="ribot-model-card-check"><i class="fa-solid fa-circle-check"></i></div>
            `;

            card.onclick = () => {
                selectedGlobalModel = m.filename;
                renderModelsGrid();
                // Actualizar modelo predeterminado de todos los PDFs en cola
                selectedFiles.forEach(item => {
                    item.modelFilename = selectedGlobalModel;
                });
                renderFileList();
            };

            modelsGrid.appendChild(card);
        });
    }

    // ── GESTIÓN DE ARCHIVOS PDF SELECCIONADOS ──────────────────────────────
    function formatBytes(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + ['Bytes', 'KB', 'MB'][i];
    }

    function handleFiles(files) {
        const pdfFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.pdf') || f.type === 'application/pdf');

        const rejected = files.length - pdfFiles.length;
        if (rejected > 0 && window.showToast) {
            window.showToast(`Se omitieron ${rejected} archivo(s) por no ser formato PDF.`, 'warning');
        }

        pdfFiles.forEach(file => {
            if (!selectedFiles.find(item => item.file.name === file.name)) {
                selectedFiles.push({
                    id: 'pdf_' + Math.random().toString(36).substr(2, 9),
                    file: file,
                    modelFilename: selectedGlobalModel,
                    texto: '',
                    editando: false
                });
            }
        });

        renderFileList();
        updateUI();
    }

    function renderFileList() {
        if (!fileList) return;
        fileList.innerHTML = '';
        selectedFiles.forEach(item => {
            addFileToUI(item);
        });
    }

    function addFileToUI(item) {
        const file = item.file;
        const div = document.createElement('div');
        div.className = 'ribot-file-item';

        // Opciones del selector individual de modelo por PDF
        let selectOptionsHtml = '';
        availableModels.forEach(m => {
            const isSel = m.filename === item.modelFilename;
            selectOptionsHtml += `<option value="${escapeHtml(m.filename)}" ${isSel ? 'selected' : ''}>${escapeHtml(m.name)}</option>`;
        });

        div.innerHTML = `
            <div class="ribot-file-info">
                <div class="ribot-file-icon"><i class="fa-solid fa-file-pdf"></i></div>
                <div style="min-width: 0; flex: 1;">
                    <div class="ribot-file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</div>
                    <div class="ribot-file-sub-row">
                        <span class="ribot-file-size">${formatBytes(file.size)}</span>
                        <div class="ribot-file-model-picker">
                            <i class="fa-solid fa-file-word" style="color: var(--primary-blue, #0163ac); font-size: 11px;"></i>
                            <select class="ribot-file-model-select" title="Seleccionar modelo para este PDF">
                                ${selectOptionsHtml}
                            </select>
                        </div>
                    </div>
                </div>
            </div>
            <div class="ribot-file-actions">
                ${evalOn() ? `<button type="button" class="ribot-btn-texto ${item.texto.trim() ? 'has-text' : ''}" title="Escribir un texto diferente para este PDF"><i class="fa-solid fa-pen"></i></button>` : ''}
                <button type="button" class="ribot-btn-remove" title="Quitar de la lista">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>
            ${evalOn() && item.editando ? `<div class="ribot-file-texto"><textarea class="ribot-file-texto-input" rows="3" maxlength="5000" placeholder="Texto solo para este PDF. Si lo dejas vacío, se usa el texto general.">${escapeHtml(item.texto)}</textarea></div>` : ''}
        `;

        const textoBtn = div.querySelector('.ribot-btn-texto');
        if (textoBtn) {
            textoBtn.onclick = () => {
                item.editando = !item.editando;
                renderFileList();
            };
        }
        const textoInput = div.querySelector('.ribot-file-texto-input');
        if (textoInput) {
            textoInput.oninput = (e) => {
                item.texto = e.target.value;
                renderEvalState();
                const btn = div.querySelector('.ribot-btn-texto');
                if (btn) btn.classList.toggle('has-text', item.texto.trim() !== '');
            };
        }

        const modelSelectEl = div.querySelector('.ribot-file-model-select');
        modelSelectEl.onchange = (e) => {
            item.modelFilename = e.target.value;
        };

        div.querySelector('.ribot-btn-remove').onclick = (e) => {
            e.stopPropagation();
            removeFile(item.id);
        };

        fileList.appendChild(div);
    }

    // ── REDACCIÓN DE LA EVALUACIÓN ────────────────────────────────────────
    function evalOn() {
        return !!(evalToggle && evalToggle.checked);
    }

    function renderEvalState() {
        if (!evalToggle) return;
        const on = evalOn();
        evalBox.classList.toggle('on', on);
        evalBody.style.display = on ? 'block' : 'none';
        evalSub.textContent = on ? 'Escribe el texto que deseas agregar a la resolución.' : 'Desactivado.';
        evalClear.style.display = on && evalText.value.trim() ? 'inline' : 'none';

        const propios = selectedFiles.filter(item => item.texto.trim()).length;
        evalInfo.textContent = on && propios
            ? `${propios} PDF usa${propios > 1 ? 'n' : ''} un texto diferente.`
            : '';
    }

    function updateAllFileModelSelectors() {
        renderFileList();
    }

    function removeFile(id) {
        selectedFiles = selectedFiles.filter(item => item.id !== id);
        renderFileList();
        updateUI();
    }

    window.ribotClearAllFiles = function () {
        selectedFiles = [];
        if (fileList) fileList.innerHTML = '';
        updateUI();
    };

    function updateUI() {
        const hasFiles = selectedFiles.length > 0;
        if (fileCard) fileCard.style.display = hasFiles ? 'block' : 'none';
        if (queueCount) queueCount.textContent = selectedFiles.length;
        if (processBtn) processBtn.disabled = !hasFiles;

        if (hasFiles) {
            ribotResetResults();
        }
    }

    // ── ACCIÓN PRINCIPAL: PROCESAR PDFS ────────────────────────────────────
    window.ribotProcessFiles = async function () {
        if (selectedFiles.length === 0) return;

        const destinationPath = destPathInput ? destPathInput.value.trim() : '';
        const formData = new FormData();
        
        const fileModelsMap = {};
        selectedFiles.forEach(item => {
            formData.append('files', item.file);
            fileModelsMap[item.file.name] = item.modelFilename || selectedGlobalModel;
        });

        formData.append('destination_path', destinationPath);
        formData.append('model_name', selectedGlobalModel);
        formData.append('file_models', JSON.stringify(fileModelsMap));

        if (evalOn()) {
            const textosMap = {};
            selectedFiles.forEach(item => {
                if (item.texto.trim()) textosMap[item.file.name] = item.texto;
            });
            formData.append('texto_libre', evalText.value);
            formData.append('file_textos', JSON.stringify(textosMap));
        }

        startLoading();

        try {
            const r = await fetch('/ribot/api/process-upload', {
                method: 'POST',
                body: formData
            });

            if (!r.ok) {
                const errData = await r.json().catch(() => ({}));
                throw new Error(errData.message || 'Error de comunicación con RIbot.');
            }

            const res = await r.json();
            stopLoading();

            if (res.status === 'success') {
                showResults(res);
                if (res.avisos && res.avisos.length) {
                    if (window.showToast) window.showToast('Generado con avisos: revisa el detalle.', 'warning');
                } else if (window.showToast) window.showToast('¡Resoluciones generadas con éxito!', 'success');
            } else {
                throw new Error(res.message || 'Error al redactar documentos.');
            }

        } catch (e) {
            stopLoading();
            if (window.showToast) window.showToast(e.message, 'error');
            else alert(e.message);
        }
    };

    // RENDERIZAR RESULTADOS
    function showResults(res) {
        if (emptyRight) emptyRight.style.display = 'none';
        if (resultCard) resultCard.style.display = 'block';

        if (downloadBtn) downloadBtn.href = res.download_url;

        if (res.saved_locally && res.physical_path) {
            resultStats.innerHTML = `Se procesaron <strong>${res.count}</strong> Ficha(s) de Valor PDF.<br>` +
                `<span style="font-weight: 600; color: #16a34a; display: block; margin-top: 6px;">` +
                `<i class="fa-solid fa-folder-open"></i> Guardado en tu PC:<br>` +
                `<code style="background: rgba(22, 163, 74, 0.08); padding: 3px 8px; border-radius: 6px; font-family: monospace; font-size: 11.5px; color: #15803d; word-break: break-all; display: inline-block; margin-top: 4px; border: 1px solid rgba(22, 163, 74, 0.2);">${escapeHtml(res.physical_path)}</code></span>`;
        } else {
            resultStats.textContent = `Se redactaron ${res.count} Resolución(es) de Intendencia correctamente.`;
        }

        if (res.avisos && res.avisos.length) {
            resultStats.innerHTML += `<div style="margin-top: 10px; padding: 8px 10px; border-radius: 8px; background: rgba(217, 119, 6, 0.08); border: 1px solid rgba(217, 119, 6, 0.3); color: #b45309; font-size: 12px; text-align: left;">` +
                `<strong><i class="fa-solid fa-triangle-exclamation"></i> Revisar:</strong><ul style="margin: 4px 0 0 16px; padding: 0;">` +
                res.avisos.map(a => `<li>${escapeHtml(a)}</li>`).join('') + `</ul></div>`;
        }

        if (res.type === 'single') {
            resultTitleText.textContent = "¡Resolución Generada!";
            downloadIcon.className = "ribot-dl-icon";
            downloadIcon.innerHTML = '<i class="fa-solid fa-file-word"></i>';
            downloadName.textContent = res.filename;
            downloadDesc.textContent = "Resolución Word (.docx) redactada con el formato oficial de SUNAT.";
            downloadBtn.className = "ribot-btn-download";
            downloadBtn.innerHTML = '<i class="fa-solid fa-download"></i> Descargar Word';
        } else {
            resultTitleText.textContent = "¡Procesamiento Masivo Listo!";
            downloadIcon.className = "ribot-dl-icon zip";
            downloadIcon.innerHTML = '<i class="fa-solid fa-file-zipper"></i>';
            downloadName.textContent = res.filename;
            downloadDesc.textContent = "Paquete comprimido ZIP con todas las Resoluciones en Word redactadas.";
            downloadBtn.className = "ribot-btn-download zip";
            downloadBtn.innerHTML = '<i class="fa-solid fa-file-arrow-down"></i> Descargar ZIP Completo';
        }

        resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    window.ribotResetResults = function () {
        if (emptyRight) emptyRight.style.display = 'flex';
        if (resultCard) resultCard.style.display = 'none';
    };

    // LOADER
    function startLoading() {
        if (!overlay) return;
        overlay.classList.add('active');
        let idx = 0;
        loaderStatus.textContent = loaderMessages[0];

        statusInterval = setInterval(() => {
            idx = (idx + 1) % loaderMessages.length;
            loaderStatus.style.opacity = '0';
            setTimeout(() => {
                loaderStatus.textContent = loaderMessages[idx];
                loaderStatus.style.opacity = '1';
            }, 180);
        }, 2200);
    }

    function stopLoading() {
        if (!overlay) return;
        overlay.classList.remove('active');
        if (statusInterval) {
            clearInterval(statusInterval);
            statusInterval = null;
        }
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }
})();
