document.addEventListener("DOMContentLoaded", () => {
    // Leer el total de registros inyectado por el servidor (sesión persistente)
    const docPage = document.querySelector(".doc-page");
    const serverRecordCount = parseInt(docPage?.dataset.recordCount || "0");
    let totalRecords = serverRecordCount;

    // Restaurar último índice visitado desde localStorage
    const STORAGE_KEY = "doccheck_last_idx";
    let currentIdx = totalRecords > 0
        ? Math.min(parseInt(localStorage.getItem(STORAGE_KEY) || "0"), totalRecords - 1)
        : 0;

    let originalValues = {}; // Para detectar cambios sin guardar
    let hasUnsavedChanges = false;

    // Referencias a elementos
    const btnUpload = document.getElementById("btn-upload");
    const fileInput = document.getElementById("file-upload");
    const btnSave = document.getElementById("btn-save");
    const btnPrev = document.getElementById("btn-prev");
    const btnNext = document.getElementById("btn-next");
    const btnExpExcel = document.getElementById("btn-export-excel");
    const btnExpTxt = document.getElementById("btn-export-txt");
    const btnConsultar = document.querySelector(".btn-consultar");

    // Referencias para indicador de progreso
    const currentPositionEl = document.getElementById("current-position");
    const totalRecordsEl = document.getElementById("total-records");

    // ===== CONVERTIR A NÚMEROS ROMANOS =====
    function convertirARomano(valor) {
        if (!valor) return "";

        // Si ya es romano (solo letras I, V, X, L, C, D, M), devolverlo
        const valorStr = String(valor).trim().toUpperCase();
        if (/^[IVXLCDM]+$/.test(valorStr)) {
            return valorStr;
        }

        // Si es número, convertir a romano
        const num = parseInt(valor);
        if (isNaN(num) || num <= 0) return valorStr;

        const romanos = [
            { valor: 1000, simbolo: 'M' },
            { valor: 900, simbolo: 'CM' },
            { valor: 500, simbolo: 'D' },
            { valor: 400, simbolo: 'CD' },
            { valor: 100, simbolo: 'C' },
            { valor: 90, simbolo: 'XC' },
            { valor: 50, simbolo: 'L' },
            { valor: 40, simbolo: 'XL' },
            { valor: 10, simbolo: 'X' },
            { valor: 9, simbolo: 'IX' },
            { valor: 5, simbolo: 'V' },
            { valor: 4, simbolo: 'IV' },
            { valor: 1, simbolo: 'I' }
        ];

        let resultado = '';
        let n = num;
        for (const { valor: v, simbolo } of romanos) {
            while (n >= v) {
                resultado += simbolo;
                n -= v;
            }
        }
        return resultado;
    }

    // ===== TOAST NOTIFICATIONS =====
    function showToast(type, message) {
        if (window.showToast) {
            window.showToast(message, type);
        }
    }


    // ===== DETECTAR CAMBIOS SIN GUARDAR =====
    const editableFields = [
        "field-folios", "field-tipo-doc", "field-num-doc", "field-razon",
        "field-ruc", "field-fecha", "field-obs", "field-x1", "field-x2", "field-x3"
    ];

    function storeOriginalValues() {
        originalValues = {};
        editableFields.forEach(id => {
            const el = document.getElementById(id);
            if (el) originalValues[id] = el.value;
        });
        hasUnsavedChanges = false;
    }

    function checkForChanges() {
        for (const id of editableFields) {
            const el = document.getElementById(id);
            if (el && originalValues[id] !== el.value) {
                return true;
            }
        }
        return false;
    }

    // Escuchar cambios en los campos
    editableFields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener("input", () => {
                hasUnsavedChanges = checkForChanges();
            });
        }
    });

    // ===== CONFIRMACIÓN ANTES DE NAVEGAR =====
    function confirmNavigation(callback) {
        if (!hasUnsavedChanges) {
            callback();
            return;
        }

        showConfirmModal(
            "Cambios sin guardar",
            "Tienes cambios sin guardar. ¿Qué deseas hacer?",
            [
                {
                    text: "Guardar y continuar", type: "primary", action: async () => {
                        await saveCurrentRecord();
                        callback();
                    }
                },
                {
                    text: "Descartar cambios", type: "secondary", action: () => {
                        hasUnsavedChanges = false;
                        callback();
                    }
                },
                { text: "Cancelar", type: "cancel", action: () => { } }
            ]
        );
    }

    function showConfirmModal(title, message, buttons) {
        const overlay = document.createElement("div");
        overlay.className = "confirm-overlay";
        overlay.innerHTML = `
            <div class="confirm-modal">
                <div class="confirm-icon">
                    <i class="fa-solid fa-exclamation-triangle"></i>
                </div>
                <h3>${title}</h3>
                <p>${message}</p>
                <div class="confirm-actions"></div>
            </div>
        `;

        const actionsContainer = overlay.querySelector(".confirm-actions");
        buttons.forEach(btn => {
            const button = document.createElement("button");
            button.className = `confirm-btn confirm-btn-${btn.type}`;
            button.textContent = btn.text;
            button.addEventListener("click", () => {
                overlay.remove();
                btn.action();
            });
            actionsContainer.appendChild(button);
        });

        document.body.appendChild(overlay);
    }

    // ===== CARGAR ARCHIVO EXCEL =====
    if (btnUpload && fileInput) {
        btnUpload.addEventListener("click", () => fileInput.click());
        fileInput.addEventListener("change", async () => {
            if (!fileInput.files.length) return;
            const formData = new FormData();
            formData.append("file", fileInput.files[0]);
            try {
                const resp = await fetch("/doccheck/upload", { method: "POST", body: formData });
                const data = await resp.json();
                if (data.status === "success") {
                    totalRecords = data.count;
                    showToast("success", `¡Archivo cargado! Total: ${data.count} registros`);
                    currentIdx = 0;
                    localStorage.setItem(STORAGE_KEY, 0); // Resetear al subir nuevo archivo
                    loadRecord(currentIdx);
                    updateNavButtons();
                    updateProgressIndicator();
                } else {
                    showToast("error", "Error: " + data.error);
                }
            } catch (e) {
                console.error("Error de red al cargar archivo:", e);
                showToast("error", "Error de red al cargar archivo");
            }
        });
    }

    // ===== ACTUALIZAR INDICADOR DE PROGRESO =====
    function updateProgressIndicator() {
        if (currentPositionEl) currentPositionEl.textContent = totalRecords > 0 ? currentIdx + 1 : 0;
        if (totalRecordsEl) totalRecordsEl.textContent = totalRecords;
    }

    // ===== NAVEGACIÓN =====
    if (btnPrev) {
        btnPrev.addEventListener("click", () => {
            if (currentIdx > 0) {
                confirmNavigation(() => {
                    currentIdx--;
                    localStorage.setItem(STORAGE_KEY, currentIdx);
                    loadRecord(currentIdx);
                    updateNavButtons();
                    updateProgressIndicator();
                });
            }
        });
    }

    if (btnNext) {
        btnNext.addEventListener("click", () => {
            if (currentIdx < totalRecords - 1) {
                confirmNavigation(() => {
                    currentIdx++;
                    localStorage.setItem(STORAGE_KEY, currentIdx);
                    loadRecord(currentIdx);
                    updateNavButtons();
                    updateProgressIndicator();
                });
            }
        });
    }

    function updateNavButtons() {
        if (btnPrev) {
            btnPrev.disabled = currentIdx <= 0;
            btnPrev.style.opacity = currentIdx <= 0 ? "0.5" : "1";
        }
        if (btnNext) {
            btnNext.disabled = currentIdx >= totalRecords - 1;
            btnNext.style.opacity = currentIdx >= totalRecords - 1 ? "0.5" : "1";
        }
    }

    // ===== IR A REGISTRO ESPECÍFICO =====
    const recordIdxInput = document.getElementById("record-idx");
    if (recordIdxInput) {
        recordIdxInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                e.preventDefault();
                const searchValue = recordIdxInput.value.trim();

                if (!searchValue) {
                    showToast("info", "Ingresa un número de registro");
                    return;
                }

                if (totalRecords === 0) {
                    showToast("info", "Primero carga un archivo Excel");
                    return;
                }

                const idx = parseInt(searchValue) - 1;

                if (!isNaN(idx) && idx >= 0 && idx < totalRecords) {
                    confirmNavigation(() => {
                        currentIdx = idx;
                        localStorage.setItem(STORAGE_KEY, currentIdx);
                        loadRecord(currentIdx);
                        updateNavButtons();
                        updateProgressIndicator();
                    });
                } else {
                    showToast("error", `Registro ${searchValue} no existe. Rango: 1 - ${totalRecords}`);
                }
            }
        });
    }

    // ===== GUARDAR REGISTRO =====
    async function saveCurrentRecord() {
        const formData = new FormData();
        const fieldMap = {
            "field-folios": "Folios",
            "field-tipo-doc": "Tipo Doc",
            "field-num-doc": "N° Documento",
            "field-razon": "Razón Social",
            "field-ruc": "RUC",
            "field-fecha": "Fecha Extrema",
            "field-obs": "Observaciones",
            "field-x1": "X1",
            "field-x2": "X2",
            "field-x3": "X3",
        };
        for (const [id, key] of Object.entries(fieldMap)) {
            const el = document.getElementById(id);
            if (el) formData.append(key, el.value);
        }
        try {
            const resp = await fetch(`/doccheck/save/${currentIdx}`, { method: "POST", body: formData });
            const res = await resp.json();
            if (res.status === "success") {
                showToast("success", "Guardado correctamente");
                storeOriginalValues();
                loadLastChange();  // Refrescar el badge
                return true;
            } else {
                showToast("error", "Error: " + res.error);
                return false;
            }
        } catch (e) {
            console.error("Error de red al guardar:", e);
            showToast("error", "Error de red al guardar");
            return false;
        }
    }

    if (btnSave) {
        btnSave.addEventListener("click", saveCurrentRecord);
    }

    // ===== ÚLTIMO CAMBIO =====
    async function loadLastChange() {
        const badge = document.getElementById("last-change-badge");
        if (!badge) return;

        try {
            const resp = await fetch("/doccheck/last-change");
            const res  = await resp.json();

            if (res.status === "success") {
                const d = res.data;
                badge.style.display = "flex";
                document.getElementById("last-change-summary").textContent =
                    `${d.campo}: ${(d.nuevo || "(vacío)").substring(0, 22)}`;
                document.getElementById("lcp-registro").textContent = d.registro || "—";
                document.getElementById("lcp-campo").textContent    = d.campo    || "—";
                document.getElementById("lcp-antiguo").textContent  = d.antiguo  || "(vacío)";
                document.getElementById("lcp-nuevo").textContent    = d.nuevo    || "(vacío)";
                document.getElementById("lcp-fecha").textContent    = d.fecha    || "—";
            } else if (res.status === "empty") {
                // Archivo cargado pero sin cambios aún → mostrar badge vacío
                badge.style.display = "flex";
                document.getElementById("last-change-summary").textContent = "Sin cambios aún";
                document.getElementById("lcp-registro").textContent = "—";
                document.getElementById("lcp-campo").textContent    = "—";
                document.getElementById("lcp-antiguo").textContent  = "—";
                document.getElementById("lcp-nuevo").textContent    = "—";
                document.getElementById("lcp-fecha").textContent    = "Aún no hay cambios registrados";
            } else {
                badge.style.display = "none";
            }
        } catch (e) {
            console.warn("No se pudo consultar el estado:", e);
            badge.style.display = "none";
        }
    }

    // ===== CONSULTAR RUC =====
    if (btnConsultar) {
        btnConsultar.addEventListener("click", () => {
            const rucField = document.getElementById("field-ruc");
            const ruc = rucField ? rucField.value.trim() : "";

            const URL_CONSULTA = "http://intranet/cl-ti-iaconsruc/jcrS01Alias"; // NOSONAR: intranet interna, sin HTTPS
            const abrirConsulta = () => window.open(URL_CONSULTA, "_blank", "noopener");

            if (!ruc) {
                showToast("info", "No hay RUC ingresado");
                abrirConsulta();
            } else {
                navigator.clipboard.writeText(ruc).then(() => {
                    showToast("success", `RUC copiado: ${ruc}`);
                    abrirConsulta();
                }).catch(() => {
                    showToast("info", `RUC: ${ruc} - Copia manualmente`);
                    abrirConsulta();
                });
            }
        });
    }

    // ===== EXPORTAR =====
    if (btnExpExcel) {
        btnExpExcel.addEventListener("click", () => {
            if (totalRecords === 0) {
                showToast("info", "Primero carga un archivo Excel");
                return;
            }
            showExportModal("excel");
        });
    }

    if (btnExpTxt) {
        btnExpTxt.addEventListener("click", () => {
            if (totalRecords === 0) {
                showToast("info", "Primero carga un archivo Excel");
                return;
            }
            showExportModal("txt");
        });
    }

    // ===== MODAL DE EXPORTACIÓN MEJORADO =====
    function showExportModal(formato) {
        const overlay = document.createElement("div");
        overlay.className = "export-overlay";
        overlay.innerHTML = `
            <div class="export-modal-v2">
                <h2 class="export-title">SELECCIONE QUE EXPORTAR</h2>
                
                <div class="export-cards">
                    <div class="export-card" data-type="todo">
                        <div class="export-card-icon">
                            <i class="fa-solid fa-boxes-stacked"></i>
                        </div>
                        <span class="export-card-label">TODO</span>
                    </div>
                    
                    <div class="export-card" data-type="caja">
                        <div class="export-card-icon">
                            <i class="fa-solid fa-box"></i>
                        </div>
                        <span class="export-card-label">CAJA</span>
                    </div>
                </div>

                <div class="export-caja-section" id="export-caja-section" style="display: none;">
                    <div class="export-caja-field">
                        <span class="export-caja-label">Nº DE CAJA</span>
                        <input type="text" id="export-caja-input" class="export-caja-input" placeholder="">
                    </div>
                </div>

                <button class="export-btn-submit" id="export-submit">
                    EXPORTAR
                </button>
            </div>
        `;

        document.body.appendChild(overlay);

        let selectedType = "todo";
        const cards = overlay.querySelectorAll(".export-card");
        const cajaSection = overlay.querySelector("#export-caja-section");
        const cajaInput = overlay.querySelector("#export-caja-input");

        cards.forEach(card => {
            card.addEventListener("click", () => {
                cards.forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");
                selectedType = card.dataset.type;

                if (selectedType === "caja") {
                    cajaSection.style.display = "flex";
                    cajaInput.focus();
                } else {
                    cajaSection.style.display = "none";
                }
            });
        });

        // Seleccionar "Todos" por defecto
        cards[0].classList.add("selected");

        overlay.querySelector("#export-submit").addEventListener("click", () => {
            let url = `/doccheck/export/${formato}`;

            if (selectedType === "caja") {
                const cajaVal = cajaInput.value.trim();
                if (!cajaVal) {
                    showToast("warning", "Ingresa el N° de Caja");
                    cajaInput.focus();
                    return;
                }
                url += `?caja=${encodeURIComponent(cajaVal)}`;
            }

            window.location.href = url;
            overlay.remove();
            showToast("success", `Exportando a ${formato.toUpperCase()}...`);
        });

        overlay.addEventListener("click", (e) => {
            if (e.target === overlay) overlay.remove();
        });

        // Cerrar con ESC
        const handleEsc = (e) => {
            if (e.key === "Escape") {
                overlay.remove();
                document.removeEventListener("keydown", handleEsc);
            }
        };
        document.addEventListener("keydown", handleEsc);
    }

    // ===== CARGAR REGISTRO =====
    async function loadRecord(idx) {
        try {
            const resp = await fetch(`/doccheck/record/${idx}`);
            const res = await resp.json();
            if (res.status === "success") {
                const d = res.data;
                const setVal = (id, val) => {
                    const el = document.getElementById(id);
                    if (el) el.value = val || "";
                };

                // Formatear N° Registro con Tomo si existe
                let registroDisplay = d["N° Registro"] || "";
                const tomo = d["Tomo"];
                if (tomo && tomo !== "" && tomo !== null) {
                    // Si el tomo ya es romano (I, II, III, etc) usarlo directo
                    // Si es número, convertirlo a romano
                    const tomoRomano = convertirARomano(tomo);
                    registroDisplay = `${d["N° Registro"]}-${tomoRomano}`;
                }
                setVal("record-idx", registroDisplay);
                setVal("field-paquete", d["Paquete"]);
                setVal("field-caja", d["Caja"]);
                setVal("field-folios", d["Folios"]);
                setVal("field-tipo-doc", d["Tipo Doc"]);
                setVal("field-num-doc", d["N° Documento"]);
                setVal("field-razon", d["Razón Social"]);
                setVal("field-ruc", d["RUC"]);
                setVal("field-fecha", d["Fecha Extrema"]);
                setVal("field-obs", d["Observaciones"]);
                setVal("field-x1", d["X1"]);
                setVal("field-x2", d["X2"]);
                setVal("field-x3", d["X3"]);

                // Guardar valores originales para detectar cambios
                storeOriginalValues();
                updateProgressIndicator();
            } else if (res.error === "Index out of range") {
                // No hacer nada
            }
        } catch (e) {
            console.error("Error cargando registro:", e);
        }
    }

    // ===== ATAJOS DE TECLADO =====
    document.addEventListener("keydown", (e) => {
        // Ctrl + S = Guardar
        if (e.ctrlKey && e.key === "s") {
            e.preventDefault();
            if (btnSave) btnSave.click();
        }
        // Flechas izq/der = Navegación (solo si no está en input)
        if (e.key === "ArrowLeft" && !e.target.matches("input")) {
            if (btnPrev && !btnPrev.disabled) btnPrev.click();
        }
        if (e.key === "ArrowRight" && !e.target.matches("input")) {
            if (btnNext && !btnNext.disabled) btnNext.click();
        }
    });

    // Inicializar — si hay archivo, cargar el último registro visitado
    if (totalRecords > 0) {
        const progressEl = document.getElementById("progress-indicator");
        if (progressEl) progressEl.style.display = "flex";
        loadRecord(currentIdx);  // ← currentIdx ya viene de localStorage
        loadLastChange();        // ← Mostrar último cambio si existe
    }
    updateNavButtons();
    updateProgressIndicator();
});
