/**
 * app.js
 * Lógica principal para el Wizard de Generación de Documentos
 */

const app = {
    state: {
        step: 1,
        zoom: 1.0,
        data: {
            documentType: null,
            caseId: null,
            formData: {}
        }
    },

    init: function () {
        console.log("App iniciada");
        // Establecer fecha actual por defecto
        const today = new Date().toISOString().split('T')[0];
        const dateInput = document.getElementById('inp-date');
        if (dateInput) {
            dateInput.value = today;
            this.updatePreview();
        }
    },

    goToStep: function (stepNumber) {
        // Ocultar todos los pasos
        document.querySelectorAll('.wizard-step').forEach(el => el.classList.remove('active'));

        // Mostrar el paso actual
        const stepEl = document.getElementById(`step-${stepNumber}`);
        if (stepEl) {
            stepEl.classList.add('active');
            this.state.step = stepNumber;
        }
    },

    selectDocument: function (docType) {
        this.state.data.documentType = docType;
        console.log("Documento seleccionado:", docType);

        // Actualizar etiqueta del paso 3
        const label = document.getElementById('current-doc-label');
        if (label) {
            // Formatear nombre para mostrar
            const names = {
                'carta_inductiva': 'Carta Inductiva',
                'esquela_citacion': 'Esquela de Citación',
                'resolucion_multa': 'Resolución de Multa'
            };
            label.innerText = `Documento: ${names[docType] || docType}`;

            // Actualizar vista previa también
            const prevType = document.getElementById('prev-type');
            if (prevType) prevType.innerText = (names[docType] || 'DOCUMENTO').toUpperCase();
        }

        this.goToStep(2);
    },

    selectCase: function (caseName) {
        this.state.data.caseId = caseName;
        console.log("Caso seleccionado:", caseName);

        // Prellenar datos simulados según el caso
        if (caseName.includes("5214")) {
            document.getElementById('inp-name').value = "EMPRESA DE TRANSPORTES S.A.C.";
            document.getElementById('inp-ruc').value = "20601234567";
            document.getElementById('inp-motive').value = "Se ha detectado que no ha presentado la declaración jurada mensual del IGV correspondiente al periodo Enero 2024, a pesar de haber emitido comprobantes de pago electrónicos.";
        } else {
            document.getElementById('inp-name').value = "JUAN PEREZ CONSULTORES E.I.R.L.";
            document.getElementById('inp-ruc').value = "20559876543";
            document.getElementById('inp-motive').value = "Inconsistencia en el cálculo de Renta de Tercera Categoría respecto a los ingresos netos declarados.";
        }

        this.updatePreview();
        this.goToStep(3);
    },

    filterCases: function (query) {
        const lowerQuery = query.toLowerCase();
        const items = document.querySelectorAll('.case-item');
        items.forEach(item => {
            const text = item.innerText.toLowerCase();
            if (text.includes(lowerQuery)) {
                item.style.display = 'flex';
            } else {
                item.style.display = 'none';
            }
        });
    },

    updatePreview: function () {
        // Mapeo de IDs de inputs a IDs de vista previa
        const map = {
            'inp-city': 'prev-city',
            'inp-date': 'prev-date',
            'inp-number': 'prev-number',
            'inp-name': 'prev-name',
            'inp-ruc': 'prev-ruc',
            'inp-address': 'prev-address',
            'inp-motive': 'prev-motive',
            'inp-result': 'prev-result'
        };

        for (const [inputId, prevId] of Object.entries(map)) {
            const input = document.getElementById(inputId);
            const prev = document.getElementById(prevId);
            if (input && prev) {
                // Formato especial para fecha
                if (inputId === 'inp-date' && input.value) {
                    const d = new Date(input.value);
                    // Ajuste de zona horaria simple
                    const userTimezoneOffset = d.getTimezoneOffset() * 60000;
                    const dateObj = new Date(d.getTime() + userTimezoneOffset);

                    const options = { year: 'numeric', month: 'long', day: 'numeric' };
                    prev.innerText = dateObj.toLocaleDateString('es-ES', options);
                } else if (input.value) {
                    prev.innerText = input.value;
                } else {
                    // Mantener placeholder si está vacío (opcional, aquí borramos o ponemos ...)
                    // prev.innerText = '...'; 
                }
            }
        }

        // Caso especial: Firma
        const signerSelect = document.getElementById('inp-signer');
        if (signerSelect) {
            const val = signerSelect.value; // "Nombre|Rol"
            if (val) {
                const parts = val.split('|');
                document.getElementById('prev-sign-name').innerText = parts[0];
                document.getElementById('prev-sign-role').innerText = parts[1];
            } else {
                document.getElementById('prev-sign-name').innerText = '';
                document.getElementById('prev-sign-role').innerText = '';
            }
        }

        // Caso especial: Asunto (toma del caso seleccionado o genera uno)
        const caseName = this.state.data.caseId || 'Sin Asignar';
        const modalPrev = document.getElementById('prev-case-name');
        if (modalPrev) modalPrev.innerText = caseName;
    },

    zoom: function (delta) {
        this.state.zoom += delta;
        // Limitar zoom
        if (this.state.zoom < 0.5) this.state.zoom = 0.5;
        if (this.state.zoom > 2.0) this.state.zoom = 2.0;

        const page = document.getElementById('a4-page');
        const level = document.getElementById('zoom-level');

        if (page) {
            page.style.transform = `scale(${this.state.zoom})`;
            // Ajustar margen para que no se corte al hacer zoom in
            if (this.state.zoom > 1) {
                page.style.marginTop = `${(this.state.zoom - 1) * 150}px`;
                page.style.marginBottom = `${(this.state.zoom - 1) * 150}px`;
            } else {
                page.style.marginTop = '0';
                page.style.marginBottom = '0';
            }
        }
        if (level) {
            level.innerText = `${Math.round(this.state.zoom * 100)}%`;
        }
    },

    downloadPDF: function () {
        showToast("Funcionalidad de descarga PDF se implementará en el backend", "info");
    }
};

// Inicializar al cargar
document.addEventListener('DOMContentLoaded', () => {
    app.init();

    // Modal global - cerrar al hacer clic en Aceptar
    const uiModalBtn = document.getElementById('uiModalBtn');
    const uiModal = document.getElementById('uiModal');

    if (uiModalBtn && uiModal) {
        uiModalBtn.addEventListener('click', () => {
            uiModal.classList.add('uimodal-hidden');
        });

        // También cerrar al hacer clic fuera del modal
        uiModal.addEventListener('click', (e) => {
            if (e.target === uiModal) {
                uiModal.classList.add('uimodal-hidden');
            }
        });
    }
});

/**
 * Global Toast Notification (DataMask Style)
 * @param {string} message - Message to show
 * @param {string} type - 'success', 'error', 'warning', or 'info'
 */
window.showToast = function(message, type) {
    type = type || 'info';

    // Get or create the toast element dynamically
    let toast = document.getElementById('sicodeToast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'sicodeToast';
        document.body.appendChild(toast);
    }

    // Reset and apply type class
    toast.className = 'sicode-toast';
    if (type === 'success') toast.classList.add('toast-success');
    else if (type === 'error') toast.classList.add('toast-error');
    else if (type === 'warning') toast.classList.add('toast-warning');

    // Icon map
    var icons = {
        success: '<i class="fa-solid fa-circle-check"></i>',
        error:   '<i class="fa-solid fa-circle-exclamation"></i>',
        warning: '<i class="fa-solid fa-triangle-exclamation"></i>',
        info:    '<i class="fa-solid fa-circle-info"></i>'
    };

    toast.innerHTML = (icons[type] || icons.info) + ' <span>' + message + '</span>';
    toast.classList.add('visible');

    // Auto-hide after 2.5s
    if (window._toastTimeout) clearTimeout(window._toastTimeout);
    window._toastTimeout = setTimeout(function() {
        toast.classList.remove('visible');
    }, 2500);
};

