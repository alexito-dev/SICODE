/**
 * AUTODOC MÓDULO - LÓGICA FRONTEND
 * --------------------------------
 * Este archivo maneja toda la interactividad de la página sin recargarla (SPA behaviour).
 * Usamos Vanilla JS (Javascript puro) para máximo rendimiento y control.
 * 
 * Estructura:
 * - Constantes de configuración
 * - Objeto principal 'app' que encapsula todo el estado y funciones
 * - Inicialización al cargar el DOM
 */

// URL Base relativa para las peticiones al Backend (FastAPI)
const API_URL = "/autodoc/api";

// Campos que serán obligatorios para permitir la descarga
const REQUIRED_FIELDS = [
    'numero_esquela', 'ciudad', 'fecha', 'nombre',
    'ruc', 'domicilio', 'expediente', 'fecha_ref', 'firmante_nombre'
];

// Mapeo de nombres técnicos (keys) a nombres legibles para el usuario
const FIELD_LABELS = {
    'numero_esquela': 'Número de Esquela',
    'ciudad': 'Ciudad',
    'fecha': 'Fecha',
    'nombre': 'Nombre',
    'ruc': 'RUC',
    'domicilio': 'Domicilio Fiscal',
    'expediente': 'Expediente',
    'fecha_ref': 'Fecha de Referencia',
    'firmante_nombre': 'Firmante'
};

/**
 * Objeto Principal 'app'
 * Patrón 'Singleton' simple para organizar el código.
 */
const app = {
    // Estado global de la aplicación
    state: {
        step: 1,                // Paso actual del Wizard (1, 2, 3)
        docType: '',            // Tipo de documento seleccionado
        selectedTemplate: null, // Objeto plantilla actual
        formData: {},           // Datos que el usuario va escribiendo
        debounceTimer: null     // Timer para no actualizar la vista previa en cada tecla
    },

    // Base de datos Local Falsa (Fallback) por si falla la API
    mockRucDB: {
        '20600000001': { nombre: 'MINERA YANACOCHA S.R.L.', domicilio: 'AV. VICTOR ANDRES BELAUNDE 147' },
        '20100000001': { nombre: 'ALICORP S.A.A.', domicilio: 'AV. ARGENTINA 4793' },
        '20555555555': { nombre: 'EMPRESA DE TRANSPORTES S.A.C.', domicilio: 'JR. LOS ALAMOS 123' }
    },

    // Tipos de documentos visuales (Cards)
    documentTypes: [
        { name: 'Pago con error', icon: 'fa-triangle-exclamation' },
        { name: 'Código no tributario', icon: 'fa-file-invoice-dollar' },
        { name: 'Esquela', icon: 'fa-envelope-open-text' },
        { name: 'Carta informativa', icon: 'fa-circle-info' },
        { name: 'Resolución de Multa', icon: 'fa-gavel' },
        { name: 'Orden de Pago', icon: 'fa-money-bill-wave' },
        { name: 'Carta Inductiva', icon: 'fa-magnifying-glass-chart' },
        { name: 'Solicitud de Devolución', icon: 'fa-file-invoice-dollar' }
    ],

    templates: [], // Se llenará con datos del servidor

    /**
     * Inicializador: Se ejecuta cuando la página carga.
     * Carga la lista inicial de tarjetas y llama a la API.
     */
    init: async function () {
        this.renderDocuments();

        // Intentamos cargar las plantillas del backend
        try {
            await this.fetchTemplates();
        } catch (e) {
            console.warn("Backend no listo, error:", e);
        }
    },

    /**
     * Llama al endpoint /api/templates para obtener los documentos disponibles
     */
    fetchTemplates: async function () {
        const container = document.getElementById('case-list-container');
        container.innerHTML = '<div style="padding:20px;text-align:center;">Cargando plantillas...</div>';

        try {
            const res = await fetch(`${API_URL}/templates`);
            if (!res.ok) throw new Error("API Response not OK");

            this.templates = await res.json();
            this.renderCases(this.templates);
        } catch (err) {
            console.error(err);
            this.templates = [
                { name: 'Esquela_Generica.docx' },
                { name: 'Carta_Inductiva_2024.docx' }
            ];
            this.renderCases(this.templates);
            if (window.showToast) window.showToast('Servidor offline — usando plantillas de muestra', 'warning');
        }
    },

    /**
     * Obtiene la estructura del formulario (Schema) para una plantilla específica.
     * Esto permite que el formulario sea dinámico.
     */
    fetchSchema: async function (filename) {
        try {
            const res = await fetch(`${API_URL}/templates/${filename}/schema`);
            if (!res.ok) throw new Error("API Schema Error");
            const data = await res.json();
            return data.fields; // Retorna array de campos
        } catch (err) {
            console.warn("Usando esquema por defecto: " + err);
            // Schema por defecto de emergencia
            return [
                { key: 'numero_esquela', label: 'Número Esquela', type: 'text' },
                { key: 'fecha', label: 'Fecha', type: 'date' },
                { key: 'nombre', label: 'Nombre Contribuyente', type: 'text' },
                { key: 'ruc', label: 'RUC', type: 'text' },
                { key: 'firmante_nombre', label: 'Firmante', type: 'select' }
            ];
        }
    },

    /**
     * Genera la vista previa del PDF llamando al backend.
     * Se usa un iframe para mostrar el blob resultante.
     */
    generatePreview: async function () {
        const payload = {
            filename: this.state.selectedTemplate ? this.state.selectedTemplate.name : 'unknown',
            data: this.state.formData
        };

        const previewContainer = document.querySelector('.preview-canvas');
        const statusLabel = document.querySelector('.preview-status'); // Indicador visual de estado

        if (statusLabel) statusLabel.innerHTML = '<span class="status-dot" style="background: #f59e0b;"></span> Generando PDF...';

        try {
            const res = await fetch(`${API_URL}/generate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }, // Importante: JSON header
                body: JSON.stringify(payload)
            });

            if (!res.ok) throw new Error("Error generando PDF");

            // Convertimos la respuesta binaria a una URL visualizable
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);

            // Inyectamos el PDF en un iframe sin bordes
            previewContainer.innerHTML = `
                <iframe 
                    src="${url}#toolbar=0&view=FitH" 
                    width="100%" 
                    height="100%" 
                    style="border:none;">
                </iframe>
            `;

            if (statusLabel) statusLabel.innerHTML = '<span class="status-dot" style="background: #10b981;"></span> Documento generado';
            if (window.showToast) window.showToast('Vista previa generada', 'success');

        } catch (err) {
            console.error(err);
            if (statusLabel) statusLabel.innerHTML = '<span class="status-dot" style="background: #ef4444;"></span> Error vista previa (Mock Backend)';
            if (window.showToast) window.showToast('No se pudo generar la vista previa', 'error');
        }
    },

    /**
     * Maneja la descarga final del archivo.
     * Incluye validación de campos obligatorios.
     */
    downloadPDF: async function () {
        // Validación omitida para demo, descomentar para producción
        // const missingFields = requiredFields.filter(f => !this.state.formData[f] || this.state.formData[f].trim() === '');
        const missingFields = [];

        if (missingFields.length > 0) {
            // Convertimos keys a nombres bonitos
            const fieldNames = missingFields.map(f => FIELD_LABELS[f] || f).join(', ');
            if (window.showToast) {
                window.showToast(`Campos incompletos: ${fieldNames}`, 'error');
            } else {
                alert(`⚠️ Campos incompletos:\n\n${fieldNames}`);
            }
            return;
        }

        const payload = {
            filename: this.state.selectedTemplate ? this.state.selectedTemplate.name : 'doc',
            data: this.state.formData
        };

        try {
            // Reutilizamos el endpoint generate
            const res = await fetch(`${API_URL}/generate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!res.ok) throw new Error("Error generando PDF");

            // Truco clásico para forzar descarga: Crear link <a> oculto y hacerle click
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Esquela_${this.state.formData.numero_esquela || 'doc'}.pdf`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            if (window.showToast) window.showToast('PDF descargado correctamente', 'success');
        } catch (err) {
            console.error('Error descargando PDF:', err);
            if (window.showToast) {
                window.showToast('Error descargando PDF (Backend Mock activo)', 'error');
            } else {
                alert('Error descargando PDF (Backend Mock activo)');
            }
        }
    },

    // --- FUNCIONES DE RENDERIZADO (DOM Manipulation) ---

    // Dibuja las tarjetas del paso 1
    renderDocuments: function () {
        const container = document.getElementById('doc-grid-container');
        if (!container) return;
        container.innerHTML = '';
        this.documentTypes.forEach(doc => {
            const card = document.createElement('div');
            card.className = 'doc-card';
            card.onclick = () => this.selectDocument(doc.name); // Closure
            card.innerHTML = `<div class="doc-icon"><i class="fa-solid ${doc.icon}"></i></div><div class="doc-name">${doc.name}</div>`;
            container.appendChild(card);
        });
    },

    // Dibuja la lista del paso 2
    renderCases: function (list) {
        const container = document.getElementById('case-list-container');
        container.innerHTML = '';
        if (list.length === 0) {
            container.innerHTML = '<div style="padding:20px;text-align:center;">No hay plantillas</div>';
            return;
        }
        list.forEach(tpl => {
            const cleanName = tpl.name.replace('.docx', '').replace(/_/g, ' ');
            const div = document.createElement('div');
            div.className = 'case-item';
            div.innerHTML = `<div class="case-info"><span class="case-desc">${cleanName}</span></div><i class="fa-solid fa-chevron-right" style="color:#ccc;"></i>`;
            div.onclick = () => this.selectTemplate(tpl);
            container.appendChild(div);
        });
    },

    /**
     * Construye el formulario dinámicamente basado en el JSON Schema.
     * Agrupa campos por secciones visuales.
     */
    renderForm: function (schema) {
        const container = document.querySelector('.form-content');
        container.innerHTML = '';

        // Caja de Alerta (inicialmente oculta)
        const alertBox = document.createElement('div');
        alertBox.className = 'form-alert';
        alertBox.style.display = 'none';
        alertBox.innerHTML = `<i class="fa-solid fa-circle-exclamation form-alert-icon"></i><div><span class="form-alert-title">Campos obligatorios incompletos</span>Complete todos los campos.</div>`;
        container.appendChild(alertBox);

        // Definición de secciones y orden deseado
        const sections = {
            'Datos Generales': { fields: ['numero_esquela', 'ciudad', 'fecha'], required: true },
            'Datos del Contribuyente': { fields: ['nombre', 'ruc', 'domicilio'], required: true },
            'Referencia': { fields: ['expediente', 'fecha_ref'], required: false },
            'Firma': { fields: ['firmante_nombre'], required: true }
        };

        // Ordenar campos según nuestra configuración de secciones
        const fieldOrder = Object.values(sections).flatMap(s => s.fields);
        const sortedSchema = schema.sort((a, b) => {
            const ia = fieldOrder.indexOf(a.key), ib = fieldOrder.indexOf(b.key);
            return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
        });

        // Agrupar campos en un diccionario temporal
        const fieldsBySection = {};
        sortedSchema.forEach(field => {
            if (field.type === 'readonly') return;
            let sectionName = 'Otros'; // Fallback category
            for (const [name, config] of Object.entries(sections)) {
                if (config.fields.includes(field.key)) sectionName = name;
            }
            if (!fieldsBySection[sectionName]) fieldsBySection[sectionName] = [];
            fieldsBySection[sectionName].push(field);
        });

        // Generar HTML para cada sección
        for (const [sectionName, config] of Object.entries(sections)) {
            const sectionFields = fieldsBySection[sectionName];
            if (!sectionFields || sectionFields.length === 0) continue;

            const sectionCard = document.createElement('div');
            sectionCard.className = 'form-section';
            const reqBadge = config.required ? '<span class="badge-required">Obligatorio</span>' : '';
            sectionCard.innerHTML = `<div class="form-section-title"><div class="section-label">${sectionName} ${reqBadge}</div><i class="fa-solid fa-chevron-down section-toggle-icon"></i></div>`;

            const body = document.createElement('div');
            body.className = 'form-section-body';

            sectionFields.forEach(field => {
                const div = document.createElement('div');
                div.className = 'form-group';
                const isRequired = REQUIRED_FIELDS.includes(field.key);
                div.innerHTML = `<label>${field.label || field.key} ${isRequired ? '<span class="required-star">*</span>' : ''}</label>`;

                let input;
                // Lógica para tipos de input
                if (field.type === 'select') {
                    input = document.createElement('select');
                    input.innerHTML = '<option value="">Seleccione...</option>';
                    if (field.key === 'firmante_nombre') {
                        input.innerHTML += '<option value="Juan Pérez|Jefe">Juan Pérez</option><option value="María López|Gerente">María López</option>';
                    }
                } else if (field.type === 'date') {
                    input = document.createElement('input');
                    input.type = 'date';
                } else if (field.type === 'textarea' || field.key === 'domicilio') {
                    input = document.createElement('textarea');
                    input.rows = 2;
                } else {
                    input = document.createElement('input');
                    input.type = 'text';

                    // --- PRE-LLENADO INTELIGENTE (Auto-Fill) ---
                    if (field.key === 'ciudad') input.value = 'Lima';
                    if (field.key === 'fecha') {
                        input.type = 'date';
                        input.value = new Date().toISOString().split('T')[0];
                    }
                    if (field.key === 'numero_esquela') input.value = 'ESQ-2025-' + (crypto.getRandomValues(new Uint32Array(1))[0] % 9000 + 1000);

                    // Guardamos valor inicial
                    if (input.value) this.state.formData[field.key] = input.value;
                }

                input.id = `field-${field.key}`;
                // Evento input: Actualizar estado y previsualización
                input.oninput = (e) => this.handleInput(field.key, e.target.value);
                div.appendChild(input);
                body.appendChild(div);
            });
            sectionCard.appendChild(body);
            container.appendChild(sectionCard);
        }

        // Validación inicial post-renderizado
        setTimeout(() => {
            Object.keys(this.state.formData).forEach(k => this.validateField(k, this.state.formData[k]));
        }, 100);
    },

    /**
     * Consulta API RUC
     * Lógica asíncrona para autocompletar datos de empresa.
     */
    applyLocalRuc: function (ruc) {
        const local = this.mockRucDB[ruc];
        if (!local) return;
        this.setFieldValue('nombre', local.nombre);
        this.setFieldValue('domicilio', local.domicilio);
    },

    lookupRUC: async function (ruc) {
        if (!/^\d{11}$/.test(ruc)) return; // Regex simple RUC 11 dígitos

        // Feedback Visual en el placeholder
        const nameInput = document.getElementById('field-nombre');
        const originalPlaceholder = nameInput ? nameInput.placeholder : '';
        if (nameInput) nameInput.placeholder = "Consultando SUNAT...";

        try {
            const res = await fetch(`${API_URL}/ruc/${ruc}`);
            if (!res.ok) throw new Error("API RUC Error");
            const data = await res.json();

            if (!data.success) {
                this.applyLocalRuc(ruc); // Fallback Local Database
                return;
            }
            this.setFieldValue('nombre', data.nombre);
            this.setFieldValue('domicilio', data.domicilio);
            if (data.ciudad) this.setFieldValue('ciudad', data.ciudad);
            if (window.showToast) window.showToast('RUC encontrado: ' + data.nombre, 'success');
        } catch (e) {
            console.error(e);
            this.applyLocalRuc(ruc); // Fallback en error de red
        } finally {
            if (nameInput) nameInput.placeholder = originalPlaceholder;
        }
    },

    // Helper para asignar valor a un input y disparar eventos
    setFieldValue: function (key, value) {
        const input = document.getElementById(`field-${key}`);
        if (input) {
            input.value = value;
            this.handleInput(key, value);
        }
    },

    /**
     * Lógica de Validación Visual (Verde/Rojo)
     */
    validateField: function (key, value) {
        const input = document.getElementById(`field-${key}`);
        if (!input) return;

        // Clase .valid pone el input en verde
        if (value && value.trim() !== '') input.classList.add('valid');
        else input.classList.remove('valid');

        // Lógica para badge de "OBLIGATORIO" -> "COMPLETO" en la tarjeta
        const sectionCard = input.closest('.form-section');
        if (sectionCard) {
            const badge = sectionCard.querySelector('.badge-required');
            if (badge) {
                // Verificar si todos los inputs de la sección están llenos
                const reqs = sectionCard.querySelectorAll('input, select, textarea');
                let allOk = true;
                reqs.forEach(el => {
                    const k = el.id.replace('field-', '');
                    if (REQUIRED_FIELDS.includes(k) && (!el.value || el.value.trim() === '')) allOk = false;
                });

                if (allOk) {
                    badge.classList.add('completed');
                    badge.innerHTML = '<i class="fa-solid fa-check"></i> COMPLETO';
                } else {
                    badge.classList.remove('completed');
                    badge.textContent = 'OBLIGATORIO';
                }
            }
        }
    },

    // Manejador central de inputs
    handleInput: function (key, value) {
        this.state.formData[key] = value;

        // Gatillo automático de RUC
        if (key === 'ruc' && value.length === 11) {
            this.lookupRUC(value);
        }

        this.validateField(key, value);

        // Debounce: Esperamos 1.5s después de que el usuario deje de escribir
        // para regenerar el PDF. Esto evita sobrecargar el servidor.
        clearTimeout(this.state.debounceTimer);
        this.state.debounceTimer = setTimeout(() => {
            this.generatePreview();
        }, 1500);
    },

    // Navegación Wizard
    selectDocument: function (category) {
        this.state.docType = category;
        this.goToStep(2);
    },

    selectTemplate: async function (template) {
        this.state.selectedTemplate = template;
        document.getElementById('current-doc-label').textContent = template.name.replace('.docx', '');
        const schema = await this.fetchSchema(template.name);
        this.renderForm(schema);
        this.generatePreview();
        this.goToStep(3);
        if (window.showToast) window.showToast('Plantilla cargada: ' + template.name.replace('.docx',''), 'info');
    },

    goToStep: function (num) {
        document.querySelectorAll('.wizard-step').forEach(el => el.classList.remove('active'));
        document.getElementById('step-' + num).classList.add('active');
        this.state.step = num;
        window.scrollTo(0, 0);
    }
};

// Evento DOMContentLoaded: El punto de entrada seguro
document.addEventListener('DOMContentLoaded', () => {
    app.init();
});
