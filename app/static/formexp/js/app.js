/* ═══════════════════════════════════════════════════
   SUNAT | Premium Registration App - Initialization
   ═══════════════════════════════════════════════════ */

let currentMode = 'formulario';

// ─── Mode Switching ───────────────────────────────
function setMode(mode) {
    currentMode = mode;

    const formContainer = document.getElementById('forms-wrapper');
    formContainer.style.opacity = '0';

    setTimeout(() => {
        // Toggle forms
        document.getElementById('form-formulario').classList.toggle('hidden', mode !== 'formulario');
        document.getElementById('form-expediente').classList.toggle('hidden', mode !== 'expediente');

        // Toggle button styles
        document.getElementById('btn-formulario').classList.toggle('active', mode === 'formulario');
        document.getElementById('btn-expediente').classList.toggle('active', mode === 'expediente');

        formContainer.style.opacity = '1';

        // Focus first field
        if (mode === 'formulario') {
            document.getElementById('f-documento').focus();
        } else {
            document.getElementById('e-p1').focus();
        }
    }, 150);
}

// ─── Quick Helpers ────────────────────────────────
function fillToday(inputId) {
    const now = new Date();
    const dd = String(now.getDate()).padStart(2, '0');
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const yyyy = now.getFullYear();

    const input = document.getElementById(inputId);
    input.value = `${dd}/${mm}/${yyyy}`;
    input.classList.add('valid');

    // Trigger nice animation
    const icon = input.previousElementSibling;
    icon.style.transform = 'scale(1.2)';
    setTimeout(() => icon.style.transform = 'scale(1)', 200);
}

function formatDateInput(input) {
    let raw = input.value.replace(/\D/g, ''); // Solo números
    let formatted = '';

    if (raw.length > 0) {
        formatted += raw.substring(0, 2);
    }
    if (raw.length >= 2) {
        // Añadir slash solo si el usuario está tipeando un mes, no si borró el slash.
        if (input.value.length > formatted.length || raw.length > 2) {
            formatted += '/';
        }
    }
    if (raw.length > 2) {
        formatted += raw.substring(2, 4);
    }
    if (raw.length >= 4) {
        if (input.value.length > formatted.length || raw.length > 4) {
            formatted += '/';
        }
    }
    if (raw.length > 4) {
        formatted += raw.substring(4, 8);
    }

    input.value = formatted;
}

function autoTab(current, nextId, maxLen) {
    if (current.value.length >= maxLen) {
        document.getElementById(nextId).focus();
    }
    updateSmartStatus();
}

function updateSmartStatus() {
    const p1 = document.getElementById('e-p1').value.trim();
    const p2 = document.getElementById('e-p2').value.trim();
    const p3 = document.getElementById('e-p3').value.trim();
    const p4 = document.getElementById('e-p4').value.trim();

    const statusIcon = document.querySelector('#e-smart-status i');
    if (p1 && p2 && p3 && p4) {
        statusIcon.className = 'ph-fill ph-check-circle';
        statusIcon.parentElement.style.color = 'var(--success)';
        statusIcon.parentElement.style.opacity = '1';
    } else {
        statusIcon.className = 'ph ph-check-circle';
        statusIcon.parentElement.style.color = 'var(--text-subtle)';
        statusIcon.parentElement.style.opacity = '0.3';
    }
}

// ─── Validation ───────────────────────────────────
function validateRUC(input) {
    const val = input.value.replace(/\D/g, '');
    input.value = val;

    const hintId = input.id.replace('ruc', 'ruc-hint');
    const hint = document.getElementById(hintId);

    if (val.length === 0) {
        input.className = '';
        hint.textContent = '';
        hint.className = 'validation-msg';
    } else if (val.length === 11) {
        input.className = 'valid';
        hint.textContent = '✓ Válido';
        hint.className = 'validation-msg ok';
    } else {
        input.className = 'invalid';
        hint.textContent = `${val.length}/11`;
        hint.className = 'validation-msg err';
    }
}

function shakeField(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.add('invalid');
    el.focus();
    setTimeout(() => el.classList.remove('invalid'), 400);
}

// ─── API Submissions ──────────────────────────────
async function saveFormulario(event) {
    event.preventDefault();
    const btn = document.getElementById('save-formulario');
    const originalText = btn.innerHTML;

    const doc = document.getElementById('f-documento').value.trim();
    const ruc = document.getElementById('f-ruc').value.trim();

    if (!doc) return shakeField('f-documento');
    if (!ruc || ruc.length !== 11) return shakeField('f-ruc');

    const body = {
        documento: doc,
        razon_social: document.getElementById('f-razon').value.trim(),
        ruc: ruc,
        fecha: document.getElementById('f-fecha').value.trim(),
        observaciones: document.getElementById('f-obs').value.trim(),
    };

    btn.innerHTML = '<i class="ph ph-spinner-gap" style="animation: spin 1s linear infinite;"></i> Guardando...';

    try {
        const res = await fetch('/formexp/api/formulario', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });
        const data = await res.json();

        if (data.success) {
            showToast('Formulario 194 registrado', 'success', data.live);
            btn.classList.add('success-state');
            btn.innerHTML = '<i class="ph ph-check"></i> ¡Guardado!';
            document.getElementById('count-formulario').textContent = data.count;

            setTimeout(() => {
                clearForm();
                btn.classList.remove('success-state');
                btn.innerHTML = originalText;
            }, 1500);
        }
    } catch (err) {
        showToast('Error al procesar: ' + err.message, 'error');
        btn.innerHTML = originalText;
    }
    return false;
}

async function saveExpediente(event) {
    event.preventDefault();
    const btn = document.getElementById('save-expediente');
    const originalText = btn.innerHTML;

    const p1 = document.getElementById('e-p1').value.trim();
    const p2 = document.getElementById('e-p2').value.trim();
    const p3 = document.getElementById('e-p3').value.trim();
    const p4 = document.getElementById('e-p4').value.trim();
    const ruc = document.getElementById('e-ruc').value.trim();

    if (!p1 || !p2 || !p3 || !p4) return shakeField('e-p1');
    if (!ruc || ruc.length !== 11) return shakeField('e-ruc');

    const body = {
        parte1: p1, parte2: p2, parte3: p3, parte4: p4,
        razon_social: document.getElementById('e-razon').value.trim(),
        ruc: ruc,
        fecha: document.getElementById('e-fecha').value.trim(),
        observaciones: document.getElementById('e-obs').value.trim(),
    };

    btn.innerHTML = '<i class="ph ph-spinner-gap" style="animation: spin 1s linear infinite;"></i> Guardando...';

    try {
        const res = await fetch('/formexp/api/expediente', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });
        const data = await res.json();

        if (data.success) {
            showToast('Expediente registrado', 'success', data.live);
            btn.classList.add('success-state');
            btn.innerHTML = '<i class="ph ph-check"></i> ¡Guardado!';
            document.getElementById('count-expediente').textContent = data.count;

            setTimeout(() => {
                clearForm();
                btn.classList.remove('success-state');
                btn.innerHTML = originalText;
            }, 1500);
        }
    } catch (err) {
        showToast('Error al procesar: ' + err.message, 'error');
        btn.innerHTML = originalText;
    }
    return false;
}

// ─── Reset ────────────────────────────────────────
function clearForm() {
    document.querySelectorAll('input').forEach(el => {
        el.value = '';
        el.className = '';
        if (el.id.startsWith('e-p')) el.className = el.id === 'e-p3' ? 'wide' : (el.id === 'e-p4' ? 'narrow' : '');
    });

    document.querySelectorAll('.validation-msg').forEach(el => {
        el.textContent = '';
        el.className = 'validation-msg';
    });

    updateSmartStatus();

    if (currentMode === 'formulario') {
        document.getElementById('f-documento').focus();
    } else {
        document.getElementById('e-p1').focus();
    }
}

// ─── UI Status / Toast ────────────────────────────
function showToast(message, type, isLive = false) {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let icon = type === 'success' ? '<i class="ph-fill ph-check-circle"></i>' : '<i class="ph-fill ph-warning-circle"></i>';
    let liveTag = isLive ? '<span style="color:#2E7D32; font-size:1.1rem" title="Tiempo real"><i class="ph-fill ph-lightning"></i></span>' : '';

    toast.innerHTML = `${icon} <span>${message}</span> ${liveTag}`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

async function checkConnection() {
    try {
        const res = await fetch('/formexp/api/status');
        const data = await res.json();

        const pill = document.getElementById('connection-status');
        const text = document.getElementById('status-text');

        if (data.excel_connected) {
            pill.className = 'status-pill connected';
            text.innerHTML = 'Excel Activo <i class="ph-fill ph-lightning"></i>';
        } else if (data.has_xlwings) {
            pill.className = 'status-pill waiting';
            text.textContent = 'Abre el Excel para sincronizar';
        } else {
            pill.className = 'status-pill default';
            text.innerHTML = '<i class="ph ph-file-xls"></i> Modo Archivo';
        }

        document.getElementById('count-formulario').textContent = data.formulario_count;
        document.getElementById('count-expediente').textContent = data.expediente_count;

    } catch (e) {
        console.log("Status API down");
    }
}

// Add spin animation CSS dynamically
const style = document.createElement('style');
style.textContent = `@keyframes spin { 100% { transform: rotate(360deg); } }`;
document.head.appendChild(style);

// Init
document.addEventListener('DOMContentLoaded', () => {
    checkConnection();
    setInterval(checkConnection, 3000);

    // Add smooth form transitions
    document.getElementById('forms-wrapper').style.transition = 'opacity 0.15s ease';
});
