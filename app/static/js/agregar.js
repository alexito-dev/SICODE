/* ════════════════════════════════════════════════════
   MÓDULO AGREGAR — Frontend Logic
   Prefijo 'ag' en funciones para evitar conflictos.
   ════════════════════════════════════════════════════ */

let agCurrentMode = 'formulario';
let agConfigOpen  = false;
let agCurrentTab  = 'upload';

// ══════════════════════════════════════════════════════
//  PANEL EXCEL — Abrir/Cerrar
// ══════════════════════════════════════════════════════
function agToggleExcelConfig() {
    agConfigOpen = !agConfigOpen;
    const panel = document.getElementById('ag-excel-config');
    const btn   = document.getElementById('ag-btn-change-toggle');
    if (agConfigOpen) {
        panel.classList.remove('hidden');
        btn.classList.add('active');
        btn.innerHTML = '<i class="ph ph-x"></i> Cerrar';
    } else {
        panel.classList.add('hidden');
        btn.classList.remove('active');
        btn.innerHTML = '<i class="ph ph-pencil-simple"></i> Cambiar';
    }
}

function agSwitchTab(tab) {
    agCurrentTab = tab;
    ['upload','path','new'].forEach(t => {
        document.getElementById(`ag-tab-${t}`).classList.toggle('active', t === tab);
        document.getElementById(`ag-body-${t}`).classList.toggle('hidden', t !== tab);
    });
}

// ══════════════════════════════════════════════════════
//  SUBIR ARCHIVO
// ══════════════════════════════════════════════════════
function agHandleFileSelect(input) {
    if (input.files && input.files[0]) {
        agUploadFile(input.files[0]);
    }
}

// Drag & drop
document.addEventListener('DOMContentLoaded', () => {
    const zone = document.getElementById('ag-dropzone');
    if (!zone) return;

    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file) agUploadFile(file);
    });
});

async function agUploadFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        agShowToast('Solo se admiten archivos .xlsx o .xls', 'error');
        return;
    }
    const content  = document.getElementById('ag-dropzone-content');
    const progress = document.getElementById('ag-upload-progress');
    content.classList.add('hidden');
    progress.classList.remove('hidden');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const res  = await fetch('/agregar/api/upload-excel', { method:'POST', body:formData });
        const data = await res.json();
        if (data.success) {
            agShowToast(`✓ Archivo cargado: ${data.excel_filename} (${data.size_kb} KB)`, 'success');
            agUpdateExcelDisplay(data.excel_filename, data.excel_path);
            agToggleExcelConfig(); // cerrar panel
            agCheckConnection();
        } else {
            agShowToast(data.error || 'Error al subir archivo', 'error');
        }
    } catch(e) {
        agShowToast('Error de conexión: ' + e.message, 'error');
    } finally {
        content.classList.remove('hidden');
        progress.classList.add('hidden');
    }
}

// ══════════════════════════════════════════════════════
//  RUTA DEL SERVIDOR
// ══════════════════════════════════════════════════════
async function agSetPath() {
    const path = document.getElementById('ag-server-path').value.trim();
    if (!path) { agShowToast('Introduce una ruta válida', 'error'); return; }

    const formData = new FormData();
    formData.append('path', path);
    try {
        const res  = await fetch('/agregar/api/set-path', { method:'POST', body:formData });
        const data = await res.json();
        if (data.success) {
            agShowToast(`✓ Excel configurado: ${data.excel_filename}`, 'success');
            agUpdateExcelDisplay(data.excel_filename, data.excel_path);
            agToggleExcelConfig();
            agCheckConnection();
        } else {
            agShowToast(data.error || 'Error al configurar ruta', 'error');
        }
    } catch(e) {
        agShowToast('Error de conexión: ' + e.message, 'error');
    }
}

// ══════════════════════════════════════════════════════
//  CREAR NUEVO EXCEL
// ══════════════════════════════════════════════════════
async function agCreateNew() {
    const filename = document.getElementById('ag-new-filename').value.trim() || 'FORMULARIO194_MESADEPARTES.xlsx';
    const formData = new FormData();
    formData.append('filename', filename);
    try {
        const res  = await fetch('/agregar/api/new-excel', { method:'POST', body:formData });
        const data = await res.json();
        if (data.success) {
            agShowToast(`✓ Excel creado: ${data.excel_filename}`, 'success');
            agUpdateExcelDisplay(data.excel_filename, data.excel_path);
            agToggleExcelConfig();
            agCheckConnection();
        } else {
            agShowToast(data.error || 'Error al crear Excel', 'error');
        }
    } catch(e) {
        agShowToast('Error de conexión: ' + e.message, 'error');
    }
}

// ══════════════════════════════════════════════════════
//  ACTUALIZAR DISPLAY DE ARCHIVO ACTIVO
// ══════════════════════════════════════════════════════
function agUpdateExcelDisplay(filename, fullPath) {
    document.getElementById('ag-excel-filename').textContent = filename || '—';
    document.getElementById('ag-excel-path').textContent    = fullPath || '';
}

// ══════════════════════════════════════════════════════
//  MODE SWITCHING (Formulario / Expediente)
// ══════════════════════════════════════════════════════
function agSetMode(mode) {
    agCurrentMode = mode;
    const wrapper = document.getElementById('ag-forms-wrapper');
    wrapper.style.opacity = '0';
    setTimeout(() => {
        document.getElementById('ag-form-formulario').classList.toggle('hidden', mode !== 'formulario');
        document.getElementById('ag-form-expediente').classList.toggle('hidden', mode !== 'expediente');
        document.getElementById('ag-btn-formulario').classList.toggle('active', mode === 'formulario');
        document.getElementById('ag-btn-expediente').classList.toggle('active', mode === 'expediente');
        wrapper.style.opacity = '1';
        document.getElementById(mode === 'formulario' ? 'ag-f-documento' : 'ag-e-p1').focus();
    }, 150);
}

// ══════════════════════════════════════════════════════
//  DATE HELPERS
// ══════════════════════════════════════════════════════
function agFillToday(inputId) {
    const now = new Date();
    const dd = String(now.getDate()).padStart(2,'0');
    const mm = String(now.getMonth()+1).padStart(2,'0');
    const input = document.getElementById(inputId);
    input.value = `${dd}/${mm}/${now.getFullYear()}`;
    input.classList.add('valid');
    const icon = input.previousElementSibling;
    if (icon) { icon.style.transform = 'scale(1.2)'; setTimeout(()=>icon.style.transform='scale(1)',200); }
}

function agFormatDate(input) {
    const raw = input.value.replace(/\D/g,'');
    let formatted = '';
    if (raw.length > 0) formatted += raw.substring(0,2);
    if (raw.length >= 2 && (input.value.length > formatted.length || raw.length > 2)) formatted += '/';
    if (raw.length > 2) formatted += raw.substring(2,4);
    if (raw.length >= 4 && (input.value.length > formatted.length || raw.length > 4)) formatted += '/';
    if (raw.length > 4) formatted += raw.substring(4,8);
    input.value = formatted;
}

// ══════════════════════════════════════════════════════
//  SMART EXPEDIENTE
// ══════════════════════════════════════════════════════
function agAutoTab(current, nextId, maxLen) {
    if (current.value.length >= maxLen) document.getElementById(nextId).focus();
    agUpdateSmartStatus();
}
function agUpdateSmartStatus() {
    const all = ['ag-e-p1','ag-e-p2','ag-e-p3','ag-e-p4'].every(id => document.getElementById(id)?.value.trim());
    const icon = document.querySelector('#ag-e-smart-status i');
    if (!icon) return;
    if (all) {
        icon.className = 'ph-fill ph-check-circle';
        icon.parentElement.style.color   = 'var(--ag-success)';
        icon.parentElement.style.opacity = '1';
    } else {
        icon.className = 'ph ph-check-circle';
        icon.parentElement.style.color   = 'var(--ag-text-subtle)';
        icon.parentElement.style.opacity = '0.3';
    }
}

// ══════════════════════════════════════════════════════
//  VALIDATION
// ══════════════════════════════════════════════════════
function agValidateRUC(input) {
    const val = input.value.replace(/\D/g,'');
    input.value = val;
    const hint = document.getElementById(input.id.replace('ruc','ruc-hint'));
    if (!hint) return;
    if (!val.length)    { input.className=''; hint.textContent=''; hint.className='ag-validation-msg'; }
    else if (val.length===11) { input.className='valid'; hint.textContent='✓ Válido'; hint.className='ag-validation-msg ok'; }
    else                { input.className='invalid'; hint.textContent=`${val.length}/11`; hint.className='ag-validation-msg err'; }
}
function agShakeField(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.add('invalid'); el.focus();
    setTimeout(()=>el.classList.remove('invalid'),400);
}

// ══════════════════════════════════════════════════════
//  API SUBMISSIONS
// ══════════════════════════════════════════════════════
async function agSaveFormulario(event) {
    event.preventDefault();
    const btn  = document.getElementById('ag-save-formulario');
    const orig = btn.innerHTML;
    const doc  = document.getElementById('ag-f-documento').value.trim();
    const ruc  = document.getElementById('ag-f-ruc').value.trim();
    if (!doc) return agShakeField('ag-f-documento');
    if (!ruc || ruc.length!==11) return agShakeField('ag-f-ruc');

    btn.innerHTML = '<i class="ph ph-spinner-gap" style="animation:agSpin 1s linear infinite"></i> Guardando...';
    try {
        const res  = await fetch('/agregar/api/formulario',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
            documento:doc, razon_social:document.getElementById('ag-f-razon').value.trim(),
            ruc, fecha:document.getElementById('ag-f-fecha').value.trim(),
            observaciones:document.getElementById('ag-f-obs').value.trim()
        })});
        const data = await res.json();
        if (data.success) {
            agShowToast('Formulario 194 registrado','success',data.live);
            btn.classList.add('ag-success-state');
            btn.innerHTML='<i class="ph ph-check"></i> ¡Guardado!';
            document.getElementById('ag-count-formulario').textContent = data.count;
            setTimeout(()=>{ agClearForm(); btn.classList.remove('ag-success-state'); btn.innerHTML=orig; },1500);
        }
    } catch(err) { agShowToast('Error: '+err.message,'error'); btn.innerHTML=orig; }
    return false;
}

async function agSaveExpediente(event) {
    event.preventDefault();
    const btn  = document.getElementById('ag-save-expediente');
    const orig = btn.innerHTML;
    const p    = ['ag-e-p1','ag-e-p2','ag-e-p3','ag-e-p4'].map(id=>document.getElementById(id).value.trim());
    const ruc  = document.getElementById('ag-e-ruc').value.trim();
    if (p.some(v=>!v)) return agShakeField('ag-e-p1');
    if (!ruc||ruc.length!==11) return agShakeField('ag-e-ruc');

    btn.innerHTML='<i class="ph ph-spinner-gap" style="animation:agSpin 1s linear infinite"></i> Guardando...';
    try {
        const res  = await fetch('/agregar/api/expediente',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
            parte1:p[0],parte2:p[1],parte3:p[2],parte4:p[3],
            razon_social:document.getElementById('ag-e-razon').value.trim(),
            ruc,fecha:document.getElementById('ag-e-fecha').value.trim(),
            observaciones:document.getElementById('ag-e-obs').value.trim()
        })});
        const data = await res.json();
        if (data.success) {
            agShowToast('Expediente registrado','success',data.live);
            btn.classList.add('ag-success-state');
            btn.innerHTML='<i class="ph ph-check"></i> ¡Guardado!';
            document.getElementById('ag-count-expediente').textContent = data.count;
            setTimeout(()=>{ agClearForm(); btn.classList.remove('ag-success-state'); btn.innerHTML=orig; },1500);
        }
    } catch(err) { agShowToast('Error: '+err.message,'error'); btn.innerHTML=orig; }
    return false;
}

// ══════════════════════════════════════════════════════
//  CLEAR FORM
// ══════════════════════════════════════════════════════
function agClearForm() {
    document.querySelectorAll('#ag-form-formulario input,#ag-form-expediente input').forEach(el=>{
        el.value=''; el.className='';
        if (el.id==='ag-e-p3') el.className='wide';
        if (el.id==='ag-e-p4') el.className='narrow';
    });
    document.querySelectorAll('.ag-validation-msg').forEach(el=>{el.textContent='';el.className='ag-validation-msg';});
    agUpdateSmartStatus();
    document.getElementById(agCurrentMode==='formulario'?'ag-f-documento':'ag-e-p1').focus();
}

// ══════════════════════════════════════════════════════
//  TOAST (Sistema Global)
// ══════════════════════════════════════════════════════
function agShowToast(message, type, isLive=false) {
    if (window.showToast) {
        window.showToast(message, type);
    }
}


// ══════════════════════════════════════════════════════
//  CONNECTION STATUS
// ══════════════════════════════════════════════════════
async function agCheckConnection() {
    try {
        const res  = await fetch('/agregar/api/status');
        const data = await res.json();
        const pill = document.getElementById('ag-connection-status');
        const text = document.getElementById('ag-status-text');
        if (data.excel_connected) {
            pill.className='ag-status-pill connected';
            text.innerHTML='Excel Activo <i class="ph-fill ph-lightning"></i>';
        } else if (data.has_xlwings) {
            pill.className='ag-status-pill waiting';
            text.textContent='Abre el Excel para sincronizar';
        } else {
            pill.className='ag-status-pill';
            text.innerHTML='<i class="ph ph-file-xls"></i> Modo Archivo';
        }
        document.getElementById('ag-count-formulario').textContent = data.formulario_count;
        document.getElementById('ag-count-expediente').textContent = data.expediente_count;
        // Actualizar display del archivo activo
        if (data.excel_filename) agUpdateExcelDisplay(data.excel_filename, data.excel_path);
    } catch(e) { /* sin conexión */ }
}

// Spin keyframes
const agStyle = document.createElement('style');
agStyle.textContent='@keyframes agSpin{100%{transform:rotate(360deg)}}';
document.head.appendChild(agStyle);

// ── Init ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', ()=>{
    document.getElementById('ag-forms-wrapper').style.transition='opacity 0.15s ease';
    agCheckConnection();
    setInterval(agCheckConnection, 3000);
});
