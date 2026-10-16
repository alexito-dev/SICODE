// Boletin - Calendar, Search and PDF Logic

document.addEventListener('DOMContentLoaded', () => {
    // --- State ---
    let currentDate = new Date();
    let selectedDates = new Set(); 
    let selectionStart = null; 
    let selectionEnd = null;
    let hoverDate = null; // For previewing range
    let dataCache = {}; 
    let currentSearchTerm = "";

    // --- DOM Elements ---
    const calendarGrid = document.getElementById('calendar-grid');
    const monthDisplay = document.getElementById('current-month');
    const prevMonthBtn = document.getElementById('prev-month');
    const nextMonthBtn = document.getElementById('next-month');
    const resultsList = document.getElementById('results-list');
    const resultsCount = document.getElementById('results-count');
    const pdfViewport = document.getElementById('pdf-viewport');
    const viewControls = document.getElementById('view-controls');
    const docTitle = document.getElementById('current-document-title');
    const viewportLoader = document.getElementById('viewport-loader');
    const emptyViewer = document.getElementById('empty-viewer');
    const btnDownloadMain = document.getElementById('btn-download-main');
    const btnOpenExternal = document.getElementById('btn-open-external');
    const btnClearSelection = document.getElementById('btn-clear-selection');
    const calendarLoadingBar = document.getElementById('calendar-loading-bar');
    
    
    // Search elements
    const searchWordInput = document.getElementById('search-word');
    const btnStartSearch = document.getElementById('btn-start-search');
    const progressContainerBar = document.getElementById('search-progress-container');
    const progressBar = document.getElementById('search-progress-bar');
    const progressStatusPercent = document.getElementById('search-status-percent');
    const progressStatusBadge = document.getElementById('search-status-badge');
    const progressStatusFilename = document.getElementById('search-status-filename');
    const progressStatusTime = document.getElementById('search-status-time');

    const months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

    renderCalendar();

    prevMonthBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() - 1); renderCalendar(); };
    nextMonthBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() + 1); renderCalendar(); };
    
    btnClearSelection.onclick = () => {
        selectedDates.clear();
        selectionStart = null;
        selectionEnd = null;
        hoverDate = null;
        resultsList.innerHTML = '<div style="text-align: center; padding: 40px 0; color: #94a3b8;"><i class="fa-solid fa-calendar-day" style="font-size: 24px; margin-bottom: 10px;"></i><p>Selecciona una fecha</p></div>';
        resultsCount.textContent = "0";
        renderCalendar();
    };

    // --- Booking.com Style Range Selection ---
    
    function initClickEvents(div, dateStr) {
        div.onclick = () => {
            if (!selectionStart || (selectionStart && selectionEnd)) {
                // Iniciar nueva selección
                selectionStart = dateStr;
                selectionEnd = null;
                selectedDates.clear();
                selectedDates.add(dateStr);
                updateCumulativeResults();
            } else if (selectionStart && !selectionEnd) {
                // Completar el rango - BIDIRECCIONAL
                let d1 = new Date(selectionStart + "T00:00:00");
                let d2 = new Date(dateStr + "T00:00:00");
                
                if (d2 < d1) {
                    // El usuario seleccionó una fecha anterior como final -> Swapiar
                    selectionEnd = selectionStart;
                    selectionStart = dateStr;
                } else {
                    selectionEnd = dateStr;
                }
                selectRange(selectionStart, selectionEnd);
                updateCumulativeResults();
            }
            hoverDate = null;
            renderCalendar();
        };

        div.onmouseenter = () => {
            if (selectionStart && !selectionEnd) {
                // Mostrar preview mientras decide el final del rango
                let targetDate = new Date(dateStr + "T00:00:00");
                let startDate = new Date(selectionStart + "T00:00:00");
                
                if (targetDate >= startDate) {
                    hoverDate = dateStr;
                    updateDayHighlights();
                }
            }
        };
    }

    function selectRange(start, end) {
        selectedDates.clear();
        let d1 = new Date(start + "T00:00:00");
        let d2 = new Date(end + "T00:00:00");

        let current = new Date(d1);
        while (current <= d2) {
            const y = current.getFullYear();
            const m = String(current.getMonth() + 1).padStart(2, '0');
            const d = String(current.getDate()).padStart(2, '0');
            selectedDates.add(`${y}-${m}-${d}`);
            current.setDate(current.getDate() + 1);
        }
    }

    function updateDayHighlights() {
        const allDays = document.querySelectorAll('.calendar-day[data-date]');
        allDays.forEach(div => {
            const dateStr = div.dataset.date;
            div.classList.remove('active', 'in-range', 'range-start', 'range-end');
            
            // Lógica cuando el rango ya está cerrado
            if (selectionStart && selectionEnd) {
                if (selectedDates.has(dateStr)) {
                    div.classList.add('active');
                    if (dateStr === selectionStart) div.classList.add('range-start');
                    if (dateStr === selectionEnd) div.classList.add('range-end');
                    if (dateStr !== selectionStart && dateStr !== selectionEnd) div.classList.add('in-range');
                }
            } 
            // Lógica de Previsualización (Hover)
            else if (selectionStart && !selectionEnd) {
                if (dateStr === selectionStart) {
                    div.classList.add('active', 'range-start', 'range-end');
                }
                
                if (hoverDate) {
                    let d = new Date(dateStr + "T00:00:00");
                    let s = new Date(selectionStart + "T00:00:00");
                    let h = new Date(hoverDate + "T00:00:00");
                    
                    // Preview bidireccional en hover
                    let start = s < h ? s : h;
                    let end = s < h ? h : s;

                    if (d >= start && d <= end) {
                        div.classList.add('active');
                        if (dateStr === (s < h ? selectionStart : hoverDate)) div.classList.add('range-start');
                        else if (dateStr === (s < h ? hoverDate : selectionStart)) div.classList.add('range-end');
                        else div.classList.add('in-range');
                    }
                }
            }
        });
    }

    // --- Search Logic ---
    btnStartSearch.onclick = async () => {
        const palabra = searchWordInput.value.trim();
        if (!palabra) { showToast('warning', 'Ingresa una palabra para buscar'); return; }
        if (selectedDates.size === 0) { showToast('warning', 'Selecciona al menos un día en el calendario'); return; }

        currentSearchTerm = palabra;
        let targetPdfs = getAllSelectedPdfs();
        if (targetPdfs.length === 0) { showToast('info', 'No hay boletines disponibles en los días seleccionados'); return; }

        progressContainerBar.style.display = 'block';
        if (progressStatusBadge) progressStatusBadge.textContent = `0 / ${targetPdfs.length}`;
        if (progressStatusFilename) progressStatusFilename.textContent = "Preparando...";
        resultsList.innerHTML = '<div style="text-align: center; padding: 20px;"><i class="fa-solid fa-spinner fa-spin"></i> Buscando coincidencias...</div>';
        resultsCount.textContent = "0";
        const totalToSearch = targetPdfs.length;
        let processedCount = 0;
        const startTime = Date.now();

        try {
            const res = await fetch('/boletin/api/search-bulk', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ palabra, pdfs: targetPdfs })
            });
            const reader = res.body.getReader();
            const decoder = new TextDecoder();
            resultsList.innerHTML = '';
            let foundCount = 0;

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                const chunk = decoder.decode(value);
                const lines = chunk.split('\n\n');
                lines.forEach(line => {
                    if (!line.startsWith('data: ')) return;
                    try {
                        const event = JSON.parse(line.substring(6));
                        if (event.type === 'progress') {
                            processedCount++;
                            const displayCount = Math.min(processedCount, totalToSearch);
                            const elapsed = Math.round((Date.now() - startTime) / 1000);
                            progressBar.style.width = `${event.percent}%`;
                            if (progressStatusPercent) progressStatusPercent.textContent = `${event.percent}%`;
                            if (progressStatusBadge) progressStatusBadge.textContent = `${displayCount} / ${totalToSearch}`;
                            if (progressStatusFilename) progressStatusFilename.textContent = `Analizando: ${event.archivo}`;
                            if (progressStatusTime) progressStatusTime.textContent = `(${elapsed}s)`;
                        } else if (event.type === 'found') {
                            foundCount++;
                            resultsCount.textContent = foundCount;
                            addSearchResult(event);
                        } else if (event.type === 'complete') {
                            if (progressStatusFilename) progressStatusFilename.innerHTML = '<i class="fa-solid fa-circle-check" style="color: #22c55e;"></i> Finalizado';
                            if (progressStatusBadge) progressStatusBadge.textContent = `${totalToSearch} / ${totalToSearch}`;
                            progressBar.style.width = '100%';
                            const msg = foundCount > 0
                                ? `${foundCount} coincidencia${foundCount > 1 ? 's' : ''} encontrada${foundCount > 1 ? 's' : ''}`
                                : 'Sin coincidencias en los días seleccionados';
                            showToast(foundCount > 0 ? 'success' : 'info', msg);
                            if (foundCount === 0) resultsList.innerHTML = '<div style="text-align: center; padding: 20px; color: #94a3b8;">No se encontraron coincidencias</div>';
                        }
                    } catch(e) {}
                });
            }
        } catch (e) {
            console.error(e);
            showToast('error', 'Error durante la búsqueda');
        }
    };

    function getAllSelectedPdfs() {
        let all = [];
        const sorted = Array.from(selectedDates).sort();
        sorted.forEach(dateKey => {
            const [y, m, d] = dateKey.split('-');
            const monthKey = `${y}-${m}`;
            const pdfs = dataCache[monthKey] || [];
            all = all.concat(pdfs.filter(p => p.dia === parseInt(d)));
        });
        return all;
    }

    function addSearchResult(event) {
        const pdfWithPages = { ...event.pdf, paginas: event.paginas };
        const div = document.createElement('div');
        div.className = 'result-item';
        div.dataset.pdfUrl = event.pdf.url_online;
        const parts = event.pdf.titulo.split(' - ');
        const codBoletin = parts[0] || event.pdf.titulo;
        const fechaBoletin = parts[1] || event.pdf.fecha;

        div.innerHTML = `
            <div class="result-info">
                <div class="result-title">${codBoletin}</div>
                <div class="result-date"><i class="fa-regular fa-calendar-days"></i> ${fechaBoletin}</div>
                <div class="result-search-meta">
                    <span class="match-count">${event.veces} coincidencias</span>
                    <span class="page-count" id="pages-container-${event.pdf.fecha}-${event.pdf.dia}">Págs: </span>
                </div>
            </div>
            <i class="fa-solid fa-chevron-right" style="opacity: 0.3; font-size: 10px;"></i>`;
        
        div.onclick = () => viewPdf(pdfWithPages, currentSearchTerm);
        resultsList.appendChild(div);

        // Insertar los links de páginas de forma segura (sin quotes de HTML)
        const pagesContainer = div.querySelector('.page-count');
        event.paginas.forEach((p, idx) => {
            const link = document.createElement('span');
            link.className = 'page-link';
            link.textContent = p;
            link.onclick = (e) => {
                e.stopPropagation();
                viewPdf(pdfWithPages, currentSearchTerm, p);
            };
            pagesContainer.appendChild(link);
            if (idx < event.paginas.length - 1) {
                pagesContainer.appendChild(document.createTextNode(', '));
            }
        });
    }

    async function renderCalendar() {
        calendarGrid.innerHTML = '';
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();

        // Month Selector
        const mSelect = document.getElementById('select-month');
        const ySelect = document.getElementById('select-year');
        if (mSelect) mSelect.value = month;
        if (ySelect) ySelect.value = year;

        // Header estático (fallback display)
        if (monthDisplay) monthDisplay.textContent = `${months[month]} ${year}`;

        // Añadir cabecera de días de la semana
        const dayNames = ['D', 'L', 'M', 'M', 'J', 'V', 'S'];
        dayNames.forEach(name => {
            const div = document.createElement('div');
            div.className = 'calendar-day-head';
            div.textContent = name;
            calendarGrid.appendChild(div);
        });

        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
            const div = document.createElement('div'); div.className = 'calendar-day empty';
            calendarGrid.appendChild(div);
        }

        const monthKey = `${year}-${String(month + 1).padStart(2, '0')}`;

        for (let day = 1; day <= daysInMonth; day++) {
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            const div = document.createElement('div');
            div.className = 'calendar-day';
            div.textContent = day;
            div.dataset.date = dateStr;
            
            initClickEvents(div, dateStr);
            
            const today = new Date();
            if (today.getFullYear() === year && today.getMonth() === month && today.getDate() === day) div.classList.add('today');

            calendarGrid.appendChild(div);
        }
        
        // Aplicar estado visual inicial al renderizar el mes
        updateDayHighlights();

        // Si los datos ya están en caché, pintar puntos inmediatamente
        if (dataCache[monthKey]) {
            const mData = dataCache[monthKey];
            document.querySelectorAll('.calendar-day[data-date]').forEach(d => {
                const day = parseInt(d.textContent);
                if (mData.some(p => p.dia === day)) d.classList.add('has-data');
            });
        } else {
            // Mostrar loading: skeleton gris en todos los días + barra animada
            if (calendarLoadingBar) calendarLoadingBar.classList.add('visible');
            document.querySelectorAll('.calendar-day[data-date]').forEach(d => d.classList.add('data-loading'));

            fetchMonthData(year, month + 1).then(() => {
                // Ocultar loading
                if (calendarLoadingBar) calendarLoadingBar.classList.remove('visible');

                // Si el mes que terminó de cargar es el actual, pintar puntos reales
                if (currentDate.getFullYear() === year && currentDate.getMonth() === month) {
                    const mData = dataCache[monthKey] || [];
                    document.querySelectorAll('.calendar-day[data-date]').forEach(d => {
                        d.classList.remove('data-loading');
                        const day = parseInt(d.textContent);
                        if (mData.some(p => p.dia === day)) d.classList.add('has-data');
                    });
                }
            });
        }
    }
    function updateCumulativeResults() {
        if (selectedDates.size === 0) {
            resultsList.innerHTML = '<div style="text-align: center; padding: 40px 0; color: #94a3b8;"><i class="fa-solid fa-calendar-day" style="font-size: 24px; margin-bottom: 10px;"></i><p>Selecciona una fecha</p></div>';
            resultsCount.textContent = "0";
            return;
        }

        let allPdfs = getAllSelectedPdfs();
        resultsList.innerHTML = '';
        resultsCount.textContent = allPdfs.length;

        if (allPdfs.length === 0) {
            resultsList.innerHTML = '<div style="text-align: center; padding: 40px 0; color: #94a3b8;"><p>No hay publicaciones disponibles para estas fechas</p></div>';
            return;
        }

        allPdfs.forEach(pdf => {
            const parts = pdf.titulo.split(' - ');
            const codBoletin = parts[0] || pdf.titulo;
            const fechaBoletin = parts[1] || pdf.fecha;

            const div = document.createElement('div');
            div.className = 'result-item';
            div.dataset.pdfUrl = pdf.url_online;
            div.innerHTML = `
                <div class="result-info">
                    <div class="result-title">${codBoletin}</div>
                    <div class="result-date"><i class="fa-regular fa-calendar-days"></i> ${fechaBoletin}</div>
                </div>
                <i class="fa-solid fa-chevron-right" style="opacity: 0.3; font-size: 10px;"></i>
            `;
            div.onclick = () => {
                viewPdf(pdf);
            };
            resultsList.appendChild(div);
        });

        // Insertar separador y historial si hay resultados
    }

    async function fetchMonthData(year, month) {
        const monthStr = String(month).padStart(2, '0');
        const mk = `${year}-${monthStr}`;
        if (dataCache[mk]) return;
        try {
            const res = await fetch(`/boletin/api/calendar?anio=${year}&mes=${monthStr}`);
            const data = await res.json();
            if (data.status === 'success') dataCache[mk] = data.data;
        } catch (e) {}
    }

    function viewPdf(pdf, term = "", forcePage = null) {
        emptyViewer.style.display = 'none';
        viewportLoader.style.display = 'flex';
        viewControls.style.display = 'none';
        docTitle.textContent = pdf.titulo;
        
        // Feedback visual de selección en la lista
        document.querySelectorAll('.result-item').forEach(item => {
            if (item.dataset.pdfUrl === pdf.url_online) item.classList.add('selected');
            else item.classList.remove('selected');
        });

        const existing = pdfViewport.querySelector('iframe');
        if (existing) existing.remove();
        const iframe = document.createElement('iframe');
        let viewerUrl = `/boletin/api/view?url=${encodeURIComponent(pdf.url_online)}&fecha=${pdf.fecha}&titulo=${encodeURIComponent(pdf.titulo)}`;
        if (term) viewerUrl += `&highlight=${encodeURIComponent(term)}`;
        
        let fragment = "";
        if (forcePage) fragment = `#page=${forcePage}&zoom=65`;
        else if (pdf.paginas && pdf.paginas.length > 0) fragment = `#page=${pdf.paginas[0]}&zoom=65`;
        else if (term) fragment = `#search="${encodeURIComponent(term)}"&zoom=65`;
        else fragment = "#zoom=65";
        
        iframe.src = viewerUrl + fragment;
        iframe.onload = () => { viewportLoader.style.display = 'none'; viewControls.style.display = 'flex'; };
        btnDownloadMain.onclick = () => {
            window.location.href = `/boletin/api/download?url=${encodeURIComponent(pdf.url_online)}&fecha=${pdf.fecha}&titulo=${encodeURIComponent(pdf.titulo)}`;
            showToast('info', 'Descargando boletín...');
        };
        btnOpenExternal.onclick = () => {
            window.open(pdf.url_online, '_blank');
            showToast('info', 'Abriendo en nueva pestaña...');
        };
        pdfViewport.appendChild(iframe);
    }

    function showToast(type, message) {
        if (window.showToast) {
            window.showToast(message, type);
        }
    }

    // --- Nav Selectors Logic ---
    const sMonth = document.getElementById('select-month');
    const sYear = document.getElementById('select-year');
    
    if (sMonth && sYear) {
        sMonth.onchange = () => { currentDate.setMonth(parseInt(sMonth.value)); renderCalendar(); };
        sYear.onchange = () => { currentDate.setFullYear(parseInt(sYear.value)); renderCalendar(); };
    }

});
