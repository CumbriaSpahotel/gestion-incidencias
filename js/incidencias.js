// Configuration
const CONFIG = {
    defaultUrl: "",
    localFileName: "1.4.2 Cuestionario de queja o sugerencia.xlsx",
    refreshInterval: 300000,
    dbName: 'incidencias_db'
};

// State
let STATE = {
    incidencias: [],
    lastUpdate: null,
    filters: { search: '', hotel: 'all', estado: 'all' }
};

window.refreshIntervalId = null;

// Charts Variables (Global)
let chartByHotel = null;
let chartTrend = null;

// IndexedDB Helper for Persistent File Handle
const DB_NAME = 'IncidenciasFilesDB';
const STORE_NAME = 'handles';
const LocalDB = {
    async get(key) {
        return new Promise((resolve) => {
            const r = indexedDB.open(DB_NAME, 1);
            r.onupgradeneeded = e => e.target.result.createObjectStore(STORE_NAME);
            r.onsuccess = e => {
                const tx = e.target.result.transaction(STORE_NAME, 'readonly');
                const req = tx.objectStore(STORE_NAME).get(key);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => resolve(null);
            };
            r.onerror = () => resolve(null);
        });
    },
    async set(key, val) {
        return new Promise((resolve) => {
            const r = indexedDB.open(DB_NAME, 1);
            r.onupgradeneeded = e => e.target.result.createObjectStore(STORE_NAME);
            r.onsuccess = e => {
                const tx = e.target.result.transaction(STORE_NAME, 'readwrite');
                tx.objectStore(STORE_NAME).put(val, key);
                tx.oncomplete = () => resolve();
            };
        });
    }
};

let fileHandle = null;

// Mappings (Strict Matches)
const FIELD_MAP = {
    id_original: ['Id'],
    fecha_creacion: ['Hora de Inicio'],
    fecha_cierre: ['Hora de finalización'],
    correo: ['Correo electrónico'],
    usuario_registro: ['Indícanos tu nombre'],
    hotel: ['¿Dónde ha ocurrido', 'el establecimiento'],
    tipo: ['¿Qué tipo de regi'],
    departamento: ['Zona o servicio'],
    cliente: ['Nombre completo del Cliente'],
    descripcion: ['¿Qué incidencia', 'incidencia, queja o sugerencia', 'relevante', 'Cuéntanos', 'Descripción'],
    accion: ['¿Se ha dado una sol'],
    solicita_respuesta: ['Solicita Respuesta', 'solicitado respuesta'],
    telefono: ['Teléfono de Contacto'],
    correo_respuesta: ['correo electrónico a continuación'],
    responsable: ['Responsable'],
    estado: ['Estado']
};

document.addEventListener('DOMContentLoaded', async () => {
    checkMigration();
    initLocalState();
    setupEventListeners();
    setupCharts();

    // Attempt local handle restore first
    await checkSavedHandle();

    window.refreshIntervalId = setInterval(() => {
        if (fileHandle) readLocalFile(true); // Silent refresh
        else if (canUseNetworkSync()) loadData(true);
    }, CONFIG.refreshInterval);
});

function isFilePage() {
    return window.location.protocol === 'file:';
}

function canUseNetworkSync() {
    return Boolean(getConfiguredSourceUrl()) && !isFilePage();
}

function checkMigration() {
    const current = localStorage.getItem('source_url');
    if (isUnsupportedSourceUrl(current)) {
        localStorage.removeItem('source_url');
    }
}

function getConfiguredSourceUrl() {
    const url = (localStorage.getItem('source_url') || CONFIG.defaultUrl || '').trim();
    return isUnsupportedSourceUrl(url) ? '' : url;
}

function isUnsupportedSourceUrl(url) {
    if (!url) return false;
    return url.includes('1drv.ms/') || url.includes('onedrive.live.com/') || url.includes('forms.cloud.microsoft/');
}

function setupCharts() { renderCharts(); }

function initLocalState() {
    const saved = localStorage.getItem(CONFIG.dbName);
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            STATE.incidencias = parsed.items || [];
            STATE.lastUpdate = parsed.lastUpdate;
            updateLastUpdateUI();
            renderDashboard();
        } catch (e) { console.error(e); }
    }
}

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', (e) => {
        STATE.filters.search = e.target.value.toLowerCase(); renderTable();
    });
    document.getElementById('filterHotel').addEventListener('change', (e) => {
        STATE.filters.hotel = e.target.value; renderTable();
    });
    document.getElementById('filterEstado').addEventListener('change', (e) => {
        STATE.filters.estado = e.target.value; renderTable();
    });

    // Smart Refresh
    document.getElementById('btnRefresh').addEventListener('click', () => {
        if (fileHandle) readLocalFile();
        else if (canUseNetworkSync()) loadData();
        else showLocalModeUI();
    });

    document.getElementById('btnExport').addEventListener('click', exportToExcel);

    // Legacy File Input Fallback
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);

    document.getElementById('btnUpload').addEventListener('click', (event) => {
        event.preventDefault();
        pickLocalFile();
    });

    document.getElementById('btnConfig').addEventListener('click', configureSourceUrl);
    document.getElementById('btnNewIncident').addEventListener('click', () => {
        window.location.href = 'RegistroIncidencia.html';
    });
    document.getElementById('btnCloseModal').addEventListener('click', closeIncidentModal);
    document.getElementById('btnCloseFormModal').addEventListener('click', closeIncidentFormModal);
    document.getElementById('btnCancelForm').addEventListener('click', closeIncidentFormModal);
    document.getElementById('incidentForm').addEventListener('submit', handleIncidentFormSubmit);
    document.getElementById('incidentModal').addEventListener('click', (event) => {
        if (event.target.id === 'incidentModal') closeIncidentModal();
    });
    document.getElementById('incidentFormModal').addEventListener('click', (event) => {
        if (event.target.id === 'incidentFormModal') closeIncidentFormModal();
    });
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') closeIncidentModal();
    });
}

// Smart Local File Logic
async function checkSavedHandle() {
    try {
        const handle = await LocalDB.get('excelHandle');
        if (handle) {
            fileHandle = handle;
            // Check permission
            const opts = { mode: 'read' };
            if ((await fileHandle.queryPermission(opts)) === 'granted') {
                await readLocalFile();
            } else {
                showReconnectUI(); // Ask user to reconnect
            }
        } else {
            if (canUseNetworkSync()) await loadData();
            else showLocalModeUI();
        }
    } catch (e) {
        if (canUseNetworkSync()) await loadData();
        else showLocalModeUI();
    }
}

async function pickLocalFile() {
    // Call this from a button execution context
    try {
        if (window.showOpenFilePicker) {
            [fileHandle] = await window.showOpenFilePicker({
                types: [{ description: 'Excel Files', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
                multiple: false
            });
            await LocalDB.set('excelHandle', fileHandle);
            await readLocalFile();
            // Remove banners
            const banners = document.querySelectorAll('#errorBanner, #reconnectBanner');
            banners.forEach(b => b.remove());
        } else {
            document.getElementById('fileInput').click(); // Fallback
        }
    } catch (err) { console.error(err); }
}

async function readLocalFile(silent = false) {
    if (!fileHandle) return;
    if (!silent) showLoading(true);
    try {
        const file = await fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        processWorkbook(workbook, silent ? null : "Sincronizado (Local)");
        document.getElementById('lastUpdateText').innerHTML = `Act: ${new Date().toLocaleTimeString()} (Local)`;

        // Clear Error Banner if exists
        const banner = document.getElementById('errorBanner');
        if (banner) banner.remove();

    } catch (e) {
        console.error("Local read error", e);
        if (!silent) showToast("Error lectura local", "error");
    } finally {
        if (!silent) showLoading(false);
    }
}

function showReconnectUI() {
    const header = document.querySelector('.header');
    if (!document.getElementById('reconnectBanner')) {
        const banner = document.createElement('div');
        banner.id = 'reconnectBanner';
        banner.style.cssText = 'background:#f0f9ff; color:#0369a1; padding:1rem; border-radius:0.5rem; margin-bottom:1rem; border:1px solid #bae6fd; display:flex; gap:1rem; align-items:center;';
        banner.innerHTML = `
            <i class="fa-solid fa-link" style="font-size:1.5rem"></i>
            <div>
                <strong>Conexión Local Detectada</strong><br>
                Permita el acceso para restablecer la actualización automática con su archivo.
            </div>
            <button class="btn btn-primary" id="btnReconnect" style="margin-left:auto; font-size:0.85rem">Reconectar</button>
        `;
        header.parentNode.insertBefore(banner, header.nextSibling);

        document.getElementById('btnReconnect').addEventListener('click', async () => {
            if (fileHandle && (await fileHandle.requestPermission({ mode: 'read' })) === 'granted') {
                banner.remove();
                readLocalFile();
            }
        });
    }
}

function showLocalModeUI() {
    const header = document.querySelector('.header');
    if (!header || document.getElementById('localModeBanner') || document.getElementById('reconnectBanner')) return;

    const banner = document.createElement('div');
    banner.id = 'localModeBanner';
    banner.style.cssText = 'background:#f0f9ff; color:#0369a1; padding:1rem; border-radius:0.5rem; margin-bottom:1rem; border:1px solid #bae6fd; display:flex; gap:1rem; align-items:center;';
    banner.innerHTML = `
        <i class="fa-solid fa-folder-open" style="font-size:1.5rem"></i>
        <div style="flex:1">
            <strong>Seleccione el Excel</strong><br>
            <div style="font-size:0.875rem;">
                Puede registrar incidencias desde la web o cargar un Excel local para importar datos existentes.
            </div>
        </div>
        <button class="btn btn-primary" id="btnOpenLocalFile" style="font-size:0.85rem">
            <i class="fa-solid fa-file-excel"></i> Abrir Excel
        </button>
    `;
    header.parentNode.insertBefore(banner, header.nextSibling);
    document.getElementById('btnOpenLocalFile').addEventListener('click', pickLocalFile);
    document.getElementById('lastUpdateText').innerText = 'Pendiente de Excel local';
}

// Legacy Handler
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            processWorkbook(workbook, "Carga Local Exitosa");
            const banner = document.getElementById('localModeBanner') || document.getElementById('errorBanner');
            if (banner) banner.remove();
        } catch (error) {
            console.error(error);
            showToast("Error al leer archivo", "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

function configureSourceUrl() {
    const current = getConfiguredSourceUrl();
    const next = prompt('Pegue una URL directa y pública del Excel. Los enlaces de OneDrive o Forms no funcionan aquí por seguridad del navegador. Si usa datos locales, deje este campo vacío.', current);
    if (next === null) return;

    const clean = next.trim();
    if (clean) {
        if (isUnsupportedSourceUrl(clean)) {
            localStorage.removeItem('source_url');
            showLocalModeUI();
            showToast("Ese enlace no se puede leer desde la web. Use Cargar Local o registre incidencias directamente.", "error");
            return;
        }
        localStorage.setItem('source_url', clean);
        fileHandle = null;
        if (canUseNetworkSync()) loadData();
        else showToast("URL guardada. Abra la página desde un servidor local para sincronizar por URL.", "success");
    } else {
        localStorage.removeItem('source_url');
        showLocalModeUI();
    }
}

// Network Load
async function loadData(silent = false) {
    if (isFilePage()) {
        showLocalModeUI();
        return;
    }

    const url = getConfiguredSourceUrl();
    if (!url && !canTryBundledWorkbook()) {
        showLocalModeUI();
        return;
    }

    if (!silent) showLoading(true);
    try {
        let response = await tryLocalWorkbook();
        let loadedFromLocalWorkbook = Boolean(response);
        let usedProxy = false;

        if (!response) {
            try {
                response = await fetch(url);
                if (!response.ok) throw new Error("Direct fail");
                if (response.headers.get("content-type")?.includes("text/html")) throw new Error("Auth Wall");
            } catch (directError) {
                console.log("Using Proxy...");
                usedProxy = true;
                let proxyTarget = url;
                if (!proxyTarget.includes('download=1')) proxyTarget += (proxyTarget.includes('?') ? '&' : '?') + 'download=1';
                const proxyUrl = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(proxyTarget);
                response = await fetch(proxyUrl);
            }
        }
        if (!response.ok) throw new Error(`Error ${response.status}`);

        const arrayBuffer = await response.arrayBuffer();
        const view = new Uint8Array(arrayBuffer.slice(0, 4));
        if (view[0] !== 0x50 || view[1] !== 0x4B) console.warn("HTML Recibido");

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const message = loadedFromLocalWorkbook ? "Sincronizado (Excel local)" : (usedProxy ? "Sincronizado (Proxy)" : "Sincronizado");
        processWorkbook(workbook, silent ? null : message);

        const banner = document.getElementById('errorBanner');
        if (banner) banner.remove();

    } catch (error) {
        console.warn("Sync failed:", error);
        if (window.refreshIntervalId) { clearInterval(window.refreshIntervalId); window.refreshIntervalId = null; }
        const header = document.querySelector('.header');
        if (header && !document.getElementById('errorBanner') && !document.getElementById('reconnectBanner')) {
            const banner = document.createElement('div');
            banner.id = 'errorBanner';
            banner.style.cssText = 'background:#fef2f2; color:#991b1b; padding:1rem; border-radius:0.5rem; margin-bottom:1rem; border:1px solid #fecaca; display:flex; gap:1rem; align-items:center;';
            banner.innerHTML = `
                <i class="fa-solid fa-triangle-exclamation" style="font-size:1.5rem"></i>
                <div style="flex:1">
                     <strong>Modo Offline</strong><br>
                     <div style="font-size:0.875rem;">
                        No se pudo cargar la fuente externa configurada.<br>
                        Use <strong>"Cargar Local"</strong> para trabajar con sus datos.
                     </div>
                </div>
                <button class="btn btn-secondary" id="btnOpenFallbackFile" style="font-size:0.85rem">
                    <i class="fa-solid fa-folder-open"></i> Abrir Archivo
                </button>
            `;
            header.parentNode.insertBefore(banner, header.nextSibling);
            document.getElementById('btnOpenFallbackFile').addEventListener('click', pickLocalFile);
        }
        document.getElementById('lastUpdateText').innerHTML = `<span style="color:#ef4444">Offline</span>`;
    } finally {
        if (!silent) showLoading(false);
    }
}

async function tryLocalWorkbook() {
    if (!canTryBundledWorkbook()) return null;
    try {
        const response = await fetch(CONFIG.localFileName, { cache: 'no-store' });
        return response.ok ? response : null;
    } catch (e) {
        return null;
    }
}

function canTryBundledWorkbook() {
    return ['localhost', '127.0.0.1', ''].includes(window.location.hostname);
}

function processWorkbook(workbook, successMsg) {
    let firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);
    const freshData = processRawData(rawData);
    mergeData(freshData);
    renderDashboard();
    if (successMsg) showToast(successMsg, "success");
}

function processRawData(data) {
    if (!data || data.length === 0) return [];

    // Normalized Key Finder
    const findKey = (row, keywords) => {
        const rowKeys = Object.keys(row);
        const normalizeHeader = (value) => value
            .replace(/[\r\n]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase();
        for (let keyword of keywords) {
            const cleanKeyword = keyword.toLowerCase().trim();
            const hit = rowKeys.find(k => normalizeHeader(k).includes(cleanKeyword));
            if (hit) return hit;
        }
        return null;
    };

    return data.map((row, index) => {
        const getVal = (targetField) => {
            const keywords = FIELD_MAP[targetField];
            if (!keywords) return "";
            const trueKey = findKey(row, keywords);
            return trueKey ? row[trueKey] : "";
        };

        const sourceId = getVal('id_original');
        const rawDate = getVal('fecha_creacion');
        const rawDesc = getVal('descripcion') || `Item #${index}`;
        const idBase = (rawDate || '') + (getVal('usuario_registro') || '') + rawDesc.substring(0, 20);
        const id = sourceId ? `excel_${sourceId}` : (idBase ? btoa(unescape(encodeURIComponent(idBase))) : `row_${index}`);

        let dateObj = new Date();
        if (rawDate && typeof rawDate === 'number') dateObj = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
        else if (rawDate) dateObj = new Date(rawDate);

        let tipo = normalizeTipo(getVal('tipo'));

        return {
            id: id,
            source: 'excel',
            id_original: sourceId || index + 1,
            fecha_creacion: dateObj,
            usuario_registro: getVal('usuario_registro') || 'Anónimo',
            hotel: getVal('hotel') || 'N/A',
            tipo: tipo,
            departamento: getVal('departamento') || 'General',
            descripcion: rawDesc,
            responsable: getVal('responsable') || '',
            estado: getVal('estado') || 'Pendiente',
            accion: getVal('accion') || '',
            cliente: getVal('cliente') || '',
            correo: getVal('correo') || '',
            solicita_respuesta: getVal('solicita_respuesta') || '',
            telefono: getVal('telefono') || '',
            correo_respuesta: getVal('correo_respuesta') || '',
            fecha_cierre: getVal('fecha_cierre') || '',
            notas_internas: ''
        };
    });
}

function normalizeTipo(value) {
    const text = (value || '').toString().toLowerCase();
    if (text.includes('reclam')) return 'Reclamación';
    if (text.includes('queja')) return 'Queja';
    if (text.includes('sugerencia')) return 'Sugerencia';
    if (text.includes('incidencia')) return 'Incidencia';
    return value || 'General';
}

function mergeData(freshItems) {
    const currentMap = new Map(STATE.incidencias.map(i => [i.id, i]));
    const nextItems = freshItems.map(item => {
        if (currentMap.has(item.id)) {
            const existing = currentMap.get(item.id);
            existing.fecha_creacion = item.fecha_creacion;
            existing.hotel = item.hotel;
            existing.tipo = item.tipo;
            existing.departamento = item.departamento;
            existing.descripcion = item.descripcion;
            existing.usuario_registro = item.usuario_registro;
            existing.accion = item.accion;
            existing.cliente = item.cliente;
            existing.correo = item.correo;
            existing.solicita_respuesta = item.solicita_respuesta;
            existing.telefono = item.telefono;
            existing.correo_respuesta = item.correo_respuesta;
            existing.fecha_cierre = item.fecha_cierre;
            existing.id_original = item.id_original;
            return existing;
        }
        return item;
    });
    const localItems = STATE.incidencias.filter(item => item.source === 'local' || item.id?.startsWith('local_'));
    STATE.incidencias = [...nextItems, ...localItems];
    STATE.lastUpdate = new Date();
    saveState();
}

function saveState() {
    localStorage.setItem(CONFIG.dbName, JSON.stringify({ items: STATE.incidencias, lastUpdate: STATE.lastUpdate }));
    updateLastUpdateUI();
}

function renderDashboard() {
    renderKPIs(); renderTable(); renderCharts(); populateSelects();
}

function renderKPIs() {
    const items = STATE.incidencias;
    document.getElementById('kpiTotal').innerText = items.length;
    document.getElementById('kpiAbiertas').innerText = items.filter(i => ['Pendiente', 'En proceso'].includes(i.estado)).length;
    document.getElementById('kpiCerradas').innerText = items.filter(i => ['Resuelto', 'Cerrado'].includes(i.estado)).length;
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    const { search, hotel, estado } = STATE.filters;
    const filtered = STATE.incidencias.filter(item => {
        const matchesSearch = (item.descripcion || '').toLowerCase().includes(search) || (item.usuario_registro || '').toLowerCase().includes(search);
        const matchesHotel = hotel === 'all' || item.hotel === hotel;
        const matchesEstado = estado === 'all' || item.estado === estado;
        return matchesSearch && matchesHotel && matchesEstado;
    });
    filtered.sort((a, b) => new Date(b.fecha_creacion) - new Date(a.fecha_creacion));

    filtered.forEach(item => {
        const tr = document.createElement('tr');
        tr.addEventListener('click', (event) => {
            if (event.target.closest('select, input, button, a')) return;
            openIncidentModal(item.id);
        });
        tr.innerHTML = `
            <td>${formatDate(item.fecha_creacion)}</td>
            <td><span class="font-medium">${escapeHtml(item.hotel)}</span></td>
            <td>${escapeHtml(item.tipo)}</td>
            <td>${escapeHtml(item.departamento)}</td>
            <td style="max-width: 300px;"><div class="truncate" title="${escapeHtml(item.descripcion)}">${escapeHtml(item.descripcion)}</div></td>
            <td><select onchange="updateStatus('${item.id}', this.value)" class="badge ${getBadgeClass(item.estado)}">
                <option value="Pendiente" ${item.estado === 'Pendiente' ? 'selected' : ''}>Pendiente</option>
                <option value="En proceso" ${item.estado === 'En proceso' ? 'selected' : ''}>En proceso</option>
                <option value="Resuelto" ${item.estado === 'Resuelto' ? 'selected' : ''}>Resuelto</option>
                <option value="Cerrado" ${item.estado === 'Cerrado' ? 'selected' : ''}>Cerrado</option>
            </select></td>
            <td><input type="text" value="${escapeHtml(item.responsable)}" onchange="updateResponsable('${item.id}', this.value)" placeholder="..." style="border:none;width:100px;"></td>
        `;
        tbody.appendChild(tr);
    });
}

function escapeHtml(value) {
    return (value ?? '').toString().replace(/[&<>"']/g, char => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    }[char]));
}

function detailField(label, value, full = false) {
    const cleanValue = value || '-';
    return `
        <div class="detail-field ${full ? 'full' : ''}">
            <span class="detail-label">${escapeHtml(label)}</span>
            <div class="detail-value">${escapeHtml(cleanValue)}</div>
        </div>
    `;
}

function openIncidentModal(id) {
    const item = STATE.incidencias.find(x => x.id === id);
    if (!item) return;

    document.getElementById('modalMeta').textContent = `ID ${item.id_original || '-'} · ${formatDate(item.fecha_creacion)} · ${item.hotel || 'Sin hotel'}`;
    document.getElementById('modalTitle').textContent = `${item.tipo || 'Incidencia'} - ${item.departamento || 'General'}`;
    document.getElementById('modalBody').innerHTML = `
        <div class="detail-grid">
            ${detailField('Hotel', item.hotel)}
            ${detailField('Fecha de registro', formatDateTime(item.fecha_creacion))}
            ${detailField('Tipo', item.tipo)}
            ${detailField('Zona o servicio', item.departamento)}
            ${detailField('Cliente', item.cliente)}
            ${detailField('Registrado por', item.usuario_registro)}
            ${detailField('Solicita respuesta', item.solicita_respuesta)}
            ${detailField('Teléfono / correo de respuesta', [item.telefono, item.correo_respuesta].filter(Boolean).join(' · '))}
            ${detailField('Correo de registro', item.correo)}
            ${detailField('Descripción completa', item.descripcion, true)}
            ${detailField('Solución inmediata', item.accion, true)}
            ${detailField('Fecha de finalización', formatDateTime(item.fecha_cierre), true)}
        </div>
    `;

    const modal = document.getElementById('incidentModal');
    modal.classList.add('active');
    modal.setAttribute('aria-hidden', 'false');
}

function closeIncidentModal() {
    const modal = document.getElementById('incidentModal');
    modal.classList.remove('active');
    modal.setAttribute('aria-hidden', 'true');
}

function openIncidentFormModal() {
    const form = document.getElementById('incidentForm');
    form.reset();
    document.getElementById('formHotel').value = 'Secotel Guadiana';
    document.getElementById('formTipo').value = 'Queja';
    const modal = document.getElementById('incidentFormModal');
    modal.classList.add('active');
    modal.setAttribute('aria-hidden', 'false');
    setTimeout(() => document.getElementById('formDepartamento').focus(), 0);
}

function closeIncidentFormModal() {
    const modal = document.getElementById('incidentFormModal');
    modal.classList.remove('active');
    modal.setAttribute('aria-hidden', 'true');
}

function handleIncidentFormSubmit(event) {
    event.preventDefault();
    const now = new Date();
    const nextLocalNumber = getNextLocalNumber();
    const item = {
        id: `local_${Date.now()}`,
        source: 'local',
        id_original: `WEB-${String(nextLocalNumber).padStart(4, '0')}`,
        fecha_creacion: now,
        usuario_registro: document.getElementById('formUsuario').value.trim() || 'Registro web',
        hotel: document.getElementById('formHotel').value,
        tipo: document.getElementById('formTipo').value,
        departamento: document.getElementById('formDepartamento').value.trim(),
        descripcion: document.getElementById('formDescripcion').value.trim(),
        responsable: '',
        estado: 'Pendiente',
        accion: document.getElementById('formAccion').value.trim(),
        cliente: document.getElementById('formCliente').value.trim(),
        correo: '',
        solicita_respuesta: document.getElementById('formSolicitaRespuesta').value,
        telefono: document.getElementById('formTelefono').value.trim(),
        correo_respuesta: document.getElementById('formCorreoRespuesta').value.trim(),
        fecha_cierre: '',
        notas_internas: ''
    };

    STATE.incidencias.unshift(item);
    STATE.lastUpdate = now;
    saveState();
    renderDashboard();
    closeIncidentFormModal();
    showToast('Incidencia guardada', 'success');
}

function getNextLocalNumber() {
    const localNumbers = STATE.incidencias
        .map(item => (item.id_original || '').toString().match(/^WEB-(\d+)$/))
        .filter(Boolean)
        .map(match => Number(match[1]));
    return localNumbers.length ? Math.max(...localNumbers) + 1 : 1;
}

function renderCharts() {
    if (!STATE.incidencias.length) return;
    const ctx1El = document.getElementById('chartIncidenciasHotel');
    const ctx2El = document.getElementById('chartTendencias');
    if (!ctx1El || !ctx2El) return;

    const byHotel = {}; STATE.incidencias.forEach(i => { const h = i.hotel || 'Otros'; byHotel[h] = (byHotel[h] || 0) + 1; });

    if (chartByHotel) chartByHotel.destroy();
    chartByHotel = new Chart(ctx1El.getContext('2d'), {
        type: 'doughnut',
        data: { labels: Object.keys(byHotel), datasets: [{ data: Object.values(byHotel), backgroundColor: ['#4f46e5', '#10b981', '#f59e0b', '#ef4444', '#64748b'] }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
    });

    const byMonth = {}; STATE.incidencias.forEach(i => {
        const d = new Date(i.fecha_creacion);
        if (!isNaN(d)) { const k = `${d.getMonth() + 1}/${d.getFullYear()}`; byMonth[k] = (byMonth[k] || 0) + 1; }
    });

    if (chartTrend) chartTrend.destroy();
    chartTrend = new Chart(ctx2El.getContext('2d'), {
        type: 'line',
        data: { labels: Object.keys(byMonth), datasets: [{ label: 'Incidencias', data: Object.values(byMonth), borderColor: '#4f46e5', tension: 0.4, fill: true }] },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

window.updateStatus = function (id, v) { const i = STATE.incidencias.find(x => x.id === id); if (i) { i.estado = v; saveState(); renderDashboard(); } };
window.updateResponsable = function (id, v) { const i = STATE.incidencias.find(x => x.id === id); if (i) { i.responsable = v; saveState(); } };

function getBadgeClass(s) { return ({ 'Pendiente': 'badge-pendiente', 'En proceso': 'badge-proceso', 'Resuelto': 'badge-resuelto', 'Cerrado': 'badge-cerrado' }[s] || 'badge-cerrado'); }
function formatDate(d) { return (!d || isNaN(new Date(d))) ? '-' : new Date(d).toLocaleDateString('es-ES'); }
function formatDateTime(d) { return (!d || isNaN(new Date(d))) ? '-' : new Date(d).toLocaleString('es-ES'); }
function populateSelects() {
    const sel = document.getElementById('filterHotel');
    const cur = sel.value;
    const h = [...new Set(STATE.incidencias.map(i => i.hotel).filter(Boolean))];
    sel.innerHTML = '<option value="all">Todos</option>' + h.map(x => `<option value="${x}">${x}</option>`).join('');
    sel.value = cur;
}
function updateLastUpdateUI() {
    if (STATE.lastUpdate) document.getElementById('lastUpdateText').innerText = "Act: " + new Date(STATE.lastUpdate).toLocaleTimeString();
}
function showLoading(s) {
    const el = document.getElementById('loadingOverlay'); if (el) el.classList.toggle('active', s);
}
function showToast(m, t) { console.log(m); }
function exportToExcel() {
    const ws = XLSX.utils.json_to_sheet(STATE.incidencias);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Incidencias");
    XLSX.writeFile(wb, "Reporte_Incidencias.xlsx");
}
