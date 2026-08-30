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
    filters: { search: '', hotel: 'all', tipo: 'all', estado: 'Pendiente', fechaDesde: '', fechaHasta: '', sortFecha: 'desc' },
    adminFilters: { search: '', hotel: 'all', tipo: 'all', estado: 'all', fechaDesde: '', fechaHasta: '', sortFecha: 'desc' },
    isAdminUnlocked: false,
    pendingDeleteIncidentId: null
};

window.refreshIntervalId = null;

// Charts Variables (Global)
let chartByHotel = null;
let chartTrend = null;
let currentModalIncidentId = null;

function matchesDateRange(itemDate, desde, hasta) {
    if (!itemDate) return !desde && !hasta;
    const d = new Date(itemDate);
    if (isNaN(d.getTime())) return true;
    if (desde) {
        const dFrom = new Date(desde + 'T00:00:00');
        if (d < dFrom) return false;
    }
    if (hasta) {
        const dTo = new Date(hasta + 'T23:59:59');
        if (d > dTo) return false;
    }
    return true;
}

// IndexedDB Helper for Persistent File Handle
const DB_NAME = 'IncidenciasFilesDB';
const STORE_NAME = 'handles';
const ATTACHMENT_DB_NAME = 'IncidenciasAttachmentsDB';
const ATTACHMENT_STORE = 'attachments';
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

function openAttachmentDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(ATTACHMENT_DB_NAME, 1);
        request.onupgradeneeded = event => {
            const db = event.target.result;
            if (!db.objectStoreNames.contains(ATTACHMENT_STORE)) {
                db.createObjectStore(ATTACHMENT_STORE, { keyPath: 'id' });
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function getAttachment(id) {
    if (!id) return null;
    try {
        const db = await openAttachmentDB();
        const record = await new Promise((resolve) => {
            const tx = db.transaction(ATTACHMENT_STORE, 'readonly');
            const req = tx.objectStore(ATTACHMENT_STORE).get(id);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => resolve(null);
        });
        db.close();
        return record;
    } catch (error) {
        console.warn('Attachment read error', error);
        return null;
    }
}

async function deleteAttachment(id) {
    if (!id) return;
    try {
        const db = await openAttachmentDB();
        await new Promise((resolve) => {
            const tx = db.transaction(ATTACHMENT_STORE, 'readwrite');
            tx.objectStore(ATTACHMENT_STORE).delete(id);
            tx.oncomplete = () => resolve();
            tx.onerror = () => resolve();
        });
        db.close();
    } catch (error) {
        console.warn('Attachment delete error', error);
    }
}

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
    descripcion: ['¿Qué incidencia', 'incidencia, queja o sugerencia', 'relevante', 'Cuéntanos', 'Descripción', 'Problema', 'Detalle', 'Motivo', 'Comentarios', 'Observaciones', 'Asunto'],
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
    renderDashboard();

    // Firebase handles real-time updates via onSnapshot automatically.
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

    // FIREBASE SYNC: Listen for real-time updates from Firestore
    if (typeof db !== 'undefined') {
        db.collection('incidencias').onSnapshot((snapshot) => {
            const firebaseItems = [];
            snapshot.forEach((doc) => {
                const data = doc.data();
                if (data.fecha_creacion && data.fecha_creacion.toDate) {
                    data.fecha_creacion = data.fecha_creacion.toDate().toISOString();
                }
                if (data.fecha_cierre && data.fecha_cierre.toDate) {
                    data.fecha_cierre = data.fecha_cierre.toDate().toISOString();
                }
                firebaseItems.push({ ...data, id: doc.id });
            });
            
            // Si Firebase está vacío pero tenemos datos locales (Migración inicial)
            if (snapshot.docs.length === 0 && STATE.incidencias.length > 0 && !localStorage.getItem('firestore_migrated')) {
                console.log("Migrando incidencias locales a Firebase...");
                localStorage.setItem('firestore_migrated', 'true');
                STATE.incidencias.forEach(item => {
                    const docId = item.id || `local_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                    item.id = docId;
                    db.collection('incidencias').doc(docId).set(item).catch(console.error);
                });
                return; // onSnapshot volverá a saltar cuando se guarden
            }

            // Firebase como fuente de verdad única para sincronizar y propagar eliminaciones
            STATE.incidencias = firebaseItems;
            STATE.lastUpdate = new Date().toISOString();
            updateLastUpdateUI();
            renderDashboard();
            if (STATE.isAdminUnlocked) {
                populateAdminSelects();
                renderAdminTable();
            }
            
            // Keep local storage in sync as a backup
            localStorage.setItem(CONFIG.dbName, JSON.stringify({ items: STATE.incidencias, lastUpdate: STATE.lastUpdate }));
        }, (error) => {
            console.error("Error fetching from Firebase:", error);
        });
    }
}

function setupEventListeners() {
    document.querySelectorAll('.nav-item[data-view]').forEach(item => {
        item.addEventListener('click', () => switchView(item.dataset.view));
    });

    document.getElementById('searchInput').addEventListener('input', (e) => {
        STATE.filters.search = e.target.value.toLowerCase(); renderTable();
    });
    document.getElementById('filterHotel').addEventListener('change', (e) => {
        STATE.filters.hotel = e.target.value; renderTable();
    });
    document.getElementById('filterTipo').addEventListener('change', (e) => {
        STATE.filters.tipo = e.target.value; renderTable();
    });
    document.getElementById('filterEstado').addEventListener('change', (e) => {
        STATE.filters.estado = e.target.value; renderTable();
    });

    // Date Filters (Operational Table)
    const fDesde = document.getElementById('filterFechaDesde');
    if (fDesde) {
        fDesde.addEventListener('change', (e) => {
            STATE.filters.fechaDesde = e.target.value;
            renderTable();
        });
    }
    const fHasta = document.getElementById('filterFechaHasta');
    if (fHasta) {
        fHasta.addEventListener('change', (e) => {
            STATE.filters.fechaHasta = e.target.value;
            renderTable();
        });
    }
    const btnClearDate = document.getElementById('btnClearDateFilter');
    if (btnClearDate) {
        btnClearDate.addEventListener('click', () => {
            if (fDesde) fDesde.value = '';
            if (fHasta) fHasta.value = '';
            STATE.filters.fechaDesde = '';
            STATE.filters.fechaHasta = '';
            renderTable();
        });
    }
    const thSort = document.getElementById('thSortFecha');
    if (thSort) {
        thSort.addEventListener('click', () => {
            STATE.filters.sortFecha = STATE.filters.sortFecha === 'desc' ? 'asc' : 'desc';
            renderTable();
        });
    }

    // Admin Filters
    const searchAdmin = document.getElementById('searchAdminInput');
    if (searchAdmin) {
        searchAdmin.addEventListener('input', (e) => {
            STATE.adminFilters.search = e.target.value.toLowerCase();
            renderAdminTable();
        });
    }
    const filterAdminH = document.getElementById('filterAdminHotel');
    if (filterAdminH) {
        filterAdminH.addEventListener('change', (e) => {
            STATE.adminFilters.hotel = e.target.value;
            renderAdminTable();
        });
    }
    const filterAdminT = document.getElementById('filterAdminTipo');
    if (filterAdminT) {
        filterAdminT.addEventListener('change', (e) => {
            STATE.adminFilters.tipo = e.target.value;
            renderAdminTable();
        });
    }
    const filterAdminE = document.getElementById('filterAdminEstado');
    if (filterAdminE) {
        filterAdminE.addEventListener('change', (e) => {
            STATE.adminFilters.estado = e.target.value;
            renderAdminTable();
        });
    }
    const fAdminDesde = document.getElementById('filterAdminFechaDesde');
    if (fAdminDesde) {
        fAdminDesde.addEventListener('change', (e) => {
            STATE.adminFilters.fechaDesde = e.target.value;
            renderAdminTable();
        });
    }
    const fAdminHasta = document.getElementById('filterAdminFechaHasta');
    if (fAdminHasta) {
        fAdminHasta.addEventListener('change', (e) => {
            STATE.adminFilters.fechaHasta = e.target.value;
            renderAdminTable();
        });
    }
    const btnClearAdminDate = document.getElementById('btnClearAdminDateFilter');
    if (btnClearAdminDate) {
        btnClearAdminDate.addEventListener('click', () => {
            if (fAdminDesde) fAdminDesde.value = '';
            if (fAdminHasta) fAdminHasta.value = '';
            STATE.adminFilters.fechaDesde = '';
            STATE.adminFilters.fechaHasta = '';
            renderAdminTable();
        });
    }
    const thAdminSort = document.getElementById('thAdminSortFecha');
    if (thAdminSort) {
        thAdminSort.addEventListener('click', () => {
            STATE.adminFilters.sortFecha = STATE.adminFilters.sortFecha === 'desc' ? 'asc' : 'desc';
            renderAdminTable();
        });
    }

    // Delete Modal Listeners
    const btnCloseDel = document.getElementById('btnCloseDeleteModal');
    if (btnCloseDel) btnCloseDel.addEventListener('click', closeDeleteModal);
    const btnCancelDel = document.getElementById('btnCancelDelete');
    if (btnCancelDel) btnCancelDel.addEventListener('click', closeDeleteModal);
    const btnConfirmDel = document.getElementById('btnConfirmDelete');
    if (btnConfirmDel) btnConfirmDel.addEventListener('click', ejecutarEliminacionIncidencia);

    const delModal = document.getElementById('deleteConfirmModal');
    if (delModal) {
        delModal.addEventListener('click', (event) => {
            if (event.target.id === 'deleteConfirmModal') closeDeleteModal();
        });
    }

    // Smart Refresh
    document.getElementById('btnRefresh').addEventListener('click', () => {
        showLoading(true);
        setTimeout(() => {
            showLoading(false);
            showToast("Datos actualizados correctamente", "success");
        }, 500);
    });

    document.getElementById('btnReport').addEventListener('click', generateManagementReport);
    document.getElementById('btnExport').addEventListener('click', exportToExcel);

    // Dark Mode Toggle
    const btnTheme = document.getElementById('btnToggleTheme');
    if (btnTheme) {
        if (localStorage.getItem('dark-mode') === 'true') {
            document.body.classList.add('dark-mode');
            btnTheme.querySelector('i').classList.replace('fa-moon', 'fa-sun');
            btnTheme.querySelector('span').innerText = 'Modo claro';
        }
        btnTheme.addEventListener('click', () => {
            const isDark = document.body.classList.toggle('dark-mode');
            localStorage.setItem('dark-mode', isDark);
            const icon = btnTheme.querySelector('i');
            const span = btnTheme.querySelector('span');
            if (isDark) {
                icon.classList.replace('fa-moon', 'fa-sun');
                span.innerText = 'Modo claro';
            } else {
                icon.classList.replace('fa-sun', 'fa-moon');
                span.innerText = 'Modo oscuro';
            }
        });
    }

    const configButton = document.getElementById('btnConfig');
    if (configButton) configButton.addEventListener('click', configureSourceUrl);
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
        if (event.key === 'Escape') {
            closeIncidentModal();
            closeDeleteModal();
            closeIncidentFormModal();
        }
    });

    // Check existing admin session
    checkAdminSession();
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
    banner.style.cssText = 'background:#f7fbf8; color:#24523f; padding:1rem; border-radius:0.75rem; margin-bottom:1rem; border:1px solid #d7e7dd; display:flex; gap:1rem; align-items:center;';
    banner.innerHTML = `
        <i class="fa-solid fa-circle-check" style="font-size:1.5rem"></i>
        <div style="flex:1">
            <strong>Q-Centros funciona de forma autónoma</strong><br>
            <div style="font-size:0.875rem;">
                Registre nuevas incidencias desde la web o importe un Excel solo cuando necesite incorporar históricos.
            </div>
        </div>
        <button class="btn btn-primary" id="btnCreateFromBanner" style="font-size:0.85rem">
            <i class="fa-solid fa-plus"></i> Nueva incidencia
        </button>
    `;
    header.parentNode.insertBefore(banner, header.nextSibling);
    document.getElementById('btnCreateFromBanner').addEventListener('click', () => {
        window.location.href = 'RegistroIncidencia.html';
    });
    document.getElementById('lastUpdateText').innerText = 'Sistema autónomo';
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

    // Sync all loaded/merged items to Firebase so they are visible across all browsers
    if (typeof db !== 'undefined') {
        STATE.incidencias.forEach(item => {
            db.collection('incidencias').doc(item.id).set(item, { merge: true }).catch(console.error);
        });
    }
}

function saveState() {
    localStorage.setItem(CONFIG.dbName, JSON.stringify({ items: STATE.incidencias, lastUpdate: STATE.lastUpdate }));
    updateLastUpdateUI();
}

function renderDashboard() {
    renderKPIs();
    renderOperations();
    renderTable();
    if (typeof renderKanban === 'function') renderKanban();
    renderCharts();
    renderAnalysis();
    populateSelects();
}

function renderKPIs() {
    const items = STATE.incidencias;
    document.getElementById('kpiTotal').innerText = items.length;
    document.getElementById('kpiAbiertas').innerText = items.filter(i => ['Pendiente', 'En proceso'].includes(i.estado)).length;
    document.getElementById('kpiCerradas').innerText = items.filter(i => ['Resuelto', 'Cerrado'].includes(i.estado)).length;
    document.getElementById('kpiTiempo').innerText = calculateAverageCloseDays(items);
}

function renderOperations() {
    const active = STATE.incidencias.filter(isOpenIncident);
    
    const threeDaysAgo = new Date();
    threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
    const oldPending = active.filter(x => new Date(x.fecha_creacion) < threeDaysAgo).length;
    const oldElement = document.getElementById('opsOldPending');
    if(oldElement) oldElement.innerText = oldPending;

    const noOwner = active.filter(x => !x.responsable || x.responsable.trim() === '').length;
    const noOwnerElement = document.getElementById('opsNoOwner');
    if(noOwnerElement) noOwnerElement.innerText = noOwner;

    const ownerCounts = {};
    active.forEach(item => {
        const owner = item.responsable ? item.responsable.trim() : 'Sin asignar';
        ownerCounts[owner] = (ownerCounts[owner] || 0) + 1;
    });
    
    const ownerDetails = Object.entries(ownerCounts)
        .sort((a, b) => b[1] - a[1])
        .map(([owner, count]) => `<div style="display:flex; justify-content:space-between; padding:0.2rem 0; border-bottom:1px solid var(--border);"><span>${escapeHtml(owner)}</span><strong>${count}</strong></div>`)
        .join('');
        
    const detailElement = document.getElementById('opsTopOwnerDetail');
    if(detailElement) detailElement.innerHTML = ownerDetails || '<span>No hay carga pendiente.</span>';
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    const { search, hotel, tipo, estado, fechaDesde, fechaHasta, sortFecha } = STATE.filters;
    const filtered = STATE.incidencias.filter(item => {
        const matchesSearch = (item.descripcion || '').toLowerCase().includes(search) || (item.usuario_registro || '').toLowerCase().includes(search);
        const matchesHotel = hotel === 'all' || item.hotel === hotel;
        const matchesTipo = tipo === 'all' || item.tipo === tipo;
        const matchesEstado = estado === 'all' || item.estado === estado;
        const matchesDate = matchesDateRange(item.fecha_creacion, fechaDesde, fechaHasta);
        return matchesSearch && matchesHotel && matchesTipo && matchesEstado && matchesDate;
    });

    filtered.sort((a, b) => {
        const da = new Date(a.fecha_creacion).getTime() || 0;
        const db = new Date(b.fecha_creacion).getTime() || 0;
        return sortFecha === 'asc' ? da - db : db - da;
    });

    const thSort = document.getElementById('thSortFecha');
    if (thSort) {
        thSort.innerHTML = `Fecha <i class="fa-solid fa-arrow-${sortFecha === 'asc' ? 'up-short-wide' : 'down-wide-short'}" style="margin-left: 4px; color: var(--primary);"></i>`;
    }

    if (!filtered.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="empty-table">No hay registros con los filtros y fechas seleccionados.</td>
            </tr>
        `;
        return;
    }

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
            <td style="max-width: 300px;">
                <div class="truncate" title="${escapeHtml(item.descripcion)}">
                    <span style="font-size: 0.75rem; color: var(--text-muted); font-weight: normal; margin-right: 5px;">#${item.id_original || 'S/N'}</span>
                    ${escapeHtml(item.descripcion || '')}
                </div>
            </td>
            <td><span class="badge ${getBadgeClass(item.estado)}">${escapeHtml(item.estado || 'Pendiente')}</span></td>
            <td><span style="color: var(--text-muted);">${escapeHtml(item.responsable || 'Sin asignar')}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderAnalysis() {
    renderMetricList('analysisByType', countBy(STATE.incidencias, item => item.tipo || 'Sin tipo'));
    renderMetricList('analysisByStatus', countBy(STATE.incidencias, item => item.estado || 'Pendiente'));
    renderMetricList('analysisByArea', countBy(STATE.incidencias, item => item.departamento || 'General'), 6);
    renderRecentList();
}

function renderMetricList(containerId, metrics, limit = 4) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const total = STATE.incidencias.length || 1;
    const rows = Object.entries(metrics)
        .sort((a, b) => b[1] - a[1])
        .slice(0, limit);

    if (!rows.length) {
        container.innerHTML = '<p class="empty-analysis">Sin datos registrados.</p>';
        return;
    }

    container.innerHTML = rows.map(([label, count]) => {
        const percent = Math.round((count / total) * 100);
        return `
            <div class="metric-row">
                <div>
                    <strong>${escapeHtml(label)}</strong>
                    <span>${count} registros</span>
                </div>
                <div class="metric-bar" aria-hidden="true">
                    <span style="width:${percent}%"></span>
                </div>
                <b>${percent}%</b>
            </div>
        `;
    }).join('');
}

function renderRecentList() {
    const container = document.getElementById('analysisRecent');
    if (!container) return;
    const recent = [...STATE.incidencias]
        .sort((a, b) => new Date(b.fecha_creacion) - new Date(a.fecha_creacion))
        .slice(0, 5);

    if (!recent.length) {
        container.innerHTML = '<p class="empty-analysis">Sin registros recientes.</p>';
        return;
    }

    container.innerHTML = recent.map(item => `
        <button class="recent-item" type="button" onclick="openIncidentModal('${item.id}')">
            <span>${escapeHtml(item.tipo || 'Registro')} · ${escapeHtml(item.hotel || 'Sin centro')}</span>
            <strong>${escapeHtml(item.departamento || 'General')}</strong>
            <small>${formatDate(item.fecha_creacion)} · ${escapeHtml(item.estado || 'Pendiente')}</small>
        </button>
    `).join('');
}

function countBy(items, getLabel) {
    return items.reduce((acc, item) => {
        const label = getLabel(item);
        acc[label] = (acc[label] || 0) + 1;
        return acc;
    }, {});
}

function getTopMetric(items, getLabel) {
    const [label = '-', count = 0] = Object.entries(countBy(items, getLabel))
        .sort((a, b) => b[1] - a[1])[0] || [];
    return { label, count };
}

function isOpenIncident(item) {
    return ['Pendiente', 'En proceso'].includes(item.estado || 'Pendiente');
}

function calculateAverageCloseDays(items) {
    const closed = items
        .map(item => {
            const start = new Date(item.fecha_creacion);
            const end = new Date(item.fecha_cierre);
            if (isNaN(start) || isNaN(end) || end < start) return null;
            return Math.max(1, Math.round((end - start) / 86400000));
        })
        .filter(value => value !== null);

    if (!closed.length) return '-';
    const average = closed.reduce((sum, value) => sum + value, 0) / closed.length;
    return `${average.toFixed(average >= 10 ? 0 : 1)} d`;
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

async function openIncidentModal(id) {
    const item = STATE.incidencias.find(x => x.id === id);
    if (!item) return;
    currentModalIncidentId = id;

    document.getElementById('modalMeta').textContent = `ID ${item.id_original || '-'} · ${formatDate(item.fecha_creacion)} · ${item.hotel || 'Sin hotel'}`;
    document.getElementById('modalTitle').textContent = `${item.tipo || 'Incidencia'} - ${item.departamento || 'General'}`;
    document.getElementById('modalBody').innerHTML = `
        <div class="modal-split">
            <!-- Columna Izquierda: Detalles (Solo lectura) -->
            <div class="modal-left">
                <section class="case-summary">
                    <div>
                        <span>Centro</span>
                        <strong>${escapeHtml(item.hotel || '-')}</strong>
                    </div>
                    <div>
                        <span>Tipo</span>
                        <strong>${escapeHtml(item.tipo || '-')}</strong>
                    </div>
                    <div>
                        <span>Estado</span>
                        <strong>${escapeHtml(item.estado || 'Pendiente')}</strong>
                    </div>
                    <div>
                        <span>Responsable</span>
                        <strong>${escapeHtml(item.responsable || 'Sin asignar')}</strong>
                    </div>
                </section>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 1rem; margin-bottom: 1.5rem; padding: 1rem; background: #f8fafc; border-radius: 0.5rem; border: 1px solid var(--border);">
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">FECHA REGISTRO</span>${formatDateTime(item.fecha_creacion) || '-'}</div>
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">ZONA / SERVICIO</span>${escapeHtml(item.departamento || '-')}</div>
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">CLIENTE</span>${escapeHtml(item.cliente || '-')}</div>
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">REGISTRADO POR</span>${escapeHtml(item.usuario_registro || '-')}</div>
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">SOLICITA RESP.</span>${escapeHtml(item.solicita_respuesta || '-')}</div>
                    <div style="font-size: 0.85rem;"><span style="display:block; color:var(--text-muted); font-size:0.7rem; font-weight:bold; margin-bottom: 2px;">CONTACTO</span>${escapeHtml([item.telefono, item.correo_respuesta].filter(Boolean).join(' · ') || '-')}</div>
                </div>

                <div class="detail-grid">
                    ${detailField('Descripción completa', item.descripcion, true)}
                    <div class="detail-field full">
                        <span class="detail-label">Adjuntos</span>
                        <div id="attachmentViewer" class="attachment-viewer"></div>
                    </div>
                </div>
            </div>

            <!-- Columna Derecha: Panel de Gestión -->
            <div class="modal-right">
                <form id="receptionManagementForm" class="reception-management" style="margin:0; padding:1.5rem; background:#f8fafc; border-radius: 0.5rem; border: 1px solid var(--border); box-shadow:none; height: 100%;">
                    <span class="section-kicker" style="font-size: 1rem;">Panel de Gestión</span>
                    <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1.5rem;">Actualice el progreso, asigne responsable y documente las acciones.</p>
                    
                    <div class="form-grid" style="display:flex; flex-direction:column; gap:1.25rem;">
                        <label class="full" style="background: white; padding: 0.75rem; border-radius: 0.5rem; border: 1px solid #e2e8f0;">
                            <span style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; font-weight: bold; margin-bottom: 0.5rem; display: block;">Estado actual</span>
                            <select id="manageEstado" style="font-weight: bold; font-size: 1.1rem; color: var(--primary); border: none; background: transparent; padding: 0; width: 100%; cursor: pointer;">
                                <option value="Pendiente" ${item.estado === 'Pendiente' ? 'selected' : ''}>Pendiente</option>
                                <option value="En proceso" ${item.estado === 'En proceso' ? 'selected' : ''}>En proceso</option>
                                <option value="Resuelto" ${item.estado === 'Resuelto' ? 'selected' : ''}>Resuelto</option>
                                <option value="Cerrado" ${item.estado === 'Cerrado' ? 'selected' : ''}>Cerrado</option>
                                <option value="Irresoluble" ${item.estado === 'Irresoluble' ? 'selected' : ''}>Irresoluble (Absurda/Imposible)</option>
                            </select>
                        </label>
                        <label class="full">
                            <span style="font-size: 0.75rem; font-weight: bold; color: var(--text-muted); text-transform: uppercase;">Asignado a (Responsable)</span>
                            <input id="manageResponsable" type="text" value="${escapeHtml(item.responsable || '')}" placeholder="Mantenimiento, Recepción..." style="font-size: 1rem; padding: 0.75rem; background: white; border: 1px solid #e2e8f0; border-radius: 0.5rem;">
                        </label>
                        <label class="full">
                            <span style="font-size: 0.75rem; font-weight: bold; color: var(--text-muted); text-transform: uppercase;">Acción inmediata / 1ª Respuesta</span>
                            <textarea id="manageAccion" rows="2" placeholder="Qué se hizo al momento de la incidencia..." style="font-size: 1rem; padding: 0.75rem; background: white; border: 1px solid #e2e8f0; border-radius: 0.5rem;">${escapeHtml(item.accion || '')}</textarea>
                        </label>
                        <label class="full">
                            <span style="font-size: 0.75rem; font-weight: bold; color: var(--text-muted); text-transform: uppercase;">Gestión de Dirección / Seguimiento</span>
                            <textarea id="manageGestionDireccion" rows="3" placeholder="Pasos dados para resolver el problema a fondo..." style="font-size: 1rem; padding: 0.75rem; background: white; border: 1px solid #e2e8f0; border-radius: 0.5rem;">${escapeHtml(item.gestion_direccion || item.notas_internas || '')}</textarea>
                        </label>
                        <label class="full">
                            <span style="font-size: 0.75rem; font-weight: bold; color: var(--text-muted); text-transform: uppercase;">Resolución Final</span>
                            <textarea id="manageResolucion" rows="2" placeholder="Cómo ha quedado resuelto el caso..." style="font-size: 1rem; padding: 0.75rem; background: white; border: 1px solid #e2e8f0; border-radius: 0.5rem;">${escapeHtml(item.resolucion || '')}</textarea>
                        </label>
                        <label class="full">
                            <span style="font-size: 0.75rem; font-weight: bold; color: var(--text-muted); text-transform: uppercase;">Fecha de cierre</span>
                            <input id="manageFechaCierre" type="datetime-local" value="${formatDateTimeInput(item.fecha_cierre)}" style="font-size: 1rem; padding: 0.75rem; background: white; border: 1px solid #e2e8f0; border-radius: 0.5rem;">
                        </label>
                    </div>
                    </div>
                </form>
            </div>
        </div>
    `;

    const modal = document.getElementById('incidentModal');
    modal.classList.add('active');
    modal.setAttribute('aria-hidden', 'false');
    document.getElementById('receptionManagementForm').addEventListener('submit', saveReceptionManagement);
    await renderIncidentAttachments(item);
}

function closeIncidentModal() {
    const modal = document.getElementById('incidentModal');
    modal.classList.remove('active');
    modal.setAttribute('aria-hidden', 'true');
    currentModalIncidentId = null;
    document.querySelectorAll('[data-object-url]').forEach(element => {
        URL.revokeObjectURL(element.dataset.objectUrl);
    });
}

function saveReceptionManagement(event) {
    event.preventDefault();
    const item = STATE.incidencias.find(x => x.id === currentModalIncidentId);
    if (!item) return;

    item.estado = document.getElementById('manageEstado').value;
    item.responsable = document.getElementById('manageResponsable').value.trim();
    item.accion = document.getElementById('manageAccion').value.trim();
    item.gestion_direccion = document.getElementById('manageGestionDireccion').value.trim();
    item.notas_internas = item.gestion_direccion; // Retrocompatibilidad
    item.resolucion = document.getElementById('manageResolucion').value.trim();
    item.fecha_cierre = document.getElementById('manageFechaCierre').value || '';
    item.fecha_ultima_gestion = new Date().toISOString();

    STATE.lastUpdate = new Date();
    saveState();
    
    if (typeof db !== 'undefined') {
        db.collection('incidencias').doc(item.id).set(item).catch(console.error);
    }
    
    renderDashboard();
    showToast('Gestión guardada', 'success');
    closeIncidentModal();
}

function switchView(view) {
    document.querySelectorAll('.view-section').forEach(section => {
        section.classList.toggle('active', section.id === `${view}Section`);
    });
    document.querySelectorAll('.nav-item[data-view]').forEach(item => {
        item.classList.toggle('active', item.dataset.view === view);
    });
    if (view === 'admin') {
        if (STATE.isAdminUnlocked) {
            populateAdminSelects();
            renderAdminTable();
        } else {
            setTimeout(() => {
                const pinInput = document.getElementById('adminPinInput');
                if (pinInput) pinInput.focus();
            }, 100);
        }
    }
}

async function renderIncidentAttachments(item) {
    const viewer = document.getElementById('attachmentViewer');
    if (!viewer) return;

    const attachments = Array.isArray(item.attachments) ? item.attachments : [];
    if (!attachments.length) {
        viewer.innerHTML = '<span class="empty-attachments">Sin adjuntos</span>';
        return;
    }

    viewer.innerHTML = '';
    for (const meta of attachments) {
        const record = await getAttachment(meta.id);
        const card = document.createElement('div');
        card.className = 'attachment-preview';

        if (!record?.blob && !meta.url) {
            card.innerHTML = `
                <i class="fa-solid fa-file-circle-exclamation"></i>
                <div>
                    <strong>${escapeHtml(meta.name)}</strong>
                    <span>No disponible en este navegador</span>
                </div>
            `;
            viewer.appendChild(card);
            continue;
        }

        const objectUrl = record?.blob ? URL.createObjectURL(record.blob) : meta.url;
        if(record?.blob) card.dataset.objectUrl = objectUrl;
        const isImage = (meta.type || "").startsWith("image/");
        card.innerHTML = isImage ? `
            <a href="${objectUrl}" target="_blank" rel="noopener">
                <img src="${objectUrl}" alt="${escapeHtml(record.name || meta.name)}">
            </a>
            <div>
                <strong>${escapeHtml(record.name || meta.name)}</strong>
                <span>${formatBytes(record.size || meta.size)}</span>
            </div>
        ` : `
            <i class="fa-solid fa-file"></i>
            <div>
                <strong>${escapeHtml(record.name || meta.name)}</strong>
                <span>${formatBytes(record.size || meta.size)}</span>
                <a href="${objectUrl}" download="${escapeHtml(record.name || meta.name)}">Descargar archivo</a>
            </div>
        `;
        viewer.appendChild(card);
    }
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
    
    if (typeof db !== 'undefined') {
        db.collection('incidencias').doc(item.id).set(item).catch(console.error);
    }
    
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

window.updateStatus = function (id, v) { 
    const i = STATE.incidencias.find(x => x.id === id); 
    if (i) { 
        i.estado = v; 
        saveState(); 
        if (typeof db !== 'undefined') db.collection('incidencias').doc(id).update({ estado: v }).catch(console.error);
        renderDashboard(); 
    } 
};
window.updateResponsable = function (id, v) { 
    const i = STATE.incidencias.find(x => x.id === id); 
    if (i) { 
        i.responsable = v; 
        saveState(); 
        if (typeof db !== 'undefined') db.collection('incidencias').doc(id).update({ responsable: v }).catch(console.error);
    } 
};

function getBadgeClass(s) { return ({ 'Pendiente': 'badge-pendiente', 'En proceso': 'badge-proceso', 'Resuelto': 'badge-resuelto', 'Cerrado': 'badge-cerrado' }[s] || 'badge-cerrado'); }
function formatDate(d) { return (!d || isNaN(new Date(d))) ? '-' : new Date(d).toLocaleDateString('es-ES'); }
function formatDateTime(d) { return (!d || isNaN(new Date(d))) ? '-' : new Date(d).toLocaleString('es-ES'); }
function formatDateTimeInput(d) {
    if (!d || isNaN(new Date(d))) return '';
    const date = new Date(d);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    return date.toISOString().slice(0, 16);
}
function formatBytes(bytes) {
    if (!bytes) return '0 KB';
    const units = ['B', 'KB', 'MB', 'GB'];
    const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / Math.pow(1024, index);
    return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
}
function populateSelects() {
    const sel = document.getElementById('filterHotel');
    const cur = sel.value;
    const h = [...new Set(STATE.incidencias.map(i => i.hotel).filter(Boolean))];
    sel.innerHTML = '<option value="all">Todos</option>' + h.map(x => `<option value="${x}">${x}</option>`).join('');
    sel.value = cur;
}
function updateLastUpdateUI() {
    if (STATE.lastUpdate) {
        document.getElementById('lastUpdateText').innerHTML = `<i class="fa-solid fa-cloud-check" style="color:var(--success)"></i> Sincronizado: ` + new Date(STATE.lastUpdate).toLocaleString('es-ES');
    } else {
        document.getElementById('lastUpdateText').innerHTML = `<i class="fa-solid fa-cloud-arrow-up" style="color:var(--primary)"></i> Conectando con la nube...`;
    }
}
function showLoading(s) {
    const el = document.getElementById('loadingOverlay'); if (el) el.classList.toggle('active', s);
}
function showToast(m, t) { console.log(m); }

function generateManagementReport() {
    const items = [...STATE.incidencias].sort((a, b) => new Date(a.fecha_creacion) - new Date(b.fecha_creacion));
    if (!items.length) {
        showToast('No hay datos para generar informe', 'error');
        alert('No hay registros para generar el informe. Importe datos o cree una incidencia primero.');
        return;
    }

    const reportDate = new Date();
    const openItems = items.filter(isOpenIncident);
    const closedItems = items.filter(item => ['Resuelto', 'Cerrado'].includes(item.estado));
    const noAction = items.filter(item => !(item.accion || '').trim()).length;
    const noResponse = items.filter(item => !(item.solicita_respuesta || '').trim()).length;
    const byHotel = countBy(items, item => item.hotel || 'Sin centro');
    const byType = countBy(items, item => item.tipo || 'Sin tipo');
    const byArea = countBy(items, item => item.departamento || 'General');
    const topHotel = getTopMetric(items, item => item.hotel || 'Sin centro');
    const topType = getTopMetric(items, item => item.tipo || 'Sin tipo');
    const topArea = getTopMetric(items, item => item.departamento || 'General');
    const risks = getPriorityRisks(items);

    const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Informe para dirección</title>
            <style>
                @page { margin: 1.8cm; }
                body { font-family: Arial, sans-serif; color: #18231f; line-height: 1.45; }
                h1 { color: #145a3f; font-size: 26pt; margin: 0 0 6px; }
                h2 { color: #145a3f; font-size: 16pt; margin: 24px 0 8px; border-bottom: 2px solid #d8a63a; padding-bottom: 4px; }
                h3 { color: #24352e; font-size: 12pt; margin: 16px 0 6px; }
                p { font-size: 10.5pt; margin: 6px 0; }
                .subtitle { color: #66736d; font-size: 11pt; margin-bottom: 18px; }
                .executive { background: #f3f8f5; border-left: 5px solid #145a3f; padding: 12px 14px; margin: 16px 0; }
                table { width: 100%; border-collapse: collapse; margin: 10px 0 16px; font-size: 9.5pt; }
                th { background: #145a3f; color: #ffffff; text-align: left; padding: 7px; }
                td { border: 1px solid #dfe6e1; padding: 7px; vertical-align: top; }
                .muted { color: #66736d; }
                .pill { color: #991b1b; font-weight: bold; }
            </style>
        </head>
        <body>
            <h1>Informe para dirección</h1>
            <p class="subtitle">Quejas, sugerencias, reclamaciones e incidencias · Q-Centros · Generado el ${escapeHtml(formatDateTime(reportDate))}</p>

            <div class="executive">
                <h2 style="margin-top:0;">Conclusión ejecutiva</h2>
                <p>${buildExecutiveConclusion(items, openItems, topHotel, topType, topArea, noAction)}</p>
            </div>

            <h2>1. Resumen ejecutivo</h2>
            <table>
                <tr><th>Indicador</th><th>Resultado</th><th>Lectura para dirección</th></tr>
                <tr><td>Registros analizados</td><td>${items.length}</td><td>Base actual registrada en el sistema.</td></tr>
                <tr><td>Distribución por centro</td><td>${formatMetricSummary(byHotel, items.length)}</td><td>${escapeHtml(topHotel.label)} concentra el mayor volumen.</td></tr>
                <tr><td>Tipo de registro</td><td>${formatMetricSummary(byType, items.length)}</td><td>${escapeHtml(topType.label)} es el tipo más repetido.</td></tr>
                <tr><td>Registros abiertos</td><td>${openItems.length}</td><td>${openItems.length ? 'Requieren seguimiento hasta cierre.' : 'No constan registros pendientes.'}</td></tr>
                <tr><td>Solución documentada</td><td>${items.length - noAction} con acción · ${noAction} sin acción</td><td>${noAction ? 'Debe reforzarse la trazabilidad de respuesta.' : 'La gestión aparece documentada.'}</td></tr>
                <tr><td>Respuesta solicitada</td><td>${items.length - noResponse} informado · ${noResponse} en blanco</td><td>${noResponse ? 'Conviene completar este campo para medir obligaciones de contestación.' : 'Campo cumplimentado de forma homogénea.'}</td></tr>
            </table>

            <h2>2. Hallazgos principales</h2>
            ${buildHotelFindings(byHotel, items)}

            <h2>3. Distribución por áreas</h2>
            ${buildMetricTable(byArea, items.length, 'Área o servicio')}

            <h2>4. Riesgos que requieren actuación prioritaria</h2>
            ${buildRiskTable(risks)}

            <h2>5. Calidad de la gestión de incidencias</h2>
            <p><strong>Debilidad de control:</strong> ${buildControlWeakness(items, noAction, noResponse)}</p>
            <table>
                <tr><th>Problema de registro</th><th>Impacto</th><th>Medida de control</th></tr>
                <tr><td>Registros sin responsable</td><td>Dificulta saber quién debe cerrar la acción.</td><td>Asignar responsable antes de pasar el estado a “En proceso”.</td></tr>
                <tr><td>Soluciones no documentadas</td><td>No permite verificar si el cliente quedó atendido.</td><td>Registrar acción realizada, fecha prevista y comprobación final.</td></tr>
                <tr><td>Campos de respuesta incompletos</td><td>Puede perderse una contestación necesaria.</td><td>Marcar siempre si solicita respuesta y añadir teléfono o correo cuando proceda.</td></tr>
            </table>

            <h2>6. Plan de actuación recomendado</h2>
            <table>
                <tr><th>Plazo</th><th>Actuación</th><th>Responsable sugerido</th><th>Indicador de cierre</th></tr>
                <tr><td>0-7 días</td><td>Revisar todos los registros pendientes y asignar responsable.</td><td>Dirección / Jefes de área</td><td>100% de incidencias abiertas con responsable.</td></tr>
                <tr><td>0-7 días</td><td>Actuar sobre el foco más repetido: ${escapeHtml(topArea.label)}.</td><td>Responsable del servicio afectado</td><td>Acción correctiva definida y fecha de revisión.</td></tr>
                <tr><td>7-30 días</td><td>Revisar recurrencias por centro y tipo de registro.</td><td>Dirección</td><td>Comparativa mensual con reducción de repetición.</td></tr>
                <tr><td>Mensual</td><td>Emitir informe de seguimiento y validar cierres.</td><td>Dirección</td><td>Informe mensual archivado y acciones cerradas.</td></tr>
            </table>
        </body>
        </html>
    `;

    downloadReport(html, `Informe_direccion_incidencias_${formatFileDate(reportDate)}.doc`);
    showToast('Informe generado', 'success');
}

function buildExecutiveConclusion(items, openItems, topHotel, topType, topArea, noAction) {
    const openText = openItems.length
        ? `Hay ${openItems.length} registros abiertos que requieren seguimiento operativo.`
        : 'No constan registros abiertos en este momento.';
    const actionText = noAction
        ? `${noAction} registros no tienen una acción o solución documentada, por lo que debe reforzarse el cierre.`
        : 'Los registros incluyen actuación documentada.';
    return `El registro actual contiene ${items.length} entradas. ${topHotel.label} concentra el mayor volumen, el tipo más repetido es ${topType.label} y el foco principal por zona o servicio es ${topArea.label}. ${openText} ${actionText}`;
}

function buildHotelFindings(byHotel, items) {
    return Object.entries(byHotel)
        .sort((a, b) => b[1] - a[1])
        .map(([hotel, count]) => {
            const hotelItems = items.filter(item => (item.hotel || 'Sin centro') === hotel);
            const topArea = getTopMetric(hotelItems, item => item.departamento || 'General');
            const openCount = hotelItems.filter(isOpenIncident).length;
            return `
                <h3>${escapeHtml(hotel)}</h3>
                <p>Registra ${count} entradas (${formatPercent(count, items.length)}). El área más repetida es <strong>${escapeHtml(topArea.label)}</strong> con ${topArea.count} registros. ${openCount ? `Quedan ${openCount} registros abiertos que deben revisarse hasta cierre.` : 'No constan registros abiertos para este centro.'}</p>
            `;
        }).join('');
}

function buildMetricTable(metrics, total, firstColumn) {
    const rows = Object.entries(metrics)
        .sort((a, b) => b[1] - a[1])
        .map(([label, count]) => `
            <tr>
                <td>${escapeHtml(label)}</td>
                <td>${count}</td>
                <td>${formatPercent(count, total)}</td>
                <td>${count >= Math.max(2, total * 0.15) ? 'Foco relevante para seguimiento.' : 'Volumen bajo o puntual.'}</td>
            </tr>
        `).join('');

    return `
        <table>
            <tr><th>${firstColumn}</th><th>Total</th><th>Peso</th><th>Interpretación</th></tr>
            ${rows}
        </table>
    `;
}

function buildRiskTable(risks) {
    if (!risks.length) {
        return '<p class="muted">No se detectan riesgos prioritarios por palabras clave. Revise igualmente los registros abiertos.</p>';
    }

    return `
        <table>
            <tr><th>Prioridad</th><th>Riesgo identificado</th><th>Evidencia del registro</th><th>Acción inmediata</th></tr>
            ${risks.map(risk => `
                <tr>
                    <td><span class="pill">${risk.priority}</span></td>
                    <td>${escapeHtml(risk.title)}</td>
                    <td>${escapeHtml(risk.evidence)}</td>
                    <td>${escapeHtml(risk.action)}</td>
                </tr>
            `).join('')}
        </table>
    `;
}

function getPriorityRisks(items) {
    const catalog = [
        { priority: 'P1', title: 'Seguridad de clientes o trabajadores', terms: ['caida', 'resbal', 'corte', 'lesion', 'quemad', 'explosion', 'seguridad'], action: 'Revisar zona, retirar el riesgo, documentar evidencias y cerrar con responsable.' },
        { priority: 'P1', title: 'Instalaciones o mantenimiento crítico', terms: ['agua', 'averia', 'rot', 'fuga', 'goter', 'electric', 'secador'], action: 'Abrir parte técnico, identificar causa, fecha de corrección y verificación posterior.' },
        { priority: 'P2', title: 'Servicio y experiencia del cliente', terms: ['buffet', 'desayuno', 'limpieza', 'habitacion', 'ruido', 'personal'], action: 'Revisar estándar de servicio, reforzar checklist y confirmar seguimiento con el área.' },
        { priority: 'P2', title: 'Reputación o comunicación al cliente', terms: ['reclam', 'respuesta', 'internet', 'aviso', 'reserva'], action: 'Validar comunicación al cliente y registrar respuesta formal si procede.' }
    ];

    return catalog.map(rule => {
        const hits = items.filter(item => {
            const text = normalizeSearchText(`${item.tipo} ${item.departamento} ${item.descripcion} ${item.accion} ${item.notas_internas}`);
            return rule.terms.some(term => text.includes(term));
        });
        return hits.length ? {
            priority: rule.priority,
            title: rule.title,
            evidence: `${hits.length} registros relacionados. Ejemplos: ${hits.slice(0, 4).map(item => item.id_original || item.id).join(', ')}`,
            action: rule.action
        } : null;
    }).filter(Boolean);
}

function buildControlWeakness(items, noAction, noResponse) {
    const noOwner = items.filter(item => isOpenIncident(item) && !(item.responsable || '').trim()).length;
    return `${noAction} de ${items.length} registros no tienen solución documentada, ${noResponse} no informan si el cliente solicita respuesta y ${noOwner} registros abiertos no tienen responsable asignado. Estos campos son clave para demostrar seguimiento y cierre.`;
}

function formatMetricSummary(metrics, total) {
    return Object.entries(metrics)
        .sort((a, b) => b[1] - a[1])
        .map(([label, count]) => `${escapeHtml(label)} ${count} (${formatPercent(count, total)})`)
        .join(' · ');
}

function formatPercent(value, total) {
    if (!total) return '0%';
    return `${((value / total) * 100).toFixed(1).replace('.', ',')}%`;
}

function normalizeSearchText(value) {
    return (value || '').toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function formatFileDate(date) {
    return date.toISOString().slice(0, 10);
}

function downloadReport(html, filename) {
    const blob = new Blob(['\ufeff', html], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function exportToExcel() {
    const exportRows = STATE.incidencias.map(item => ({
        id: item.id_original || item.id,
        fecha_creacion: formatDateTime(item.fecha_creacion),
        hotel: item.hotel,
        tipo: item.tipo,
        departamento: item.departamento,
        descripcion: item.descripcion,
        estado: item.estado,
        responsable: item.responsable,
        cliente: item.cliente,
        solicita_respuesta: item.solicita_respuesta,
        telefono: item.telefono,
        correo_respuesta: item.correo_respuesta,
        solucion: item.accion,
        datos_especificos: item.notas_internas,
        adjuntos: Array.isArray(item.attachments) ? item.attachments.map(file => file.name).join('; ') : ''
    }));
    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Incidencias");
    XLSX.writeFile(wb, "Reporte_Incidencias.xlsx");
}

function renderKanban() {
  const board = document.querySelector('.kanban-board');
  if (!board) return;
  
  const normalizeState = (s) => {
    if (!s) return 'Pendiente';
    const l = s.toLowerCase();
    if (l.includes('pendiente')) return 'Pendiente';
    if (l.includes('proceso')) return 'En proceso';
    if (l.includes('resuelto')) return 'Resuelto';
    if (l.includes('cerrado')) return 'Cerrado';
    if (l.includes('irresoluble')) return 'Irresoluble';
    return 'Pendiente';
  };

  ['Pendiente', 'En proceso', 'Resuelto', 'Cerrado', 'Irresoluble'].forEach(status => {
    const col = board.querySelector('.kanban-column[data-status="' + status + '"] .kanban-cards');
    const countBadge = board.querySelector('.kanban-column[data-status="' + status + '"] .count');
    if (!col) return;
    col.innerHTML = '';
    const items = STATE.incidencias.filter(x => normalizeState(x.estado) === status);
    if(countBadge) countBadge.innerText = items.length;
    items.sort((a, b) => new Date(b.fecha_creacion) - new Date(a.fecha_creacion)).forEach(item => {
      const card = document.createElement('div');
      card.className = 'kanban-card';
      card.draggable = true;
      card.dataset.id = item.id;
      card.innerHTML = `<h4><span style="color:var(--text-muted); font-size:0.75rem; font-weight:normal; margin-right:5px;">#${item.id_original || 'S/N'}</span>${escapeHtml(item.tipo || 'Incidencia')}</h4><p>${escapeHtml(item.descripcion || '')}</p><div class="kanban-meta"><span>${escapeHtml(item.hotel || '-')}</span><span>${escapeHtml(item.responsable || 'Sin asignar')}</span></div>`;
      card.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', item.id); card.classList.add('dragging'); });
      card.addEventListener('dragend', () => card.classList.remove('dragging'));
      card.addEventListener('click', () => openIncidentModal(item.id));
      col.appendChild(card);
    });
  });
}
function setupKanbanEvents() {
  document.querySelectorAll('.kanban-column').forEach(col => {
    col.addEventListener('dragover', e => { e.preventDefault(); const cards = col.querySelector('.kanban-cards'); cards.style.background = 'rgba(0,0,0,0.05)'; });
    col.addEventListener('dragleave', e => { const cards = col.querySelector('.kanban-cards'); cards.style.background = ''; });
    col.addEventListener('drop', e => {
      e.preventDefault();
      const cards = col.querySelector('.kanban-cards');
      cards.style.background = '';
      const id = e.dataTransfer.getData('text/plain');
      const newStatus = col.dataset.status;
      const item = STATE.incidencias.find(x => x.id === id);
      if(item && item.estado !== newStatus) {
        item.estado = newStatus;
        if(newStatus==='Cerrado' || newStatus==='Resuelto') item.fecha_cierre = new Date().toISOString();
        saveState();
        if(typeof db !== 'undefined') { db.collection('incidencias').doc(item.id).update({estado: item.estado, fecha_cierre: item.fecha_cierre}).catch(console.error); }
        renderDashboard();
        if(typeof showToast === 'function') showToast('Estado actualizado', 'success');
      }
    });
  });
}
document.addEventListener('DOMContentLoaded', setupKanbanEvents);

// ==========================================
// MÓDULO DE ADMINISTRACIÓN Y ELIMINACIÓN
// ==========================================

function getAdminPin() {
    return localStorage.getItem('admin_pin') || '1234';
}

function checkAdminSession() {
    if (sessionStorage.getItem('admin_unlocked') === 'true') {
        STATE.isAdminUnlocked = true;
        const authCard = document.getElementById('adminAuthCard');
        const dashContent = document.getElementById('adminDashboardContent');
        if (authCard) authCard.style.display = 'none';
        if (dashContent) dashContent.style.display = 'block';
        populateAdminSelects();
        renderAdminTable();
    }
}

function handleAdminLogin(event) {
    if (event) event.preventDefault();
    const pinInput = document.getElementById('adminPinInput');
    const pinError = document.getElementById('adminPinError');
    const enteredPin = (pinInput?.value || '').trim();

    if (enteredPin === getAdminPin()) {
        STATE.isAdminUnlocked = true;
        sessionStorage.setItem('admin_unlocked', 'true');
        if (pinError) pinError.style.display = 'none';
        if (pinInput) pinInput.value = '';
        
        const authCard = document.getElementById('adminAuthCard');
        const dashContent = document.getElementById('adminDashboardContent');
        if (authCard) authCard.style.display = 'none';
        if (dashContent) dashContent.style.display = 'block';

        populateAdminSelects();
        renderAdminTable();
        showToast('Acceso a Administración concedido', 'success');
    } else {
        if (pinError) pinError.style.display = 'block';
        if (pinInput) {
            pinInput.focus();
            pinInput.select();
        }
    }
}

function lockAdminSession() {
    STATE.isAdminUnlocked = false;
    sessionStorage.removeItem('admin_unlocked');
    const authCard = document.getElementById('adminAuthCard');
    const dashContent = document.getElementById('adminDashboardContent');
    const pinError = document.getElementById('adminPinError');
    if (authCard) authCard.style.display = 'block';
    if (dashContent) dashContent.style.display = 'none';
    if (pinError) pinError.style.display = 'none';
    showToast('Sesión de administración bloqueada', 'info');
}

function promptChangeAdminPin() {
    const current = prompt('Introduzca el PIN actual de administración:');
    if (current === null) return;
    if (current.trim() !== getAdminPin()) {
        alert('El PIN actual introducido no es correcto.');
        return;
    }
    const next = prompt('Introduzca el nuevo PIN (mínimo 4 caracteres):');
    if (next === null) return;
    const clean = next.trim();
    if (clean.length < 4) {
        alert('El PIN debe tener un mínimo de 4 caracteres.');
        return;
    }
    localStorage.setItem('admin_pin', clean);
    alert('PIN de administración actualizado correctamente.');
    showToast('PIN actualizado con éxito', 'success');
}

function populateAdminSelects() {
    const sel = document.getElementById('filterAdminHotel');
    if (!sel) return;
    const cur = sel.value;
    const h = [...new Set(STATE.incidencias.map(i => i.hotel).filter(Boolean))];
    sel.innerHTML = '<option value="all">Todos los Hoteles</option>' + h.map(x => `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`).join('');
    sel.value = cur;
}

function renderAdminTable() {
    if (!STATE.isAdminUnlocked) return;
    const tbody = document.getElementById('adminTableBody');
    if (!tbody) return;

    tbody.innerHTML = '';
    const { search, hotel, tipo, estado, fechaDesde, fechaHasta, sortFecha } = STATE.adminFilters;
    const filtered = STATE.incidencias.filter(item => {
        const term = search.toLowerCase();
        const matchesSearch = !term ||
            (item.descripcion || '').toLowerCase().includes(term) ||
            (item.usuario_registro || '').toLowerCase().includes(term) ||
            (item.id_original || '').toString().toLowerCase().includes(term) ||
            (item.cliente || '').toLowerCase().includes(term) ||
            (item.departamento || '').toLowerCase().includes(term);

        const matchesHotel = hotel === 'all' || item.hotel === hotel;
        const matchesTipo = tipo === 'all' || item.tipo === tipo;
        const matchesEstado = estado === 'all' || item.estado === estado;
        const matchesDate = matchesDateRange(item.fecha_creacion, fechaDesde, fechaHasta);
        return matchesSearch && matchesHotel && matchesTipo && matchesEstado && matchesDate;
    });

    filtered.sort((a, b) => {
        const da = new Date(a.fecha_creacion).getTime() || 0;
        const db = new Date(b.fecha_creacion).getTime() || 0;
        return sortFecha === 'asc' ? da - db : db - da;
    });

    const thAdminSort = document.getElementById('thAdminSortFecha');
    if (thAdminSort) {
        thAdminSort.innerHTML = `ID / Fecha <i class="fa-solid fa-arrow-${sortFecha === 'asc' ? 'up-short-wide' : 'down-wide-short'}" style="margin-left: 4px; color: var(--primary);"></i>`;
    }

    const adminTotalEl = document.getElementById('adminTotalCount');
    if (adminTotalEl) adminTotalEl.innerText = filtered.length;

    if (!filtered.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="empty-table">No se encontraron incidencias en administración con los filtros y fechas aplicados.</td>
            </tr>
        `;
        return;
    }

    filtered.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <strong>#${escapeHtml(item.id_original || 'S/N')}</strong><br>
                <small style="color:var(--text-muted); font-size:0.75rem;">${formatDate(item.fecha_creacion)}</small>
            </td>
            <td><span class="font-medium">${escapeHtml(item.hotel || '-')}</span></td>
            <td>${escapeHtml(item.tipo || '-')}</td>
            <td>${escapeHtml(item.departamento || '-')}</td>
            <td style="max-width: 250px;">
                <div class="truncate" title="${escapeHtml(item.descripcion || '')}">
                    ${escapeHtml(item.descripcion || '')}
                </div>
            </td>
            <td><span class="badge ${getBadgeClass(item.estado)}">${escapeHtml(item.estado || 'Pendiente')}</span></td>
            <td><span style="color: var(--text-muted); font-size:0.85rem;">${escapeHtml(item.responsable || 'Sin asignar')}</span></td>
            <td style="text-align: center;">
                <button class="btn-delete-row" type="button" data-id="${escapeHtml(item.id)}" title="Eliminar definitivamente">
                    <i class="fa-solid fa-trash-can"></i> Eliminar
                </button>
            </td>
        `;

        tr.addEventListener('click', (event) => {
            const deleteBtn = event.target.closest('.btn-delete-row');
            if (deleteBtn) {
                event.stopPropagation();
                confirmarEliminarIncidencia(item.id);
            } else {
                openIncidentModal(item.id);
            }
        });

        tbody.appendChild(tr);
    });
}

function confirmarEliminarIncidencia(id) {
    const item = STATE.incidencias.find(x => x.id === id);
    if (!item) return;

    STATE.pendingDeleteIncidentId = id;
    const previewEl = document.getElementById('deleteIncidentPreview');
    if (previewEl) {
        previewEl.innerHTML = `
            <div style="font-size:0.9rem; margin-bottom:0.4rem; display:flex; justify-content:space-between; align-items:center;">
                <strong>#${escapeHtml(item.id_original || 'S/N')} · ${escapeHtml(item.tipo || 'Incidencia')}</strong>
                <span class="badge ${getBadgeClass(item.estado)}">${escapeHtml(item.estado || 'Pendiente')}</span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-muted); margin-bottom:0.5rem;">
                <span>${escapeHtml(item.hotel || '-')}</span> · <span>${escapeHtml(item.departamento || '-')}</span> · <span>${formatDateTime(item.fecha_creacion)}</span>
            </div>
            <div style="font-size:0.85rem; color:var(--text-main); max-height:80px; overflow-y:auto; background:white; padding:0.5rem; border-radius:0.375rem; border:1px solid var(--border);">
                ${escapeHtml(item.descripcion || 'Sin descripción')}
            </div>
        `;
    }

    const modal = document.getElementById('deleteConfirmModal');
    if (modal) {
        modal.classList.add('active');
        modal.setAttribute('aria-hidden', 'false');
    }
}

function closeDeleteModal() {
    const modal = document.getElementById('deleteConfirmModal');
    if (modal) {
        modal.classList.remove('active');
        modal.setAttribute('aria-hidden', 'true');
    }
    STATE.pendingDeleteIncidentId = null;
}

async function ejecutarEliminacionIncidencia() {
    const id = STATE.pendingDeleteIncidentId;
    if (!id) return;

    const item = STATE.incidencias.find(x => x.id === id);
    const idOriginal = item ? (item.id_original || item.id) : id;

    // 1. Eliminar archivos adjuntos asociados en IndexedDB
    if (item && Array.isArray(item.attachments)) {
        for (const att of item.attachments) {
            if (att.id) await deleteAttachment(att.id);
        }
    }

    // 2. Eliminar documento de Firebase Firestore
    if (typeof db !== 'undefined') {
        try {
            await db.collection('incidencias').doc(id).delete();
        } catch (err) {
            console.error("Error al eliminar documento en Firestore:", err);
        }
    }

    // 3. Eliminar de STATE en memoria y LocalStorage
    STATE.incidencias = STATE.incidencias.filter(x => x.id !== id);
    STATE.lastUpdate = new Date().toISOString();
    saveState();

    // 4. Si el modal de detalle estaba abierto con esta incidencia, cerrarlo
    if (currentModalIncidentId === id) {
        closeIncidentModal();
    }

    // 5. Cerrar modal de confirmación
    closeDeleteModal();

    // 6. Refrescar todas las vistas (Dashboard, KPIs, Kanban, Tabla, Análisis y Admin)
    renderDashboard();
    renderAdminTable();

    // 7. Notificación de confirmación
    alert(`Incidencia #${idOriginal} eliminada definitivamente de todos los sitios.`);
}

function solicitarEliminarDesdeModal() {
    if (!currentModalIncidentId) return;
    
    // Si la sesión de administración no está desbloqueada, pedir el PIN
    if (!STATE.isAdminUnlocked) {
        const enteredPin = prompt('Acción restringida a Administración.\nIntroduzca el PIN de administración (PIN por defecto: 1234):');
        if (enteredPin === null) return; // cancelado
        if (enteredPin.trim() !== getAdminPin()) {
            alert('PIN de administración incorrecto. Solo administración puede eliminar incidencias.');
            return;
        }
        // Desbloquear sesión de administración
        STATE.isAdminUnlocked = true;
        sessionStorage.setItem('admin_unlocked', 'true');
        const authCard = document.getElementById('adminAuthCard');
        const dashContent = document.getElementById('adminDashboardContent');
        if (authCard) authCard.style.display = 'none';
        if (dashContent) dashContent.style.display = 'block';
    }
    
    // Abrir confirmación de eliminación para la incidencia actual
    confirmarEliminarIncidencia(currentModalIncidentId);
}



