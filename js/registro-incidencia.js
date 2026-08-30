const DB_KEY = 'incidencias_db';
const ATTACHMENT_DB_NAME = 'IncidenciasAttachmentsDB';
const ATTACHMENT_STORE = 'attachments';
const NOTIFICATION_EMAIL = 'comunicaciones@hotelguadiana.es';
const NOTIFICATION_WEBHOOK_URL = 'https://formspree.io/f/xbgjlpnr';

// ── CONFIGURACIÓN DE SYNOLOGY NAS ──────────────────────────────────────
// Rellena con tu dirección de QuickConnect o IP, p.ej. "https://tuempresa.synology.me"
const SYNOLOGY_URL = 'https://miempresa.synology.me'; 
// Usuario creado en tu DSM con permisos de escritura/subida en la carpeta de incidencias
const SYNOLOGY_USER = 'incidencias_upload';
const SYNOLOGY_PASS = 'CONTRASEÑA_DE_USUARIO_AQUI';
// Ruta exacta donde se guardarán los archivos en Synology
const SYNOLOGY_FOLDER = '/Carpeta de equipo/compartidas A B C/B/Reclamaciones';
const SYNOLOGY_TIMEOUT = 10000; // ms máximo por subida para evitar bloqueos


document.addEventListener('DOMContentLoaded', () => {
    updateCurrentTotal();
    setupPickers();
    updateRecordType('Queja');
    document.getElementById('standaloneIncidentForm').addEventListener('submit', saveIncident);
    document.getElementById('formAdjuntos').addEventListener('change', renderSelectedAttachments);
    document.getElementById('btnClearForm').addEventListener('click', () => {
        document.getElementById('standaloneIncidentForm').reset();
        document.getElementById('formHotel').value = 'Secotel Guadiana';
        document.getElementById('formTipo').value = 'Queja';
        setActiveButton('.center-card', 'center', 'Secotel Guadiana');
        setActiveButton('.type-card', 'type', 'Queja');
        updateRecordType('Queja');
        renderSelectedAttachments();
        showMessage('');
    });
});

const TYPE_CONFIG = {
    Queja: {
        title: 'Queja',
        description: 'Use este registro cuando un cliente manifiesta malestar por un servicio recibido.',
        label: 'Descripción de la queja',
        placeholder: 'Describa el motivo de la queja, quién la comunica, cuándo ocurrió y cómo se ha atendido.',
        actionLabel: 'Respuesta dada al cliente'
    },
    Reclamación: {
        title: 'Reclamación',
        description: 'Use este registro cuando requiere respuesta formal, trazabilidad documental o posible seguimiento de dirección.',
        label: 'Motivo de la reclamación',
        placeholder: 'Describa los hechos, petición del cliente, datos de contacto y cualquier referencia documental.',
        actionLabel: 'Respuesta o medida inicial'
    },
    Incidencia: {
        title: 'Incidencia',
        description: 'Use este registro para riesgos operativos, seguridad, averías relevantes o hechos con posible impacto legal.',
        label: 'Descripción de la incidencia',
        placeholder: 'Describa el hecho, zona afectada, personas implicadas, riesgo detectado y actuación inicial.',
        actionLabel: 'Medida inmediata'
    },
    Sugerencia: {
        title: 'Sugerencia',
        description: 'Use este registro para propuestas de mejora recibidas de clientes o del equipo.',
        label: 'Propuesta recibida',
        placeholder: 'Describa la sugerencia, el origen de la propuesta y el cambio que se plantea.',
        actionLabel: 'Acción propuesta'
    }
};

function setupPickers() {
    document.querySelectorAll('.center-card').forEach(button => {
        button.addEventListener('click', () => {
            document.getElementById('formHotel').value = button.dataset.center;
            setActiveButton('.center-card', 'center', button.dataset.center);
            updateCurrentTotal();
        });
    });

    document.querySelectorAll('.type-card').forEach(button => {
        button.addEventListener('click', () => {
            document.getElementById('formTipo').value = button.dataset.type;
            setActiveButton('.type-card', 'type', button.dataset.type);
            updateRecordType(button.dataset.type);
        });
    });
}

function setActiveButton(selector, dataKey, value) {
    document.querySelectorAll(selector).forEach(button => {
        button.classList.toggle('active', button.dataset[dataKey] === value);
    });
}

function updateRecordType(type) {
    const config = TYPE_CONFIG[type] || TYPE_CONFIG.Queja;
    document.getElementById('sideTitle').textContent = config.title;
    document.getElementById('sideDescription').textContent = config.description;
    document.getElementById('descriptionLabel').textContent = config.label;
    document.getElementById('formDescripcion').placeholder = config.placeholder;

    document.querySelectorAll('.context-panel').forEach(panel => {
        const shouldShow = panel.dataset.context === type || (type === 'Reclamación' && panel.dataset.context === 'Queja');
        panel.classList.toggle('active', shouldShow);
        panel.querySelectorAll('input, select, textarea').forEach(field => {
            field.disabled = !shouldShow;
        });
    });

    const actionLabel = document.querySelector('#contextQueja label.full span');
    if (actionLabel) actionLabel.textContent = config.actionLabel;
    updateCurrentTotal();
}

function loadState() {
    const saved = localStorage.getItem(DB_KEY);
    if (!saved) return { items: [], lastUpdate: null };
    try {
        const parsed = JSON.parse(saved);
        return {
            items: Array.isArray(parsed.items) ? parsed.items : [],
            lastUpdate: parsed.lastUpdate || null
        };
    } catch (error) {
        return { items: [], lastUpdate: null };
    }
}

function saveState(state) {
    localStorage.setItem(DB_KEY, JSON.stringify(state));
}

async function saveIncident(event) {
    event.preventDefault();

    const form = event.target;
    form.classList.add('form-submitted');

    const submitBtn = form.querySelector('button[type="submit"]');
    const originalBtnHtml = submitBtn ? submitBtn.innerHTML : '';
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Guardando...';
    }

    try {
        const state = loadState();
        const now = new Date();
        const localId = `local_${Date.now()}`;
        const attachments = await saveAttachments(localId);
        
        const item = {
            id: localId,
            source: 'local',
            id_original: `WEB-${String(getNextLocalNumber(state.items)).padStart(4, '0')}`,
            fecha_creacion: now.toISOString(),
            usuario_registro: valueOf('formUsuario') || 'Registro web',
            hotel: valueOf('formHotel'),
            tipo: valueOf('formTipo'),
            departamento: valueOf('formDepartamento'),
            descripcion: valueOf('formDescripcion'),
            responsable: valueOfOptional('formResponsableInicial'),
            estado: 'Pendiente',
            accion: buildActionSummary(),
            cliente: valueOfOptional('formCliente') || valueOfOptional('formClienteIncidencia'),
            correo: valueOfOptional('formCorreoRespuesta') || '',
            solicita_respuesta: valueOfOptional('formSolicitaRespuesta'),
            telefono: valueOfOptional('formTelefono'),
            correo_respuesta: valueOfOptional('formCorreoRespuesta'),
            fecha_cierre: '',
            notas_internas: buildInternalNotes(),
            attachments: attachments.map(({ id, name, type, size, url, synologyPath }) => ({
                id,
                name,
                type,
                size,
                url: url || null,
                synologyPath: synologyPath || null
            }))
        };

        state.items.unshift(item);
        state.lastUpdate = now.toISOString();
        saveState(state);
        
        if (typeof db !== 'undefined' && db) {
            db.collection('incidencias').doc(item.id).set(item).catch(err => {
                console.warn('Error guardando en Firestore:', err);
            });
        }

        updateCurrentTotal();
        showMessage(`¡Incidencia ${item.id_original} guardada con éxito!`);
        
        // Reset form y quitar clase de validación visual
        const standaloneForm = document.getElementById('standaloneIncidentForm');
        standaloneForm.reset();
        standaloneForm.classList.remove('form-submitted');
        document.getElementById('formHotel').value = 'Secotel Guadiana';
        document.getElementById('formTipo').value = 'Queja';
        setActiveButton('.center-card', 'center', 'Secotel Guadiana');
        setActiveButton('.type-card', 'type', 'Queja');
        updateRecordType('Queja');
        renderSelectedAttachments();
        window.scrollTo({ top: 0, behavior: 'smooth' });

        sendAutomaticNotification(item).then(notificationSent => {
            if (NOTIFICATION_WEBHOOK_URL) {
                showMessage(notificationSent
                    ? `¡Incidencia ${item.id_original} guardada y aviso enviado!`
                    : `Incidencia ${item.id_original} guardada.`);
            }
        }).catch(console.warn);

    } catch (err) {
        console.error("Error al guardar incidencia:", err);
        showMessage(`Error al guardar: ${err.message || 'Inténtelo de nuevo'}`);
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalBtnHtml;
        }
    }
}

function renderSelectedAttachments() {
    const list = document.getElementById('attachmentList');
    const files = Array.from(document.getElementById('formAdjuntos').files || []);
    if (!files.length) {
        list.innerHTML = '';
        return;
    }

    list.innerHTML = files.map(file => `
        <div class="attachment-chip">
            <i class="fa-solid ${file.type.startsWith('image/') ? 'fa-image' : 'fa-file'}"></i>
            <span>${escapeHtml(file.name)}</span>
            <small>${formatBytes(file.size)}</small>
        </div>
    `).join('');
}

function fileToDataUrl(file, maxWidth = 1000, maxHeight = 1000, quality = 0.7) {
    return new Promise((resolve) => {
        if (!file.type || !file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => resolve(null);
            reader.readAsDataURL(file);
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                let { width, height } = img;
                if (width > maxWidth || height > maxHeight) {
                    if (width > height) {
                        height = Math.round((height * maxWidth) / width);
                        width = maxWidth;
                    } else {
                        width = Math.round((width * maxHeight) / height);
                        height = maxHeight;
                    }
                }
                const canvas = document.createElement('canvas');
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, width, height);
                resolve(canvas.toDataURL('image/jpeg', quality));
            };
            img.onerror = () => resolve(reader.result);
            img.src = e.target.result;
        };
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(file);
    });
}

// ── Synology FileStation API ──────────────────────────────────────────

async function synologyLogin() {
    const url = `${SYNOLOGY_URL}/webapi/auth.cgi?` + new URLSearchParams({
        api: 'SYNO.API.Auth',
        version: '3',
        method: 'login',
        account: SYNOLOGY_USER,
        passwd: SYNOLOGY_PASS,
        session: 'FileStation',
        format: 'sid'
    });
    
    const res = await fetch(url, { credentials: 'omit' });
    const data = await res.json();
    if (!data.success) {
        throw new Error(`Synology login error: ${JSON.stringify(data.error)}`);
    }
    return data.data.sid;
}

async function synologyLogout(sid) {
    try {
        const url = `${SYNOLOGY_URL}/webapi/auth.cgi?` + new URLSearchParams({
            api: 'SYNO.API.Auth',
            version: '1',
            method: 'logout',
            session: 'FileStation',
            _sid: sid
        });
        await fetch(url, { credentials: 'omit' });
    } catch (e) {
        console.warn("Synology logout error:", e);
    }
}

async function synologyUploadFile(sid, file, incidentId) {
    const prefixedName = `${incidentId}__${file.name}`;
    const formData = new FormData();
    formData.append('api', 'SYNO.FileStation.Upload');
    formData.append('version', '2');
    formData.append('method', 'upload');
    formData.append('path', SYNOLOGY_FOLDER);
    formData.append('create_parents', 'true');
    formData.append('overwrite', 'false');
    formData.append('file', file, prefixedName);
    formData.append('_sid', sid);

    const uploadUrl = `${SYNOLOGY_URL}/webapi/entry.cgi?_sid=${encodeURIComponent(sid)}`;
    const res = await fetch(uploadUrl, {
        method: 'POST',
        body: formData,
        credentials: 'omit'
    });
    const data = await res.json();
    if (!data.success) {
        throw new Error(`Synology upload error: ${JSON.stringify(data.error)}`);
    }
    return `${SYNOLOGY_URL}/webapi/entry.cgi?api=SYNO.FileStation.Download&version=2&method=download&path=${encodeURIComponent(SYNOLOGY_FOLDER + '/' + prefixedName)}&mode=open`;
}

async function uploadFilesToSynology(files, incidentId) {
    if (!SYNOLOGY_URL || SYNOLOGY_URL === 'https://miempresa.synology.me') {
        return null; // Omitir si no está configurado o tiene la URL por defecto
    }

    let sid = null;
    const results = [];
    try {
        const loginPromise = synologyLogin();
        const timeoutPromise = new Promise((_, reject) => 
            setTimeout(() => reject(new Error("Timeout conectando a Synology")), SYNOLOGY_TIMEOUT)
        );
        sid = await Promise.race([loginPromise, timeoutPromise]);

        for (const file of files) {
            try {
                const downloadURL = await synologyUploadFile(sid, file, incidentId);
                results.push({ name: file.name, url: downloadURL, path: `${SYNOLOGY_FOLDER}/${incidentId}__${file.name}`, ok: true });
            } catch (fileErr) {
                console.warn(`Error al subir archivo ${file.name} a Synology:`, fileErr);
                results.push({ name: file.name, url: null, path: null, ok: false });
            }
        }
    } catch (err) {
        console.warn("No se pudo conectar a Synology (se usará Base64 local):", err.message || err);
        return null;
    } finally {
        if (sid) {
            await synologyLogout(sid);
        }
    }
    return results;
}

// ── Guardado de adjuntos ──────────────────────────────────────────────

async function saveAttachments(incidentId) {
    const files = Array.from(document.getElementById('formAdjuntos').files || []);
    if (!files.length) return [];

    // Intentamos subir a Synology primero
    const synologyResults = await uploadFilesToSynology(files, incidentId);
    const records = [];
    
    for (let index = 0; index < files.length; index++) {
        const file = files[index];
        const recordId = `${incidentId}_att_${index}`;
        const synResult = synologyResults ? synologyResults[index] : null;
        
        let downloadURL = synResult && synResult.ok ? synResult.url : null;
        let synologyPath = synResult && synResult.ok ? synResult.path : null;

        // Fallback si falló o no se subió a Synology
        if (!downloadURL) {
            try {
                downloadURL = await fileToDataUrl(file);
            } catch (e) {
                console.warn("Error generando Data URL para adjunto:", e);
            }
        }
        
        const record = {
            id: recordId,
            incidentId,
            name: file.name,
            type: file.type || 'application/octet-stream',
            size: file.size,
            createdAt: new Date().toISOString(),
            blob: file,
            url: downloadURL,
            synologyPath: synologyPath
        };
        
        records.push(record);
    }

    try {
        const localDb = await openAttachmentDB();
        await Promise.all(records.map(record => putAttachment(localDb, record)));
        localDb.close();
    } catch (e) {
        console.warn("Error guardando adjuntos en IndexedDB:", e);
    }
    
    return records;
}

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

function putAttachment(db, record) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(ATTACHMENT_STORE, 'readwrite');
        tx.objectStore(ATTACHMENT_STORE).put(record);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

function valueOf(id) {
    return document.getElementById(id).value.trim();
}

function valueOfOptional(id) {
    const element = document.getElementById(id);
    return element && !element.disabled ? element.value.trim() : '';
}

function buildActionSummary() {
    const type = valueOf('formTipo');
    if (type === 'Incidencia') return valueOfOptional('formAccionIncidencia');
    if (type === 'Sugerencia') return valueOfOptional('formBeneficio');
    return valueOfOptional('formAccion');
}

function buildInternalNotes() {
    const notes = [
        ['Canal', valueOfOptional('formCanal')],
        ['Compensación ofrecida', valueOfOptional('formCompensacion')],
        ['Gravedad', valueOfOptional('formGravedad')],
        ['Requiere cierre formal', valueOfOptional('formCierreFormal')],
        ['Área de mejora', valueOfOptional('formAreaMejora')],
        ['Prioridad sugerida', valueOfOptional('formPrioridadSugerida')]
    ].filter(([, value]) => value);

    return notes.map(([label, value]) => `${label}: ${value}`).join('\n');
}

function getNextLocalNumber(items) {
    const localNumbers = items
        .map(item => (item.id_original || '').toString().match(/^WEB-(\d+)$/))
        .filter(Boolean)
        .map(match => Number(match[1]));
    return localNumbers.length ? Math.max(...localNumbers) + 1 : 1;
}

function updateCurrentTotal() {
    const selectedCenter = valueOf('formHotel');
    const selectedType = valueOf('formTipo');
    const items = loadState().items;
    const total = items.filter(item => (
        matchesCenter(item.hotel, selectedCenter) && matchesType(item.tipo, selectedType)
    )).length;
    const centerLabel = getCenterLabel(selectedCenter);

    document.getElementById('currentTotal').textContent = total;
    document.getElementById('currentFilterText').textContent = `${centerLabel} · ${selectedType}`;
}

function matchesCenter(itemCenter, selectedCenter) {
    const item = normalizeText(itemCenter);
    const selected = normalizeText(selectedCenter);
    if (!item || !selected) return false;

    if (selected.includes('guadiana')) return item.includes('guadiana');
    if (selected.includes('bienestar')) return item.includes('bienestar');
    if (selected.includes('cumbria')) return item.includes('cumbria') && !item.includes('bienestar');
    return item === selected;
}

function matchesType(itemType, selectedType) {
    return normalizeText(itemType) === normalizeText(selectedType);
}

function getCenterLabel(center) {
    const normalized = normalizeText(center);
    if (normalized.includes('guadiana')) return 'Guadiana';
    if (normalized.includes('bienestar')) return 'Cumbria Bienestar';
    if (normalized.includes('cumbria')) return 'Cumbria Hotel';
    return center || 'Centro';
}

function normalizeText(value) {
    return (value || '')
        .toString()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function showMessage(message) {
    const el = document.getElementById('saveMessage');
    if (el) {
        el.textContent = message;
        el.classList.toggle('active', Boolean(message));
    }
    const formEl = document.getElementById('formSaveMessage');
    if (formEl) {
        formEl.textContent = message;
        formEl.classList.toggle('active', Boolean(message));
    }
}

async function sendAutomaticNotification(item) {
    if (!NOTIFICATION_WEBHOOK_URL) return false;
    const subject = `[Q-Centros] Nuevo registro ${item.tipo} - ${item.hotel}`;
    const attachmentNames = Array.isArray(item.attachments) && item.attachments.length
        ? item.attachments.map(file => file.name)
        : [];

    try {
        const response = await fetch(NOTIFICATION_WEBHOOK_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify({
                _subject: subject,
                email_destino: NOTIFICATION_EMAIL,
                id_registro: item.id_original,
                tipo: item.tipo,
                centro: item.hotel,
                zona_servicio: item.departamento,
                fecha: new Date(item.fecha_creacion).toLocaleString('es-ES'),
                registrado_por: item.usuario_registro,
                cliente: item.cliente || '-',
                solicita_respuesta: item.solicita_respuesta || '-',
                telefono: item.telefono || '-',
                correo_respuesta: item.correo_respuesta || '-',
                descripcion: item.descripcion || '-',
                actuacion_inicial: item.accion || '-',
                datos_especificos: item.notas_internas || '-',
                adjuntos: attachmentNames.length ? attachmentNames.join(', ') : 'Sin adjuntos'
            })
        });
        return response.ok;
    } catch (error) {
        console.warn('Notification error', error);
        return false;
    }
}

function formatBytes(bytes) {
    if (!bytes) return '0 KB';
    const units = ['B', 'KB', 'MB', 'GB'];
    const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / Math.pow(1024, index);
    return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
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

// Dark Mode Setup
document.addEventListener('DOMContentLoaded', () => {
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
if (isDark) { icon.classList.replace('fa-moon', 'fa-sun'); span.innerText = 'Modo claro'; }
else { icon.classList.replace('fa-sun', 'fa-moon'); span.innerText = 'Modo oscuro'; }
});
}
});

