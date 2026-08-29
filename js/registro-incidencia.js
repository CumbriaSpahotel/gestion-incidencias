const DB_KEY = 'incidencias_db';
const ATTACHMENT_DB_NAME = 'IncidenciasAttachmentsDB';
const ATTACHMENT_STORE = 'attachments';
const NOTIFICATION_EMAIL = 'comunicaciones@hotelguadiana.es';
const NOTIFICATION_WEBHOOK_URL = 'https://formspree.io/f/xvgawqya';

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

    const state = loadState();
    const now = new Date();
    const localId = `local_${Date.now()}`;
    const attachments = await saveAttachments(localId);
    const item = {
        id: localId,
        source: 'local',
        id_original: `WEB-${String(getNextLocalNumber(state.items)).padStart(4, '0')}`,
        fecha_creacion: now,
        usuario_registro: valueOf('formUsuario') || 'Registro web',
        hotel: valueOf('formHotel'),
        tipo: valueOf('formTipo'),
        departamento: valueOf('formDepartamento'),
        descripcion: valueOf('formDescripcion'),
        responsable: valueOfOptional('formResponsableInicial'),
        estado: 'Pendiente',
        accion: buildActionSummary(),
        cliente: valueOfOptional('formCliente') || valueOfOptional('formClienteIncidencia'),
        correo: '',
        solicita_respuesta: valueOfOptional('formSolicitaRespuesta'),
        telefono: valueOfOptional('formTelefono'),
        correo_respuesta: valueOfOptional('formCorreoRespuesta'),
        fecha_cierre: '',
        notas_internas: buildInternalNotes(),
        attachments: attachments.map(({ id, name, type, size }) => ({ id, name, type, size }))
    };

    state.items.unshift(item);
    state.lastUpdate = now;
    saveState(state);
    updateCurrentTotal();
    showMessage(`Incidencia ${item.id_original} guardada correctamente.`);
    document.getElementById('standaloneIncidentForm').reset();
    document.getElementById('formHotel').value = 'Secotel Guadiana';
    document.getElementById('formTipo').value = 'Queja';
    setActiveButton('.center-card', 'center', 'Secotel Guadiana');
    setActiveButton('.type-card', 'type', 'Queja');
    updateRecordType('Queja');
    renderSelectedAttachments();

    let notificationSent = false;
    if (document.getElementById('formNotifyEmail').checked) {
        notificationSent = await sendAutomaticNotification(item);
    }
    if (document.getElementById('formNotifyEmail').checked && NOTIFICATION_WEBHOOK_URL) {
        showMessage(notificationSent
            ? `Incidencia ${item.id_original} guardada y aviso enviado.`
            : `Incidencia ${item.id_original} guardada. No se pudo enviar el aviso automático.`);
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

async function saveAttachments(incidentId) {
    const files = Array.from(document.getElementById('formAdjuntos').files || []);
    if (!files.length) return [];

    const records = files.map((file, index) => ({
        id: `${incidentId}_att_${index}`,
        incidentId,
        name: file.name,
        type: file.type || 'application/octet-stream',
        size: file.size,
        createdAt: new Date().toISOString(),
        blob: file
    }));

    const db = await openAttachmentDB();
    await Promise.all(records.map(record => putAttachment(db, record)));
    db.close();
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
    document.getElementById('currentTotal').textContent = loadState().items.length;
}

function showMessage(message) {
    const el = document.getElementById('saveMessage');
    el.textContent = message;
    el.classList.toggle('active', Boolean(message));
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
