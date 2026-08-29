const DB_KEY = 'incidencias_db';

document.addEventListener('DOMContentLoaded', () => {
    updateCurrentTotal();
    document.getElementById('standaloneIncidentForm').addEventListener('submit', saveIncident);
    document.getElementById('btnClearForm').addEventListener('click', () => {
        document.getElementById('standaloneIncidentForm').reset();
        document.getElementById('formHotel').value = 'Secotel Guadiana';
        document.getElementById('formTipo').value = 'Queja';
        showMessage('');
    });
});

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

function saveIncident(event) {
    event.preventDefault();

    const state = loadState();
    const now = new Date();
    const item = {
        id: `local_${Date.now()}`,
        source: 'local',
        id_original: `WEB-${String(getNextLocalNumber(state.items)).padStart(4, '0')}`,
        fecha_creacion: now,
        usuario_registro: valueOf('formUsuario') || 'Registro web',
        hotel: valueOf('formHotel'),
        tipo: valueOf('formTipo'),
        departamento: valueOf('formDepartamento'),
        descripcion: valueOf('formDescripcion'),
        responsable: '',
        estado: 'Pendiente',
        accion: valueOf('formAccion'),
        cliente: valueOf('formCliente'),
        correo: '',
        solicita_respuesta: valueOf('formSolicitaRespuesta'),
        telefono: valueOf('formTelefono'),
        correo_respuesta: valueOf('formCorreoRespuesta'),
        fecha_cierre: '',
        notas_internas: ''
    };

    state.items.unshift(item);
    state.lastUpdate = now;
    saveState(state);
    updateCurrentTotal();
    showMessage(`Incidencia ${item.id_original} guardada correctamente.`);
    document.getElementById('standaloneIncidentForm').reset();
    document.getElementById('formHotel').value = 'Secotel Guadiana';
    document.getElementById('formTipo').value = 'Queja';
}

function valueOf(id) {
    return document.getElementById(id).value.trim();
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
