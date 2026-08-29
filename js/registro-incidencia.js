const DB_KEY = 'incidencias_db';

document.addEventListener('DOMContentLoaded', () => {
    updateCurrentTotal();
    setupPickers();
    updateRecordType('Queja');
    document.getElementById('standaloneIncidentForm').addEventListener('submit', saveIncident);
    document.getElementById('btnClearForm').addEventListener('click', () => {
        document.getElementById('standaloneIncidentForm').reset();
        document.getElementById('formHotel').value = 'Secotel Guadiana';
        document.getElementById('formTipo').value = 'Queja';
        setActiveButton('.center-card', 'center', 'Secotel Guadiana');
        setActiveButton('.type-card', 'type', 'Queja');
        updateRecordType('Queja');
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
        responsable: valueOfOptional('formResponsableInicial'),
        estado: 'Pendiente',
        accion: buildActionSummary(),
        cliente: valueOfOptional('formCliente') || valueOfOptional('formClienteIncidencia'),
        correo: '',
        solicita_respuesta: valueOfOptional('formSolicitaRespuesta'),
        telefono: valueOfOptional('formTelefono'),
        correo_respuesta: valueOfOptional('formCorreoRespuesta'),
        fecha_cierre: '',
        notas_internas: buildInternalNotes()
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
