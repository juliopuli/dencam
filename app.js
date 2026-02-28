/**
 * app.js - Versión 2.1 (Light Theme & Refinements)
 * Gestión de Atenciones con Interfaz de Pestañas y Búsqueda Global.
 */

const BRIDGE_URL = "http://localhost:8080";
let currentPerson = null;
let people = [];
let appConfig = null;

// UI Elements
const ui = {
    // Header & Status
    connectionStatus: document.getElementById('connection-status'),
    statusDot: document.querySelector('.status-dot'),
    statusText: document.querySelector('.status-text'),
    loginBtn: document.getElementById('login-btn'),
    logoutBtn: document.getElementById('logout-btn'),
    userInfo: document.getElementById('user-info'),

    // Config
    configSection: document.getElementById('config-section'),
    welcomeSection: document.getElementById('welcome-section'),
    mainSection: document.getElementById('main-section'),
    filePath: document.getElementById('file-path'),
    recordPath: document.getElementById('record-path'),
    saveConfigBtn: document.getElementById('save-config'),

    // Search
    searchPeople: document.getElementById('search-people'),
    searchResults: document.getElementById('search-results'),

    // Selected User Bar
    userBar: document.getElementById('selected-user-bar'),
    barNombre: document.getElementById('bar-nombre'),
    barApellidos: document.getElementById('bar-apellidos'),
    barDni: document.getElementById('bar-dni'),
    barHab: document.getElementById('bar-hab'),

    // Tabs
    tabLinks: document.querySelectorAll('.tab-link'),
    tabPanes: document.querySelectorAll('.tab-pane'),

    // Tab Panes Content
    detailsGrid: document.getElementById('person-details-grid'),
    selectedPersonHeader: document.getElementById('selected-person-header'),
    attendanceForm: document.getElementById('attendance-form')
};

// Initialization
async function init() {
    const savedConfig = JSON.parse(localStorage.getItem('bridge_config'));
    if (savedConfig) {
        appConfig = savedConfig;
        ui.filePath.value = savedConfig.filePath || '';
        ui.recordPath.value = savedConfig.recordPath || '';
        showMainApp();
    } else {
        ui.welcomeSection.classList.remove('hidden');
    }

    setupEventListeners();
    setupTabs();
}

function showMainApp() {
    ui.welcomeSection.classList.add('hidden');
    ui.configSection.classList.add('hidden');
    ui.mainSection.classList.remove('hidden');
    ui.userInfo.classList.remove('hidden');
    ui.loginBtn.classList.add('hidden');

    loadPeople();
    // history loading will be tab-specific in the future
}

// --- Status Indicator Management ---
function setStatus(status, text) {
    if (status === 'loading') {
        ui.statusDot.className = 'status-dot red';
        ui.statusText.textContent = text || 'Conectando...';
    } else if (status === 'ready') {
        ui.statusDot.className = 'status-dot green';
        ui.statusText.textContent = text || 'Conectado';
    } else {
        ui.statusDot.className = 'status-dot red';
        ui.statusText.textContent = text || 'Error';
    }
}

// --- Data Loading ---
async function loadPeople() {
    setStatus('loading', 'Cargando registros...');
    try {
        const url = `${BRIDGE_URL}/people?path=${encodeURIComponent(appConfig.filePath)}`;
        const response = await fetch(url);
        if (!response.ok) throw new Error("Fallo en la respuesta del bridge");
        const data = await response.json();

        if (data.error) throw new Error(data.error);

        people = data;
        setStatus('ready', `${people.length} registros`);
    } catch (error) {
        console.error("Load people error:", error);
        setStatus('error', 'Error de conexión');
        showToast("No se pudo conectar con el Excel Madre", "error");
    }
}

// --- Search Logic ---
function setupSearch() {
    ui.searchPeople.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (query.length < 2) {
            ui.searchResults.classList.add('hidden');
            return;
        }

        const filtered = people.filter(p =>
            p.nombre.toLowerCase().includes(query) ||
            p.apellidos.toLowerCase().includes(query) ||
            p.dni.toLowerCase().includes(query) ||
            p.hab.toLowerCase().includes(query)
        ).slice(0, 10);

        renderSearchResults(filtered);
    });

    document.addEventListener('click', (e) => {
        if (!ui.searchPeople.contains(e.target) && !ui.searchResults.contains(e.target)) {
            ui.searchResults.classList.add('hidden');
        }
    });
}

function renderSearchResults(results) {
    ui.searchResults.innerHTML = '';
    if (results.length === 0) {
        ui.searchResults.innerHTML = '<div class="search-result-item">No se encontraron coincidencias</div>';
    } else {
        results.forEach(person => {
            const div = document.createElement('div');
            div.className = 'search-result-item';
            div.innerHTML = `
                <div>
                    <span class="name">${person.nombre} ${person.apellidos}</span>
                    <span class="meta">${person.dni} | ${person.pais}</span>
                </div>
                <span class="room">${person.hab}</span>
            `;
            div.onclick = () => {
                selectPerson(person);
                ui.searchResults.classList.add('hidden');
                ui.searchPeople.value = ""; // Clear search after selection as requested by bar presence
            };
            ui.searchResults.appendChild(div);
        });
    }
    ui.searchResults.classList.remove('hidden');
}

// --- Selection & Tabs Logic ---
function setupTabs() {
    ui.tabLinks.forEach(link => {
        link.onclick = () => {
            const tabId = link.getAttribute('data-tab');
            ui.tabLinks.forEach(l => l.classList.remove('active'));
            ui.tabPanes.forEach(p => p.classList.remove('active'));
            link.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        };
    });
}

function selectPerson(person) {
    currentPerson = person;

    // Update Top User Bar
    ui.barNombre.textContent = person.nombre;
    ui.barApellidos.textContent = person.apellidos;
    ui.barDni.textContent = person.dni;
    ui.barHab.textContent = person.hab;
    ui.userBar.classList.remove('hidden');

    // Update "Datos Generales" Tab
    renderPersonDetails(person);

    // Update "Enfermería" Tab internal header
    if (ui.selectedPersonHeader) {
        ui.selectedPersonHeader.innerHTML = `<strong>Registrando atención para:</strong> ${person.nombre} ${person.apellidos}`;
        ui.selectedPersonHeader.classList.remove('empty');
    }
    ui.attendanceForm.classList.remove('hidden');

    // Pre-fill fields
    const piSelect = document.getElementById('attendance-pi');
    const csSelect = document.getElementById('attendance-conf-salida');
    if (piSelect) piSelect.value = person.pi || "";
    if (csSelect) csSelect.value = person.conf_salida || "";

    document.getElementById('observations').value = "";
    showToast(`Usuario seleccionado: ${person.nombre}`);

    // Auto-switch to "Datos Generales" if on a different tab? 
    // User didn't specify, but usually good to show the details.
}

function renderPersonDetails(p) {
    const fields = [
        { label: 'Nombre', value: p.nombre },
        { label: 'Apellidos', value: p.apellidos },
        { label: 'DNI / NIE', value: p.dni },
        { label: 'Habitación', value: p.hab },
        { label: 'Nº SIRIA', value: p.siria },
        { label: 'País de Origen', value: p.pais },
        { label: 'Fecha Nacimiento', value: p.fn },
        { label: 'P.I. (SAP)', value: p.pi || 'No asignado' },
        { label: 'Confirmación Salida', value: p.conf_salida || 'Pendiente' }
    ];

    ui.detailsGrid.innerHTML = fields.map(f => `
        <div class="detail-item">
            <span class="detail-label">${f.label}</span>
            <span class="detail-value">${f.value}</span>
        </div>
    `).join('');
}

// --- Event Handlers ---
async function handleAttendanceSubmit(e) {
    e.preventDefault();
    if (!currentPerson) return;

    const btn = document.getElementById('save-attendance');
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "Guardando...";

    try {
        const response = await fetch(`${BRIDGE_URL}/save`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                path: appConfig.recordPath,
                nombre: currentPerson.nombre,
                apellidos: currentPerson.apellidos,
                dni: currentPerson.dni,
                hab: currentPerson.hab,
                siria: currentPerson.siria,
                pais: currentPerson.pais,
                fn: currentPerson.fn,
                pi: document.getElementById('attendance-pi').value,
                conf_salida: document.getElementById('attendance-conf-salida').value,
                tipo: document.getElementById('attendance-type').value,
                observations: document.getElementById('observations').value
            })
        });

        const res = await response.json();
        if (res.error) throw new Error(res.error);

        showToast("Registro guardado con éxito");
        document.getElementById('observations').value = "";
    } catch (error) {
        showToast("Error al guardar: " + error.message, "error");
    } finally {
        btn.disabled = false;
        btn.textContent = originalText;
    }
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container') || document.body;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
}

function setupEventListeners() {
    ui.saveConfigBtn.onclick = () => {
        const config = {
            filePath: ui.filePath.value,
            recordPath: ui.recordPath.value
        };
        if (config.filePath && config.recordPath) {
            localStorage.setItem('bridge_config', JSON.stringify(config));
            window.location.reload();
        } else {
            showToast("Introduce todas las rutas", "error");
        }
    };

    ui.loginBtn.onclick = () => ui.configSection.classList.remove('hidden');
    ui.attendanceForm.onsubmit = handleAttendanceSubmit;
    setupSearch();
}

// Run!
init();
