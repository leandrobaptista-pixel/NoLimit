const DB_NAME = "cabinets-control-db";
const DB_VERSION = 4;
const UNIT_STORE = "units";
const PHOTO_STORE = "photos";
const USER_STORE = "users";
const SETTINGS_STORE = "settings";
const CLIENT_STORE = "clients";
const PROJECT_STORE = "projects";
const CONTACT_STORE = "contacts";
const CONTAINER_STORE = "containers";
const SESSION_KEY = "cc-session-user-id";

const STAGES = [
  { key: "warehouse", label: "Warehouse" },
  { key: "transportation", label: "Transportation" },
  { key: "siteDelivery", label: "Site Job" },
  { key: "distribution", label: "Distribuicao" },
  { key: "quality", label: "Qualidade" },
  { key: "installation", label: "Instalacao" },
];

const ROLE_LABEL = {
  admin: "Admin",
  warehouse: "Warehouse",
  transport: "Transport",
  qa: "QA",
  installer: "Installer",
};

const CONTACT_ROLE_LABEL = {
  subcontractor: "Sub Contractor",
  foreman: "Foreman",
  "project-manager": "Project Manager",
  client: "Cliente",
  other: "Other",
};

const CONTAINER_STATUS_LABEL = {
  scheduled: "Scheduled",
  "in-transit": "In Transit",
  arrived: "Arrived",
  delayed: "Delayed",
};

const STAGE_ROLE_ACCESS = {
  warehouse: ["warehouse"],
  transport: ["transportation", "siteDelivery"],
  qa: ["quality"],
  installer: ["distribution", "installation"],
};

const DEFAULT_SYNC = {
  supabaseUrl: "",
  supabaseAnonKey: "",
  tenant: "",
  autoSync: false,
  lastSyncAt: null,
};

let db;
let units = [];
let photos = [];
let users = [];
let clients = [];
let projects = [];
let contacts = [];
let containers = [];
let syncConfig = { ...DEFAULT_SYNC };
let photosByUnit = new Map();
let currentUser = null;
let deferredPrompt = null;
let autoSyncTimer = null;

const authView = document.getElementById("authView");
const setupPanel = document.getElementById("setupPanel");
const loginPanel = document.getElementById("loginPanel");
const setupForm = document.getElementById("setupForm");
const loginForm = document.getElementById("loginForm");

const appMain = document.getElementById("appMain");
const roleLine = document.getElementById("roleLine");
const permissionLine = document.getElementById("permissionLine");

const masterDataPanel = document.getElementById("masterDataPanel");
const clientForm = document.getElementById("clientForm");
const projectForm = document.getElementById("projectForm");
const contactForm = document.getElementById("contactForm");
const clientsTable = document.getElementById("clientsTable");
const projectsTable = document.getElementById("projectsTable");
const contactsTable = document.getElementById("contactsTable");
const projectClientSelect = document.getElementById("projectClientSelect");
const contactClientSelect = document.getElementById("contactClientSelect");
const contactProjectSelect = document.getElementById("contactProjectSelect");

const containerPanel = document.getElementById("containerPanel");
const containerForm = document.getElementById("containerForm");
const containerClientSelect = document.getElementById("containerClientSelect");
const containerProjectSelect = document.getElementById("containerProjectSelect");
const containersBoard = document.getElementById("containersBoard");
const containerTemplate = document.getElementById("containerTemplate");

const unitForm = document.getElementById("unitForm");
const unitEntryPanel = document.getElementById("unitEntryPanel");
const unitProjectSelect = document.getElementById("unitProjectSelect");
const unitClientName = document.getElementById("unitClientName");
const unitProjectName = document.getElementById("unitProjectName");
const unitJobSite = document.getElementById("unitJobSite");
const unitProjectHint = document.getElementById("unitProjectHint");

const unitsContainer = document.getElementById("unitsContainer");
const unitTemplate = document.getElementById("unitTemplate");
const statsPanel = document.getElementById("statsPanel");
const searchInput = document.getElementById("searchInput");
const filterSelect = document.getElementById("filterSelect");

const exportBtn = document.getElementById("exportBtn");
const importInput = document.getElementById("importInput");
const projectReportSelect = document.getElementById("projectReportSelect");
const projectReportBtn = document.getElementById("projectReportBtn");

const syncPanel = document.getElementById("syncPanel");
const syncConfigForm = document.getElementById("syncConfigForm");
const pushSyncBtn = document.getElementById("pushSyncBtn");
const pullSyncBtn = document.getElementById("pullSyncBtn");
const syncStatus = document.getElementById("syncStatus");

const usersPanel = document.getElementById("usersPanel");
const userForm = document.getElementById("userForm");
const usersTable = document.getElementById("usersTable");

const connectionBadge = document.getElementById("connectionBadge");
const userBadge = document.getElementById("userBadge");
const logoutBtn = document.getElementById("logoutBtn");
const installBtn = document.getElementById("installBtn");

const fmtDate = (iso) => {
  if (!iso) return "-";
  return new Date(iso).toLocaleString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
};

const uid = () => {
  if (crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
};

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function toNumber(value) {
  const num = Number(value || 0);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const now = new Date();
  const target = new Date(`${dateStr}T00:00:00`);
  const diff = target.getTime() - new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const dbRef = event.target.result;
      if (!dbRef.objectStoreNames.contains(UNIT_STORE)) dbRef.createObjectStore(UNIT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(PHOTO_STORE)) {
        const store = dbRef.createObjectStore(PHOTO_STORE, { keyPath: "id" });
        store.createIndex("unitId", "unitId", { unique: false });
      }
      if (!dbRef.objectStoreNames.contains(USER_STORE)) dbRef.createObjectStore(USER_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(SETTINGS_STORE)) dbRef.createObjectStore(SETTINGS_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CLIENT_STORE)) dbRef.createObjectStore(CLIENT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(PROJECT_STORE)) dbRef.createObjectStore(PROJECT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CONTACT_STORE)) dbRef.createObjectStore(CONTACT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CONTAINER_STORE)) dbRef.createObjectStore(CONTAINER_STORE, { keyPath: "id" });
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function tx(storeName, mode = "readonly") {
  return db.transaction(storeName, mode).objectStore(storeName);
}

function requestToPromise(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function getAll(storeName) {
  return requestToPromise(tx(storeName).getAll());
}

async function put(storeName, value) {
  return requestToPromise(tx(storeName, "readwrite").put(value));
}

async function del(storeName, id) {
  return requestToPromise(tx(storeName, "readwrite").delete(id));
}

function normalizeUnit(unit) {
  const normalized = { ...unit };
  normalized.stages = normalized.stages || {};
  for (const stage of STAGES) {
    normalized.stages[stage.key] = normalized.stages[stage.key] || { done: false, at: null };
  }
  normalized.checkItems = Array.isArray(normalized.checkItems) ? normalized.checkItems : [];
  normalized.deliveryQuality = normalized.deliveryQuality || "pending";
  normalized.installationStatus = normalized.installationStatus || "not-started";
  normalized.deliveryNotes = normalized.deliveryNotes || "";
  normalized.issuesText = normalized.issuesText || "";
  normalized.clientId = normalized.clientId || "";
  normalized.projectId = normalized.projectId || "";
  normalized.clientName = normalized.clientName || "";
  normalized.projectName = normalized.projectName || "";
  normalized.createdAt = normalized.createdAt || new Date().toISOString();
  normalized.updatedAt = normalized.updatedAt || normalized.createdAt;
  return normalized;
}

function normalizeClient(client) {
  return {
    id: client.id,
    name: client.name || "",
    contactPerson: client.contactPerson || "",
    phone: client.phone || "",
    email: client.email || "",
    createdAt: client.createdAt || new Date().toISOString(),
    updatedAt: client.updatedAt || client.createdAt || new Date().toISOString(),
  };
}

function normalizeProject(project) {
  return {
    id: project.id,
    clientId: project.clientId || "",
    name: project.name || "",
    code: project.code || "",
    address: project.address || "",
    createdAt: project.createdAt || new Date().toISOString(),
    updatedAt: project.updatedAt || project.createdAt || new Date().toISOString(),
  };
}

function normalizeContact(contact) {
  return {
    id: contact.id,
    name: contact.name || "",
    role: contact.role || "other",
    company: contact.company || "",
    clientId: contact.clientId || "",
    projectId: contact.projectId || "",
    phone: contact.phone || "",
    email: contact.email || "",
    createdAt: contact.createdAt || new Date().toISOString(),
    updatedAt: contact.updatedAt || contact.createdAt || new Date().toISOString(),
  };
}

function normalizeContainer(container) {
  return {
    id: container.id,
    containerCode: container.containerCode || "",
    manufacturer: container.manufacturer || "",
    clientId: container.clientId || "",
    projectId: container.projectId || "",
    etaDate: container.etaDate || "",
    arrivalStatus: container.arrivalStatus || "scheduled",
    qtyKitchens: toNumber(container.qtyKitchens),
    qtyVanities: toNumber(container.qtyVanities),
    qtyMedCabinets: toNumber(container.qtyMedCabinets),
    qtyCountertops: toNumber(container.qtyCountertops),
    notes: container.notes || "",
    materialItems: Array.isArray(container.materialItems) ? container.materialItems : [],
    movements: Array.isArray(container.movements) ? container.movements : [],
    createdAt: container.createdAt || new Date().toISOString(),
    updatedAt: container.updatedAt || container.createdAt || new Date().toISOString(),
  };
}

async function loadAll() {
  const [unitRows, photoRows, userRows, settingsRows, clientRows, projectRows, contactRows, containerRows] =
    await Promise.all([
      getAll(UNIT_STORE),
      getAll(PHOTO_STORE),
      getAll(USER_STORE),
      getAll(SETTINGS_STORE),
      getAll(CLIENT_STORE),
      getAll(PROJECT_STORE),
      getAll(CONTACT_STORE),
      getAll(CONTAINER_STORE),
    ]);

  units = unitRows.map(normalizeUnit).sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  photos = photoRows;
  users = userRows.sort((a, b) => a.username.localeCompare(b.username));
  clients = clientRows.map(normalizeClient).sort((a, b) => a.name.localeCompare(b.name));
  projects = projectRows.map(normalizeProject).sort((a, b) => a.name.localeCompare(b.name));
  contacts = contactRows.map(normalizeContact).sort((a, b) => a.name.localeCompare(b.name));
  containers = containerRows.map(normalizeContainer).sort((a, b) => {
    const ad = a.etaDate ? new Date(a.etaDate).getTime() : 0;
    const bd = b.etaDate ? new Date(b.etaDate).getTime() : 0;
    return ad - bd;
  });

  const savedSync = settingsRows.find((row) => row.id === "syncConfig") || {};
  syncConfig = {
    ...DEFAULT_SYNC,
    ...savedSync,
    autoSync: Boolean(savedSync.autoSync),
  };

  photosByUnit = new Map();
  for (const photo of photos) {
    if (!photosByUnit.has(photo.unitId)) photosByUnit.set(photo.unitId, []);
    photosByUnit.get(photo.unitId).push(photo);
  }
}

function clientById(id) {
  return clients.find((client) => client.id === id);
}

function projectById(id) {
  return projects.find((project) => project.id === id);
}

function contactsForProject(projectId) {
  return contacts.filter((contact) => contact.projectId === projectId);
}

function isAdmin() {
  return currentUser?.role === "admin";
}

function can(action, stageKey = "") {
  if (!currentUser) return false;
  if (isAdmin()) return true;

  if (action === "manageUsers") return false;
  if (action === "sync") return false;
  if (action === "manageCatalog") return false;
  if (action === "manageContainers") return ["warehouse", "transport"].includes(currentUser.role);
  if (action === "manageContainerManifest") return currentUser.role === "warehouse";
  if (action === "manageContainerRelease") return currentUser.role === "warehouse";
  if (action === "deleteContainer") return false;
  if (action === "deleteUnit") return false;
  if (action === "createUnit") return currentUser.role === "warehouse";
  if (action === "toggleStage") return STAGE_ROLE_ACCESS[currentUser.role]?.includes(stageKey) || false;
  if (action === "quality") return currentUser.role === "qa";
  if (action === "install") return currentUser.role === "installer";
  if (action === "notes") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "issues") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "photos") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "checklist") return ["warehouse", "transport", "qa"].includes(currentUser.role);
  if (action === "report") return true;

  return false;
}

function rolePermissionsSummary(role) {
  if (role === "admin") return "Acesso total: cadastro base, containers, usuarios, sync, QA, instalacao e relatorios.";
  if (role === "warehouse") return "Cria unidades, controla containers, manifesto e envio/retencao do warehouse.";
  if (role === "transport") return "Atualiza transportation/site job e acompanha schedule de containers.";
  if (role === "qa") return "Valida qualidade, pendencias e pode anexar fotos.";
  if (role === "installer") return "Controla distribuicao, instalacao, pendencias e fotos.";
  return "";
}

function stageDoneCount(unit) {
  return STAGES.filter((stage) => unit.stages[stage.key]?.done).length;
}

function statusBadge(type) {
  if (type === "ok") return "OK";
  if (type === "missing") return "Faltante";
  if (type === "damaged") return "Defeito";
  if (type === "adjustment") return "Ajuste";
  return "-";
}

async function hashPassword(password) {
  const encoded = new TextEncoder().encode(password);
  const buffer = await crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(buffer))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function sessionUserId() {
  return localStorage.getItem(SESSION_KEY);
}

function setSession(userId) {
  localStorage.setItem(SESSION_KEY, userId);
}

function clearSession() {
  localStorage.removeItem(SESSION_KEY);
}

async function tryRestoreSession() {
  const id = sessionUserId();
  if (!id) return;
  const found = users.find((user) => user.id === id);
  if (found) currentUser = found;
}

function renderAuth() {
  const noUsers = users.length === 0;

  if (currentUser) {
    authView.classList.add("hidden");
    appMain.classList.remove("hidden");
    userBadge.classList.remove("hidden");
    logoutBtn.classList.remove("hidden");
    userBadge.textContent = `${currentUser.name} (${ROLE_LABEL[currentUser.role]})`;
    render();
    return;
  }

  appMain.classList.add("hidden");
  authView.classList.remove("hidden");
  userBadge.classList.add("hidden");
  logoutBtn.classList.add("hidden");

  setupPanel.classList.toggle("hidden", !noUsers);
  loginPanel.classList.toggle("hidden", noUsers);
}

function setFormEnabled(form, enabled) {
  form.querySelectorAll("input, select, textarea, button").forEach((field) => {
    field.disabled = !enabled;
  });
}

function renderStats() {
  const total = units.length;
  const completed = units.filter((unit) => unit.installationStatus === "completed").length;
  const pending = units.filter((unit) => stageDoneCount(unit) < STAGES.length).length;
  const blocked = units.filter((unit) => unit.installationStatus === "blocked").length;
  const qualityIssues = units.filter((unit) => unit.deliveryQuality === "rejected").length;

  const arriving7 = containers.filter((container) => {
    if (!container.etaDate || container.arrivalStatus === "arrived") return false;
    const d = daysUntil(container.etaDate);
    return d !== null && d >= 0 && d <= 7;
  }).length;

  const delayed = containers.filter((container) => {
    if (container.arrivalStatus === "delayed") return true;
    if (!container.etaDate || container.arrivalStatus === "arrived") return false;
    const d = daysUntil(container.etaDate);
    return d !== null && d < 0;
  }).length;

  const cards = [
    ["Total unidades", total],
    ["Pendentes", pending],
    ["Instalacao concluida", completed],
    ["Instalacao bloqueada", blocked],
    ["Qualidade rejeitada", qualityIssues],
    ["Containers", containers.length],
    ["Chegam em 7 dias", arriving7],
    ["Containers atrasados", delayed],
  ];

  statsPanel.innerHTML = cards
    .map(([label, value]) => `<div class="stat-item"><span>${label}</span><strong>${value}</strong></div>`)
    .join("");
}

function populateSelect(
  select,
  list,
  {
    placeholder = "Selecione",
    includeEmpty = true,
    labelFn = (item) => item.name,
  } = {}
) {
  const previous = select.value;
  const options = [];
  if (includeEmpty) options.push(`<option value="">${escapeHtml(placeholder)}</option>`);
  for (const item of list) options.push(`<option value="${escapeHtml(item.id)}">${escapeHtml(labelFn(item))}</option>`);

  select.innerHTML = options.join("");
  if (list.some((item) => item.id === previous)) select.value = previous;
}

function projectsForClient(clientId) {
  if (!clientId) return projects;
  return projects.filter((project) => project.clientId === clientId);
}

function populateProjectSelectByClient(select, clientId, placeholder, includeEmpty = true) {
  populateSelect(select, projectsForClient(clientId), {
    includeEmpty,
    placeholder,
    labelFn: (project) => `${project.name}${project.code ? ` (${project.code})` : ""}`,
  });
}

function syncUnitProjectInfo() {
  const projectId = unitProjectSelect.value;
  if (!projectId) {
    unitClientName.value = "";
    unitProjectName.value = "";
    unitJobSite.value = "";
    return;
  }

  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;
  unitProjectName.value = project?.name || "";
  unitClientName.value = client?.name || "";
  unitJobSite.value = project?.address || "";
}

function syncContainerProjectSelect() {
  const clientId = containerClientSelect.value;
  populateProjectSelectByClient(
    containerProjectSelect,
    clientId,
    clientId ? "Selecione o projeto" : "Selecione um cliente",
    true
  );
}

function syncContactProjectSelect() {
  const clientId = contactClientSelect.value;
  populateProjectSelectByClient(contactProjectSelect, clientId, "Projeto (opcional)", true);
}

function renderMasterSelects() {
  populateSelect(projectClientSelect, clients, {
    includeEmpty: true,
    placeholder: clients.length ? "Selecione o cliente" : "Cadastre um cliente",
  });

  populateSelect(contactClientSelect, clients, {
    includeEmpty: true,
    placeholder: "Cliente (opcional)",
  });

  syncContactProjectSelect();

  populateSelect(unitProjectSelect, projects, {
    includeEmpty: true,
    placeholder: projects.length ? "Selecione o projeto" : "Cadastre cliente e projeto primeiro",
    labelFn: (project) => {
      const client = clientById(project.clientId);
      return `${project.name} - ${client?.name || "Sem cliente"}`;
    },
  });

  populateSelect(containerClientSelect, clients, {
    includeEmpty: true,
    placeholder: clients.length ? "Selecione o cliente" : "Cadastre cliente primeiro",
  });

  syncContainerProjectSelect();

  const previousReport = projectReportSelect.value;
  projectReportSelect.innerHTML = `<option value="all">Projeto para PDF (todos)</option>${projects
    .map((project) => `<option value="${escapeHtml(project.id)}">${escapeHtml(project.name)}</option>`)
    .join("")}`;
  if (projects.some((project) => project.id === previousReport)) projectReportSelect.value = previousReport;

  syncUnitProjectInfo();
}

function renderClientsTable() {
  const canManage = can("manageCatalog");

  const rows = clients
    .map((client) => {
      const linkedProjects = projects.filter((project) => project.clientId === client.id).length;
      return `<tr>
        <td>${escapeHtml(client.name)}</td>
        <td>${escapeHtml(client.contactPerson)}</td>
        <td>${escapeHtml(client.phone)}</td>
        <td>${escapeHtml(client.email)}</td>
        <td>${linkedProjects}</td>
        <td>${canManage ? `<button class="danger xs-btn" data-del-client="${client.id}">Excluir</button>` : "-"}</td>
      </tr>`;
    })
    .join("");

  clientsTable.innerHTML = `<table class="data-table"><thead><tr><th>Cliente</th><th>Contato</th><th>Telefone</th><th>Email</th><th>Projetos</th><th>Acoes</th></tr></thead><tbody>${
    rows || '<tr><td colspan="6">Nenhum cliente cadastrado.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;

  clientsTable.querySelectorAll("[data-del-client]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.delClient;
      if (projects.some((project) => project.clientId === id)) {
        alert("Nao e possivel excluir cliente com projetos vinculados.");
        return;
      }
      if (contacts.some((contact) => contact.clientId === id)) {
        alert("Nao e possivel excluir cliente com pessoas vinculadas.");
        return;
      }
      if (containers.some((container) => container.clientId === id)) {
        alert("Nao e possivel excluir cliente com containers vinculados.");
        return;
      }
      await del(CLIENT_STORE, id);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderProjectsTable() {
  const canManage = can("manageCatalog");

  const rows = projects
    .map((project) => {
      const client = clientById(project.clientId);
      return `<tr>
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project.code)}</td>
        <td>${escapeHtml(project.address)}</td>
        <td>${canManage ? `<button class="danger xs-btn" data-del-project="${project.id}">Excluir</button>` : "-"}</td>
      </tr>`;
    })
    .join("");

  projectsTable.innerHTML = `<table class="data-table"><thead><tr><th>Projeto</th><th>Cliente</th><th>Codigo</th><th>Endereco</th><th>Acoes</th></tr></thead><tbody>${
    rows || '<tr><td colspan="5">Nenhum projeto cadastrado.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;

  projectsTable.querySelectorAll("[data-del-project]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.delProject;
      if (units.some((unit) => unit.projectId === id)) {
        alert("Nao e possivel excluir projeto com unidades vinculadas.");
        return;
      }
      if (contacts.some((contact) => contact.projectId === id)) {
        alert("Nao e possivel excluir projeto com pessoas vinculadas.");
        return;
      }
      if (containers.some((container) => container.projectId === id)) {
        alert("Nao e possivel excluir projeto com containers vinculados.");
        return;
      }
      await del(PROJECT_STORE, id);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderContactsTable() {
  const canManage = can("manageCatalog");

  const rows = contacts
    .map((contact) => {
      const client = clientById(contact.clientId);
      const project = projectById(contact.projectId);
      return `<tr>
        <td>${escapeHtml(contact.name)}</td>
        <td>${escapeHtml(CONTACT_ROLE_LABEL[contact.role] || contact.role)}</td>
        <td>${escapeHtml(contact.company)}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project?.name || "-")}</td>
        <td>${escapeHtml(contact.phone)}</td>
        <td>${escapeHtml(contact.email)}</td>
        <td>${canManage ? `<button class="danger xs-btn" data-del-contact="${contact.id}">Excluir</button>` : "-"}</td>
      </tr>`;
    })
    .join("");

  contactsTable.innerHTML = `<table class="data-table"><thead><tr><th>Nome</th><th>Funcao</th><th>Empresa</th><th>Cliente</th><th>Projeto</th><th>Telefone</th><th>Email</th><th>Acoes</th></tr></thead><tbody>${
    rows || '<tr><td colspan="8">Nenhuma pessoa cadastrada.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;
  contactsTable.querySelectorAll("[data-del-contact]").forEach((button) => {
    button.addEventListener("click", async () => {
      await del(CONTACT_STORE, button.dataset.delContact);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderMasterData() {
  masterDataPanel.classList.toggle("hidden", !currentUser);

  const canManage = can("manageCatalog");
  setFormEnabled(clientForm, canManage);
  setFormEnabled(projectForm, canManage);
  setFormEnabled(contactForm, canManage);

  renderMasterSelects();
  renderClientsTable();
  renderProjectsTable();
  renderContactsTable();
}

function containerReleasedQty(container) {
  const totals = { kitchen: 0, vanity: 0, medCabinet: 0, countertop: 0 };
  for (const move of container.movements) {
    if (move.action !== "release") continue;
    totals.kitchen += toNumber(move.kitchens);
    totals.vanity += toNumber(move.vanities);
    totals.medCabinet += toNumber(move.medCabinets);
    totals.countertop += toNumber(move.countertops);
  }
  return totals;
}

function containerAvailableQty(container) {
  const released = containerReleasedQty(container);
  return {
    kitchen: container.qtyKitchens - released.kitchen,
    vanity: container.qtyVanities - released.vanity,
    medCabinet: container.qtyMedCabinets - released.medCabinet,
    countertop: container.qtyCountertops - released.countertop,
  };
}

function containerAlertLine(container) {
  if (!container.etaDate) return "Sem ETA informado.";
  const d = daysUntil(container.etaDate);
  if (container.arrivalStatus === "arrived") return `Container recebido em status Arrived. ETA: ${container.etaDate}.`;
  if (container.arrivalStatus === "delayed" || d < 0) return `Atencao: container atrasado (${Math.abs(d)} dia(s) de atraso).`;
  if (d === 0) return "Container previsto para hoje.";
  if (d <= 7) return `Container chega em ${d} dia(s).`;
  return `ETA em ${d} dia(s).`;
}

async function saveContainer(container) {
  container.updatedAt = new Date().toISOString();
  await put(CONTAINER_STORE, normalizeContainer(container));
  await loadAll();
  render();
  queueAutoSync();
}

function renderManifestTable(wrapper, container, editable) {
  const rows = container.materialItems
    .map(
      (item) => `<tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.description)}</td>
      <td>${escapeHtml(item.category)}</td>
      <td>${escapeHtml(item.qty)}</td>
      <td>${escapeHtml(item.unit)}</td>
      <td>${editable ? `<button class="danger xs-btn" data-manifest-del="${item.id}">x</button>` : "-"}</td>
    </tr>`
    )
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Cod.</th><th>Descricao</th><th>Categoria</th><th>Qtd</th><th>Unidade</th><th>Acoes</th></tr></thead><tbody>${
    rows || '<tr><td colspan="6">Sem itens no manifesto.</td></tr>'
  }</tbody></table>`;
}

function renderReleaseTable(wrapper, container) {
  const rows = container.movements
    .slice()
    .reverse()
    .map((move) => {
      return `<tr>
      <td>${escapeHtml(fmtDate(move.createdAt))}</td>
      <td>${escapeHtml(move.action === "release" ? "Envio" : "Retencao")}</td>
      <td>${escapeHtml(move.unitLabel || "-")}</td>
      <td>${escapeHtml(move.destination || "-")}</td>
      <td>K:${toNumber(move.kitchens)} V:${toNumber(move.vanities)} M:${toNumber(move.medCabinets)} C:${toNumber(move.countertops)}</td>
      <td>${escapeHtml(move.operatorName || "-")}</td>
      <td>${escapeHtml(move.note || "-")}</td>
    </tr>`;
    })
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Data</th><th>Movimento</th><th>Unidade</th><th>Destino</th><th>Qtd</th><th>Operador</th><th>Nota</th></tr></thead><tbody>${
    rows || '<tr><td colspan="7">Sem movimentos registrados.</td></tr>'
  }</tbody></table>`;
}

function renderContainers() {
  containerPanel.classList.toggle("hidden", !currentUser);
  containersBoard.innerHTML = "";

  if (!containers.length) {
    containersBoard.innerHTML = `<div class="panel empty">Nenhum container cadastrado.</div>`;
    return;
  }

  for (const container of containers) {
    const node = containerTemplate.content.firstElementChild.cloneNode(true);
    const client = clientById(container.clientId);
    const project = projectById(container.projectId);
    const available = containerAvailableQty(container);
    const released = containerReleasedQty(container);

    node.querySelector(".container-title").textContent = `${container.containerCode} - ${CONTAINER_STATUS_LABEL[container.arrivalStatus] || container.arrivalStatus}`;
    node.querySelector(".container-meta").textContent = [
      client?.name && `Cliente: ${client.name}`,
      project?.name && `Projeto: ${project.name}`,
      container.manufacturer && `Fabricante: ${container.manufacturer}`,
      container.etaDate && `ETA: ${container.etaDate}`,
    ]
      .filter(Boolean)
      .join(" | ");

    node.querySelector(".container-kpi").innerHTML = `
      <div class="kpi-box"><span>Planejado (K/V/M/C)</span><strong>${container.qtyKitchens}/${container.qtyVanities}/${container.qtyMedCabinets}/${container.qtyCountertops}</strong></div>
      <div class="kpi-box"><span>Enviado (K/V/M/C)</span><strong>${released.kitchen}/${released.vanity}/${released.medCabinet}/${released.countertop}</strong></div>
      <div class="kpi-box"><span>Saldo (K/V/M/C)</span><strong>${available.kitchen}/${available.vanity}/${available.medCabinet}/${available.countertop}</strong></div>
    `;

    node.querySelector(".container-alert").textContent = containerAlertLine(container);

    const deleteBtn = node.querySelector(".container-delete-btn");
    deleteBtn.classList.toggle("hidden", !can("deleteContainer"));
    deleteBtn.addEventListener("click", async () => {
      if (!can("deleteContainer")) return;
      await del(CONTAINER_STORE, container.id);
      await loadAll();
      render();
      queueAutoSync();
    });

    const manifestEditable = can("manageContainerManifest");
    const manifestForm = node.querySelector(".manifest-form");
    if (!manifestEditable) manifestForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    const manifestTable = node.querySelector(".manifest-table");
    renderManifestTable(manifestTable, container, manifestEditable);

    manifestForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!manifestEditable) return;

      const code = manifestForm.querySelector(".mf-code").value.trim();
      const description = manifestForm.querySelector(".mf-desc").value.trim();
      const category = manifestForm.querySelector(".mf-category").value;
      const qty = toNumber(manifestForm.querySelector(".mf-qty").value);
      const unit = manifestForm.querySelector(".mf-unit").value.trim();

      if (!description || qty <= 0) {
        alert("Descricao e quantidade sao obrigatorias.");
        return;
      }

      container.materialItems.push({
        id: uid(),
        code,
        description,
        category,
        qty,
        unit,
        createdAt: new Date().toISOString(),
      });

      manifestForm.reset();
      await saveContainer(container);
    });

    if (manifestEditable) {
      manifestTable.querySelectorAll("[data-manifest-del]").forEach((btn) => {
        btn.addEventListener("click", async () => {
          container.materialItems = container.materialItems.filter((item) => item.id !== btn.dataset.manifestDel);
          await saveContainer(container);
        });
      });
    }

    const releaseEditable = can("manageContainerRelease");
    const releaseForm = node.querySelector(".release-form");
    const releaseUnitSelect = releaseForm.querySelector(".rf-unit");
    const releaseDestination = releaseForm.querySelector(".rf-destination");

    const unitOptions = units.filter((u) => u.projectId === container.projectId || u.projectName === project?.name);
    releaseUnitSelect.innerHTML = `<option value="">Unidade destino (opcional)</option>${unitOptions
      .map((u) => `<option value="${escapeHtml(u.id)}">${escapeHtml(u.unitCode)} - ${escapeHtml(u.jobSite || u.projectName)}</option>`)
      .join("")}`;

    releaseUnitSelect.addEventListener("change", () => {
      const selected = unitOptions.find((u) => u.id === releaseUnitSelect.value);
      if (selected) releaseDestination.value = selected.jobSite || selected.unitCode;
    });

    if (!releaseEditable) releaseForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    const releaseTable = node.querySelector(".release-table");
    renderReleaseTable(releaseTable, container);

    releaseForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!releaseEditable) return;

      const unitId = releaseUnitSelect.value;
      const selectedUnit = units.find((u) => u.id === unitId);
      const action = releaseForm.querySelector(".rf-action").value;
      const destination = releaseDestination.value.trim();
      const kitchens = toNumber(releaseForm.querySelector(".rf-k").value);
      const vanities = toNumber(releaseForm.querySelector(".rf-v").value);
      const medCabinets = toNumber(releaseForm.querySelector(".rf-m").value);
      const countertops = toNumber(releaseForm.querySelector(".rf-c").value);
      const note = releaseForm.querySelector(".rf-note").value.trim();

      const totalMoved = kitchens + vanities + medCabinets + countertops;
      if (totalMoved <= 0) {
        alert("Informe ao menos uma quantidade para registrar movimento.");
        return;
      }

      if (action === "release") {
        const balance = containerAvailableQty(container);
        if (
          kitchens > balance.kitchen ||
          vanities > balance.vanity ||
          medCabinets > balance.medCabinet ||
          countertops > balance.countertop
        ) {
          alert("Quantidade excede o saldo disponivel no container.");
          return;
        }
      }

      const now = new Date().toISOString();
      container.movements.push({
        id: uid(),
        unitId: unitId || "",
        unitLabel: selectedUnit ? selectedUnit.unitCode : "",
        action,
        destination,
        kitchens,
        vanities,
        medCabinets,
        countertops,
        note,
        operatorUserId: currentUser.id,
        operatorName: currentUser.name,
        createdAt: now,
      });

      container.updatedAt = now;
      await put(CONTAINER_STORE, normalizeContainer(container));

      if (action === "release" && selectedUnit) {
        selectedUnit.stages.warehouse = { done: true, at: now };
        const noteLine = `Warehouse release ${container.containerCode}: K${kitchens} V${vanities} M${medCabinets} C${countertops} (${fmtDate(now)})`;
        selectedUnit.deliveryNotes = selectedUnit.deliveryNotes
          ? `${selectedUnit.deliveryNotes}\n${noteLine}`
          : noteLine;
        selectedUnit.updatedAt = now;
        await put(UNIT_STORE, normalizeUnit(selectedUnit));
      }

      releaseForm.reset();
      releaseUnitSelect.value = "";
      releaseDestination.value = "";
      await loadAll();
      render();
      queueAutoSync();
    });

    containersBoard.appendChild(node);
  }
}

function matchesFilter(unit, query, filter) {
  const q = query.trim().toLowerCase();
  const text = [
    unit.clientName,
    unit.projectName,
    unit.jobSite,
    unit.unitCode,
    unit.building,
    unit.floor,
    unit.category,
    unit.unitType,
    unit.kitchenModel,
  ]
    .join(" ")
    .toLowerCase();

  const passQuery = !q || text.includes(q);
  if (!passQuery) return false;

  if (filter === "pending") return stageDoneCount(unit) < STAGES.length;
  if (filter === "qualityIssues") return unit.deliveryQuality === "rejected";
  if (filter === "blockedInstall") return unit.installationStatus === "blocked";
  return true;
}

async function saveSetting(row) {
  await put(SETTINGS_STORE, row);
  await loadAll();
  render();
}

function updateSyncStatus(message, isError = false) {
  syncStatus.textContent = message;
  syncStatus.style.borderColor = isError ? "#e7b6b6" : "#94c5a8";
  syncStatus.style.background = isError ? "#fff5f5" : "#f4fcf7";
}

async function saveUnit(unit) {
  unit.updatedAt = new Date().toISOString();
  await put(UNIT_STORE, normalizeUnit(unit));
  await loadAll();
  render();
  queueAutoSync();
}

async function addPhotos(unitId, files) {
  for (const file of files) {
    const dataUrl = await fileToDataUrl(file);
    await put(PHOTO_STORE, {
      id: uid(),
      unitId,
      name: file.name,
      type: file.type,
      dataUrl,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  }
  await loadAll();
  render();
  queueAutoSync();
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function checklistSummary(checkItems) {
  const total = checkItems.length;
  const missing = checkItems.filter((item) => item.status === "missing").length;
  const damaged = checkItems.filter((item) => item.status === "damaged").length;
  const adjustment = checkItems.filter((item) => item.status === "adjustment").length;
  return `Itens: ${total} | Faltantes: ${missing} | Defeito: ${damaged} | Ajuste: ${adjustment}`;
}

function renderChecklistTable(wrapper, unit) {
  const editable = can("checklist");
  if (!unit.checkItems.length) {
    wrapper.innerHTML = `<p class="hint">Sem itens cadastrados.</p>`;
    return;
  }

  wrapper.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th>Cod.</th>
          <th>Descricao</th>
          <th>Esperada</th>
          <th>Conferida</th>
          <th>Status</th>
          <th>Obs</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${unit.checkItems
          .map(
            (item) => `
            <tr>
              <td>${escapeHtml(item.code)}</td>
              <td>${escapeHtml(item.description)}</td>
              <td>${escapeHtml(item.expectedQty)}</td>
              <td><input data-item-action="got" data-item-id="${item.id}" type="number" min="0" value="${Number(item.checkedQty || 0)}" ${
                editable ? "" : "disabled"
              } /></td>
              <td>
                <select data-item-action="status" data-item-id="${item.id}" ${editable ? "" : "disabled"}>
                  <option value="ok" ${item.status === "ok" ? "selected" : ""}>OK</option>
                  <option value="missing" ${item.status === "missing" ? "selected" : ""}>Faltante</option>
                  <option value="damaged" ${item.status === "damaged" ? "selected" : ""}>Defeito</option>
                  <option value="adjustment" ${item.status === "adjustment" ? "selected" : ""}>Ajuste</option>
                </select>
              </td>
              <td><input data-item-action="note" data-item-id="${item.id}" value="${escapeHtml(item.note || "")}" ${
                editable ? "" : "disabled"
              } /></td>
              <td><button data-item-remove="${item.id}" class="danger xs-btn" ${editable ? "" : "disabled"}>x</button></td>
            </tr>`
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function bindChecklistEvents(node, unit) {
  const editable = can("checklist");
  const wrapper = node.querySelector(".checklist-table");
  renderChecklistTable(wrapper, unit);
  node.querySelector(".checklist-summary").textContent = checklistSummary(unit.checkItems);

  if (!editable) return;

  wrapper.querySelectorAll("[data-item-action]").forEach((el) => {
    el.addEventListener("change", () => {
      const item = unit.checkItems.find((entry) => entry.id === el.dataset.itemId);
      if (!item) return;
      if (el.dataset.itemAction === "got") item.checkedQty = Number(el.value || 0);
      if (el.dataset.itemAction === "status") item.status = el.value;
      if (el.dataset.itemAction === "note") item.note = el.value;
      item.updatedAt = new Date().toISOString();
      saveUnit(unit);
    });
  });

  wrapper.querySelectorAll("[data-item-remove]").forEach((button) => {
    button.addEventListener("click", () => {
      unit.checkItems = unit.checkItems.filter((entry) => entry.id !== button.dataset.itemRemove);
      saveUnit(unit);
    });
  });
}

function renderUnits() {
  const query = searchInput.value;
  const filter = filterSelect.value;
  const filtered = units.filter((unit) => matchesFilter(unit, query, filter));

  unitsContainer.innerHTML = "";

  if (!filtered.length) {
    unitsContainer.innerHTML = `<div class="panel empty">Nenhum registro encontrado.</div>`;
    return;
  }

  for (const unit of filtered) {
    const node = unitTemplate.content.firstElementChild.cloneNode(true);
    const progress = Math.round((stageDoneCount(unit) / STAGES.length) * 100);

    node.querySelector(".unit-title").textContent = `${unit.unitCode} - ${unit.category}`;
    node.querySelector(".unit-meta").textContent = [
      unit.clientName && `Cliente: ${unit.clientName}`,
      unit.projectName && `Projeto: ${unit.projectName}`,
      unit.jobSite,
      unit.building && `Building: ${unit.building}`,
      unit.floor && `Floor: ${unit.floor}`,
      unit.unitType && `Type: ${unit.unitType}`,
      unit.kitchenModel && `Kitchen: ${unit.kitchenModel}`,
    ]
      .filter(Boolean)
      .join(" | ");

    node.querySelector(".progress-bar").style.width = `${progress}%`;
    node.querySelector(".progress-label").textContent = `${progress}% concluido (${stageDoneCount(unit)}/${STAGES.length})`;

    const stageGrid = node.querySelector(".stage-grid");
    for (const stage of STAGES) {
      const stageValue = unit.stages[stage.key] || { done: false, at: null };
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `stage-btn ${stageValue.done ? "done" : ""}`;
      btn.disabled = !can("toggleStage", stage.key);
      btn.innerHTML = `<strong>${stage.label}</strong><span>${stageValue.done ? "Baixado: " + fmtDate(stageValue.at) : "Pendente"}</span>`;
      btn.addEventListener("click", () => {
        if (!can("toggleStage", stage.key)) return;
        const latest = unit.stages[stage.key];
        unit.stages[stage.key] = {
          done: !latest.done,
          at: !latest.done ? new Date().toISOString() : null,
        };
        saveUnit(unit);
      });
      stageGrid.appendChild(btn);
    }

    const qualitySelect = node.querySelector(".quality-status");
    const installSelect = node.querySelector(".install-status");
    const notesInput = node.querySelector(".delivery-notes");
    const issuesInput = node.querySelector(".issues-text");

    qualitySelect.value = unit.deliveryQuality;
    installSelect.value = unit.installationStatus;
    notesInput.value = unit.deliveryNotes;
    issuesInput.value = unit.issuesText;

    qualitySelect.disabled = !can("quality");
    installSelect.disabled = !can("install");
    notesInput.disabled = !can("notes");
    issuesInput.disabled = !can("issues");

    qualitySelect.addEventListener("change", () => {
      if (!can("quality")) return;
      unit.deliveryQuality = qualitySelect.value;
      saveUnit(unit);
    });

    installSelect.addEventListener("change", () => {
      if (!can("install")) return;
      unit.installationStatus = installSelect.value;
      saveUnit(unit);
    });

    notesInput.addEventListener("blur", () => {
      if (!can("notes")) return;
      unit.deliveryNotes = notesInput.value.trim();
      saveUnit(unit);
    });

    issuesInput.addEventListener("blur", () => {
      if (!can("issues")) return;
      unit.issuesText = issuesInput.value.trim();
      saveUnit(unit);
    });

    const checkForm = node.querySelector(".check-item-form");
    if (!can("checklist")) checkForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    checkForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!can("checklist")) return;

      const code = checkForm.querySelector(".ci-code").value.trim();
      const description = checkForm.querySelector(".ci-desc").value.trim();
      const expectedQty = Number(checkForm.querySelector(".ci-exp").value || 0);
      const checkedQty = Number(checkForm.querySelector(".ci-got").value || 0);
      const status = checkForm.querySelector(".ci-status").value;

      if (!description) return;

      unit.checkItems.push({
        id: uid(),
        code,
        description,
        expectedQty,
        checkedQty,
        status,
        note: "",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });

      checkForm.reset();
      saveUnit(unit);
    });

    bindChecklistEvents(node, unit);

    node.querySelector(".scope-line").textContent = unit.scopeWork ? `Scope: ${unit.scopeWork}` : "Scope: nao informado";
    node.querySelector(".shop-line").textContent = unit.shopdrawingRef
      ? `Shopdrawing: ${unit.shopdrawingRef}`
      : "Shopdrawing: nao informado";

    const photoInput = node.querySelector(".photo-input");
    const photoLabel = node.querySelector(".photo-label");
    photoLabel.classList.toggle("hidden", !can("photos"));

    photoInput.addEventListener("change", async () => {
      if (!can("photos")) return;
      if (!photoInput.files?.length) return;
      await addPhotos(unit.id, photoInput.files);
      photoInput.value = "";
    });

    const gallery = node.querySelector(".photo-gallery");
    const unitPhotos = photosByUnit.get(unit.id) || [];
    for (const photo of unitPhotos) {
      const p = document.createElement("div");
      p.className = "photo";
      p.innerHTML = `<img loading="lazy" src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" /><button type="button" title="Remover foto" ${
        can("photos") ? "" : "disabled"
      }>x</button>`;
      p.querySelector("button").addEventListener("click", async () => {
        if (!can("photos")) return;
        await del(PHOTO_STORE, photo.id);
        await loadAll();
        render();
        queueAutoSync();
      });
      gallery.appendChild(p);
    }

    const deleteBtn = node.querySelector(".delete-btn");
    deleteBtn.classList.toggle("hidden", !can("deleteUnit"));
    deleteBtn.addEventListener("click", async () => {
      if (!can("deleteUnit")) return;
      await del(UNIT_STORE, unit.id);
      const toDelete = photosByUnit.get(unit.id) || [];
      await Promise.all(toDelete.map((photo) => del(PHOTO_STORE, photo.id)));
      await loadAll();
      render();
      queueAutoSync();
    });

    node.querySelector(".report-btn").addEventListener("click", () => {
      generateUnitReport(unit.id);
    });

    unitsContainer.appendChild(node);
  }
}

function renderUsers() {
  if (!can("manageUsers")) {
    usersPanel.classList.add("hidden");
    return;
  }

  usersPanel.classList.remove("hidden");

  const rows = users
    .map(
      (user) => `<tr>
      <td>${escapeHtml(user.name)}</td>
      <td>${escapeHtml(user.username)}</td>
      <td>${escapeHtml(ROLE_LABEL[user.role])}</td>
      <td>${fmtDate(user.updatedAt)}</td>
    </tr>`
    )
    .join("");

  usersTable.innerHTML = `<table class="data-table"><thead><tr><th>Nome</th><th>Usuario</th><th>Perfil</th><th>Atualizado</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function renderRoleStrip() {
  roleLine.textContent = `Usuario: ${currentUser.name} | Perfil: ${ROLE_LABEL[currentUser.role]}`;
  permissionLine.textContent = rolePermissionsSummary(currentUser.role);
}

function renderSyncPanel() {
  const allowed = can("sync");
  syncPanel.classList.toggle("hidden", !allowed);
  if (!allowed) return;

  syncConfigForm.supabaseUrl.value = syncConfig.supabaseUrl;
  syncConfigForm.supabaseAnonKey.value = syncConfig.supabaseAnonKey;
  syncConfigForm.tenant.value = syncConfig.tenant;
  syncConfigForm.autoSync.value = String(syncConfig.autoSync);

  const statusText = syncConfig.lastSyncAt ? `Ultimo sync: ${fmtDate(syncConfig.lastSyncAt)}` : "Sem sincronizacao";
  updateSyncStatus(statusText, false);
}

function renderAccessControl() {
  unitEntryPanel.classList.toggle("hidden", !can("createUnit"));
  exportBtn.disabled = !isAdmin();
  importInput.disabled = !isAdmin();
  projectReportBtn.disabled = !can("report");

  if (can("createUnit")) {
    const hasProjects = projects.length > 0;
    setFormEnabled(unitForm, hasProjects);
    unitProjectHint.textContent = hasProjects
      ? ""
      : "Para cadastrar unidades, primeiro crie cliente e projeto no bloco de cadastro base.";
  }

  const canManageCont = can("manageContainers");
  const hasContBase = clients.length > 0 && projects.length > 0;
  setFormEnabled(containerForm, canManageCont && hasContBase);
}

function render() {
  if (!currentUser) return;
  renderRoleStrip();
  renderMasterData();
  renderContainers();
  renderStats();
  renderSyncPanel();
  renderUsers();
  renderAccessControl();
  renderUnits();
}

function toUnit(formData) {
  const projectId = formData.get("projectId")?.toString();
  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;

  return normalizeUnit({
    id: uid(),
    clientId: client?.id || "",
    projectId: project?.id || "",
    clientName: client?.name || formData.get("clientName")?.toString().trim(),
    projectName: project?.name || formData.get("projectName")?.toString().trim(),
    jobSite: formData.get("jobSite")?.toString().trim(),
    unitCode: formData.get("unitCode")?.toString().trim(),
    building: formData.get("building")?.toString().trim(),
    floor: formData.get("floor")?.toString().trim(),
    category: formData.get("category")?.toString().trim(),
    unitType: formData.get("unitType")?.toString().trim(),
    kitchenModel: formData.get("kitchenModel")?.toString().trim(),
    shopdrawingRef: formData.get("shopdrawingRef")?.toString().trim(),
    scopeWork: formData.get("scopeWork")?.toString().trim(),
    deliveryQuality: "pending",
    installationStatus: "not-started",
    deliveryNotes: "",
    issuesText: "",
    checkItems: [],
    stages: Object.fromEntries(STAGES.map((stage) => [stage.key, { done: false, at: null }])),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
}

function syncEndpoint() {
  const url = syncConfig.supabaseUrl.trim().replace(/\/$/, "");
  if (!url || !syncConfig.supabaseAnonKey || !syncConfig.tenant) return null;
  return `${url}/rest/v1/app_records`;
}

function toCloudRecords() {
  const tenant = syncConfig.tenant.trim();

  const mapRecords = (arr, kind, updatedField = "updatedAt") =>
    arr.map((item) => ({
      tenant,
      kind,
      id: item.id,
      payload: item,
      updated_at: item[updatedField] || item.createdAt || new Date().toISOString(),
    }));

  return [
    ...mapRecords(units, "unit"),
    ...mapRecords(photos, "photo", "updatedAt"),
    ...mapRecords(users, "user"),
    ...mapRecords(clients, "client"),
    ...mapRecords(projects, "project"),
    ...mapRecords(contacts, "contact"),
    ...mapRecords(containers, "container"),
  ];
}

async function pushCloud({ silent = false } = {}) {
  if (!can("sync")) return;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    updateSyncStatus("Configurar Supabase URL/Key/Tenant", true);
    return;
  }
  if (!navigator.onLine) {
    updateSyncStatus("Sem internet para sincronizar", true);
    return;
  }

  try {
    updateSyncStatus("Enviando dados...");

    const response = await fetch(`${endpoint}?on_conflict=tenant,kind,id`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(toCloudRecords()),
    });

    if (!response.ok) throw new Error(await response.text());

    syncConfig.lastSyncAt = new Date().toISOString();
    await saveSetting({ id: "syncConfig", ...syncConfig });
    updateSyncStatus(`Push concluido: ${fmtDate(syncConfig.lastSyncAt)}`);
  } catch (error) {
    updateSyncStatus("Falha no push. Verifique a configuracao.", true);
    if (!silent) console.error(error);
  }
}

function newerThan(remoteTs, localTs) {
  return new Date(remoteTs || 0).getTime() > new Date(localTs || 0).getTime();
}

async function pullCloud() {
  if (!can("sync")) return;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    updateSyncStatus("Configurar Supabase URL/Key/Tenant", true);
    return;
  }
  if (!navigator.onLine) {
    updateSyncStatus("Sem internet para sincronizar", true);
    return;
  }

  try {
    updateSyncStatus("Baixando dados...");
    const qs = new URLSearchParams({
      select: "kind,id,payload,updated_at",
      tenant: `eq.${syncConfig.tenant.trim()}`,
      order: "updated_at.desc",
      limit: "20000",
    });

    const response = await fetch(`${endpoint}?${qs.toString()}`, {
      headers: {
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
      },
    });

    if (!response.ok) throw new Error(await response.text());

    const records = await response.json();
    const unitMap = new Map(units.map((item) => [item.id, item]));
    const photoMap = new Map(photos.map((item) => [item.id, item]));
    const userMap = new Map(users.map((item) => [item.id, item]));
    const clientMap = new Map(clients.map((item) => [item.id, item]));
    const projectMap = new Map(projects.map((item) => [item.id, item]));
    const contactMap = new Map(contacts.map((item) => [item.id, item]));
    const containerMap = new Map(containers.map((item) => [item.id, item]));

    for (const row of records) {
      const item = row.payload || {};
      const remoteUpdatedAt = item.updatedAt || row.updated_at || new Date().toISOString();
      item.updatedAt = remoteUpdatedAt;

      if (row.kind === "unit") {
        const local = unitMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(UNIT_STORE, normalizeUnit(item));
      }

      if (row.kind === "photo") {
        const local = photoMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt || local.createdAt)) await put(PHOTO_STORE, item);
      }

      if (row.kind === "user") {
        const local = userMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(USER_STORE, item);
      }

      if (row.kind === "client") {
        const local = clientMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CLIENT_STORE, normalizeClient(item));
      }

      if (row.kind === "project") {
        const local = projectMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(PROJECT_STORE, normalizeProject(item));
      }

      if (row.kind === "contact") {
        const local = contactMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CONTACT_STORE, normalizeContact(item));
      }

      if (row.kind === "container") {
        const local = containerMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CONTAINER_STORE, normalizeContainer(item));
      }
    }

    syncConfig.lastSyncAt = new Date().toISOString();
    await saveSetting({ id: "syncConfig", ...syncConfig });
    await loadAll();

    if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;

    render();
    updateSyncStatus(`Pull concluido: ${fmtDate(syncConfig.lastSyncAt)}`);
  } catch (error) {
    updateSyncStatus("Falha no pull. Verifique a configuracao.", true);
    console.error(error);
  }
}

function queueAutoSync() {
  if (!can("sync")) return;
  if (!syncConfig.autoSync) return;
  if (!navigator.onLine) return;

  clearTimeout(autoSyncTimer);
  autoSyncTimer = setTimeout(() => {
    pushCloud({ silent: true });
  }, 1200);
}

function reportShell(title, bodyHtml) {
  return `
  <!doctype html>
  <html lang="pt-BR">
    <head>
      <meta charset="UTF-8" />
      <title>${escapeHtml(title)}</title>
      <style>
        body { font-family: Georgia, "Times New Roman", serif; margin: 24px; color: #222; }
        h1 { margin: 0 0 6px; }
        p { margin: 2px 0; }
        .meta { margin-bottom: 14px; color: #555; }
        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { border: 1px solid #bbb; padding: 6px; font-size: 12px; text-align: left; vertical-align: top; }
        th { background: #efefef; }
        .photos { margin-top: 16px; display: flex; flex-wrap: wrap; gap: 8px; }
        .photos img { width: 150px; height: 120px; object-fit: cover; border: 1px solid #aaa; }
      </style>
    </head>
    <body>${bodyHtml}</body>
  </html>`;
}

function warehouseMovesForUnit(unitId) {
  const out = [];
  for (const container of containers) {
    for (const move of container.movements) {
      if (move.unitId === unitId) {
        out.push({ ...move, containerCode: container.containerCode });
      }
    }
  }
  return out.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
}

function generateUnitReport(unitId) {
  const unit = units.find((entry) => entry.id === unitId);
  if (!unit) return;

  const projectPeople = contactsForProject(unit.projectId);
  const photosRows = photosByUnit.get(unit.id) || [];
  const warehouseMoves = warehouseMovesForUnit(unit.id);

  const stagesTable = STAGES.map((stage) => {
    const entry = unit.stages[stage.key] || { done: false, at: null };
    return `<tr><td>${escapeHtml(stage.label)}</td><td>${entry.done ? "Concluido" : "Pendente"}</td><td>${escapeHtml(
      fmtDate(entry.at)
    )}</td></tr>`;
  }).join("");

  const checklistRows = unit.checkItems
    .map(
      (item) => `<tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.description)}</td>
      <td>${escapeHtml(item.expectedQty)}</td>
      <td>${escapeHtml(item.checkedQty)}</td>
      <td>${escapeHtml(statusBadge(item.status))}</td>
      <td>${escapeHtml(item.note || "")}</td>
    </tr>`
    )
    .join("");

  const peopleRows = projectPeople
    .map(
      (person) => `<tr>
      <td>${escapeHtml(person.name)}</td>
      <td>${escapeHtml(CONTACT_ROLE_LABEL[person.role] || person.role)}</td>
      <td>${escapeHtml(person.company)}</td>
      <td>${escapeHtml(person.phone)}</td>
      <td>${escapeHtml(person.email)}</td>
    </tr>`
    )
    .join("");

  const warehouseRows = warehouseMoves
    .map(
      (move) => `<tr>
      <td>${escapeHtml(fmtDate(move.createdAt))}</td>
      <td>${escapeHtml(move.containerCode)}</td>
      <td>${escapeHtml(move.action === "release" ? "Envio" : "Retencao")}</td>
      <td>K:${toNumber(move.kitchens)} V:${toNumber(move.vanities)} M:${toNumber(move.medCabinets)} C:${toNumber(move.countertops)}</td>
      <td>${escapeHtml(move.destination || "-")}</td>
      <td>${escapeHtml(move.operatorName || "-")}</td>
    </tr>`
    )
    .join("");

  const html = reportShell(
    `Relatorio unidade ${unit.unitCode}`,
    `
      <h1>Relatorio de unidade ${escapeHtml(unit.unitCode)}</h1>
      <div class="meta">
        <p><strong>Cliente:</strong> ${escapeHtml(unit.clientName)}</p>
        <p><strong>Projeto:</strong> ${escapeHtml(unit.projectName)}</p>
        <p><strong>Job Site:</strong> ${escapeHtml(unit.jobSite)}</p>
        <p><strong>Categoria:</strong> ${escapeHtml(unit.category)}</p>
        <p><strong>Type:</strong> ${escapeHtml(unit.unitType)}</p>
        <p><strong>Modelo cozinha:</strong> ${escapeHtml(unit.kitchenModel)}</p>
        <p><strong>Shopdrawing:</strong> ${escapeHtml(unit.shopdrawingRef || "-")}</p>
        <p><strong>Scope:</strong> ${escapeHtml(unit.scopeWork || "-")}</p>
        <p><strong>Qualidade:</strong> ${escapeHtml(unit.deliveryQuality)}</p>
        <p><strong>Instalacao:</strong> ${escapeHtml(unit.installationStatus)}</p>
        <p><strong>Pendencias:</strong> ${escapeHtml(unit.issuesText || "-")}</p>
        <p><strong>Notas:</strong> ${escapeHtml(unit.deliveryNotes || "-")}</p>
        <p><strong>Emissao:</strong> ${escapeHtml(fmtDate(new Date().toISOString()))}</p>
      </div>

      <h3>Fluxo operacional</h3>
      <table>
        <thead><tr><th>Etapa</th><th>Status</th><th>Data/Hora</th></tr></thead>
        <tbody>${stagesTable}</tbody>
      </table>

      <h3>Checklist detalhado</h3>
      <table>
        <thead><tr><th>Cod.</th><th>Descricao</th><th>Qtd Esperada</th><th>Qtd Conferida</th><th>Status</th><th>Obs</th></tr></thead>
        <tbody>${checklistRows || '<tr><td colspan="6">Sem itens.</td></tr>'}</tbody>
      </table>

      <h3>Movimentos de warehouse (containers)</h3>
      <table>
        <thead><tr><th>Data</th><th>Container</th><th>Movimento</th><th>Qtd</th><th>Destino</th><th>Operador</th></tr></thead>
        <tbody>${warehouseRows || '<tr><td colspan="6">Sem movimentos vinculados.</td></tr>'}</tbody>
      </table>

      <h3>Pessoas do projeto</h3>
      <table>
        <thead><tr><th>Nome</th><th>Funcao</th><th>Empresa</th><th>Telefone</th><th>Email</th></tr></thead>
        <tbody>${peopleRows || '<tr><td colspan="5">Sem pessoas vinculadas.</td></tr>'}</tbody>
      </table>

      <h3>Fotos</h3>
      <div class="photos">${photosRows.map((photo) => `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" />`).join("") || "Sem fotos"}</div>
    `
  );

  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 300);
}

function generateProjectReport(projectId) {
  const project = projectById(projectId);
  const scopedUnits =
    projectId === "all"
      ? units
      : units.filter((unit) => unit.projectId === projectId || (project && unit.projectName === project.name));

  if (!scopedUnits.length) {
    alert("Nao ha unidades para gerar o relatorio.");
    return;
  }

  const scopedContainers =
    projectId === "all" ? containers : containers.filter((container) => container.projectId === projectId);

  const unitRows = scopedUnits
    .map(
      (unit) => `<tr>
      <td>${escapeHtml(unit.clientName)}</td>
      <td>${escapeHtml(unit.projectName)}</td>
      <td>${escapeHtml(unit.unitCode)}</td>
      <td>${escapeHtml(unit.category)}</td>
      <td>${escapeHtml(unit.deliveryQuality)}</td>
      <td>${escapeHtml(unit.installationStatus)}</td>
      <td>${stageDoneCount(unit)}/${STAGES.length}</td>
      <td>${escapeHtml(checklistSummary(unit.checkItems))}</td>
      <td>${escapeHtml(unit.issuesText || "-")}</td>
    </tr>`
    )
    .join("");

  const containerRows = scopedContainers
    .map((container) => {
      const client = clientById(container.clientId);
      const projectEntry = projectById(container.projectId);
      const av = containerAvailableQty(container);
      return `<tr>
      <td>${escapeHtml(container.containerCode)}</td>
      <td>${escapeHtml(client?.name || "-")}</td>
      <td>${escapeHtml(projectEntry?.name || "-")}</td>
      <td>${escapeHtml(CONTAINER_STATUS_LABEL[container.arrivalStatus] || container.arrivalStatus)}</td>
      <td>${escapeHtml(container.etaDate || "-")}</td>
      <td>${container.qtyKitchens}/${container.qtyVanities}/${container.qtyMedCabinets}/${container.qtyCountertops}</td>
      <td>${av.kitchen}/${av.vanity}/${av.medCabinet}/${av.countertop}</td>
    </tr>`;
    })
    .join("");

  const title = projectId === "all" ? "Todos os projetos" : project?.name || "Projeto";

  const html = reportShell(
    `Relatorio projeto ${title}`,
    `
      <h1>Relatorio de projeto: ${escapeHtml(title)}</h1>
      <div class="meta">
        <p><strong>Total de unidades:</strong> ${scopedUnits.length}</p>
        <p><strong>Total de containers:</strong> ${scopedContainers.length}</p>
        <p><strong>Gerado em:</strong> ${escapeHtml(fmtDate(new Date().toISOString()))}</p>
      </div>
      <h3>Unidades</h3>
      <table>
        <thead>
          <tr>
            <th>Cliente</th>
            <th>Projeto</th>
            <th>Unidade</th>
            <th>Categoria</th>
            <th>Qualidade</th>
            <th>Instalacao</th>
            <th>Progresso</th>
            <th>Checklist</th>
            <th>Pendencias</th>
          </tr>
        </thead>
        <tbody>${unitRows}</tbody>
      </table>
      <h3>Schedule e saldo dos containers</h3>
      <table>
        <thead>
          <tr>
            <th>Container</th>
            <th>Cliente</th>
            <th>Projeto</th>
            <th>Status</th>
            <th>ETA</th>
            <th>Qtd planejada K/V/M/C</th>
            <th>Saldo K/V/M/C</th>
          </tr>
        </thead>
        <tbody>${containerRows || '<tr><td colspan="7">Sem containers.</td></tr>'}</tbody>
      </table>
    `
  );

  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 300);
}

setupForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(setupForm);

  const user = {
    id: uid(),
    name: data.get("name")?.toString().trim(),
    username: data.get("username")?.toString().trim().toLowerCase(),
    passwordHash: await hashPassword(data.get("password")?.toString() || ""),
    role: "admin",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  await put(USER_STORE, user);
  await loadAll();
  currentUser = user;
  setSession(user.id);
  setupForm.reset();
  renderAuth();
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(loginForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  const passwordHash = await hashPassword(data.get("password")?.toString() || "");

  const user = users.find((entry) => entry.username === username && entry.passwordHash === passwordHash);
  if (!user) {
    alert("Usuario ou senha invalidos.");
    return;
  }

  currentUser = user;
  setSession(user.id);
  loginForm.reset();
  renderAuth();
});

logoutBtn.addEventListener("click", () => {
  currentUser = null;
  clearSession();
  renderAuth();
});

userForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageUsers")) return;

  const data = new FormData(userForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  if (users.some((user) => user.username === username)) {
    alert("Usuario ja existe.");
    return;
  }

  const user = {
    id: uid(),
    name: data.get("name")?.toString().trim(),
    username,
    passwordHash: await hashPassword(data.get("password")?.toString() || ""),
    role: data.get("role")?.toString(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  await put(USER_STORE, user);
  userForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

clientForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(clientForm);
  const name = data.get("name")?.toString().trim();
  if (!name) return;

  if (clients.some((client) => client.name.toLowerCase() === name.toLowerCase())) {
    alert("Cliente ja cadastrado.");
    return;
  }

  await put(
    CLIENT_STORE,
    normalizeClient({
      id: uid(),
      name,
      contactPerson: data.get("contactPerson")?.toString().trim(),
      phone: data.get("phone")?.toString().trim(),
      email: data.get("email")?.toString().trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  );

  clientForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

projectForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(projectForm);
  const clientId = data.get("clientId")?.toString();
  const name = data.get("name")?.toString().trim();

  if (!clientId || !name) {
    alert("Cliente e nome do projeto sao obrigatorios.");
    return;
  }

  if (projects.some((project) => project.clientId === clientId && project.name.toLowerCase() === name.toLowerCase())) {
    alert("Projeto ja cadastrado para este cliente.");
    return;
  }

  await put(
    PROJECT_STORE,
    normalizeProject({
      id: uid(),
      clientId,
      name,
      code: data.get("code")?.toString().trim(),
      address: data.get("address")?.toString().trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  );

  projectForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

contactForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(contactForm);
  const name = data.get("name")?.toString().trim();
  if (!name) return;

  await put(
    CONTACT_STORE,
    normalizeContact({
      id: uid(),
      name,
      role: data.get("role")?.toString(),
      company: data.get("company")?.toString().trim(),
      clientId: data.get("clientId")?.toString() || "",
      projectId: data.get("projectId")?.toString() || "",
      phone: data.get("phone")?.toString().trim(),
      email: data.get("email")?.toString().trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  );

  contactForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

containerClientSelect.addEventListener("change", syncContainerProjectSelect);
contactClientSelect.addEventListener("change", syncContactProjectSelect);
unitProjectSelect.addEventListener("change", syncUnitProjectInfo);

containerForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageContainers")) return;

  const data = new FormData(containerForm);
  const containerCode = data.get("containerCode")?.toString().trim();
  const clientId = data.get("clientId")?.toString();
  const projectId = data.get("projectId")?.toString();

  if (!containerCode || !clientId || !projectId) {
    alert("Container, cliente e projeto sao obrigatorios.");
    return;
  }

  if (containers.some((c) => c.containerCode.toLowerCase() === containerCode.toLowerCase())) {
    alert("Container ID ja cadastrado.");
    return;
  }

  await put(
    CONTAINER_STORE,
    normalizeContainer({
      id: uid(),
      containerCode,
      manufacturer: data.get("manufacturer")?.toString().trim(),
      clientId,
      projectId,
      etaDate: data.get("etaDate")?.toString(),
      arrivalStatus: data.get("arrivalStatus")?.toString(),
      qtyKitchens: toNumber(data.get("qtyKitchens")?.toString()),
      qtyVanities: toNumber(data.get("qtyVanities")?.toString()),
      qtyMedCabinets: toNumber(data.get("qtyMedCabinets")?.toString()),
      qtyCountertops: toNumber(data.get("qtyCountertops")?.toString()),
      notes: data.get("notes")?.toString().trim(),
      materialItems: [],
      movements: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  );

  containerForm.reset();
  syncContainerProjectSelect();
  await loadAll();
  render();
  queueAutoSync();
});

unitForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("createUnit")) return;

  const projectId = unitProjectSelect.value;
  if (!projectId) {
    alert("Selecione um projeto.");
    return;
  }

  const unit = toUnit(new FormData(unitForm));
  await put(UNIT_STORE, unit);
  unitForm.reset();
  unitProjectSelect.value = projectId;
  syncUnitProjectInfo();
  await loadAll();
  render();
  queueAutoSync();
});

searchInput.addEventListener("input", renderUnits);
filterSelect.addEventListener("change", renderUnits);

projectReportBtn.addEventListener("click", () => {
  if (!can("report")) return;
  generateProjectReport(projectReportSelect.value);
});

exportBtn.addEventListener("click", async () => {
  if (!isAdmin()) return;

  const payload = {
    exportedAt: new Date().toISOString(),
    units,
    photos,
    users,
    clients,
    projects,
    contacts,
    containers,
    settings: [{ id: "syncConfig", ...syncConfig }],
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `cabinets-control-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

importInput.addEventListener("change", async () => {
  if (!isAdmin()) return;
  const file = importInput.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const payload = JSON.parse(text);

    if (!Array.isArray(payload.units) || !Array.isArray(payload.photos)) {
      throw new Error("Formato invalido");
    }

    for (const unit of payload.units) await put(UNIT_STORE, normalizeUnit(unit));
    for (const photo of payload.photos) await put(PHOTO_STORE, photo);
    if (Array.isArray(payload.users)) for (const user of payload.users) await put(USER_STORE, user);
    if (Array.isArray(payload.clients)) for (const client of payload.clients) await put(CLIENT_STORE, normalizeClient(client));
    if (Array.isArray(payload.projects)) for (const project of payload.projects) await put(PROJECT_STORE, normalizeProject(project));
    if (Array.isArray(payload.contacts)) for (const contact of payload.contacts) await put(CONTACT_STORE, normalizeContact(contact));
    if (Array.isArray(payload.containers)) for (const container of payload.containers) await put(CONTAINER_STORE, normalizeContainer(container));
    if (Array.isArray(payload.settings)) for (const setting of payload.settings) await put(SETTINGS_STORE, setting);

    await loadAll();
    if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;
    render();
  } catch {
    alert("Falha ao importar backup. Verifique o arquivo.");
  } finally {
    importInput.value = "";
  }
});

syncConfigForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("sync")) return;

  const data = new FormData(syncConfigForm);
  syncConfig = {
    ...syncConfig,
    id: "syncConfig",
    supabaseUrl: data.get("supabaseUrl")?.toString().trim(),
    supabaseAnonKey: data.get("supabaseAnonKey")?.toString().trim(),
    tenant: data.get("tenant")?.toString().trim(),
    autoSync: data.get("autoSync")?.toString() === "true",
  };

  await saveSetting(syncConfig);
  updateSyncStatus("Configuracao salva.");
});

pushSyncBtn.addEventListener("click", () => pushCloud());
pullSyncBtn.addEventListener("click", () => pullCloud());

function updateConnectivity() {
  const online = navigator.onLine;
  connectionBadge.textContent = online ? "Online" : "Offline";
  connectionBadge.style.borderColor = online ? "#94c5a8" : "#e7b6b6";
  connectionBadge.style.background = online ? "#f4fcf7" : "#fff5f5";

  if (online) queueAutoSync();
}

window.addEventListener("online", updateConnectivity);
window.addEventListener("offline", updateConnectivity);

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  deferredPrompt = event;
  installBtn.classList.remove("hidden");
});

installBtn.addEventListener("click", async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  installBtn.classList.add("hidden");
});

async function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    try {
      await navigator.serviceWorker.register("sw.js");
    } catch (error) {
      console.error("Falha ao registrar Service Worker", error);
    }
  }
}

(async function init() {
  db = await openDB();
  await loadAll();
  await tryRestoreSession();
  updateConnectivity();
  renderAuth();
  await registerServiceWorker();
})();
