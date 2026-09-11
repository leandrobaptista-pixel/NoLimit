const defaultState = {
  customServices: [
    { id: "SVC-DEMO-001", title: "Custom door modification", description: "Modify door height or width to fit an existing opening.", unit: "each", unitPrice: 600 },
  ],
  requests: [
    { id: "REQ-DEMO-001", clientId: "CL-DEMO-001", clientName: "Harbor Residence (Demo)", service: "Custom Trim", submitted: "Sep 8, 2026", status: "to-contact", nextAction: "Site visit · Sep 12" },
    { id: "REQ-DEMO-002", clientId: "CL-DEMO-002", clientName: "North Shore Builders (Demo)", service: "Stairs & Railings", submitted: "Sep 6, 2026", status: "proposal-sent", nextAction: "Follow up · Sep 13" },
    { id: "REQ-DEMO-003", clientId: "CL-DEMO-003", clientName: "Oak House (Demo)", service: "Kitchen Millwork", submitted: "Aug 19, 2026", status: "converted", nextAction: "Opened PR-DEMO-001" },
  ],
  siteVisits: [
    { id: "VIS-DEMO-001", requestId: "REQ-DEMO-001", clientId: "CL-DEMO-001", clientName: "Harbor Residence (Demo)", projectId: "", visitDate: "Sep 12, 2026", assignedTo: "Alex Morgan (Demo)", status: "scheduled", notes: "Measure trim scope and confirm finish selections." },
  ],
  estimates: [
    { id: "EST-DEMO-001", requestId: "REQ-DEMO-002", clientId: "CL-DEMO-002", clientName: "North Shore Builders (Demo)", projectId: "PR-DEMO-003", issueDate: "Sep 10, 2026", status: "sent", revision: 2, validUntil: "Sep 30, 2026", discount: 600, tax: 0, items: [{ category: "Stairs", description: "Custom stair and railing installation", quantity: 1, unitPrice: 20000 }], total: 19400, contractId: "", notes: "Scope and final measurements are subject to client approval." },
    { id: "EST-DEMO-002", requestId: "REQ-DEMO-003", clientId: "CL-DEMO-003", clientName: "Oak House (Demo)", projectId: "PR-DEMO-001", issueDate: "Aug 20, 2026", status: "approved", revision: 1, validUntil: "Sep 18, 2026", discount: 0, tax: 0, items: [{ category: "Kitchen & Vanities", description: "Kitchen millwork and built-ins", quantity: 1, unitPrice: 48500 }], total: 48500, contractId: "CTR-DEMO-001", notes: "Includes the work described above. Changes require written approval." },
  ],
  contracts: [
    { id: "CTR-DEMO-001", estimateId: "EST-DEMO-002", clientId: "CL-DEMO-003", clientName: "Oak House (Demo)", projectId: "PR-DEMO-001", status: "active", signedDate: "Sep 1, 2026", value: 48500, invoiceId: "INV-DEMO-001" },
  ],
  invoices: [
    { id: "INV-DEMO-001", contractId: "CTR-DEMO-001", projectId: "PR-DEMO-001", clientId: "CL-DEMO-003", clientName: "Oak House (Demo)", issueDate: "Sep 1, 2026", dueDate: "Oct 1, 2026", terms: "Net 30", status: "partially-paid", items: [{ category: "Kitchen & Vanities", description: "Kitchen millwork and built-ins", quantity: 1, unitPrice: 48500 }], schedule: [{ label: "Initial deposit", percent: 40, amount: 19400, status: "paid" }, { label: "Mid-project milestone", percent: 30, amount: 14550, status: "upcoming" }, { label: "Final completion", percent: 30, amount: 14550, status: "upcoming" }], total: 48500, paid: 19600, balance: 28900, notes: "Thank you for choosing No Limit Carpentry." },
  ],
  payments: [
    { id: "PAY-DEMO-001", invoiceId: "INV-DEMO-001", projectId: "PR-DEMO-001", date: "Sep 3, 2026", method: "Check", checkNumber: "DEMO-1001", milestone: "Initial deposit", amount: 19600, status: "cleared" },
  ],
  changeOrders: [
    { id: "CO-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", status: "pending", description: "Custom / New Work example", amount: 1800 },
  ],
  expenseReceipts: [
    { id: "RCP-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", vendor: "Demo Hardware Store", category: "Job supplies", purchaseDate: "2026-09-08", amount: 286.42, paymentMethod: "Company card", reference: "DEMO-4582", receiptFileName: "sample-receipt.pdf", notes: "Fasteners and installation supplies." },
  ],
  insuranceRenewals: [],
  clients: [
    { id: "CL-DEMO-001", name: "Harbor Residence (Demo)", type: "Homeowner", contactName: "Demo Contact", personalPhone: "(555) 010-1101", companyPhone: "", email: "harbor@example.invalid", website: "", street: "10 Sample Lane", city: "Red Bank", state: "NJ", postalCode: "00000", billingStreet: "", billingCity: "", billingState: "", billingPostalCode: "", projectIds: ["PR-DEMO-002"] },
    { id: "CL-DEMO-002", name: "North Shore Builders (Demo)", type: "General contractor", contactName: "Demo Office", personalPhone: "", companyPhone: "(555) 010-2202", email: "office@example.invalid", website: "https://example.invalid", street: "20 Example Avenue", city: "Rumson", state: "NJ", postalCode: "00000", billingStreet: "", billingCity: "", billingState: "", billingPostalCode: "", projectIds: ["PR-DEMO-003"] },
    { id: "CL-DEMO-003", name: "Oak House (Demo)", type: "Homeowner", contactName: "Demo Contact", personalPhone: "(555) 010-3303", companyPhone: "", email: "oak@example.invalid", website: "", street: "30 Preview Road", city: "Long Branch", state: "NJ", postalCode: "00000", billingStreet: "", billingCity: "", billingState: "", billingPostalCode: "", projectIds: ["PR-DEMO-001"] },
  ],
  projects: [
    { id: "PR-DEMO-001", name: "Oak House Millwork", clientId: "CL-DEMO-003", clientName: "Oak House (Demo)", status: "active", service: "Kitchen & Built-ins", siteStreet: "31 Preview Road", siteCity: "Long Branch", siteState: "NJ", sitePostalCode: "00000", startDate: "Sep 2, 2026", progress: 42, contractValue: 48500, cost: 17280, outstanding: 28900, managerId: "PE-DEMO-001" },
    { id: "PR-DEMO-002", name: "Harbor Architectural Trim", clientId: "CL-DEMO-001", clientName: "Harbor Residence (Demo)", status: "planned", service: "Custom Trim", siteStreet: "11 Sample Lane", siteCity: "Red Bank", siteState: "NJ", sitePostalCode: "00000", startDate: "Sep 21, 2026", progress: 8, contractValue: 22500, cost: 2500, outstanding: 15750, managerId: "PE-DEMO-001" },
    { id: "PR-DEMO-003", name: "Seaview Stair Upgrade", clientId: "CL-DEMO-002", clientName: "North Shore Builders (Demo)", status: "completed", service: "Stairs & Railings", siteStreet: "21 Example Avenue", siteCity: "Rumson", siteState: "NJ", sitePostalCode: "00000", startDate: "Jun 3, 2026", progress: 100, contractValue: 18400, cost: 11950, outstanding: 0, managerId: "PE-DEMO-002" },
  ],
  transactions: [
    { id: "TX-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", date: "Sep 8, 2026", type: "receivable", category: "Client invoice", party: "Oak House (Demo)", amount: 28900, status: "open" },
    { id: "TX-DEMO-002", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", date: "Sep 3, 2026", type: "receivable", category: "Client payment", party: "Oak House (Demo)", amount: 19600, status: "paid" },
    { id: "TX-DEMO-003", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", date: "Sep 7, 2026", type: "cost", category: "Materials", party: "Atlantic Millwork Supply (Demo)", amount: 5250, status: "due" },
    { id: "TX-DEMO-004", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", date: "Sep 6, 2026", type: "cost", category: "Labor", party: "Demo field crew", amount: 4230, status: "paid" },
    { id: "TX-DEMO-005", projectId: "PR-DEMO-003", projectName: "Seaview Stair Upgrade", date: "Jul 22, 2026", type: "receivable", category: "Final payment", party: "North Shore Builders (Demo)", amount: 18400, status: "paid" },
  ],
  people: [
    { id: "PE-DEMO-001", name: "Alex Morgan (Demo)", type: "Team Member", role: "Project manager", contactName: "Alex Morgan", personalPhone: "(555) 010-4101", companyPhone: "", email: "alex@example.invalid", website: "", street: "", city: "", state: "NJ", postalCode: "", projectIds: ["PR-DEMO-001", "PR-DEMO-002"], status: "active" },
    { id: "PE-DEMO-002", name: "Jordan Lee (Demo)", type: "Team Member", role: "Lead carpenter", contactName: "Jordan Lee", personalPhone: "(555) 010-4202", companyPhone: "", email: "jordan@example.invalid", website: "", street: "", city: "", state: "NJ", postalCode: "", projectIds: ["PR-DEMO-003"], status: "active" },
    { id: "PE-DEMO-003", name: "Atlantic Millwork Supply (Demo)", type: "Vendor", role: "Moldings & lumber", contactName: "Demo Sales", personalPhone: "", companyPhone: "(555) 010-4303", email: "sales@example.invalid", website: "https://example.invalid", street: "40 Vendor Drive", city: "Ocean", state: "NJ", postalCode: "00000", projectIds: ["PR-DEMO-001"], status: "active" },
    { id: "PE-DEMO-004", name: "Precision Rail Co. (Demo)", type: "Subcontractor", role: "Metal railings", contactName: "Demo Operations", personalPhone: "", companyPhone: "(555) 010-4404", email: "operations@example.invalid", website: "https://example.invalid", street: "50 Trade Road", city: "Rumson", state: "NJ", postalCode: "00000", projectIds: ["PR-DEMO-003"], status: "approved", w9Status: "verified", w9ReceivedDate: "2026-05-12", w9FileName: "precision-rail-w9.pdf", insuranceCompany: "Demo Insurance Co.", insuranceType: "General Liability", policyNumber: "DEMO-GL-2044", coverageAmount: 2000000, insuranceEffectiveDate: "2025-11-05", insuranceExpirationDate: "2026-11-05", insuranceFileName: "precision-rail-coi.pdf", renewalNoticeDays: "60,30" },
  ],
  media: [
    { id: "ME-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", phase: "In progress", fileType: "Photo", publishStatus: "awaiting-review", date: "Sep 8, 2026" },
    { id: "ME-DEMO-002", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", phase: "Before", fileType: "Photo", publishStatus: "internal", date: "Sep 2, 2026" },
    { id: "ME-DEMO-003", projectId: "PR-DEMO-003", projectName: "Seaview Stair Upgrade", phase: "Completed", fileType: "Photo", publishStatus: "approved", date: "Jul 22, 2026" },
    { id: "ME-DEMO-004", projectId: "PR-DEMO-003", projectName: "Seaview Stair Upgrade", phase: "Completed", fileType: "Video", publishStatus: "published", date: "Jul 22, 2026" },
    { id: "ME-DEMO-005", projectId: "PR-DEMO-002", projectName: "Harbor Architectural Trim", phase: "Site visit", fileType: "Photo", publishStatus: "awaiting-review", date: "Sep 9, 2026" },
  ],
  materials: [
    { id: "MT-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", vendorId: "PE-DEMO-003", vendor: "Atlantic Millwork Supply (Demo)", description: "White oak panels", quantity: 30, unit: "panels", unitCost: 175, amount: 5250, reference: "PO-DEMO-101", purchaseDate: "Sep 7, 2026", status: "ordered" },
    { id: "MT-DEMO-002", projectId: "PR-DEMO-002", projectName: "Harbor Architectural Trim", vendorId: "PE-DEMO-003", vendor: "Atlantic Millwork Supply (Demo)", description: "Custom trim package", quantity: 1, unit: "package", unitCost: 2500, amount: 2500, reference: "QUOTE-DEMO-202", purchaseDate: "Sep 9, 2026", status: "quoted" },
  ],
  schedule: [
    { id: "SCH-DEMO-001", personId: "PE-DEMO-001", personName: "Alex Morgan (Demo)", personType: "Team Member", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", workDate: "Sep 11, 2026", shift: "7:00 AM – 3:30 PM", instructions: "Continue kitchen panel installation.", status: "scheduled" },
    { id: "SCH-DEMO-002", personId: "PE-DEMO-004", personName: "Precision Rail Co. (Demo)", personType: "Subcontractor", projectId: "PR-DEMO-003", projectName: "Seaview Stair Upgrade", workDate: "Sep 11, 2026", shift: "9:00 AM – 1:00 PM", instructions: "Final railing inspection and punch list.", status: "scheduled" },
  ],
  attendance: [
    { id: "ATT-DEMO-001", scheduleId: "SCH-DEMO-001", personId: "PE-DEMO-001", personName: "Alex Morgan (Demo)", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", checkedInAt: "7:04 AM", locationStatus: "within-project-area", consentConfirmed: true, trackingEnds: "6:00 PM" },
  ],
  compliance: [
    { id: "CMP-DEMO-001", projectId: "PR-DEMO-001", projectName: "Oak House Millwork", jurisdiction: "New Jersey", municipality: "Long Branch", requirement: "Permit and inspection applicability reviewed", category: "Permit / Inspection", status: "pending", sourceLabel: "NJ DCA Uniform Construction Code", sourceUrl: "https://www.nj.gov/dca/codes/codreg/ucc.shtml" },
    { id: "CMP-DEMO-002", projectId: "PR-DEMO-002", projectName: "Harbor Architectural Trim", jurisdiction: "New Jersey", municipality: "Red Bank", requirement: "Local construction office requirements confirmed", category: "Municipal Review", status: "pending", sourceLabel: "NJ DCA Codes and Standards", sourceUrl: "https://www.nj.gov/dca/codes/" },
    { id: "CMP-DEMO-003", projectId: "PR-DEMO-003", projectName: "Seaview Stair Upgrade", jurisdiction: "New York State example", municipality: "Outside NYC", requirement: "Applicable Uniform Code edition documented", category: "Building Code", status: "reviewed", sourceLabel: "NYS Building Standards and Codes", sourceUrl: "https://dos.ny.gov/building-standards-and-codes" },
  ],
  consents: [
    { id: "CNS-DEMO-001", partyType: "Client", document: "Project terms and change-order approval", version: "Draft v1", status: "legal-review" },
    { id: "CNS-DEMO-002", partyType: "Vendor", document: "Schedule access and work-hour location consent", version: "Draft v1", status: "legal-review" },
    { id: "CNS-DEMO-003", partyType: "Subcontractor", document: "Schedule, project assignment, insurance, and location consent", version: "Draft v1", status: "legal-review" },
  ],
};

const serviceCategories = ["Trim", "Wainscoting", "Stairs", "Crown Molding · Ceiling · Coffered Ceiling", "Decks", "Kitchen & Vanities", "Fireplaces & Bars", "Outside Doors & Windows", "Pergola", "Port & Portal", "Commercial", "Trash Container", "Wall Paneling", "Custom / New Work"];
const scheduleAccessWindow = { starts: "6:00 AM", ends: "6:00 PM" };

const previewStorageKey = "no-limit-admin-preview-demo-v1";

function cloneDefaultState() {
  return JSON.parse(JSON.stringify(defaultState));
}

function loadPreviewState() {
  try {
    const saved = localStorage.getItem(previewStorageKey);
    if (!saved) return cloneDefaultState();
    const merged = { ...cloneDefaultState(), ...JSON.parse(saved) };
    merged.people = merged.people.map((person) => {
      const demoDefaults = defaultState.people.find((item) => item.id === person.id) || {};
      return {
        w9Status: "missing",
        w9ReceivedDate: "",
        w9FileName: "",
        insuranceCompany: "",
        insuranceType: "General Liability",
        policyNumber: "",
        coverageAmount: 0,
        insuranceEffectiveDate: "",
        insuranceExpirationDate: "",
        insuranceFileName: "",
        renewalNoticeDays: "60,30",
        ...demoDefaults,
        ...person,
        type: ["Employee", "Employer"].includes(person.type) ? "Team Member" : person.type === "Supplier" ? "Vendor" : person.type,
      };
    });
    merged.expenseReceipts = merged.expenseReceipts || [];
    merged.insuranceRenewals = merged.insuranceRenewals || [];
    merged.materials = merged.materials.map((item) => ({ ...item, vendor: item.vendor || item.supplier || "" }));
    merged.estimates = merged.estimates.map((estimate) => {
      const linkedContract = merged.contracts.find((contract) => contract.estimateId === estimate.id);
      const linkedProject = merged.projects.find((project) => project.id === linkedContract?.projectId) || merged.projects.find((project) => project.clientId === estimate.clientId);
      return { issueDate: estimate.issueDate || "Sep 10, 2026", projectId: estimate.projectId || linkedProject?.id || "", notes: estimate.notes || "Scope and final measurements are subject to client approval.", ...estimate };
    });
    merged.invoices = merged.invoices.map((invoice) => ({ dueDate: invoice.dueDate || "Oct 10, 2026", terms: invoice.terms || "Net 30", notes: invoice.notes || "Thank you for choosing No Limit Carpentry.", ...invoice }));
    return merged;
  } catch {
    return cloneDefaultState();
  }
}

let state = loadPreviewState();

const betaWorkspaceId = "shared-v1";
const isLocalPreview = ["127.0.0.1", "localhost"].includes(location.hostname);
let currentBetaUser = null;
let currentAuthSession = null;
let cloudSaveTimer = null;

function savePreviewState() {
  localStorage.setItem(previewStorageKey, JSON.stringify(state));
  if (currentBetaUser && currentAuthSession) {
    window.clearTimeout(cloudSaveTimer);
    syncBadge.textContent = "Saving…";
    syncBadge.className = "sync-badge is-saving";
    cloudSaveTimer = window.setTimeout(() => void pushBetaWorkspace(), 450);
  }
}

const routes = {
  overview: {
    title: "Overview",
    kicker: "Operational summary",
    heading: "Everything connected, without the noise.",
    description: "A single view of requests, projects, financial activity, people, media, and website performance. The No Limit database will be connected only after its isolated security foundation is approved.",
  },
  requests: {
    title: "Estimate Requests",
    kicker: "Lead management",
    heading: "Every request receives a clear next step.",
    description: "Requests from all 13 service categories will enter one organized workflow, from first contact through proposal and project conversion.",
  },
  documents: {
    title: "Estimates & Contracts",
    kicker: "Sales documents",
    heading: "From estimate to contract without entering data twice.",
    description: "Build itemized estimates, keep every revision, record client approval, and generate the related contract and invoice from the same approved information.",
  },
  clients: {
    title: "Clients",
    kicker: "Client relationships",
    heading: "One client, every project in context.",
    description: "Client profiles will connect contacts, addresses, requests, proposals, documents, and all related projects without duplicate entry.",
  },
  projects: {
    title: "Projects",
    kicker: "Job-site operations",
    heading: "Follow each project from plan to completion.",
    description: "Projects connect the client, team, subcontractors, schedule, contracts, materials, costs, progress media, and final results.",
  },
  financial: {
    title: "Financial",
    kicker: "Financial control",
    heading: "Know the cost, balance, and result of every job.",
    description: "Track proposed and contracted values, materials, labor, subcontractors, payments, receipts, change orders, and project results.",
  },
  materials: {
    title: "Materials",
    kicker: "Project purchasing",
    heading: "Every material, Vendor, cost, and project in one record.",
    description: "Track quotes, purchase orders, quantities, unit costs, totals, receipts, delivery status, and the project that received each item.",
  },
  team: {
    title: "Team & Partners",
    kicker: "People and companies",
    heading: "The right access for every person involved.",
    description: "Team members, vendors, and subcontractors remain distinct while connecting to the projects, services, documents, hours, and payments that concern them.",
  },
  schedule: {
    title: "Work Schedule",
    kicker: "Individual assignments",
    heading: "Everyone sees the right work for the next day.",
    description: "Administrators assign work by person, company, project, date, and shift. Future private accounts will show each authorized user only the schedule that concerns them.",
  },
  media: {
    title: "Media Library",
    kicker: "Photos and videos",
    heading: "Find every project image without searching folders.",
    description: "Media will be organized by client, project, date, category, and phase, with separate permissions for upload, approval, and public website publishing.",
  },
  map: {
    title: "Project Map",
    kicker: "Location overview",
    heading: "See where every active job stands.",
    description: "Filter planned, active, paused, and completed projects while keeping private addresses protected by user permissions.",
  },
  compliance: {
    title: "Compliance",
    kicker: "Codes, permits, and consent",
    heading: "Document the rules that apply to every project.",
    description: "Use project-specific checklists for New Jersey, New York State, New York City, and municipal requirements, while keeping signed consent records versioned by party type.",
  },
  reports: {
    title: "Reports",
    kicker: "Custom report builder",
    heading: "Choose exactly what the report should contain.",
    description: "Mix operational, financial, people, material, contract, and progress sections into a clean No Limit report before printing or saving as PDF.",
  },
  security: {
    title: "Security & Settings",
    kicker: "Protected foundation",
    heading: "No Limit data stays separate by design.",
    description: "Authentication, database records, files, permissions, audit history, and backups will remain isolated from TAG and every unrelated system.",
  },
};

const content = document.getElementById("appContent");
const nav = document.getElementById("primaryNav");
const topbarTitle = document.getElementById("topbarTitle");
const menuButton = document.getElementById("menuButton");
const mobileOverlay = document.getElementById("mobileOverlay");
const dataDialog = document.getElementById("dataDialog");
const dataForm = document.getElementById("dataForm");
const dialogTitle = document.getElementById("dialogTitle");
const dialogFields = document.getElementById("dialogFields");
const documentDialog = document.getElementById("documentDialog");
const documentDialogTitle = document.getElementById("documentDialogTitle");
const documentCanvas = document.getElementById("documentCanvas");
const documentNotice = document.getElementById("documentNotice");
const editDocumentButton = document.getElementById("editDocument");
const addDocumentItemButton = document.getElementById("addDocumentItem");
const saveDocumentButton = document.getElementById("saveDocument");
const printDocumentButton = document.getElementById("printDocument");
const emailDocumentButton = document.getElementById("emailDocument");
const messageDocumentButton = document.getElementById("messageDocument");
const closeDocumentButton = document.getElementById("closeDocument");
const customServiceDialog = document.getElementById("customServiceDialog");
const customServiceForm = document.getElementById("customServiceForm");
const closeCustomServiceButton = document.getElementById("closeCustomService");
const cancelCustomServiceButton = document.getElementById("cancelCustomService");
const authView = document.getElementById("authView");
const loginForm = document.getElementById("loginForm");
const setPasswordForm = document.getElementById("setPasswordForm");
const resetPasswordButton = document.getElementById("resetPasswordButton");
const authStatus = document.getElementById("authStatus");
const syncBadge = document.getElementById("syncBadge");
const profileButton = document.getElementById("profileButton");

let pendingRecordType = "";
let pendingRecordId = "";
let activeDocument = null;
let pendingCustomServiceTarget = null;

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0));
}

function formatStatus(value = "") {
  return String(value).replaceAll("-", " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function statusClass(value = "") {
  return ["active", "paid", "approved", "completed", "published", "converted"].includes(value) ? "green" : "amber";
}

function projectOptions(selected = "") {
  return state.projects.map((project) => `<option value="${escapeHtml(project.id)}" ${project.id === selected ? "selected" : ""}>${escapeHtml(project.name)} · ${escapeHtml(project.id)}</option>`).join("");
}

function clientOptions(selected = "") {
  return state.clients.map((client) => `<option value="${escapeHtml(client.id)}" ${client.id === selected ? "selected" : ""}>${escapeHtml(client.name)}</option>`).join("");
}

function serviceOptions(selected = "", selectedCustomId = "") {
  const standard = serviceCategories.map((service) => `<option value="${escapeHtml(service)}" ${service === selected && !selectedCustomId ? "selected" : ""}>${escapeHtml(service)}</option>`).join("");
  const custom = state.customServices.length ? `<optgroup label="Saved Custom / New Jobs">${state.customServices.map((service) => `<option value="custom-service:${escapeHtml(service.id)}" ${service.id === selectedCustomId ? "selected" : ""}>${escapeHtml(service.title)}</option>`).join("")}</optgroup>` : "";
  return standard + custom;
}

function customServiceFromSelection(value = "") {
  if (!String(value).startsWith("custom-service:")) return null;
  return state.customServices.find((service) => service.id === String(value).slice("custom-service:".length)) || null;
}

function requestOptions() {
  return state.requests.map((request) => `<option value="${escapeHtml(request.id)}">${escapeHtml(request.id)} · ${escapeHtml(request.clientName)}</option>`).join("");
}

function invoiceOptions() {
  return state.invoices.map((invoice) => `<option value="${escapeHtml(invoice.id)}">${escapeHtml(invoice.id)} · ${escapeHtml(invoice.clientName)} · ${formatCurrency(invoice.balance)} balance</option>`).join("");
}

function peopleOptions() {
  return state.people.map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)} · ${escapeHtml(person.type)}</option>`).join("");
}

function vendorOptions() {
  return state.people.filter((person) => person.type === "Vendor").map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)}</option>`).join("");
}

function reportPeopleOptions() {
  return state.people.map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)} · ${escapeHtml(person.type)}</option>`).join("");
}

function materialOptions() {
  return state.materials.map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.description)} · ${escapeHtml(item.projectName)}</option>`).join("");
}

function clientLocation(client) {
  if (client.street || client.state || client.postalCode) return [client.street, client.city, client.state, client.postalCode].filter(Boolean).join(", ");
  return client.city || "Not entered";
}

function projectLocation(project) {
  if (project.siteStreet || project.siteState || project.sitePostalCode) return [project.siteStreet, project.siteCity, project.siteState, project.sitePostalCode].filter(Boolean).join(", ");
  return project.city || "Not entered";
}

function clientContact(client) {
  return [client.contactName || client.contact, client.companyPhone || client.personalPhone, client.email].filter(Boolean).join(" · ") || "Not entered";
}

function partyLocation(party) {
  return [party.street, party.city, party.state, party.postalCode].filter(Boolean).join(", ") || "Not entered";
}

function partyContact(party) {
  return [party.contactName, party.companyPhone || party.personalPhone, party.email, party.website].filter(Boolean).join(" · ") || "Not entered";
}

function demoTable(headers, rows) {
  return `<div class="table-shell"><table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.join("")}</tbody></table></div>`;
}

function addressLines(parts) {
  return parts.filter(Boolean).map((part) => escapeHtml(part)).join("<br />") || "Not entered";
}

function billingAddress(client) {
  const useBilling = client.billingStreet || client.billingCity || client.billingState || client.billingPostalCode;
  return addressLines(useBilling
    ? [client.billingStreet, [client.billingCity, client.billingState, client.billingPostalCode].filter(Boolean).join(" ")]
    : [client.street, [client.city, client.state, client.postalCode].filter(Boolean).join(" ")]);
}

function projectAddress(project) {
  if (!project) return "Not assigned yet";
  return addressLines([project.siteStreet, [project.siteCity, project.siteState, project.sitePostalCode].filter(Boolean).join(" ")]);
}

function editableDocumentCell(value, name, index, type = "text") {
  const step = type === "number" ? ' step="0.01" min="0" inputmode="decimal"' : "";
  return `<input class="document-input" type="${type}" name="${name}" data-document-item="${index}" value="${escapeHtml(value)}"${step} />`;
}

function documentContext(type, id) {
  const record = type === "proposal" ? state.estimates.find((item) => item.id === id) : state.invoices.find((item) => item.id === id);
  if (!record) return null;
  const client = state.clients.find((item) => item.id === record.clientId) || {};
  const project = state.projects.find((item) => item.id === record.projectId) || null;
  return { type, record, client, project };
}

function renderBusinessDocument(context, editing = false, itemDraft = null) {
  const { type, record, client, project } = context;
  const isProposal = type === "proposal";
  const title = isProposal ? "PROPOSAL" : "INVOICE";
  const documentItems = itemDraft || record.items;
  const subtotal = documentItems.reduce((sum, item) => sum + Number(item.quantity || 0) * Number(item.unitPrice || 0), 0);
  const discount = isProposal ? Number(record.discount || 0) : 0;
  const tax = isProposal ? Number(record.tax || 0) : 0;
  const grandTotal = isProposal ? Math.max(0, subtotal - discount) + tax : Number(record.total || subtotal);
  const rows = documentItems.map((item, index) => `
    <tr>
      <td>${index + 1}</td>
      <td>${editing ? `<select class="document-input" name="category" data-document-item="${index}">${serviceOptions(item.category, item.customServiceId)}</select><button class="remove-line" data-remove-document-item="${index}" type="button">Remove</button>` : escapeHtml(item.category)}</td>
      <td>${editing ? editableDocumentCell(item.title || item.category, "title", index) : `<strong>${escapeHtml(item.title || item.category)}</strong>`}</td>
      <td>${editing ? editableDocumentCell(item.description, "description", index) : escapeHtml(item.description)}</td>
      <td>${editing ? editableDocumentCell(item.quantity, "quantity", index, "number") : Number(item.quantity || 0).toLocaleString("en-US")}</td>
      <td>${editing ? editableDocumentCell(item.unitPrice, "unitPrice", index, "number") : formatCurrency(item.unitPrice)}</td>
      <td>${formatCurrency(Number(item.quantity || 0) * Number(item.unitPrice || 0))}</td>
    </tr>`).join("");
  const documentMeta = isProposal
    ? `<dl><div><dt>Proposal no.</dt><dd>${escapeHtml(record.id)}</dd></div><div><dt>Revision</dt><dd>R${Number(record.revision || 1)}</dd></div><div><dt>Issue date</dt><dd>${escapeHtml(record.issueDate || "Not entered")}</dd></div><div><dt>Valid until</dt><dd>${escapeHtml(record.validUntil || "Not entered")}</dd></div></dl>`
    : `<dl><div><dt>Invoice no.</dt><dd>${escapeHtml(record.id)}</dd></div><div><dt>Terms</dt><dd>${escapeHtml(record.terms || "Net 30")}</dd></div><div><dt>Invoice date</dt><dd>${escapeHtml(record.issueDate || "Not entered")}</dd></div><div><dt>Due date</dt><dd>${escapeHtml(record.dueDate || "Not entered")}</dd></div></dl>`;
  const totals = isProposal
    ? `<div><span>Subtotal</span><strong>${formatCurrency(subtotal)}</strong></div><div><span>Discount</span><strong>− ${formatCurrency(discount)}</strong></div><div><span>Tax</span><strong>${formatCurrency(tax)}</strong></div><div class="grand-total"><span>Proposal total</span><strong>${formatCurrency(grandTotal)}</strong></div>`
    : `<div><span>Total</span><strong>${formatCurrency(grandTotal)}</strong></div><div><span>Paid</span><strong>− ${formatCurrency(record.paid || 0)}</strong></div><div class="grand-total"><span>Balance due</span><strong>${formatCurrency(record.balance ?? grandTotal)}</strong></div>`;
  const stages = !isProposal && record.schedule?.length ? `<section class="document-section"><h3>Payment stages</h3><div class="document-stage-list">${record.schedule.map((stage) => `<div><span>${escapeHtml(stage.label)} · ${Number(stage.percent || 0).toLocaleString("en-US")}%</span><strong>${formatCurrency(stage.amount)}</strong><small>${escapeHtml(formatStatus(stage.status))}</small></div>`).join("")}</div></section>` : "";
  return `
    <header class="document-brand">
      <div class="document-company"><img src="logo.png" alt="No Limit Carpentry" /><p><strong>No Limit Contractor, LLC.</strong><br />2137 Aldrin Rd Apt 6B<br />Ocean, NJ 07712-2466</p></div>
      <div class="document-company-contact"><strong>${title}</strong><p>leandrobaptista@me.com<br />+1 (848) 466-3339</p></div>
    </header>
    <section class="document-parties">
      <div><span>Bill to</span><h2>${escapeHtml(client.name || record.clientName)}</h2><p>${escapeHtml(client.contactName || "")}<br />${billingAddress(client)}<br />${escapeHtml(client.email || "")}</p></div>
      <div><span>Project</span><h2>${escapeHtml(project?.name || "Project pending")}</h2><p>${projectAddress(project)}</p></div>
    </section>
    <section class="document-meta"><h3>${title.toLowerCase().replace(/^./, (letter) => letter.toUpperCase())} details</h3>${documentMeta}</section>
    <section class="document-items">
      <table><thead><tr><th>#</th><th>Category</th><th>Service</th><th>Description</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody>${rows}</tbody></table>
    </section>
    <section class="document-summary">
      <div class="document-notes"><h3>${isProposal ? "Scope and terms" : "Notes"}</h3>${editing ? `<textarea class="document-input document-notes-input" name="documentNotes">${escapeHtml(record.notes || "")}</textarea>` : `<p>${escapeHtml(record.notes || "No notes entered.")}</p>`}</div>
      <div class="document-totals">${totals}</div>
    </section>
    ${stages}
    ${isProposal ? '<section class="document-approval"><div><span>Client approval</span><i></i></div><div><span>Date</span><i></i></div></section>' : ""}
    <footer class="document-footer"><span>No Limit Carpentry</span><span>Crafted for the way you live.</span></footer>`;
}

function openBusinessDocument(type, id) {
  const context = documentContext(type, id);
  if (!context) return;
  activeDocument = { type, id, editing: false };
  documentDialogTitle.textContent = type === "proposal" ? "Proposal preview" : "Invoice preview";
  documentCanvas.innerHTML = renderBusinessDocument(context, false);
  documentNotice.textContent = "Nothing is sent automatically. Review the document before preparing a message.";
  editDocumentButton.hidden = false;
  addDocumentItemButton.hidden = true;
  saveDocumentButton.hidden = true;
  documentDialog.showModal();
}

function setDocumentEditing(editing) {
  if (!activeDocument) return;
  activeDocument.editing = editing;
  const context = documentContext(activeDocument.type, activeDocument.id);
  if (editing && !activeDocument.draftItems) activeDocument.draftItems = JSON.parse(JSON.stringify(context.record.items));
  documentCanvas.innerHTML = renderBusinessDocument(context, editing, editing ? activeDocument.draftItems : null);
  editDocumentButton.hidden = editing;
  addDocumentItemButton.hidden = !editing;
  saveDocumentButton.hidden = !editing;
  documentNotice.textContent = editing ? "Adjust descriptions, quantities, rates, and notes. Saving creates the next proposal revision or updates the invoice draft in this preview." : "Changes saved in this private preview. Nothing has been sent.";
  if (editing) bindDocumentEditor();
}

function harvestDocumentDraft() {
  if (!activeDocument?.draftItems) return;
  activeDocument.draftItems.forEach((item, index) => {
    const selectedCategory = documentCanvas.querySelector(`[name="category"][data-document-item="${index}"]`)?.value || "Custom / New Work";
    const savedService = customServiceFromSelection(selectedCategory);
    item.category = savedService ? "Custom / New Work" : selectedCategory;
    item.customServiceId = savedService?.id || (selectedCategory === "Custom / New Work" ? item.customServiceId || "" : "");
    item.title = documentCanvas.querySelector(`[name="title"][data-document-item="${index}"]`)?.value || item.category;
    item.description = documentCanvas.querySelector(`[name="description"][data-document-item="${index}"]`)?.value || "";
    item.quantity = Number(documentCanvas.querySelector(`[name="quantity"][data-document-item="${index}"]`)?.value || 0);
    item.unitPrice = Number(documentCanvas.querySelector(`[name="unitPrice"][data-document-item="${index}"]`)?.value || 0);
  });
  activeDocument.draftNotes = documentCanvas.querySelector('[name="documentNotes"]')?.value || "";
}

function bindDocumentEditor() {
  documentCanvas.querySelectorAll('select[name="category"]').forEach((select) => select.addEventListener("change", () => {
    harvestDocumentDraft();
    const index = Number(select.dataset.documentItem);
    const savedService = customServiceFromSelection(select.value);
    if (savedService) {
      activeDocument.draftItems[index] = { ...activeDocument.draftItems[index], category: "Custom / New Work", customServiceId: savedService.id, title: savedService.title, description: savedService.description, unitPrice: savedService.unitPrice };
      setDocumentEditing(true);
      return;
    }
    if (select.value === "Custom / New Work") {
      pendingCustomServiceTarget = { mode: "document", index };
      customServiceForm.reset();
      customServiceDialog.showModal();
    }
  }));
  documentCanvas.querySelectorAll("[data-remove-document-item]").forEach((button) => button.addEventListener("click", () => {
    harvestDocumentDraft();
    if (activeDocument.draftItems.length <= 1) {
      documentNotice.textContent = "A Proposal or Invoice must keep at least one line item.";
      return;
    }
    activeDocument.draftItems.splice(Number(button.dataset.removeDocumentItem), 1);
    setDocumentEditing(true);
  }));
}

function addDocumentLineItem() {
  if (!activeDocument?.editing) return;
  harvestDocumentDraft();
  activeDocument.draftItems.push({ category: "Custom / New Work", title: "", description: "", quantity: 1, unitPrice: 0 });
  setDocumentEditing(true);
}

function closeCustomServiceDialog() {
  customServiceDialog.close();
  pendingCustomServiceTarget = null;
}

function saveCustomService(formData) {
  const service = {
    id: nextId(state.customServices, "SVC"),
    title: formData.get("title"),
    description: formData.get("description"),
    unit: formData.get("unit"),
    unitPrice: Number(formData.get("unitPrice")),
  };
  state.customServices.unshift(service);
  savePreviewState();
  if (pendingCustomServiceTarget?.mode === "document" && activeDocument?.draftItems) {
    const item = activeDocument.draftItems[pendingCustomServiceTarget.index];
    activeDocument.draftItems[pendingCustomServiceTarget.index] = { ...item, category: "Custom / New Work", customServiceId: service.id, title: service.title, description: service.description, unitPrice: service.unitPrice };
    customServiceDialog.close();
    pendingCustomServiceTarget = null;
    setDocumentEditing(true);
    return;
  }
  if (pendingCustomServiceTarget?.mode === "data-form") {
    const categoryControl = dataForm.elements.namedItem("category");
    categoryControl.innerHTML = serviceOptions("Custom / New Work", service.id);
    categoryControl.value = `custom-service:${service.id}`;
    const titleControl = dataForm.elements.namedItem("title");
    const descriptionControl = dataForm.elements.namedItem("description");
    const priceControl = dataForm.elements.namedItem("unitPrice");
    if (titleControl) titleControl.value = service.title;
    if (descriptionControl) descriptionControl.value = service.description;
    if (priceControl) priceControl.value = service.unitPrice;
  }
  customServiceDialog.close();
  pendingCustomServiceTarget = null;
}

function saveDocumentEdits() {
  if (!activeDocument) return;
  const context = documentContext(activeDocument.type, activeDocument.id);
  if (!context) return;
  harvestDocumentDraft();
  context.record.items = JSON.parse(JSON.stringify(activeDocument.draftItems));
  context.record.notes = activeDocument.draftNotes || "";
  const subtotal = context.record.items.reduce((sum, item) => sum + Number(item.quantity) * Number(item.unitPrice), 0);
  if (activeDocument.type === "proposal") {
    context.record.total = Math.max(0, subtotal - Number(context.record.discount || 0)) + Number(context.record.tax || 0);
    context.record.revision = Number(context.record.revision || 1) + 1;
  } else {
    context.record.total = subtotal;
    context.record.balance = Math.max(0, subtotal - Number(context.record.paid || 0));
  }
  savePreviewState();
  delete activeDocument.draftItems;
  delete activeDocument.draftNotes;
  setDocumentEditing(false);
  renderRoute();
}

function prepareDocumentDelivery(channel) {
  if (!activeDocument) return;
  const context = documentContext(activeDocument.type, activeDocument.id);
  if (!context) return;
  const label = activeDocument.type === "proposal" ? "Proposal" : "Invoice";
  const recipient = channel === "email" ? context.client.email : (context.client.companyPhone || context.client.personalPhone);
  if (!recipient) {
    documentNotice.textContent = `Add the client's ${channel === "email" ? "email address" : "phone number"} before preparing this ${channel}.`;
    return;
  }
  if (!window.confirm(`Prepare this ${label.toLowerCase()} for ${recipient}? Nothing will be sent until you confirm it in your ${channel === "email" ? "email" : "messaging"} app.`)) return;
  const subject = `${label} ${context.record.id} — No Limit Carpentry`;
  const body = `Hello ${context.client.contactName || context.client.name || ""},\n\nYour ${label.toLowerCase()} ${context.record.id} for ${context.project?.name || "your project"} is ready. Please review the attached PDF.\n\nThank you,\nNo Limit Carpentry\n(848) 466-3339`;
  const href = channel === "email"
    ? `mailto:${encodeURIComponent(recipient)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
    : `sms:${String(recipient).replace(/[^+\d]/g, "")}?&body=${encodeURIComponent(`${subject}. Please review the PDF from No Limit Carpentry.`)}`;
  window.location.href = href;
}

function field(label, name, options = {}) {
  const required = options.required === false ? "" : "required";
  if (options.options) {
    return `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}" ${required}>${options.options}</select></label>`;
  }
  if (options.multiline) {
    return `<label class="field-wide">${escapeHtml(label)}<textarea name="${escapeHtml(name)}" ${options.placeholder ? `placeholder="${escapeHtml(options.placeholder)}"` : ""} ${required}></textarea></label>`;
  }
  const isMoney = ["unitPrice", "discount", "contractValue", "cost", "outstanding", "amount", "coverageAmount"].includes(name);
  return `<label>${escapeHtml(label)}<input name="${escapeHtml(name)}" type="${escapeHtml(options.type || "text")}" ${options.min !== undefined ? `min="${options.min}"` : ""} ${options.max !== undefined ? `max="${options.max}"` : ""} ${isMoney ? 'step="0.01" inputmode="decimal" placeholder="$0.00"' : options.placeholder ? `placeholder="${escapeHtml(options.placeholder)}"` : ""} ${required} /></label>`;
}

function openDataEntry(type, recordId = "") {
  pendingRecordType = type;
  pendingRecordId = recordId;
  const statusOptions = (items) => items.map((item) => `<option value="${item}">${formatStatus(item)}</option>`).join("");
  const forms = {
    request: {
      title: "Add estimate request",
      fields: field("Existing client", "clientId", { required: false, options: '<option value="">New or unlinked lead</option>' + clientOptions() }) + field("Client or lead name", "clientName", { placeholder: "Demo name" }) + field("Service", "service", { placeholder: "Custom Trim" }) + field("Status", "status", { options: statusOptions(["new", "to-contact", "contacted", "proposal-sent", "converted"]) }) + field("Next action", "nextAction", { placeholder: "Site visit · Sep 15" }),
    },
    siteVisit: {
      title: "Schedule Free On-Site Visit",
      fields: field("Estimate request", "requestId", { options: requestOptions() }) + field("Client", "clientId", { options: clientOptions() }) + field("Related project", "projectId", { required: false, options: '<option value="">Not created yet</option>' + projectOptions() }) + field("Visit date", "visitDate", { type: "date" }) + field("Assigned to", "assignedTo", { placeholder: "Team Member" }) + field("Status", "status", { options: statusOptions(["scheduled", "confirmed", "completed", "cancelled"]) }) + field("Notes", "notes", { placeholder: "Measurements, scope, finish selections..." }),
    },
    estimate: {
      title: recordId ? "Edit proposal" : "Add proposal",
      fields: field("Estimate request", "requestId", { options: requestOptions() }) + field("Client", "clientId", { options: clientOptions() }) + field("Related project", "projectId", { required: false, options: '<option value="">Not assigned yet</option>' + projectOptions() }) + field("Issue date", "issueDate", { type: "date" }) + field("Service category", "category", { options: serviceOptions() }) + field("Service title", "title", { placeholder: "Door installation, custom stair, drop ceiling..." }) + field("Line item description", "description", { placeholder: "Detailed scope of work" }) + field("Quantity", "quantity", { type: "number", min: 0 }) + field("Unit price", "unitPrice", { type: "number", min: 0 }) + field("Discount", "discount", { required: false, type: "number", min: 0 }) + field("Tax %", "taxPercent", { required: false, type: "number", min: 0 }) + field("Valid until", "validUntil", { type: "date" }) + field("Status", "status", { options: statusOptions(["draft", "sent", "approved", "declined"]) }) + field("Proposal notes and terms", "notes", { required: false, multiline: true, placeholder: "Scope, exclusions, payment stages, and approval notes" }),
    },
    client: {
      title: recordId ? "Edit client" : "Add client",
      fields: field("Client or company name", "name", { placeholder: "Demo client" }) + field("Client type", "type", { options: '<option>Homeowner</option><option>General contractor</option><option>Architect / Designer</option>' }) + field("Primary contact", "contactName", { placeholder: "Contact name" }) + field("Personal phone", "personalPhone", { required: false, type: "tel" }) + field("Company phone", "companyPhone", { required: false, type: "tel" }) + field("Primary email", "email", { required: false, type: "email" }) + field("Website", "website", { required: false, type: "url", placeholder: "https://" }) + field("Client street address", "street", { required: false }) + field("Client city", "city", { required: false }) + field("Client state", "state", { required: false, placeholder: "NJ" }) + field("Client ZIP code", "postalCode", { required: false }) + field("Billing street (if different)", "billingStreet", { required: false }) + field("Billing city", "billingCity", { required: false }) + field("Billing state", "billingState", { required: false }) + field("Billing ZIP code", "billingPostalCode", { required: false }),
    },
    project: {
      title: recordId ? "Edit project" : "Add project",
      fields: field("Project name", "name", { placeholder: "Demo project" }) + field("Client", "clientId", { options: clientOptions() }) + field("Service", "service", { placeholder: "Kitchen & Built-ins" }) + field("Project street address", "siteStreet") + field("Project city", "siteCity") + field("Project state", "siteState", { placeholder: "NJ" }) + field("Project ZIP code", "sitePostalCode") + field("Status", "status", { options: statusOptions(["planned", "active", "paused", "completed"]) }) + field("Start date", "startDate", { type: "date" }) + field("Progress %", "progress", { type: "number", min: 0, max: 100 }) + field("Contract value", "contractValue", { type: "number", min: 0 }) + field("Current cost", "cost", { type: "number", min: 0 }) + field("Outstanding", "outstanding", { type: "number", min: 0 }),
    },
    transaction: {
      title: "Add financial transaction",
      fields: field("Project", "projectId", { options: projectOptions() }) + field("Transaction type", "type", { options: '<option value="receivable">Client receivable</option><option value="cost">Project cost</option>' }) + field("Category", "category", { placeholder: "Materials, labor, payment..." }) + field("Related party", "party", { placeholder: "Demo client or vendor" }) + field("Amount", "amount", { type: "number", min: 0 }) + field("Status", "status", { options: statusOptions(["open", "due", "paid"]) }),
    },
    invoiceItem: {
      title: "Add invoice item",
      fields: field("Invoice", "invoiceId", { options: invoiceOptions() }) + field("Service category", "category", { options: serviceOptions() }) + field("Service title", "title", { placeholder: "Door installation, casing, baseboard..." }) + field("Description", "description", { placeholder: "Detailed scope for this line" }) + field("Quantity", "quantity", { type: "number", min: 0 }) + field("Unit price", "unitPrice", { type: "number", min: 0 }),
    },
    payment: {
      title: "Record client payment",
      fields: field("Invoice", "invoiceId", { options: invoiceOptions() }) + field("Payment date", "date", { type: "date" }) + field("Payment method", "method", { options: '<option>Check</option><option>ACH / Bank transfer</option><option>Credit card</option><option>Cash</option><option>Other</option>' }) + field("Check or reference number", "checkNumber", { required: false }) + field("Project milestone", "milestone", { placeholder: "Initial deposit, midpoint, completion..." }) + field("Amount", "amount", { type: "number", min: 0 }) + field("Status", "status", { options: statusOptions(["received", "pending", "cleared", "returned"]) }),
    },
    changeOrder: {
      title: "Add change order / new work",
      fields: field("Project", "projectId", { options: projectOptions() }) + field("Service category", "category", { options: serviceOptions("Custom / New Work") }) + field("Service title", "title", { placeholder: "Additional custom work" }) + field("Description", "description", { placeholder: "Describe the additional work" }) + field("Quantity", "quantity", { type: "number", min: 0 }) + field("Unit price", "unitPrice", { type: "number", min: 0 }) + field("Status", "status", { options: statusOptions(["draft", "pending", "approved", "declined"]) }),
    },
    person: {
      title: recordId ? "Edit Team Member, Vendor, or Subcontractor" : "Add Team Member, Vendor, or Subcontractor",
      fields: field("Person or company name", "name", { placeholder: "Demo name" }) + field("Type", "type", { options: '<option>Team Member</option><option>Vendor</option><option>Subcontractor</option>' }) + field("Role or specialty", "role", { placeholder: "Lead carpenter, lumber, railings..." }) + field("Primary contact", "contactName", { required: false }) + field("Personal phone", "personalPhone", { required: false, type: "tel" }) + field("Company phone", "companyPhone", { required: false, type: "tel" }) + field("Primary email", "email", { required: false, type: "email" }) + field("Website", "website", { required: false, type: "url", placeholder: "https://" }) + field("Street address", "street", { required: false }) + field("City", "city", { required: false }) + field("State", "state", { required: false, placeholder: "NJ" }) + field("ZIP code", "postalCode", { required: false }) + field("Project", "projectId", { required: false, options: '<option value="">Not assigned</option>' + projectOptions() }) + field("Status", "status", { options: statusOptions(["active", "approved", "inactive"]) }) + '<div class="form-section-title field-wide"><strong>Subcontractor compliance</strong><span>Used only when the record type is Subcontractor.</span></div>' + field("W-9 status", "w9Status", { required: false, options: '<option value="missing">Missing</option><option value="requested">Requested</option><option value="received">Received</option><option value="verified">Verified</option>' }) + field("W-9 received date", "w9ReceivedDate", { required: false, type: "date" }) + field("Upload W-9", "w9File", { required: false, type: "file" }) + field("Insurance company", "insuranceCompany", { required: false }) + field("Insurance type", "insuranceType", { required: false, options: '<option>General Liability</option><option>Workers’ Compensation</option><option>Commercial Auto</option><option>Umbrella</option><option>Other</option>' }) + field("Policy number", "policyNumber", { required: false }) + field("Coverage amount", "coverageAmount", { required: false, type: "number", min: 0 }) + field("Effective date", "insuranceEffectiveDate", { required: false, type: "date" }) + field("Expiration date", "insuranceExpirationDate", { required: false, type: "date" }) + field("Upload Certificate of Insurance", "insuranceFile", { required: false, type: "file" }) + field("Renewal alerts", "renewalNoticeDays", { required: false, options: '<option value="60,30">60 and 30 days before</option><option value="90,60,30">90, 60, and 30 days before</option><option value="30">30 days before</option>' }),
    },
    receipt: {
      title: "Add project receipt or expense",
      fields: field("Project", "projectId", { options: projectOptions() }) + field("Vendor or store", "vendor", { placeholder: "Company or person paid" }) + field("Expense category", "category", { options: '<option>Materials</option><option>Job supplies</option><option>Equipment rental</option><option>Subcontractor</option><option>Permits and fees</option><option>Travel and delivery</option><option>Other</option>' }) + field("Purchase date", "purchaseDate", { type: "date" }) + field("Amount", "amount", { type: "number", min: 0 }) + field("Payment method", "paymentMethod", { options: '<option>Company card</option><option>Check</option><option>Cash</option><option>ACH / Bank transfer</option><option>Other</option>' }) + field("Receipt or reference number", "reference", { required: false }) + field("Upload receipt", "receiptFile", { required: false, type: "file" }) + field("Notes", "notes", { required: false, multiline: true, placeholder: "What this expense covered" }),
    },
    material: {
      title: "Add project material",
      fields: field("Project", "projectId", { options: projectOptions() }) + field("Vendor", "vendorId", { options: vendorOptions() }) + field("Material or Custom / New Item", "description", { placeholder: "Material description" }) + field("Quantity", "quantity", { type: "number", min: 0 }) + field("Unit", "unit", { placeholder: "pieces, boxes, linear feet..." }) + field("Unit cost", "unitCost", { type: "number", min: 0 }) + field("Purchase or quote date", "purchaseDate", { type: "date" }) + field("PO, receipt, or quote number", "reference", { required: false }) + field("Status", "status", { options: statusOptions(["quoted", "ordered", "received", "installed", "returned"]) }),
    },
    schedule: {
      title: "Add work assignment",
      fields: field("Team Member, Vendor, or Subcontractor", "personId", { options: peopleOptions() }) + field("Project", "projectId", { options: projectOptions() }) + field("Work date", "workDate", { type: "date" }) + field("Shift", "shift", { placeholder: "7:00 AM – 3:30 PM" }) + field("Instructions", "instructions", { placeholder: "Work planned for this date" }) + field("Status", "status", { options: statusOptions(["scheduled", "confirmed", "completed", "cancelled"]) }),
    },
    compliance: {
      title: "Add compliance requirement",
      fields: field("Project", "projectId", { options: projectOptions() }) + field("Jurisdiction", "jurisdiction", { options: '<option>New Jersey</option><option>New York State</option><option>New York City</option><option>Other / Municipal</option>' }) + field("Municipality", "municipality", { placeholder: "City, town, or borough" }) + field("Category", "category", { options: '<option>Building Code</option><option>Permit / Inspection</option><option>Municipal Review</option><option>Safety</option><option>Insurance / License</option><option>Custom Requirement</option>' }) + field("Requirement", "requirement", { placeholder: "Describe what must be checked" }) + field("Official source label", "sourceLabel", { required: false }) + field("Official source URL", "sourceUrl", { required: false, type: "url" }) + field("Status", "status", { options: statusOptions(["pending", "reviewed", "complete", "not-applicable"]) }),
    },
    consent: {
      title: "Add consent requirement",
      fields: field("Party type", "partyType", { options: '<option>Client</option><option>Vendor</option><option>Subcontractor</option><option>Team Member</option>' }) + field("Document or consent", "document", { placeholder: "Describe the required consent" }) + field("Version", "version", { placeholder: "Draft v1" }) + field("Status", "status", { options: statusOptions(["draft", "legal-review", "approved", "active", "retired"]) }),
    },
  };
  const config = forms[type];
  if (!config) return;
  dialogTitle.textContent = config.title;
  dialogFields.innerHTML = config.fields;
  dataForm.reset();
  const source = type === "client" ? state.clients.find((item) => item.id === recordId) : type === "project" ? state.projects.find((item) => item.id === recordId) : type === "person" ? state.people.find((item) => item.id === recordId) : type === "estimate" ? state.estimates.find((item) => item.id === recordId) : null;
  if (source) {
    Object.entries(source).forEach(([name, value]) => {
      const control = dataForm.elements.namedItem(name);
      if (control && control.type !== "file" && typeof value !== "object") control.value = value ?? "";
    });
    if (type === "estimate" && source.items?.[0]) {
      ["category", "title", "description", "quantity", "unitPrice"].forEach((name) => {
        const control = dataForm.elements.namedItem(name);
        if (control) control.value = source.items[0][name] ?? (name === "title" ? source.items[0].category : "");
      });
      const subtotal = source.items.reduce((sum, item) => sum + Number(item.quantity) * Number(item.unitPrice), 0);
      const taxable = Math.max(0, subtotal - Number(source.discount || 0));
      const taxPercent = taxable ? Number(source.tax || 0) / taxable * 100 : 0;
      if (dataForm.elements.namedItem("taxPercent")) dataForm.elements.namedItem("taxPercent").value = taxPercent || "";
    }
  }
  dataDialog.showModal();
  const categoryControl = dataForm.elements.namedItem("category");
  categoryControl?.addEventListener("change", () => {
    const savedService = customServiceFromSelection(categoryControl.value);
    if (savedService) {
      const titleControl = dataForm.elements.namedItem("title");
      const descriptionControl = dataForm.elements.namedItem("description");
      const priceControl = dataForm.elements.namedItem("unitPrice");
      if (titleControl) titleControl.value = savedService.title;
      if (descriptionControl) descriptionControl.value = savedService.description;
      if (priceControl) priceControl.value = savedService.unitPrice;
      return;
    }
    if (categoryControl.value === "Custom / New Work") {
      pendingCustomServiceTarget = { mode: "data-form" };
      customServiceForm.reset();
      customServiceDialog.showModal();
    }
  });
}

function nextId(collection, prefix) {
  return `${prefix}-LOCAL-${String(collection.length + 1).padStart(3, "0")}`;
}

function readableDate(value) {
  if (!value) return new Intl.DateTimeFormat("en-US", { dateStyle: "medium" }).format(new Date());
  return new Intl.DateTimeFormat("en-US", { dateStyle: "medium", timeZone: "UTC" }).format(new Date(`${value}T00:00:00Z`));
}

function daysUntilDate(value) {
  if (!value) return null;
  const target = new Date(`${value}T00:00:00`);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (Number.isNaN(target.getTime())) return null;
  return Math.ceil((target.getTime() - today.getTime()) / 86400000);
}

function insuranceStatus(person) {
  if (!person.insuranceExpirationDate || !person.insuranceFileName) return { key: "missing", label: "Missing", days: null };
  const days = daysUntilDate(person.insuranceExpirationDate);
  if (days === null) return { key: "missing", label: "Missing", days: null };
  if (days < 0) return { key: "expired", label: "Expired", days };
  if (days <= 30) return { key: "expires-30", label: `Expires in ${days} days`, days };
  if (days <= 60) return { key: "expires-60", label: `Expires in ${days} days`, days };
  return { key: "active", label: "Active", days };
}

function insuranceStatusClass(key) {
  return key === "active" ? "green" : key === "expired" || key === "missing" ? "red" : "amber";
}

function projectReceiptTotal(projectId) {
  return state.expenseReceipts.filter((item) => item.projectId === projectId).reduce((sum, item) => sum + Number(item.amount || 0), 0);
}

function projectActualCost(project) {
  return Number(project.cost || 0) + projectReceiptTotal(project.id);
}

function prepareInsuranceRenewal(personId) {
  const person = state.people.find((item) => item.id === personId);
  if (!person) return;
  const recipient = person.email || person.companyPhone || person.personalPhone;
  if (!recipient) {
    window.alert("Add a company email or phone number before preparing a renewal request.");
    return;
  }
  const status = insuranceStatus(person);
  const due = person.insuranceExpirationDate ? readableDate(person.insuranceExpirationDate) : "not recorded";
  const channel = person.email ? "email" : "message";
  if (!window.confirm(`Prepare an insurance renewal ${channel} for ${person.name}? Nothing will be sent until you confirm it in your ${channel} app.`)) return;
  const subject = "Updated Certificate of Insurance requested — No Limit Carpentry";
  const body = `Hello ${person.contactName || person.name},\n\nOur records show that your Certificate of Insurance is ${status.key === "missing" ? "not currently on file" : `due to expire on ${due}`}. Please send an updated certificate showing the required coverage for our records.\n\nThank you,\nNo Limit Carpentry\n(848) 466-3339`;
  state.insuranceRenewals.unshift({ id: nextId(state.insuranceRenewals, "REN"), personId: person.id, personName: person.name, preparedAt: new Date().toISOString(), channel, expirationDate: person.insuranceExpirationDate || "", status: "prepared" });
  savePreviewState();
  window.location.href = person.email
    ? `mailto:${encodeURIComponent(person.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
    : `sms:${String(recipient).replace(/[^+\d]/g, "")}?&body=${encodeURIComponent(`${subject}. ${body}`)}`;
}

function betaConfig() {
  const config = window.NO_LIMIT_SUPABASE || {};
  const supabaseUrl = String(config.url || "").trim().replace(/\/$/, "");
  const supabaseAnonKey = String(config.publishableKey || "").trim();
  const organizationId = String(config.organizationId || "").trim();
  return supabaseUrl && supabaseAnonKey && organizationId ? { supabaseUrl, supabaseAnonKey, organizationId } : null;
}

function betaEndpoint() {
  const config = betaConfig();
  return config ? `${config.supabaseUrl}/rest/v1/beta_workspaces` : "";
}

function normalizedSharedState(candidate = {}) {
  const base = cloneDefaultState();
  const merged = { ...base, ...(candidate && typeof candidate === "object" ? candidate : {}) };
  Object.keys(base).forEach((key) => {
    if (Array.isArray(base[key]) && !Array.isArray(merged[key])) merged[key] = base[key];
  });
  merged.people = merged.people.map((person) => ({
    w9Status: "missing",
    w9ReceivedDate: "",
    w9FileName: "",
    insuranceCompany: "",
    insuranceType: "General Liability",
    policyNumber: "",
    coverageAmount: 0,
    insuranceEffectiveDate: "",
    insuranceExpirationDate: "",
    insuranceFileName: "",
    renewalNoticeDays: "60,30",
    ...person,
  }));
  return merged;
}

function betaRequestHeaders(extra = {}) {
  const config = betaConfig();
  return {
    apikey: config?.supabaseAnonKey || "",
    Authorization: `Bearer ${currentAuthSession?.access_token || ""}`,
    ...extra,
  };
}

function setSyncStatus(label, mode = "ready") {
  if (!syncBadge) return;
  syncBadge.textContent = label;
  syncBadge.className = `sync-badge${mode === "saving" ? " is-saving" : mode === "error" ? " is-error" : ""}`;
}

async function pullBetaWorkspace({ quiet = false } = {}) {
  const endpoint = betaEndpoint();
  if (!endpoint) {
    if (!quiet) setSyncStatus("Sync unavailable", "error");
    return false;
  }
  try {
    if (!quiet) setSyncStatus("Loading shared test…", "saving");
    const query = new URLSearchParams({ select: "payload,updated_at", organization_id: `eq.${betaConfig().organizationId}`, id: `eq.${betaWorkspaceId}`, limit: "1" });
    const response = await fetch(`${endpoint}?${query.toString()}`, { headers: betaRequestHeaders(), cache: "no-store" });
    if (!response.ok) throw new Error(`Shared beta returned HTTP ${response.status}`);
    const rows = await response.json();
    const remoteState = rows?.[0]?.payload?.state;
    if (remoteState) {
      state = normalizedSharedState(remoteState);
      localStorage.setItem(previewStorageKey, JSON.stringify(state));
    } else if (currentBetaUser?.role === "admin") {
      await pushBetaWorkspace();
    }
    setSyncStatus("Shared test synced");
    return true;
  } catch (error) {
    console.error("Could not load shared beta workspace", error);
    setSyncStatus("Offline test copy", "error");
    return false;
  }
}

async function pushBetaWorkspace() {
  const endpoint = betaEndpoint();
  if (!endpoint || !currentBetaUser) return false;
  try {
    setSyncStatus("Saving…", "saving");
    const now = new Date().toISOString();
    const record = {
      organization_id: betaConfig().organizationId,
      id: betaWorkspaceId,
      payload: { state, betaOnly: true, updatedAt: now, updatedBy: currentBetaUser.name, updatedByRole: currentBetaUser.role },
      updated_by: currentBetaUser.id,
      updated_at: now,
    };
    const response = await fetch(`${endpoint}?on_conflict=organization_id,id`, {
      method: "POST",
      headers: betaRequestHeaders({ "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }),
      body: JSON.stringify([record]),
      cache: "no-store",
    });
    if (!response.ok) throw new Error(`Shared beta save returned HTTP ${response.status}`);
    setSyncStatus("Shared test saved");
    return true;
  } catch (error) {
    console.error("Could not save shared beta workspace", error);
    setSyncStatus("Save pending", "error");
    return false;
  }
}

function routesForBetaRole(role = "admin") {
  if (role === "collaborator") return ["overview", "projects", "schedule", "media", "map"];
  if (role === "vendor") return ["overview", "projects", "materials", "schedule", "team"];
  if (role === "subcontractor") return ["overview", "projects", "schedule", "team", "compliance"];
  return Object.keys(routes);
}

function betaCanCreate(type) {
  const role = currentBetaUser?.role || "admin";
  if (role === "admin") return true;
  if (role === "collaborator") return ["receipt", "schedule"].includes(type);
  if (role === "vendor") return ["material", "receipt"].includes(type);
  if (role === "subcontractor") return ["person", "receipt"].includes(type);
  return false;
}

function applyBetaAccess() {
  const role = currentBetaUser?.role || "admin";
  const allowedRoutes = new Set(routesForBetaRole(role));
  nav.querySelectorAll("a[data-route]").forEach((link) => { link.hidden = !allowedRoutes.has(link.dataset.route); });
  content.querySelectorAll("[data-create]").forEach((button) => { button.hidden = !betaCanCreate(button.dataset.create); });
  if (role !== "admin") {
    content.querySelectorAll("[data-edit-client], [data-edit-project], [data-edit-estimate], [data-convert-estimate]").forEach((button) => { button.hidden = true; });
  }
  const name = currentBetaUser?.name || "Beta tester";
  const initials = name.split(/\s+/).map((part) => part[0]).join("").slice(0, 2).toUpperCase();
  const profileParts = profileButton?.querySelectorAll("span");
  if (profileParts?.[0]) profileParts[0].textContent = initials;
  const strong = profileButton?.querySelector("strong");
  const small = profileButton?.querySelector("small");
  if (strong) strong.textContent = name;
  if (small) small.textContent = `${formatStatus(role)} beta`;
}

async function activateBetaUser(user, session = currentAuthSession) {
  currentBetaUser = user;
  currentAuthSession = session;
  document.body.classList.remove("auth-required");
  await pullBetaWorkspace();
  renderRoute();
}

function mappedBetaRole(role) {
  if (["owner", "admin", "manager", "office"].includes(role)) return "admin";
  if (role === "team_member") return "collaborator";
  if (["vendor", "subcontractor"].includes(role)) return role;
  return "viewer";
}

async function userFromSession(session) {
  const client = window.noLimitSupabaseClient;
  if (!client || !session?.user) return null;
  const [{ data: profile, error: profileError }, { data: membership, error: membershipError }] = await Promise.all([
    client.from("profiles").select("id,email,full_name").eq("id", session.user.id).single(),
    client.from("organization_members").select("role,status").eq("organization_id", betaConfig().organizationId).eq("user_id", session.user.id).eq("status", "active").single(),
  ]);
  if (profileError || membershipError || !membership) return null;
  return {
    id: session.user.id,
    username: profile?.email || session.user.email,
    name: profile?.full_name || session.user.email?.split("@")[0] || "No Limit user",
    role: mappedBetaRole(membership.role),
  };
}

async function startBeta() {
  if (isLocalPreview && new URLSearchParams(location.search).get("local-demo") === "1") {
    currentBetaUser = { id: "local-preview", username: "local", name: "Leandro", role: "admin" };
    document.body.classList.remove("auth-required");
    setSyncStatus("Local preview");
    renderRoute();
    return;
  }
  const config = betaConfig();
  if (!config || !window.supabase?.createClient) {
    authStatus.textContent = "The secure login service is temporarily unavailable.";
    return;
  }
  const authFlowType = new URLSearchParams(location.hash.replace(/^#/, "")).get("type");
  window.noLimitSupabaseClient = window.supabase.createClient(config.supabaseUrl, config.supabaseAnonKey, {
    auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
  });
  const { data: { session } } = await window.noLimitSupabaseClient.auth.getSession();
  if (session) {
    if (["invite", "recovery"].includes(authFlowType)) {
      currentAuthSession = session;
      loginForm.hidden = true;
      setPasswordForm.hidden = false;
      resetPasswordButton.hidden = true;
      authStatus.textContent = "Choose a password for your No Limit account.";
      return;
    }
    const user = await userFromSession(session);
    if (user) await activateBetaUser(user, session);
    else authStatus.textContent = "Your account exists, but it has not been authorized for No Limit yet.";
  }
}

function saveDataEntry(formData) {
  if (pendingRecordType === "request") {
    const client = state.clients.find((item) => item.id === formData.get("clientId"));
    state.requests.unshift({ id: nextId(state.requests, "REQ"), clientId: client?.id || "", clientName: client?.name || formData.get("clientName"), service: formData.get("service"), submitted: readableDate(), status: formData.get("status"), nextAction: formData.get("nextAction") });
  }
  if (pendingRecordType === "siteVisit") {
    const client = state.clients.find((item) => item.id === formData.get("clientId"));
    state.siteVisits.unshift({ id: nextId(state.siteVisits, "VIS"), requestId: formData.get("requestId"), clientId: client.id, clientName: client.name, projectId: formData.get("projectId"), visitDate: readableDate(formData.get("visitDate")), assignedTo: formData.get("assignedTo"), status: formData.get("status"), notes: formData.get("notes") });
  }
  if (pendingRecordType === "client") {
    const record = { id: pendingRecordId || nextId(state.clients, "CL"), name: formData.get("name"), type: formData.get("type"), contactName: formData.get("contactName"), personalPhone: formData.get("personalPhone"), companyPhone: formData.get("companyPhone"), email: formData.get("email"), website: formData.get("website"), street: formData.get("street"), city: formData.get("city"), state: formData.get("state"), postalCode: formData.get("postalCode"), billingStreet: formData.get("billingStreet"), billingCity: formData.get("billingCity"), billingState: formData.get("billingState"), billingPostalCode: formData.get("billingPostalCode"), projectIds: state.clients.find((item) => item.id === pendingRecordId)?.projectIds || [] };
    if (pendingRecordId) state.clients = state.clients.map((item) => item.id === pendingRecordId ? record : item);
    else state.clients.unshift(record);
  }
  if (pendingRecordType === "estimate") {
    const client = state.clients.find((item) => item.id === formData.get("clientId"));
    const quantity = Number(formData.get("quantity"));
    const savedService = customServiceFromSelection(formData.get("category"));
    const category = savedService ? "Custom / New Work" : formData.get("category");
    const unitPrice = Number(formData.get("unitPrice") || savedService?.unitPrice || 0);
    const discount = Number(formData.get("discount"));
    const taxPercent = Number(formData.get("taxPercent"));
    const taxable = Math.max(0, quantity * unitPrice - discount);
    const existing = state.estimates.find((item) => item.id === pendingRecordId);
    const record = { id: pendingRecordId || nextId(state.estimates, "EST"), requestId: formData.get("requestId"), clientId: client.id, clientName: client.name, projectId: formData.get("projectId"), issueDate: readableDate(formData.get("issueDate")), status: formData.get("status"), revision: existing ? Number(existing.revision || 1) + 1 : 1, validUntil: readableDate(formData.get("validUntil")), discount, tax: taxable * taxPercent / 100, items: [{ category, customServiceId: savedService?.id || "", title: formData.get("title") || savedService?.title || category, description: formData.get("description") || savedService?.description || "", quantity, unitPrice }], total: taxable * (1 + taxPercent / 100), contractId: existing?.contractId || "", notes: formData.get("notes") };
    if (pendingRecordId) state.estimates = state.estimates.map((item) => item.id === pendingRecordId ? record : item);
    else state.estimates.unshift(record);
  }
  if (pendingRecordType === "project") {
    const client = state.clients.find((item) => item.id === formData.get("clientId"));
    const project = { id: pendingRecordId || nextId(state.projects, "PR"), name: formData.get("name"), clientId: client.id, clientName: client.name, status: formData.get("status"), service: formData.get("service"), siteStreet: formData.get("siteStreet"), siteCity: formData.get("siteCity"), siteState: formData.get("siteState"), sitePostalCode: formData.get("sitePostalCode"), startDate: readableDate(formData.get("startDate")), progress: Number(formData.get("progress")), contractValue: Number(formData.get("contractValue")), cost: Number(formData.get("cost")), outstanding: Number(formData.get("outstanding")), managerId: state.projects.find((item) => item.id === pendingRecordId)?.managerId || "" };
    if (pendingRecordId) state.projects = state.projects.map((item) => item.id === pendingRecordId ? project : item);
    else state.projects.unshift(project);
    state.clients.forEach((item) => { item.projectIds = (item.projectIds || []).filter((id) => id !== project.id); });
    client.projectIds = [...new Set([...(client.projectIds || []), project.id])];
  }
  if (pendingRecordType === "transaction") {
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    state.transactions.unshift({ id: nextId(state.transactions, "TX"), projectId: project.id, projectName: project.name, date: readableDate(), type: formData.get("type"), category: formData.get("category"), party: formData.get("party"), amount: Number(formData.get("amount")), status: formData.get("status") });
  }
  if (pendingRecordType === "invoiceItem") {
    const invoice = state.invoices.find((item) => item.id === formData.get("invoiceId"));
    const savedService = customServiceFromSelection(formData.get("category"));
    const category = savedService ? "Custom / New Work" : formData.get("category");
    invoice.items.push({ category, customServiceId: savedService?.id || "", title: formData.get("title") || savedService?.title || category, description: formData.get("description") || savedService?.description || "", quantity: Number(formData.get("quantity")), unitPrice: Number(formData.get("unitPrice") || savedService?.unitPrice || 0) });
    invoice.total = invoice.items.reduce((sum, item) => sum + Number(item.quantity) * Number(item.unitPrice), 0);
    invoice.balance = Math.max(0, invoice.total - Number(invoice.paid || 0));
  }
  if (pendingRecordType === "payment") {
    const invoice = state.invoices.find((item) => item.id === formData.get("invoiceId"));
    const amount = Number(formData.get("amount"));
    state.payments.unshift({ id: nextId(state.payments, "PAY"), invoiceId: invoice.id, projectId: invoice.projectId, date: readableDate(formData.get("date")), method: formData.get("method"), checkNumber: formData.get("checkNumber"), milestone: formData.get("milestone"), amount, status: formData.get("status") });
    if (["received", "cleared"].includes(formData.get("status"))) invoice.paid = Number(invoice.paid || 0) + amount;
    invoice.balance = Math.max(0, invoice.total - invoice.paid);
    invoice.status = invoice.balance === 0 ? "paid" : invoice.paid > 0 ? "partially-paid" : "open";
  }
  if (pendingRecordType === "changeOrder") {
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    const savedService = customServiceFromSelection(formData.get("category"));
    const category = savedService ? "Custom / New Work" : formData.get("category");
    const quantity = Number(formData.get("quantity"));
    const unitPrice = Number(formData.get("unitPrice") || savedService?.unitPrice || 0);
    const amount = quantity * unitPrice;
    state.changeOrders.unshift({ id: nextId(state.changeOrders, "CO"), projectId: project.id, projectName: project.name, status: formData.get("status"), category, customServiceId: savedService?.id || "", title: formData.get("title") || savedService?.title || category, description: formData.get("description") || savedService?.description || "", quantity, unitPrice, amount });
  }
  if (pendingRecordType === "person") {
    const projectId = formData.get("projectId");
    const existing = state.people.find((item) => item.id === pendingRecordId);
    const w9File = formData.get("w9File");
    const insuranceFile = formData.get("insuranceFile");
    const record = { id: pendingRecordId || nextId(state.people, "PE"), name: formData.get("name"), type: formData.get("type"), role: formData.get("role"), contactName: formData.get("contactName"), personalPhone: formData.get("personalPhone"), companyPhone: formData.get("companyPhone"), email: formData.get("email"), website: formData.get("website"), street: formData.get("street"), city: formData.get("city"), state: formData.get("state"), postalCode: formData.get("postalCode"), projectIds: projectId ? [...new Set([...(existing?.projectIds || []), projectId])] : existing?.projectIds || [], status: formData.get("status"), w9Status: formData.get("w9Status") || "missing", w9ReceivedDate: formData.get("w9ReceivedDate") || "", w9FileName: w9File?.name || existing?.w9FileName || "", insuranceCompany: formData.get("insuranceCompany") || "", insuranceType: formData.get("insuranceType") || "", policyNumber: formData.get("policyNumber") || "", coverageAmount: Number(formData.get("coverageAmount") || 0), insuranceEffectiveDate: formData.get("insuranceEffectiveDate") || "", insuranceExpirationDate: formData.get("insuranceExpirationDate") || "", insuranceFileName: insuranceFile?.name || existing?.insuranceFileName || "", renewalNoticeDays: formData.get("renewalNoticeDays") || "60,30" };
    if (pendingRecordId) state.people = state.people.map((item) => item.id === pendingRecordId ? record : item);
    else state.people.unshift(record);
  }
  if (pendingRecordType === "receipt") {
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    const receiptFile = formData.get("receiptFile");
    state.expenseReceipts.unshift({ id: nextId(state.expenseReceipts, "RCP"), projectId: project.id, projectName: project.name, vendor: formData.get("vendor"), category: formData.get("category"), purchaseDate: formData.get("purchaseDate"), amount: Number(formData.get("amount")), paymentMethod: formData.get("paymentMethod"), reference: formData.get("reference"), receiptFileName: receiptFile?.name || "", notes: formData.get("notes") });
  }
  if (pendingRecordType === "material") {
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    const vendor = state.people.find((item) => item.id === formData.get("vendorId"));
    const quantity = Number(formData.get("quantity"));
    const unitCost = Number(formData.get("unitCost"));
    state.materials.unshift({ id: nextId(state.materials, "MT"), projectId: project.id, projectName: project.name, vendorId: vendor.id, vendor: vendor.name, description: formData.get("description"), quantity, unit: formData.get("unit"), unitCost, amount: quantity * unitCost, reference: formData.get("reference"), purchaseDate: readableDate(formData.get("purchaseDate")), status: formData.get("status") });
  }
  if (pendingRecordType === "schedule") {
    const person = state.people.find((item) => item.id === formData.get("personId"));
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    state.schedule.unshift({ id: nextId(state.schedule, "SCH"), personId: person.id, personName: person.name, personType: person.type, projectId: project.id, projectName: project.name, workDate: readableDate(formData.get("workDate")), shift: formData.get("shift"), instructions: formData.get("instructions"), status: formData.get("status") });
  }
  if (pendingRecordType === "compliance") {
    const project = state.projects.find((item) => item.id === formData.get("projectId"));
    state.compliance.unshift({ id: nextId(state.compliance, "CMP"), projectId: project.id, projectName: project.name, jurisdiction: formData.get("jurisdiction"), municipality: formData.get("municipality"), requirement: formData.get("requirement"), category: formData.get("category"), status: formData.get("status"), sourceLabel: formData.get("sourceLabel"), sourceUrl: formData.get("sourceUrl") });
  }
  if (pendingRecordType === "consent") {
    state.consents.unshift({ id: nextId(state.consents, "CNS"), partyType: formData.get("partyType"), document: formData.get("document"), version: formData.get("version"), status: formData.get("status") });
  }
  savePreviewState();
}

function pageHead(route, actions = "") {
  return `
    <header class="page-head">
      <div>
        <p class="page-kicker">${route.kicker}</p>
        <h1>${route.heading}</h1>
        <p class="page-description">${route.description}</p>
      </div>
      ${actions ? `<div class="button-row">${actions}</div>` : ""}
    </header>`;
}

function emptyState(title, copy) {
  return `
    <div class="empty-state">
      <div>
        <img src="../assets/brand-kit/no-limit-carpentry-nl-monogram.png" alt="" />
        <h3>${title}</h3>
        <p>${copy}</p>
      </div>
    </div>`;
}

function metric(label, value, note, route) {
  return `<a class="metric-card" href="#${route}"><span class="metric-label">${label}</span><strong class="metric-value">${value}</strong><span class="metric-note">${note}</span></a>`;
}

function renderOverview() {
  const activeProjects = state.projects.filter((item) => item.status === "active").length;
  const awaitingContact = state.requests.filter((item) => ["new", "to-contact"].includes(item.status)).length;
  const outstanding = state.transactions.filter((item) => item.type === "receivable" && item.status !== "paid").reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const approvedMedia = state.media.filter((item) => item.publishStatus === "approved").length;

  return `
    <section class="page">
      ${pageHead(routes.overview, '<a class="button" href="#requests">Review new requests</a><a class="button secondary" href="#reports">Build a report</a>')}
      <div class="metric-grid">
        ${metric("Awaiting contact", awaitingContact, "New website requests that still need a response.", "requests")}
        ${metric("Active projects", activeProjects, "Projects currently in execution.", "projects")}
        ${metric("Outstanding", formatCurrency(outstanding), "Open client balances across projects.", "financial")}
        ${metric("Approved media", approvedMedia, "Files cleared for possible website publishing.", "media")}
      </div>
      <div class="content-grid">
        <article class="panel">
          <div class="panel-head"><div><h2>Current priorities</h2><p>The dashboard will rank work that needs attention.</p></div></div>
          <div class="priority-list">
            <a href="#requests"><span class="priority-dot amber"></span><div><strong>Contact Harbor Residence (Demo)</strong><p>Custom Trim request · site visit due Sep 12</p></div><small>Request</small></a>
            <a href="#financial"><span class="priority-dot amber"></span><div><strong>Review open material bill</strong><p>Oak House Millwork · Atlantic Millwork Supply (Demo)</p></div><small>${formatCurrency(5250)}</small></a>
            <a href="#media"><span class="priority-dot green"></span><div><strong>Review project media</strong><p>Two demo uploads are awaiting classification.</p></div><small>Media</small></a>
          </div>
        </article>
        <aside class="panel">
          <div class="panel-head"><div><h2>Connected workflow</h2><p>Each stage opens the next record without retyping information.</p></div></div>
          <div class="workflow">
            ${["Request received", "Contact and site visit", "Proposal approval", "Client and project", "Execution and costs", "Completion and publishing"].map((item, index) => `<div class="workflow-item"><span class="workflow-number">${index + 1}</span><strong>${item}</strong><small>${index < 5 ? "Connected" : "Approval required"}</small></div>`).join("")}
          </div>
        </aside>
      </div>
    </section>`;
}

function renderRequests() {
  return `
    <section class="page">
      ${pageHead(routes.requests, '<button class="button" data-create="request" type="button">New request</button><button class="button secondary" data-create="siteVisit" type="button">Schedule Free On-Site Visit</button>')}
      <article class="panel">
        <div class="filter-bar">
          <label>Search<input type="search" placeholder="Name, phone, email, or service" /></label>
          <label>Status<select><option>All statuses</option><option>New</option><option>To Contact</option><option>Contacted</option><option>Proposal Sent</option></select></label>
          <label>Service category<select><option>All 13 categories</option></select></label>
          <button class="button secondary" type="button">Apply filters</button>
        </div>
        ${demoTable(["Request", "Client", "Service", "Submitted", "Status", "Next action"], state.requests.map((item) => `<tr><td><strong>${escapeHtml(item.id)}</strong></td><td>${escapeHtml(item.clientName)}</td><td>${escapeHtml(item.service)}</td><td>${escapeHtml(item.submitted)}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td><td>${escapeHtml(item.nextAction)}</td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Free On-Site Visits</h2><p>Visits connect the original request, client, future project, assigned person, and notes.</p></div></div>
        ${demoTable(["Visit", "Date", "Client", "Request", "Project", "Assigned to", "Status", "Notes"], state.siteVisits.map((visit) => `<tr><td><strong>${escapeHtml(visit.id)}</strong></td><td>${escapeHtml(visit.visitDate)}</td><td>${escapeHtml(visit.clientName)}</td><td>${escapeHtml(visit.requestId)}</td><td>${escapeHtml(visit.projectId || "Not created")}</td><td>${escapeHtml(visit.assignedTo)}</td><td><span class="status-pill ${statusClass(visit.status)}">${escapeHtml(formatStatus(visit.status))}</span></td><td>${escapeHtml(visit.notes)}</td></tr>`))}
      </article>
    </section>`;
}

function renderDocuments() {
  return `
    <section class="page">
      ${pageHead(routes.documents, '<button class="button" data-create="estimate" type="button">New estimate</button><button class="button secondary" data-create="changeOrder" type="button">New work / change order</button>')}
      <article class="panel">
        <div class="panel-head"><div><h2>Proposals and revisions</h2><p>Save drafts, revise quantities and prices, preview the branded document, then approve it to create the contract and invoice.</p></div></div>
        ${demoTable(["Proposal", "Request", "Client", "Revision", "Valid until", "Total", "Status", "Document", "Next step"], state.estimates.map((estimate) => `<tr><td><strong>${escapeHtml(estimate.id)}</strong><small class="record-id">${estimate.items.length} line item${estimate.items.length === 1 ? "" : "s"}</small></td><td>${escapeHtml(estimate.requestId)}</td><td><a href="#clients">${escapeHtml(estimate.clientName)}</a></td><td>R${estimate.revision}</td><td>${escapeHtml(estimate.validUntil)}</td><td>${formatCurrency(estimate.total)}</td><td><span class="status-pill ${statusClass(estimate.status)}">${escapeHtml(formatStatus(estimate.status))}</span></td><td><button class="text-button" data-view-document="proposal" data-document-id="${escapeHtml(estimate.id)}" type="button">Preview / send</button><small class="record-id"><button class="text-button" data-edit-estimate="${escapeHtml(estimate.id)}" type="button">Edit details</button></small></td><td>${estimate.contractId ? `<a href="#documents">${escapeHtml(estimate.contractId)}</a>` : `<button class="text-button" data-convert-estimate="${escapeHtml(estimate.id)}" type="button">Approve & generate</button>`}</td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Contracts</h2><p>Every contract stays linked to its estimate, client, project, and invoice.</p></div></div>
        ${demoTable(["Contract", "Estimate", "Client", "Project", "Signed", "Value", "Status", "Invoice"], state.contracts.map((contract) => `<tr><td><strong>${escapeHtml(contract.id)}</strong></td><td>${escapeHtml(contract.estimateId)}</td><td><a href="#clients">${escapeHtml(contract.clientName)}</a></td><td><a href="#projects">${escapeHtml(contract.projectId)}</a></td><td>${escapeHtml(contract.signedDate)}</td><td>${formatCurrency(contract.value)}</td><td><span class="status-pill ${statusClass(contract.status)}">${escapeHtml(formatStatus(contract.status))}</span></td><td><a href="#financial">${escapeHtml(contract.invoiceId)}</a></td></tr>`))}
      </article>
    </section>`;
}

function renderClients() {
  return `
    <section class="page">
      ${pageHead(routes.clients, '<button class="button" data-create="client" type="button">Add client</button>')}
      <div class="module-grid">
        <article class="module-card"><span class="initial">CL</span><h3>Client directory</h3><p>Company, contacts, addresses, notes, documents, and full relationship history.</p><a href="#clients">Open directory</a></article>
        <article class="module-card"><span class="initial">PR</span><h3>Projects by client</h3><p>Review every planned, active, paused, and completed job for one client.</p><a href="#projects">Open projects</a></article>
        <article class="module-card"><span class="initial">CT</span><h3>Contracts</h3><p>Keep primary contracts and approved change orders connected to the correct project.</p><a href="#financial">Open financial records</a></article>
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Connected client records</h2><p>Each client shows the projects linked to the same identifier.</p></div></div>
        ${demoTable(["Client", "Type", "Client address", "Primary contact", "Linked projects", "Contract value", ""], state.clients.map((client) => {
          const projects = state.projects.filter((project) => project.clientId === client.id);
          const total = projects.reduce((sum, project) => sum + project.contractValue, 0);
          return `<tr><td><strong>${escapeHtml(client.name)}</strong><small class="record-id">${escapeHtml(client.id)}</small></td><td>${escapeHtml(client.type)}</td><td>${escapeHtml(clientLocation(client))}</td><td>${escapeHtml(clientContact(client))}</td><td><a href="#projects">${projects.length} project${projects.length === 1 ? "" : "s"}</a></td><td>${formatCurrency(total)}</td><td><button class="text-button" type="button" data-edit-client="${escapeHtml(client.id)}">Edit</button></td></tr>`;
        }))}
      </article>
    </section>`;
}

function renderProjects() {
  return `
    <section class="page">
      ${pageHead(routes.projects, '<button class="button" data-create="project" type="button">Add project</button><button class="button secondary" data-create="receipt" type="button">Add receipt / expense</button><a class="button secondary" href="#map">Open map</a>')}
      <article class="panel">
        <div class="filter-bar">
          <label>Search<input type="search" placeholder="Project, client, or address" /></label>
          <label>Status<select><option>All statuses</option><option>Planned</option><option>Active</option><option>Paused</option><option>Completed</option></select></label>
          <label>Client<select><option>All clients</option>${clientOptions()}</select></label>
          <button class="button secondary" type="button">Apply filters</button>
        </div>
        ${demoTable(["Project", "Client", "Service", "Project address", "Progress", "Financial", "Status", ""], state.projects.map((project) => `<tr><td><strong>${escapeHtml(project.name)}</strong><small class="record-id">${escapeHtml(project.id)}</small></td><td><a href="#clients">${escapeHtml(project.clientName)}</a></td><td>${escapeHtml(project.service)}</td><td>${escapeHtml(projectLocation(project))}</td><td><div class="progress"><span style="width:${Math.max(0, Math.min(100, project.progress))}%"></span></div><small>${project.progress}%</small></td><td>${formatCurrency(project.contractValue)}<small class="record-id">Base cost ${formatCurrency(project.cost)}</small><small class="record-id">Receipts ${formatCurrency(projectReceiptTotal(project.id))}</small><strong class="record-id">Actual cost ${formatCurrency(projectActualCost(project))}</strong></td><td><span class="status-pill ${statusClass(project.status)}">${escapeHtml(formatStatus(project.status))}</span></td><td><button class="text-button" type="button" data-edit-project="${escapeHtml(project.id)}">Edit</button></td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Additional costs and receipts</h2><p>Every receipt is connected to a project and included in its actual cost.</p></div><button class="button secondary" data-create="receipt" type="button">Add receipt / expense</button></div>
        ${demoTable(["Receipt", "Project", "Vendor / store", "Category", "Date", "Amount", "Payment", "Attachment"], state.expenseReceipts.map((item) => `<tr><td><strong>${escapeHtml(item.id)}</strong><small class="record-id">${escapeHtml(item.reference || "No reference")}</small></td><td>${escapeHtml(item.projectName)}</td><td>${escapeHtml(item.vendor)}</td><td>${escapeHtml(item.category)}</td><td>${escapeHtml(readableDate(item.purchaseDate))}</td><td>${formatCurrency(item.amount)}</td><td>${escapeHtml(item.paymentMethod)}</td><td>${escapeHtml(item.receiptFileName || "Not attached")}</td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>New work and change orders</h2><p>Additional scope remains separate from the original contract until approved.</p></div><button class="button secondary" data-create="changeOrder" type="button">Add new work</button></div>
        ${demoTable(["Change order", "Project", "Category", "Description", "Amount", "Status"], state.changeOrders.map((item) => `<tr><td><strong>${escapeHtml(item.id)}</strong></td><td>${escapeHtml(item.projectName)}</td><td>${escapeHtml(item.category || "Custom / New Work")}</td><td>${escapeHtml(item.description)}</td><td>${formatCurrency(item.amount)}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
      </article>
    </section>`;
}

function renderFinancial() {
  const contracted = state.projects.reduce((sum, item) => sum + Number(item.contractValue || 0), 0);
  const received = state.transactions.filter((item) => item.type === "receivable" && item.status === "paid").reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const actualCost = state.projects.reduce((sum, item) => sum + projectActualCost(item), 0);
  const outstanding = state.transactions.filter((item) => item.type === "receivable" && item.status !== "paid").reduce((sum, item) => sum + Number(item.amount || 0), 0);
  return `
    <section class="page">
      ${pageHead(routes.financial, '<button class="button" data-create="payment" type="button">Record payment</button><button class="button secondary" data-create="invoiceItem" type="button">Add invoice item</button><button class="button secondary" data-create="transaction" type="button">Add expense</button>')}
      <div class="metric-grid">
        ${metric("Contracted value", formatCurrency(contracted), "Total value of current project contracts.", "financial")}
        ${metric("Received", formatCurrency(received), "Client payments received.", "financial")}
        ${metric("Actual cost", formatCurrency(actualCost), "Materials, labor, subcontractors, and expenses.", "financial")}
        ${metric("Outstanding", formatCurrency(outstanding), "Open client balances.", "financial")}
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Invoices and payment stages</h2><p>Invoices are generated from contracts and remain itemized by service.</p></div><a class="button secondary" href="#reports">Invoice report</a></div>
        ${demoTable(["Invoice", "Contract", "Client", "Project", "Items", "Total", "Paid", "Balance", "Status", "Document"], state.invoices.map((invoice) => `<tr><td><strong>${escapeHtml(invoice.id)}</strong><small class="record-id">Issued ${escapeHtml(invoice.issueDate)}</small></td><td><a href="#documents">${escapeHtml(invoice.contractId)}</a></td><td>${escapeHtml(invoice.clientName)}</td><td><a href="#projects">${escapeHtml(invoice.projectId || "Pending project")}</a></td><td>${invoice.items.length}</td><td>${formatCurrency(invoice.total)}</td><td>${formatCurrency(invoice.paid)}</td><td>${formatCurrency(invoice.balance)}</td><td><span class="status-pill ${statusClass(invoice.status)}">${escapeHtml(formatStatus(invoice.status))}</span></td><td><button class="text-button" data-view-document="invoice" data-document-id="${escapeHtml(invoice.id)}" type="button">Preview / send</button></td></tr>`))}
        <h3 class="subsection-title">Invoice line items</h3>
        ${demoTable(["Invoice", "Service category", "Description", "Quantity", "Unit price", "Line total"], state.invoices.flatMap((invoice) => invoice.items.map((item) => `<tr><td>${escapeHtml(invoice.id)}</td><td>${escapeHtml(item.category)}</td><td>${escapeHtml(item.description)}</td><td>${Number(item.quantity).toLocaleString("en-US")}</td><td>${formatCurrency(item.unitPrice)}</td><td>${formatCurrency(Number(item.quantity) * Number(item.unitPrice))}</td></tr>`)))}
        <h3 class="subsection-title">Payment milestones</h3>
        ${demoTable(["Invoice", "Milestone", "Percent", "Amount", "Status"], state.invoices.flatMap((invoice) => invoice.schedule.map((stage) => `<tr><td>${escapeHtml(invoice.id)}</td><td>${escapeHtml(stage.label)}</td><td>${Number(stage.percent).toLocaleString("en-US")}%</td><td>${formatCurrency(stage.amount)}</td><td><span class="status-pill ${statusClass(stage.status)}">${escapeHtml(formatStatus(stage.status))}</span></td></tr>`)))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Deposits, checks, and payments</h2><p>Each payment is tied to an invoice, project milestone, and reference number.</p></div></div>
        ${demoTable(["Payment", "Invoice", "Date", "Method", "Check / reference", "Milestone", "Amount", "Status"], state.payments.map((payment) => `<tr><td><strong>${escapeHtml(payment.id)}</strong></td><td>${escapeHtml(payment.invoiceId)}</td><td>${escapeHtml(payment.date)}</td><td>${escapeHtml(payment.method)}</td><td>${escapeHtml(payment.checkNumber || "—")}</td><td>${escapeHtml(payment.milestone)}</td><td>${formatCurrency(payment.amount)}</td><td><span class="status-pill ${statusClass(payment.status)}">${escapeHtml(formatStatus(payment.status))}</span></td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Transactions by project</h2><p>Every amount is connected to a project and related party.</p></div></div>
        ${demoTable(["Date", "Transaction", "Project", "Category", "Related party", "Amount", "Status"], state.transactions.map((item) => `<tr><td>${escapeHtml(item.date)}</td><td><strong>${escapeHtml(item.id)}</strong></td><td><a href="#projects">${escapeHtml(item.projectName)}</a></td><td>${escapeHtml(item.category)}</td><td>${escapeHtml(item.party)}</td><td>${formatCurrency(item.amount)}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
      </article>
    </section>`;
}

function renderMaterials() {
  const total = state.materials.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  return `
    <section class="page">
      ${pageHead(routes.materials, '<button class="button" data-create="material" type="button">Add material</button><button class="button secondary" data-create="person" type="button">Add Vendor</button>')}
      <div class="metric-grid">
        ${metric("Material records", state.materials.length, "Quotes, orders, receipts, and installed items.", "materials")}
        ${metric("Committed value", formatCurrency(total), "Current demo material value across projects.", "materials")}
        ${metric("Vendors", state.people.filter((person) => person.type === "Vendor").length, "Registered material and service Vendors.", "team")}
        ${metric("Awaiting receipt", state.materials.filter((item) => item.status === "ordered").length, "Orders that have not been marked received.", "materials")}
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Materials by project</h2><p>Every line connects quantity, unit cost, Vendor, purchasing reference, and project.</p></div></div>
        ${demoTable(["Material", "Project", "Vendor", "Quantity", "Unit cost", "Total", "PO / receipt / quote", "Date", "Status"], state.materials.map((item) => `<tr><td><strong>${escapeHtml(item.description)}</strong><small class="record-id">${escapeHtml(item.id)}</small></td><td><a href="#projects">${escapeHtml(item.projectName)}</a></td><td><a href="#team">${escapeHtml(item.vendor || "Not entered")}</a></td><td>${Number(item.quantity || 0).toLocaleString("en-US")} ${escapeHtml(item.unit || "")}</td><td>${formatCurrency(item.unitCost || 0)}</td><td>${formatCurrency(item.amount)}</td><td>${escapeHtml(item.reference || "—")}</td><td>${escapeHtml(item.purchaseDate || "—")}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
      </article>
    </section>`;
}

function renderTeam() {
  const subcontractors = state.people.filter((person) => person.type === "Subcontractor");
  const insuranceCounts = subcontractors.reduce((counts, person) => {
    const key = insuranceStatus(person).key;
    if (key === "active") counts.active += 1;
    else if (["expires-30", "expires-60"].includes(key)) counts.expiring += 1;
    else counts.action += 1;
    return counts;
  }, { active: 0, expiring: 0, action: 0 });
  return `
    <section class="page">
      ${pageHead(routes.team, '<button class="button" data-create="person" type="button">Add person or company</button>')}
      <div class="metric-grid">
        ${metric("Subcontractors", subcontractors.length, "Companies tracked for projects and compliance.", "team")}
        ${metric("Insurance active", insuranceCounts.active, "Certificates valid for more than 60 days.", "team")}
        ${metric("Renewal approaching", insuranceCounts.expiring, "Certificates inside the 60-day renewal window.", "team")}
        ${metric("Compliance action", insuranceCounts.action, "Missing or expired insurance records.", "team")}
      </div>
      <div class="module-grid">
        <article class="module-card"><span class="initial">TM</span><h3>Team Members</h3><p>Roles, assignments, time records, documents, and payments.</p><a href="#team">Open team members</a></article>
        <article class="module-card"><span class="initial">VE</span><h3>Vendors</h3><p>Contacts, supplied products, purchases, expenses, and project history.</p><a href="#team">Open vendors</a></article>
        <article class="module-card"><span class="initial">SC</span><h3>Subcontractors</h3><p>Specialties, contracts, compliance documents, projects, and payments.</p><a href="#team">Open subcontractors</a></article>
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Assignments</h2><p>Team members, vendors, and subcontractors remain distinct but connect to the same projects.</p></div></div>
        ${demoTable(["Person or company", "Type", "Role / specialty", "Primary contact", "Address", "Linked projects", "Status", ""], state.people.map((person) => `<tr><td><strong>${escapeHtml(person.name)}</strong><small class="record-id">${escapeHtml(person.id)}</small></td><td>${escapeHtml(person.type)}</td><td>${escapeHtml(person.role)}</td><td>${escapeHtml(partyContact(person))}</td><td>${escapeHtml(partyLocation(person))}</td><td>${person.projectIds.map((id) => `<a href="#projects">${escapeHtml(id)}</a>`).join(" · ")}</td><td><span class="status-pill ${statusClass(person.status)}">${escapeHtml(formatStatus(person.status))}</span></td><td><button class="text-button" type="button" data-edit-person="${escapeHtml(person.id)}">Edit</button></td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Subcontractor documents and renewals</h2><p>W-9 and insurance records show their current status automatically. Renewal requests are prepared at 60 and 30 days and require confirmation before sending.</p></div></div>
        ${demoTable(["Subcontractor", "W-9", "Insurance", "Policy / coverage", "Expires", "Renewal activity", ""], subcontractors.map((person) => {
          const insurance = insuranceStatus(person);
          const lastRenewal = state.insuranceRenewals.find((item) => item.personId === person.id);
          return `<tr><td><strong>${escapeHtml(person.name)}</strong><small class="record-id">${escapeHtml(person.contactName || "No contact")}</small></td><td><span class="status-pill ${person.w9Status === "verified" ? "green" : "amber"}">${escapeHtml(formatStatus(person.w9Status || "missing"))}</span><small class="record-id">${escapeHtml(person.w9FileName || "No file")}</small></td><td><span class="status-pill ${insuranceStatusClass(insurance.key)}">${escapeHtml(insurance.label)}</span><small class="record-id">${escapeHtml(person.insuranceCompany || "No carrier")}</small></td><td>${escapeHtml(person.policyNumber || "—")}<small class="record-id">${person.coverageAmount ? formatCurrency(person.coverageAmount) : "No coverage entered"}</small></td><td>${person.insuranceExpirationDate ? escapeHtml(readableDate(person.insuranceExpirationDate)) : "—"}<small class="record-id">${escapeHtml(person.insuranceFileName || "No certificate")}</small></td><td>${lastRenewal ? `<span class="status-pill amber">Prepared</span><small class="record-id">${escapeHtml(new Intl.DateTimeFormat("en-US", { dateStyle: "medium" }).format(new Date(lastRenewal.preparedAt)))}</small>` : "No request prepared"}</td><td><button class="text-button" type="button" data-renewal-person="${escapeHtml(person.id)}">Prepare renewal</button><small class="record-id"><button class="text-button" type="button" data-edit-person="${escapeHtml(person.id)}">Edit record</button></small></td></tr>`;
        }))}
      </article>
    </section>`;
}

function renderSchedule() {
  return `
    <section class="page">
      ${pageHead(routes.schedule, '<button class="button" data-create="schedule" type="button">Add assignment</button>')}
      <div class="metric-grid">
        ${metric("Scheduled", state.schedule.filter((item) => item.status === "scheduled").length, "Assignments awaiting completion.", "schedule")}
        ${metric("Confirmed", state.schedule.filter((item) => item.status === "confirmed").length, "People who confirmed their next workday.", "schedule")}
        ${metric("On site", state.attendance.filter((item) => item.locationStatus === "within-project-area").length, "Location-authorized work check-ins.", "map")}
        ${metric("Completed", state.schedule.filter((item) => item.status === "completed").length, "Assignments completed in this preview.", "schedule")}
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Individual work assignments</h2><p>Future login permissions will show each person only their own schedule.</p></div></div>
        ${demoTable(["Date", "Team Member / Vendor / Subcontractor", "Type", "Project", "Shift", "Instructions", "Status"], state.schedule.map((item) => `<tr><td>${escapeHtml(item.workDate)}</td><td><strong>${escapeHtml(item.personName)}</strong></td><td>${escapeHtml(item.personType)}</td><td><a href="#projects">${escapeHtml(item.projectName)}</a></td><td>${escapeHtml(item.shift)}</td><td>${escapeHtml(item.instructions)}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
      </article>
      <article class="panel"><div class="panel-head"><div><h2>Private account behavior</h2><p>Planned for the isolated authentication phase.</p></div></div><div class="security-list"><div class="security-row"><div><strong>Individual login</strong><p>Every authorized person or company receives a separate account.</p></div><span class="status-pill amber">Planned</span></div><div class="security-row"><div><strong>Schedule access window</strong><p>Daily access is available from ${scheduleAccessWindow.starts} to ${scheduleAccessWindow.ends}.</p></div><span class="status-pill green">Defined</span></div><div class="security-row"><div><strong>Schedule-only access</strong><p>Users see their own assignments, project instructions, and permitted documents.</p></div><span class="status-pill amber">Planned</span></div><div class="security-row"><div><strong>Location-authorized check-in</strong><p>Opening the schedule does not capture a home location. Verification begins only after an explicit Start Workday / Check In action and ends at Check Out or ${scheduleAccessWindow.ends}.</p></div><span class="status-pill amber">Planned</span></div></div></article>
    </section>`;
}

function renderMedia() {
  const countByStatus = (status) => state.media.filter((item) => item.publishStatus === status).length;
  return `
    <section class="page">
      ${pageHead(routes.media, '<button class="button" type="button">Upload files</button>')}
      <div class="metric-grid">
        ${metric("Internal", countByStatus("internal"), "Private files visible only to authorized users.", "media")}
        ${metric("Awaiting review", countByStatus("awaiting-review"), "Uploads that need approval or classification.", "media")}
        ${metric("Approved", countByStatus("approved"), "Files approved for possible public use.", "media")}
        ${metric("Published", countByStatus("published"), "Files currently visible on the website.", "media")}
      </div>
      <article class="panel">
        <div class="panel-head"><div><h2>Media index</h2><p>Demo metadata only; no image files were copied or published.</p></div></div>
        ${demoTable(["Media", "Project", "Phase", "Type", "Date", "Publishing status"], state.media.map((item) => `<tr><td><strong>${escapeHtml(item.id)}</strong></td><td><a href="#projects">${escapeHtml(item.projectName)}</a></td><td>${escapeHtml(item.phase)}</td><td>${escapeHtml(item.fileType)}</td><td>${escapeHtml(item.date)}</td><td><span class="status-pill ${statusClass(item.publishStatus)}">${escapeHtml(formatStatus(item.publishStatus))}</span></td></tr>`))}
      </article>
    </section>`;
}

function renderMap() {
  return `
    <section class="page">
      ${pageHead(routes.map, '<a class="button secondary" href="#projects">View project list</a>')}
      <article class="panel">
        <div class="filter-bar">
          <label>Project status<select><option>All statuses</option><option>Planned</option><option>Active</option><option>Paused</option><option>Completed</option></select></label>
          <label>Client<select><option>All clients</option></select></label>
          <label>Date range<select><option>All dates</option><option>This month</option><option>This year</option></select></label>
          <button class="button secondary" type="button">Apply filters</button>
        </div>
        <div class="location-grid">
          ${state.projects.map((project) => `<a class="location-card" href="#projects"><span class="map-pin">${project.status === "active" ? "A" : project.status === "planned" ? "P" : "C"}</span><div><strong>${escapeHtml(project.name)}</strong><p>${escapeHtml(projectLocation(project))} · ${escapeHtml(project.service)}</p></div><span class="status-pill ${statusClass(project.status)}">${escapeHtml(formatStatus(project.status))}</span></a>`).join("")}
        </div>
        <p class="privacy-note">Only city-level demo locations are shown in this preview. Exact job-site addresses will require permission.</p>
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Authorized work check-ins</h2><p>Location is verified only after Start Workday / Check In and only during the defined work window.</p></div></div>
        ${demoTable(["Person", "Assigned project", "Check-in", "Verification", "Consent", "Automatic end"], state.attendance.map((item) => `<tr><td><strong>${escapeHtml(item.personName)}</strong></td><td>${escapeHtml(item.projectName)}</td><td>${escapeHtml(item.checkedInAt)}</td><td><span class="status-pill ${item.locationStatus === "within-project-area" ? "green" : "amber"}">${item.locationStatus === "within-project-area" ? "On site" : "Outside project area"}</span></td><td>${item.consentConfirmed ? "Confirmed" : "Not confirmed"}</td><td>${escapeHtml(item.trackingEnds)}</td></tr>`))}
      </article>
    </section>`;
}

function renderCompliance() {
  return `
    <section class="page">
      ${pageHead(routes.compliance, '<button class="button" data-create="compliance" type="button">Add requirement</button><button class="button secondary" data-create="consent" type="button">Add consent</button>')}
      <article class="panel">
        <div class="panel-head"><div><h2>Project compliance checklist</h2><p>Requirements remain connected to a project, jurisdiction, municipality, and official source.</p></div></div>
        ${demoTable(["Project", "Jurisdiction", "Municipality", "Category", "Requirement", "Official source", "Status"], state.compliance.map((item) => `<tr><td><a href="#projects">${escapeHtml(item.projectName)}</a></td><td>${escapeHtml(item.jurisdiction)}</td><td>${escapeHtml(item.municipality)}</td><td>${escapeHtml(item.category)}</td><td>${escapeHtml(item.requirement)}</td><td>${item.sourceUrl ? `<a href="${escapeHtml(item.sourceUrl)}" target="_blank" rel="noopener">${escapeHtml(item.sourceLabel || "Official source")}</a>` : escapeHtml(item.sourceLabel || "Not entered")}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
      </article>
      <article class="panel">
        <div class="panel-head"><div><h2>Contract consent library</h2><p>Draft legal language must be reviewed before activation. Signatures will preserve document version, signer, timestamp, and evidence.</p></div></div>
        ${demoTable(["Party type", "Consent or document", "Version", "Status"], state.consents.map((item) => `<tr><td><strong>${escapeHtml(item.partyType)}</strong></td><td>${escapeHtml(item.document)}</td><td>${escapeHtml(item.version)}</td><td><span class="status-pill ${statusClass(item.status)}">${escapeHtml(formatStatus(item.status))}</span></td></tr>`))}
        <p class="privacy-note">This module organizes compliance records; it does not replace advice from a licensed attorney, design professional, building official, or code professional.</p>
      </article>
    </section>`;
}

function renderReports() {
  const sections = ["Project summary", "Client and billing information", "Free On-Site Visit", "Estimates and contracts", "Invoices and payment stages", "Financial overview", "Team Members", "Vendors", "Subcontractors", "Materials in project", "Change orders / New Work", "Schedule and attendance", "Compliance and consent", "Progress photos and timeline", "Audit history"];
  return `
    <section class="page">
      ${pageHead(routes.reports)}
      <div class="report-layout">
        <form class="panel report-controls" id="reportBuilder">
          <div class="panel-head"><div><h2>Report settings</h2><p>Combine only the sections needed for this report.</p></div></div>
          <label>Report type<select name="type"><option>Project Report</option><option>Financial Report</option><option>Client History</option><option>Team & Payments</option><option>Materials & Vendors</option><option>Custom Report</option></select></label>
          <label>Project<select name="project"><option value="all">All projects</option>${projectOptions()}</select></label>
          <label>Client<select name="client"><option value="all">All clients</option>${clientOptions()}</select></label>
          <label>Team Member, Vendor, or Subcontractor<select name="person"><option value="all">All people and companies</option>${reportPeopleOptions()}</select></label>
          <label>Material<select name="material"><option value="all">All materials</option>${materialOptions()}</select></label>
          <label>Date range<select name="period"><option>All dates</option><option>This week</option><option>This month</option><option>This quarter</option><option>This year</option><option>Custom range</option></select></label>
          <fieldset>
            <legend>Include sections</legend>
            ${sections.map((section, index) => `<label class="check-row"><input type="checkbox" name="sections" value="${escapeHtml(section)}" ${index < 3 ? "checked" : ""} />${escapeHtml(section)}</label>`).join("")}
          </fieldset>
          <div class="button-row">
            <button class="button" type="submit">Update preview</button>
            <button class="button secondary" id="printReport" type="button">Print / Save PDF</button>
          </div>
        </form>
        <article class="report-preview" id="reportPreview" aria-live="polite"></article>
      </div>
    </section>`;
}

function renderSecurity() {
  const items = [
    ["Dedicated No Limit database", "Required before live operational data is connected.", "Planned"],
    ["Individual authentication", "Every person receives a separate account.", "Planned"],
    ["Role-based permissions", "Administrator, project manager, financial, and authorized team member.", "Planned"],
    ["Two-factor authentication", "Required for administrator accounts.", "Planned"],
    ["Audit history", "Important reads, writes, exports, and permission changes.", "Planned"],
    ["Work-hour location controls", "Explicit check-in consent, project-area verification, and automatic stop at 6:00 PM.", "Planned"],
    ["Backup and recovery", "Automatic backups plus a tested recovery procedure.", "Planned"],
  ];
  return `
    <section class="page">
      ${pageHead(routes.security)}
      <div class="content-grid">
        <article class="panel">
          <div class="panel-head"><div><h2>Isolation checklist</h2><p>No production connection will be made before these controls are approved.</p></div></div>
          <div class="security-list">
            ${items.map(([title, copy, status]) => `<div class="security-row"><div><strong>${title}</strong><p>${copy}</p></div><span class="status-pill amber">${status}</span></div>`).join("")}
          </div>
        </article>
        <aside class="panel">
          <div class="panel-head"><div><h2>Current preview state</h2><p>Verified separation</p></div></div>
          <div class="security-list">
            <div class="security-row"><div><strong>TAG data</strong><p>No connection in this preview.</p></div><span class="status-pill green">Separated</span></div>
            <div class="security-row"><div><strong>Public website</strong><p>No admin changes published.</p></div><span class="status-pill green">Preserved</span></div>
            <div class="security-row"><div><strong>Live writes</strong><p>Disabled until the dedicated database is ready.</p></div><span class="status-pill green">Protected</span></div>
          </div>
        </aside>
      </div>
    </section>`;
}

const renderers = {
  overview: renderOverview,
  requests: renderRequests,
  documents: renderDocuments,
  clients: renderClients,
  projects: renderProjects,
  financial: renderFinancial,
  materials: renderMaterials,
  team: renderTeam,
  schedule: renderSchedule,
  media: renderMedia,
  map: renderMap,
  compliance: renderCompliance,
  reports: renderReports,
  security: renderSecurity,
};

function convertEstimateToContract(estimateId) {
  const estimate = state.estimates.find((item) => item.id === estimateId);
  if (!estimate || estimate.contractId) return;
  const client = state.clients.find((item) => item.id === estimate.clientId);
  const primaryService = estimate.items[0]?.category || "Custom / New Work";
  const newProject = {
    id: nextId(state.projects, "PR"),
    name: `${client.name} — Contract Work`,
    clientId: client.id,
    clientName: client.name,
    status: "planned",
    service: primaryService,
    siteStreet: "",
    siteCity: "",
    siteState: "NJ",
    sitePostalCode: "",
    startDate: "To be scheduled",
    progress: 0,
    contractValue: estimate.total,
    cost: 0,
    outstanding: estimate.total,
    managerId: "",
  };
  const project = state.projects.find((item) => item.id === estimate.projectId) || newProject;
  const contractId = nextId(state.contracts, "CTR");
  const invoiceId = nextId(state.invoices, "INV");
  if (project === newProject) state.projects.unshift(project);
  else {
    project.contractValue = estimate.total;
    project.outstanding = estimate.total;
  }
  client.projectIds = [...new Set([...(client.projectIds || []), project.id])];
  state.contracts.unshift({ id: contractId, estimateId: estimate.id, clientId: client.id, clientName: client.name, projectId: project.id, status: "active", signedDate: readableDate(), value: estimate.total, invoiceId });
  state.invoices.unshift({
    id: invoiceId,
    contractId,
    projectId: project.id,
    clientId: client.id,
    clientName: client.name,
    issueDate: readableDate(),
    dueDate: readableDate(new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10)),
    terms: "Net 30",
    status: "open",
    items: JSON.parse(JSON.stringify(estimate.items)),
    schedule: [
      { label: "Initial deposit", percent: 30, amount: estimate.total * 0.3, status: "due" },
      { label: "Mid-project milestone", percent: 40, amount: estimate.total * 0.4, status: "upcoming" },
      { label: "Final completion", percent: 30, amount: estimate.total * 0.3, status: "upcoming" },
    ],
    total: estimate.total,
    paid: 0,
    balance: estimate.total,
    notes: "Thank you for choosing No Limit Carpentry.",
  });
  estimate.status = "approved";
  estimate.contractId = contractId;
  savePreviewState();
  renderRoute();
}

function updateReportPreview() {
  const form = document.getElementById("reportBuilder");
  const preview = document.getElementById("reportPreview");
  if (!form || !preview) return;
  const formData = new FormData(form);
  const type = formData.get("type") || "Project Report";
  const projectId = formData.get("project") || "all";
  const clientId = formData.get("client") || "all";
  const personId = formData.get("person") || "all";
  const materialId = formData.get("material") || "all";
  const period = formData.get("period") || "All dates";
  const sections = formData.getAll("sections");
  let selectedProjects = projectId === "all" ? state.projects : state.projects.filter((project) => project.id === projectId);
  if (clientId !== "all") selectedProjects = selectedProjects.filter((project) => project.clientId === clientId);
  const selectedProjectIds = selectedProjects.map((project) => project.id);
  const selectedClientIds = [...new Set(selectedProjects.map((project) => project.clientId))];
  const selectedClients = state.clients.filter((client) => selectedClientIds.includes(client.id));
  const selectedTransactions = state.transactions.filter((item) => selectedProjectIds.includes(item.projectId));
  let selectedPeople = state.people.filter((person) => person.projectIds.some((id) => selectedProjectIds.includes(id)));
  if (personId !== "all") selectedPeople = selectedPeople.filter((person) => person.id === personId);
  let selectedMaterials = state.materials.filter((item) => selectedProjectIds.includes(item.projectId));
  if (materialId !== "all") selectedMaterials = selectedMaterials.filter((item) => item.id === materialId);
  const selectedMedia = state.media.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedContracts = state.contracts.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedEstimates = state.estimates.filter((item) => selectedClientIds.includes(item.clientId));
  const selectedInvoices = state.invoices.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedInvoiceIds = selectedInvoices.map((item) => item.id);
  const selectedPayments = state.payments.filter((item) => selectedInvoiceIds.includes(item.invoiceId));
  const selectedChanges = state.changeOrders.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedSchedule = state.schedule.filter((item) => selectedProjectIds.includes(item.projectId) && (personId === "all" || item.personId === personId));
  const selectedAttendance = state.attendance.filter((item) => selectedProjectIds.includes(item.projectId) && (personId === "all" || item.personId === personId));
  const selectedCompliance = state.compliance.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedReceipts = state.expenseReceipts.filter((item) => selectedProjectIds.includes(item.projectId));
  const selectedVisits = state.siteVisits.filter((item) => selectedClientIds.includes(item.clientId) && (!item.projectId || selectedProjectIds.includes(item.projectId)));
  const projectLabel = projectId === "all" ? "All demo projects" : selectedProjects[0]?.name || "No project selected";
  const issueDate = new Intl.DateTimeFormat("en-US", { dateStyle: "long" }).format(new Date());
  const reportSections = {
    "Project summary": demoTable(["Project", "Client", "Project address", "Status", "Progress", "Contract value"], selectedProjects.map((project) => `<tr><td>${escapeHtml(project.name)}</td><td>${escapeHtml(project.clientName)}</td><td>${escapeHtml(projectLocation(project))}</td><td>${escapeHtml(formatStatus(project.status))}</td><td>${project.progress}%</td><td>${formatCurrency(project.contractValue)}</td></tr>`)),
    "Client and billing information": demoTable(["Client", "Primary contact", "Client address", "Billing address", "Website"], selectedClients.map((client) => `<tr><td>${escapeHtml(client.name)}</td><td>${escapeHtml(clientContact(client))}</td><td>${escapeHtml(clientLocation(client))}</td><td>${escapeHtml([client.billingStreet || client.street, client.billingCity || client.city, client.billingState || client.state, client.billingPostalCode || client.postalCode].filter(Boolean).join(", "))}</td><td>${escapeHtml(client.website || "—")}</td></tr>`)),
    "Free On-Site Visit": demoTable(["Date", "Client", "Request", "Assigned to", "Status", "Notes"], selectedVisits.map((visit) => `<tr><td>${escapeHtml(visit.visitDate)}</td><td>${escapeHtml(visit.clientName)}</td><td>${escapeHtml(visit.requestId)}</td><td>${escapeHtml(visit.assignedTo)}</td><td>${escapeHtml(formatStatus(visit.status))}</td><td>${escapeHtml(visit.notes)}</td></tr>`)),
    "Estimates and contracts": `<p>${selectedEstimates.length} estimate${selectedEstimates.length === 1 ? "" : "s"} · ${selectedContracts.length} contract${selectedContracts.length === 1 ? "" : "s"} · ${formatCurrency(selectedContracts.reduce((sum, item) => sum + item.value, 0))} contracted.</p>`,
    "Invoices and payment stages": demoTable(["Invoice", "Total", "Paid", "Balance", "Status"], selectedInvoices.map((invoice) => `<tr><td>${escapeHtml(invoice.id)}</td><td>${formatCurrency(invoice.total)}</td><td>${formatCurrency(invoice.paid)}</td><td>${formatCurrency(invoice.balance)}</td><td>${escapeHtml(formatStatus(invoice.status))}</td></tr>`)),
    "Financial overview": `<p>${formatCurrency(selectedTransactions.filter((item) => item.type === "receivable" && item.status === "paid").reduce((sum, item) => sum + item.amount, 0))} received · ${formatCurrency(selectedTransactions.filter((item) => item.type === "receivable" && item.status !== "paid").reduce((sum, item) => sum + item.amount, 0))} outstanding · ${formatCurrency(selectedProjects.reduce((sum, item) => sum + projectActualCost(item), 0))} actual project cost, including ${selectedReceipts.length} additional receipt${selectedReceipts.length === 1 ? "" : "s"}.</p>`,
    "Team Members": `<p>${selectedPeople.filter((person) => person.type === "Team Member").map((person) => `${escapeHtml(person.name)} — ${escapeHtml(person.role)}`).join(" · ") || "No linked Team Members"}</p>`,
    "Vendors": `<p>${selectedPeople.filter((person) => person.type === "Vendor").map((person) => `${escapeHtml(person.name)} — ${escapeHtml(person.role)}`).join(" · ") || "No linked Vendors"}</p>`,
    "Subcontractors": `<p>${selectedPeople.filter((person) => person.type === "Subcontractor").map((person) => `${escapeHtml(person.name)} — ${escapeHtml(person.role)}`).join(" · ") || "No linked Subcontractors"}</p>`,
    "Materials in project": demoTable(["Project", "Material", "Vendor", "Amount", "Status"], selectedMaterials.map((item) => `<tr><td>${escapeHtml(item.projectName)}</td><td>${escapeHtml(item.description)}</td><td>${escapeHtml(item.vendor || "Not entered")}</td><td>${formatCurrency(item.amount)}</td><td>${escapeHtml(formatStatus(item.status))}</td></tr>`)),
    "Change orders / New Work": demoTable(["Project", "Category", "Description", "Amount", "Status"], selectedChanges.map((item) => `<tr><td>${escapeHtml(item.projectName)}</td><td>${escapeHtml(item.category || "Custom / New Work")}</td><td>${escapeHtml(item.description)}</td><td>${formatCurrency(item.amount)}</td><td>${escapeHtml(formatStatus(item.status))}</td></tr>`)),
    "Schedule and attendance": `<p>${selectedSchedule.length} work assignment${selectedSchedule.length === 1 ? "" : "s"} · ${selectedAttendance.length} authorized check-in${selectedAttendance.length === 1 ? "" : "s"}. Schedule access window: ${scheduleAccessWindow.starts} to ${scheduleAccessWindow.ends}.</p>`,
    "Compliance and consent": `<p>${selectedCompliance.length} project compliance requirement${selectedCompliance.length === 1 ? "" : "s"} · ${state.consents.length} consent template${state.consents.length === 1 ? "" : "s"} in the library.</p>`,
    "Progress photos and timeline": `<p>${selectedMedia.length} media record${selectedMedia.length === 1 ? "" : "s"} connected to the selected project scope.</p>`,
    "Audit history": `<p>${selectedPayments.length} payment record${selectedPayments.length === 1 ? "" : "s"} included. Preview activity is stored only in this browser; the production audit log is not connected.</p>`,
  };
  preview.innerHTML = `
    <div class="report-brand"><img src="logo.png" alt="No Limit Carpentry" /><span>Private administrative report</span></div>
    <h2>${escapeHtml(type)}</h2>
    <p class="report-meta">${escapeHtml(projectLabel)} · ${escapeHtml(period)} · Issued ${escapeHtml(issueDate)} · Demo preview only</p>
    ${sections.length ? sections.map((section) => `<section class="report-section"><h3>${escapeHtml(section)}</h3>${reportSections[section] || ""}</section>`).join("") : '<section class="report-section"><h3>No sections selected</h3><p>Select at least one report section and update the preview.</p></section>'}`;
}

function bindPageEvents(routeName) {
  content.querySelectorAll("[data-create]").forEach((button) => button.addEventListener("click", () => openDataEntry(button.dataset.create)));
  content.querySelectorAll("[data-edit-client]").forEach((button) => button.addEventListener("click", () => openDataEntry("client", button.dataset.editClient)));
  content.querySelectorAll("[data-edit-project]").forEach((button) => button.addEventListener("click", () => openDataEntry("project", button.dataset.editProject)));
  content.querySelectorAll("[data-edit-person]").forEach((button) => button.addEventListener("click", () => openDataEntry("person", button.dataset.editPerson)));
  content.querySelectorAll("[data-edit-estimate]").forEach((button) => button.addEventListener("click", () => openDataEntry("estimate", button.dataset.editEstimate)));
  content.querySelectorAll("[data-view-document]").forEach((button) => button.addEventListener("click", () => openBusinessDocument(button.dataset.viewDocument, button.dataset.documentId)));
  content.querySelectorAll("[data-convert-estimate]").forEach((button) => button.addEventListener("click", () => convertEstimateToContract(button.dataset.convertEstimate)));
  content.querySelectorAll("[data-renewal-person]").forEach((button) => button.addEventListener("click", () => prepareInsuranceRenewal(button.dataset.renewalPerson)));
  if (routeName === "reports") {
    const form = document.getElementById("reportBuilder");
    form?.addEventListener("submit", (event) => {
      event.preventDefault();
      updateReportPreview();
    });
    document.getElementById("printReport")?.addEventListener("click", () => window.print());
    updateReportPreview();
  }
}

function closeNavigation() {
  document.body.classList.remove("nav-open");
  menuButton.setAttribute("aria-expanded", "false");
}

function renderRoute() {
  const requested = location.hash.replace(/^#/, "") || "overview";
  const routeName = routes[requested] ? requested : "overview";
  const route = routes[routeName];
  topbarTitle.textContent = route.title;
  document.title = `No Limit | ${route.title}`;
  nav.querySelectorAll("a[data-route]").forEach((link) => link.classList.toggle("active", link.dataset.route === routeName));
  content.innerHTML = renderers[routeName]();
  content.focus({ preventScroll: true });
  window.scrollTo({ top: 0, behavior: "instant" });
  bindPageEvents(routeName);
  applyBetaAccess();
  closeNavigation();
}

menuButton.addEventListener("click", () => {
  const isOpen = document.body.classList.toggle("nav-open");
  menuButton.setAttribute("aria-expanded", String(isOpen));
});
mobileOverlay.addEventListener("click", closeNavigation);
dataForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (event.submitter?.value === "cancel") {
    dataDialog.close();
    return;
  }
  if (!dataForm.reportValidity()) return;
  saveDataEntry(new FormData(dataForm));
  dataDialog.close();
  renderRoute();
});
dataDialog.addEventListener("click", (event) => {
  if (event.target === dataDialog) dataDialog.close();
});
closeDocumentButton.addEventListener("click", () => documentDialog.close());
documentDialog.addEventListener("click", (event) => {
  if (event.target === documentDialog) documentDialog.close();
});
editDocumentButton.addEventListener("click", () => setDocumentEditing(true));
addDocumentItemButton.addEventListener("click", addDocumentLineItem);
saveDocumentButton.addEventListener("click", saveDocumentEdits);
printDocumentButton.addEventListener("click", () => {
  document.body.classList.add("printing-document");
  window.print();
  setTimeout(() => document.body.classList.remove("printing-document"), 250);
});
window.addEventListener("afterprint", () => document.body.classList.remove("printing-document"));
emailDocumentButton.addEventListener("click", () => prepareDocumentDelivery("email"));
messageDocumentButton.addEventListener("click", () => prepareDocumentDelivery("message"));
closeCustomServiceButton.addEventListener("click", closeCustomServiceDialog);
cancelCustomServiceButton.addEventListener("click", closeCustomServiceDialog);
customServiceForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!customServiceForm.reportValidity()) return;
  saveCustomService(new FormData(customServiceForm));
});
window.addEventListener("hashchange", renderRoute);
window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeNavigation();
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(loginForm);
  const email = String(formData.get("email") || "").trim().toLowerCase();
  const password = String(formData.get("password") || "");
  if (!email || !password || !window.noLimitSupabaseClient) return;
  authStatus.textContent = "Signing in securely…";
  const { data, error } = await window.noLimitSupabaseClient.auth.signInWithPassword({ email, password });
  if (error || !data.session) {
    authStatus.textContent = error?.message || "We could not sign you in.";
    return;
  }
  const user = await userFromSession(data.session);
  if (!user) {
    await window.noLimitSupabaseClient.auth.signOut();
    authStatus.textContent = "This account is not authorized for No Limit Administration.";
    return;
  }
  await activateBetaUser(user, data.session);
});

setPasswordForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(setPasswordForm);
  const password = String(formData.get("newPassword") || "");
  const confirmation = String(formData.get("confirmPassword") || "");
  if (password !== confirmation) {
    authStatus.textContent = "The passwords do not match.";
    return;
  }
  authStatus.textContent = "Saving your password…";
  const { error } = await window.noLimitSupabaseClient.auth.updateUser({ password });
  if (error) {
    authStatus.textContent = error.message;
    return;
  }
  history.replaceState(null, "", `${location.pathname}${location.search}#overview`);
  const { data: { session } } = await window.noLimitSupabaseClient.auth.getSession();
  const user = await userFromSession(session);
  if (!user) {
    authStatus.textContent = "Password saved, but this account is not authorized yet.";
    return;
  }
  await activateBetaUser(user, session);
});

resetPasswordButton.addEventListener("click", async () => {
  const emailInput = loginForm.elements.email;
  const email = String(emailInput.value || "").trim().toLowerCase();
  if (!email) {
    authStatus.textContent = "Enter your email first, then request the password link.";
    emailInput.focus();
    return;
  }
  const redirectTo = `${location.origin}${location.pathname}`;
  authStatus.textContent = "Sending a secure password link…";
  const { error } = await window.noLimitSupabaseClient.auth.resetPasswordForEmail(email, { redirectTo });
  authStatus.textContent = error ? error.message : "Check your email for the secure password link.";
});

profileButton.addEventListener("click", async () => {
  if (window.noLimitSupabaseClient) await window.noLimitSupabaseClient.auth.signOut();
  currentBetaUser = null;
  currentAuthSession = null;
  document.body.classList.add("auth-required");
  authStatus.textContent = "You have signed out.";
});

renderRoute();
void startBeta();
