/* Independent service catalog for the existing Administration workspace.
 * Loaded after app.js. Reuses its state, authorization, sync and document UI.
 * Catalog records are templates; document items always remain snapshots.
 */
"use strict";

const NoLimitServiceCatalog = (() => {
  const fallbackCategory = "Custom / New Work";
  let editingId = "";
  let busy = false;
  let query = "";
  let categoryFilter = "";
  let showRemoved = false;
  let notice = "";
  const originalSaveDataEntry = saveDataEntry;
  const originalNormalize = normalizedSharedState;
  const originalRefresh = refreshWorkspaceIfChanged;

  function authorize() {
    if (!betaCanCreate("catalogService")) throw new Error("Only authorized office users can change the service catalog.");
  }

  function getService(id) {
    return state.customServices.find((service) => service.id === id) || null;
  }

  function validate(values, id = "") {
    const title = String(values.title || "").trim();
    const description = String(values.description || "").trim();
    const unit = String(values.unit || "").trim();
    const category = String(values.category || fallbackCategory).trim();
    const priceText = String(values.unitPrice ?? "").trim();
    const unitPrice = Number(priceText);
    if (!title || !description || !unit) throw new Error("Enter a service title, description and unit.");
    if (!serviceCategories.includes(category)) throw new Error("Choose a valid service category.");
    if (!priceText || !Number.isFinite(unitPrice) || unitPrice < 0) throw new Error("Enter a valid price of zero or more.");
    const duplicate = state.customServices.find((service) => service.id !== id && serviceCatalogKey(service.title) === serviceCatalogKey(title));
    if (duplicate) throw new Error(duplicate.archived ? "A removed service already has this title. Restore it from Removed items, or use a different title." : "A service already has this title. Edit that service or use a different title.");
    return { title, description, unit, category, unitPrice };
  }

  // One confirmed write, with rollback on failure. The existing RPC also checks
  // the workspace version, so stale devices cannot silently overwrite changes.
  async function commit(change) {
    authorize();
    if (busy) throw new Error("A catalog change is still saving. Please wait.");
    busy = true;
    const before = structuredClone(state);
    try {
      const result = change();
      if (!await savePreviewState()) throw new Error("The change was not saved. Refresh the workspace and try again.");
      return result;
    } catch (error) {
      state = before;
      throw error;
    } finally {
      busy = false;
    }
  }

  async function save(values, id = "") {
    authorize();
    const fields = validate(values, id);
    return commit(() => {
      const previous = id ? getService(id) : null;
      if (id && !previous) throw new Error("This service no longer exists. Refresh the catalog.");
      if (previous?.archived) throw new Error("Restore this service before editing it.");
      const now = new Date().toISOString();
      const service = { ...previous, ...fields, id: previous?.id || nextId(state.customServices, "SVC"), archived: false, createdAt: previous?.createdAt || now, updatedAt: now };
      if (previous) state.customServices = state.customServices.map((item) => item.id === id ? service : item);
      else state.customServices.unshift(service);
      return service;
    });
  }

  async function setRemoved(id, removed) {
    return commit(() => {
      const service = getService(id);
      if (!service) throw new Error("This service no longer exists. Refresh the catalog.");
      service.archived = removed;
      service.updatedAt = new Date().toISOString();
      // Keep the record and its stable ID. No estimate/invoice/order is touched.
      return service;
    });
  }

  function snapshot(service, values = {}) {
    const price = values.unitPrice === undefined || values.unitPrice === null || values.unitPrice === "" ? service.unitPrice : Number(values.unitPrice);
    const quantity = values.quantity === undefined ? 1 : Number(values.quantity);
    if (!Number.isFinite(price) || price < 0 || !Number.isFinite(quantity) || quantity < 0) throw new Error("Quantity and price must be valid numbers of zero or more.");
    return { category: service.category || fallbackCategory, customServiceId: service.id, title: values.title ?? service.title, description: values.description ?? service.description, unit: service.unit || "each", quantity, unitPrice: price };
  }

  function options(selected = "", selectedId = "") {
    const standard = serviceCategories.map((category) => `<option value="${escapeHtml(category)}"${category === selected && !selectedId ? " selected" : ""}>${escapeHtml(category)}</option>`).join("");
    const active = state.customServices.filter((service) => !service.archived);
    const saved = active.length ? `<optgroup label="Service catalog">${active.map((service) => `<option value="custom-service:${escapeHtml(service.id)}"${service.id === selectedId ? " selected" : ""}>${escapeHtml(service.title)}</option>`).join("")}</optgroup>` : "";
    const previous = selectedId ? getService(selectedId) : null;
    // An old document may retain an archived (or legacy missing) selection.
    // It stays selected but cannot be newly selected for another line item.
    const historical = selectedId && (!previous || previous.archived) ? `<optgroup label="Existing document item"><option value="custom-service:${escapeHtml(selectedId)}" selected disabled>${escapeHtml(previous?.title || "Previously saved service")} (removed)</option></optgroup>` : "";
    return standard + saved + historical;
  }
  serviceOptions = options;

  // The legacy change-order migration matched only by title. Follow stable IDs
  // first so renaming/removing a catalog service cannot recreate its old title.
  normalizedSharedState = function normalizeCatalog(candidate = {}) {
    const source = candidate && typeof candidate === "object" ? candidate : {};
    const base = isLocalPreview ? cloneDefaultState() : emptyOperationalState();
    const orders = structuredClone(Array.isArray(source.changeOrders) ? source.changeOrders : base.changeOrders);
    const services = structuredClone(Array.isArray(source.customServices) ? source.customServices : base.customServices);
    const result = originalNormalize({ ...source, customServices: services, changeOrders: [] });
    result.changeOrders = orders;
    orders.forEach((order) => {
      let service = result.customServices.find((item) => item.id === order.customServiceId);
      if (service) return;
      const name = String(order.title || order.description || "").trim();
      if (!name) return;
      service = result.customServices.find((item) => serviceCatalogKey(item.title) === serviceCatalogKey(name));
      if (!service) {
        const id = order.customServiceId || `SVC-CHANGE-${String(order.id || serviceCatalogKey(name)).replace(/[^A-Za-z0-9_-]+/g, "-")}`;
        service = { id, title: name, description: String(order.description || "").trim(), category: order.category || fallbackCategory, unit: order.unit || "each", unitPrice: Number(order.unitPrice || 0) };
        result.customServices.push(service);
      }
      if (!order.customServiceId) order.customServiceId = service.id;
    });
    return result;
  };

  refreshWorkspaceIfChanged = function refreshOutsideCatalogEdit(...args) {
    if (busy || customServiceDialog.open) return Promise.resolve(false);
    return originalRefresh(...args);
  };

  function rows() {
    const key = serviceCatalogKey(query);
    const items = state.customServices.filter((service) => Boolean(service.archived) === showRemoved && (!categoryFilter || (service.category || fallbackCategory) === categoryFilter) && (!key || serviceCatalogKey(`${service.title} ${service.description} ${service.category || fallbackCategory} ${service.unit || "each"}`).includes(key)));
    if (!items.length) return emptyState(showRemoved ? "No removed items in this view" : "No services in this view", "Add a new item or change the filters. No invoice or estimate is required to maintain this catalog.");
    return demoTable(["Service", "Category", "Description", "Unit", "Default price", "Actions"], items.map((service) => `<tr><td><strong>${escapeHtml(service.title)}</strong><small class="record-id">${escapeHtml(service.id)}</small></td><td>${escapeHtml(service.category || fallbackCategory)}</td><td>${escapeHtml(service.description || "")}</td><td>${escapeHtml(service.unit || "each")}</td><td>${formatCurrency(service.unitPrice)}</td><td><div class="button-row">${service.archived ? `<button class="text-button" type="button" data-catalog-restore="${escapeHtml(service.id)}">Restore</button>` : `<button class="text-button" type="button" data-catalog-edit="${escapeHtml(service.id)}">Edit</button><button class="text-button danger-button" type="button" data-catalog-remove="${escapeHtml(service.id)}">Remove</button>`}</div></td></tr>`));
  }

  function render() {
    if (!betaCanCreate("catalogService")) return emptyState("Office access required", "Contact the office to request service catalog changes.");
    return `<section class="page">${pageHead(routes["line-items"], '<button class="button" type="button" data-catalog-add>Add new item</button>')}<article class="panel"><div class="panel-head"><div><h2>Service catalog</h2><p>Create services here first, then select them when preparing estimates and invoices. Editing a default never changes an existing document.</p></div></div><div class="filter-bar"><label>Search services<input type="search" data-catalog-search value="${escapeHtml(query)}" placeholder="Service, description or unit" /></label><label>Category<select data-catalog-category><option value="">All categories</option>${serviceCategories.map((category) => `<option value="${escapeHtml(category)}"${categoryFilter === category ? " selected" : ""}>${escapeHtml(category)}</option>`).join("")}</select></label><label>Show<select data-catalog-status><option value="active"${!showRemoved ? " selected" : ""}>Active items</option><option value="removed"${showRemoved ? " selected" : ""}>Removed items</option></select></label></div><p data-catalog-notice role="status">${escapeHtml(notice)}</p><div id="serviceCatalogRows">${rows()}</div></article><div class="button-row"><a class="button secondary" href="#documents">Open estimates</a><a class="button secondary" href="#financial">Open invoices</a></div></section>`;
  }

  function renderRows() {
    const target = document.getElementById("serviceCatalogRows");
    if (target) target.innerHTML = rows();
  }

  function openEditor(id = "") {
    authorize();
    if (busy) return;
    const service = id ? getService(id) : null;
    if (id && (!service || service.archived)) throw new Error("This item is not active. Refresh the catalog or restore it first.");
    editingId = id;
    pendingCustomServiceTarget = { mode: "catalog" };
    customServiceForm.reset();
    document.getElementById("customServiceTitle").textContent = id ? "Edit catalog item" : "Add new item";
    for (const name of ["title", "description", "unit", "unitPrice", "category"]) {
      const control = customServiceForm.elements.namedItem(name);
      if (control) control.value = service?.[name] ?? (name === "category" ? fallbackCategory : name === "unit" ? "each" : "");
    }
    customServiceDialog.showModal();
  }

  saveCustomService = async function saveCatalogForm(formData) {
    if (busy) return;
    const buttons = [...customServiceForm.querySelectorAll("button")];
    buttons.forEach((button) => { button.disabled = true; });
    const target = pendingCustomServiceTarget;
    try {
      const service = await save(Object.fromEntries(formData), editingId);
      if (target?.mode === "document" && activeDocument?.draftItems?.[target.index]) {
        const previous = activeDocument.draftItems[target.index];
        activeDocument.draftItems[target.index] = { ...previous, ...snapshot(service, { quantity: previous.quantity }) };
        setDocumentEditing(true);
      } else if (target?.mode === "data-form") {
        const category = dataForm.elements.namedItem("category");
        if (category) { category.innerHTML = options(service.category, service.id); category.value = `custom-service:${service.id}`; }
        for (const name of ["title", "description", "unitPrice"]) {
          const control = dataForm.elements.namedItem(name);
          if (control) control.value = service[name];
        }
      }
      notice = editingId ? "Service updated. Existing estimates and invoices were not changed." : "Service added to the catalog and available in estimates and invoices.";
      customServiceDialog.close();
      pendingCustomServiceTarget = null;
      if (target?.mode === "catalog") renderRoute();
      // Audit telemetry must never turn a confirmed save into a failure.
      void Promise.resolve(recordActivity("record_saved", "catalogService", service.id, { section: "line-items" })).catch(() => {});
    } catch (error) {
      window.alert(error?.message || "The service could not be saved.");
    } finally {
      buttons.forEach((button) => { button.disabled = false; });
    }
  };

  // Store a selected service as a fresh line-item snapshot, including category
  // and unit. Other entry types keep their existing implementation.
  saveDataEntry = async function saveCatalogDocumentEntry(formData) {
    if (busy) throw new Error("A catalog change is still saving. Please wait.");
    const type = pendingRecordType;
    const service = customServiceFromSelection(formData.get("category"));
    if (!["estimate", "invoiceItem", "changeOrder"].includes(type) || !service) return originalSaveDataEntry(formData);
    const existing = type === "estimate" ? state.estimates.find((item) => item.id === pendingRecordId) : null;
    const previousItem = existing?.items?.[0];
    if (service.archived && previousItem?.customServiceId !== service.id) throw new Error("This service was removed. Restore it in Line Items before using it again.");
    const item = snapshot(service, { title: String(formData.get("title") ?? service.title), description: String(formData.get("description") ?? service.description), quantity: Number(formData.get("quantity")), unitPrice: formData.get("unitPrice") });
    if (previousItem?.customServiceId === service.id) {
      item.category = previousItem.category || fallbackCategory;
      item.unit = previousItem.unit || "each";
    }
    return commit(() => {
      if (type === "estimate") {
        const client = state.clients.find((entry) => entry.id === formData.get("clientId"));
        if (!client) throw new Error("Choose an existing client for this estimate.");
        const discount = Number(formData.get("discount") || 0);
        const taxPercent = Number(formData.get("taxPercent") || 0);
        if (!Number.isFinite(discount) || discount < 0 || !Number.isFinite(taxPercent) || taxPercent < 0) throw new Error("Enter valid discount and tax values.");
        const items = [item, ...structuredClone(existing?.items?.slice(1) || [])];
        const subtotal = items.reduce((total, line) => total + Number(line.quantity) * Number(line.unitPrice), 0);
        const taxable = Math.max(0, subtotal - discount);
        const record = { ...existing, id: existing?.id || nextId(state.estimates, "EST"), requestId: formData.get("requestId") || "", clientId: client.id, clientName: client.name, projectId: formData.get("projectId") || "", issueDate: readableDate(formData.get("issueDate")), status: formData.get("status"), revision: existing ? Number(existing.revision || 1) + 1 : 1, validUntil: readableDate(formData.get("validUntil")), discount, tax: taxable * taxPercent / 100, items, total: taxable * (1 + taxPercent / 100), contractId: existing?.contractId || "", notes: formData.get("notes") || "" };
        if (existing) state.estimates = state.estimates.map((entry) => entry.id === existing.id ? record : entry);
        else state.estimates.unshift(record);
      } else if (type === "invoiceItem") {
        const invoice = state.invoices.find((entry) => entry.id === formData.get("invoiceId"));
        if (!invoice) throw new Error("Choose an existing invoice.");
        invoice.items.push(item);
        invoice.total = invoice.items.reduce((total, line) => total + Number(line.quantity) * Number(line.unitPrice), 0);
        invoice.balance = Math.max(0, invoice.total - Number(invoice.paid || 0));
      } else {
        const project = state.projects.find((entry) => entry.id === formData.get("projectId"));
        if (!project) throw new Error("Choose an existing project.");
        state.changeOrders.unshift({ ...item, id: nextId(state.changeOrders, "CO"), projectId: project.id, projectName: project.name, status: formData.get("status"), amount: item.quantity * item.unitPrice });
      }
    });
  };

  harvestDocumentDraft = function harvestCatalogDocumentDraft() {
    if (!activeDocument?.draftItems) return;
    activeDocument.draftItems.forEach((item, index) => {
      const control = (name) => documentCanvas.querySelector(`[name="${name}"][data-document-item="${index}"]`);
      const selected = control("category")?.value || item.category || fallbackCategory;
      const id = selected.startsWith("custom-service:") ? selected.slice("custom-service:".length) : "";
      const service = id ? getService(id) : null;
      if (id && id !== item.customServiceId && service && !service.archived) {
        item.category = service.category || fallbackCategory;
        item.unit = service.unit || "each";
      } else if (!id) {
        item.category = selected;
      }
      item.customServiceId = id;
      item.title = control("title")?.value ?? item.title ?? item.category;
      item.description = control("description")?.value ?? item.description ?? "";
      item.quantity = Number(control("quantity")?.value ?? item.quantity ?? 0);
      item.unitPrice = Number(control("unitPrice")?.value ?? item.unitPrice ?? 0);
    });
    activeDocument.draftNotes = documentCanvas.querySelector('[name="documentNotes"]')?.value || "";
  };

  bindDocumentEditor = function bindCatalogDocumentEditor() {
    documentCanvas.querySelectorAll('select[name="category"]').forEach((select) => select.addEventListener("change", () => {
      harvestDocumentDraft();
      const index = Number(select.dataset.documentItem);
      const service = customServiceFromSelection(select.value);
      if (service && !service.archived) {
        const previous = activeDocument.draftItems[index];
        activeDocument.draftItems[index] = { ...previous, ...snapshot(service, { quantity: previous.quantity }) };
        setDocumentEditing(true);
      } else if (select.value === fallbackCategory) {
        pendingCustomServiceTarget = { mode: "document", index };
        customServiceForm.reset();
        customServiceDialog.showModal();
      }
    }));
    documentCanvas.querySelectorAll("[data-remove-document-item]").forEach((button) => button.addEventListener("click", () => {
      harvestDocumentDraft();
      if (activeDocument.draftItems.length <= 1) { documentNotice.textContent = "A Proposal or Invoice must keep at least one line item."; return; }
      activeDocument.draftItems.splice(Number(button.dataset.removeDocumentItem), 1);
      setDocumentEditing(true);
    }));
  };

  routes["line-items"] = { title: "Line Items", kicker: "Reusable services", heading: "Manage services before creating a document.", description: "Add, edit and remove catalog items independently of invoices, estimates, clients or projects." };
  renderers["line-items"] = render;
  const originalFinancial = renderers.financial;
  renderers.financial = () => originalFinancial().replace('<h3 class="subsection-title">Invoice line items</h3>', '<div class="panel-head"><div><h3>Line Items - service catalog</h3><p>Maintain reusable services separately from document history.</p></div><div class="button-row"><button class="button" type="button" data-catalog-add>Add new item</button><a class="button secondary" href="#line-items">Manage catalog</a></div></div><h3 class="subsection-title">Invoice line items - document history</h3>');
  const originalDocuments = renderers.documents;
  renderers.documents = () => originalDocuments().replace(/<\/section>\s*$/, '<article class="panel"><div class="panel-head"><div><h2>Service catalog</h2><p>Prepare reusable line items without creating an estimate or invoice. Select a saved service in the Service category field when building a document.</p></div><a class="button secondary" href="#line-items">Manage Line Items</a></div></article></section>');

  const categoryLabel = document.createElement("label");
  categoryLabel.innerHTML = `Service category<select name="category" required>${serviceCategories.map((category) => `<option value="${escapeHtml(category)}"${category === fallbackCategory ? " selected" : ""}>${escapeHtml(category)}</option>`).join("")}</select>`;
  customServiceForm.querySelector(".form-grid").prepend(categoryLabel);
  customServiceDialog.addEventListener("close", () => {
    editingId = "";
    pendingCustomServiceTarget = null;
    document.getElementById("customServiceTitle").textContent = "Add new catalog item";
  });
  customServiceDialog.addEventListener("cancel", (event) => { if (busy) event.preventDefault(); });
  document.addEventListener("input", (event) => {
    if (event.target.matches("[data-catalog-search]")) { query = event.target.value; renderRows(); }
  });
  document.addEventListener("change", (event) => {
    if (event.target.matches("[data-catalog-category]")) { categoryFilter = event.target.value; renderRows(); }
    if (event.target.matches("[data-catalog-status]")) { showRemoved = event.target.value === "removed"; renderRows(); }
  });
  document.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-catalog-add], [data-catalog-edit], [data-catalog-remove], [data-catalog-restore]");
    if (!button || busy) return;
    try {
      authorize();
      if (button.hasAttribute("data-catalog-add")) { openEditor(); return; }
      if (button.hasAttribute("data-catalog-edit")) { openEditor(button.dataset.catalogEdit); return; }
      const removed = button.hasAttribute("data-catalog-remove");
      const id = removed ? button.dataset.catalogRemove : button.dataset.catalogRestore;
      const service = getService(id);
      if (!service) throw new Error("This service no longer exists. Refresh the catalog.");
      if (removed && !window.confirm(`Remove "${service.title}" from the active catalog? Existing estimates and invoices will not change. You can restore it from Removed items.`)) return;
      button.disabled = true;
      await setRemoved(id, removed);
      notice = removed ? "Service removed from the active catalog. Existing documents were not changed." : "Service restored and available for new documents.";
      renderRoute();
    } catch (error) {
      window.alert(error?.message || "The catalog could not be updated.");
    } finally {
      button.disabled = false;
    }
  });

  // Local preview starts synchronously in app.js, before this extension loads.
  if (isLocalPreview && new URLSearchParams(location.search).get("local-demo") === "1") renderRoute();
  return Object.freeze({ save, setRemoved, snapshot, options, validate, render, openEditor });
})();
