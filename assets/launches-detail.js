import {
  createEditableDraft,
  getLaunchById,
  getLaunchListRoute,
  getLaunchRoute,
  getLaunchSequence,
  saveLaunch
} from "./launches-store.js";
import { currency, lineTotal, mountGlobalSearch, setBackLinks, showMessage } from "./launches-ui.js";

let persistedEntry = null;
let draftEntry = null;

function getRouteId() {
  const parts = window.location.pathname.split("/").filter(Boolean);
  const last = decodeURIComponent(parts[parts.length - 1] || "");
  if (last && last !== "lancamento" && last !== "index.html") return last;
  const queryId = new URLSearchParams(window.location.search).get("id");
  return queryId ? decodeURIComponent(queryId) : "";
}

function updateTitle(entry) {
  document.title = `${entry.title} | Launch entry`;
  const title = document.getElementById("launchDetailTitle");
  const subtitle = document.getElementById("launchDetailSubtitle");
  const status = document.getElementById("launchDetailStatus");
  const amount = document.getElementById("launchDetailAmount");

  if (title) title.textContent = entry.title;
  if (subtitle) subtitle.textContent = `${entry.category} · ${entry.entryDate} · ${entry.client}`;
  if (status) status.textContent = entry.status;
  if (amount) amount.textContent = currency(entry.amount);
}

function fillHeaderFields(entry) {
  const mapping = {
    launchTitleInput: entry.title,
    launchDateInput: entry.entryDate,
    launchCategoryInput: entry.category,
    launchClientInput: entry.client,
    launchLocationInput: entry.location,
    launchStatusInput: entry.status,
    launchAmountInput: entry.amount,
    launchSummaryInput: entry.summary,
    launchNotesInput: entry.notes
  };

  Object.entries(mapping).forEach(([id, value]) => {
    const element = document.getElementById(id);
    if (element) element.value = value ?? "";
  });
}

function readHeaderFields() {
  return {
    ...draftEntry,
    title: document.getElementById("launchTitleInput")?.value.trim() || "",
    entryDate: document.getElementById("launchDateInput")?.value || "",
    category: document.getElementById("launchCategoryInput")?.value.trim() || "",
    client: document.getElementById("launchClientInput")?.value.trim() || "",
    location: document.getElementById("launchLocationInput")?.value.trim() || "",
    status: document.getElementById("launchStatusInput")?.value || "",
    amount: Number(document.getElementById("launchAmountInput")?.value || 0),
    summary: document.getElementById("launchSummaryInput")?.value.trim() || "",
    notes: document.getElementById("launchNotesInput")?.value.trim() || ""
  };
}

function renderTable(entry) {
  const body = document.getElementById("launchTableBody");
  const total = document.getElementById("launchTableGrandTotal");
  if (!body || !total) return;

  body.innerHTML = entry.lineItems
    .map(
      (item, index) => `
        <tr data-line-index="${index}">
          <td><input type="text" value="${item.description}" data-field="description" /></td>
          <td><input type="text" value="${item.owner}" data-field="owner" /></td>
          <td><input type="number" min="0" step="0.01" value="${item.quantity}" data-field="quantity" /></td>
          <td><input type="number" min="0" step="0.01" value="${item.unitPrice}" data-field="unitPrice" /></td>
          <td data-line-total>${currency(lineTotal(item))}</td>
        </tr>
      `
    )
    .join("");

  total.textContent = currency(entry.lineItems.reduce((sum, item) => sum + lineTotal(item), 0));

  body.querySelectorAll("input").forEach((input) => {
    input.addEventListener("input", handleLineItemChange);
  });
}

function handleLineItemChange(event) {
  const row = event.target.closest("tr[data-line-index]");
  if (!row || !draftEntry) return;
  const index = Number(row.dataset.lineIndex);
  const field = event.target.dataset.field;
  if (!draftEntry.lineItems[index] || !field) return;

  draftEntry.lineItems[index][field] =
    field === "quantity" || field === "unitPrice" ? Number(event.target.value || 0) : event.target.value;

  const lineTotalCell = row.querySelector("[data-line-total]");
  if (lineTotalCell) lineTotalCell.textContent = currency(lineTotal(draftEntry.lineItems[index]));

  const grandTotal = document.getElementById("launchTableGrandTotal");
  if (grandTotal) {
    grandTotal.textContent = currency(draftEntry.lineItems.reduce((sum, item) => sum + lineTotal(item), 0));
  }
}

function syncDraftFromForm() {
  const headerValues = readHeaderFields();
  draftEntry = {
    ...headerValues,
    lineItems: draftEntry.lineItems.map((item) => ({ ...item }))
  };
  updateTitle(draftEntry);
}

function persistDraft(message, shouldReturn = false) {
  syncDraftFromForm();
  const totalFromTable = draftEntry.lineItems.reduce((sum, item) => sum + lineTotal(item), 0);
  draftEntry.amount = Number(draftEntry.amount || totalFromTable);
  persistedEntry = saveLaunch(draftEntry);
  draftEntry = createEditableDraft(persistedEntry);
  fillHeaderFields(draftEntry);
  updateTitle(draftEntry);
  renderTable(draftEntry);
  showMessage(document.getElementById("launchDetailMessage"), message, "success");

  if (shouldReturn) {
    window.location.href = getLaunchListRoute();
  }
}

function bindActions() {
  document.getElementById("launchSaveButton")?.addEventListener("click", () => {
    persistDraft("Launch entry saved successfully.");
  });

  document.getElementById("launchOkButton")?.addEventListener("click", () => {
    persistDraft("Launch entry saved and confirmed.", true);
  });

  document.getElementById("launchCancelButton")?.addEventListener("click", () => {
    draftEntry = createEditableDraft(persistedEntry);
    fillHeaderFields(draftEntry);
    renderTable(draftEntry);
    updateTitle(draftEntry);
    showMessage(document.getElementById("launchDetailMessage"), "Unsaved changes were cancelled.", "neutral");
  });

  document.getElementById("launchBackButton")?.addEventListener("click", () => {
    window.location.href = getLaunchListRoute();
  });

  document.querySelectorAll("[data-sync-launch]").forEach((input) => {
    input.addEventListener("input", () => {
      syncDraftFromForm();
      showMessage(document.getElementById("launchDetailMessage"), "Unsaved changes in progress.", "neutral");
    });
  });
}

function renderNavigation(entry) {
  const { previous, next } = getLaunchSequence(entry.id);
  const previousLink = document.getElementById("launchPreviousLink");
  const nextLink = document.getElementById("launchNextLink");

  if (previousLink) {
    if (previous) {
      previousLink.href = getLaunchRoute(previous.id);
      previousLink.innerHTML = `<strong>Previous</strong><span>${previous.title}</span>`;
      previousLink.hidden = false;
    } else {
      previousLink.hidden = true;
    }
  }

  if (nextLink) {
    if (next) {
      nextLink.href = getLaunchRoute(next.id);
      nextLink.innerHTML = `<strong>Next</strong><span>${next.title}</span>`;
      nextLink.hidden = false;
    } else {
      nextLink.hidden = true;
    }
  }
}

function initSearch() {
  mountGlobalSearch({
    input: document.getElementById("launchGlobalSearch"),
    results: document.getElementById("launchSearchResults"),
    emptyState: document.getElementById("launchSearchEmpty")
  });
}

function init() {
  const id = getRouteId();
  const entry = getLaunchById(id);
  setBackLinks(document);
  initSearch();

  const notFound = document.getElementById("launchNotFound");
  const detail = document.getElementById("launchDetail");
  if (!entry) {
    if (detail) detail.hidden = true;
    if (notFound) notFound.hidden = false;
    return;
  }

  if (detail) detail.hidden = false;
  if (notFound) notFound.hidden = true;

  persistedEntry = createEditableDraft(entry);
  draftEntry = createEditableDraft(entry);

  updateTitle(draftEntry);
  fillHeaderFields(draftEntry);
  renderTable(draftEntry);
  renderNavigation(draftEntry);
  bindActions();
}

init();
