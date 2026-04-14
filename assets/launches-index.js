import { getLaunchCategories, getLaunchListRoute, getLaunchRoute, getLaunches } from "./launches-store.js";
import { currency, formatEntryMeta, mountGlobalSearch, setBackLinks } from "./launches-ui.js";

const state = {
  query: "",
  category: "",
  date: ""
};

function renderSummary(items) {
  const totalEl = document.getElementById("launchSummaryTotal");
  const amountEl = document.getElementById("launchSummaryAmount");
  const categoryEl = document.getElementById("launchSummaryCategories");

  const totalAmount = items.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  if (totalEl) totalEl.textContent = String(items.length);
  if (amountEl) amountEl.textContent = currency(totalAmount);
  if (categoryEl) categoryEl.textContent = String(new Set(items.map((item) => item.category)).size);
}

function renderCards(items) {
  const container = document.getElementById("launchCards");
  if (!container) return;

  if (!items.length) {
    container.innerHTML = `
      <article class="launch-card">
        <h4>No entries found</h4>
        <p>Adjust the filters or search for another launch entry.</p>
      </article>
    `;
    return;
  }

  container.innerHTML = items
    .map(
      (item) => `
        <article class="launch-card">
          <div class="launch-card-top">
            <div>
              <h4>${item.title}</h4>
              <p class="launch-card-meta">${formatEntryMeta(item)}</p>
            </div>
            <span class="launch-pill">${item.status}</span>
          </div>
          <p>${item.summary}</p>
          <div class="launch-card-top">
            <strong>${currency(item.amount)}</strong>
            <span>${item.location}</span>
          </div>
          <div class="launch-card-actions">
            <a class="btn" href="${getLaunchRoute(item.id)}">Open entry</a>
            <a class="btn btn-secondary" href="${getLaunchRoute(item.id)}#table">Open table</a>
          </div>
        </article>
      `
    )
    .join("");
}

function renderList() {
  const items = getLaunches(state);
  renderSummary(items);
  renderCards(items);
}

function bindFilters() {
  const queryInput = document.getElementById("launchFilterQuery");
  const categoryInput = document.getElementById("launchFilterCategory");
  const dateInput = document.getElementById("launchFilterDate");
  const resetButton = document.getElementById("launchFilterReset");

  if (queryInput) {
    queryInput.addEventListener("input", () => {
      state.query = queryInput.value.trim();
      renderList();
    });
  }

  if (categoryInput) {
    categoryInput.innerHTML =
      `<option value="">All categories</option>` +
      getLaunchCategories()
        .map((category) => `<option value="${category}">${category}</option>`)
        .join("");

    categoryInput.addEventListener("change", () => {
      state.category = categoryInput.value;
      renderList();
    });
  }

  if (dateInput) {
    dateInput.addEventListener("change", () => {
      state.date = dateInput.value;
      renderList();
    });
  }

  resetButton?.addEventListener("click", () => {
    state.query = "";
    state.category = "";
    state.date = "";
    if (queryInput) queryInput.value = "";
    if (categoryInput) categoryInput.value = "";
    if (dateInput) dateInput.value = "";
    renderList();
  });
}

function initSearch() {
  mountGlobalSearch({
    input: document.getElementById("launchGlobalSearch"),
    results: document.getElementById("launchSearchResults"),
    emptyState: document.getElementById("launchSearchEmpty")
  });
}

function init() {
  setBackLinks(document);
  initSearch();
  bindFilters();
  renderList();

  const backHome = document.getElementById("launchBackHome");
  if (backHome) {
    backHome.setAttribute("href", "/");
  }

  const breadcrumb = document.getElementById("launchBreadcrumbList");
  if (breadcrumb) {
    breadcrumb.setAttribute("href", getLaunchListRoute());
  }
}

init();
