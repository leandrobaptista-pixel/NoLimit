import { getLaunchListRoute, getLaunchRoute, searchLaunches } from "./launches-store.js";

export function currency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 2
  }).format(Number(value || 0));
}

export function lineTotal(item) {
  return Number(item.quantity || 0) * Number(item.unitPrice || 0);
}

export function mountGlobalSearch({ input, results, emptyState }) {
  if (!input || !results) return;

  const hideResults = () => {
    results.innerHTML = "";
    results.hidden = true;
    if (emptyState) emptyState.hidden = true;
  };

  const renderResults = (items) => {
    results.innerHTML = "";
    if (!items.length) {
      results.hidden = true;
      if (emptyState) emptyState.hidden = false;
      return;
    }

    if (emptyState) emptyState.hidden = true;
    results.hidden = false;

    items.forEach((item) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "launch-search-result";
      button.innerHTML = `
        <strong>${item.title}</strong>
        <span>${item.category} · ${item.entryDate} · ${item.client}</span>
      `;
      button.addEventListener("click", () => {
        window.location.href = getLaunchRoute(item.id);
      });
      results.appendChild(button);
    });
  };

  input.addEventListener("input", () => {
    const query = input.value.trim();
    if (!query) {
      hideResults();
      return;
    }

    renderResults(searchLaunches(query));
  });

  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      const match = searchLaunches(input.value.trim(), 1)[0];
      if (match) {
        window.location.href = getLaunchRoute(match.id);
      }
    }

    if (event.key === "Escape") {
      hideResults();
      input.blur();
    }
  });

  document.addEventListener("click", (event) => {
    if (event.target === input || results.contains(event.target)) return;
    hideResults();
  });
}

export function setBackLinks(root = document) {
  root.querySelectorAll("[data-launch-list-link]").forEach((link) => {
    link.setAttribute("href", getLaunchListRoute());
  });
}

export function showMessage(target, message, tone = "neutral") {
  if (!target) return;
  target.textContent = message;
  target.dataset.tone = tone;
  target.hidden = !message;
}

export function formatEntryMeta(entry) {
  return `${entry.category} · ${entry.entryDate} · ${entry.client}`;
}
