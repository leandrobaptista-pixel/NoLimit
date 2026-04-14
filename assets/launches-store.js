import { LAUNCHES_SEED } from "./launches-data.js";

const STORAGE_KEY = "nolimit.website2.launches.v1";

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function ensureLineTotals(entry) {
  return {
    ...entry,
    amount: Number(entry.amount || 0),
    lineItems: Array.isArray(entry.lineItems)
      ? entry.lineItems.map((item) => ({
          ...item,
          quantity: Number(item.quantity || 0),
          unitPrice: Number(item.unitPrice || 0)
        }))
      : []
  };
}

function normalizeList(list) {
  return (Array.isArray(list) ? list : []).map((entry) => ensureLineTotals(entry));
}

function sortEntries(list) {
  return [...list].sort((left, right) => {
    const byDate = String(right.entryDate || "").localeCompare(String(left.entryDate || ""));
    if (byDate !== 0) return byDate;
    return String(left.id || "").localeCompare(String(right.id || ""));
  });
}

function saveRaw(list) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(normalizeList(list)));
}

function bootSeed() {
  try {
    const existing = localStorage.getItem(STORAGE_KEY);
    if (existing) return;
    saveRaw(LAUNCHES_SEED);
  } catch {
    // Keep seed in memory only if storage is unavailable.
  }
}

function loadRaw() {
  bootSeed();
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    const normalized = normalizeList(parsed);
    if (normalized.length) return normalized;
  } catch {
    // Fall back to bundled seed.
  }
  return normalizeList(LAUNCHES_SEED);
}

function textIndex(entry) {
  return [
    entry.id,
    entry.title,
    entry.category,
    entry.entryDate,
    entry.client,
    entry.location,
    entry.status,
    entry.summary,
    entry.notes,
    ...(entry.lineItems || []).flatMap((item) => [item.description, item.owner])
  ]
    .join(" ")
    .toLowerCase();
}

export function getLaunches(filters = {}) {
  const { query = "", category = "", date = "" } = filters;
  const normalizedQuery = String(query).trim().toLowerCase();
  const normalizedCategory = String(category).trim().toLowerCase();
  const normalizedDate = String(date).trim();

  return sortEntries(loadRaw()).filter((entry) => {
    if (normalizedCategory && String(entry.category || "").toLowerCase() !== normalizedCategory) return false;
    if (normalizedDate && String(entry.entryDate || "") !== normalizedDate) return false;
    if (normalizedQuery && !textIndex(entry).includes(normalizedQuery)) return false;
    return true;
  });
}

export function getLaunchById(id) {
  if (!id) return null;
  return sortEntries(loadRaw()).find((entry) => entry.id === id) || null;
}

export function getLaunchCategories() {
  return Array.from(new Set(loadRaw().map((entry) => entry.category).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b)
  );
}

export function saveLaunch(updatedLaunch) {
  const launches = loadRaw();
  const index = launches.findIndex((entry) => entry.id === updatedLaunch.id);
  const normalized = ensureLineTotals(updatedLaunch);

  if (index >= 0) {
    launches[index] = normalized;
  } else {
    launches.push(normalized);
  }

  saveRaw(launches);
  return normalized;
}

export function getLaunchSequence(id) {
  const launches = sortEntries(loadRaw());
  const index = launches.findIndex((entry) => entry.id === id);
  return {
    launches,
    index,
    previous: index > 0 ? launches[index - 1] : null,
    next: index >= 0 && index < launches.length - 1 ? launches[index + 1] : null
  };
}

export function searchLaunches(query, limit = 8) {
  if (!String(query).trim()) return [];
  return getLaunches({ query }).slice(0, limit);
}

export function getLaunchRoute(id) {
  return `/lancamento/${encodeURIComponent(id)}`;
}

export function getLaunchListRoute() {
  return "/lancamentos/";
}

export function resetLaunchDraft(id) {
  return clone(getLaunchById(id));
}

export function createEditableDraft(entry) {
  return clone(ensureLineTotals(entry));
}
