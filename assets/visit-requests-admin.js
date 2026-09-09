const REQUEST_SESSION_KEY = 'nolimitAdminPortalSession:v1';
const REQUEST_AREAS = ['Trim', 'Wainscoting', 'Stairs', 'Ceiling', 'Decks', 'Kitchen & Vanities', 'Fireplaces & Bars', 'Outside Doors & Windows', 'Pergola', 'Port & Portal', 'Commercial', 'Trash Container', 'Wall Paneling'];
const statusEl = document.getElementById('visitRequestsStatus');
const metricsEl = document.getElementById('visitRequestMetrics');
const tableEl = document.getElementById('visitRequestsTable');
const emptyEl = document.getElementById('visitRequestsEmpty');
const countEl = document.getElementById('visitRequestCount');

function getConfig() {
  const config = window.CABINETS_SYNC || {};
  const supabaseUrl = String(config.supabaseUrl || '').trim().replace(/\/$/, '');
  const key = String(config.supabaseAnonKey || '').trim();
  const tenant = String(config.tenant || '').trim();
  return supabaseUrl && key && tenant ? { supabaseUrl, key, tenant } : null;
}
function escapeHtml(value) { const div = document.createElement('div'); div.textContent = String(value || '—'); return div.innerHTML; }
function readSession() { try { return JSON.parse(sessionStorage.getItem(REQUEST_SESSION_KEY) || 'null'); } catch { return null; } }
function formatDate(value) { const date = new Date(value); return Number.isNaN(date.getTime()) ? '—' : date.toLocaleString(); }

function renderMetrics(rows) {
  const counts = Object.fromEntries(REQUEST_AREAS.map((area) => [area, 0]));
  let other = 0;
  rows.forEach((row) => { const type = String(row.payload?.projectType || '').trim(); if (counts[type] !== undefined) counts[type] += 1; else if (type) other += 1; });
  metricsEl.innerHTML = REQUEST_AREAS.map((area) => `<article class="card visit-request-metric"><strong>${escapeHtml(area)}</strong><span>${counts[area]}</span></article>`).join('') + `<article class="card visit-request-metric"><strong>Other</strong><span>${other}</span></article>`;
}
function renderRows(rows) {
  countEl.textContent = `${rows.length} request${rows.length === 1 ? '' : 's'}`;
  emptyEl.hidden = rows.length > 0;
  tableEl.innerHTML = rows.map((row) => { const p = row.payload || {}; return `<article class="visit-request-row"><div><p class="section-kicker">${escapeHtml(p.projectType || 'Other')} · ${escapeHtml(formatDate(p.submittedAt || row.updated_at))}</p><h3>${escapeHtml(p.name)}</h3><p class="hint">${escapeHtml(p.email)}${p.phone ? ` · ${escapeHtml(p.phone)}` : ''}</p></div><div><strong>Location</strong><p>${escapeHtml([p.address, p.city].filter(Boolean).join(', ') || 'Not provided')}</p></div><div><strong>Preferred date</strong><p>${escapeHtml(p.preferredDate || 'Not provided')}</p></div><div class="visit-request-details"><strong>Project details</strong><p>${escapeHtml(p.details || 'No details provided')}</p></div></article>`; }).join('');
}
async function loadRequests() {
  const session = readSession();
  if (!session?.id) { statusEl.textContent = 'Please sign in through the Admin Portal first.'; statusEl.classList.add('is-error'); return; }
  const config = getConfig();
  if (!config) { statusEl.textContent = 'The request database is not configured on this device.'; statusEl.classList.add('is-error'); return; }
  statusEl.textContent = 'Loading visit requests…'; statusEl.classList.remove('is-error');
  const query = new URLSearchParams({ select: 'id,updated_at,payload', tenant: `eq.${config.tenant}`, kind: 'eq.visitRequest', order: 'updated_at.desc', limit: '500' });
  try { const response = await fetch(`${config.supabaseUrl}/rest/v1/app_records?${query}`, { headers: { apikey: config.key, Authorization: `Bearer ${config.key}` }, cache: 'no-store' }); if (!response.ok) throw new Error(`HTTP ${response.status}`); const rows = await response.json(); renderMetrics(rows); renderRows(rows); statusEl.textContent = 'Up to date.'; } catch { statusEl.textContent = 'Requests could not be loaded right now. Please try again.'; statusEl.classList.add('is-error'); }
}
document.getElementById('refreshVisitRequests')?.addEventListener('click', loadRequests);
loadRequests();
