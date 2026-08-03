const adminPortalForm = document.getElementById('adminPortalLoginForm');
const adminPortalUsername = document.getElementById('adminPortalUsername');
const adminPortalPassword = document.getElementById('adminPortalPassword');
const adminPortalStatus = document.getElementById('adminPortalStatus');
const adminToolsSection = document.getElementById('adminToolsSection');
const adminSessionMeta = document.getElementById('adminSessionMeta');
const adminPortalLogout = document.getElementById('adminPortalLogout');
const adminToolCards = Array.from(document.querySelectorAll('[data-tool]'));
const adminNavWrap = document.querySelector('.nav');
const adminNavMenu = document.getElementById('primaryNav');
const adminNavToggle = document.querySelector('.menu-toggle');
const adminHeader = document.querySelector('.site-header');

const ADMIN_PORTAL_SESSION_KEY = 'nolimitAdminPortalSession:v1';
const ADMIN_COMPANY_MATCH = 'no limit contractor';
const ADMIN_SESSION_MAX_AGE_MS = 12 * 60 * 60 * 1000;
const ADMIN_TOOL_QUERY = new URLSearchParams(window.location.search).get('tool') || '';
const AUTH_USER_SELECT = [
  'id',
  'updated_at',
  'userId:payload->>id',
  'name:payload->>name',
  'firstName:payload->>firstName',
  'lastName:payload->>lastName',
  'companyName:payload->>companyName',
  'jobTitle:payload->>jobTitle',
  'email:payload->>email',
  'username:payload->>username',
  'passwordHash:payload->>passwordHash',
  'legacyPassword:payload->>legacyPassword',
  'legacyPlainPassword:payload->>password',
  'systemRole:payload->>systemRole',
  'accessProfile:payload->>accessProfile',
  'updatedAt:payload->>updatedAt'
].join(',');

function syncHeaderHeight() {
  if (!adminHeader) return;
  const height = Math.ceil(adminHeader.getBoundingClientRect().height);
  document.documentElement.style.setProperty('--header-h', `${height}px`);
}

function closeMobileNav() {
  adminNavToggle?.setAttribute('aria-expanded', 'false');
  adminNavWrap?.classList.remove('nav-open');
  syncHeaderHeight();
}

function setStatus(message, type = 'info') {
  if (!adminPortalStatus) return;
  adminPortalStatus.textContent = message;
  adminPortalStatus.classList.remove('is-error', 'is-success');
  if (type === 'error') adminPortalStatus.classList.add('is-error');
  if (type === 'success') adminPortalStatus.classList.add('is-success');
}

function normalizeText(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function companyIsAuthorized(companyName) {
  return normalizeText(companyName).includes(ADMIN_COMPANY_MATCH);
}

function isSha256Hash(value) {
  return /^[a-f0-9]{64}$/i.test(String(value || '').trim());
}

async function hashPassword(password) {
  const encoded = new TextEncoder().encode(String(password || ''));
  const buffer = await crypto.subtle.digest('SHA-256', encoded);
  return Array.from(new Uint8Array(buffer))
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');
}

function userPasswordMatches(user, plainPassword, passwordHash) {
  const stored = String(user?.passwordHash || '').trim();
  const legacy = String(user?.legacyPassword || user?.password || '').trim();
  const inputHash = String(passwordHash || '').trim().toLowerCase();

  if (stored) {
    if (isSha256Hash(stored)) return stored.toLowerCase() === inputHash;
    return stored === plainPassword || stored.toLowerCase() === inputHash;
  }

  if (!legacy) return false;
  return legacy === plainPassword || legacy.toLowerCase() === inputHash;
}

function getSyncConfig() {
  const config = window.CABINETS_SYNC || {};
  const supabaseUrl = String(config.supabaseUrl || '').trim().replace(/\/$/, '');
  const supabaseAnonKey = String(config.supabaseAnonKey || '').trim();
  const tenant = String(config.tenant || '').trim();
  if (!supabaseUrl || !supabaseAnonKey || !tenant) return null;
  return { supabaseUrl, supabaseAnonKey, tenant };
}

function syncEndpoint() {
  const config = getSyncConfig();
  if (!config) return '';
  return `${config.supabaseUrl}/rest/v1/app_records`;
}

function authDirectoryUnavailableMessage(code = '') {
  if (code === 'no-config') return 'System database is not available on this device.';
  if (code === 'offline') return 'Connect to the internet to reach the central user directory.';
  if (code === 'timeout') return 'The central user directory took too long to respond.';
  if (String(code || '').startsWith('http-')) return `Central user directory request failed (${String(code).replace('http-', 'HTTP ')}).`;
  return 'Could not reach the central user directory right now.';
}

function normalizeDirectoryUser(row = {}) {
  const payload = row?.payload && typeof row.payload === 'object' ? row.payload : row;
  const name = String(payload.name || `${payload.firstName || ''} ${payload.lastName || ''}` || '').trim();
  return {
    id: String(payload.id || payload.userId || row.id || '').trim(),
    name,
    firstName: String(payload.firstName || '').trim(),
    lastName: String(payload.lastName || '').trim(),
    companyName: String(payload.companyName || '').trim(),
    jobTitle: String(payload.jobTitle || '').trim(),
    email: String(payload.email || '').trim(),
    username: String(payload.username || '').trim().toLowerCase(),
    passwordHash: String(payload.passwordHash || '').trim(),
    legacyPassword: String(payload.legacyPassword || payload.legacyPlainPassword || payload.password || '').trim(),
    systemRole: String(payload.systemRole || '').trim(),
    accessProfile: String(payload.accessProfile || '').trim(),
    updatedAt: String(payload.updatedAt || row.updated_at || '').trim()
  };
}

async function fetchCloudUsersByUsername(username, timeoutMs = 12000) {
  const config = getSyncConfig();
  const endpoint = syncEndpoint();
  if (!config || !endpoint) {
    return { ok: false, code: 'no-config', users: [], message: authDirectoryUnavailableMessage('no-config') };
  }
  if (!navigator.onLine) {
    return { ok: false, code: 'offline', users: [], message: authDirectoryUnavailableMessage('offline') };
  }

  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), timeoutMs);

  try {
    const qs = new URLSearchParams({
      select: AUTH_USER_SELECT,
      tenant: `eq.${config.tenant}`,
      kind: 'eq.user',
      order: 'updated_at.desc',
      limit: '20'
    });
    qs.set('payload->>username', `eq.${String(username || '').trim().toLowerCase()}`);

    const response = await fetch(`${endpoint}?${qs.toString()}`, {
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`
      },
      cache: 'no-store',
      signal: controller.signal
    });

    if (!response.ok) {
      return {
        ok: false,
        code: `http-${response.status}`,
        users: [],
        message: authDirectoryUnavailableMessage(`http-${response.status}`)
      };
    }

    const rows = await response.json();
    const users = Array.isArray(rows)
      ? rows.map((entry) => normalizeDirectoryUser(entry)).filter((entry) => entry.id && entry.username)
      : [];
    return { ok: true, code: 'ok', users, message: users.length ? '' : 'No matching user was found in the central directory.' };
  } catch (error) {
    const code = error?.name === 'AbortError' ? 'timeout' : 'network';
    return { ok: false, code, users: [], message: authDirectoryUnavailableMessage(code) };
  } finally {
    window.clearTimeout(timer);
  }
}

function saveAdminSession(user) {
  const session = {
    id: user.id,
    username: user.username,
    name: user.name || user.username,
    companyName: user.companyName || '',
    jobTitle: user.jobTitle || '',
    systemRole: user.systemRole || '',
    accessProfile: user.accessProfile || '',
    authenticatedAt: new Date().toISOString()
  };
  sessionStorage.setItem(ADMIN_PORTAL_SESSION_KEY, JSON.stringify(session));
  return session;
}

function readAdminSession() {
  try {
    const raw = sessionStorage.getItem(ADMIN_PORTAL_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    const authenticatedAt = new Date(parsed.authenticatedAt || '').getTime();
    if (!authenticatedAt || Date.now() - authenticatedAt > ADMIN_SESSION_MAX_AGE_MS) {
      sessionStorage.removeItem(ADMIN_PORTAL_SESSION_KEY);
      return null;
    }
    return parsed;
  } catch {
    sessionStorage.removeItem(ADMIN_PORTAL_SESSION_KEY);
    return null;
  }
}

function clearAdminSession() {
  sessionStorage.removeItem(ADMIN_PORTAL_SESSION_KEY);
}

function highlightRequestedTool() {
  adminToolCards.forEach((card) => {
    const isTarget = ADMIN_TOOL_QUERY && card.dataset.tool === ADMIN_TOOL_QUERY;
    card.classList.toggle('is-target', Boolean(isTarget));
  });
}

function showTools(session) {
  if (adminToolsSection) adminToolsSection.hidden = false;
  highlightRequestedTool();

  if (adminSessionMeta) {
    const role = session.systemRole || session.accessProfile || 'registered user';
    const title = session.jobTitle ? ` • ${session.jobTitle}` : '';
    adminSessionMeta.textContent = `Signed in as ${session.name} (${role})${title} • ${session.companyName || 'No Limit Contractor'}`;
  }
}

function hideTools() {
  if (adminToolsSection) adminToolsSection.hidden = true;
}

async function handleAdminLogin(event) {
  event.preventDefault();
  const username = String(adminPortalUsername?.value || '').trim().toLowerCase();
  const plainPassword = String(adminPortalPassword?.value || '');

  if (!username || !plainPassword) {
    setStatus('Enter username and password to continue.', 'error');
    return;
  }

  const submitButton = adminPortalForm?.querySelector('button[type="submit"]');
  submitButton?.setAttribute('disabled', 'disabled');
  setStatus('Checking the central No Limit directory...');

  try {
    const result = await fetchCloudUsersByUsername(username);
    if (!result.ok) {
      setStatus(result.message || 'Could not reach the central user directory.', 'error');
      return;
    }

    const allowedUsers = result.users.filter((user) => companyIsAuthorized(user.companyName));
    if (!allowedUsers.length) {
      setStatus('This account is not registered under No Limit Contractor.', 'error');
      return;
    }

    const passwordHash = await hashPassword(plainPassword);
    const matchedUser = allowedUsers.find((user) => userPasswordMatches(user, plainPassword, passwordHash)) || null;
    if (!matchedUser) {
      setStatus('Invalid username or password.', 'error');
      return;
    }

    const session = saveAdminSession(matchedUser);
    adminPortalForm.reset();
    showTools(session);
    setStatus('Access granted. Admin tools are now available below.', 'success');
    adminToolsSection?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } finally {
    submitButton?.removeAttribute('disabled');
  }
}

function restoreExistingSession() {
  const session = readAdminSession();
  if (!session) {
    hideTools();
    return;
  }
  showTools(session);
  setStatus(`Welcome back, ${session.name}.`, 'success');
}

adminPortalForm?.addEventListener('submit', handleAdminLogin);

adminPortalLogout?.addEventListener('click', () => {
  clearAdminSession();
  hideTools();
  setStatus('You signed out of the Admin portal.', 'success');
  adminPortalUsername?.focus();
});

adminNavToggle?.addEventListener('click', () => {
  const expanded = adminNavToggle.getAttribute('aria-expanded') === 'true';
  adminNavToggle.setAttribute('aria-expanded', expanded ? 'false' : 'true');
  adminNavWrap?.classList.toggle('nav-open', !expanded);
  syncHeaderHeight();
});

adminNavMenu?.querySelectorAll('a').forEach((link) => {
  link.addEventListener('click', () => {
    if (window.innerWidth <= 980) closeMobileNav();
  });
});

window.addEventListener('resize', syncHeaderHeight);
window.addEventListener('load', () => {
  syncHeaderHeight();
  restoreExistingSession();
});
