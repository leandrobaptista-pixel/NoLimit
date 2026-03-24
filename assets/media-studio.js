const studioYear = document.getElementById('studioYear');
if (studioYear) studioYear.textContent = new Date().getFullYear();

const studioAlert = document.getElementById('studioAlert');
const accessForm = document.getElementById('studioAccessForm');
const accessModeChip = document.getElementById('studioAccessModeChip');
const accessStatus = document.getElementById('studioAccessStatus');
const simulateUploadButton = document.getElementById('studioSimulateUploadButton');
const apiUrlInput = document.getElementById('studioApiUrlInput');
const adminTokenInput = document.getElementById('studioAdminTokenInput');
const unlockButton = document.getElementById('studioUnlockButton');
const testModeButton = document.getElementById('studioTestModeButton');
const lockButton = document.getElementById('studioLockButton');
const uploadForm = document.getElementById('studioUploadForm');
const uploadButton = document.getElementById('studioUploadButton');
const uploadPanel = document.getElementById('studioUploadPanel');
const refreshButton = document.getElementById('studioRefreshButton');
const categorySelect = document.getElementById('studioCategorySelect');
const categoryFilter = document.getElementById('studioCategoryFilter');
const searchInput = document.getElementById('studioSearchInput');
const publishedFilter = document.getElementById('studioPublishedFilter');
const tableBody = document.getElementById('studioTableBody');
const countBadge = document.getElementById('studioCountBadge');
const apiBadge = document.getElementById('studioApiBadge');
const metricTotal = document.getElementById('metricTotal');
const metricPublished = document.getElementById('metricPublished');
const metricGenerated = document.getElementById('metricGenerated');
const metricSocial = document.getElementById('metricSocial');
const studioPostsBadge = document.getElementById('studioPostsBadge');
const studioReadyTodayCount = document.getElementById('studioReadyTodayCount');
const studioSentTodayCount = document.getElementById('studioSentTodayCount');
const studioNeedsArtCount = document.getElementById('studioNeedsArtCount');
const studioReadyQueue = document.getElementById('studioReadyQueue');
const studioSentQueue = document.getElementById('studioSentQueue');

const previewTitle = document.getElementById('studioPreviewTitle');
const originalPreview = document.getElementById('studioOriginalPreview');
const generatedPreview = document.getElementById('studioGeneratedPreview');
const metaCategory = document.getElementById('studioMetaCategory');
const metaCreated = document.getElementById('studioMetaCreated');
const metaSocial = document.getElementById('studioMetaSocial');
const captionField = document.getElementById('studioCaption');

const state = {
  apiBaseUrl: '',
  adminToken: '',
  mode: 'idle',
  canManage: false,
  categories: [],
  media: [],
  filters: {
    search: '',
    category: '',
    published: 'all'
  },
  selectedId: '',
  busyId: '',
  connected: false
};

const DEMO_CATEGORIES = [
  { id: 'trim', name: 'Trim', slug: 'trim' },
  { id: 'kitchens', name: 'Kitchens', slug: 'kitchens' },
  { id: 'decks', name: 'Decks', slug: 'decks' },
  { id: 'stairs', name: 'Stairs', slug: 'stairs' },
  { id: 'wainscoting', name: 'Wainscoting', slug: 'wainscoting' }
];

function getMediaStudioConfig() {
  return window.MEDIA_STUDIO || {};
}

function getUrlParams() {
  return new URLSearchParams(window.location.search || '');
}

function getRequestedEntry() {
  return String(getUrlParams().get('entry') || '').trim().toLowerCase();
}

function isUploadEntryRequested() {
  return getRequestedEntry() === 'upload';
}

function isLocalHost() {
  return ['localhost', '127.0.0.1'].includes(window.location.hostname);
}

function normalizeApiBaseUrl(value) {
  return String(value || '').trim().replace(/\/$/, '');
}

function getApiUrlStorageKey() {
  return String(getMediaStudioConfig().apiUrlStorageKey || 'nolimit_gallery_control_api_url').trim();
}

function getAuthStorageKey() {
  return String(getMediaStudioConfig().authStorageKey || 'nolimit_gallery_control_admin_token').trim();
}

function getDemoStorageKey() {
  return 'nolimit_gallery_control_demo_media';
}

function getDemoModeStorageKey() {
  return 'nolimit_gallery_control_demo_mode';
}

function readStoredValue(key) {
  try {
    return localStorage.getItem(key) || '';
  } catch {
    return '';
  }
}

function writeStoredValue(key, value) {
  try {
    if (value) {
      localStorage.setItem(key, value);
    } else {
      localStorage.removeItem(key);
    }
  } catch {
    // Ignore storage failures.
  }
}

function getApiBaseUrl() {
  const config = getMediaStudioConfig();
  const stored = normalizeApiBaseUrl(readStoredValue(getApiUrlStorageKey()));
  if (stored) return stored;

  const direct = normalizeApiBaseUrl(config.apiBaseUrl);
  if (direct) return direct;

  if (isLocalHost()) {
    const local = normalizeApiBaseUrl(config.localApiBaseUrl);
    if (local) return local;
  }

  return normalizeApiBaseUrl(config.productionApiBaseUrl);
}

function getStoredAdminToken() {
  return String(readStoredValue(getAuthStorageKey()) || '').trim();
}

function isDemoModeStored() {
  return readStoredValue(getDemoModeStorageKey()) === 'true';
}

function normalizeTableList(value, fallback) {
  const list = Array.isArray(value) ? value : value ? [value] : fallback;
  return list.map((item) => String(item || '').trim()).filter(Boolean);
}

function slugify(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function getWebsiteContentConfig() {
  const sharedConfig = window.CONTACT_FORM || {};
  const config = window.WEBSITE_CONTENT || {};
  const supabaseUrl = String(config.supabaseUrl || sharedConfig.supabaseUrl || '').trim().replace(/\/$/, '');
  const supabaseAnonKey = String(config.supabaseAnonKey || sharedConfig.supabaseAnonKey || '').trim();

  if (!supabaseUrl || !supabaseAnonKey) return null;

  return {
    supabaseUrl,
    supabaseAnonKey,
    categoriesTables: normalizeTableList(config.categoriesTables || config.categoriesTable, ['website_categories', 'categories']),
    galleryItemsTables: normalizeTableList(config.galleryItemsTables || config.galleryItemsTable, ['website_gallery_items', 'gallery_items', 'portfolio_items'])
  };
}

async function fetchSupabaseRows(tableNames, queryEntries = []) {
  const config = getWebsiteContentConfig();
  if (!config) return null;

  for (const table of tableNames) {
    const url = new URL(`${config.supabaseUrl}/rest/v1/${table}`);
    queryEntries.forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== '') url.searchParams.append(key, value);
    });

    try {
      const response = await fetch(url.toString(), {
        headers: {
          apikey: config.supabaseAnonKey,
          Authorization: `Bearer ${config.supabaseAnonKey}`,
          Accept: 'application/json'
        }
      });

      if (response.ok) {
        const data = await response.json();
        if (Array.isArray(data)) return data;
        if (data && typeof data === 'object') return [data];
        return [];
      }

      const errorText = await response.text().catch(() => '');
      const missingRelation =
        response.status === 404 ||
        /does not exist/i.test(errorText) ||
        /could not find the table/i.test(errorText) ||
        /relation/i.test(errorText);

      if (missingRelation) continue;
    } catch (error) {
      console.warn(`Gallery Control Supabase fetch failed for table "${table}".`, error);
    }
  }

  return null;
}

function normalizeCategoryRow(row) {
  if (!row || typeof row !== 'object') return null;
  const name = String(row.name || row.title || '').trim();
  const slug = String(row.slug || slugify(name || row.id)).trim();
  if (!name || !slug) return null;
  return {
    id: String(row.id ?? slug),
    name,
    slug
  };
}

function createCategoryMap(categories) {
  const byId = new Map();
  const bySlug = new Map();

  categories.forEach((category) => {
    byId.set(String(category.id), category);
    bySlug.set(category.slug, category);
  });

  return { byId, bySlug };
}

function normalizeMediaItem(row, categoriesById) {
  if (!row || typeof row !== 'object') return null;

  const imageUrl = String(row.original_url || row.originalUrl || row.image_url || row.imageUrl || '').trim();
  if (!imageUrl) return null;

  const categoryId = row.category_id ?? row.categoryId ?? '';
  const linkedCategory = categoryId ? categoriesById.get(String(categoryId)) : null;
  const categoryName =
    linkedCategory?.name ||
    (typeof row.category === 'object' ? row.category?.name : row.category) ||
    row.category_name ||
    row.categoryName ||
    'Uncategorized';
  const categorySlug =
    linkedCategory?.slug ||
    (typeof row.category === 'object' ? row.category?.slug : '') ||
    row.category_slug ||
    row.categorySlug ||
    slugify(categoryName) ||
    'uncategorized';

  return {
    id: String(row.id || imageUrl),
    title: String(row.title || 'Untitled media').trim() || 'Untitled media',
    categoryId: String(linkedCategory?.id || categoryId || categorySlug),
    categoryName: String(categoryName).trim() || 'Uncategorized',
    categorySlug: String(categorySlug).trim() || 'uncategorized',
    originalUrl: imageUrl,
    generatedArtUrl: String(row.generated_art_url || row.generatedArtUrl || '').trim(),
    createdAt: String(row.created_at || row.createdAt || '').trim(),
    published: Boolean(row.published),
    usedInSocial: Boolean(row.used_in_social || row.usedInSocial),
    captionText: String(row.caption_text || row.captionText || '').trim()
  };
}

function getBrandContactDetails() {
  const companyEmail =
    String(window.CONTACT_FORM?.email || '').trim() ||
    String(window.WEBSITE_CONTENT?.email || '').trim() ||
    'hello@nolimitcontractor.com';
  const companyPhone = String(window.CONTACT_FORM?.phone || '').trim() || '(732) 555-0178';

  return {
    phone: companyPhone,
    email: companyEmail
  };
}

function buildDemoCaption(categoryName) {
  const contact = getBrandContactDetails();
  return `${categoryName} by No Limit. Clean finish carpentry, ready for your next project. ${contact.phone} • ${contact.email}`;
}

function readDemoMedia() {
  try {
    const parsed = JSON.parse(localStorage.getItem(getDemoStorageKey()) || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeDemoMedia(records) {
  try {
    localStorage.setItem(getDemoStorageKey(), JSON.stringify(records));
  } catch (error) {
    console.warn('Could not persist Gallery Control demo media.', error);
  }
}

function getAvailableCategories() {
  return state.categories.length ? state.categories : DEMO_CATEGORIES;
}

function getCategoryDetailsById(categoryId) {
  const categories = getAvailableCategories();
  return categories.find((category) => String(category.id) === String(categoryId)) || categories[0] || DEMO_CATEGORIES[0];
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('Could not read the selected image.'));
    reader.readAsDataURL(file);
  });
}

function loadImage(source) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error(`Could not load image: ${source}`));
    image.src = source;
  });
}

async function generateDemoArtUrl(item) {
  const canvas = document.createElement('canvas');
  canvas.width = 1080;
  canvas.height = 1080;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('Could not start browser artwork generation.');

  const [sourceImage, logoImage, anniversaryImage] = await Promise.all([
    loadImage(item.originalUrl),
    loadImage('assets/brand.png').catch(() => null),
    loadImage('assets/anniversary-18.png').catch(() => null)
  ]);

  const scale = Math.max(canvas.width / sourceImage.width, canvas.height / sourceImage.height);
  const drawWidth = sourceImage.width * scale;
  const drawHeight = sourceImage.height * scale;
  const drawX = (canvas.width - drawWidth) / 2;
  const drawY = (canvas.height - drawHeight) / 2;

  context.fillStyle = '#0f0f0f';
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(sourceImage, drawX, drawY, drawWidth, drawHeight);

  const gradient = context.createLinearGradient(0, 0, 0, canvas.height);
  gradient.addColorStop(0, 'rgba(10,10,10,0.08)');
  gradient.addColorStop(0.55, 'rgba(10,10,10,0.22)');
  gradient.addColorStop(1, 'rgba(10,10,10,0.72)');
  context.fillStyle = gradient;
  context.fillRect(0, 0, canvas.width, canvas.height);

  if (logoImage) {
    context.drawImage(logoImage, 72, 68, 280, 120);
  }

  if (anniversaryImage) {
    context.drawImage(anniversaryImage, 902, 56, 120, 120);
  }

  context.fillStyle = '#ffffff';
  context.font = '700 76px Georgia, serif';
  context.fillText(item.categoryName, 72, 808);

  context.font = '500 34px Arial, sans-serif';
  context.fillStyle = 'rgba(255,255,255,0.92)';
  context.fillText('Premium finish carpentry with clean detail and dependable delivery.', 72, 874);

  const contact = getBrandContactDetails();
  context.font = '600 28px Arial, sans-serif';
  context.fillText(`${contact.phone}  |  ${contact.email}`, 72, 938);

  return canvas.toDataURL('image/jpeg', 0.92);
}

async function uploadDemoMedia(formData) {
  const file = formData.get('image');
  if (!(file instanceof File) || !file.size) {
    throw new Error('Choose an image file before uploading in test mode.');
  }

  const title = String(formData.get('title') || '').trim();
  if (!title) {
    throw new Error('Title is required in test mode too.');
  }

  const category = getCategoryDetailsById(String(formData.get('categoryId') || '').trim());
  const originalUrl = await readFileAsDataUrl(file);
  const record = {
    id: `demo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    title,
    categoryId: category.id,
    categoryName: category.name,
    categorySlug: category.slug,
    originalUrl,
    generatedArtUrl: '',
    createdAt: new Date().toISOString(),
    published: String(formData.get('published') || '') === 'on',
    usedInSocial: false,
    captionText: buildDemoCaption(category.name),
    demo: true
  };

  const nextRecords = [record, ...readDemoMedia()];
  writeDemoMedia(nextRecords);
  return record;
}

async function patchDemoMedia(item, action) {
  const records = readDemoMedia();
  const index = records.findIndex((entry) => entry.id === item.id);
  if (index === -1) throw new Error('This demo media item is no longer available.');

  const nextItem = { ...records[index] };

  if (action === 'generate') {
    nextItem.generatedArtUrl = await generateDemoArtUrl(nextItem);
  }

  if (action === 'publish') {
    nextItem.published = !Boolean(nextItem.published);
  }

  if (action === 'social') {
    nextItem.usedInSocial = !Boolean(nextItem.usedInSocial);
  }

  records[index] = nextItem;
  writeDemoMedia(records);
  return nextItem;
}

function buildUrl(path, params = {}) {
  const base = state.apiBaseUrl;
  const url = new URL(`${base}${path}`, window.location.origin);
  Object.entries(params).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '' || value === 'all') return;
    url.searchParams.set(key, value);
  });
  return url.toString();
}

function updateAccessPanel() {
  if (apiUrlInput) apiUrlInput.value = state.apiBaseUrl || '';
  if (lockButton) lockButton.disabled = !state.canManage && state.mode !== 'demo';
  if (testModeButton) testModeButton.disabled = state.mode === 'demo';

  if (state.mode === 'demo') {
    if (accessModeChip) accessModeChip.textContent = 'Test mode';
    if (accessStatus) {
      accessStatus.textContent =
        'Browser test mode is active. Uploads, artwork generation, publish status, and social tracking are being saved only on this device.';
    }
    if (adminTokenInput) {
      adminTokenInput.value = '';
      adminTokenInput.placeholder = 'Not required in test mode';
    }
    return;
  }

  if (state.canManage) {
    if (accessModeChip) accessModeChip.textContent = 'Management unlocked';
    if (accessStatus) {
      accessStatus.textContent =
        'Signed in successfully. Upload, artwork generation, publish status, and social tracking are unlocked for this browser session.';
    }
    if (adminTokenInput) {
      adminTokenInput.value = '';
      adminTokenInput.placeholder = 'Session is already active in this browser';
    }
    return;
  }

  if (adminTokenInput) {
    adminTokenInput.value = '';
    adminTokenInput.placeholder = 'Enter the admin password';
  }

  if (state.apiBaseUrl) {
    if (accessModeChip) accessModeChip.textContent = 'API locked';
    if (accessStatus) {
      accessStatus.textContent =
        'The secure media API is configured. Sign in with the admin password to unlock upload and publish controls, or stay in public tracking mode.';
    }
  } else {
    if (accessModeChip) accessModeChip.textContent = 'Public view';
    if (accessStatus) {
      accessStatus.textContent =
        'Public tracking is live. Add the secure media API URL and sign in with the admin password to unlock upload, artwork generation, and publishing.';
    }
  }
}

function setApiSession({ apiBaseUrl = '', adminToken = '' } = {}) {
  state.apiBaseUrl = normalizeApiBaseUrl(apiBaseUrl);
  state.adminToken = String(adminToken || '').trim();
  state.mode = state.apiBaseUrl && state.adminToken ? 'api' : 'public';

  writeStoredValue(getApiUrlStorageKey(), state.apiBaseUrl);
  writeStoredValue(getAuthStorageKey(), state.adminToken);
  writeStoredValue(getDemoModeStorageKey(), '');
  setManagementAvailability(Boolean(state.apiBaseUrl && state.adminToken));
  updateAccessPanel();
}

function clearAdminSession({ keepApiUrl = true } = {}) {
  const apiBaseUrl = keepApiUrl ? state.apiBaseUrl : '';
  setApiSession({ apiBaseUrl, adminToken: '' });
}

function enableDemoMode() {
  state.mode = 'demo';
  writeStoredValue(getDemoModeStorageKey(), 'true');
  setManagementAvailability(true);
  updateAccessPanel();
}

function disableDemoMode() {
  writeStoredValue(getDemoModeStorageKey(), '');
  state.mode = state.apiBaseUrl && state.adminToken ? 'api' : 'public';
  setManagementAvailability(Boolean(state.apiBaseUrl && state.adminToken));
  updateAccessPanel();
}

function focusUploadFlow() {
  const target = uploadPanel || uploadForm;
  target?.scrollIntoView({ behavior: 'smooth', block: 'start' });

  window.setTimeout(() => {
    uploadForm?.querySelector('input[name="title"]')?.focus();
  }, 160);
}

function applyRequestedEntryExperience() {
  if (!isUploadEntryRequested()) return;

  focusUploadFlow();

  if (state.mode === 'demo') {
    setAlert(
      'Final upload flow simulation is active. You can upload photos, generate artwork, and test publishing from this browser right now.',
      'success'
    );
    return;
  }

  if (state.canManage) {
    setAlert('Definitive upload flow is ready. You can start uploading photos now.', 'success');
    return;
  }

  setAlert(
    'This is the definitive upload route. Sign in to the media API to upload live, or start test mode to simulate the final workflow now.',
    'success'
  );
}

async function apiFetch(path, options = {}, params) {
  if (!state.apiBaseUrl) {
    throw new Error('Gallery Control API is not configured yet. Add the media API URL to unlock management.');
  }

  const headers = new Headers(options.headers || {});
  if (state.adminToken) headers.set('Authorization', `Bearer ${state.adminToken}`);

  const response = await fetch(buildUrl(path, params), {
    ...options,
    headers
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      clearAdminSession();
      throw new Error('Admin access was rejected. Sign in again and check the media API connection.');
    }
    throw new Error(data.message || 'Gallery Control request failed.');
  }
  return data;
}

async function loginWithApiPassword(apiBaseUrl, password) {
  const response = await fetch(`${apiBaseUrl}/api/auth/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ password })
  });

  const data = await response.json().catch(() => ({}));
  if (response.ok && data?.token) {
    return {
      token: String(data.token).trim(),
      expiresAt: data.expiresAt || null,
      mode: 'session'
    };
  }

  if (response.status === 404 || response.status === 405) {
    return {
      token: String(password || '').trim(),
      expiresAt: null,
      mode: 'legacy-token'
    };
  }

  throw new Error(data.message || 'Could not sign in to the media API.');
}

async function publicApiFetch(path, params) {
  if (!state.apiBaseUrl) return null;

  const response = await fetch(buildUrl(path, params));
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.message || 'Gallery Control public request failed.');
  }
  return data;
}

function setAlert(message, tone = '') {
  if (!studioAlert) return;
  if (!message) {
    studioAlert.hidden = true;
    studioAlert.textContent = '';
    studioAlert.className = 'studio-alert';
    return;
  }

  studioAlert.hidden = false;
  studioAlert.textContent = message;
  studioAlert.className = `studio-alert${tone ? ` is-${tone}` : ''}`;
}

function formatDate(value) {
  if (!value) return '-';
  try {
    return new Intl.DateTimeFormat('en-US', {
      dateStyle: 'medium',
      timeStyle: 'short'
    }).format(new Date(value));
  } catch {
    return value;
  }
}

function setBusy(button, busy, busyText, idleText) {
  if (!button) return;
  button.disabled = busy;
  button.textContent = busy ? busyText : idleText;
}

function setManagementAvailability(enabled) {
  state.canManage = Boolean(enabled);

  if (uploadButton) uploadButton.disabled = !enabled;
  if (uploadForm) {
    uploadForm.querySelectorAll('input, select, button').forEach((field) => {
      field.disabled = !enabled;
    });
  }
}

function renderCategories() {
  if (categorySelect) {
    categorySelect.innerHTML = '<option value="">Select category</option>';
  }

  if (categoryFilter) {
    categoryFilter.innerHTML = '<option value="">All categories</option>';
  }

  state.categories.forEach((category) => {
    if (categorySelect) {
      const option = document.createElement('option');
      option.value = category.id;
      option.textContent = category.name;
      categorySelect.appendChild(option);
    }

    if (categoryFilter) {
      const option = document.createElement('option');
      option.value = category.slug;
      option.textContent = category.name;
      categoryFilter.appendChild(option);
    }
  });
}

function renderMetrics() {
  const total = state.media.length;
  const published = state.media.filter((item) => item.published).length;
  const generated = state.media.filter((item) => item.generatedArtUrl).length;
  const social = state.media.filter((item) => item.usedInSocial).length;

  if (metricTotal) metricTotal.textContent = String(total);
  if (metricPublished) metricPublished.textContent = String(published);
  if (metricGenerated) metricGenerated.textContent = String(generated);
  if (metricSocial) metricSocial.textContent = String(social);
  if (countBadge) countBadge.textContent = `${total} item${total === 1 ? '' : 's'}`;
}

function renderPostQueueList(container, list, emptyText) {
  if (!container) return;
  if (!list.length) {
    container.innerHTML = `<p class="studio-post-empty">${emptyText}</p>`;
    return;
  }

  container.innerHTML = list
    .slice(0, 6)
    .map(
      (item) => `
        <article class="studio-post-item">
          <strong>${item.title}</strong>
          <span>${item.categoryName} · ${item.generatedArtUrl ? 'Artwork ready' : 'Waiting for artwork'}</span>
        </article>
      `
    )
    .join('');
}

function renderDailyPostsBoard() {
  const readyQueue = state.media.filter((item) => item.generatedArtUrl && !item.usedInSocial);
  const sentQueue = state.media.filter((item) => item.usedInSocial);
  const needsArt = state.media.filter((item) => !item.generatedArtUrl);

  if (studioPostsBadge) studioPostsBadge.textContent = `${readyQueue.length} ready`;
  if (studioReadyTodayCount) studioReadyTodayCount.textContent = String(readyQueue.length);
  if (studioSentTodayCount) studioSentTodayCount.textContent = String(sentQueue.length);
  if (studioNeedsArtCount) studioNeedsArtCount.textContent = String(needsArt.length);

  renderPostQueueList(studioReadyQueue, readyQueue, 'No ready items yet.');
  renderPostQueueList(studioSentQueue, sentQueue, 'No sent items yet.');
}

function buildStatus(text, active = false) {
  return `<span class="studio-status${active ? ' is-active' : ''}">${text}</span>`;
}

function ensureSelectedItem() {
  if (state.selectedId && state.media.some((item) => item.id === state.selectedId)) return;
  state.selectedId = state.media[0]?.id || '';
}

function renderPreview() {
  const item = state.media.find((entry) => entry.id === state.selectedId);
  if (!item) {
    if (previewTitle) previewTitle.textContent = 'No media selected';
    if (originalPreview) originalPreview.textContent = 'Select a media item to preview the original photo.';
    if (generatedPreview) generatedPreview.textContent = 'Generate the branded asset to preview it here.';
    if (metaCategory) metaCategory.textContent = '-';
    if (metaCreated) metaCreated.textContent = '-';
    if (metaSocial) metaSocial.textContent = '-';
    if (captionField) captionField.value = '';
    return;
  }

  if (previewTitle) previewTitle.textContent = item.title;
  if (originalPreview) {
    originalPreview.innerHTML = `<img src="${item.originalUrl}" alt="${item.title}" />`;
  }
  if (generatedPreview) {
    generatedPreview.innerHTML = item.generatedArtUrl
      ? `<img src="${item.generatedArtUrl}" alt="${item.title} generated art" />`
      : 'Generate the branded asset to preview it here.';
  }
  if (metaCategory) metaCategory.textContent = item.categoryName || '-';
  if (metaCreated) metaCreated.textContent = formatDate(item.createdAt);
  if (metaSocial) metaSocial.textContent = item.usedInSocial ? 'Used in social' : 'Not used yet';
  if (captionField) captionField.value = item.captionText || '';
}

function renderTable() {
  ensureSelectedItem();
  renderMetrics();
  renderDailyPostsBoard();
  renderPreview();

  if (!tableBody) return;
  tableBody.innerHTML = '';

  if (!state.media.length) {
    tableBody.innerHTML = '<tr><td colspan="6" class="studio-empty">No media items match the active filters.</td></tr>';
    return;
  }

  state.media.forEach((item) => {
    const row = document.createElement('tr');
    const actionMarkup = state.canManage
      ? `
        <div class="studio-row-actions">
          <button class="btn btn-secondary" type="button" data-action="generate" data-id="${item.id}">Generate art</button>
          <button class="btn btn-secondary" type="button" data-action="publish" data-id="${item.id}">
            ${item.published ? 'Unpublish' : 'Publish'}
          </button>
          <button class="btn btn-secondary" type="button" data-action="social" data-id="${item.id}">
            ${item.usedInSocial ? 'Mark unused' : 'Mark social'}
          </button>
        </div>
      `
      : '<span class="studio-status">Read-only view</span>';
    row.innerHTML = `
      <td><img class="studio-thumb" src="${item.originalUrl}" alt="${item.title}" /></td>
      <td>
        <span class="studio-row-title">${item.title}</span>
        <span class="studio-row-subtitle">${item.usedInSocial ? 'Already used in social' : 'Ready for next publication step'}</span>
      </td>
      <td>${item.categoryName}</td>
      <td>${buildStatus(item.generatedArtUrl ? 'Generated' : 'Pending', Boolean(item.generatedArtUrl))}</td>
      <td>${buildStatus(item.published ? 'Published' : 'Draft', item.published)}</td>
      <td>${actionMarkup}</td>
    `;

    row.addEventListener('click', (event) => {
      if (event.target.closest('button')) return;
      state.selectedId = item.id;
      renderPreview();
    });

    tableBody.appendChild(row);
  });
}

async function loadCategories() {
  if (state.mode === 'demo' && state.categories.length) {
    renderCategories();
    return;
  }

  if (state.mode === 'api') {
    const data = await apiFetch('/api/categories');
    state.categories = Array.isArray(data.categories) ? data.categories : [];
    renderCategories();
    return;
  }

  if (state.apiBaseUrl) {
    try {
      const data = await publicApiFetch('/api/public/categories');
      state.categories = Array.isArray(data?.categories) ? data.categories : [];
      renderCategories();
      return;
    } catch (error) {
      console.warn('Gallery Control public categories fallback enabled.', error);
    }
  }

  const categoryRows = await fetchSupabaseRows(getWebsiteContentConfig()?.categoriesTables || [], [
    ['select', '*'],
    ['order', 'name.asc']
  ]);
  state.categories = (Array.isArray(categoryRows) ? categoryRows : []).map(normalizeCategoryRow).filter(Boolean);
  if (!state.categories.length) state.categories = [...DEMO_CATEGORIES];
  renderCategories();
}

async function loadMedia() {
  if (state.mode === 'demo') {
    const demoMedia = readDemoMedia();
    state.media = demoMedia.filter((item) => {
      if (state.filters.search) {
        const haystack = `${item.title} ${item.categoryName}`.toLowerCase();
        if (!haystack.includes(state.filters.search.toLowerCase())) return false;
      }
      if (state.filters.category && item.categorySlug !== state.filters.category) return false;
      if (state.filters.published === 'true' && !item.published) return false;
      if (state.filters.published === 'false' && item.published) return false;
      return true;
    });
    state.connected = true;
    if (apiBadge) apiBadge.textContent = 'Browser test mode';
    renderTable();
    return;
  }

  if (state.mode === 'api') {
    const data = await apiFetch('/api/media', {}, state.filters);
    state.media = Array.isArray(data.media) ? data.media : [];
    state.connected = true;
    if (apiBadge) apiBadge.textContent = 'API connected';
    renderTable();
    return;
  }

  if (state.apiBaseUrl) {
    try {
      const data = await publicApiFetch('/api/public/media', {
        category: state.filters.category,
        search: state.filters.search,
        published: state.filters.published
      });
      state.media = Array.isArray(data?.media) ? data.media : [];
      state.connected = true;
      if (apiBadge) apiBadge.textContent = 'Public API connected';
      renderTable();
      return;
    } catch (error) {
      console.warn('Gallery Control public media fallback enabled.', error);
    }
  }

  const galleryRows = await fetchSupabaseRows(getWebsiteContentConfig()?.galleryItemsTables || [], [
    ['select', '*'],
    ['order', 'created_at.desc']
  ]);
  const { byId } = createCategoryMap(state.categories);
  const normalizedMedia = (Array.isArray(galleryRows) ? galleryRows : [])
    .map((row) => normalizeMediaItem(row, byId))
    .filter(Boolean);

  state.media = normalizedMedia.filter((item) => {
    if (state.filters.search) {
      const haystack = `${item.title} ${item.categoryName}`.toLowerCase();
      if (!haystack.includes(state.filters.search.toLowerCase())) return false;
    }
    if (state.filters.category && item.categorySlug !== state.filters.category) return false;
    if (state.filters.published === 'true' && !item.published) return false;
    if (state.filters.published === 'false' && item.published) return false;
    return true;
  });
  state.connected = true;
  if (apiBadge) apiBadge.textContent = 'Public gallery connected';
  renderTable();
}

async function refreshAll() {
  try {
    if (state.mode === 'api') {
      await Promise.all([loadCategories(), loadMedia()]);
    } else if (state.mode === 'demo') {
      await loadCategories();
      await loadMedia();
    } else if (state.apiBaseUrl) {
      await Promise.all([loadCategories(), loadMedia()]);
    } else {
      await loadCategories();
      await loadMedia();
    }
    if (state.mode === 'demo') {
      setAlert('Browser test mode is active. You can upload photos and test the tools safely on this device.', 'success');
    } else if (state.canManage) {
      setAlert('Gallery Control connected successfully. You can upload, generate, and publish from here.', 'success');
    } else {
      setAlert(
        'Gallery Control is running in public view mode. You can review published items and daily post status here, and secure upload/publish actions can be connected next.',
        'success'
      );
    }
  } catch (error) {
    state.connected = false;
    if (apiBadge) apiBadge.textContent = state.mode === 'api' ? 'API not connected' : 'Gallery not connected';
    state.media = [];
    renderTable();
    setAlert(error.message, 'error');
  }
}

async function handleUpload(event) {
  event.preventDefault();
  const formData = new FormData(uploadForm);
  if (!formData.get('image')) {
    setAlert('Please choose an image file before uploading.', 'error');
    return;
  }

  try {
    setBusy(uploadButton, true, 'Uploading...', 'Upload image');
    if (state.mode === 'demo') {
      await uploadDemoMedia(formData);
    } else {
      await apiFetch('/api/media/upload', {
        method: 'POST',
        body: formData
      });
    }
    uploadForm.reset();
    await loadMedia();
    setAlert(
      state.mode === 'demo'
        ? 'Image uploaded in browser test mode. You can now generate art or publish it inside this test workspace.'
        : 'Image uploaded successfully. You can now generate the branded social art or publish it.',
      'success'
    );
  } catch (error) {
    setAlert(error.message, 'error');
  } finally {
    setBusy(uploadButton, false, 'Uploading...', 'Upload image');
  }
}

async function handleTableAction(event) {
  const button = event.target.closest('button[data-action]');
  if (!button) return;
  if (!state.canManage) {
    setAlert('This page is currently in read-only mode. Connect the secure media API to manage uploads and status changes.', 'error');
    return;
  }

  const { action, id } = button.dataset;
  const item = state.media.find((entry) => entry.id === id);
  if (!item) return;

  try {
    button.disabled = true;
    if (action === 'generate') {
      if (state.mode === 'demo') {
        await patchDemoMedia(item, 'generate');
      } else {
        await apiFetch(`/api/media/${id}/generate-art`, { method: 'POST' });
      }
      setAlert(`Branded social art generated for "${item.title}".`, 'success');
    }
    if (action === 'publish') {
      if (state.mode === 'demo') {
        await patchDemoMedia(item, 'publish');
      } else {
        await apiFetch(`/api/media/${id}/status`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ published: !item.published })
        });
      }
      setAlert(item.published ? 'Media moved back to draft.' : 'Media published to the gallery pipeline.', 'success');
    }
    if (action === 'social') {
      if (state.mode === 'demo') {
        await patchDemoMedia(item, 'social');
      } else {
        await apiFetch(`/api/media/${id}/status`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ usedInSocial: !item.usedInSocial })
        });
      }
      setAlert(item.usedInSocial ? 'Media marked as not yet used in social.' : 'Media marked as used in social.', 'success');
    }

    await loadMedia();
  } catch (error) {
    setAlert(error.message, 'error');
  } finally {
    button.disabled = false;
  }
}

async function handleAccessSubmit(event) {
  event.preventDefault();

  const apiBaseUrl = normalizeApiBaseUrl(apiUrlInput?.value || state.apiBaseUrl);
  const adminPassword = String(adminTokenInput?.value || '').trim();

  if (!apiBaseUrl) {
    setAlert('Enter the secure media API URL before unlocking management.', 'error');
    return;
  }

  if (!adminPassword) {
    setAlert('Enter the admin password to unlock upload and publish controls.', 'error');
    return;
  }

  try {
    setBusy(unlockButton, true, 'Signing in...', 'Sign in');
    const session = await loginWithApiPassword(apiBaseUrl, adminPassword);
    setApiSession({ apiBaseUrl, adminToken: session.token });
    await refreshAll();
    setAlert(
      session.mode === 'legacy-token'
        ? 'Gallery management unlocked with legacy token access.'
        : 'Gallery management unlocked successfully.',
      'success'
    );
  } catch (error) {
    clearAdminSession();
    await refreshAll().catch(() => {});
    setAlert(error.message, 'error');
  } finally {
    setBusy(unlockButton, false, 'Signing in...', 'Sign in');
  }
}

function handleLockManagement() {
  if (state.mode === 'demo') {
    disableDemoMode();
  } else {
    clearAdminSession();
  }
  void refreshAll();
  setAlert('Gallery management was locked for this browser. Public tracking view remains available.', 'success');
}

function handleStartTestMode() {
  enableDemoMode();
  void refreshAll();
}

function handleSimulateUploadFlow() {
  if (!state.canManage && state.mode !== 'demo') {
    enableDemoMode();
  }
  void refreshAll().finally(() => {
    applyRequestedEntryExperience();
  });
}

function bindFilters() {
  searchInput?.addEventListener('input', () => {
    state.filters.search = searchInput.value.trim();
    void loadMedia().catch((error) => setAlert(error.message, 'error'));
  });

  categoryFilter?.addEventListener('change', () => {
    state.filters.category = categoryFilter.value;
    void loadMedia().catch((error) => setAlert(error.message, 'error'));
  });

  publishedFilter?.addEventListener('change', () => {
    state.filters.published = publishedFilter.value;
    void loadMedia().catch((error) => setAlert(error.message, 'error'));
  });

  refreshButton?.addEventListener('click', () => {
    void refreshAll();
  });

  accessForm?.addEventListener('submit', (event) => {
    void handleAccessSubmit(event);
  });

  lockButton?.addEventListener('click', handleLockManagement);
  testModeButton?.addEventListener('click', handleStartTestMode);
  simulateUploadButton?.addEventListener('click', handleSimulateUploadFlow);
}

function init() {
  setApiSession({
    apiBaseUrl: getApiBaseUrl(),
    adminToken: getStoredAdminToken()
  });

  if (isDemoModeStored() || (isUploadEntryRequested() && !state.canManage)) enableDemoMode();

  uploadForm?.addEventListener('submit', handleUpload);

  tableBody?.addEventListener('click', handleTableAction);
  bindFilters();
  void refreshAll().finally(() => {
    applyRequestedEntryExperience();
  });
}

init();
