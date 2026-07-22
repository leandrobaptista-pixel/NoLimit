const yearEl = document.getElementById('year');
if (yearEl) yearEl.textContent = new Date().getFullYear();

const header = document.querySelector('.site-header');
const navWrap = document.querySelector('.nav');
const navMenu = document.getElementById('primaryNav');
const navToggle = document.querySelector('.menu-toggle');
const form = document.getElementById('visitForm');
const note = document.getElementById('formNote');
const dateField = form?.querySelector('input[name="date"]');
const submitButton = form?.querySelector('button[type="submit"]');

const metaDescription = document.querySelector('meta[name="description"]');
const brandLogo = document.getElementById('brandLogo');
const brandName = document.getElementById('brandName');
const heroLogo = document.getElementById('heroLogo');
const aboutHeading = document.getElementById('aboutHeading');
const anniversaryLogo = document.getElementById('anniversaryLogo');
const galleryHint = document.getElementById('galleryHint');
const galleryControls = document.getElementById('galleryControls');
const galleryCount = document.getElementById('galleryCount');
const galleryMoreButton = document.getElementById('galleryMoreBtn');
const featuredMetrics = document.getElementById('featuredMetrics');
const featuredLead = document.getElementById('featuredLead');
const featuredRail = document.getElementById('featuredRail');
const categorySpotlights = document.getElementById('categorySpotlights');
const catalogsPreview = document.getElementById('catalogsPreview');
const galleryLightbox = document.getElementById('galleryLightbox');
const galleryLightboxImage = document.getElementById('galleryLightboxImage');
const galleryLightboxTitle = document.getElementById('galleryLightboxTitle');
const galleryLightboxEyebrow = document.getElementById('galleryLightboxEyebrow');
const galleryLightboxMeta = document.getElementById('galleryLightboxMeta');
const galleryLightboxPrev = document.getElementById('galleryLightboxPrev');
const galleryLightboxNext = document.getElementById('galleryLightboxNext');
const galleryLightboxClose = document.getElementById('galleryLightboxClose');
const galleryLightboxBack = document.getElementById('galleryLightboxBack');
const footerCompany = document.getElementById('footerCompany');
const footerPhoneLink = document.getElementById('footerPhoneLink');
const footerEmailLink = document.getElementById('footerEmailLink');
const WEBSITE_CONTENT_CACHE_KEY = 'websiteContentCache:v1';
const GALLERY_BATCH_SIZE = 12;
const galleryUiState = {
  data: null,
  categorySlug: '',
  visibleBySlug: {}
};
const galleryLightboxState = {
  items: [],
  index: 0,
  opener: null
};

const DEFAULT_SITE_PROFILE = {
  companyName: 'No Limit Carpentry',
  phone: '',
  email: form?.dataset?.to || 'info@your-company.com',
  logoUrl: brandLogo?.getAttribute('src') || 'assets/brand.png',
  anniversaryLogoUrl: anniversaryLogo?.getAttribute('src') || 'assets/anniversary-18.png',
  defaultCta: 'Request a Visit'
};

const DEFAULT_WEBSITE_CONTENT_TABLES = {
  settingsTables: ['website_settings', 'site_settings', 'company_profile'],
  categoriesTables: ['website_categories', 'categories'],
  galleryItemsTables: ['website_gallery_items', 'gallery_items', 'portfolio_items']
};
const IMAGE_VARIANTS = window.OPTIMIZED_IMAGE_VARIANTS || {};

function syncHeaderHeight() {
  if (!header) return;
  const height = Math.ceil(header.getBoundingClientRect().height);
  document.documentElement.style.setProperty('--header-h', `${height}px`);
}

function closeMobileNav() {
  navToggle?.setAttribute('aria-expanded', 'false');
  navWrap?.classList.remove('nav-open');
  syncHeaderHeight();
}

function scrollToSection(hash, { updateHistory = true } = {}) {
  const id = String(hash || '').replace(/^#/, '').trim();
  if (!id) return false;
  const target = document.getElementById(id);
  if (!target) return false;

  target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  if (updateHistory) history.replaceState(null, '', `#${id}`);
  return true;
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

function normalizeTableList(value, fallback) {
  const list = Array.isArray(value) ? value : value ? [value] : fallback;
  return list.map((item) => String(item || '').trim()).filter(Boolean);
}

function isAbsoluteUrl(value) {
  return /^(?:[a-z]+:)?\/\//i.test(value) || String(value || '').startsWith('data:') || String(value || '').startsWith('blob:');
}

function toURL(path) {
  const value = String(path || '').trim();
  if (!value) return '';
  if (isAbsoluteUrl(value)) return value;
  return value.split('/').map(encodeURIComponent).join('/');
}

function humanizePhotoTitle(source) {
  const fileName = decodeURIComponent(String(source || '').split('/').pop() || 'photo');
  const baseName = fileName.replace(/\.[^.]+$/, '').replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
  return baseName || 'Project photo';
}

function normalizeAssetPath(value) {
  return decodeURIComponent(String(value || '').trim())
    .replace(/^\.\/+/, '')
    .replace(/^\/+/, '');
}

function resolveImageVariants(imageUrl) {
  const normalizedPath = normalizeAssetPath(imageUrl);
  const localVariants = IMAGE_VARIANTS[normalizedPath];
  return {
    thumbUrl: localVariants?.thumb || '',
    featureUrl: localVariants?.feature || localVariants?.thumb || ''
  };
}

function getCatalogLibrary() {
  return Array.isArray(window.CATALOG_LIBRARY) ? window.CATALOG_LIBRARY : [];
}

function formatPhoneHref(phone) {
  const digits = String(phone || '').replace(/[^\d+]/g, '');
  return digits ? `tel:${digits}` : '';
}

function scheduleNonCriticalTask(callback) {
  if (typeof callback !== 'function') return;
  if (typeof window.requestIdleCallback === 'function') {
    window.requestIdleCallback(() => callback(), { timeout: 1500 });
    return;
  }
  window.setTimeout(() => callback(), 120);
}

function readStoredJson(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function readCachedSiteSettings() {
  const cached = readStoredJson('websiteSiteSettings');
  return cached && typeof cached === 'object' ? cached : null;
}

function readWebsiteContentCache() {
  const cached = readStoredJson(WEBSITE_CONTENT_CACHE_KEY);
  return cached && typeof cached === 'object' ? cached : null;
}

function writeWebsiteContentCache(content) {
  try {
    localStorage.setItem(
      WEBSITE_CONTENT_CACHE_KEY,
      JSON.stringify({
        savedAt: new Date().toISOString(),
        settings: content?.settings || null,
        gallery: content?.gallery || null
      })
    );
  } catch {
    // Ignore storage errors.
  }
}

function hasRenderableGallery(gallery) {
  return Boolean(gallery?.categories?.length && Object.keys(gallery.itemsBySlug || {}).length);
}

function setGalleryHint(message) {
  if (!galleryHint) return;
  galleryHint.textContent = message;
}

function getContactFormConfig() {
  const config = window.CONTACT_FORM || {};
  const supabaseUrl = String(config.supabaseUrl || '').trim().replace(/\/$/, '');
  const supabaseAnonKey = String(config.supabaseAnonKey || '').trim();
  const table = String(config.table || 'public_visit_requests').trim();

  if (!supabaseUrl || !supabaseAnonKey || !table) return null;

  return {
    supabaseUrl,
    supabaseAnonKey,
    table
  };
}

function getWebsiteContentConfig() {
  const config = window.WEBSITE_CONTENT || {};
  const sharedConfig = window.CONTACT_FORM || {};
  const supabaseUrl = String(config.supabaseUrl || sharedConfig.supabaseUrl || '').trim().replace(/\/$/, '');
  const supabaseAnonKey = String(config.supabaseAnonKey || sharedConfig.supabaseAnonKey || '').trim();

  if (!supabaseUrl || !supabaseAnonKey) return null;

  return {
    supabaseUrl,
    supabaseAnonKey,
    settingsTables: normalizeTableList(config.settingsTables || config.settingsTable, DEFAULT_WEBSITE_CONTENT_TABLES.settingsTables),
    categoriesTables: normalizeTableList(config.categoriesTables || config.categoriesTable, DEFAULT_WEBSITE_CONTENT_TABLES.categoriesTables),
    galleryItemsTables: normalizeTableList(config.galleryItemsTables || config.galleryItemsTable, DEFAULT_WEBSITE_CONTENT_TABLES.galleryItemsTables)
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
      console.warn(`Website content fetch failed for table "${table}".`, error);
    }
  }

  return null;
}

function normalizeSiteSettingsRow(row) {
  if (!row || typeof row !== 'object') return null;

  return {
    companyName: String(row.company_name || row.companyName || DEFAULT_SITE_PROFILE.companyName).trim() || DEFAULT_SITE_PROFILE.companyName,
    phone: String(row.phone || '').trim(),
    email: String(row.email || DEFAULT_SITE_PROFILE.email).trim() || DEFAULT_SITE_PROFILE.email,
    logoUrl: String(row.logo_url || row.logoUrl || DEFAULT_SITE_PROFILE.logoUrl).trim() || DEFAULT_SITE_PROFILE.logoUrl,
    anniversaryLogoUrl: String(row.logo_18_years_url || row.logo18YearsUrl || DEFAULT_SITE_PROFILE.anniversaryLogoUrl).trim() || DEFAULT_SITE_PROFILE.anniversaryLogoUrl,
    defaultCta: String(row.default_cta || row.defaultCta || DEFAULT_SITE_PROFILE.defaultCta).trim() || DEFAULT_SITE_PROFILE.defaultCta
  };
}

function normalizeCategoryRow(row) {
  if (!row || typeof row !== 'object') return null;

  const name = String(row.name || row.title || '').trim();
  const slug = String(row.slug || slugify(name || row.id)).trim();
  if (!name || !slug) return null;

  return {
    id: row.id ?? slug,
    name,
    slug
  };
}

function isPublishedValue(value) {
  if (value === undefined || value === null || value === '') return true;
  if (typeof value === 'boolean') return value;
  return String(value).trim().toLowerCase() === 'true';
}

function createCategoryMap(categories) {
  const byId = new Map();
  const bySlug = new Map();

  categories.forEach((category) => {
    if (category?.id !== undefined && category?.id !== null) byId.set(String(category.id), category);
    if (category?.slug) bySlug.set(category.slug, category);
  });

  return { byId, bySlug };
}

function buildStaticGalleryData(manifest) {
  const categories = [];
  const itemsBySlug = {};

  Object.entries(manifest || {}).forEach(([name, list]) => {
    const slug = slugify(name) || name;
    const uniqueList = dedupeByBase(Array.isArray(list) ? list : []);
    if (!uniqueList.length) return;

    categories.push({ id: slug, name, slug });
    itemsBySlug[slug] = uniqueList.map((imageUrl, index) => ({
      ...resolveImageVariants(imageUrl),
      id: `${slug}-${index}`,
      title: humanizePhotoTitle(imageUrl),
      imageUrl,
      createdAt: '',
      usedInSocial: false
    }));
  });

  return {
    source: 'static',
    categories,
    itemsBySlug
  };
}

function buildDynamicGalleryData(categoryRows, galleryRows) {
  const declaredCategories = (Array.isArray(categoryRows) ? categoryRows : []).map(normalizeCategoryRow).filter(Boolean);
  const { byId, bySlug } = createCategoryMap(declaredCategories);
  const categoryOrder = declaredCategories.map((category) => category.slug);
  const itemsBySlug = {};

  function ensureCategory(nextCategory) {
    if (!nextCategory?.slug || bySlug.has(nextCategory.slug)) return;
    bySlug.set(nextCategory.slug, nextCategory);
    if (nextCategory.id !== undefined && nextCategory.id !== null) byId.set(String(nextCategory.id), nextCategory);
    categoryOrder.push(nextCategory.slug);
  }

  (Array.isArray(galleryRows) ? galleryRows : []).forEach((row, index) => {
    if (!row || typeof row !== 'object' || !isPublishedValue(row.published)) return;

    const imageUrl = String(row.image_url || row.imageUrl || '').trim();
    if (!imageUrl) return;

    const categoryId = row.category_id ?? row.categoryId ?? '';
    const linkedCategory = categoryId !== '' ? byId.get(String(categoryId)) : null;
    const rawCategoryName =
      linkedCategory?.name ||
      (typeof row.category === 'object' ? row.category?.name : row.category) ||
      row.category_name ||
      row.categoryName ||
      '';
    const rawCategorySlug =
      linkedCategory?.slug ||
      (typeof row.category === 'object' ? row.category?.slug : '') ||
      row.category_slug ||
      row.categorySlug ||
      slugify(rawCategoryName);

    const fallbackCategoryName = rawCategoryName ? String(rawCategoryName).trim() : 'Uncategorized';
    const fallbackCategorySlug = String(rawCategorySlug || slugify(fallbackCategoryName) || 'uncategorized').trim();

    ensureCategory({
      id: linkedCategory?.id ?? categoryId ?? fallbackCategorySlug,
      name: linkedCategory?.name || fallbackCategoryName,
      slug: linkedCategory?.slug || fallbackCategorySlug
    });

    const targetCategory = bySlug.get(fallbackCategorySlug);
    if (!targetCategory) return;

    if (!itemsBySlug[targetCategory.slug]) itemsBySlug[targetCategory.slug] = [];
    const variants = resolveImageVariants(imageUrl);
    itemsBySlug[targetCategory.slug].push({
      ...variants,
      id: row.id ?? `${targetCategory.slug}-${index}`,
      title: String(row.title || humanizePhotoTitle(imageUrl)).trim() || humanizePhotoTitle(imageUrl),
      imageUrl,
      createdAt: String(row.created_at || row.createdAt || '').trim(),
      usedInSocial: Boolean(row.used_in_social || row.usedInSocial)
    });
  });

  Object.values(itemsBySlug).forEach((list) => {
    list.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || '') || String(a.title || '').localeCompare(String(b.title || '')));
  });

  const categories = categoryOrder
    .map((slug) => bySlug.get(slug))
    .filter((category) => category && Array.isArray(itemsBySlug[category.slug]) && itemsBySlug[category.slug].length);

  return {
    source: 'supabase',
    categories,
    itemsBySlug
  };
}

function applySiteSettings(settings) {
  const profile = { ...DEFAULT_SITE_PROFILE, ...(settings || {}) };

  if (brandLogo && profile.logoUrl) {
    brandLogo.src = profile.logoUrl;
    brandLogo.alt = `${profile.companyName} logo`;
  }

  if (heroLogo && profile.logoUrl) {
    heroLogo.src = profile.logoUrl;
    heroLogo.alt = profile.companyName;
  }

  if (anniversaryLogo && profile.anniversaryLogoUrl) {
    anniversaryLogo.src = profile.anniversaryLogoUrl;
    anniversaryLogo.alt = `${profile.companyName} anniversary badge`;
  }

  if (brandName) brandName.textContent = profile.companyName;
  if (aboutHeading) aboutHeading.textContent = `About ${profile.companyName}`;
  if (footerCompany) footerCompany.textContent = profile.companyName;
  if (metaDescription) {
    metaDescription.setAttribute(
      'content',
      `${profile.companyName} delivers premium finish carpentry: custom built-ins, crown molding, wall paneling, doors, and more. Book a free on-site visit.`
    );
  }
  document.title = `${profile.companyName} | Premium Finish Carpentry`;

  if (form && profile.email) form.dataset.to = profile.email;

  document.querySelectorAll('[data-default-cta]').forEach((element) => {
    element.textContent = profile.defaultCta;
  });

  if (footerPhoneLink) {
    if (profile.phone) {
      footerPhoneLink.hidden = false;
      footerPhoneLink.href = formatPhoneHref(profile.phone);
      footerPhoneLink.textContent = profile.phone;
    } else {
      footerPhoneLink.hidden = true;
      footerPhoneLink.removeAttribute('href');
      footerPhoneLink.textContent = '';
    }
  }

  if (footerEmailLink) {
    if (profile.email) {
      footerEmailLink.hidden = false;
      footerEmailLink.href = `mailto:${profile.email}`;
      footerEmailLink.textContent = profile.email;
    } else {
      footerEmailLink.hidden = true;
      footerEmailLink.removeAttribute('href');
      footerEmailLink.textContent = '';
    }
  }

  try {
    localStorage.setItem('websiteSiteSettings', JSON.stringify(profile));
  } catch {
    // Ignore storage errors.
  }
}

async function loadWebsiteContent() {
  const config = getWebsiteContentConfig();
  if (!config) return null;

  try {
    const [settingsRows, categoryRows, galleryRows] = await Promise.all([
      fetchSupabaseRows(config.settingsTables, [
        ['select', 'company_name,phone,email,logo_url,logo_18_years_url,default_cta'],
        ['limit', '1']
      ]),
      fetchSupabaseRows(config.categoriesTables, [
        ['select', 'id,name,slug'],
        ['order', 'name.asc']
      ]),
      fetchSupabaseRows(config.galleryItemsTables, [
        ['select', 'id,title,category,image_url,created_at,published,used_in_social'],
        ['order', 'created_at.desc']
      ])
    ]);

    const settings = normalizeSiteSettingsRow(settingsRows?.[0]);
    const gallery = buildDynamicGalleryData(categoryRows || [], galleryRows || []);

    if (!settings && !gallery.categories.length) return null;

    return {
      settings,
      gallery
    };
  } catch (error) {
    console.warn('Website content fallback enabled.', error);
    return null;
  }
}

function encodeMailto(fields) {
  const to = form?.dataset?.to || DEFAULT_SITE_PROFILE.email;
  const subject = encodeURIComponent('New Visit Request - A No Limit');
  const lines = [
    `Name: ${fields.name || ''}`,
    `Email: ${fields.email || ''}`,
    `Phone: ${fields.phone || ''}`,
    `Address: ${fields.address || ''}`,
    `City: ${fields.city || ''}`,
    `Preferred Date: ${fields.date || ''}`,
    `Project Type: ${fields.type || ''}`,
    '',
    'Details:',
    fields.details || ''
  ];
  const body = encodeURIComponent(lines.join('\n'));
  return `mailto:${to}?subject=${subject}&body=${body}`;
}

async function submitVisitRequest(fields) {
  const config = getContactFormConfig();
  if (!config) return { mode: 'mailto' };

  const response = await fetch(`${config.supabaseUrl}/rest/v1/${config.table}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      apikey: config.supabaseAnonKey,
      Authorization: `Bearer ${config.supabaseAnonKey}`,
      Prefer: 'return=representation'
    },
    body: JSON.stringify({
      source: 'website',
      page_url: window.location.href,
      name: fields.name,
      email: fields.email,
      phone: fields.phone || null,
      address: fields.address || null,
      city: fields.city || null,
      preferred_date: fields.date || null,
      project_type: fields.type || null,
      details: fields.details || null
    })
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => '');
    throw new Error(errorText || 'Could not submit request.');
  }

  return { mode: 'supabase' };
}

function getFields() {
  const data = {};
  if (!form) return data;
  new FormData(form).forEach((value, key) => {
    data[key] = typeof value === 'string' ? value.trim() : value;
  });
  return data;
}

function setFormNote(message, state = '') {
  if (!note) return;
  note.textContent = message;
  if (state) {
    note.dataset.state = state;
  } else {
    delete note.dataset.state;
  }
}

function getFieldWrapper(name) {
  return form?.querySelector(`.field[data-field="${name}"]`) || null;
}

function setFieldState(name, value, error = '', forceReveal = false) {
  const wrapper = getFieldWrapper(name);
  if (!wrapper) return;

  const isCheckbox = wrapper.classList.contains('consent-field');
  const hasValue = isCheckbox ? Boolean(value) : Boolean(String(value || '').trim());
  const shouldRevealError = forceReveal || wrapper.dataset.touched === 'true';
  wrapper.classList.toggle('is-invalid', shouldRevealError && Boolean(error));
  wrapper.classList.toggle('is-valid', !error && hasValue);

  const errorEl = wrapper.querySelector('.field-error');
  if (errorEl) errorEl.textContent = shouldRevealError ? error : '';
}

function validate(fields) {
  const errors = {};

  if (!fields.name) {
    errors.name = 'Please enter your full name.';
  }

  if (!fields.email) {
    errors.email = 'Please enter your email address.';
  } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(fields.email)) {
    errors.email = 'Please enter a valid email address.';
  }

  if (form && !form.querySelector('input[name="agree"]')?.checked) {
    errors.agree = 'Please confirm that we may contact you about this request.';
  }

  return errors;
}

function syncFieldStates(fields, errors, forceReveal = false) {
  ['name', 'email', 'phone', 'city', 'address', 'date', 'type', 'details', 'agree'].forEach((name) => {
    const value = name === 'agree' ? form?.querySelector('input[name="agree"]')?.checked : fields[name];
    setFieldState(name, value, errors[name] || '', forceReveal);
  });
}

function resetVisitForm() {
  form?.reset();
  ['name', 'email', 'phone', 'city', 'address', 'date', 'type', 'details', 'agree'].forEach((name) => {
    const wrapper = getFieldWrapper(name);
    if (!wrapper) return;
    delete wrapper.dataset.touched;
    wrapper.classList.remove('is-valid', 'is-invalid');
    const errorEl = wrapper.querySelector('.field-error');
    if (errorEl) errorEl.textContent = '';
  });
}

const FALLBACK_GALLERY = {
  Trim: ['assets/00-Trim/IMG_7031.jpg'],
  Wainscoting: ['assets/00-Wainscoting/IMG_1663.JPG'],
  Stairs: ['assets/00-Stairs/IMG_9268.JPG'],
  Ceiling: ['assets/00-Ceiling/IMG_6650.jpg'],
  Decks: ['assets/00-Decks/IMG_4694.jpg'],
  'Kitchen & Vanities': ['assets/00-kitchen & Vanities/IMG_5865.JPG'],
  'Fireplaces & Bars': ['assets/00-Fireplaces & Bars/IMG_1816.JPG'],
  'Outside Doors & Windows': ['assets/00-Outside Doors & Windows/IMG_3311.JPG'],
  Pergola: ['assets/00-Pergola/IMG_2744.jpg'],
  'Port & Portal': ['assets/00-Port & Portal/IMG_5811.JPG']
};

function dedupeByBase(list) {
  const preferOrder = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'JPG', 'JPEG', 'PNG', 'WEBP', 'GIF'];
  const byBase = new Map();

  list.forEach((item) => {
    const file = String(item || '').split('/').pop() || '';
    const ext = file.includes('.') ? file.split('.').pop() : '';
    const base = file.replace(/\.[^.]+$/, '');
    const current = byBase.get(base);

    if (!current) {
      byBase.set(base, { path: item, ext });
      return;
    }

    if (preferOrder.indexOf(ext) < preferOrder.indexOf(current.ext)) {
      byBase.set(base, { path: item, ext });
    }
  });

  return Array.from(byBase.values()).map((value) => value.path);
}

async function loadManifest() {
  if (window.GALLERY_MANIFEST && typeof window.GALLERY_MANIFEST === 'object') {
    return window.GALLERY_MANIFEST;
  }

  try {
    const resp = await fetch(`assets/gallery-manifest.json?t=${Date.now()}`, { cache: 'no-store' });
    if (!resp.ok) throw new Error('manifest not available');
    const data = await resp.json();
    if (!data || typeof data !== 'object') throw new Error('invalid manifest');
    return data;
  } catch {
    return FALLBACK_GALLERY;
  }
}

function getInitialCategory(categories) {
  const urlCat = new URLSearchParams(location.search).get('cat');
  if (urlCat) {
    const normalized = slugify(urlCat);
    const matched = categories.find((category) => category.slug === normalized || slugify(category.name) === normalized);
    if (matched) return matched.slug;
  }

  const firstActive = document.querySelector('.filter-btn.active')?.dataset?.cat;
  if (firstActive) {
    const normalized = slugify(firstActive);
    const matched = categories.find((category) => category.slug === normalized || slugify(category.name) === normalized);
    if (matched) return matched.slug;
  }

  return categories[0]?.slug || 'trim';
}

function renderGalleryControls(categories, activeSlug, onChange) {
  if (!galleryControls) return;
  galleryControls.innerHTML = '';

  categories.forEach((category) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `filter-btn${category.slug === activeSlug ? ' active' : ''}`;
    button.dataset.cat = category.slug;
    button.textContent = category.name;
    button.addEventListener('click', () => onChange(category.slug));
    galleryControls.appendChild(button);
  });
}

function buildGalleryItem(item, category, index) {
  const fig = document.createElement('figure');
  const link = document.createElement('a');
  link.href = buildViewerHref(item.imageUrl, category.slug, index);

  const img = document.createElement('img');
  img.loading = 'lazy';
  img.decoding = 'async';
  img.fetchPriority = 'low';
  img.sizes = '(max-width: 600px) 100vw, (max-width: 900px) 50vw, 33vw';
  img.src = toURL(item.thumbUrl || item.imageUrl);
  img.alt = item.title || `${category.name} project`;
  img.onerror = () => {
    img.src = 'assets/placeholder.svg';
  };

  const cap = document.createElement('figcaption');
  cap.textContent = item.title || humanizePhotoTitle(item.imageUrl);

  link.appendChild(img);
  bindLightboxTrigger(link, category.items || [], index);
  fig.appendChild(link);
  fig.appendChild(cap);
  return fig;
}

function buildViewerHref(imageUrl, categorySlug, index = 0) {
  return `viewer.html?img=${encodeURIComponent(imageUrl)}&cat=${encodeURIComponent(categorySlug)}&i=${index}`;
}

function buildCategoryHref(categorySlug) {
  return `?cat=${encodeURIComponent(categorySlug)}#gallery`;
}

function canUseInlineLightbox(event) {
  return Boolean(
    event &&
      event.button === 0 &&
      !event.metaKey &&
      !event.ctrlKey &&
      !event.shiftKey &&
      !event.altKey
  );
}

function closeGalleryLightbox({ returnFocus = true } = {}) {
  if (!galleryLightbox || galleryLightbox.hidden) return;
  galleryLightbox.hidden = true;
  galleryLightbox.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('lightbox-open');

  if (returnFocus && galleryLightboxState.opener instanceof HTMLElement) {
    galleryLightboxState.opener.focus({ preventScroll: true });
  }
}

function renderGalleryLightbox() {
  if (!galleryLightbox || !galleryLightboxImage || !galleryLightboxTitle || !galleryLightboxMeta) return;
  const item = galleryLightboxState.items[galleryLightboxState.index];
  if (!item) return;

  const positionText = galleryLightboxState.items.length > 1 ? ` · ${galleryLightboxState.index + 1} of ${galleryLightboxState.items.length}` : '';
  const categoryText = item.categoryName || 'Gallery photo';

  galleryLightboxImage.src = toURL(item.featureUrl || item.imageUrl || item.thumbUrl);
  galleryLightboxImage.alt = item.title || humanizePhotoTitle(item.imageUrl);
  galleryLightboxTitle.textContent = item.title || humanizePhotoTitle(item.imageUrl);
  galleryLightboxEyebrow.textContent = categoryText;
  galleryLightboxMeta.textContent = `${formatPublishedDate(item.createdAt)}${positionText} · Tap the image, click outside, or use X to return.`;

  if (galleryLightboxPrev) {
    const hasMultiple = galleryLightboxState.items.length > 1;
    galleryLightboxPrev.disabled = !hasMultiple;
    galleryLightboxPrev.hidden = !hasMultiple;
  }

  if (galleryLightboxNext) {
    const hasMultiple = galleryLightboxState.items.length > 1;
    galleryLightboxNext.disabled = !hasMultiple;
    galleryLightboxNext.hidden = !hasMultiple;
  }
}

function stepGalleryLightbox(direction) {
  if (galleryLightboxState.items.length <= 1) return;
  const total = galleryLightboxState.items.length;
  galleryLightboxState.index = (galleryLightboxState.index + direction + total) % total;
  renderGalleryLightbox();
}

function openGalleryLightbox(items, index = 0, opener = null) {
  if (!galleryLightbox) return;
  const normalizedItems = (Array.isArray(items) ? items : []).filter((item) => item?.imageUrl);
  if (!normalizedItems.length) return;

  galleryLightboxState.items = normalizedItems;
  galleryLightboxState.index = Math.max(0, Math.min(index, normalizedItems.length - 1));
  galleryLightboxState.opener = opener instanceof HTMLElement ? opener : null;

  renderGalleryLightbox();
  galleryLightbox.hidden = false;
  galleryLightbox.setAttribute('aria-hidden', 'false');
  document.body.classList.add('lightbox-open');
  galleryLightboxClose?.focus({ preventScroll: true });
}

function bindLightboxTrigger(trigger, items, index) {
  if (!trigger || !galleryLightbox) return;

  trigger.addEventListener('click', (event) => {
    if (!canUseInlineLightbox(event)) return;
    event.preventDefault();
    openGalleryLightbox(items, index, trigger);
  });
}

function flattenGalleryItems(galleryData) {
  const items = [];
  let sequence = 0;

  galleryData.categories.forEach((category) => {
    const list = Array.isArray(galleryData.itemsBySlug[category.slug]) ? galleryData.itemsBySlug[category.slug] : [];
    list.forEach((item, categoryIndex) => {
      items.push({
        ...item,
        categoryName: category.name,
        categorySlug: category.slug,
        categoryIndex,
        sequence: sequence++
      });
    });
  });

  return items.sort(
    (a, b) => (b.createdAt || '').localeCompare(a.createdAt || '') || a.sequence - b.sequence
  );
}

function formatPublishedDate(value) {
  if (!value) return 'Published gallery item';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return 'Published gallery item';
  return date.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });
}

function createFeaturedImage(imageUrl, alt, sizes) {
  const image = document.createElement('img');
  image.loading = 'lazy';
  image.decoding = 'async';
  image.fetchPriority = 'low';
  image.sizes = sizes;
  image.src = toURL(imageUrl);
  image.alt = alt;
  image.onerror = () => {
    image.src = 'assets/placeholder.svg';
  };
  return image;
}

function renderFeaturedWork(galleryData) {
  if (!featuredMetrics || !featuredLead || !featuredRail) return;

  const items = flattenGalleryItems(galleryData);
  if (!items.length) {
    featuredMetrics.innerHTML = '';
    featuredLead.innerHTML = `
      <div class="featured-lead-copy">
        <p class="featured-eyebrow">Selected project</p>
        <h3>No published highlights yet</h3>
        <p class="hint">As soon as new gallery items are published, this area will update automatically.</p>
      </div>
    `;
    featuredRail.innerHTML = '';
    return;
  }

  const lead = items[0];
  const railItems = items.slice(1, 5);
  const totalPhotos = items.length;
  const totalSegments = galleryData.categories.length;

  featuredMetrics.innerHTML = '';
  [
    { label: 'Published photos', value: String(totalPhotos) },
    { label: 'Active segments', value: String(totalSegments) },
    { label: 'Current focus', value: lead.categoryName }
  ].forEach((entry) => {
    const metric = document.createElement('article');
    metric.className = 'featured-metric card';

    const label = document.createElement('span');
    label.textContent = entry.label;
    const value = document.createElement('strong');
    value.textContent = entry.value;

    metric.append(label, value);
    featuredMetrics.appendChild(metric);
  });

  featuredLead.innerHTML = '';
  const leadLink = document.createElement('a');
  leadLink.className = 'featured-lead-link';
  leadLink.href = buildViewerHref(lead.imageUrl, lead.categorySlug, lead.categoryIndex);
  bindLightboxTrigger(leadLink, items, 0);

  const leadMedia = document.createElement('div');
  leadMedia.className = 'featured-lead-media';
  const leadChip = document.createElement('span');
  leadChip.className = 'featured-chip';
  leadChip.textContent = lead.categoryName;
  leadMedia.appendChild(leadChip);
  leadMedia.appendChild(
    createFeaturedImage(
      lead.featureUrl || lead.thumbUrl || lead.imageUrl,
      lead.title || `${lead.categoryName} featured project`,
      '(max-width: 900px) 100vw, 40vw'
    )
  );

  const leadCopy = document.createElement('div');
  leadCopy.className = 'featured-lead-copy';
  const leadEyebrow = document.createElement('p');
  leadEyebrow.className = 'featured-eyebrow';
  leadEyebrow.textContent = 'Most recent published highlight';
  const leadTitle = document.createElement('h3');
  leadTitle.textContent = lead.title || humanizePhotoTitle(lead.imageUrl);
  const leadMeta = document.createElement('p');
  leadMeta.className = 'hint';
  leadMeta.textContent = `${formatPublishedDate(lead.createdAt)} · ${lead.categoryName}`;

  leadCopy.append(leadEyebrow, leadTitle, leadMeta);
  leadLink.append(leadMedia, leadCopy);
  featuredLead.appendChild(leadLink);

  featuredRail.innerHTML = '';
  railItems.forEach((item, index) => {
    const card = document.createElement('article');
    card.className = 'featured-card card';

    const link = document.createElement('a');
    link.className = 'featured-card-link';
    link.href = buildViewerHref(item.imageUrl, item.categorySlug, item.categoryIndex);
    bindLightboxTrigger(link, items, index + 1);

    const media = document.createElement('div');
    media.className = 'featured-card-media';
    media.appendChild(
      createFeaturedImage(
        item.thumbUrl || item.imageUrl,
        item.title || `${item.categoryName} project`,
        '(max-width: 600px) 100vw, (max-width: 900px) 50vw, 18vw'
      )
    );

    const copy = document.createElement('div');
    copy.className = 'featured-card-copy';

    const meta = document.createElement('div');
    meta.className = 'featured-card-meta';
    const category = document.createElement('span');
    category.className = 'featured-card-category';
    category.textContent = item.categoryName;
    const position = document.createElement('span');
    position.className = 'featured-card-index';
    position.textContent = `Highlight ${index + 1}`;
    meta.append(category, position);

    const title = document.createElement('h3');
    title.textContent = item.title || humanizePhotoTitle(item.imageUrl);

    const detail = document.createElement('p');
    detail.textContent = formatPublishedDate(item.createdAt);

    copy.append(meta, title, detail);
    link.append(media, copy);
    card.appendChild(link);
    featuredRail.appendChild(card);
  });
}

function renderCategorySpotlights(galleryData) {
  if (!categorySpotlights) return;

  const cards = galleryData.categories
    .map((category) => {
      const list = Array.isArray(galleryData.itemsBySlug[category.slug]) ? galleryData.itemsBySlug[category.slug] : [];
      return {
        category,
        count: list.length,
        cover: list[0] || null
      };
    })
    .filter((entry) => entry.cover)
    .sort((a, b) => b.count - a.count || a.category.name.localeCompare(b.category.name))
    .slice(0, 6);

  if (!cards.length) {
    categorySpotlights.innerHTML = '<p class="hint">Category previews will appear here as soon as published gallery items are available.</p>';
    return;
  }

  categorySpotlights.innerHTML = '';
  cards.forEach(({ category, count, cover }) => {
    const article = document.createElement('article');
    article.className = 'spotlight-card card';

    const link = document.createElement('a');
    link.className = 'spotlight-card-link';
    link.href = buildCategoryHref(category.slug);

    const media = document.createElement('div');
    media.className = 'spotlight-card-media';
    media.appendChild(
      createFeaturedImage(
        cover.featureUrl || cover.thumbUrl || cover.imageUrl,
        cover.title || `${category.name} project`,
        '(max-width: 600px) 100vw, (max-width: 900px) 50vw, 28vw'
      )
    );

    const copy = document.createElement('div');
    copy.className = 'spotlight-card-copy';

    const meta = document.createElement('div');
    meta.className = 'spotlight-card-meta';
    const badge = document.createElement('span');
    badge.className = 'spotlight-card-badge';
    badge.textContent = category.name;
    const total = document.createElement('span');
    total.className = 'spotlight-card-count';
    total.textContent = `${count} photos`;
    meta.append(badge, total);

    const title = document.createElement('h3');
    title.textContent = cover.title || humanizePhotoTitle(cover.imageUrl);

    const description = document.createElement('p');
    description.textContent = `Open ${category.name} and jump directly into a curated set of published project photos.`;

    copy.append(meta, title, description);
    link.append(media, copy);
    article.appendChild(link);
    categorySpotlights.appendChild(article);
  });
}

function renderHomepageCatalogPreview() {
  if (!catalogsPreview) return;

  const library = getCatalogLibrary();
  if (!library.length) {
    catalogsPreview.innerHTML = '<p class="hint">Catalog groups will appear here once the catalog library data is connected.</p>';
    return;
  }

  catalogsPreview.innerHTML = '';
  library.forEach((group) => {
    const card = document.createElement('article');
    card.className = 'catalog-preview-card card';

    const head = document.createElement('div');
    head.className = 'catalog-preview-head';

    const brand = document.createElement('div');
    brand.className = 'catalog-preview-brand';

    const logo = document.createElement('div');
    logo.className = 'catalog-preview-logo';
    logo.textContent = String(group.mark || group.title || 'C').slice(0, 2).toUpperCase();

    const nameWrap = document.createElement('div');
    const title = document.createElement('h3');
    title.className = 'catalog-preview-title';
    title.textContent = group.title;
    const text = document.createElement('p');
    text.className = 'catalog-preview-text';
    text.textContent = group.description || '';
    nameWrap.append(title, text);

    brand.append(logo, nameWrap);

    const count = document.createElement('span');
    count.className = 'catalog-preview-count';
    count.textContent = `${Array.isArray(group.documents) ? group.documents.length : 0} files`;

    head.append(brand, count);

    const links = document.createElement('div');
    links.className = 'catalog-preview-links';
    (Array.isArray(group.documents) ? group.documents.slice(0, 2) : []).forEach((doc) => {
      const docLink = document.createElement('a');
      docLink.href = doc.href;
      docLink.target = '_blank';
      docLink.rel = 'noopener';
      docLink.textContent = doc.label;
      links.appendChild(docLink);
    });

    const openGroup = document.createElement('a');
    openGroup.className = 'btn btn-secondary';
    openGroup.href = `./catalogs.html#${encodeURIComponent(group.id)}`;
    openGroup.textContent = 'Open Group';

    card.append(head, links, openGroup);
    catalogsPreview.appendChild(card);
  });
}

function renderGallery(categorySlug, galleryData) {
  const container = document.getElementById('galleryContent');
  if (!container) return;

  container.innerHTML = '';

  const category = galleryData.categories.find((entry) => entry.slug === categorySlug) || galleryData.categories[0];
  if (!category) {
    container.innerHTML = '<p class="hint">No gallery categories available yet.</p>';
    return;
  }

  const list = Array.isArray(galleryData.itemsBySlug[category.slug]) ? galleryData.itemsBySlug[category.slug] : [];
  const scopedItems = list.map((entry) => ({
    ...entry,
    categoryName: category.name,
    categorySlug: category.slug
  }));
  const visibleCount = Math.min(
    scopedItems.length,
    galleryUiState.visibleBySlug[category.slug] || GALLERY_BATCH_SIZE
  );

  try {
    localStorage.setItem(`galleryList:${category.slug}`, JSON.stringify(list.map((item) => item.imageUrl)));
    localStorage.setItem(`galleryCategoryLabel:${category.slug}`, category.name);
  } catch {
    // Ignore storage errors and continue with regular links.
  }

  if (!scopedItems.length) {
    container.innerHTML = '<p class="hint">No published photos in this segment yet.</p>';
    return;
  }

  const fragment = document.createDocumentFragment();
  scopedItems.slice(0, visibleCount).forEach((item, index) => {
    fragment.appendChild(buildGalleryItem(item, { ...category, items: scopedItems }, index));
  });
  container.appendChild(fragment);

  if (galleryCount) {
    galleryCount.textContent =
      scopedItems.length > visibleCount
        ? `Showing ${visibleCount} of ${scopedItems.length} photos in ${category.name}.`
        : `Showing all ${scopedItems.length} photos in ${category.name}.`;
  }

  if (galleryMoreButton) {
    const hasMore = scopedItems.length > visibleCount;
    galleryMoreButton.hidden = !hasMore;
    galleryMoreButton.disabled = !hasMore;
  }
}

function updateGalleryHistory(categorySlug) {
  const url = new URL(location.href);
  url.searchParams.set('cat', categorySlug);
  history.replaceState(null, '', url.toString());
}

async function initGallery() {
  const cachedContent = readWebsiteContentCache();
  const cachedSettings = cachedContent?.settings || readCachedSiteSettings();
  if (cachedSettings) applySiteSettings(cachedSettings);

  const manifest = await loadManifest();

  const fallbackGalleryData = buildStaticGalleryData(manifest);
  const galleryData = hasRenderableGallery(cachedContent?.gallery) ? cachedContent.gallery : fallbackGalleryData;

  if (!galleryData.categories.length) {
    setGalleryHint('No published gallery items are available yet.');
    const container = document.getElementById('galleryContent');
    if (container) container.innerHTML = '<p class="hint">No published photos are available yet.</p>';
    if (galleryControls) galleryControls.innerHTML = '';
    if (galleryCount) galleryCount.textContent = '';
    if (galleryMoreButton) galleryMoreButton.hidden = true;
    return;
  }

  if (galleryData.source === 'supabase') {
    setGalleryHint('Browse by segment. Showing published gallery items from your website content tables.');
  } else {
    setGalleryHint('Browse by segment. Static gallery is shown first so the page opens faster while dynamic content updates in the background.');
  }

  let currentCategory = getInitialCategory(galleryData.categories);
  const handleCategoryChange = (nextCategory) => {
    currentCategory = nextCategory;
    galleryUiState.categorySlug = currentCategory;
    if (!galleryUiState.visibleBySlug[currentCategory]) galleryUiState.visibleBySlug[currentCategory] = GALLERY_BATCH_SIZE;
    renderGalleryControls(galleryData.categories, currentCategory, handleCategoryChange);
    updateGalleryHistory(currentCategory);
    renderGallery(currentCategory, galleryData);
  };

  galleryUiState.data = galleryData;
  galleryUiState.categorySlug = currentCategory;
  if (!galleryUiState.visibleBySlug[currentCategory]) galleryUiState.visibleBySlug[currentCategory] = GALLERY_BATCH_SIZE;

  renderGalleryControls(galleryData.categories, currentCategory, handleCategoryChange);
  renderFeaturedWork(galleryData);
  renderCategorySpotlights(galleryData);
  renderGallery(currentCategory, galleryData);

  galleryMoreButton?.addEventListener('click', () => {
    const activeGallery = galleryUiState.data;
    const activeSlug = galleryUiState.categorySlug;
    if (!activeGallery || !activeSlug) return;
    const list = Array.isArray(activeGallery.itemsBySlug[activeSlug]) ? activeGallery.itemsBySlug[activeSlug] : [];
    if (!list.length) return;
    const nextVisible = Math.min(
      list.length,
      (galleryUiState.visibleBySlug[activeSlug] || GALLERY_BATCH_SIZE) + GALLERY_BATCH_SIZE
    );
    galleryUiState.visibleBySlug[activeSlug] = nextVisible;
    renderGallery(activeSlug, activeGallery);
  });

  scheduleNonCriticalTask(async () => {
    const dynamicContent = await loadWebsiteContent();
    if (!dynamicContent) return;
    if (dynamicContent.settings) applySiteSettings(dynamicContent.settings);

    const nextGalleryData = hasRenderableGallery(dynamicContent.gallery) ? dynamicContent.gallery : fallbackGalleryData;
    writeWebsiteContentCache({
      settings: dynamicContent.settings || cachedSettings || null,
      gallery: nextGalleryData
    });

    if (!hasRenderableGallery(nextGalleryData)) return;
    galleryUiState.data = nextGalleryData;
    if (!nextGalleryData.categories.some((category) => category.slug === galleryUiState.categorySlug)) {
      galleryUiState.categorySlug = nextGalleryData.categories[0]?.slug || '';
    }
    renderGalleryControls(nextGalleryData.categories, galleryUiState.categorySlug, handleCategoryChange);
    renderFeaturedWork(nextGalleryData);
    renderCategorySpotlights(nextGalleryData);
    renderGallery(galleryUiState.categorySlug, nextGalleryData);
  });
}

renderHomepageCatalogPreview();

syncHeaderHeight();
window.addEventListener('resize', syncHeaderHeight);

if (galleryLightbox) {
  galleryLightbox.querySelectorAll('[data-lightbox-close]').forEach((element) => {
    element.addEventListener('click', () => closeGalleryLightbox());
  });

  galleryLightboxClose?.addEventListener('click', () => closeGalleryLightbox());
  galleryLightboxBack?.addEventListener('click', () => closeGalleryLightbox());
  galleryLightboxImage?.addEventListener('click', () => closeGalleryLightbox());
  galleryLightboxPrev?.addEventListener('click', () => stepGalleryLightbox(-1));
  galleryLightboxNext?.addEventListener('click', () => stepGalleryLightbox(1));
}

navToggle?.addEventListener('click', () => {
  const expanded = navToggle.getAttribute('aria-expanded') === 'true';
  navToggle.setAttribute('aria-expanded', String(!expanded));
  navWrap?.classList.toggle('nav-open', !expanded);
  syncHeaderHeight();
});

navMenu?.querySelectorAll('a').forEach((link) => {
  link.addEventListener('click', () => {
    if (window.innerWidth > 980) return;
    closeMobileNav();
  });
});

document.querySelectorAll('a[href^="#"]').forEach((link) => {
  link.addEventListener('click', (event) => {
    const href = link.getAttribute('href');
    if (!href || href === '#') return;
    if (!scrollToSection(href)) return;
    event.preventDefault();
    closeMobileNav();
  });
});

window.addEventListener('keydown', (event) => {
  if (!galleryLightbox?.hidden) {
    if (event.key === 'Escape') {
      event.preventDefault();
      closeGalleryLightbox();
      return;
    }
    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      stepGalleryLightbox(-1);
      return;
    }
    if (event.key === 'ArrowRight') {
      event.preventDefault();
      stepGalleryLightbox(1);
      return;
    }
  }

  if (event.key !== 'Escape') return;
  closeMobileNav();
});

window.addEventListener('load', () => {
  if (!window.location.hash) return;
  window.setTimeout(() => {
    scrollToSection(window.location.hash, { updateHistory: false });
  }, 0);
});

if (dateField) {
  const today = new Date();
  today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
  dateField.min = today.toISOString().split('T')[0];
}

form?.querySelectorAll('input, select, textarea').forEach((field) => {
  const eventName = field.type === 'checkbox' || field.tagName === 'SELECT' ? 'change' : 'input';
  field.addEventListener(eventName, () => {
    const wrapper = getFieldWrapper(field.name);
    if (wrapper) wrapper.dataset.touched = 'true';
    const fields = getFields();
    const errors = validate(fields);
    syncFieldStates(fields, errors);
    if (!Object.keys(errors).length && note?.dataset.state === 'error') {
      setFormNote('');
    }
  });

  field.addEventListener('blur', () => {
    const wrapper = getFieldWrapper(field.name);
    if (wrapper) wrapper.dataset.touched = 'true';
    const fields = getFields();
    const errors = validate(fields);
    syncFieldStates(fields, errors);
  });
});

form?.addEventListener('submit', async (event) => {
  event.preventDefault();
  setFormNote('');

  const fields = getFields();
  const errors = validate(fields);
  syncFieldStates(fields, errors, true);

  if (Object.keys(errors).length) {
    setFormNote('Please review the highlighted fields and try again.', 'error');
    return;
  }

  if (submitButton) {
    submitButton.disabled = true;
    submitButton.textContent = 'Sending...';
  }

  try {
    const result = await submitVisitRequest(fields);

    if (result.mode === 'mailto') {
      window.location.href = encodeMailto(fields);
      setFormNote('Opening your email client...', 'success');
    } else {
      resetVisitForm();
      setFormNote('Request sent successfully. We will get back to you shortly.', 'success');
    }
  } catch (error) {
    console.error(error);
    window.location.href = encodeMailto(fields);
    setFormNote('Could not reach the server. Opening your email client instead.', 'error');
  } finally {
    if (submitButton) {
      submitButton.disabled = false;
      submitButton.textContent = 'Send Request';
    }
  }
});

void initGallery();
