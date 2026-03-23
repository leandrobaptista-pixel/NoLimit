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

syncHeaderHeight();
window.addEventListener('resize', syncHeaderHeight);

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
  if (event.key !== 'Escape') return;
  closeMobileNav();
});

window.addEventListener('load', () => {
  if (!window.location.hash) return;
  window.setTimeout(() => {
    scrollToSection(window.location.hash, { updateHistory: false });
  }, 0);
});

function encodeMailto(fields) {
  const to = form?.dataset?.to || 'info@your-company.com';
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

const FALLBACK_GALLERY = {
  'Trim': ['assets/00-Trim/IMG_7031.jpg'],
  'Wainscoting': ['assets/00-Wainscoting/IMG_1663.JPG'],
  'Stairs': ['assets/00-Stairs/IMG_9268.JPG'],
  'Ceiling': ['assets/00-Ceiling/IMG_6650.jpg'],
  'Decks': ['assets/00-Decks/IMG_4694.jpg'],
  'Kitchen & Vanities': ['assets/00-kitchen & Vanities/IMG_5865.JPG'],
  'Fireplaces & Bars': ['assets/00-Fireplaces & Bars/IMG_1816.JPG'],
  'Outside Doors & Windows': ['assets/00-Outside Doors & Windows/IMG_3311.JPG'],
  'Pergola': ['assets/00-Pergola/IMG_2744.jpg'],
  'Port & Portal': ['assets/00-Port & Portal/IMG_5811.JPG']
};

function toURL(path) {
  return String(path || '').split('/').map(encodeURIComponent).join('/');
}

function dedupeByBase(list) {
  const preferOrder = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'JPG', 'JPEG', 'PNG', 'WEBP', 'GIF'];
  const byBase = new Map();

  list.forEach((item) => {
    const file = item.split('/').pop() || '';
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
  if (urlCat && categories.includes(urlCat)) return urlCat;

  const firstActive = document.querySelector('.filter-btn.active')?.dataset?.cat;
  if (firstActive && categories.includes(firstActive)) return firstActive;

  return categories[0] || 'Trim';
}

function bindCategoryButtons(categories, onChange) {
  const buttons = document.querySelectorAll('.filter-btn');
  buttons.forEach((button) => {
    const cat = button.dataset.cat;
    if (!categories.includes(cat)) {
      button.disabled = true;
      button.title = 'No photos available in this segment yet';
      return;
    }

    button.addEventListener('click', () => {
      buttons.forEach((b) => b.classList.remove('active'));
      button.classList.add('active');
      onChange(cat);
    });
  });
}

function buildGalleryItem(src, cat, index) {
  const fig = document.createElement('figure');
  const link = document.createElement('a');
  link.href = `viewer.html?img=${encodeURIComponent(src)}&cat=${encodeURIComponent(cat)}&i=${index}`;
  link.target = '_blank';
  link.rel = 'noopener';

  const img = document.createElement('img');
  img.loading = 'lazy';
  img.decoding = 'async';
  img.fetchPriority = 'low';
  img.src = toURL(src);
  img.alt = `${cat} project`;
  img.onerror = () => {
    img.src = 'assets/placeholder.svg';
  };

  const cap = document.createElement('figcaption');
  const fileName = decodeURIComponent(src.split('/').pop() || 'photo');
  const baseName = fileName.replace(/\.[^.]+$/, '');
  cap.textContent = baseName;

  link.appendChild(img);
  fig.appendChild(link);
  fig.appendChild(cap);
  return fig;
}

function renderGallery(cat, manifest) {
  const container = document.getElementById('galleryContent');
  if (!container) return;

  container.innerHTML = '';

  const raw = Array.isArray(manifest[cat]) ? manifest[cat] : [];
  const list = dedupeByBase(raw);

  try {
    localStorage.setItem(`galleryList:${cat}`, JSON.stringify(list));
  } catch {
    // Ignore storage errors and continue with regular links.
  }

  const fragment = document.createDocumentFragment();
  list.forEach((src, index) => {
    fragment.appendChild(buildGalleryItem(src, cat, index));
  });
  container.appendChild(fragment);
}

(async function initGallery() {
  const manifest = await loadManifest();
  const categories = Object.keys(manifest).filter((cat) => Array.isArray(manifest[cat]) && manifest[cat].length);
  if (!categories.length) return;

  let currentCat = getInitialCategory(categories);
  bindCategoryButtons(categories, (cat) => {
    currentCat = cat;
    renderGallery(currentCat, manifest);
  });
  renderGallery(currentCat, manifest);
})();
