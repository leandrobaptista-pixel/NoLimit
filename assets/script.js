const yearEl = document.getElementById('year');
if (yearEl) yearEl.textContent = new Date().getFullYear();

const header = document.querySelector('.site-header');
const navWrap = document.querySelector('.nav');
const navMenu = document.getElementById('primaryNav');
const navToggle = document.querySelector('.menu-toggle');
const form = document.getElementById('visitForm');
const note = document.getElementById('formNote');

function syncHeaderHeight() {
  if (!header) return;
  const height = Math.ceil(header.getBoundingClientRect().height);
  document.documentElement.style.setProperty('--header-h', `${height}px`);
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
    navToggle?.setAttribute('aria-expanded', 'false');
    navWrap?.classList.remove('nav-open');
    syncHeaderHeight();
  });
});

window.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return;
  navToggle?.setAttribute('aria-expanded', 'false');
  navWrap?.classList.remove('nav-open');
  syncHeaderHeight();
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

function getFields() {
  const data = {};
  if (!form) return data;
  new FormData(form).forEach((value, key) => {
    data[key] = value;
  });
  return data;
}

function validate(fields) {
  const errs = [];
  if (!fields.name) errs.push('Full Name is required.');
  if (!fields.email) errs.push('Email is required.');
  if (fields.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(fields.email)) {
    errs.push('Enter a valid email.');
  }
  if (form && !form.querySelector('input[name="agree"]')?.checked) {
    errs.push('Consent is required.');
  }
  return errs;
}

form?.addEventListener('submit', (event) => {
  event.preventDefault();
  if (note) note.textContent = '';

  const fields = getFields();
  const errs = validate(fields);

  if (errs.length) {
    if (note) note.textContent = errs.join(' ');
    return;
  }

  window.location.href = encodeMailto(fields);
  if (note) note.textContent = 'Opening your email client...';
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
