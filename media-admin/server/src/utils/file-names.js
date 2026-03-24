import path from 'node:path';

export function slugify(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

export function getSafeExtension(fileName, mimeType = '') {
  const extFromName = path.extname(fileName || '').toLowerCase();
  if (extFromName) return extFromName;

  const map = {
    'image/jpeg': '.jpg',
    'image/png': '.png',
    'image/webp': '.webp',
    'image/heic': '.heic',
    'image/heif': '.heif'
  };

  return map[mimeType] || '.jpg';
}

export function createStoredFileName({ title, originalName, prefix = 'asset' }) {
  const ext = getSafeExtension(originalName);
  const cleanTitle = slugify(title) || prefix;
  return `${prefix}-${cleanTitle}-${Date.now()}${ext}`;
}
