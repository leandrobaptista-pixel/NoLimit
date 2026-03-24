import fs from 'node:fs/promises';
import path from 'node:path';
import { env } from '../config/env.js';
import { createStoredFileName } from '../utils/file-names.js';

async function ensureDirectory(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

function toPublicUrl(relativePath) {
  const normalized = String(relativePath || '').replace(/\\/g, '/');
  return `${env.publicBaseUrl}/storage/${normalized}`;
}

function toSupabasePublicUrl(relativePath) {
  const normalized = String(relativePath || '').replace(/\\/g, '/');
  return `${env.supabaseUrl}/storage/v1/object/public/${env.supabaseStorageBucket}/${normalized}`;
}

async function saveBufferToSupabase(relativePath, buffer, contentType) {
  if (!env.supabaseUrl || !env.supabaseServiceRoleKey || !env.supabaseStorageBucket) {
    throw new Error('SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, and SUPABASE_STORAGE_BUCKET are required for STORAGE_DRIVER=supabase.');
  }

  const normalized = String(relativePath || '').replace(/^\/+/, '').replace(/\\/g, '/');
  const uploadUrl = `${env.supabaseUrl}/storage/v1/object/${env.supabaseStorageBucket}/${normalized}`;
  const response = await fetch(uploadUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${env.supabaseServiceRoleKey}`,
      apikey: env.supabaseServiceRoleKey,
      'Content-Type': contentType || 'application/octet-stream',
      'x-upsert': 'true'
    },
    body: buffer
  });

  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Supabase storage upload failed: ${response.status} ${details}`.trim());
  }

  return {
    relativePath: normalized,
    url: toSupabasePublicUrl(normalized)
  };
}

export async function saveUploadedFile(file, title) {
  const fileName = createStoredFileName({
    title,
    originalName: file.originalname,
    prefix: 'upload'
  });

  const relativePath = path.posix.join('uploads', fileName);
  if (env.storageDriver === 'supabase') {
    const saved = await saveBufferToSupabase(relativePath, file.buffer, file.mimetype);
    return {
      fileName,
      relativePath: saved.relativePath,
      absolutePath: null,
      url: saved.url
    };
  }

  const directory = path.resolve(env.storageRoot, 'uploads');
  await ensureDirectory(directory);
  const absolutePath = path.resolve(env.storageRoot, relativePath);
  await fs.writeFile(absolutePath, file.buffer);

  return {
    fileName,
    relativePath,
    absolutePath,
    url: toPublicUrl(relativePath)
  };
}

export async function saveGeneratedAsset(buffer, baseTitle) {
  const safeBase = createStoredFileName({
    title: baseTitle,
    originalName: `${baseTitle}.png`,
    prefix: 'promo'
  }).replace(/\.[a-z0-9]+$/i, '.png');

  const relativePath = path.posix.join('generated', safeBase);
  if (env.storageDriver === 'supabase') {
    const saved = await saveBufferToSupabase(relativePath, buffer, 'image/png');
    return {
      fileName: safeBase,
      relativePath: saved.relativePath,
      absolutePath: null,
      url: saved.url
    };
  }

  const directory = path.resolve(env.storageRoot, 'generated');
  await ensureDirectory(directory);
  const absolutePath = path.resolve(env.storageRoot, relativePath);
  await fs.writeFile(absolutePath, buffer);

  return {
    fileName: safeBase,
    relativePath,
    absolutePath,
    url: toPublicUrl(relativePath)
  };
}

export async function readStoredFile(relativePath, fallbackUrl = '') {
  const normalized = String(relativePath || '').replace(/^\/+/, '').replace(/\\/g, '/');

  if (env.storageDriver === 'supabase') {
    const downloadUrl = fallbackUrl || toSupabasePublicUrl(normalized);
    const response = await fetch(downloadUrl);
    if (!response.ok) {
      const details = await response.text().catch(() => '');
      throw new Error(`Supabase storage download failed: ${response.status} ${details}`.trim());
    }
    return Buffer.from(await response.arrayBuffer());
  }

  return fs.readFile(path.resolve(env.storageRoot, normalized));
}
