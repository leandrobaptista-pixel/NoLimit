import { query } from '../lib/db.js';

function normalizeMediaRow(row) {
  if (!row) return null;
  return {
    id: row.id,
    title: row.title,
    categoryId: row.category_id,
    categoryName: row.category_name,
    categorySlug: row.category_slug,
    originalUrl: row.original_url,
    generatedArtUrl: row.generated_art_url,
    captionText: row.caption_text,
    commercialCopy: row.commercial_copy,
    mimeType: row.mime_type,
    fileSizeBytes: Number(row.file_size_bytes || 0),
    published: Boolean(row.published),
    usedInSocial: Boolean(row.used_in_social),
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

export async function listMedia({ categorySlug = '', search = '', published } = {}) {
  const conditions = [];
  const params = [];

  if (categorySlug) {
    params.push(categorySlug);
    conditions.push(`c.slug = $${params.length}`);
  }

  if (search) {
    params.push(`%${search}%`);
    conditions.push(`m.title ilike $${params.length}`);
  }

  if (typeof published === 'boolean') {
    params.push(published);
    conditions.push(`m.published = $${params.length}`);
  }

  const whereClause = conditions.length ? `where ${conditions.join(' and ')}` : '';

  const { rows } = await query(
    `select
      m.id,
      m.title,
      m.category_id,
      m.original_url,
      m.generated_art_url,
      m.caption_text,
      m.commercial_copy,
      m.mime_type,
      m.file_size_bytes,
      m.published,
      m.used_in_social,
      m.created_at,
      m.updated_at,
      c.name as category_name,
      c.slug as category_slug
     from media_items m
     join categories c on c.id = m.category_id
     ${whereClause}
     order by m.created_at desc`,
    params
  );

  return rows.map(normalizeMediaRow);
}

export async function getMediaById(id) {
  const { rows } = await query(
    `select
      m.id,
      m.title,
      m.category_id,
      m.original_path,
      m.original_url,
      m.generated_art_path,
      m.generated_art_url,
      m.caption_text,
      m.commercial_copy,
      m.mime_type,
      m.file_size_bytes,
      m.published,
      m.used_in_social,
      m.created_at,
      m.updated_at,
      c.name as category_name,
      c.slug as category_slug
     from media_items m
     join categories c on c.id = m.category_id
     where m.id = $1
     limit 1`,
    [id]
  );

  const row = rows[0];
  if (!row) return null;

  return {
    ...normalizeMediaRow(row),
    originalPath: row.original_path,
    generatedArtPath: row.generated_art_path
  };
}

export async function createMedia(payload) {
  const { rows } = await query(
    `insert into media_items (
      title,
      category_id,
      original_filename,
      original_path,
      original_url,
      caption_text,
      commercial_copy,
      mime_type,
      file_size_bytes,
      published,
      used_in_social
    ) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
    returning id`,
    [
      payload.title,
      payload.categoryId,
      payload.originalFileName,
      payload.originalPath,
      payload.originalUrl,
      payload.captionText,
      payload.commercialCopy,
      payload.mimeType,
      payload.fileSizeBytes,
      payload.published,
      payload.usedInSocial
    ]
  );

  return getMediaById(rows[0].id);
}

export async function updateMediaDetails(id, payload) {
  await query(
    `update media_items
     set title = coalesce($2, title),
         category_id = coalesce($3, category_id),
         caption_text = coalesce($4, caption_text),
         updated_at = now()
     where id = $1`,
    [id, payload.title ?? null, payload.categoryId ?? null, payload.captionText ?? null]
  );

  return getMediaById(id);
}

export async function updateMediaStatus(id, payload) {
  await query(
    `update media_items
     set published = coalesce($2, published),
         used_in_social = coalesce($3, used_in_social),
         updated_at = now()
     where id = $1`,
    [id, payload.published ?? null, payload.usedInSocial ?? null]
  );

  return getMediaById(id);
}

export async function updateGeneratedAsset(id, payload) {
  await query(
    `update media_items
     set generated_art_path = $2,
         generated_art_url = $3,
         caption_text = $4,
         updated_at = now()
     where id = $1`,
    [id, payload.generatedArtPath, payload.generatedArtUrl, payload.captionText]
  );

  return getMediaById(id);
}
