import { env } from '../config/env.js';
import { generateCaption } from '../lib/caption.js';
import { generatePromotionalImage } from '../lib/promo-art.js';
import { readStoredFile, saveGeneratedAsset, saveUploadedFile } from '../lib/storage.js';
import { findCategory, listCategories } from '../repositories/category-repository.js';
import {
  createMedia,
  getMediaById,
  listMedia,
  updateGeneratedAsset,
  updateMediaDetails,
  updateMediaStatus
} from '../repositories/media-repository.js';

function createError(message, statusCode = 400, details = null) {
  const error = new Error(message);
  error.statusCode = statusCode;
  error.details = details;
  return error;
}

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (value === 'true') return true;
  if (value === 'false') return false;
  return undefined;
}

export async function fetchCategories() {
  return listCategories();
}

export async function fetchMedia(filters = {}) {
  return listMedia({
    categorySlug: String(filters.category || '').trim(),
    search: String(filters.search || '').trim(),
    published: parseBoolean(filters.published)
  });
}

export async function fetchPublishedMedia(filters = {}) {
  return listMedia({
    categorySlug: String(filters.category || '').trim(),
    search: String(filters.search || '').trim(),
    published: true
  });
}

export async function uploadMedia({ file, body }) {
  if (!file) {
    throw createError('An image file is required.', 400);
  }

  if (!String(file.mimetype || '').startsWith('image/')) {
    throw createError('Only image uploads are allowed.', 400);
  }

  const title = String(body.title || '').trim();
  if (!title) {
    throw createError('Title is required.', 400);
  }

  const category = await findCategory({
    categoryId: String(body.categoryId || '').trim(),
    categorySlug: String(body.categorySlug || '').trim()
  });

  if (!category) {
    throw createError('A valid category is required.', 400);
  }

  const storedFile = await saveUploadedFile(file, title);
  const captionText = generateCaption({
    categoryName: category.name,
    categorySlug: category.slug,
    phone: env.brandPhone,
    email: env.brandEmail
  });

  return createMedia({
    title,
    categoryId: category.id,
    originalFileName: file.originalname,
    originalPath: storedFile.relativePath,
    originalUrl: storedFile.url,
    captionText,
    commercialCopy: env.brandCopy,
    mimeType: file.mimetype,
    fileSizeBytes: file.size || 0,
    published: parseBoolean(body.published) ?? false,
    usedInSocial: parseBoolean(body.usedInSocial) ?? false
  });
}

export async function generateArtwork(mediaId) {
  const record = await getMediaById(mediaId);
  if (!record) {
    throw createError('Media item not found.', 404);
  }

  const captionText = generateCaption({
    categoryName: record.categoryName,
    categorySlug: record.categorySlug,
    phone: env.brandPhone,
    email: env.brandEmail
  });

  const sourceImage = await readStoredFile(record.originalPath, record.originalUrl);
  const promoBuffer = await generatePromotionalImage({
    sourceImage,
    categoryName: record.categoryName
  });

  const storedAsset = await saveGeneratedAsset(promoBuffer, `${record.title}-${record.categorySlug}`);

  return updateGeneratedAsset(record.id, {
    generatedArtPath: storedAsset.relativePath,
    generatedArtUrl: storedAsset.url,
    captionText
  });
}

export async function patchMediaStatus(mediaId, body) {
  const record = await getMediaById(mediaId);
  if (!record) {
    throw createError('Media item not found.', 404);
  }

  return updateMediaStatus(mediaId, {
    published: parseBoolean(body.published),
    usedInSocial: parseBoolean(body.usedInSocial)
  });
}

export async function patchMediaDetails(mediaId, body) {
  const record = await getMediaById(mediaId);
  if (!record) {
    throw createError('Media item not found.', 404);
  }

  const nextTitle = String(body.title || '').trim() || undefined;
  const nextCategory = body.categoryId || body.categorySlug
    ? await findCategory({
        categoryId: String(body.categoryId || '').trim(),
        categorySlug: String(body.categorySlug || '').trim()
      })
    : null;

  if ((body.categoryId || body.categorySlug) && !nextCategory) {
    throw createError('Category not found.', 404);
  }

  const category = nextCategory || { id: undefined, name: record.categoryName, slug: record.categorySlug };
  const nextCaption = generateCaption({
    categoryName: category.name,
    categorySlug: category.slug,
    phone: env.brandPhone,
    email: env.brandEmail
  });

  return updateMediaDetails(mediaId, {
    title: nextTitle,
    categoryId: nextCategory?.id,
    captionText: nextCaption
  });
}
