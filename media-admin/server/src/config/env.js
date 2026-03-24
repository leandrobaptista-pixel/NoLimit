import dotenv from 'dotenv';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const SERVER_ROOT = path.resolve(__dirname, '../..');
const MODULE_ROOT = path.resolve(SERVER_ROOT, '..');
const REPO_ROOT = path.resolve(MODULE_ROOT, '..');

function resolvePath(maybeRelative, fallbackAbsolute) {
  if (!maybeRelative) return fallbackAbsolute;
  if (/^https?:\/\//i.test(maybeRelative)) return maybeRelative;
  return path.isAbsolute(maybeRelative)
    ? maybeRelative
    : path.resolve(SERVER_ROOT, maybeRelative);
}

function toNumber(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

export const env = {
  port: toNumber(process.env.PORT, 4000),
  databaseUrl: String(process.env.DATABASE_URL || '').trim(),
  publicBaseUrl: String(process.env.PUBLIC_BASE_URL || 'http://localhost:4000').trim().replace(/\/$/, ''),
  mediaAdminPassword: String(process.env.MEDIA_ADMIN_PASSWORD || '').trim(),
  mediaAdminToken: String(process.env.MEDIA_ADMIN_TOKEN || '').trim(),
  mediaAdminSessionSecret: String(process.env.MEDIA_ADMIN_SESSION_SECRET || '').trim(),
  mediaAdminSessionTtlHours: Math.max(1, toNumber(process.env.MEDIA_ADMIN_SESSION_TTL_HOURS, 24)),
  maxUploadMb: toNumber(process.env.MAX_UPLOAD_MB, 20),
  storageDriver: String(process.env.STORAGE_DRIVER || 'local').trim(),
  supabaseUrl: String(process.env.SUPABASE_URL || '').trim().replace(/\/$/, ''),
  supabaseServiceRoleKey: String(process.env.SUPABASE_SERVICE_ROLE_KEY || '').trim(),
  supabaseStorageBucket: String(process.env.SUPABASE_STORAGE_BUCKET || 'nolimit-media').trim(),
  serverRoot: SERVER_ROOT,
  moduleRoot: MODULE_ROOT,
  repoRoot: REPO_ROOT,
  storageRoot: path.resolve(SERVER_ROOT, 'storage'),
  brandName: String(process.env.BRAND_NAME || 'No Limit').trim(),
  brandPhone: String(process.env.BRAND_PHONE || '(732) 555-0178').trim(),
  brandEmail: String(process.env.BRAND_EMAIL || 'hello@nolimitcontractor.com').trim(),
  brandCopy: String(
    process.env.BRAND_COPY ||
      'Custom finish carpentry with premium detail, clean execution, and dependable delivery.'
  ).trim(),
  logoPath: resolvePath(process.env.LOGO_PATH, path.resolve(REPO_ROOT, 'assets/brand.png')),
  anniversaryBadgePath: resolvePath(
    process.env.ANNIVERSARY_BADGE_PATH,
    path.resolve(REPO_ROOT, 'assets/anniversary-18.png')
  )
};
