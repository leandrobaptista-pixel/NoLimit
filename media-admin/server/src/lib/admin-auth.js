import { createHmac, timingSafeEqual } from 'node:crypto';
import { env } from '../config/env.js';

const SESSION_PREFIX = 'nlm';

function toBuffer(value) {
  return Buffer.from(String(value || ''), 'utf8');
}

function safeEqual(left, right) {
  const leftBuffer = toBuffer(left);
  const rightBuffer = toBuffer(right);
  if (!leftBuffer.length || !rightBuffer.length) return false;
  if (leftBuffer.length !== rightBuffer.length) return false;
  return timingSafeEqual(leftBuffer, rightBuffer);
}

function encodeBase64Url(value) {
  return Buffer.from(value, 'utf8').toString('base64url');
}

function decodeBase64Url(value) {
  return Buffer.from(String(value || ''), 'base64url').toString('utf8');
}

function getSessionSigningSecret() {
  return env.mediaAdminSessionSecret || env.mediaAdminPassword || env.mediaAdminToken;
}

function signPayloadSegment(segment) {
  const secret = getSessionSigningSecret();
  if (!secret) return '';
  return createHmac('sha256', secret).update(segment).digest('base64url');
}

function getAcceptedLoginSecrets() {
  return [env.mediaAdminPassword, env.mediaAdminToken].filter(Boolean);
}

export function extractAdminToken(req) {
  const authorization = String(req.headers.authorization || '').trim();
  if (authorization.toLowerCase().startsWith('bearer ')) {
    return authorization.slice(7).trim();
  }

  return String(req.headers['x-media-admin-token'] || '').trim();
}

export function canIssueAdminSessions() {
  return Boolean(getAcceptedLoginSecrets().length && getSessionSigningSecret());
}

export function canAuthorizeAdminRequests() {
  return Boolean(getSessionSigningSecret() || env.mediaAdminToken);
}

export function verifyAdminLoginSecret(receivedSecret) {
  const trimmed = String(receivedSecret || '').trim();
  if (!trimmed) return false;
  return getAcceptedLoginSecrets().some((expectedSecret) => safeEqual(trimmed, expectedSecret));
}

export function issueAdminSession() {
  if (!canIssueAdminSessions()) {
    throw new Error('Admin session signing is not configured on the server.');
  }

  const issuedAt = Math.floor(Date.now() / 1000);
  const expiresAtSeconds = issuedAt + env.mediaAdminSessionTtlHours * 60 * 60;
  const payload = {
    sub: 'media-admin',
    role: 'admin',
    iat: issuedAt,
    exp: expiresAtSeconds
  };
  const payloadSegment = encodeBase64Url(JSON.stringify(payload));
  const signature = signPayloadSegment(payloadSegment);

  return {
    token: `${SESSION_PREFIX}.${payloadSegment}.${signature}`,
    expiresAt: new Date(expiresAtSeconds * 1000).toISOString()
  };
}

export function verifyAdminSession(token) {
  const trimmed = String(token || '').trim();
  if (!trimmed) return null;

  const [prefix, payloadSegment, signature] = trimmed.split('.');
  if (prefix !== SESSION_PREFIX || !payloadSegment || !signature) return null;

  const expectedSignature = signPayloadSegment(payloadSegment);
  if (!safeEqual(signature, expectedSignature)) return null;

  try {
    const payload = JSON.parse(decodeBase64Url(payloadSegment));
    const expiration = Number(payload?.exp || 0);
    if (!expiration || expiration <= Math.floor(Date.now() / 1000)) return null;
    if (payload?.role !== 'admin') return null;
    return {
      ...payload,
      expiresAt: new Date(expiration * 1000).toISOString()
    };
  } catch {
    return null;
  }
}

export function verifyAdminAccessToken(token) {
  const trimmed = String(token || '').trim();
  if (!trimmed) return false;
  if (env.mediaAdminToken && safeEqual(trimmed, env.mediaAdminToken)) return true;
  return Boolean(verifyAdminSession(trimmed));
}
