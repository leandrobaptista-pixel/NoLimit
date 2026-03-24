import { timingSafeEqual } from 'node:crypto';
import { env } from '../config/env.js';

function extractToken(req) {
  const authorization = String(req.headers.authorization || '').trim();
  if (authorization.toLowerCase().startsWith('bearer ')) {
    return authorization.slice(7).trim();
  }

  return String(req.headers['x-media-admin-token'] || '').trim();
}

function tokensMatch(received, expected) {
  if (!received || !expected) return false;

  const left = Buffer.from(received);
  const right = Buffer.from(expected);
  if (left.length !== right.length) return false;

  return timingSafeEqual(left, right);
}

export function requireAdminApi(req, res, next) {
  if (!env.mediaAdminToken) {
    res.status(503).json({
      message: 'Media admin token is not configured on the server yet.'
    });
    return;
  }

  const token = extractToken(req);
  if (!token) {
    res.status(401).json({
      message: 'Admin authorization is required for media management.'
    });
    return;
  }

  if (!tokensMatch(token, env.mediaAdminToken)) {
    res.status(403).json({
      message: 'Invalid media admin token.'
    });
    return;
  }

  next();
}
