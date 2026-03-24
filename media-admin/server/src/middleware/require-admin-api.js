import { canAuthorizeAdminRequests, extractAdminToken, verifyAdminAccessToken } from '../lib/admin-auth.js';

export function requireAdminApi(req, res, next) {
  if (!canAuthorizeAdminRequests()) {
    res.status(503).json({
      message: 'Media admin authentication is not configured on the server yet.'
    });
    return;
  }

  const token = extractAdminToken(req);
  if (!token) {
    res.status(401).json({
      message: 'Admin authorization is required for media management.'
    });
    return;
  }

  if (!verifyAdminAccessToken(token)) {
    res.status(403).json({
      message: 'Invalid media admin session.'
    });
    return;
  }

  next();
}
