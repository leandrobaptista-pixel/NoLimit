import { Router } from 'express';
import { asyncHandler } from '../middleware/async-handler.js';
import {
  canIssueAdminSessions,
  extractAdminToken,
  issueAdminSession,
  verifyAdminAccessToken,
  verifyAdminLoginSecret,
  verifyAdminSession
} from '../lib/admin-auth.js';

export const authRouter = Router();

authRouter.post(
  '/login',
  asyncHandler(async (req, res) => {
    if (!canIssueAdminSessions()) {
      res.status(503).json({
        message: 'Admin password login is not configured on the server yet.'
      });
      return;
    }

    const password = String(req.body?.password || '').trim();
    if (!password) {
      res.status(400).json({
        message: 'Admin password is required.'
      });
      return;
    }

    if (!verifyAdminLoginSecret(password)) {
      res.status(401).json({
        message: 'Invalid admin password.'
      });
      return;
    }

    const session = issueAdminSession();
    res.json({
      token: session.token,
      expiresAt: session.expiresAt
    });
  })
);

authRouter.get(
  '/session',
  asyncHandler(async (req, res) => {
    const token = extractAdminToken(req);
    if (!token) {
      res.status(401).json({
        authenticated: false,
        message: 'Admin session not found.'
      });
      return;
    }

    if (!verifyAdminAccessToken(token)) {
      res.status(401).json({
        authenticated: false,
        message: 'Admin session is invalid or expired.'
      });
      return;
    }

    const session = verifyAdminSession(token);
    res.json({
      authenticated: true,
      mode: session ? 'session' : 'token',
      expiresAt: session?.expiresAt || null
    });
  })
);
