import { Router } from 'express';
import { asyncHandler } from '../middleware/async-handler.js';
import { fetchCategories } from '../services/media-service.js';
import { fetchPublishedMedia } from '../services/media-service.js';

export const publicRouter = Router();

publicRouter.get(
  '/categories',
  asyncHandler(async (_req, res) => {
    const categories = await fetchCategories();
    res.json({ categories });
  })
);

publicRouter.get(
  '/media',
  asyncHandler(async (req, res) => {
    const media = await fetchPublishedMedia(req.query);
    res.json({ media });
  })
);
