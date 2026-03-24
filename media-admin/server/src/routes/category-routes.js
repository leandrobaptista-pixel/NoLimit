import { Router } from 'express';
import { asyncHandler } from '../middleware/async-handler.js';
import { fetchCategories } from '../services/media-service.js';

export const categoryRouter = Router();

categoryRouter.get(
  '/',
  asyncHandler(async (_req, res) => {
    const categories = await fetchCategories();
    res.json({ categories });
  })
);
