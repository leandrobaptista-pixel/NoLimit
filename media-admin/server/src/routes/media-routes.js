import { Router } from 'express';
import multer from 'multer';
import { env } from '../config/env.js';
import { asyncHandler } from '../middleware/async-handler.js';
import { fetchMedia, generateArtwork, patchMediaDetails, patchMediaStatus, uploadMedia } from '../services/media-service.js';

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: env.maxUploadMb * 1024 * 1024 }
});

export const mediaRouter = Router();

mediaRouter.get(
  '/',
  asyncHandler(async (req, res) => {
    const media = await fetchMedia(req.query);
    res.json({ media });
  })
);

mediaRouter.post(
  '/upload',
  upload.single('image'),
  asyncHandler(async (req, res) => {
    const media = await uploadMedia({ file: req.file, body: req.body });
    res.status(201).json({ media });
  })
);

mediaRouter.post(
  '/:mediaId/generate-art',
  asyncHandler(async (req, res) => {
    const media = await generateArtwork(req.params.mediaId);
    res.json({ media });
  })
);

mediaRouter.patch(
  '/:mediaId/status',
  asyncHandler(async (req, res) => {
    const media = await patchMediaStatus(req.params.mediaId, req.body);
    res.json({ media });
  })
);

mediaRouter.patch(
  '/:mediaId',
  asyncHandler(async (req, res) => {
    const media = await patchMediaDetails(req.params.mediaId, req.body);
    res.json({ media });
  })
);
