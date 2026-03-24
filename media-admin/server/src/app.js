import cors from 'cors';
import express from 'express';
import helmet from 'helmet';
import morgan from 'morgan';
import path from 'node:path';
import { env } from './config/env.js';
import { errorHandler, notFoundHandler } from './middleware/error-handler.js';
import { requireAdminApi } from './middleware/require-admin-api.js';
import { categoryRouter } from './routes/category-routes.js';
import { healthRouter } from './routes/health-routes.js';
import { mediaRouter } from './routes/media-routes.js';
import { publicRouter } from './routes/public-routes.js';

export function createApp() {
  const app = express();

  app.use(helmet());
  app.use(cors());
  app.use(express.json({ limit: `${env.maxUploadMb}mb` }));
  app.use(express.urlencoded({ extended: true }));
  app.use(morgan('dev'));

  app.use('/storage', express.static(path.resolve(env.storageRoot)));
  app.use('/api/health', healthRouter);
  app.use('/api/categories', requireAdminApi, categoryRouter);
  app.use('/api/media', requireAdminApi, mediaRouter);
  app.use('/api/public', publicRouter);

  app.use(notFoundHandler);
  app.use(errorHandler);

  return app;
}
