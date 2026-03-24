import fs from 'node:fs/promises';
import path from 'node:path';
import { createApp } from './app.js';
import { env } from './config/env.js';

async function ensureStorageDirectories() {
  await fs.mkdir(path.resolve(env.storageRoot, 'uploads'), { recursive: true });
  await fs.mkdir(path.resolve(env.storageRoot, 'generated'), { recursive: true });
}

async function start() {
  await ensureStorageDirectories();
  const app = createApp();

  app.listen(env.port, () => {
    console.log(`No Limit media admin API listening on http://localhost:${env.port}`);
  });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
