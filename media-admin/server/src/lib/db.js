import pg from 'pg';
import { env } from '../config/env.js';

const { Pool } = pg;

if (!env.databaseUrl) {
  throw new Error('DATABASE_URL is required to run the media admin API.');
}

const pool = new Pool({
  connectionString: env.databaseUrl
});

export async function query(text, params = []) {
  return pool.query(text, params);
}

export async function withTransaction(callback) {
  const client = await pool.connect();
  try {
    await client.query('begin');
    const result = await callback(client);
    await client.query('commit');
    return result;
  } catch (error) {
    await client.query('rollback');
    throw error;
  } finally {
    client.release();
  }
}

export async function closeDb() {
  await pool.end();
}
