import { query } from '../lib/db.js';

export async function listCategories() {
  const { rows } = await query(
    `select id, name, slug
     from categories
     order by name asc`
  );

  return rows;
}

export async function findCategory({ categoryId, categorySlug }) {
  if (categoryId) {
    const { rows } = await query(
      `select id, name, slug from categories where id = $1 limit 1`,
      [categoryId]
    );
    return rows[0] || null;
  }

  if (categorySlug) {
    const { rows } = await query(
      `select id, name, slug from categories where slug = $1 limit 1`,
      [categorySlug]
    );
    return rows[0] || null;
  }

  return null;
}
