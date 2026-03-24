# No Limit Media Admin

Modular web system for managing project photos, publishing gallery-ready assets, and generating square promotional artwork for the No Limit brand.

## Stack

- Backend: Node.js + Express
- Frontend: React + Vite
- Database: PostgreSQL
- Storage: Local disk (adapter-ready for S3 later)

## Features

- Upload original work photos with title and category
- Organize and list images by category
- Filter gallery records in the admin panel
- Publish and unpublish media for website consumption
- Generate 1080x1080 promotional artwork with:
  - base image
  - brand logo overlay
  - 18 years seal overlay
  - category title
  - fixed commercial copy
  - phone and email
- Auto-generate caption text based on category
- Expose public REST endpoints for the main website

## Folder Structure

```text
media-admin/
  client/     React admin panel
  server/     Express API, upload logic, promo art generation
  database/   PostgreSQL schema and seed data
```

## Quick Start

1. Create a PostgreSQL database.
2. Run [`schema.sql`](/Users/lemacbook/Desktop/WebSite%202/media-admin/database/schema.sql).
3. Run [`seed.sql`](/Users/lemacbook/Desktop/WebSite%202/media-admin/database/seed.sql).
4. Copy [`server/.env.example`](/Users/lemacbook/Desktop/WebSite%202/media-admin/server/.env.example) to `.env` and adjust values.
5. Install and run the API:
   - `cd media-admin/server`
   - `npm install`
   - `npm run dev`
6. Install and run the admin panel:
   - `cd media-admin/client`
   - `npm install`
   - `npm run dev`

## Deployment Paths

- [Render or Railway deploy](/Users/lemacbook/Desktop/WebSite%202/media-admin/docs/deploy-render-railway.md)
- [Cloudflare + Supabase hybrid deploy](/Users/lemacbook/Desktop/WebSite%202/media-admin/docs/deploy-cloudflare-supabase.md)

## Public API Examples

- `GET /api/health`
- `POST /api/auth/login`
- `GET /api/auth/session`
- `GET /api/public/categories`
- `GET /api/categories`
- `GET /api/media?category=trim`
- `GET /api/public/media?category=trim`
- `POST /api/media/upload`
- `POST /api/media/:id/generate-art`
- `PATCH /api/media/:id/status`

## Admin protection

- `GET /api/public/*` stays public for website consumption
- `POST /api/auth/login` accepts the shared admin password and returns a signed session token
- `GET /api/categories` and all `/api/media/*` routes accept `Authorization: Bearer <session-token>`
- `MEDIA_ADMIN_TOKEN` can still be used for legacy direct bearer access if needed
- set `MEDIA_ADMIN_PASSWORD` and `MEDIA_ADMIN_SESSION_SECRET` in the API environment before going live

## Notes

- The current storage driver is local disk for a fast MVP.
- The backend was kept isolated from the existing static website so the live Cloudflare Pages deployment keeps working unchanged.
- The API structure is ready for future social scheduling and S3 storage adapters.
