# Deploy Option 1: Render or Railway

This path keeps the backend as a standard Node.js service and is the best fit when you want the fastest route to production with minimal platform-specific rewrites.

## What gets deployed

- API: [`media-admin/server`](/Users/lemacbook/Desktop/WebSite%202/media-admin/server)
- Admin frontend: [`media-admin/client`](/Users/lemacbook/Desktop/WebSite%202/media-admin/client)
- Database: PostgreSQL or Supabase Postgres
- Storage: local disk for MVP, or Supabase Storage using `STORAGE_DRIVER=supabase`

## Recommended architecture

- Cloudflare Pages: public website and optional admin frontend
- Render or Railway: Express API
- Supabase: PostgreSQL database and object storage

## Render setup

1. Create a new Web Service from this repository.
2. Set the service root directory to `media-admin/server`.
3. Build command: `npm install`
4. Start command: `npm start`
5. Add environment variables from [`server/.env.example`](/Users/lemacbook/Desktop/WebSite%202/media-admin/server/.env.example)
6. If you want infrastructure-as-code, start from [`render.example.yaml`](/Users/lemacbook/Desktop/WebSite%202/media-admin/render.example.yaml)

### Render environment variables

- `DATABASE_URL`
- `PUBLIC_BASE_URL`
- `MEDIA_ADMIN_PASSWORD`
- `MEDIA_ADMIN_TOKEN`
- `MEDIA_ADMIN_SESSION_SECRET`
- `MEDIA_ADMIN_SESSION_TTL_HOURS`
- `STORAGE_DRIVER`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_STORAGE_BUCKET`
- `BRAND_PHONE`
- `BRAND_EMAIL`
- `BRAND_COPY`
- `LOGO_PATH`
- `ANNIVERSARY_BADGE_PATH`

### Suggested production values

- `STORAGE_DRIVER=supabase`
- `MEDIA_ADMIN_PASSWORD=<shared admin password used by Gallery Control sign-in>`
- `MEDIA_ADMIN_SESSION_SECRET=<long random secret used to sign browser sessions>`
- `MEDIA_ADMIN_TOKEN=<optional long random token for legacy direct bearer access>`
- `LOGO_PATH=https://nolimitcontractor.pages.dev/assets/brand.png`
- `ANNIVERSARY_BADGE_PATH=https://nolimitcontractor.pages.dev/assets/anniversary-18.png`

## Railway setup

1. Create a new service from the same repository.
2. Set the service root directory to `media-admin/server`.
3. Railway can use:
   - Build command: `npm install`
   - Start command: `npm start`
4. Add the same environment variables listed above.
5. If you prefer Railway config as code, use [`server/railway.json`](/Users/lemacbook/Desktop/WebSite%202/media-admin/server/railway.json)

## Optional admin frontend deploy

You can deploy the React admin as:

- a second Render static service
- a second Railway service
- or, more naturally for this repo, a dedicated Cloudflare Pages project

For the admin frontend:

1. Root directory: `media-admin/client`
2. Build command: `npm install && npm run build`
3. Output directory: `dist`
4. Environment variable:
   - `VITE_API_BASE_URL=https://your-api-domain`

## Docker alternative

The API also includes a Dockerfile at [`server/Dockerfile`](/Users/lemacbook/Desktop/WebSite%202/media-admin/server/Dockerfile) if you prefer Docker-based deploys.
