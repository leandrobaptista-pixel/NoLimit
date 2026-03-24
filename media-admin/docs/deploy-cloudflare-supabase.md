# Deploy Option 2: Cloudflare + Supabase Hybrid

This path keeps your existing Cloudflare Pages setup for the website and uses Supabase for data and storage, while the new media API runs as a separate Node service.

## Why this is the best fit for the current repo

- The live website is already deployed as a static Cloudflare Pages project.
- The repository already uses Supabase for other parts of the business tools.
- The new media admin needs server-side image generation, which is much simpler and safer in a Node runtime than forcing it into the static site deployment.

## Recommended hybrid layout

- Main website: existing Cloudflare Pages project
- Media admin frontend: new Cloudflare Pages project from `media-admin/client`
- Media API: Render or Railway from `media-admin/server`
- Database: Supabase Postgres
- File storage: Supabase Storage bucket `nolimit-media`

## Supabase steps

1. In Supabase SQL Editor, run [`schema.sql`](/Users/lemacbook/Desktop/WebSite%202/media-admin/database/schema.sql)
2. Run [`seed.sql`](/Users/lemacbook/Desktop/WebSite%202/media-admin/database/seed.sql)
3. Run [`supabase-setup.sql`](/Users/lemacbook/Desktop/WebSite%202/media-admin/database/supabase-setup.sql)
4. Copy your project values:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `DATABASE_URL`

## API environment

Use these values in the Node API service:

- `DATABASE_URL=<Supabase Postgres connection string>`
- `STORAGE_DRIVER=supabase`
- `MEDIA_ADMIN_TOKEN=<long random token used by the Gallery Control page>`
- `SUPABASE_URL=<your Supabase URL>`
- `SUPABASE_SERVICE_ROLE_KEY=<service role key>`
- `SUPABASE_STORAGE_BUCKET=nolimit-media`
- `PUBLIC_BASE_URL=<your API base URL>`

## Cloudflare Pages admin frontend

Create a second Pages project for the admin interface:

1. Repository: same repo
2. Root directory: `media-admin/client`
3. Build command: `npm install && npm run build`
4. Build output directory: `dist`
5. Environment variable:
   - `VITE_API_BASE_URL=https://your-media-api-domain`

## Public website integration

The public website can consume the new media library through:

- `GET https://your-media-api-domain/api/public/media`
- `GET https://your-media-api-domain/api/public/media?category=trim`
- `GET https://your-media-api-domain/api/categories`

That means:

- the main website stays on Cloudflare Pages
- the media catalog becomes dynamic
- uploads and generated art live in Supabase Storage

## Asset overlays in production

For easiest production setup, point these environment variables to public brand assets:

- `LOGO_PATH=https://nolimitcontractor.pages.dev/assets/brand.png`
- `ANNIVERSARY_BADGE_PATH=https://nolimitcontractor.pages.dev/assets/anniversary-18.png`
