# Deploy Automation (GitHub -> Cloudflare Pages)

This repository now deploys automatically to Cloudflare Pages on every push to `main`.

## One-time setup

1. In GitHub repo settings, add this Actions secret:
   - `CLOUDFLARE_API_TOKEN`
2. Cloudflare API token permissions:
   - `Account` -> `Cloudflare Pages:Edit`
   - `Zone` -> `Zone:Read`
   - Scope: your account that contains project `nolimitcontractor`
   - Save the token as a single line, without quotes, extra spaces, or line breaks
   - If deploy still reports invalid header/token characters, delete the secret and create it again from a fresh token copy
3. `CLOUDFLARE_ACCOUNT_ID` is already set in workflow as:
   - `47f6c1dee23ca2adee98043a07a5cd23`

## How deploy works

- Trigger: push to `main` (or manual run via Actions).
- Pre-check: workflow fails if any file is larger than 25 MiB.
- Credentials are normalized before deploy to strip hidden line breaks from secrets.
- Publish command:
  - `wrangler pages deploy . --project-name=nolimitcontractor --branch=main`

## Daily usage

```bash
git add .
git commit -m "your changes"
git push origin main
```

After push, deployment runs automatically in GitHub Actions.
