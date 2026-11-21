# Private cards and financing

The Admin now has a separate `#credit` route. It reads only protected RPCs backed by `nl_credit_*` tables. Credit records never enter `beta_workspaces`, operational reports, ordinary workspace exports or shared snapshots. There is no browser-storage copy and no financial data in the source bundle.

Access is an explicit two-slot allowlist of Supabase Auth UUIDs, plus active No Limit organization membership. An Owner, Developer, Finance group or profile display name does not grant access. The application fails closed when the migration or allowlist is absent. Direct table access is revoked for browser roles and RLS is enabled without permissive policies. Administrative database operators retain their infrastructure privileges.

## Activation order

1. Apply `supabase/migrations/20261121010000_private_credit_control.sql` to the Admin project `powxzhcscqkhtaicxkyt`, using an authorized database administrator.
2. Apply the separately delivered private account activation SQL. It resolves the verified existing Leandro and Geanniny accounts by their confirmed login emails and freezes their Auth UUIDs. Keep this file outside the public website repository. The Gmail manager account is excluded.
3. Deploy the frontend changes through the existing GitHub / Cloudflare Pages process. Sign out and sign in again so the two accounts load the new access context.
4. In Cards & Financing, select the separately delivered private import JSON and review the counts and budget. Confirm the initial import. Import is atomic, validated on the server, and rejected when any debt already exists. Do not commit or host the import JSON or original workbook.
5. Verify both allowed accounts can read and write, and an unrelated account and an anonymous visitor cannot use the route or RPCs. Do not mark the rollout complete until these live checks pass.

## Behavior

Payments, actual interest and purchases / fees update balances. Monthly totals use posting dates. Current statement minimums use the payment's associated statement due date, even when its posting month differs. A missing minimum remains unavailable, while a real zero is preserved. Incorrect entries are voided and retained, then corrected with a new entry. Entry request UUIDs prevent duplicate retries after a lost response. Debt details and budgets use optimistic revisions.

The workbook's original monthly payoff budget applies to selected cards. Additional cards and financing are tracked separately unless a card is explicitly added to that plan. The projection pays all monthly minimums, then allocates extra cash to the highest APR, stopping if the budget cannot cover minimums. An absent card minimum uses the workbook estimate. Projections do not post interest to the ledger. Promotional balances and deadlines are manual statement inputs and are not incorporated into the standard APR simulation.

New tables are deliberately excluded from general workspace backup routines to preserve the requested access boundary. Database-level recovery and any dedicated credit backup need the same restricted access policy. No independent credit backup exporter is introduced by this change.

## Validation

Run `node --test tests/credit-math.test.mjs tests/credit-access.test.mjs tests/operational-workspace.test.mjs`.

For an isolated PostgreSQL check without touching Supabase: `npm ci --prefix tests/credit-db`, then `npm test --prefix tests/credit-db`. The PGlite fixture verifies actual SQL privileges, the lack of an Owner override, active membership, atomic imports, stale revisions, entry validation, void history and request retry behavior. This fixture is separate from live rollout verification.

Browser checks used synthetic records. Production application and live two-account checks remain pending database access.

Security patterns follow [Supabase database functions](https://supabase.com/docs/guides/database/functions) and [row level security](https://supabase.com/docs/guides/database/postgres/row-level-security).
