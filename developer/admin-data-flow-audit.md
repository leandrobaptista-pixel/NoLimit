# Admin data-flow audit — 2026-09-20

## Scope reviewed

- Shared workspace read/write flow (`pullBetaWorkspace`, `pushBetaWorkspace`, `savePreviewState`).
- Vendor and subcontractor directory rendering and project assignment options.
- Change-order service catalog linkage.
- Invoice line-item creation and reusable service catalog flow.
- Role/access separation from commercial vendor/subcontractor records.

## Findings and corrections

1. The Vendors screen filtered only records whose type was `Vendor`.  `Subcontractor` records, including records created by an authorized administrator, were therefore excluded from that directory.  The directory and project selector now include both commercial types and display their type explicitly.
2. Change orders created outside the custom-service category were stored without a reusable catalog reference.  New change orders now create or reuse one catalog service, matched by normalized title, while preserving the order's stored title, description and amount as its historical record.
3. Invoice users lacked an explicit way to create a reusable service from the line being entered.  The invoice form now includes a `Save new service to catalog` action; it pre-fills the service dialog and selects the saved service back into that invoice line.
4. Commercial directory data remains distinct from authentication and permissions.  Adding a vendor or subcontractor neither grants access nor changes a user's role.

## Validation performed

- `node --check no-limit-admin-beta/app.js` completed successfully.
- `git diff --check` completed successfully before commit.
- Production read synchronization was previously checked from the public Admin: authorized user data and Visit Requests loaded successfully, including a refresh without an error.
- No write test was performed against production so no real operational data, messages, invoices or invitations were created.

## Publication status and limit

The implementation is committed as `c661aed` (`Unify subcontractor directory and service catalog flows`). It is **not yet published**. The official Git push is blocked before authentication because this machine's terminal DNS cannot resolve `github.com`. No fallback upload, cache workaround or direct production mutation was used. The public version therefore remains unchanged until the official repository connection is restored and the normal deployment completes.

## Recovery point

- Prior committed state: `b4b1da1`
- Current committed state: `c661aed`
- No uncommitted files at the time this audit was created.
