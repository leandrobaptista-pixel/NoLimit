# Daily website measurement and security review

Public measurement stores only fixed event/dimension counters grouped by UTC hour in `site_metric_hourly`. There are no raw event rows, visitor IDs, IP columns, arbitrary URLs, query strings, form values, user-agent strings or user-level histories. The frontend uses a yes/no per-tab session-storage flag for approximate visit counts. It respects GPC, DNT and the local opt-out on `/website-privacy.html`.

The same-origin Cloudflare endpoint accepts only a bounded JSON object and a fixed event vocabulary. It forwards no client network headers. The private Edge Function increments counters through a service-only RPC. RLS and revoked browser grants deny reads and direct writes to counters. Aggregates expire after 90 days. Request counts come from the actual saved intake, not the optional browser signal.

`python3 scripts/daily_site_report.py --daily` reads the last completed 20:00–20:00 America/New_York period, accounting for daylight saving. Without `--daily`, it previews the last day up to now. It uses `SUPABASE_ACCESS_TOKEN` or the existing local Supabase CLI token file and never prints credentials. Scheduled delivery runs in this Codex chat; the local host must be available.

Reports include pages, approximate tab visits, gallery interest, broad traffic sources, busy hours, contact/catalog clicks, form starts, saved requests and form error counts. Do not present unique people, unidentified sources, blocked tracking or missing data as known facts. These counters are not fraud detection and cannot prove an attack.

Security reporting is separate: existing RLS protects administrative data, existing workspace snapshots/backups remain in place, and a new trigger records removed record counts with the authorized account. Staff Visit Request edits use one service-only transaction to update and append audit metadata. Public counters cannot identify anonymous visitors. Account audit records contain operational metadata only, without customer messages, photos or internal notes. Review an account action before treating it as misconduct.

Existing backups preserve application records and storage manifests, not copies of file bytes. WAF alerts, failed authentication reports, MFA enforcement and independent storage-file recovery are separate capabilities, not implicitly enabled by this change. Do not claim complete intrusion monitoring or file recovery.

## Protected Admin entry

The topbar chart/shield icon and `Website Insights` navigation item open `#website-insights`. Production visibility follows active Owner/Developer membership, matching existing audit controls. Ordinary Admin, Manager and other memberships do not receive access. Direct navigation shows a restricted message; hiding the icon is not the authorization boundary.

The `get-site-report` Edge Function validates the signed-in user and current active membership on every read. The service-only `get_private_site_report` RPC repeats membership validation and bounds the Today, Yesterday, Last 7 days and last completed 8 PM reporting periods on the server. Browser roles cannot execute this RPC directly. Responses use `Cache-Control: no-store`; report data is cleared on sign-out/account changes, and entering the view revalidates access. No secrets or management token are delivered to the browser.

The report displays verified MFA enrollment for the caller only. It does not enroll factors, change credentials, enforce MFA or configure copies of uploaded file bytes. These remaining safeguards are labeled explicitly. Local preview contains layout placeholders and never calls the protected report service.
