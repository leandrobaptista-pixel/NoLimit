# GitHub connectivity audit

## Purpose

Use this record before treating a failed `git fetch`, `git push`, or deployment
as a repository or application failure. The check distinguishes DNS, network
route, SSH authentication, and remote-repository failures without changing the
No Limit site or production data.

## Current incident — 2026-09-17

- Network interface `en0` was active with a private LAN address.
- macOS reported **`No DNS configuration available`** via `scutil --dns`.
- SSH failed before authentication with: `Could not resolve hostname github.com`.
- The Git remote and the dedicated `github-nolimit` SSH host alias were valid.
- Several tunnel interfaces were active. VPN/WARP or a malformed DNS profile
  must be checked before changing repository credentials.

## Required verification order

1. Confirm macOS has one or more DNS resolvers (`scutil --dns`).
2. Resolve `github.com` using the system resolver.
3. Confirm SSH transport with `ssh -T git@github-nolimit`.
4. Confirm the exact remote with `git ls-remote origin HEAD`.
5. Only then run `git push origin HEAD:main`.

## Safe operating rule

Never work around a DNS failure by replacing the remote, changing repository
history, using a hard-coded GitHub IP address, or publishing a second copy of
the site. Fix the network resolver first; the normal Git remote remains the
single publication path.

## Recovery verification

After DNS is restored, record the timestamp and the outputs of steps 2–4 in
the current change log. A successful `git ls-remote` proves DNS, TLS/SSH route,
key selection, and repository access independently of the application deploy.
