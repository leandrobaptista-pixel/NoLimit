#!/bin/zsh
set -euo pipefail

# Read-only preflight for the No Limit publication path. It intentionally does
# not alter DNS, Git configuration, the remote, branches, or repository files.

repo_root="${1:-$(git rev-parse --show-toplevel)}"
remote_name="${2:-origin}"

echo "[1/4] macOS resolver"
scutil --dns | sed -n '1,120p'

echo "[2/4] github.com resolution"
python3 - <<'PY'
import socket
print(socket.getaddrinfo("github.com", 443, type=socket.SOCK_STREAM)[0][4][0])
PY

echo "[3/4] SSH authentication"
ssh -T git@github-nolimit || test $? -eq 1

echo "[4/4] repository transport"
git -C "$repo_root" ls-remote "$remote_name" HEAD

echo "GitHub publication path: healthy"
