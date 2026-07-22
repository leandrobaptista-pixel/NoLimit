#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
MANIFEST_JSON = ASSETS / "gallery-manifest.json"
MANIFEST_JS = ASSETS / "gallery-manifest.js"
OPTIMIZE_SCRIPT = ROOT / "scripts" / "optimize-homepage-images.sh"
VARIANTS_JS = ASSETS / "optimized" / "variants-manifest.js"
PRODUCTION_BASE_URL = "https://nolimitcontractor.pages.dev/"

CATEGORY_DIRS = {
    "Trim": "assets/00-Trim",
    "Wainscoting": "assets/00-Wainscoting",
    "Stairs": "assets/00-Stairs",
    "Ceiling": "assets/00-Ceiling",
    "Decks": "assets/00-Decks",
    "Kitchen & Vanities": "assets/00-kitchen & Vanities",
    "Fireplaces & Bars": "assets/00-Fireplaces & Bars",
    "Outside Doors & Windows": "assets/00-Outside Doors & Windows",
    "Pergola": "assets/00-Pergola",
    "Port & Portal": "assets/00-Port & Portal",
}

ALLOWED_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".JPG", ".JPEG", ".PNG", ".WEBP", ".GIF"}
SKIP_NAME_MARKERS = ("copy", "originalfullresolutionimage")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync local gallery folders into the Website 2 manifest, optimization pipeline, and optional auto-publish flow."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would change without writing manifests or running optimization.",
    )
    parser.add_argument(
        "--skip-optimize",
        action="store_true",
        help="Do not regenerate optimized homepage images after syncing the manifest.",
    )
    parser.add_argument(
        "--skip-check",
        action="store_true",
        help="Do not run the gallery verification after syncing.",
    )
    parser.add_argument(
        "--full-site-check",
        action="store_true",
        help="Also run the broader site validator after the gallery-specific checks.",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Stage gallery changes, create a commit, push to main, and rely on GitHub Actions auto-deploy.",
    )
    parser.add_argument(
        "--commit-message",
        default="Sync gallery photo folders",
        help="Commit message to use together with --publish.",
    )
    return parser.parse_args()


def is_gallery_asset(path: Path) -> bool:
    return path.is_file() and path.suffix in ALLOWED_SUFFIXES and not any(marker in path.name.lower() for marker in SKIP_NAME_MARKERS)


def build_manifest(previous_manifest: dict[str, list[str]] | None = None) -> dict[str, list[str]]:
    manifest: dict[str, list[str]] = {}
    previous_manifest = previous_manifest or {}
    for category, rel_dir in CATEGORY_DIRS.items():
        directory = ROOT / rel_dir
        items: list[str] = []
        if directory.exists():
            current_files = [f"{rel_dir}/{file.name}" for file in directory.iterdir() if is_gallery_asset(file)]
            existing_order = [asset for asset in previous_manifest.get(category, []) if asset in current_files]
            new_assets = sorted((asset for asset in current_files if asset not in existing_order), key=str.lower)
            items = existing_order + new_assets
        manifest[category] = items
    return manifest


def load_current_manifest() -> dict[str, list[str]]:
    if not MANIFEST_JSON.exists():
        return {}
    try:
        data = json.loads(MANIFEST_JSON.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except json.JSONDecodeError:
        return {}


def load_published_manifest_order() -> dict[str, list[str]]:
    try:
        result = subprocess.run(
            ["git", "-C", str(ROOT), "show", "HEAD:assets/gallery-manifest.json"],
            check=True,
            capture_output=True,
            text=True,
        )
        data = json.loads(result.stdout)
        return data if isinstance(data, dict) else {}
    except (subprocess.CalledProcessError, json.JSONDecodeError):
        return {}


def compare_manifests(previous: dict[str, list[str]], current: dict[str, list[str]]) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    added: dict[str, list[str]] = {}
    removed: dict[str, list[str]] = {}

    for category in CATEGORY_DIRS:
        before = set(previous.get(category, []))
        after = set(current.get(category, []))
        next_added = sorted(after - before, key=str.lower)
        next_removed = sorted(before - after, key=str.lower)
        if next_added:
            added[category] = next_added
        if next_removed:
            removed[category] = next_removed

    return added, removed


def write_manifests(manifest: dict[str, list[str]]) -> None:
    payload = json.dumps(manifest, indent=2, ensure_ascii=True) + "\n"
    MANIFEST_JSON.write_text(payload, encoding="utf-8")
    MANIFEST_JS.write_text(f"window.GALLERY_MANIFEST = {payload.rstrip()};\n", encoding="utf-8")


def run_command(command: list[str], label: str) -> None:
    print(f"\n[{label}]")
    subprocess.run(command, cwd=ROOT, check=True)


def read_manifest_js_payload() -> dict[str, list[str]]:
    if not MANIFEST_JS.exists():
        raise FileNotFoundError(f"{MANIFEST_JS} not found")

    text = MANIFEST_JS.read_text(encoding="utf-8").strip()
    prefix = "window.GALLERY_MANIFEST = "
    suffix = ";"
    if not text.startswith(prefix) or not text.endswith(suffix):
        raise ValueError("gallery-manifest.js is not in the expected format")

    payload = text[len(prefix):-len(suffix)].strip()
    data = json.loads(payload)
    return data if isinstance(data, dict) else {}


def verify_gallery_pipeline(manifest: dict[str, list[str]]) -> None:
    print("\n[verify gallery pipeline]")

    missing_assets: list[str] = []
    for items in manifest.values():
        for asset in items:
            if not (ROOT / asset).exists():
                missing_assets.append(asset)

    manifest_js_payload = read_manifest_js_payload()
    if manifest_js_payload != manifest:
        raise SystemExit("gallery-manifest.js does not match gallery-manifest.json")

    if missing_assets:
        preview = "\n".join(f" - {asset}" for asset in missing_assets[:20])
        raise SystemExit(f"Gallery manifest references missing files:\n{preview}")

    if not VARIANTS_JS.exists():
        raise SystemExit("optimized variants manifest was not generated")

    print("Gallery manifest, JS manifest, and optimized variants are in sync.")


def build_stage_targets() -> list[str]:
    targets = ["assets/gallery-manifest.json", "assets/gallery-manifest.js", "assets/optimized"]
    targets.extend(CATEGORY_DIRS.values())
    return targets


def stage_commit_and_push(commit_message: str) -> None:
    stage_targets = build_stage_targets()
    run_command(["git", "-C", str(ROOT), "add", "--", *stage_targets], "stage gallery changes")

    status = subprocess.run(
        ["git", "-C", str(ROOT), "diff", "--cached", "--quiet"],
        cwd=ROOT,
        check=False,
    )
    if status.returncode == 0:
        print("\n[publish]\nNo staged gallery changes were detected. Nothing to commit.")
        return

    run_command(["git", "-C", str(ROOT), "commit", "-m", commit_message], "commit gallery changes")
    run_command(["git", "-C", str(ROOT), "push", "origin", "main"], "push to main")


def print_summary(added: dict[str, list[str]], removed: dict[str, list[str]]) -> None:
    total_added = sum(len(items) for items in added.values())
    total_removed = sum(len(items) for items in removed.values())
    print("\n[gallery sync summary]")
    print(f"Added: {total_added} file(s)")
    print(f"Removed: {total_removed} file(s)")

    if not added and not removed:
        print("No manifest differences found. The folder structure is already in sync.")
        return

    for category in CATEGORY_DIRS:
        if category in added:
            print(f"\n+ {category}")
            for asset in added[category]:
                print(f"  - {asset}")
                print(f"    host: {PRODUCTION_BASE_URL}{asset}")
        if category in removed:
            print(f"\n- {category}")
            for asset in removed[category]:
                print(f"  - {asset}")


def main() -> int:
    args = parse_args()

    current_manifest = load_current_manifest()
    baseline_manifest = load_published_manifest_order() or current_manifest
    next_manifest = build_manifest(baseline_manifest)
    added, removed = compare_manifests(current_manifest, next_manifest)

    print_summary(added, removed)

    if args.dry_run:
        print("\n[dry-run]\nNo files were written.")
        return 0

    write_manifests(next_manifest)
    print("\n[manifest]\nassets/gallery-manifest.json and assets/gallery-manifest.js updated.")

    if not args.skip_optimize:
        run_command([str(OPTIMIZE_SCRIPT)], "optimize homepage images")

    if not args.skip_check:
        verify_gallery_pipeline(next_manifest)
        if args.full_site_check:
            run_command(["python3", str(ROOT / "scripts" / "check_links.py")], "verify full site links")

    if args.publish:
        stage_commit_and_push(args.commit_message)
        print("\n[publish]\nPush completed. GitHub Actions should deploy the updated website automatically.")
    else:
        print(
            "\n[next step]\nIf everything looks good, publish with:\n"
            "python3 scripts/gallery_sync_pipeline.py --publish"
        )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as error:
        print(f"\n[error]\nCommand failed with exit code {error.returncode}: {' '.join(error.cmd)}", file=sys.stderr)
        raise SystemExit(error.returncode)
