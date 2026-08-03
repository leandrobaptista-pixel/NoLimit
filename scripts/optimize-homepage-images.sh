#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
ASSETS_DIR="$ROOT/assets"
MANIFEST_JSON="$ASSETS_DIR/gallery-manifest.json"
THUMB_DIR="$ASSETS_DIR/optimized/thumb"
FEATURE_DIR="$ASSETS_DIR/optimized/feature"
VARIANTS_JS="$ASSETS_DIR/optimized/variants-manifest.js"

export ROOT ASSETS_DIR MANIFEST_JSON THUMB_DIR FEATURE_DIR VARIANTS_JS

python3 <<'PY'
from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

ROOT = Path(os.environ["ROOT"])
ASSETS_DIR = Path(os.environ["ASSETS_DIR"])
MANIFEST_JSON = Path(os.environ["MANIFEST_JSON"])
THUMB_DIR = Path(os.environ["THUMB_DIR"])
FEATURE_DIR = Path(os.environ["FEATURE_DIR"])
VARIANTS_JS = Path(os.environ["VARIANTS_JS"])

THUMB_LIMIT = 12
FEATURE_LIMIT = 2
THUMB_WIDTH = 640
FEATURE_WIDTH = 1280
THUMB_QUALITY = "65"
FEATURE_QUALITY = "72"

manifest = json.loads(MANIFEST_JSON.read_text())

selected: dict[str, dict[str, str]] = {}
expected_thumb_paths: set[Path] = set()
expected_feature_paths: set[Path] = set()

for _category, items in manifest.items():
    items = [str(item).strip() for item in items if str(item).strip()]
    for index, rel in enumerate(items):
        normalized = rel.lstrip("./")
        entry = selected.setdefault(normalized, {})
        source = ROOT / normalized
        relative_asset = Path(normalized)
        if relative_asset.parts and relative_asset.parts[0] == "assets":
            relative_asset = Path(*relative_asset.parts[1:])
        stem = relative_asset.stem + ".jpg"
        parent = relative_asset.parent

        if index < THUMB_LIMIT:
          thumb_path = THUMB_DIR / parent / stem
          if thumb_path.exists():
              thumb_path.unlink()
          thumb_path.parent.mkdir(parents=True, exist_ok=True)
          subprocess.run(
              [
                  "sips",
                  "-s",
                  "format",
                  "jpeg",
                  "-s",
                  "formatOptions",
                  THUMB_QUALITY,
                  "--resampleWidth",
                  str(THUMB_WIDTH),
                  str(source),
                  "--out",
                  str(thumb_path),
              ],
              check=True,
              stdout=subprocess.DEVNULL,
              stderr=subprocess.DEVNULL,
          )
          expected_thumb_paths.add(thumb_path)
          entry["thumb"] = str(thumb_path.relative_to(ROOT)).replace("\\", "/")

        if index < FEATURE_LIMIT:
          feature_path = FEATURE_DIR / parent / stem
          if feature_path.exists():
              feature_path.unlink()
          feature_path.parent.mkdir(parents=True, exist_ok=True)
          subprocess.run(
              [
                  "sips",
                  "-s",
                  "format",
                  "jpeg",
                  "-s",
                  "formatOptions",
                  FEATURE_QUALITY,
                  "--resampleWidth",
                  str(FEATURE_WIDTH),
                  str(source),
                  "--out",
                  str(feature_path),
              ],
              check=True,
              stdout=subprocess.DEVNULL,
              stderr=subprocess.DEVNULL,
          )
          expected_feature_paths.add(feature_path)
          entry["feature"] = str(feature_path.relative_to(ROOT)).replace("\\", "/")

serializable = {key: value for key, value in sorted(selected.items()) if value}

for directory, expected_paths in ((THUMB_DIR, expected_thumb_paths), (FEATURE_DIR, expected_feature_paths)):
    if not directory.exists():
        continue
    for generated in directory.rglob("*.jpg"):
        if generated not in expected_paths:
            generated.unlink()

VARIANTS_JS.parent.mkdir(parents=True, exist_ok=True)
VARIANTS_JS.write_text(
    "window.OPTIMIZED_IMAGE_VARIANTS = "
    + json.dumps(serializable, indent=2, sort_keys=True)
    + ";\n"
)

print(f"generated {len(serializable)} optimized image mappings")
PY
