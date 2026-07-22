#!/bin/zsh
set -euo pipefail

ROOT="/Users/lemacbook/Desktop/WebSite 2"
ASSETS_DIR="$ROOT/assets"
THUMB_DIR="$ASSETS_DIR/optimized/thumb"
FEATURE_DIR="$ASSETS_DIR/optimized/feature"

typeset -a FILES=(
  "00-Trim/IMG_7031.jpg"
  "00-Wainscoting/IMG_1663.JPG"
  "00-Stairs/IMG_9268.JPG"
  "00-Ceiling/IMG_6650.jpg"
  "00-Decks/IMG_4694.jpg"
  "00-kitchen & Vanities/IMG_5865.JPG"
  "00-Fireplaces & Bars/IMG_1816.JPG"
  "00-Outside Doors & Windows/IMG_3311.JPG"
  "00-Pergola/IMG_2744.jpg"
  "00-Port & Portal/IMG_5811.JPG"
)

mkdir -p "$THUMB_DIR" "$FEATURE_DIR"

for rel in "${FILES[@]}"; do
  source_file="$ASSETS_DIR/$rel"
  base_dir="$(dirname "$rel")"
  base_name="$(basename "$rel")"
  stem="${base_name%.*}"

  thumb_out="$THUMB_DIR/$base_dir/$stem.jpg"
  feature_out="$FEATURE_DIR/$base_dir/$stem.jpg"

  mkdir -p "$(dirname "$thumb_out")" "$(dirname "$feature_out")"

  sips -s format jpeg -s formatOptions 65 --resampleWidth 640 "$source_file" --out "$thumb_out" >/dev/null
  sips -s format jpeg -s formatOptions 72 --resampleWidth 1280 "$source_file" --out "$feature_out" >/dev/null

  echo "optimized $rel"
done
