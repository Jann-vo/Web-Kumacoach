#!/bin/bash
# KumaCoach Deploy Script
# Run: bash deploy.sh
# Copies HTML + all local images into this folder with relative paths for Netlify

SRC="/Users/cps/.gemini/antigravity-ide/scratch/kumacoach-website/kumacoach_website.html"
DEST_DIR="$(cd "$(dirname "$0")" && pwd)"
IMG_DIR="$DEST_DIR/assets"
HTML_OUT="$DEST_DIR/index.html"

mkdir -p "$IMG_DIR"

echo "📦 Copying HTML..."
cp "$SRC" "$HTML_OUT"

echo "🖼️  Copying images and rewriting paths..."

# Find all local absolute paths in the HTML and copy files
grep -oE '/Users/cps/[^"'\'']+\.(png|jpg|jpeg|webp|gif|svg)' "$HTML_OUT" | sort -u | while read -r path; do
  if [ -f "$path" ]; then
    filename=$(basename "$path")
    # Sanitize filename (remove spaces)
    safe=$(echo "$filename" | tr ' ' '_')
    cp "$path" "$IMG_DIR/$safe"
    # Replace in HTML
    escaped_path=$(echo "$path" | sed 's/[\/&]/\\&/g')
    sed -i '' "s|$escaped_path|assets/$safe|g" "$HTML_OUT"
    echo "  ✓ $filename → assets/$safe"
  else
    echo "  ✗ Missing: $path"
  fi
done

echo "🎓 Copying certificate images..."
CER_SRC="/Users/cps/Downloads/Cer Kuma"
CER_DEST="$IMG_DIR/cer"
mkdir -p "$CER_DEST"
for f in "$CER_SRC"/*.jpeg "$CER_SRC"/*.jpg; do
  [ -f "$f" ] || continue
  fname=$(basename "$f" | tr ' ' '_')
  cp "$f" "$CER_DEST/$fname"
done
sed -i '' "s|/Users/cps/Downloads/Cer Kuma/|assets/cer/|g" "$HTML_OUT"
echo "  ✓ Certificates copied"

echo "🏋️  Converting feedback training images to JPG (compressed)..."
TRAINING_SRC="/Users/cps/Downloads/Kuma Coach/Bài Đăng Training Kuma"
TRAINING_DEST="$IMG_DIR/training"
mkdir -p "$TRAINING_DEST"
# Remove old files to start clean
rm -f "$TRAINING_DEST"/*.png "$TRAINING_DEST"/*.jpg 2>/dev/null
count=0
for i in $(seq 181 263); do
  src="$TRAINING_SRC/${i}.png"
  if [ -f "$src" ]; then
    sips -Z 540 -s format jpeg -s formatOptions 72 "$src" --out "$TRAINING_DEST/${i}.jpg" > /dev/null 2>&1
    count=$((count+1))
  fi
done
echo "  ✓ $count training images converted to JPG"

echo "🔗 Rewriting JS feedback gallery path and extension..."
sed -i '' "s|const BASE = '/Users/cps/Downloads/Kuma Coach/Bài Đăng Training Kuma/';|const BASE = 'assets/training/';|g" "$HTML_OUT"
sed -i '' "s|const EXT = '.png';|const EXT = '.jpg';|g" "$HTML_OUT"

echo ""
echo "✅ Deploy ready at: $DEST_DIR"
echo "   Upload the entire kumacoach-deploy/ folder to Netlify"
echo "   Or drag & drop at: https://app.netlify.com/drop"
