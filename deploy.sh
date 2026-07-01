#!/bin/bash
# KumaCoach Deploy Script — build + git push
# Run: bash ~/web-kumacoach/Web-Kumacoach/deploy.sh

SRC="/Users/cps/.gemini/antigravity-ide/scratch/kumacoach-website/kumacoach_website.html"
DEST_DIR="$(cd "$(dirname "$0")" && pwd)"
HTML_OUT="$DEST_DIR/index.html"
ASSETS="$DEST_DIR/assets"

CER_DIR="$ASSETS/cer"
ATHLETES_DIR="$ASSETS/athletes"
EBOOKS_DIR="$ASSETS/ebooks"
LOGOS_DIR="$ASSETS/logos"
COACHES_DIR="$ASSETS/coaches"
TRAINING_DIR="$ASSETS/training"

rm -f "$ASSETS"/*.png "$ASSETS"/*.jpg "$ASSETS"/*.jpeg 2>/dev/null
mkdir -p "$CER_DIR" "$ATHLETES_DIR" "$EBOOKS_DIR" "$LOGOS_DIR" "$COACHES_DIR" "$TRAINING_DIR"

echo "📦 Copying HTML..."
cp "$SRC" "$HTML_OUT"

copy_rewrite() {
  local src="$1" dest_dir="$2" rel="$3"
  [ -f "$src" ] || { echo "  ✗ Missing: $src"; return; }
  local fname
  fname=$(basename "$src" | tr ' ' '_')
  cp "$src" "$dest_dir/$fname"
  local esc
  esc=$(printf '%s\n' "$src" | sed 's/[\/&]/\\&/g')
  sed -i '' "s|$esc|$rel/$fname|g" "$HTML_OUT"
  echo "  ✓ $(basename "$src") → $rel/$fname"
}

echo "🎓 Certificates..."
for f in "/Users/cps/Downloads/Cer Kuma/"*.jpeg "/Users/cps/Downloads/Cer Kuma/"*.jpg; do
  [ -f "$f" ] && copy_rewrite "$f" "$CER_DIR" "assets/cer"
done

echo "🏆 Athlete photos..."
for f in "/Users/cps/Downloads/Kuma Coach/ảnh vđv/"*.jpg "/Users/cps/Downloads/Kuma Coach/ảnh vđv/"*.png; do
  [ -f "$f" ] && copy_rewrite "$f" "$ATHLETES_DIR" "assets/athletes"
done

echo "📚 Ebook covers..."
for f in \
  "/Users/cps/Downloads/Kuma Coach/The Model Elevate program. KUMATRAINING.png" \
  "/Users/cps/Downloads/Kuma Coach/A science-based guide to building muscle. KUMATRAINING.png" \
  "/Users/cps/Downloads/Kuma Coach/THE CLASSIC AESTHETICS PROGRAM. KUMATRAINING.png" \
  "/Users/cps/Downloads/Kuma Coach/THE BIKINI HEAVEN cover.png" \
  "/Users/cps/Downloads/Kuma Coach/THE WELLNESS BLUEPRINT cover.png" \
  "/Users/cps/Downloads/Kuma Coach/The Mens Physique Training Method.png"; do
  copy_rewrite "$f" "$EBOOKS_DIR" "assets/ebooks"
done

echo "🎨 Logos..."
for f in \
  "/Users/cps/Downloads/Kuma Coach/Logo Kumacoach 2.png" \
  "/Users/cps/Downloads/Kuma Coach/IFBB.png"; do
  copy_rewrite "$f" "$LOGOS_DIR" "assets/logos"
done

echo "👤 Coach & team photos..."
for f in \
  "/Users/cps/Downloads/Kuma Coach/Đạt kuma.png" \
  "/Users/cps/kumacoach/anh kuma/DSC00031.jpg" \
  "/Users/cps/kumacoach/anh kuma/DSC00133.jpg" \
  "/Users/cps/kumacoach/anh kuma/Đạt kuma.png" \
  "/Users/cps/Downloads/1389c2a2-8319-4f61-83c9-1bdb388918ed.jpeg" \
  "/Users/cps/Desktop/Ảnh màn hình 2026-06-25 lúc 13.38.19.png"; do
  copy_rewrite "$f" "$COACHES_DIR" "assets/coaches"
done

echo "🏋️  Training images (PNG → JPG compressed)..."
TRAINING_SRC="/Users/cps/Downloads/Kuma Coach/Bài Đăng Training Kuma"
rm -f "$TRAINING_DIR"/*.png "$TRAINING_DIR"/*.jpg 2>/dev/null
count=0
for i in $(seq 181 263); do
  src="$TRAINING_SRC/${i}.png"
  if [ -f "$src" ]; then
    sips -Z 540 -s format jpeg -s formatOptions 72 "$src" --out "$TRAINING_DIR/${i}.jpg" > /dev/null 2>&1
    count=$((count+1))
  fi
done
echo "  ✓ $count training images → assets/training/"

sed -i '' 's|/Users/cps/Downloads/Cer Kuma/tải-xuống.jpg|assets/cer/tải-xuống.jpg|g' "$HTML_OUT"

echo "🔗 Fixing JS gallery path..."
sed -i '' "s|const BASE = '/Users/cps/Downloads/Kuma Coach/Bài Đăng Training Kuma/';|const BASE = 'assets/training/';|g" "$HTML_OUT"
sed -i '' "s|const EXT = '.png';|const EXT = '.jpg';|g" "$HTML_OUT"

leftover=$(grep -c '/Users/cps/' "$HTML_OUT" 2>/dev/null || echo 0)
if [ "$leftover" -gt 0 ]; then
  echo "⚠️  $leftover unreplaced local paths remain:"
  grep -o '/Users/cps/[^"'\'']*' "$HTML_OUT" | sort -u | sed 's/^/    /'
else
  echo "✅ All paths resolved — no local references remain"
fi

echo ""
echo "📁 Asset structure:"
echo "   assets/cer/      — $(ls "$CER_DIR" | wc -l | tr -d ' ') certificates"
echo "   assets/athletes/ — $(ls "$ATHLETES_DIR" | wc -l | tr -d ' ') athlete photos"
echo "   assets/ebooks/   — $(ls "$EBOOKS_DIR" | wc -l | tr -d ' ') ebook covers"
echo "   assets/logos/    — $(ls "$LOGOS_DIR" | wc -l | tr -d ' ') logos"
echo "   assets/coaches/  — $(ls "$COACHES_DIR" | wc -l | tr -d ' ') coach photos"
echo "   assets/training/ — $(ls "$TRAINING_DIR" | wc -l | tr -d ' ') training images"

# ── GIT PUSH ──────────────────────────────────────────────
echo ""
echo "🚀 Pushing to GitHub..."
cd "$DEST_DIR"
git add .
if git diff --cached --quiet; then
  echo "✅ Nothing new to commit — already up to date."
else
  git commit -m "auto update $(date '+%Y-%m-%d %H:%M')"
  git push
  echo "✅ Pushed to GitHub! Cloudflare Pages will auto-deploy in ~1 min."
fi
