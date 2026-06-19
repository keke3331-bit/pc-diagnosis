#!/bin/bash
# src/ → docs/ に同期してコミット＆プッシュ（GitHub Pages公開用）
set -e
mkdir -p docs/js docs/css docs/img
cp src/index.html      docs/index.html
cp src/css/style.css   docs/css/style.css
cp src/js/blueprint.js docs/js/blueprint.js
cp src/js/firebase-config.js docs/js/firebase-config.js
cp src/js/app.js       docs/js/app.js
cp src/img/logo.png    docs/img/logo.png
echo "✅ docs/ 同期完了"

MSG="${1:-docs/ を src/ と同期}"
git add src/ docs/ 2>/dev/null || true
git commit -m "$MSG

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>" 2>/dev/null || echo "（変更なし、コミットスキップ）"
git push origin main 2>/dev/null || echo "（リモート未設定 or プッシュ不可）"
echo "🚀 完了"
