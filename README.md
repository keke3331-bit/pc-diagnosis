# パソコン診断処方箋 Web版

Excel（`☆(社内)…青白共用パソコン診断処方箋.xlsm`）をWeb化したもの。
**どのPCからでも作成・編集・印刷**でき、データは Firebase（クラウド）で共有されます。

## 構成
- `src/` … 開発用（編集はこちら）
- `docs/` … 公開用（GitHub Pages はこのフォルダを配信）。`deploy.sh` で `src/ → docs/` 同期
- `src/js/blueprint.js` … Excelから自動抽出した設計図（印刷レイアウト・数式・選択肢・初期値）。**手編集しない**
- `src/js/app.js` … アプリ本体（フォーム描画・数式計算・Firebase・印刷）
- `src/js/firebase-config.js` … FC作業進捗管理と同じ apo-dashboard プロジェクトを共用。保存先ノード = `pc_prescriptions`

## 使い方
1. 画面左の**入力フォーム**に記入（Excelの「入力」シートと同じ並び。赤=必須/緑=選択/青=入力）。
   - 「システム選択」を Windows / Mac に切り替えると、その下の項目と右の処方箋が自動で切り替わります。
   - CPUメーカーを変えると CPUシリーズの選択肢が連動します。
2. 画面右に**A4処方箋（「反映」シートと同じレイアウト）**がリアルタイム表示されます。
3. **💾保存** … クラウドに保存（どのPCからでも一覧に出ます）。
4. **📋一覧** … 保存済みを検索→「開く」で再編集 / 「複製」/「削除」。
5. **🖨印刷/PDF** … ブラウザの印刷からA4印刷、または「PDFとして保存」。

## どのPCでも使えるしくみ
データは Firebase Realtime Database（クラウド）に入るため、**HTMLを開く場所がどこでも同じ記録を共有**します。
そのため公開方法は2通り：
- **(推奨) GitHub Pages で公開** … `deploy.sh` で `docs/` をpush。URLを共有すれば各PCはブラウザで開くだけ。
- **ファイル配布** … `docs/index.html` を各PCで開くだけでもOK（記録はクラウド共有）。

### GitHub Pages 公開手順（任意）
```
cd ~/projects/パソコン診断処方箋
git init && git add . && git commit -m "init"
# GitHubで空リポジトリを作成し、リモート登録後:
git remote add origin <repoのURL>
git push -u origin main
# GitHubの Settings > Pages で「Branch: main / フォルダ: /docs」を選択
# 以後の更新は ./deploy.sh "メッセージ" で同期＆push
```

## Excelからの変更点（仕様）
- 「保存・印刷」マクロ → Web の「保存（クラウド）」「印刷/PDF」に置き換え。
- 「りすと」シートの選択肢・計算・Win/Mac分岐・ストレージ使用率バーは再現済み。
- ふりがな（ルビ）は本文のみ表示。スタッフ用のDDR規格参照表（別シート画像）は印刷対象外。
