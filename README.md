# 人事図書館 メンバー推移ダッシュボード

週次・月次のメンバー入退会・退会率を可視化する静的 Vercel サイト。

## 構成

- `public/index.html` — ダッシュボード本体(Chart.js / static)
- `public/data.json` — 可視化対象データ(`extract.py` で生成)
- `extract.py` — Excel → JSON 変換スクリプト
- `vercel.json` — `data.json` はキャッシュ無効

## 週次更新フロー

### 自動(毎週月曜 21:00 JST)

GitHub Actions (`.github/workflows/weekly-update.yml`) がサービスアカウント経由で
Google スプレッドシート(週次グラフ用 / 月次グラフ用シート)から直接データを取得し、
`public/data.json` を更新して自動 push する。push されると Vercel の Git 連携で
本番に自動反映される。

初回セットアップは [SETUP_AUTO_UPDATE.md](./SETUP_AUTO_UPDATE.md) を参照。

### 手動更新(スポット反映したい時)

1. 対象の Google スプレッドシートを Google Drive から xlsx としてダウンロード
   (または `数字で見る人事図書館変遷_*.xlsx` を直接用意)
2. `python extract.py <xlsxファイルパス>` を実行 → `public/data.json` が更新される
3. `git add public/data.json && git commit && git push`(Vercel が自動デプロイ)

## 退会率の定義

`当期退会者数 ÷ 期首(前期末)アクティブメンバー数 × 100`

種別(フルアクセス / オンライン / 法人一括 / 合計)別に算出。
