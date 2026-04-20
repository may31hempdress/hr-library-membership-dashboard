# 人事図書館 メンバー推移ダッシュボード

週次・月次のメンバー入退会・退会率を可視化する静的 Vercel サイト。

## 構成

- `public/index.html` — ダッシュボード本体(Chart.js / static)
- `public/data.json` — 可視化対象データ(`extract.py` で生成)
- `extract.py` — Excel → JSON 変換スクリプト
- `vercel.json` — `data.json` はキャッシュ無効

## 週次更新フロー

1. 最新の `数字で見る人事図書館変遷_*.xlsx` を `../参考資料/` に置く
2. `python extract.py` を実行 → `public/data.json` が更新される
3. `vercel --prod` で再デプロイ

## 退会率の定義

`当期退会者数 ÷ 期首(前期末)アクティブメンバー数 × 100`

種別(フルアクセス / オンライン / 法人一括 / 合計)別に算出。
