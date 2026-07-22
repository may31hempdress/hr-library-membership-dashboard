# 自動更新セットアップ手順(毎週月曜21:00 JST)

GitHub Actions がサービスアカウント経由で Google スプレッドシートから直接データを
取得し、`public/data.json` を更新 → push → Vercel が自動デプロイする。
初回のみ以下のセットアップが必要(所要時間15分程度)。

すでに `採用図書館作戦会議/hr_library_crm` で GCP プロジェクト/サービスアカウントを
作っている場合は、手順1〜3は使い回してよい(手順4のスプレッドシート共有だけ追加すればよい)。

---

## 手順1: GCP プロジェクト作成

1. [Google Cloud Console](https://console.cloud.google.com/) を開く
2. 左上のプロジェクトセレクタ → **新しいプロジェクト**
3. プロジェクト名: `hr-library-dashboard`(任意。既存プロジェクトの流用も可)

## 手順2: API 有効化

1. 左メニュー → **API とサービス** → **ライブラリ**
2. `Google Drive API` を検索して **有効にする**
   (スプレッドシートを xlsx としてエクスポートするのに使う。Sheets API は不要)

## 手順3: サービスアカウント作成

1. 左メニュー → **IAM と管理** → **サービスアカウント** → **サービスアカウントを作成**
2. 名前: `dashboard-reader`(任意)
3. **作成して続行** → ロール付与はスキップ → **完了**
4. 作成されたサービスアカウントをクリック → **キー** タブ → **鍵を追加** →
   **新しい鍵を作成** → **JSON** を選択 → **作成**(JSON ファイルがダウンロードされる)
5. JSON を開き `"client_email": "xxxx@xxxx.iam.gserviceaccount.com"` をコピー
   (以後「SAメアド」と呼ぶ)

> **重要**: この JSON 鍵は秘密情報。Slack に貼ったり他人に共有したりしないこと。

## 手順4: スプレッドシートをサービスアカウントと共有

1. 対象スプレッドシートを開く:
   `https://docs.google.com/spreadsheets/d/1Dul08bK15WSTpRQhK1tYWiOQRulkuMz0lCycFpx9XRk/edit`
2. 右上の **共有** → 手順3の SAメアドを追加 → **閲覧者** 権限で共有

## 手順5: GitHub Secrets に登録

リポジトリ: `may31hempdress/hr-library-membership-dashboard`

1. GitHub の当該リポジトリ → **Settings** → **Secrets and variables** → **Actions**
2. **New repository secret** を2つ作成:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` — 手順3でダウンロードした JSON ファイルの中身を丸ごと貼り付け
   - `SPREADSHEET_ID` — `1Dul08bK15WSTpRQhK1tYWiOQRulkuMz0lCycFpx9XRk`

ターミナルから `gh` CLI で登録する場合(JSON鍵をダウンロードしたローカル環境で):

```bash
gh secret set GOOGLE_SERVICE_ACCOUNT_JSON --repo may31hempdress/hr-library-membership-dashboard < path/to/service_account.json
gh secret set SPREADSHEET_ID --repo may31hempdress/hr-library-membership-dashboard --body "1Dul08bK15WSTpRQhK1tYWiOQRulkuMz0lCycFpx9XRk"
```

## 手順6: Actions の書き込み権限を確認

1. GitHub リポジトリ → **Settings** → **Actions** → **General**
2. **Workflow permissions** で **Read and write permissions** が選択されていることを確認
   (これが Read-only になっていると `data.json` の自動 push が失敗する)

## 手順7: 動作確認

1. GitHub リポジトリ → **Actions** タブ → **Weekly data update** を選択
2. **Run workflow** で手動実行
3. 成功すれば `public/data.json` を更新するコミットが自動で積まれる
4. Vercel と GitHub が連携済みなら、そのまま本番に自動デプロイされる

以後は毎週月曜 21:00 JST に自動実行される。

---

## トラブルシュート

- `PermissionError` / 403 → 手順4の共有漏れ、または SAメアドの入力ミス
- `has not been used in project` / API disabled エラー → 手順2の Drive API 有効化漏れ
- push が失敗する(403) → 手順6のワークフロー権限を確認
- 手動でローカル実行してテストしたい場合:
  ```bash
  export GOOGLE_SERVICE_ACCOUNT_JSON=secrets/service_account.json
  export SPREADSHEET_ID=1Dul08bK15WSTpRQhK1tYWiOQRulkuMz0lCycFpx9XRk
  pip install -r requirements.txt
  python fetch_from_gsheet.py
  ```
  (`secrets/service_account.json` は `.gitignore` 済みなのでリポジトリには含まれない)
