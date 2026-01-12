# Vercelデプロイ手順

## クイックスタート

### 1. Vercel CLIのインストール

```bash
npm install -g vercel
```

### 2. ログイン

```bash
vercel login
```

### 3. デプロイ

```bash
cd C:\devlop\excel
vercel
```

初回デプロイ時は、いくつかの質問に答えます：
- Set up and deploy? **Y**
- Which scope? アカウントを選択
- Link to existing project? **N** (初回)
- Project name? `excel-keyword-search` (任意)
- Directory? `.` (現在のディレクトリ)
- Override settings? **N**

### 4. 本番環境にデプロイ

```bash
vercel --prod
```

## 作成されたファイル

- `vercel.json`: Vercelの設定ファイル
- `api/index.py`: Serverless Functionラッパー
- `.vercelignore`: デプロイから除外するファイル
- `.gitignore`: Gitから除外するファイル
- `README_VERCEL.md`: 詳細なドキュメント

## 重要な注意事項

⚠️ **VercelのServerless Functionsには以下の制限があります：**

1. **一時ファイル**: `/tmp`ディレクトリのみ書き込み可能（永続化されない）
2. **実行時間**: 最大10秒（Hobby）または60秒（Pro）
3. **メモリ**: 128MB（Hobby）または1024MB（Pro）
4. **ファイルシステム**: 読み取り専用（`/tmp`を除く）

## トラブルシューティング

### ビルドエラー

```bash
# ローカルでビルドをテスト
npm run build
```

### 関数が動作しない

Vercelダッシュボードの「Functions」タブでログを確認してください。

### CORSエラー

`vercel.json`の`headers`設定を確認してください。

## 環境変数の設定

Vercelダッシュボードで設定：
1. プロジェクトを選択
2. 「Settings」→「Environment Variables」
3. 以下を追加（必要に応じて）:
   - `FLASK_DEBUG`: `false`
   - `DEFAULT_SEARCH_FOLDER`: (オプション)
