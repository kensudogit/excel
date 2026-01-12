# Vercelデプロイ手順

## 前提条件

1. Vercelアカウントを作成: https://vercel.com
2. Vercel CLIをインストール: `npm i -g vercel`

## デプロイ手順

### 方法1: Vercel CLIを使用（推奨）

```bash
# プロジェクトディレクトリに移動
cd C:\devlop\excel

# Vercelにログイン
vercel login

# デプロイ（初回）
vercel

# 本番環境にデプロイ
vercel --prod
```

### 方法2: GitHub連携

1. GitHubリポジトリにプッシュ
2. Vercelダッシュボードで「New Project」をクリック
3. GitHubリポジトリを選択
4. プロジェクト設定:
   - Framework Preset: Other
   - Build Command: `npm run build`
   - Output Directory: `dist`
   - Install Command: `npm install`
5. 「Deploy」をクリック

## 環境変数

Vercelダッシュボードで以下の環境変数を設定（必要に応じて）:

- `FLASK_PORT`: ポート番号（通常は不要）
- `FLASK_DEBUG`: デバッグモード（`false`に設定）
- `DEFAULT_SEARCH_FOLDER`: デフォルト検索フォルダ（オプション）

## 注意事項

### 制限事項

1. **ファイルシステム**: VercelのServerless Functionsは読み取り専用のファイルシステムを使用します。一時ファイルは`/tmp`ディレクトリに保存されますが、永続化されません。

2. **ファイルアップロード**: アップロードされたファイルは一時的に`/tmp`ディレクトリに保存されますが、関数の実行が終了すると削除される可能性があります。

3. **実行時間**: Serverless Functionsの最大実行時間は10秒（Hobbyプラン）または60秒（Proプラン）です。大きなExcelファイルの処理には時間がかかる場合があります。

4. **メモリ制限**: 関数のメモリ制限は128MB（Hobbyプラン）または1024MB（Proプラン）です。

### 推奨事項

- 大きなExcelファイルを処理する場合は、ファイルサイズ制限を考慮してください
- 結果ファイルのダウンロードは、処理完了後すぐに行ってください（一時ファイルが削除される前に）
- 本番環境では`FLASK_DEBUG=false`に設定してください

## トラブルシューティング

### ビルドエラー

```bash
# ローカルでビルドをテスト
npm run build
```

### 関数エラー

Vercelダッシュボードの「Functions」タブでログを確認してください。

### CORSエラー

`vercel.json`の`headers`設定を確認してください。完全公開モードで`Access-Control-Allow-Origin: *`が設定されています。
