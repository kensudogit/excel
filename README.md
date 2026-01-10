# Excel キーワード検索アプリケーション

指定したフォルダ内のExcelファイルから複数のキーワードを検索し、結果を別のブックに出力するアプリケーションです。

## 機能

- 📁 指定フォルダ内のExcelファイル（.xlsx, .xls）を自動検索
- 🔍 複数のキーワードを同時に検索
- 📊 検索結果をExcelブックに出力
- 📋 検索結果をクリックして該当セルの詳細情報を表示
- 🎨 キーワードごとに色分け表示

## 技術スタック

### バックエンド
- Python 3.8+
- Flask
- openpyxl（Excel操作）
- pandas

### フロントエンド
- React 18
- TypeScript
- Vite

## セットアップ

### 1. バックエンドのセットアップ

```bash
# 仮想環境の作成（推奨）
python -m venv venv

# 仮想環境の有効化
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 依存関係のインストール
pip install -r requirements.txt
```

### 2. フロントエンドのセットアップ

```bash
# 依存関係のインストール
npm install
```

## 実行方法

### 1. バックエンドサーバーの起動

```bash
python app.py
```

バックエンドサーバーは `http://localhost:5001` で起動します。

### 2. フロントエンド開発サーバーの起動

```bash
npm run dev
```

フロントエンドは `http://localhost:3001` で起動します。

## 使用方法

1. ブラウザで `http://localhost:3001` にアクセス
2. 「検索対象フォルダ」にExcelファイルが保存されているフォルダのパスを入力
3. 「検索キーワード」に検索したいキーワードを入力（複数可）
4. 「🔍 検索開始」ボタンをクリック
5. 検索結果がテーブルに表示されます
6. 検索結果の行をクリックすると、該当セルの詳細情報（周辺セルを含む）が表示されます
7. 「📥 結果をExcelでダウンロード」ボタンで検索結果をExcelファイルとしてダウンロードできます

## プロジェクト構造

```
excel/
├── app.py                 # Flaskバックエンド
├── requirements.txt       # Python依存関係
├── package.json          # Node.js依存関係
├── vite.config.ts        # Vite設定
├── tsconfig.json         # TypeScript設定
├── index.html            # HTMLエントリーポイント
├── src/
│   ├── main.tsx          # Reactエントリーポイント
│   ├── App.tsx           # メインアプリケーションコンポーネント
│   ├── App.css           # アプリケーションスタイル
│   ├── index.css         # グローバルスタイル
│   ├── types/
│   │   └── index.ts      # TypeScript型定義
│   └── components/
│       ├── SearchForm.tsx        # 検索フォーム
│       ├── SearchForm.css
│       ├── ResultsTable.tsx      # 検索結果テーブル
│       ├── ResultsTable.css
│       ├── CellDetails.tsx       # セル詳細モーダル
│       └── CellDetails.css
├── uploads/               # アップロードファイル（.gitignore）
└── results/              # 検索結果ファイル（.gitignore）
```

## API エンドポイント

### POST /api/search
Excelファイルを検索

**リクエスト:**
```json
{
  "folder_path": "C:\\Users\\Documents\\ExcelFiles",
  "keywords": ["キーワード1", "キーワード2"]
}
```

**レスポンス:**
```json
{
  "success": true,
  "results": [...],
  "total_matches": 10,
  "files_searched": 3,
  "output_file": "results/search_results_20240101_120000.xlsx"
}
```

### POST /api/get-cell-details
セルの詳細情報を取得

**リクエスト:**
```json
{
  "file_path": "C:\\Users\\Documents\\ExcelFiles\\file.xlsx",
  "sheet_name": "Sheet1",
  "row": 5,
  "col": 3,
  "keyword": "キーワード1",
  "context_rows": 5
}
```

### GET /api/download-results
検索結果ファイルをダウンロード

**クエリパラメータ:**
- `file_path`: ダウンロードするファイルのパス

## ライセンス

MIT
"# excel" 
