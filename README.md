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

## 必要な環境

- **Python 3.8以上** - [ダウンロード](https://www.python.org/downloads/)
- **Node.js 16以上** - [ダウンロード](https://nodejs.org/)
- **npm** - Node.jsに含まれています

## クイックスタート（推奨）

### Windows

1. **セットアップ**
   ```bash
   setup.bat
   ```
   または、`setup.bat`をダブルクリック

2. **起動**
   ```bash
   run.bat
   ```
   または、`run.bat`をダブルクリック

3. **ブラウザでアクセス**
   - http://localhost:3001 にアクセス

### Linux/Mac

1. **セットアップ**
   ```bash
   chmod +x setup.sh run.sh
   ./setup.sh
   ```

2. **起動**
   ```bash
   ./run.sh
   ```

3. **ブラウザでアクセス**
   - http://localhost:3001 にアクセス

## 手動セットアップ

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

### 方法1: 自動起動スクリプト（推奨）

**Windows:**
```bash
run.bat
```

**Linux/Mac:**
```bash
./run.sh
```

### 方法2: 手動起動

#### 1. バックエンドサーバーの起動

```bash
# Windows
venv\Scripts\python.exe app.py

# Linux/Mac
source venv/bin/activate
python app.py
```

バックエンドサーバーは `http://localhost:5001` で起動します。

#### 2. フロントエンド開発サーバーの起動

別のターミナルで：

```bash
npm run dev
```

フロントエンドは `http://localhost:3001` で起動します。

## 使用方法

1. **ブラウザでアクセス**
   - `http://localhost:3001` にアクセス

2. **フォルダの選択**
   - 「📁 フォルダ選択」ボタンをクリックしてフォルダを選択
   - または、フォルダ/Excelファイルをドラッグ&ドロップ
   - または、フォルダパスを直接入力

3. **キーワードの入力**
   - 「検索キーワード」に検索したいキーワードを入力（複数可）
   - 「+ キーワードを追加」ボタンでキーワードを追加

4. **検索実行**
   - 「🔍 検索開始」ボタンをクリック
   - 検索中はカーソルが「wait」に変わり、ローディングスピナーが表示されます

5. **結果の確認**
   - 検索結果がテーブルに表示されます（20件ずつページネーション）
   - 検索結果の「📋 詳細」ボタンをクリックすると、該当セルの詳細情報（周辺セルを含む）が表示されます
   - キーワードをクリックすると、該当のExcelファイルが開きます

6. **結果のダウンロード**
   - 「📥 結果をExcelでダウンロード」ボタンで検索結果をExcelファイルとしてダウンロードできます

## プロジェクト構造

```
excel/
├── app.py                 # Flaskバックエンド
├── requirements.txt       # Python依存関係
├── package.json          # Node.js依存関係
├── vite.config.ts        # Vite設定
├── tsconfig.json         # TypeScript設定
├── index.html            # HTMLエントリーポイント
├── setup.bat             # Windowsセットアップスクリプト
├── setup.sh              # Linux/Macセットアップスクリプト
├── run.bat               # Windows起動スクリプト
├── run.sh                # Linux/Mac起動スクリプト
├── build.bat             # ビルドスクリプト
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
│       ├── ResultsTable.tsx      # 検索結果テーブル（ページネーション対応）
│       ├── ResultsTable.css
│       ├── CellDetails.tsx      # セル詳細モーダル
│       └── CellDetails.css
├── uploads/               # アップロードファイル（自動生成）
└── results/              # 検索結果ファイル（自動生成）
```

## 他のPCへのコピー方法

1. **プロジェクトフォルダ全体をコピー**
   - `excel`フォルダ全体をUSBメモリやネットワーク経由でコピー

2. **新しいPCでセットアップ**
   - コピーしたフォルダに移動
   - `setup.bat`（Windows）または`setup.sh`（Linux/Mac）を実行

3. **起動**
   - `run.bat`（Windows）または`run.sh`（Linux/Mac）を実行

**注意:** 以下のフォルダ/ファイルはコピー不要です（自動生成されます）：
- `venv/` - Python仮想環境
- `node_modules/` - Node.js依存関係
- `uploads/` - アップロードファイル
- `results/` - 検索結果ファイル
- `*.log` - ログファイル

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

## トラブルシューティング

### ポートが既に使用されている場合

- **バックエンド（ポート5001）**: `app.py`の最後の行でポート番号を変更
- **フロントエンド（ポート3001）**: `vite.config.ts`の`server.port`を変更

### 仮想環境が作成できない場合

- Pythonが正しくインストールされているか確認
- `python --version`でバージョンを確認（3.8以上が必要）

### 依存関係のインストールに失敗する場合

- インターネット接続を確認
- プロキシ設定が必要な場合は、環境変数を設定
- Windowsの場合、管理者権限で実行してみる

### バックエンドサーバーが起動しない場合

- ポート5001が使用されていないか確認
- ファイアウォールの設定を確認
- `venv\Scripts\python.exe app.py`を直接実行してエラーメッセージを確認

## ビルド（本番環境用）

本番環境で使用する場合は、フロントエンドをビルドしてください：

```bash
# Windows
build.bat

# Linux/Mac
npm run build
```

ビルドされたファイルは`dist`フォルダに生成されます。

## ライセンス

MIT 
