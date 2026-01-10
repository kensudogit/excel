#!/bin/bash

echo "========================================"
echo "Excel キーワード検索アプリケーション"
echo "========================================"
echo ""

# 仮想環境が存在するか確認
if [ ! -d "venv" ]; then
    echo "[エラー] 仮想環境が見つかりません。"
    echo "まず setup.sh を実行してセットアップしてください。"
    exit 1
fi

# node_modulesが存在するか確認
if [ ! -d "node_modules" ]; then
    echo "[エラー] node_modulesが見つかりません。"
    echo "まず setup.sh を実行してセットアップしてください。"
    exit 1
fi

echo "[起動中] バックエンドサーバーを起動しています..."
source venv/bin/activate
python app.py &
BACKEND_PID=$!

# バックエンドサーバーの起動を待つ
sleep 3

echo "[起動中] フロントエンドサーバーを起動しています..."
echo ""
echo "========================================"
echo "ブラウザで http://localhost:3001 にアクセスしてください"
echo "========================================"
echo ""
echo "サーバーを停止するには、Ctrl+C を押してください。"
echo ""

# フロントエンドサーバーを起動
npm run dev

# クリーンアップ
kill $BACKEND_PID 2>/dev/null
