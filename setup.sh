#!/bin/bash

echo "========================================"
echo "Excel キーワード検索アプリケーション セットアップ"
echo "========================================"
echo ""

# Pythonのバージョンチェック
if ! command -v python3 &> /dev/null; then
    echo "[エラー] Python3がインストールされていません。"
    echo "Python 3.8以上をインストールしてください。"
    echo "https://www.python.org/downloads/"
    exit 1
fi

echo "[1/4] Python仮想環境を作成中..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "[エラー] 仮想環境の作成に失敗しました。"
        exit 1
    fi
    echo "[完了] 仮想環境を作成しました。"
else
    echo "[スキップ] 仮想環境は既に存在します。"
fi

echo ""
echo "[2/4] Python依存関係をインストール中..."
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[エラー] Python依存関係のインストールに失敗しました。"
    exit 1
fi
echo "[完了] Python依存関係をインストールしました。"

echo ""
echo "[3/4] Node.jsのバージョンチェック中..."
if ! command -v node &> /dev/null; then
    echo "[エラー] Node.jsがインストールされていません。"
    echo "Node.js 16以上をインストールしてください。"
    echo "https://nodejs.org/"
    exit 1
fi

echo ""
echo "[4/4] Node.js依存関係をインストール中..."
npm install
if [ $? -ne 0 ]; then
    echo "[エラー] Node.js依存関係のインストールに失敗しました。"
    exit 1
fi
echo "[完了] Node.js依存関係をインストールしました。"

echo ""
echo "========================================"
echo "セットアップが完了しました！"
echo "========================================"
echo ""
echo "起動方法:"
echo "  1. ./run.sh を実行"
echo "  または"
echo "  2. コマンドで以下を実行:"
echo "     bash run.sh"
echo ""
