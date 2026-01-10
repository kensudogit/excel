@echo off
title Excel キーワード検索アプリケーション

echo ========================================
echo Excel キーワード検索アプリケーション
echo ========================================
echo.

REM 仮想環境が存在するか確認
if not exist venv (
    echo [エラー] 仮想環境が見つかりません。
    echo まず setup.bat を実行してセットアップしてください。
    pause
    exit /b 1
)

REM node_modulesが存在するか確認
if not exist node_modules (
    echo [エラー] node_modulesが見つかりません。
    echo まず setup.bat を実行してセットアップしてください。
    pause
    exit /b 1
)

echo [起動中] バックエンドサーバーを起動しています...
start "Excel Search Backend" cmd /k "venv\Scripts\python.exe app.py"

REM バックエンドサーバーの起動を待つ
timeout /t 3 /nobreak >nul

echo [起動中] フロントエンドサーバーを起動しています...
echo.
echo ========================================
echo ブラウザで http://localhost:3001 にアクセスしてください
echo ========================================
echo.
echo サーバーを停止するには、このウィンドウを閉じてください。
echo.

call npm run dev
