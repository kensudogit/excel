@echo off
echo ========================================
echo Excel キーワード検索アプリケーション セットアップ
echo ========================================
echo.

REM Pythonのバージョンチェック
python --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Pythonがインストールされていません。
    echo Python 3.8以上をインストールしてください。
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Python仮想環境を作成中...
if not exist venv (
    python -m venv venv
    if errorlevel 1 (
        echo [エラー] 仮想環境の作成に失敗しました。
        pause
        exit /b 1
    )
    echo [完了] 仮想環境を作成しました。
) else (
    echo [スキップ] 仮想環境は既に存在します。
)

echo.
echo [2/4] Python依存関係をインストール中...
call venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo [エラー] Python依存関係のインストールに失敗しました。
    pause
    exit /b 1
)
echo [完了] Python依存関係をインストールしました。

echo.
echo [3/4] Node.jsのバージョンチェック中...
node --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Node.jsがインストールされていません。
    echo Node.js 16以上をインストールしてください。
    echo https://nodejs.org/
    pause
    exit /b 1
)

echo.
echo [4/4] Node.js依存関係をインストール中...
call npm install
if errorlevel 1 (
    echo [エラー] Node.js依存関係のインストールに失敗しました。
    pause
    exit /b 1
)
echo [完了] Node.js依存関係をインストールしました。

echo.
echo ========================================
echo セットアップが完了しました！
echo ========================================
echo.
echo 起動方法:
echo   1. run.bat をダブルクリック
echo   または
echo   2. コマンドプロンプトで以下を実行:
echo      run.bat
echo.
pause
