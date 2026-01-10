@echo off
echo ========================================
echo Excel キーワード検索アプリケーション ビルド
echo ========================================
echo.

REM node_modulesが存在するか確認
if not exist node_modules (
    echo [エラー] node_modulesが見つかりません。
    echo まず setup.bat を実行してセットアップしてください。
    pause
    exit /b 1
)

echo [ビルド中] フロントエンドをビルドしています...
call npm run build
if errorlevel 1 (
    echo [エラー] ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo [完了] ビルドが完了しました。
echo ビルドされたファイルは dist フォルダにあります。
echo.
pause
