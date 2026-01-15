"""
Vercel Serverless Function wrapper for Flask app

このファイルは、VercelのServerless Functions環境でFlaskアプリケーションを
実行するためのラッパーです。VercelのPython runtimeは、WSGIアプリケーションを
直接実行することはできないため、このラッパーを使用してFlaskアプリケーションを
Serverless Functionとして実行します。

主な機能:
- Vercelのリクエスト形式をFlaskのリクエスト形式に変換
- FlaskのレスポンスをVercelのレスポンス形式に変換
- 一時ディレクトリの設定（Vercelの/tmpディレクトリを使用）
- エラーハンドリングとログ出力

技術的な詳細:
- VercelのPython runtimeは、`handler`関数を探して実行します
- `handler`関数は、リクエスト辞書を受け取り、レスポンス辞書を返します
- Flaskアプリケーションは、`test_request_context`を使用して実行されます
"""
import sys  # システム固有のパラメータと関数
import os  # オペレーティングシステム関連の機能
from pathlib import Path  # パス操作のためのクラス
import json  # JSONデータの処理

# ============================================================================
# 環境変数の設定
# ============================================================================

# Vercel環境であることを示す環境変数を設定
# これにより、app.pyでVercel環境かどうかを判定できる
os.environ['VERCEL'] = '1'

# ============================================================================
# パスの設定
# ============================================================================

# プロジェクトルートをパスに追加
# api/index.pyの親ディレクトリ（プロジェクトルート）をPythonのパスに追加
# これにより、app.pyをインポートできるようになる
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# ============================================================================
# 一時ディレクトリの設定
# ============================================================================

# VercelのServerless Functionsでは、/tmpディレクトリのみが書き込み可能
# そのため、アップロードファイルや結果ファイルは/tmpディレクトリに保存する
TMP_DIR = Path('/tmp')
UPLOAD_DIR = TMP_DIR / 'uploads'  # アップロードされたファイルの一時保存先
RESULTS_DIR = TMP_DIR / 'results'  # 検索結果のExcelファイルの保存先

# ディレクトリを作成（存在しない場合）
# parents=True: 親ディレクトリも含めて作成
# exist_ok=True: 既に存在する場合はエラーを出さない
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# 環境変数で一時ディレクトリを設定
# app.pyでこれらの環境変数を読み取って、一時ディレクトリを使用する
os.environ['UPLOAD_FOLDER'] = str(UPLOAD_DIR)
os.environ['RESULTS_FOLDER'] = str(RESULTS_DIR)

# ============================================================================
# Flaskアプリケーションのインポート
# ============================================================================

# Flaskアプリをインポート
# VercelのPython runtimeがFlaskアプリを正しく認識できるようにする
# エラーが発生した場合は、詳細なエラー情報をログに出力してから再発生させる
try:
    from app import app
except Exception as e:
    import traceback
    error_trace = traceback.format_exc()
    import logging
    logging.error(f"Error importing Flask app: {error_trace}")
    raise

# ============================================================================
# Vercel Serverless Function Handler
# ============================================================================

def handler(request):
    """
    Vercel Serverless Function handler
    
    この関数は、VercelのPython runtimeから呼び出されます。
    Vercelのリクエスト形式をFlaskのリクエスト形式に変換し、
    Flaskアプリケーションを実行して、結果をVercelのレスポンス形式に変換します。
    
    引数:
        request: Vercelのリクエストオブジェクト（辞書形式またはオブジェクト形式）
            - method: HTTPメソッド（GET, POST, PUT, DELETEなど）
            - path: リクエストパス（例: /api/search）
            - queryStringParameters: クエリパラメータ（辞書形式）
            - headers: HTTPヘッダー（辞書形式）
            - body: リクエストボディ（文字列またはバイト）
    
    戻り値:
        dict: Vercelのレスポンス形式
            - statusCode: HTTPステータスコード（例: 200, 404, 500）
            - headers: HTTPレスポンスヘッダー（辞書形式）
            - body: レスポンスボディ（文字列）
    
    処理の流れ:
        1. リクエスト情報を取得（メソッド、パス、クエリパラメータ、ヘッダー、ボディ）
        2. パスから/api/プレフィックスを削除（Flaskアプリのルーティングに合わせる）
        3. クエリ文字列を構築
        4. ボディをバイト形式に変換
        5. Flaskのtest_request_contextを使用してリクエストを実行
        6. レスポンスを取得して、Vercelの形式に変換
        7. CORSヘッダーを追加（完全公開モード）
        8. エラーが発生した場合は、エラー情報を含むレスポンスを返す
    """
    try:
        # ====================================================================
        # リクエスト情報の取得
        # ====================================================================
        
        # リクエスト情報を取得
        # Vercelのリクエスト形式に合わせる
        # リクエストは辞書形式またはオブジェクト形式で渡される可能性がある
        if isinstance(request, dict):
            # 辞書形式のリクエスト
            method = request.get('method', 'GET')  # HTTPメソッド（デフォルト: GET）
            path = request.get('path', '/')  # リクエストパス（デフォルト: /）
            query_string = request.get('queryStringParameters', {}) or {}  # クエリパラメータ
            headers = request.get('headers', {}) or {}  # HTTPヘッダー
            body = request.get('body', '')  # リクエストボディ
        else:
            # オブジェクト形式のリクエスト
            method = getattr(request, 'method', 'GET')
            path = getattr(request, 'path', '/')
            query_string = getattr(request, 'queryStringParameters', {}) or {}
            headers = getattr(request, 'headers', {}) or {}
            body = getattr(request, 'body', '')
        
        # ====================================================================
        # パスの処理
        # ====================================================================
        
        # パスから/api/プレフィックスを削除
        # Vercelのrewrites設定により、/api/*のリクエストが/api/index.pyにルーティングされる
        # しかし、Flaskアプリのルーティングでは/api/プレフィックスは不要
        # そのため、/api/プレフィックスを削除してFlaskアプリに渡す
        if path.startswith('/api/'):
            path = path[4:]  # '/api/'を削除（例: /api/search -> /search）
        elif path.startswith('api/'):
            path = path[4:]  # 'api/'を削除（例: api/search -> /search）
        
        # パスが空の場合はルートに設定
        # パスが空文字列の場合は、ルートパス（/）に設定
        if not path or path == '':
            path = '/'
        
        # ====================================================================
        # クエリ文字列の構築
        # ====================================================================
        
        # クエリ文字列を構築
        # クエリパラメータをkey=value形式の文字列に変換
        query_parts = []
        if query_string:
            for key, value in query_string.items():
                if value:  # 値が存在する場合のみ追加
                    query_parts.append(f"{key}={value}")
        query_string_str = '&'.join(query_parts)  # クエリパラメータを&で結合
        
        # ====================================================================
        # リクエストボディの処理
        # ====================================================================
        
        # ボディをバイトに変換
        # Flaskのtest_request_contextはバイト形式のボディを期待する
        # VercelのServerless Functionsでは、リクエストボディは文字列またはbase64エンコードされた文字列として渡される可能性がある
        if isinstance(body, str):
            # base64エンコードされている可能性をチェック
            try:
                import base64
                # base64デコードを試みる（失敗した場合は通常の文字列として扱う）
                body_bytes = base64.b64decode(body)
            except:
                body_bytes = body.encode('utf-8')  # 文字列をUTF-8エンコード
        elif body is None:
            body_bytes = b''  # Noneの場合は空のバイト列
        else:
            body_bytes = body  # 既にバイト形式の場合はそのまま使用
        
        # ====================================================================
        # Content-Typeヘッダーの確認
        # ====================================================================
        
        # Content-Typeヘッダーを確認
        # multipart/form-dataの場合は、Flaskがファイルアップロードを処理できるようにする
        content_type = headers.get('content-type', '').lower() if isinstance(headers, dict) else ''
        
        # multipart/form-dataの場合は、リクエストオブジェクトから直接取得
        # VercelのServerless Functionsでは、multipart/form-dataは特殊な形式で渡される
        # しかし、Flaskのtest_request_contextでは、通常の形式で処理できる
        request_data = body_bytes
        request_content_type = content_type
        
        # ====================================================================
        # Flaskアプリケーションの実行
        # ====================================================================
        
        # Flaskのテストクライアントを使用
        # test_request_contextを使用して、Flaskアプリケーションを実行する
        # これにより、実際のHTTPリクエストをシミュレートできる
        # 注意: multipart/form-dataの場合は、Vercelが既にパースしている可能性があるため、
        # リクエストオブジェクトから直接取得する必要がある場合がある
        try:
            with app.test_request_context(
                path=path,  # リクエストパス
                method=method,  # HTTPメソッド
                query_string=query_string_str,  # クエリ文字列
                headers=headers,  # HTTPヘッダー
                data=request_data,  # リクエストボディ
                content_type=request_content_type if request_content_type else None  # Content-Type
            ):
            # Flaskアプリケーションを実行してレスポンスを取得
            # full_dispatch_response()は、リクエストを処理してレスポンスを返す
            response = app.full_dispatch_response()
            
            # ================================================================
            # レスポンスヘッダーの構築
            # ================================================================
            
            # レスポンスヘッダーを構築
            # Flaskのレスポンスヘッダーを辞書形式に変換
            response_headers = {}
            for key, value in response.headers:
                response_headers[key] = value
            
            # CORSヘッダーを追加（完全公開モード）
            # これにより、すべてのオリジンからのリクエストを許可する
            response_headers['Access-Control-Allow-Origin'] = '*'  # すべてのオリジンを許可
            response_headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'  # 許可するHTTPメソッド
            response_headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'  # 許可するHTTPヘッダー
            
            # ================================================================
            # レスポンスボディの処理
            # ================================================================
            
            # レスポンスボディを取得
            # as_text=Falseでバイト形式で取得（バイナリデータに対応）
            response_body = response.get_data(as_text=False)
            if response_body:
                try:
                    # バイナリデータの場合はbase64エンコード
                    if isinstance(response_body, bytes):
                        # JSONレスポンスの場合は文字列として返す
                        try:
                            response_body_str = response_body.decode('utf-8')  # UTF-8でデコード
                        except UnicodeDecodeError:
                            # バイナリデータ（画像、Excelファイルなど）の場合はbase64エンコード
                            import base64
                            response_body_str = base64.b64encode(response_body).decode('utf-8')
                            response_headers['Content-Encoding'] = 'base64'  # base64エンコードであることを示す
                    else:
                        response_body_str = str(response_body)  # 文字列形式の場合はそのまま使用
                except:
                    response_body_str = str(response_body)  # エラーが発生した場合は文字列に変換
            else:
                response_body_str = ''  # ボディが空の場合は空文字列
            
            # ================================================================
            # Vercelのレスポンス形式で返す
            # ================================================================
            
            # VercelのResponse形式で返す
            # VercelのServerless Functionsは、この形式の辞書を期待する
            return {
                'statusCode': response.status_code,  # HTTPステータスコード（200, 404, 500など）
                'headers': response_headers,  # HTTPレスポンスヘッダー
                'body': response_body_str  # レスポンスボディ（文字列形式）
            }
        except Exception as context_error:
            # test_request_context内でエラーが発生した場合
            import traceback
            error_trace = traceback.format_exc()
            import logging
            logging.error(f"Error in Flask request context: {error_trace}")
            logging.error(f"Path: {path}, Method: {method}, Content-Type: {content_type}")
            
            # エラーレスポンスを返す
            return {
                'statusCode': 500,
                'headers': {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': '*'
                },
                'body': json.dumps({
                    'success': False,
                    'error': f'Flask request context error: {str(context_error)}',
                    'error_type': type(context_error).__name__,
                    'traceback': error_trace
                }, ensure_ascii=False)
            }
            
    except Exception as e:
        # ====================================================================
        # エラーハンドリング
        # ====================================================================
        
        import traceback
        error_trace = traceback.format_exc()
        
        # エラーをログに出力（Vercelのログで確認可能）
        # Vercelダッシュボードの「Runtime Logs」で確認できる
        import logging
        logging.error(f"Error in handler: {error_trace}")
        logging.error(f"Request: {request}")
        
        # エラーレスポンスを返す
        # エラーが発生した場合でも、適切な形式でレスポンスを返す
        return {
            'statusCode': 500,  # 内部サーバーエラー
            'headers': {
                'Content-Type': 'application/json',  # JSON形式のレスポンス
                'Access-Control-Allow-Origin': '*'  # CORSヘッダー
            },
            'body': json.dumps({
                'success': False,  # エラーが発生したことを示す
                'error': str(e),  # エラーメッセージ
                'error_type': type(e).__name__,  # エラーの種類
                'traceback': error_trace  # スタックトレース（デバッグ用）
            }, ensure_ascii=False)  # 日本語文字を正しく表示するため、ensure_ascii=Falseを指定
        }
