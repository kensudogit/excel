"""
Vercel Serverless Function wrapper for Flask app
"""
import sys
import os
from pathlib import Path
import json

# Vercel環境であることを示す環境変数を設定
os.environ['VERCEL'] = '1'

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 一時ディレクトリを設定（Vercelの/tmpディレクトリを使用）
TMP_DIR = Path('/tmp')
UPLOAD_DIR = TMP_DIR / 'uploads'
RESULTS_DIR = TMP_DIR / 'results'

# ディレクトリを作成
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# 環境変数で一時ディレクトリを設定
os.environ['UPLOAD_FOLDER'] = str(UPLOAD_DIR)
os.environ['RESULTS_FOLDER'] = str(RESULTS_DIR)

# Flaskアプリをインポート
from app import app

def handler(request):
    """
    Vercel Serverless Function handler
    VercelのPython runtimeは、リクエストオブジェクトを受け取り、レスポンス辞書を返す
    """
    try:
        # リクエスト情報を取得
        # Vercelのリクエスト形式に合わせる
        if isinstance(request, dict):
            method = request.get('method', 'GET')
            path = request.get('path', '/')
            # パスから/api/プレフィックスを削除
            if path.startswith('/api/'):
                path = path[4:]  # '/api/'を削除
            query_string = request.get('queryStringParameters', {}) or {}
            headers = request.get('headers', {}) or {}
            body = request.get('body', '')
        else:
            # リクエストオブジェクトから属性を取得
            method = getattr(request, 'method', 'GET')
            path = getattr(request, 'path', '/')
            if path.startswith('/api/'):
                path = path[4:]
            query_string = getattr(request, 'queryStringParameters', {}) or {}
            headers = getattr(request, 'headers', {}) or {}
            body = getattr(request, 'body', '')
        
        # クエリ文字列を構築
        query_parts = []
        if query_string:
            for key, value in query_string.items():
                if value:
                    query_parts.append(f"{key}={value}")
        query_string_str = '&'.join(query_parts)
        
        # ボディをバイトに変換
        if isinstance(body, str):
            body_bytes = body.encode('utf-8')
        elif body is None:
            body_bytes = b''
        else:
            body_bytes = body
        
        # Content-Typeヘッダーを確認
        content_type = headers.get('content-type', '').lower() if isinstance(headers, dict) else ''
        
        # Flaskのテストクライアントを使用
        with app.test_request_context(
            path=path,
            method=method,
            query_string=query_string_str,
            headers=headers,
            data=body_bytes,
            content_type=content_type if content_type else None
        ):
            response = app.full_dispatch_response()
            
            # レスポンスヘッダーを構築
            response_headers = {}
            for key, value in response.headers:
                response_headers[key] = value
            
            # CORSヘッダーを追加（完全公開モード）
            response_headers['Access-Control-Allow-Origin'] = '*'
            response_headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'
            response_headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
            
            # レスポンスボディを取得
            response_body = response.get_data(as_text=False)
            if response_body:
                try:
                    # バイナリデータの場合はbase64エンコード
                    if isinstance(response_body, bytes):
                        # JSONレスポンスの場合は文字列として返す
                        try:
                            response_body_str = response_body.decode('utf-8')
                        except UnicodeDecodeError:
                            # バイナリデータの場合はbase64エンコード
                            import base64
                            response_body_str = base64.b64encode(response_body).decode('utf-8')
                            response_headers['Content-Encoding'] = 'base64'
                    else:
                        response_body_str = str(response_body)
                except:
                    response_body_str = str(response_body)
            else:
                response_body_str = ''
            
            # VercelのResponse形式で返す
            return {
                'statusCode': response.status_code,
                'headers': response_headers,
                'body': response_body_str
            }
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Error in handler: {error_trace}")
        return {
            'statusCode': 500,
            'headers': {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*'
            },
            'body': json.dumps({
                'success': False,
                'error': str(e),
                'traceback': error_trace
            })
        }
