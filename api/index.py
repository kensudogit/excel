"""
Vercel Serverless Function wrapper for Flask app
"""
import sys
import os
from pathlib import Path

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
    from flask import Response as FlaskResponse
    
    # リクエスト情報を取得
    method = request.get('method', 'GET')
    path = request.get('path', '/')
    query_string = request.get('queryStringParameters', {}) or {}
    headers = request.get('headers', {}) or {}
    body = request.get('body', '')
    
    # クエリ文字列を構築
    query_parts = []
    for key, value in query_string.items():
        query_parts.append(f"{key}={value}")
    query_string_str = '&'.join(query_parts)
    
    # ボディをバイトに変換
    if isinstance(body, str):
        body_bytes = body.encode('utf-8')
    else:
        body_bytes = body
    
    # Flaskのテストクライアントを使用
    with app.test_request_context(
        path=path,
        method=method,
        query_string=query_string_str,
        headers=headers,
        data=body_bytes
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
                response_body_str = response_body.decode('utf-8')
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
