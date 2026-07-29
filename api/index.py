# -*- coding: utf-8 -*-
"""
Vercel Serverless Function — Flask WSGI 适配器
"""

import sys
import os

# 将项目根目录加入 Python 路径
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# 标记为 Vercel 环境
os.environ['NETLIFY'] = 'true'
os.environ['VERCEL'] = 'true'

from app import app as flask_app
from flask import Request


def handler(request):
    """Vercel Python 入口点"""
    from io import BytesIO

    # 读取请求体
    body = request.body or b''
    if hasattr(body, 'read'):
        body = body.read()

    # 构建 WSGI environ
    environ = {
        'REQUEST_METHOD': request.method,
        'SCRIPT_NAME': '',
        'PATH_INFO': request.path or '/',
        'QUERY_STRING': request.query_string or '',
        'SERVER_NAME': 'vercel',
        'SERVER_PORT': '443',
        'SERVER_PROTOCOL': 'HTTP/1.1',
        'wsgi.version': (1, 0),
        'wsgi.url_scheme': 'https',
        'wsgi.input': BytesIO(body if isinstance(body, bytes) else body.encode()),
        'wsgi.errors': BytesIO(),
        'wsgi.multithread': False,
        'wsgi.multiprocess': False,
        'wsgi.run_once': True,
        'CONTENT_LENGTH': str(len(body) if body else 0),
    }

    # 添加请求头
    for key, value in request.headers.items():
        wsgi_key = 'HTTP_' + key.upper().replace('-', '_')
        environ[wsgi_key] = value

    if request.headers.get('content-type'):
        environ['CONTENT_TYPE'] = request.headers['content-type']

    # 调用 Flask
    response_body = []
    response_headers = {}
    status_code = 500

    def start_response(status, headers_list, exc_info=None):
        nonlocal status_code, response_headers
        status_code = int(status.split()[0])
        for h_name, h_value in headers_list:
            response_headers[h_name] = h_value

    try:
        body_iter = flask_app(environ, start_response)
        for chunk in body_iter:
            if isinstance(chunk, str):
                chunk = chunk.encode('utf-8')
            response_body.append(chunk)
    except Exception as e:
        import json
        error_body = json.dumps({'error': str(e)}, ensure_ascii=False).encode('utf-8')
        response_body = [error_body]
        status_code = 500
        response_headers = {'Content-Type': 'application/json'}

    # 返回响应
    from flask import Response as FlaskResponse
    return FlaskResponse(
        b''.join(response_body),
        status=status_code,
        headers=response_headers,
    )
