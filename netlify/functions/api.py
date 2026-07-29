#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Netlify Serverless Function — Flask WSGI 适配器
将所有 HTTP 请求转发到 Flask 应用处理

支持：登录认证、Supabase 持久化、PDF/Excel/PPT 解析、AI 出题
"""

import sys
import os
import json
import base64
from io import BytesIO
from urllib.parse import urlencode

# 将项目根目录加入 Python 路径
# __file__ = netlify/functions/api.py → 上三级 = 项目根目录
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

# Netlify 部署时设置环境标记（app.py 会根据此标记切换 /tmp/ 路径）
os.environ['NETLIFY'] = 'true'

from app import app as flask_app


def handler(event, context):
    """
    Netlify Functions 入口点。
    将 API Gateway 格式的 event 转为 WSGI environ，交给 Flask 处理。
    """
    # ---- 解析请求 ----
    http_method = (event.get('httpMethod') or 'GET').upper()
    path = event.get('path') or '/'

    # 去掉 Netlify 函数路径前缀
    prefix = '/.netlify/functions/api'
    if path.startswith(prefix):
        path = path[len(prefix):]
    if not path:
        path = '/'

    # 查询参数
    query_params = event.get('queryStringParameters') or {}
    query_string = urlencode(query_params) if query_params else ''

    # 请求头
    headers = event.get('headers') or {}

    # 请求体
    raw_body = event.get('body') or ''
    is_base64 = event.get('isBase64Encoded', False)

    if is_base64:
        try:
            raw_body = base64.b64decode(raw_body)
        except Exception:
            raw_body = b''
    elif isinstance(raw_body, str):
        raw_body = raw_body.encode('utf-8')
    else:
        raw_body = b''

    # ---- 构建 WSGI environ ----
    environ = {
        'REQUEST_METHOD': http_method,
        'SCRIPT_NAME': '',
        'PATH_INFO': path,
        'QUERY_STRING': query_string,
        'SERVER_NAME': 'netlify',
        'SERVER_PORT': '443',
        'SERVER_PROTOCOL': 'HTTP/1.1',
        'wsgi.version': (1, 0),
        'wsgi.url_scheme': 'https',
        'wsgi.input': BytesIO(raw_body),
        'wsgi.errors': BytesIO(),
        'wsgi.multithread': False,
        'wsgi.multiprocess': False,
        'wsgi.run_once': True,
        'CONTENT_LENGTH': str(len(raw_body)),
    }

    # 转换请求头为 WSGI 格式（包含 Cookie 用于 session 认证）
    for key, value in headers.items():
        wsgi_key = 'HTTP_' + key.upper().replace('-', '_')
        environ[wsgi_key] = value

    # Content-Type 特殊处理
    ct = headers.get('content-type') or headers.get('Content-Type') or ''
    if ct:
        environ['CONTENT_TYPE'] = ct

    # ---- 调用 Flask 应用 ----
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
        error_msg = json.dumps({'error': f'Server Error: {str(e)}'}, ensure_ascii=False)
        response_body = [error_msg.encode('utf-8')]
        status_code = 500
        response_headers = {'Content-Type': 'application/json; charset=utf-8'}

    # ---- 构建 Netlify 响应 ----
    body_bytes = b''.join(response_body)

    content_type = response_headers.get('Content-Type', 'text/html')
    is_binary = any(t in content_type for t in ['image', 'octet-stream', 'pdf'])

    if is_binary:
        return {
            'statusCode': status_code,
            'headers': response_headers,
            'body': base64.b64encode(body_bytes).decode('utf-8'),
            'isBase64Encoded': True,
        }

    return {
        'statusCode': status_code,
        'headers': response_headers,
        'body': body_bytes.decode('utf-8', errors='replace'),
    }
