"""WSGI 入口 — 兼容 Railway / Render / 任何平台。

启动命令（Procfile / Render）:
    gunicorn --timeout 120 --workers 1 wsgi:application

gunicorn 会自动读取平台注入的 PORT 环境变量（默认 8000）。
"""
import os
from app import app as application


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    application.run(host='0.0.0.0', port=port)
