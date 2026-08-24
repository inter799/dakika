"""WSGI 入口 — 用 Python 端直接启动应用（完全避免 --bind 解析问题）。

Railway / Render 平台注入 PORT 环境变量，本文件读取后传给 Flask。
启动命令（Procfile / Render）:
    python wsgi.py
"""
import os
from app import app


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    # debug=False, 监听 0.0.0.0 接受外部访问
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
