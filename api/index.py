# -*- coding: utf-8 -*-
"""Vercel Serverless — Flask 入口"""

import sys
import os

# 项目根目录加入路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 标记为生产环境
os.environ['NETLIFY'] = 'true'

# 直接导出 Flask app
from app import app
