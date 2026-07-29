# -*- coding: utf-8 -*-
"""
文件上传 → 题目生成器 v2.0
- 支持多厂商大模型 API（OpenAI / DeepSeek / Qwen / 智谱 / Kimi / 百度 / Claude / Gemini / 自定义）
- 内置无需 API 的纯本地模式
- 支持 txt / doc / docx / pdf / rtf / md / html / json / csv / xlsx / pptx / jpg / png / gif 等文件
- 本地登录认证，每次登录自动清空上一会话数据
"""

import os, json, re, base64, uuid, random, zipfile, shutil, hashlib, functools, secrets
from io import BytesIO
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, session, redirect, url_for, g, make_response
from flask_cors import CORS

# ==================== 可选依赖检测 ====================
try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

# PDF 解析（可选）
try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

# Excel / PPT 解析（可选）
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# OpenAI SDK 用于所有兼容接口
try:
    from openai import OpenAI
    HAS_OPENAI_SDK = True
except ImportError:
    HAS_OPENAI_SDK = False

# Google Gemini SDK（可选）
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# Anthropic SDK（可选）
try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

# DuckDuckGo 搜索（可选 - 网络搜题功能）
try:
    from duckduckgo_search import DDGS
    HAS_DDG = True
except ImportError:
    HAS_DDG = False

# BeautifulSoup + requests（可选 - 网页内容提取）
try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

try:
    import requests as http_requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# ==================== 平台检测 ====================
import platform as _platform
_IS_WINDOWS = _platform.system() == 'Windows'
_IS_NETLIFY = os.environ.get('NETLIFY', '') == 'true' or 'netlify' in os.environ.get('_', '')

# ==================== 初始化 ====================
app = Flask(__name__)
app.secret_key = 'quiz-generator-v2-local-auth-2024'  # 固定密钥，保证 session 持久
app.config['JSON_AS_ASCII'] = False  # 关键：确保中文JSON输出不被转义为\uXXXX
CORS(app, supports_credentials=True)

# 尝试初始化 Supabase（Netlify 部署用）
import supabase_client as db
_USE_SUPABASE = db.is_configured()
if _USE_SUPABASE:
    db.initialize()

if _IS_NETLIFY:
    DATA_ROOT = Path('/tmp/data')
    USER_FILE = Path('/tmp/users.json')
    UPLOAD_FOLDER = Path('/tmp/uploads')
    BANK_FILE = Path('/tmp/question_bank.json')
else:
    DATA_ROOT = Path(__file__).parent / 'data'
    USER_FILE = Path(__file__).parent / 'users.json'
    UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
    BANK_FILE = Path(__file__).parent / 'question_bank.json'

DATA_ROOT.mkdir(parents=True, exist_ok=True)
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)

REMEMBER_COOKIE_NAME = 'quiz_remember'
REMEMBER_DAYS = 30

# ==================== 用户持久化存储 ====================
def _load_users():
    """从文件加载用户数据（兼容旧文件存储）"""
    if _USE_SUPABASE:
        return {}  # Supabase 模式下不需要加载文件
    if USER_FILE.exists():
        try:
            with open(USER_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_users(users):
    """保存用户数据到文件（兼容旧文件存储）"""
    if _USE_SUPABASE:
        return
    with open(USER_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def _hash_pw(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()


def _generate_token():
    return secrets.token_hex(32)


# 启动时加载用户，如果为空则创建默认 admin 用户
if _USE_SUPABASE:
    # Supabase 模式下已在 db.initialize() 中创建默认用户
    _users_db = {}
else:
    _users_db = _load_users()
    if not _users_db:
        default_user = os.environ.get('QUIZ_USERNAME', 'admin')
        default_pw = os.environ.get('QUIZ_PASSWORD', 'admin123')
        _users_db[default_user] = {
            'password': _hash_pw(default_pw),
            'display_name': default_user,
            'created_at': datetime.now().isoformat(),
            'remember_token': None
        }
        _save_users(_users_db)


# ==================== 用户工具函数 ====================
def _verify_password(username, password):
    """验证用户名密码"""
    if _USE_SUPABASE:
        return db.verify_password(username, password)
    user = _users_db.get(username)
    if not user:
        return False
    return user['password'] == _hash_pw(password)


def _get_user_uploads(username):
    """获取用户专属 uploads 目录（本地/临时）"""
    d = DATA_ROOT / username / 'uploads'
    d.mkdir(parents=True, exist_ok=True)
    return d


def _get_user_bank_file(username):
    """获取用户专属题库文件路径（仅文件存储模式）"""
    return DATA_ROOT / username / 'question_bank.json'


def _clear_user_data(username):
    """清空指定用户的所有数据"""
    if _USE_SUPABASE:
        db.clear_all_spaces(username)
        # 清空临时上传目录
        uploads_dir = DATA_ROOT / username / 'uploads'
        if uploads_dir.exists():
            shutil.rmtree(str(uploads_dir), ignore_errors=True)
        return
    user_dir = DATA_ROOT / username
    uploads_dir = user_dir / 'uploads'
    bank_file = user_dir / 'question_bank.json'
    # 清空 uploads
    if uploads_dir.exists():
        for item in uploads_dir.iterdir():
            try:
                if item.is_file():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception:
                pass
    # 重置题库
    _save_bank_for_user(username, _new_bank())


def _load_bank_for_user(username):
    """加载指定用户的题库"""
    if _USE_SUPABASE:
        return {"spaces": {}, "questions": [], "next_id": 1}  # Supabase 模式不使用 JSON
    bf = _get_user_bank_file(username)
    if bf.exists():
        try:
            with open(bf, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            return _new_bank()
        if "spaces" not in data:
            old_questions = data.get("questions", [])
            data["spaces"] = {}
            data["questions"] = []
            if old_questions:
                data["spaces"]["__legacy__"] = {
                    "name": "旧题库（迁移）",
                    "created_at": datetime.now().isoformat(),
                    "questions": old_questions,
                    "next_id": data.get("next_id", 1)
                }
            data["next_id"] = 1
            _save_bank_for_user(username, data)
        return data
    return _new_bank()


def _new_bank():
    return {"spaces": {}, "questions": [], "next_id": 1}


def _save_bank_for_user(username, bank):
    """保存指定用户的题库（仅文件存储模式）"""
    if _USE_SUPABASE:
        return
    bf = _get_user_bank_file(username)
    bf.parent.mkdir(parents=True, exist_ok=True)
    with open(bf, 'w', encoding='utf-8') as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


# ==================== 登录保护 + 自动登录 ====================
@app.before_request
def check_auto_login():
    """请求前检查：自动登录（remember-me cookie）+ 设置用户专属路径"""
    # 已登录 → 设置路径
    if session.get('logged_in'):
        g.user_uploads = _get_user_uploads(session['username'])
        g.user_bank_file = _get_user_bank_file(session['username'])
        return

    # 未登录 → 尝试自动登录
    token = request.cookies.get(REMEMBER_COOKIE_NAME)
    if token:
        if _USE_SUPABASE:
            udata = db.find_by_remember_token(token)
            if udata:
                session['logged_in'] = True
                session['username'] = udata['username']
                session['display_name'] = udata.get('display_name', udata['username'])
                g.user_uploads = _get_user_uploads(udata['username'])
                g.user_bank_file = _get_user_bank_file(udata['username'])
                return
        else:
            global _users_db
            _users_db = _load_users()
            for uname, udata in _users_db.items():
                if udata.get('remember_token') and udata.get('remember_token') == token:
                    session['logged_in'] = True
                    session['username'] = uname
                    session['display_name'] = udata.get('display_name', uname)
                    g.user_uploads = _get_user_uploads(uname)
                    g.user_bank_file = _get_user_bank_file(uname)
                    return

    # 未登录也未匹配 token → 设置默认路径（防止 500）
    g.user_uploads = None
    g.user_bank_file = None


def login_required(f):
    """登录保护装饰器：API 路由返回 401，页面路由重定向到登录页"""
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            if request.path.startswith('/api/'):
                return jsonify({"error": "未登录", "login_required": True}), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated


# ==================== 兼容旧函数名的便捷方法 ====================
def _load_bank():
    """加载当前用户题库（兼容旧调用）"""
    return _load_bank_for_user(session.get('username', '_default'))


def _save_bank(bank):
    """保存当前用户题库（兼容旧调用）"""
    _save_bank_for_user(session.get('username', '_default'), bank)


def _generate_space_id(name):
    """根据文档名生成唯一 space_id"""
    base = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '_', name)[:30]
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    return f"{base}_{ts}"

ALLOWED_EXTENSIONS = {
    # 文本文档
    'txt', 'md', 'json', 'csv', 'xml', 'log', 'yaml', 'yml',
    # 富文本文档
    'rtf', 'html', 'htm',
    # Word
    'doc', 'docx',
    # PDF
    'pdf',
    # Excel
    'xlsx', 'xlsm',
    # PowerPoint
    'pptx',
    # 图片
    'jpg', 'jpeg', 'png', 'bmp', 'gif'
}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB


def safe_filename(filename):
    """保留中文等非ASCII字符的安全文件名处理"""
    # 只移除路径分隔符和危险字符，保留中文等Unicode字符
    unsafe_chars = '<>:"/\\|?*'
    name = filename
    for ch in unsafe_chars:
        name = name.replace(ch, '_')
    # 去除首尾空格和点
    name = name.strip(' .')
    # 限制长度
    if len(name) > 200:
        base, ext = (name.rsplit('.', 1) if '.' in name else (name, ''))
        name = base[:195] + ('.' + ext if ext else '')
    return name or 'unnamed_file'


# ==================== 大模型提供商注册表 ====================
PROVIDERS = {
    "openai": {
        "name": "OpenAI",
        "base_url": "https://api.openai.com/v1",
        "models": ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-4", "gpt-3.5-turbo"],
        "default_model": "gpt-3.5-turbo",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",     # 使用 openai SDK 调用
        "desc": "最通用的大模型，效果优秀"
    },
    "deepseek": {
        "name": "DeepSeek",
        "base_url": "https://api.deepseek.com/v1",
        "models": ["deepseek-chat", "deepseek-reasoner"],
        "default_model": "deepseek-chat",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "国产高性价比模型，中文能力出色"
    },
    "qwen": {
        "name": "通义千问 (Qwen)",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "models": ["qwen-plus", "qwen-max", "qwen-turbo", "qwen-long"],
        "default_model": "qwen-plus",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "阿里云出品，综合能力强"
    },
    "zhipu": {
        "name": "智谱AI (GLM)",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "models": ["glm-4-plus", "glm-4", "glm-4-flash", "glm-4-air", "glm-4-long"],
        "default_model": "glm-4-flash",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "清华系大模型，性价比高"
    },
    "moonshot": {
        "name": "Moonshot (Kimi)",
        "base_url": "https://api.moonshot.cn/v1",
        "models": ["moonshot-v1-8k", "moonshot-v1-32k", "moonshot-v1-128k"],
        "default_model": "moonshot-v1-8k",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "月之暗面出品，超长上下文"
    },
    "baidu": {
        "name": "百度文心 (ERNIE)",
        "base_url": "https://qianfan.baidubce.com/v2",
        "models": ["ernie-4.0-turbo-8k", "ernie-3.5-8k", "ernie-speed-8k", "ernie-lite-8k"],
        "default_model": "ernie-speed-8k",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "百度大模型，中文理解能力强"
    },
    "claude": {
        "name": "Anthropic Claude",
        "base_url": "https://api.anthropic.com",
        "models": ["claude-3-5-sonnet-20241022", "claude-3-opus-20240229", "claude-3-haiku-20240307"],
        "default_model": "claude-3-5-sonnet-20241022",
        "auth_header": "x-api-key: {api_key}",
        "sdk_type": "anthropic",
        "desc": "逻辑推理能力极强"
    },
    "gemini": {
        "name": "Google Gemini",
        "base_url": "https://generativelanguage.googleapis.com",
        "models": ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"],
        "default_model": "gemini-2.0-flash",
        "auth_header": "key={api_key}",
        "sdk_type": "gemini",
        "desc": "Google 多模态大模型"
    },
    "custom": {
        "name": "自定义接口",
        "base_url": "",
        "models": [],
        "default_model": "",
        "auth_header": "Bearer {api_key}",
        "sdk_type": "openai",
        "desc": "兼容 OpenAI 格式的任意接口"
    }
}

# ==================== 文本提取 ====================
def extract_text_from_txt(filepath):
    for enc in ['utf-8', 'utf-16', 'utf-16le', 'utf-16be', 'gbk', 'gb2312', 'latin-1']:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    try:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            return f.read()
    except Exception:
        return "[错误] 无法解码该文本文件"


def _is_zip_docx(filepath):
    """检测文件是否为标准的 zip 格式 docx"""
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            return 'word/document.xml' in z.namelist()
    except Exception:
        return False


def extract_text_from_docx(filepath):
    if not HAS_DOCX:
        return "[错误] 请安装 python-docx: pip install python-docx"

    # 先检查是否为合法 docx（zip 包）
    if not _is_zip_docx(filepath):
        return "[错误] 该文件不是标准的 .docx 格式。可能是旧版 .doc 文件改名，请用 Word 打开后另存为 .docx，或转换为 .txt 上传。"

    try:
        doc = DocxDocument(filepath)
        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        paras.append(cell.text.strip())
        text = '\n'.join(paras)
        if not text.strip():
            return "[提示] 该 docx 文件内容为空，或文字位于图片/文本框中无法提取。"
        return text
    except Exception as e:
        return f"[错误] 解析docx失败: {e}"


def extract_text_from_doc(filepath):
    """.doc 旧版二进制格式：尝试提取可读文本"""
    com_errors = []
    try:
        # 方法1: 尝试作为 zip 读取（有些 .doc 实际是 .docx）
        if _is_zip_docx(filepath):
            return extract_text_from_docx(filepath)

        # 方法2: 读取原始字节，提取可读文本段落
        try:
            with open(filepath, 'rb') as f:
                raw = f.read()

            text_parts = []
            for enc in ['utf-8', 'gbk', 'gb2312', 'utf-16-le', 'latin-1']:
                try:
                    decoded = raw.decode(enc, errors='ignore')
                    # 提取连续的可读中英文段落（至少20个字符）
                    paragraphs = re.findall('[\\u4e00-\\u9fa5a-zA-Z0-9\\s，。！？；：、""''《》（）…—.,!?;:()\\[\\]【】]{20,}', decoded)
                    if paragraphs:
                        text_parts.extend(paragraphs)
                        break  # 找到合适编码就停止
                except Exception:
                    continue

            if text_parts:
                result = '\n\n'.join(text_parts)
                if len(result) >= 100:
                    return result
        except Exception:
            pass

        # 方法3: Windows 上调用 Word / WPS COM 接口提取文本
        # Flask 是多线程环境，必须在每个线程中初始化 COM
        word = None
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception as e:
            com_errors.append(f"COM初始化失败: {e}")

        try:
            import win32com.client as win32
            for progid in ["Word.Application", "KWPS.Application", "WPS.Application", "Kwps.Application", "Et.Application"]:
                try:
                    word = win32.Dispatch(progid)
                    break
                except Exception as e:
                    com_errors.append(f"{progid}: {e}")
                    word = None

            if word:
                word.Visible = False
                word.DisplayAlerts = 0
                abs_path = os.path.abspath(filepath)
                doc = word.Documents.Open(abs_path, ReadOnly=True)
                text = doc.Content.Text
                doc.Close(SaveChanges=False)
                extracted = text.strip() if text else ''
                # 即使 Word 退出报错，也不应影响已提取的文本
                try:
                    word.Quit()
                except Exception as e:
                    com_errors.append(f"Word.Application.Quit: {e}")
                if extracted:
                    return extracted
        except Exception as e:
            com_errors.append(f"COM提取失败: {e}")
        finally:
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

        # 方法4: 尝试用 antiword 命令行工具（Linux/macOS）
        import subprocess
        try:
            proc = subprocess.run(['antiword', str(filepath)], capture_output=True, text=True, timeout=10)
            if proc.returncode == 0 and proc.stdout.strip():
                return proc.stdout.strip()
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        err_detail = '; '.join(com_errors) if com_errors else '未知原因'
        return (f"[错误] 无法解析旧版 .doc 文件。\n"
                f"[建议] 请用 Word/WPS 打开后另存为 .docx 或 .txt 再上传。\n"
                f"[详情] {err_detail}")
    except Exception as e:
        return f"[错误] 无法解析旧版 .doc 文件: {e}。请确保电脑已安装 Word 或 WPS，或者手动另存为 .docx / .txt 后重新上传。"


def extract_text_from_image(filepath):
    if not HAS_PIL:
        return "[错误] 请安装 Pillow: pip install Pillow"
    try:
        img = Image.open(filepath)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        if HAS_TESSERACT:
            text = pytesseract.image_to_string(img, lang='chi_sim+eng')
            return text if text.strip() else "[提示] 图片中未检测到文字"
        else:
            buffered = BytesIO()
            img.save(buffered, format='PNG')
            return "[IMAGE_BASE64]" + base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        return f"[错误] 解析图片失败: {e}"


def extract_text_from_pdf(filepath):
    """解析 PDF 文件，优先使用 pdfplumber，回退到 PyPDF2"""
    if HAS_PDFPLUMBER:
        try:
            pages_text = []
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        pages_text.append(page_text)
            text = '\n\n'.join(pages_text)
            if text.strip():
                return text
        except Exception as e:
            pass  # 回退到 PyPDF2

    if HAS_PYPDF2:
        try:
            pages_text = []
            with open(filepath, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        pages_text.append(page_text)
            text = '\n\n'.join(pages_text)
            if text.strip():
                return text
            return "[提示] PDF 中未检测到可提取文字（可能是扫描版图片PDF，需要 OCR）"
        except Exception as e:
            return f"[错误] 解析 PDF 失败: {e}。请安装依赖: py -m pip install pdfplumber PyPDF2"

    return "[错误] 请安装 PDF 解析依赖: py -m pip install pdfplumber PyPDF2"


def extract_text_from_rtf(filepath):
    """解析 RTF 富文本格式：剥离控制符，保留可见文本"""
    try:
        with open(filepath, 'rb') as f:
            raw = f.read()
        # 尝试不同编码
        for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1']:
            try:
                text = raw.decode(enc, errors='ignore')
                break
            except Exception:
                text = ''
        if not text:
            return "[错误] 无法读取 RTF 文件"
        # 简单剥离 RTF 控制符
        text = re.sub(r'\\[a-z]+\d*\s?', '', text)
        text = re.sub(r'\\[*\{\}\~\-]', '', text)
        text = re.sub(r'\{\\[^}]+\}', '', text)
        text = re.sub(r'[\{\}]', '', text)
        text = re.sub(r'\\par\b', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'\\tab\b', '\t', text, flags=re.IGNORECASE)
        text = re.sub(r'\\line\b', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'\\[0-9a-fA-F]{2}', '', text)
        text = re.sub(r'\\$', '', text)
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        text = '\n'.join(line.strip() for line in text.splitlines() if line.strip())
        return text if text.strip() else "[提示] RTF 文件中未检测到文字"
    except Exception as e:
        return f"[错误] 解析 RTF 失败: {e}"


def extract_text_from_plain(filepath):
    """读取普通文本类文件：md / html / htm / json / csv / xml / log / yaml / yml"""
    try:
        with open(filepath, 'rb') as f:
            raw = f.read()
        for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1']:
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                text = raw.decode('utf-8', errors='replace')
        return text
    except Exception as e:
        return f"[错误] 读取文本文件失败: {e}"


def extract_text_from_xlsx(filepath):
    """解析 Excel 表格文件，读取每个单元格的文本"""
    if not HAS_OPENPYXL:
        return "[错误] 请安装 openpyxl: py -m pip install openpyxl"
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        lines = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_text = ' '.join(str(cell) for cell in row if cell is not None)
                if row_text.strip():
                    lines.append(row_text)
        text = '\n'.join(lines)
        return text if text.strip() else "[提示] Excel 文件中未检测到文字"
    except Exception as e:
        return f"[错误] 解析 Excel 失败: {e}"


def extract_text_from_pptx(filepath):
    """解析 PowerPoint 演示文稿"""
    if not HAS_PPTX:
        return "[错误] 请安装 python-pptx: py -m pip install python-pptx"
    try:
        prs = Presentation(filepath)
        texts = []
        for slide_idx, slide in enumerate(prs.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text.strip():
                    texts.append(shape.text.strip())
        text = '\n\n'.join(texts)
        return text if text.strip() else "[提示] PPT 中未检测到文字"
    except Exception as e:
        return f"[错误] 解析 PPT 失败: {e}"


def extract_text(filepath, original_filename):
    ext = original_filename.rsplit('.', 1)[1].lower() if '.' in original_filename else ''
    handlers = {
        # Word
        'docx': extract_text_from_docx,
        'doc': extract_text_from_doc,
        # PDF
        'pdf': extract_text_from_pdf,
        # RTF
        'rtf': extract_text_from_rtf,
        # Excel
        'xlsx': extract_text_from_xlsx,
        'xlsm': extract_text_from_xlsx,
        # PowerPoint
        'pptx': extract_text_from_pptx,
        # 图片
        'jpg': extract_text_from_image, 'jpeg': extract_text_from_image,
        'png': extract_text_from_image, 'bmp': extract_text_from_image,
        'gif': extract_text_from_image,
        # 纯文本类
        'txt': extract_text_from_plain,
        'md': extract_text_from_plain,
        'json': extract_text_from_plain,
        'csv': extract_text_from_plain,
        'xml': extract_text_from_plain,
        'log': extract_text_from_plain,
        'yaml': extract_text_from_plain,
        'yml': extract_text_from_plain,
        'html': extract_text_from_plain,
        'htm': extract_text_from_plain,
    }
    return handlers.get(ext, lambda _: f"[错误] 不支持 .{ext} 格式")(filepath)


# ==================== 文本规范化 ====================

def normalize_text(text):
    """规范化从各种来源提取的文本，统一字符格式，提高题目识别率"""
    if not text:
        return text
    # 全角字母 → 半角
    text = text.translate(str.maketrans(
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
        'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
        '０１２３４５６７８９',
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        'abcdefghijklmnopqrstuvwxyz'
        '0123456789'
    ))
    # 全角标点 → 半角
    text = text.replace('．', '.').replace('：', ':').replace('，', ',')
    text = text.replace('（', '(').replace('）', ')').replace('【', '[').replace('】', ']')
    text = text.replace('《', '<').replace('》', '>').replace('；', ';')
    text = text.replace('？', '?').replace('！', '!').replace('～', '~')
    # 全角空格 → 半角空格
    text = text.replace('\u3000', ' ')
    # 统一换行符
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    # 合并多余空行（最多保留一个空行）
    text = re.sub(r'\n{3,}', '\n\n', text)
    # 清理每行首尾空白
    lines = [line.strip() for line in text.split('\n')]
    text = '\n'.join(lines)
    return text


# ==================== 题目识别与解析 ====================

_Q_SPLIT = re.compile(
    r'(?:^|\n)\s*'
    r'(?:'
    r'\d+[\.、．\)）]\s*'
    r'|[(（]\s*\d+\s*[)）]\s*'
    r'|第\s*\d+\s*题\s*'
    r'|第\s*[一二三四五六七八九十]+\s*题\s*'
    r')',
    re.MULTILINE
)

_OPT_RE = re.compile(r'([A-E])[\.、．\)）]\s*')
_OPT_PAREN_RE = re.compile(r'[（(]\s*([A-E])\s*[)）]\s*')
# 题干中直接标注答案：...（ B ）... / ...(ABC)... / ...（正确）... 或 【答案：B】/【答案：ABC】
_STEM_INLINE_ANSWER_RE = re.compile(
    r'[（(]\s*([A-Ea-e]+)\s*[)）]|'
    r'【\s*答案\s*[：:]\s*([A-Ea-e]+)\s*】|'
    r'[（(]\s*(正确|错误|对|错)\s*[)）]',
    re.IGNORECASE
)



_ANS_RE = re.compile(
    r'(?:标准答案|正确答案|参考答案|答案|建议答案|正确选项)'
    r'[是为]?\s*[：:]\s*[（(]?\s*'
    r'([A-Ea-e]+(?:[、，,\s/]+[A-Ea-e]+)*)\s*[)）]?'
    r'|(?:标准答案|正确答案|参考答案|答案|建议答案)[是为]?\s*[：:]\s*(正确|错误|对|错)'
    r'|[【\[]\s*(?:标准答案|正确答案|参考答案|答案)\s*[】\]]\s*[（(]?\s*([A-Ea-e]+)\s*[)）]?',
    re.IGNORECASE
)
_ANS_TF_RE = re.compile(r'(?:答案|正确答案)[：:是为]?\s*(正确|错误|对|错)')
_ANS_LETTER_RE = re.compile(r'(?:答案|正确答案|参考答案)[：:是为]?\s*[（(]?\s*([A-Ea-e]+)\s*[)）]?')


def _count_question_numbers(text):
    return len(_Q_SPLIT.findall(text))


def detect_existing_questions(text):
    if not text or len(text) < 20:
        return False, 0, ''
    q_count = _count_question_numbers(text)
    if q_count < 1:
        return False, 0, ''
    opt_count = len(_OPT_RE.findall(text)) + len(_OPT_PAREN_RE.findall(text))
    ans_count = len(_ANS_RE.findall(text)) + len(_ANS_TF_RE.findall(text))
    if q_count >= 1 and (opt_count >= 2 or ans_count >= 1):
        if opt_count >= ans_count:
            return True, q_count, 'choice'
        return True, q_count, 'mixed'
    return False, 0, ''


def parse_questions_from_text(text, question_type='choice'):
    questions = []
    parts = _Q_SPLIT.split(text)
    if not parts or len(parts) < 2:
        q = _parse_one_block(text.strip(), 1)
        return [q] if q else []

    q_num = 0
    for i in range(1, len(parts)):
        block = parts[i].strip()
        if not block or len(block) < 4:
            continue
        q_num += 1
        q = _parse_one_block(block, q_num)
        if q:
            questions.append(q)

    if not questions:
        questions = _parse_tf_questions(text)

    return questions


def _parse_one_block(block, q_id):
    if len(block) < 4:
        return None

    answer = ''
    ans_pos = len(block)

    # 1. 先尝试识别题干中直接括号标注的答案，例如 "...（ B ）的关系。" / "...（ABC）..." / "...（正确）..."
    inline_m = _STEM_INLINE_ANSWER_RE.search(block)
    if inline_m:
        raw = inline_m.group(1) or inline_m.group(2) or inline_m.group(3)
        raw = (raw or '').strip()
        # 如果答案前面已经出现 A. B. C. 类选项标记，说明这个括号是选项本身，不是答案
        text_before = block[:inline_m.start()]
        if re.search(r'[A-E][\.、．\)）]', text_before[-200:] if len(text_before) > 200 else text_before):
            answer = ''
        else:
            if raw in ('正确', '对'):
                answer = '正确'
            elif raw in ('错误', '错'):
                answer = '错误'
            else:
                # 多选/单选答案字母，去重排序（如 ABC -> ABC）
                letters = sorted(set(re.findall(r'[A-Ea-e]', raw)))
                answer = ''.join(letters).upper()
            # 内联答案在题干中：把答案标记从 block 中移除，避免影响选项提取
            block = block[:inline_m.start()] + block[inline_m.end():]




    # 2. 显式 "答案：B" 类标注
    if not answer:
        m_ans = _ANS_RE.search(block)
        if m_ans:
            raw = m_ans.group(1) or m_ans.group(3) or ''
            tf = m_ans.group(2) or ''
            if tf:
                answer = '正确' if tf in ('正确', '对') else '错误'
            elif raw:
                answer = ''.join(re.findall(r'[A-Ea-e]', raw)).upper()
                if len(answer) > 1:
                    answer = ''.join(sorted(set(answer)))
            if answer:
                ans_pos = m_ans.start()

    if not answer:
        m_tf = _ANS_TF_RE.search(block)
        if m_tf:
            tf = m_tf.group(1)
            answer = '正确' if tf in ('正确', '对') else '错误'
            ans_pos = m_tf.start()


    options = []
    opt_start = len(block)

    for m in _OPT_RE.finditer(block):
        if m.start() >= ans_pos:
            break
        letter = m.group(1).upper()
        opt_start = min(opt_start, m.start())
        content_start = m.end()
        next_m = _OPT_RE.search(block, content_start)
        content_end = next_m.start() if (next_m and next_m.start() < ans_pos) else ans_pos
        content = block[content_start:content_end].strip()
        options.append((letter, content))

    if not options:
        for m in _OPT_PAREN_RE.finditer(block):
            if m.start() >= ans_pos:
                break
            letter = m.group(1).upper()
            opt_start = min(opt_start, m.start())
            content_start = m.end()
            next_m = _OPT_PAREN_RE.search(block, content_start)
            content_end = next_m.start() if (next_m and next_m.start() < ans_pos) else ans_pos
            content = block[content_start:content_end].strip()
            options.append((letter, content))

    if options:
        stem = block[:opt_start]
    else:
        stem = block[:ans_pos] if ans_pos < len(block) else block

    stem = _clean_stem(stem)

    if not stem or len(stem.strip()) < 1:
        return None

    qtype = _detect_type(block, answer, options)

    opt_list = [f'{letter}. {content}' for letter, content in options]

    if not opt_list and qtype == 'tf':
        opt_list = ['正确. 正确', '错误. 错误']
    elif not opt_list and qtype in ('choice', 'multi'):
        opt_list = ['A. 选项A', 'B. 选项B', 'C. 选项C', 'D. 选项D']

    answer_guessed = False
    if not answer:
        answer_guessed = True
        if qtype == 'tf':
            answer = '正确' if q_id % 2 == 0 else '错误'
        else:
            answer = options[0][0] if options else 'A'

    analysis = _gen_analysis(qtype, stem, answer, opt_list)

    result = {
        'id': q_id,
        'type': qtype,
        'stem': stem.strip(),
        'options': opt_list,
        'answer': answer,
        'analysis': analysis,
    }
    if answer_guessed:
        result['answer_guessed'] = True
    return result


def _clean_stem(stem):
    stem = _ANS_RE.sub('', stem)
    stem = _ANS_TF_RE.sub('', stem)
    stem = _ANS_LETTER_RE.sub('', stem)
    # 去除题干中直接括号标注的答案，如 （ B ） 或 (B)
    stem = _STEM_INLINE_ANSWER_RE.sub('', stem)
    stem = _OPT_RE.sub(' ', stem)
    stem = _OPT_PAREN_RE.sub(' ', stem)
    stem = re.sub(r'\s+[（(]?\s*[A-Ea-e]\s*[)）]?\s*$', '', stem)
    stem = re.sub(r'(?:标准答案|正确答案|参考答案|答案|建议答案)\s*[：:是为]?\s*$', '', stem)
    stem = re.sub(r'\s+', ' ', stem)
    stem = stem.strip().rstrip('.,，。、;；:：')
    return stem



def _detect_type(block, answer, options):
    # 判断题：2个选项 + 选项包含"正确/错误/对/错" + 答案也是正确/错误
    if len(options) == 2:
        texts = ' '.join(o[1] for o in options if isinstance(o, tuple) and len(o) > 1)
        if re.search(r'正确|错误|对|错', texts):
            if answer in ('正确', '错误', '对', '错'):
                return 'tf'

    # 答案长度为 >=2 且全是字母 → 多选题
    if len(answer) >= 2 and re.match(r'^[A-E]+$', answer):
        return 'multi'

    # 题干明确标注多选/不定项/多项
    if re.search(r'多选|不定项|多项|至少.*选|选出.?包括', block):
        return 'multi'

    # 选项数量 >= 5 时，如果是多选题答案字母数 >=2 才判多选
    if len(options) >= 5 and len(answer) >= 2 and re.match(r'^[A-E]+$', answer):
        return 'multi'

    return 'choice'


def _gen_analysis(qtype, stem, answer, opt_list):
    if not opt_list:
        return '请参考原文内容。'

    ans_letters = list(answer) if answer else []
    correct_texts = []
    for opt in opt_list:
        parts = opt.split('.', 1) if '.' in opt else (opt[:1], opt[1:])
        letter = parts[0].strip()
        content = parts[1].strip() if len(parts) > 1 else opt
        if letter in ans_letters:
            correct_texts.append(content)

    correct_str = '；'.join(correct_texts[:3]) if correct_texts else '相关选项'

    if qtype == 'tf':
        state = '符合' if answer == '正确' else '不符合'
        return f'该陈述{state}原文内容。'
    elif qtype == 'multi':
        return f'本题为多选题，共 {len(ans_letters)} 个正确选项。正确答案 {answer}：{correct_str}。'
    else:
        return f'正确答案为 {answer}：{correct_str}。'

def _parse_tf_questions(text):
    """解析判断题：题号 + 陈述 + 答案标记"""
    questions = []

    blocks = re.split(
        r'(?:^|\n)\s*(?:\d+[\.、．\)]\s*|[(（]\d+[)）]\s*|第[一二三四五六七八九十\d]+题\s*)',
        text
    )
    if len(blocks) < 2:
        return []

    for idx, block in enumerate(blocks[1:], start=1):
        block = block.strip()
        if not block or len(block) < 8:
            continue

        # 提取答案
        answer = ""
        ans_patterns = [
            r'(?:答案|正确答案)[：:是为]?\s*(正确|错误|对|错|True|False|T|F)',
            r'[（(]\s*(正确|错误|对|错|[×✓✔✗✘√])\s*[)）]',
            r'(正确|错误|对|错)\s*[（(]\s*[)）]',
        ]
        for pat in ans_patterns:
            m = re.search(pat, block)
            if m:
                raw = m.group(1)
                if raw in ('正确', '对', '✓', '✔', '√', 'True', 'T', 'true', 't'):
                    answer = "正确"
                elif raw in ('错误', '错', '×', '✗', '✘', 'False', 'F', 'false', 'f'):
                    answer = "错误"
                break

        # 去除答案部分，剩余作为题干
        stem = block
        for pat in ans_patterns:
            stem = re.sub(pat, '', stem).strip()
        stem = re.sub(r'\s*[（(]\s*[)）]\s*', '', stem).strip()

        if not stem or len(stem) < 4:
            continue

        analysis = ""
        am = re.search(r'(?:解析|分析)[：:]\s*(.+?)(?=\n|$)', block)
        if am:
            analysis = am.group(1).strip()

        questions.append({
            "id": len(questions) + 1,
            "type": "tf",
            "stem": stem,
            "answer": answer if answer else ("正确" if idx % 2 == 0 else "错误"),
            "analysis": analysis if analysis else "从原文解析提取"
        })

    return questions


def _guess_answer(options):
    """当没有明确答案时，默认选A"""
    return options[0][0] if options else "A"


# ==================== 智能题数估算 ====================
def estimate_question_count(text, max_count=50):
    """根据文本长度和内容密度估算合理的题目数量"""
    text_len = len(text)
    # 统计句子数
    sentences = re.split(r'[。！？\n；;]+', text)
    sentence_count = len([s for s in sentences if len(s.strip()) >= 10])
    # 每2-3个有效句子可出1道题
    estimated = max(3, min(max_count, sentence_count // 2))
    return estimated


# ==================== Prompt 构建 ====================
def build_prompt(text, question_type, question_count):
    """构建 AI 生成单类题型的 prompt"""
    text_content = text[:15000]
    count_instruction = "尽可能多地生成题目（根据内容量自行决定数量，至少5道），完整覆盖全部关键知识点" if question_count == 0 else f"生成 {question_count} 道题"

    base = f"""你是一位资深教育专家。请仔细阅读以下内容，{count_instruction}。题目必须严格基于原文内容，不要脱离材料编造。每题给出正确答案和简短解析。题目之间要有区分度，避免重复考察同一知识点。

【内容】
---
{text_content}
---

【要求】"""

    if question_type == 'choice':
        return base + """每道题4个选项(A/B/C/D)，仅1个正确答案，选项要有迷惑性。

【输出JSON格式】只输出JSON：
{"questions":[{"id":1,"stem":"题目","options":["A. ...","B. ...","C. ...","D. ..."],"answer":"A","analysis":"解析"}]}"""

    elif question_type == 'tf':
        return base + """正确和错误的题目各占一半左右，错误题目要有合理误导性。

【输出JSON格式】只输出JSON：
{"questions":[{"id":1,"stem":"题目陈述","answer":"正确/错误","analysis":"解析"}]}"""

    elif question_type == 'fill':
        return base + """每道题是一个句子，在关键处留出1个空位（用____表示），答案为原文中的关键词或短语。

【输出JSON格式】只输出JSON：
{"questions":[{"id":1,"stem":"带有____的句子","answer":"正确答案","analysis":"解析"}]}"""

    elif question_type == 'essay':
        return base + """每道题为大题/问答题，要求综合分析原文，给出参考答案和评分要点。

【输出JSON格式】只输出JSON：
{"questions":[{"id":1,"stem":"问题","answer":"参考答案","key_points":["要点1","要点2","要点3"],"analysis":"解析","difficulty":"easy/medium/hard"}]}"""

    return ""


def build_prompt_all(text, count_per_type):
    """一次性生成全部四种题型，JSON 中按类型分组"""
    text_content = text[:15000]
    count_text = "每种题型根据内容量自行决定数量，至少3道，完整覆盖全部关键知识点" if count_per_type == 0 else f"每种题型生成 {count_per_type} 道题"
    return f"""你是一位资深教育专家。请仔细阅读以下内容，{count_text}。

【内容】
---
{text_content}
---

【要求】
1. 严格基于原文内容，不要脱离材料编造。
2. 按以下四种题型分别出题，并把题目放在对应的类型数组中：
   - choice: 单项选择题（4个选项A/B/C/D，仅1个正确答案）
   - tf: 判断题（正确/错误）
   - fill: 填空题（stem中用____表示空缺，answer为原文关键词）
   - essay: 大题/问答题（给出参考答案和评分要点 key_points）
3. 每题都要给出 answer 和 analysis，essay 题要有 key_points 和 difficulty。
4. 题目之间要有区分度。

【输出JSON格式】只输出JSON，不要其他文字：
{{"choice":{{"questions":[{{"id":1,"stem":"...","options":["A. ...","B. ...","C. ...","D. ..."],"answer":"A","analysis":"..."}}]}},
"tf":{{"questions":[{{"id":1,"stem":"...","answer":"正确/错误","analysis":"..."}}]}},
"fill":{{"questions":[{{"id":1,"stem":"... ____ ...","answer":"...","analysis":"..."}}]}},
"essay":{{"questions":[{{"id":1,"stem":"...","answer":"...","key_points":["..."],"analysis":"...","difficulty":"medium"}}]}}}}"""



# ==================== AI 调用引擎 ====================
def call_openai_compatible(provider_id, api_key, api_base, model, messages, **kwargs):
    """通过 OpenAI SDK 调用（OpenAI/DeepSeek/Qwen/智谱/Kimi/百度/自定义）"""
    if not HAS_OPENAI_SDK:
        raise RuntimeError("请安装 openai: pip install openai")
    client = OpenAI(api_key=api_key, base_url=api_base.rstrip('/') + '/v1' if not api_base.rstrip('/').endswith('/v1') else api_base.rstrip('/'))
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=messages,
            temperature=0.7,
            max_tokens=kwargs.get('max_tokens', 3000)
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        raise RuntimeError(f"{PROVIDERS.get(provider_id,{}).get('name',provider_id)} 调用失败: {str(e)}")


def call_anthropic(provider_id, api_key, api_base, model, messages, **kwargs):
    """通过 Anthropic SDK 调用 Claude"""
    if not HAS_ANTHROPIC:
        raise RuntimeError("请安装 anthropic: pip install anthropic")
    client = anthropic.Anthropic(api_key=api_key)
    # 提取 system prompt
    system_msg = ""
    user_msgs = []
    for m in messages:
        if m["role"] == "system":
            system_msg = m["content"]
        else:
            user_msgs.append(m["content"])
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=kwargs.get('max_tokens', 3000),
            system=system_msg if system_msg else anthropic.NOT_GIVEN,
            messages=[{"role": "user", "content": "\n".join(user_msgs)}]
        )
        return resp.content[0].text
    except Exception as e:
        raise RuntimeError(f"Claude 调用失败: {str(e)}")


def call_gemini(provider_id, api_key, api_base, model, messages, **kwargs):
    """通过 Google SDK 调用 Gemini"""
    if not HAS_GEMINI:
        raise RuntimeError("请安装 google-generativeai: pip install google-generativeai")
    genai.configure(api_key=api_key)
    system_msg = ""
    user_content = ""
    for m in messages:
        if m["role"] == "system":
            system_msg = m["content"]
        else:
            user_content += m["content"] + "\n"

    full_prompt = f"[系统指令: {system_msg}]\n\n{user_content}" if system_msg else user_content
    try:
        gm = genai.GenerativeModel(model)
        resp = gm.generate_content(full_prompt)
        return resp.text
    except Exception as e:
        raise RuntimeError(f"Gemini 调用失败: {str(e)}")


def call_ai_api(provider_id, api_key, api_base, model, text, question_type, question_count):
    """统一 AI 调用入口。
    question_type='all' 时一次性生成四种题型；
    question_count=0 表示自动决定数量。"""
    if not api_key:
        return {"error": "请先配置 API Key"}

    is_all = (question_type == 'all')
    prompt = build_prompt_all(text, question_count) if is_all else build_prompt(text, question_type, question_count)

    messages = [
        {"role": "system", "content": "你是专业出题助手。严格按JSON格式返回结果，不要输出其他内容。"},
        {"role": "user", "content": prompt}
    ]

    max_tok = 8000 if (question_count == 0 or is_all) else 3000

    provider = PROVIDERS.get(provider_id, PROVIDERS["custom"])
    sdk_type = provider["sdk_type"]

    try:
        if sdk_type == "anthropic":
            result_text = call_anthropic(provider_id, api_key, api_base, model, messages, max_tokens=max_tok)
        elif sdk_type == "gemini":
            result_text = call_gemini(provider_id, api_key, api_base, model, messages, max_tokens=max_tok)
        else:
            result_text = call_openai_compatible(provider_id, api_key, api_base, model, messages, max_tokens=max_tok)

        result_text = result_text.strip()
        if result_text.startswith('```'):
            result_text = re.sub(r'^```\w*\n?', '', result_text)
            result_text = re.sub(r'\n?```$', '', result_text)
            result_text = result_text.strip()

        data = json.loads(result_text)

        if is_all:
            # 按类型分组，保留空数组也返回
            grouped = {}
            total = 0
            for t in ['choice', 'tf', 'fill', 'essay']:
                arr = data.get(t, {}).get('questions', []) if isinstance(data.get(t), dict) else data.get(t, [])
                if not isinstance(arr, list):
                    arr = []
                grouped[t] = arr
                total += len(arr)
            return {"error": None, "grouped": grouped, "total": total, "provider_used": provider["name"]}
        else:
            return {"error": None, "questions": data.get('questions', []), "provider_used": provider["name"]}
    except json.JSONDecodeError:
        raw_preview = result_text[:200] if isinstance(result_text, str) else ""
        return {"error": "AI返回格式异常，请重试", "questions": [], "raw_preview": raw_preview}
    except Exception as e:
        return {"error": str(e), "questions": []}



# ==================== 纯本地模式（无需API） ====================
def _split_sentences(text):
    """智能分句"""
    sents = re.split(r'[。！？\n；;]+', text)
    return [s.strip() for s in sents if len(s.strip()) >= 12]


def _extract_keywords(text, topn=5):
    """简易关键词提取"""
    # 按词频统计2-4字中文词 + 英文词
    cn_words = re.findall(r'[\u4e00-\u9fa5]{2,4}', text)
    en_words = re.findall(r'[a-zA-Z]{3,}', text)
    all_words = cn_words + en_words

    freq = {}
    for w in all_words:
        if w.lower() not in {'这是', '不是', '一个', '可以', '进行', '使用', '没有',
                              '他们', '我们', '这个', '那个', '什么', '因为', '所以',
                              '而且', '但是', '如果', '就是', '已经', '还是', '或者'}:
            freq[w] = freq.get(w, 0) + 1

    sorted_words = sorted(freq.items(), key=lambda x: x[1], reverse=True)
    return [w for w, _ in sorted_words[:topn]]


def _extract_numbers(text):
    """提取数字相关信息"""
    patterns = [
        r'(\d+\.?\d*\s*[%％])',         # 百分比
        r'(\d{4}\s*年)',                  # 年份
        r'([约近超达]?\s*\d+\.?\d*\s*[万亿千百]?\s*[个只条项次元美元吨米克])'  # 数量
    ]
    results = []
    for p in patterns:
        results.extend(re.findall(p, text))
    return results[:5]


def generate_choice_local(text, count):
    """本地生成选择题（count=0 时自动使用全部句子）"""
    sents = _split_sentences(text)
    if not sents:
        return []

    keywords = _extract_keywords(text, 10)
    numbers = _extract_numbers(text)

    # count=0 表示使用全部句子，但最多50题
    max_q = min(len(sents), 50) if count == 0 else min(count, len(sents))

    questions = []
    for i in range(max_q):
        sent = sents[i % len(sents)]
        words = re.findall(r'[\u4e00-\u9fa5a-zA-Z]{2,}', sent)
        if not words:
            words = re.findall(r'\S+', sent)
        if not words:
            words = ['核心要素']

        # 用不同关键词做不同题型风格
        kw = words[i % len(words)]

        templates = [
            {
                "stem": f'根据原文内容，"{sent[:45]}..."，以下哪项说法是正确的？',
                "options": [
                    f"A. {kw}是核心要素",
                    f"B. 相关内容与{kw}无关",
                    f"C. 以上说法均不正确",
                    f"D. 需要结合上下文综合判断"
                ],
                "answer": "A",
                "analysis": f'原文"{sent[:80]}..."明确指出{kw}的关键地位。'
            },
            {
                "stem": f'"{sent[:45]}..."，在这段内容中主要强调的是？',
                "options": [
                    f"A. 方法论的重要性",
                    f"B. {kw}的核心作用",
                    f"C. 外部因素的影响",
                    f"D. 历史背景的分析"
                ],
                "answer": "B",
                "analysis": f'内容重点围绕"{kw}"展开论述。'
            },
            {
                "stem": f'以下哪项最符合原文中"{sent[:35]}..."的表述？',
                "options": [
                    f"A. 原文明确否定此观点",
                    f"B. {kw}在其中起决定性作用",
                    f"C. 没有足够信息判断",
                    f"D. 此观点需要更多证据支持"
                ],
                "answer": "B",
                "analysis": f'根据上下文，"{kw}"是核心概念。'
            },
            {
                "stem": f'关于"{sent[:45]}..."，以下理解最准确的是？',
                "options": [
                    f"A. 该描述主要指{kw}",
                    f"B. 该描述缺乏明确指向",
                    f"C. 该描述与主题无关",
                    f"D. 该描述存在歧义"
                ],
                "answer": "A",
                "analysis": f'综合分析，原文讨论的核心是{kw}。'
            }
        ]

        tmpl = templates[i % len(templates)]
        questions.append({
            "id": len(questions) + 1,
            "stem": tmpl["stem"],
            "options": tmpl["options"],
            "answer": tmpl["answer"],
            "analysis": tmpl["analysis"]
        })

        if len(questions) >= count:
            break

    return questions


def generate_essay_local(text, count):
    """本地生成问答题"""
    sents = _split_sentences(text)
    if not sents:
        return []

    keywords = _extract_keywords(text, 15)
    questions = []

    for i in range(count):
        sent = sents[i % len(sents)]
        kw = keywords[i % len(keywords)] if keywords else "相关内容"
        kw2 = keywords[(i + 1) % len(keywords)] if keywords else "相关概念"

        templates = [
            {
                "stem": f'请简述"{sent[:60]}..."的主要内容，并说明{kw}在其中的作用。',
                "answer": sent if len(sent) < 300 else sent[:300] + "...",
                "key_points": [
                    f"准确概括核心内容",
                    f"阐述{kw}的定义与内涵",
                    f"分析{kw}与整体内容的关系",
                    f"给出自己的理解或总结"
                ],
                "difficulty": "medium"
            },
            {
                "stem": f'结合原文分析，{kw}与{kw2}之间存在怎样的关系？请举例说明。',
                "answer": f"根据原文内容，{kw}和{kw2}存在相互关联。{sent[:200]}",
                "key_points": [
                    f"明确{kw}的定义",
                    f"分析{kw2}的内涵",
                    f"阐述两者之间的关系",
                    f"举出原文中的具体例子"
                ],
                "difficulty": "hard"
            },
            {
                "stem": f'阅读"{sent[:50]}..."，请用简洁的语言概括其表达的核心观点（不超过100字）。',
                "answer": f"核心观点：{sent[:150]}",
                "key_points": [
                    "准确提取核心观点",
                    "表述清晰简洁",
                    "不超过字数限制",
                    "不遗漏关键信息"
                ],
                "difficulty": "easy"
            }
        ]

        tmpl = templates[i % len(templates)]
        questions.append({
            "id": len(questions) + 1,
            "stem": tmpl["stem"],
            "answer": tmpl["answer"],
            "key_points": tmpl["key_points"],
            "difficulty": tmpl["difficulty"]
        })

        if len(questions) >= count:
            break

    return questions


def generate_fill_local(text, count):
    """本地生成填空题"""
    sents = _split_sentences(text)
    keywords = _extract_keywords(text, 20)

    questions = []
    for i in range(count):
        sent = sents[i % len(sents)]
        kw = keywords[i % len(keywords)] if keywords else "___"
        # 用关键词填空
        blank_sent = sent.replace(kw, "___", 1) if kw in sent else sent[:40] + "___" + sent[40:80]
        questions.append({
            "id": len(questions) + 1,
            "stem": blank_sent[:150],
            "answer": kw,
            "analysis": f'根据原文语境，此处应填入"{kw}"。'
        })
        if len(questions) >= count:
            break
    return questions


def generate_tf_local(text, count):
    """本地生成判断题（count=0 时自动使用全部句子）"""
    sents = _split_sentences(text)
    if not sents:
        return []

    keywords = _extract_keywords(text, 10)

    # count=0 表示使用全部句子，但最多50题
    max_q = min(len(sents), 50) if count == 0 else min(count, len(sents))

    questions = []
    for i in range(max_q):
        sent = sents[i % len(sents)]
        kw = keywords[i % len(keywords)] if keywords else "核心要素"
        is_true = i % 2 == 0

        if is_true:
            stem = f'{kw}{"" if kw in sent else "在原文内容中"}可以找到相关依据。'
            answer = "正确"
            analysis = f'原文"{sent[:60]}..."证实了这一点。'
        else:
            opposite_kw = keywords[(i + 1) % len(keywords)] if len(keywords) > 1 else "无关因素"
            stem = f'{opposite_kw}是影响{kw}的唯一决定因素。'
            answer = "错误"
            analysis = f'原文表明{kw}受多种因素影响，并非仅由{opposite_kw}决定。'

        questions.append({
            "id": len(questions) + 1,
            "stem": stem,
            "answer": answer,
            "analysis": analysis
        })
    return questions


def generate_questions_local(text, question_type, question_count=5):
    """纯本地模式题目生成入口。
    question_type='all' 时同时生成四种题型；
    question_count=0 时自动按文本长度估算。"""
    generators = {
        "choice": generate_choice_local,
        "tf": generate_tf_local,
        "fill": generate_fill_local,
        "essay": generate_essay_local,
    }

    if question_type == 'all':
        # 自动估算每类题数
        if question_count == 0:
            per_type = max(3, min(15, estimate_question_count(text, max_count=50) // 4))
        else:
            per_type = max(2, question_count // 4)
        grouped = {}
        total = 0
        for t in ['choice', 'tf', 'fill', 'essay']:
            arr = generators[t](text, per_type)
            grouped[t] = arr
            total += len(arr)
        return {
            "error": None,
            "grouped": grouped,
            "total": total,
            "provider_used": "本地模式（无需API）",
            "notice": "使用本地规则生成，题目质量有限。建议配置AI API Key以获得更好的效果。"
        }

    gen = generators.get(question_type, generate_choice_local)
    questions = gen(text, question_count)
    return {
        "error": None,
        "questions": questions,
        "provider_used": "本地模式（无需API）",
        "notice": "使用本地规则生成，题目质量有限。建议配置AI API Key以获得更好的效果。"
    }



# ==================== 登录/注册/登出路由 ====================
@app.route('/login')
def login_page():
    """登录页面"""
    if session.get('logged_in'):
        return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/api/register', methods=['POST'])
def api_register():
    """注册 API — 创建新用户，自动登录"""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "请输入注册信息"}), 400

    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    display_name = data.get('display_name', '').strip() or username

    # 验证用户名：3-20 位字母/数字/中文/下划线
    if not re.match(r'^[\w\u4e00-\u9fa5]{3,20}$', username):
        return jsonify({"error": "用户名需3-20位，支持字母/数字/中文/下划线"}), 400
    if len(password) < 4:
        return jsonify({"error": "密码至少需要4位"}), 400

    if _USE_SUPABASE:
        if db.user_exists(username):
            return jsonify({"error": "该用户名已被注册"}), 409
        user_data = db.create_user(username, password, display_name)
        token = user_data['remember_token']
    else:
        global _users_db
        _users_db = _load_users()

        if username.lower() in {k.lower() for k in _users_db}:
            return jsonify({"error": "该用户名已被注册"}), 409

        # 创建用户
        token = _generate_token()
        _users_db[username] = {
            'password': _hash_pw(password),
            'display_name': display_name,
            'created_at': datetime.now().isoformat(),
            'remember_token': token
        }
        _save_users(_users_db)

    # 自动登录
    session.clear()
    session['logged_in'] = True
    session['username'] = username
    session['display_name'] = display_name

    # 设置 remember-me cookie
    resp = make_response(jsonify({
        "success": True,
        "message": f"注册成功，欢迎 {display_name}！",
        "username": username,
        "display_name": display_name
    }))
    resp.set_cookie(
        REMEMBER_COOKIE_NAME, token,
        max_age=REMEMBER_DAYS * 86400,
        httponly=True,
        samesite='Lax'
    )
    return resp


@app.route('/api/login', methods=['POST'])
def api_login():
    """登录 API — 支持自动登录"""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "请输入用户名和密码"}), 400

    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    remember = data.get('remember', True)  # 默认开启自动登录

    if not username or not password:
        return jsonify({"error": "用户名和密码不能为空"}), 400

    if _verify_password(username, password):
        session.clear()
        session['logged_in'] = True
        session['username'] = username

        if _USE_SUPABASE:
            user_data = db.get_user(username)
            display_name = user_data.get('display_name', username) if user_data else username
        else:
            global _users_db
            _users_db = _load_users()
            display_name = _users_db.get(username, {}).get('display_name', username)

        session['display_name'] = display_name

        resp = make_response(jsonify({
            "success": True,
            "message": f"登录成功，欢迎回来！",
            "username": username,
            "display_name": display_name
        }))

        if remember:
            token = _generate_token()
            if _USE_SUPABASE:
                db.update_remember_token(username, token)
            else:
                _users_db[username]['remember_token'] = token
                _save_users(_users_db)
            resp.set_cookie(
                REMEMBER_COOKIE_NAME, token,
                max_age=REMEMBER_DAYS * 86400,
                httponly=True,
                samesite='Lax'
            )
        else:
            if _USE_SUPABASE:
                db.update_remember_token(username, None)
            else:
                _users_db[username]['remember_token'] = None
                _save_users(_users_db)
            resp.set_cookie(REMEMBER_COOKIE_NAME, '', max_age=0)

        return resp
    else:
        return jsonify({"error": "用户名或密码错误"}), 401


@app.route('/api/logout', methods=['POST'])
def api_logout():
    """登出 — 清除 cookie 和 token"""
    username = session.get('username')
    if username:
        if _USE_SUPABASE:
            db.update_remember_token(username, None)
        else:
            global _users_db
            _users_db = _load_users()
            if username in _users_db:
                _users_db[username]['remember_token'] = None
                _save_users(_users_db)
    session.clear()
    resp = make_response(jsonify({"success": True}))
    resp.set_cookie(REMEMBER_COOKIE_NAME, '', max_age=0)
    return resp


# ==================== API 路由 ====================
@app.after_request
def add_utf8_charset(response):
    """确保所有响应都有正确的 UTF-8 编码声明"""
    content_type = response.headers.get('Content-Type', '')
    if content_type and 'charset' not in content_type:
        if 'text/html' in content_type:
            response.headers['Content-Type'] = 'text/html; charset=utf-8'
        elif 'application/json' in content_type:
            response.headers['Content-Type'] = 'application/json; charset=utf-8'
        elif 'text/plain' in content_type:
            response.headers['Content-Type'] = 'text/plain; charset=utf-8'
    return response


@app.route('/')
@login_required
def index():
    return render_template('index.html',
        username=session.get('username', ''),
        display_name=session.get('display_name', ''))


@app.route('/api/providers', methods=['GET'])
@login_required
def list_providers():
    """返回所有大模型提供商信息"""
    result = {}
    for pid, pdata in PROVIDERS.items():
        result[pid] = {
            "name": pdata["name"],
            "default_model": pdata["default_model"],
            "models": pdata["models"],
            "base_url": pdata["base_url"],
            "desc": pdata["desc"],
            "sdk_type": pdata["sdk_type"]
        }
    return jsonify(result)


@app.route('/api/upload', methods=['POST'])
@login_required
def upload_file():
    """统一的文件上传接口，支持 multipart/form-data 和 JSON+base64 两种方式"""
    try:
        filename = None
        filepath = None
        user_uploads = g.user_uploads

        # ---- 方式1: 传统 multipart/form-data ----
        if 'file' in request.files:
            file = request.files['file']
            if not file.filename:
                return jsonify({"error": "文件名为空"}), 400
            ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
            if ext not in ALLOWED_EXTENSIONS:
                return jsonify({"error": f"不支持 .{ext}，支持: {', '.join(ALLOWED_EXTENSIONS)}"}), 400
            filename = safe_filename(file.filename)
            filepath = user_uploads / filename
            file.seek(0, os.SEEK_END)
            if file.tell() > MAX_FILE_SIZE:
                return jsonify({"error": "文件超过 10MB"}), 400
            file.seek(0)
            file.save(str(filepath))

        # ---- 方式2: JSON + base64（Netlify / serverless 兼容）----
        elif request.is_json:
            data = request.get_json(silent=True)
            if data and data.get('file_base64') and data.get('filename'):
                filename = safe_filename(data['filename'])
                ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
                if ext not in ALLOWED_EXTENSIONS:
                    return jsonify({"error": f"不支持 .{ext}"}), 400
                try:
                    file_bytes = base64.b64decode(data['file_base64'])
                except Exception:
                    return jsonify({"error": "Base64 解码失败"}), 400
                if len(file_bytes) > MAX_FILE_SIZE:
                    return jsonify({"error": "文件超过 10MB"}), 400
                filepath = user_uploads / filename
                with open(str(filepath), 'wb') as f:
                    f.write(file_bytes)
            else:
                return jsonify({"error": "JSON 数据缺少 file_base64 或 filename"}), 400
        else:
            return jsonify({"error": "没有上传文件"}), 400

        # ---- 统一文本提取和返回 ----
        if not filepath or not filename:
            return jsonify({"error": "文件保存失败"}), 500

        full_text = extract_text(str(filepath), filename)
        if full_text.startswith("[错误]"):
            return jsonify({"error": full_text.strip("[]错误 ").strip()}), 400
        if full_text.startswith("[提示]"):
            return jsonify({"error": full_text.strip("[]提示 ").strip()}), 400
        # 规范化文本，统一全角/半角字符
        full_text = normalize_text(full_text)
        truncated = len(full_text) > 6000
        display_text = full_text[:6000] + ("\n\n...[文本过长已截断]..." if truncated else "")

        # ---- 检测文件中是否包含已格式化的题目 ----
        has_questions, detected_count, detected_type = detect_existing_questions(full_text)
        parsed_questions = None
        if has_questions:
            parsed = parse_questions_from_text(full_text, detected_type)
            if parsed:
                parsed_questions = {
                    "type": detected_type,
                    "questions": parsed,
                    "count": len(parsed)
                }

        return jsonify({
            "filename": filename,
            "text": display_text,
            "full_text": full_text,
            "full_length": len(full_text),
            "truncated": truncated,
            "has_existing_questions": has_questions,
            "parsed_questions": parsed_questions
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"服务器处理上传失败: {str(e)}"}), 500


@app.route('/api/generate', methods=['POST'])
@login_required
def generate_questions():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "请求体不是有效的 JSON"}), 400

        text = data.get('text', '')
        question_type = data.get('type', 'choice')
        question_count = int(data.get('count', 5))
        mode = data.get('mode', 'local')  # 'ai' or 'local'

        # 支持全部题型与单类题型
        valid_types = ('choice', 'tf', 'fill', 'essay', 'all')
        if question_type not in valid_types:
            question_type = 'all'

        if not text.strip():
            return jsonify({"error": "文本内容为空"}), 400

        # 规范化文本
        text = normalize_text(text)

        # 优先识别文档中已有的题目，避免重新生成导致答案不一致
        has_existing, _, detected_type = detect_existing_questions(text)
        if has_existing:
            parsed = parse_questions_from_text(text, detected_type)
            if parsed:
                return jsonify({
                    "error": None,
                    "questions": parsed,
                    "total": len(parsed),
                    "provider_used": "原文档解析（保持原答案）",
                    "notice": "已检测到文件中包含现成题目，直接解析并保留原答案。"
                })

        # count=0 表示自动提取全部题目：估算合理数量
        is_all_mode = (question_count == 0)
        if is_all_mode:
            question_count = estimate_question_count(text, max_count=50)
        else:
            question_count = min(question_count, 50)

        if mode == 'ai':
            provider_id = data.get('provider', 'openai')
            api_key = data.get('api_key', '')
            api_base = data.get('api_base', '')
            model = data.get('model', '')

            provider = PROVIDERS.get(provider_id, PROVIDERS['custom'])
            if not api_base:
                api_base = provider.get('base_url', '')
            if not model:
                model = provider.get('default_model', '')

            result = call_ai_api(provider_id, api_key, api_base, model,
                                 text, question_type, question_count)

            # 如果 AI 失败，回退本地（all 模式也支持）
            if result.get('error'):
                local_result = generate_questions_local(text, question_type, question_count)
                local_result['notice'] = f"AI调用失败({result['error'][:80]})，已自动切换本地模式"
                return jsonify(local_result)

            return jsonify(result)
        else:
            result = generate_questions_local(text, question_type, question_count)
            return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"生成题目失败: {str(e)}"}), 500





# ==================== 题库管理 API（支持文档空间） ====================
@app.route('/api/bank', methods=['GET'])
@login_required
def get_bank():
    """获取题库摘要（所有空间概览）"""
    if _USE_SUPABASE:
        username = session.get('username', '')
        spaces_summary = db.get_spaces_summary(username)
        return jsonify({
            "spaces": spaces_summary,
            "total_spaces": len(spaces_summary),
            "total_questions": sum(s.get("question_count", 0) for s in spaces_summary),
            "questions": [],
        })
    bank = _load_bank()
    spaces_summary = []
    for sid, sdata in bank.get("spaces", {}).items():
        wrong_book = sdata.get("wrong_book", {})
        spaces_summary.append({
            "space_id": sid,
            "name": sdata.get("name", sid),
            "created_at": sdata.get("created_at", ""),
            "question_count": len(sdata.get("questions", [])),
            "question_types": list(set(q.get("type", "choice") for q in sdata.get("questions", []))),
            "wrong_count": len(wrong_book),
        })
    spaces_summary.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return jsonify({
        "spaces": spaces_summary,
        "total_spaces": len(spaces_summary),
        "total_questions": sum(s.get("question_count", 0) for s in spaces_summary),
        "questions": bank.get("questions", []),
    })


@app.route('/api/bank/spaces', methods=['GET'])
@login_required
def list_spaces():
    """列出所有文档空间"""
    if _USE_SUPABASE:
        username = session.get('username', '')
        spaces = db.list_spaces(username)
        result = []
        for s in spaces:
            questions = db.get_space_questions(s['id'])
            result.append({
                "space_id": s['id'],
                "name": s.get('name', s['id']),
                "created_at": s.get('created_at', ''),
                "question_count": len(questions),
            })
        return jsonify(result)
    bank = _load_bank()
    result = []
    for sid, sdata in bank.get("spaces", {}).items():
        result.append({
            "space_id": sid,
            "name": sdata.get("name", sid),
            "created_at": sdata.get("created_at", ""),
            "question_count": len(sdata.get("questions", [])),
        })
    result.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return jsonify(result)


@app.route('/api/bank/space/<space_id>', methods=['GET'])
@login_required
def get_space(space_id):
    """获取指定空间的题目"""
    if _USE_SUPABASE:
        space = db.get_space(space_id)
        if not space:
            return jsonify({"error": "空间未找到"}), 404
        questions = db.get_space_questions(space_id)
        # 将 JSON 字符串转回 Python 对象
        for q in questions:
            if isinstance(q.get('options'), str):
                try:
                    q['options'] = json.loads(q['options'])
                except Exception:
                    q['options'] = []
        return jsonify({
            "space_id": space_id,
            "name": space.get("name", space_id),
            "created_at": space.get("created_at", ""),
            "questions": questions,
            "total": len(questions)
        })
    bank = _load_bank()
    space = bank.get("spaces", {}).get(space_id)
    if not space:
        return jsonify({"error": "空间未找到"}), 404
    return jsonify({
        "space_id": space_id,
        "name": space.get("name", space_id),
        "created_at": space.get("created_at", ""),
        "questions": space.get("questions", []),
        "total": len(space.get("questions", []))
    })


@app.route('/api/bank/save-batch', methods=['POST'])
@login_required
def save_batch_to_bank():
    """批量保存题目到指定文档空间"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请求数据为空"}), 400

    questions = data.get('questions', [])
    space_name = data.get('space_name', '未命名文档')
    space_id = data.get('space_id', None)
    source_text = data.get('source_text', '')
    username = session.get('username', '')

    if not questions or not isinstance(questions, list):
        return jsonify({"error": "没有题目数据"}), 400

    if _USE_SUPABASE:
        if not space_id:
            # 检查同名空间是否存在
            spaces = db.list_spaces(username)
            for s in spaces:
                if s.get('name') == space_name:
                    space_id = s['id']
                    break
            if not space_id:
                space_data = db.create_space(username, space_name, source_text)
                space_id = space_data['id']

        saved_count = db.save_questions_to_space(space_id, questions)
        total = len(db.get_space_questions(space_id))
        return jsonify({
            "success": True,
            "space_id": space_id,
            "space_name": space_name,
            "added": saved_count,
            "total_in_space": total
        })

    bank = _load_bank()

    if not space_id:
        space_id = _generate_space_id(space_name)

    existing = bank.get("spaces", {}).get(space_id, None)
    if existing:
        start_id = existing.get("next_id", len(existing.get("questions", [])) + 1)
        for q in questions:
            q["bank_id"] = start_id
            start_id += 1
            existing["questions"].append(q)
        existing["next_id"] = start_id
        if source_text:
            existing["source_text"] = source_text
    else:
        start_id = 1
        for q in questions:
            q["bank_id"] = start_id
            start_id += 1
        bank["spaces"][space_id] = {
            "name": space_name,
            "created_at": datetime.now().isoformat(),
            "questions": list(questions),
            "next_id": start_id,
            "source_text": source_text
        }

    _save_bank(bank)
    return jsonify({
        "success": True,
        "space_id": space_id,
        "space_name": space_name,
        "added": len(questions),
        "total_in_space": len(bank["spaces"][space_id]["questions"])
    })


@app.route('/api/bank', methods=['POST'])
@login_required
def add_to_bank():
    """添加单个/批量题目到题库（兼容旧接口，支持可选空间参数）"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请求数据为空"}), 400

    questions = data.get('questions', [])
    if not isinstance(questions, list):
        questions = [data]
    if not questions:
        return jsonify({"error": "没有题目数据"}), 400

    space_id = data.get('space_id') or data.get('space_name', '未命名文档')
    space_name = data.get('space_name', '未命名文档')
    source_text = data.get('source_text', '')
    username = session.get('username', '')

    if _USE_SUPABASE:
        # 如果 space_id 不像 ID，按名称查找
        if not re.search(r'\d{14}$', space_id or ''):
            spaces = db.list_spaces(username)
            for s in spaces:
                if s.get('name') == space_name:
                    space_id = s['id']
                    break
            else:
                space_id = None

        if not space_id:
            space_data = db.create_space(username, space_name, source_text)
            space_id = space_data['id']
        elif not db.space_exists(space_id):
            space_data = db.create_space(username, space_name, source_text)
            space_id = space_data['id']

        saved_count = db.save_questions_to_space(space_id, questions)
        total = len(db.get_space_questions(space_id))
        return jsonify({
            "success": True,
            "space_id": space_id,
            "space_name": space_name,
            "added": saved_count,
            "total": total,
            "total_in_space": total
        })

    # 如果 space_id 看起来不像 ID，则用文档名生成
    if not re.search(r'\d{14}$', space_id or ''):
        space_id = None

    bank = _load_bank()

    if not space_id:
        space_id = _generate_space_id(space_name)

    existing = bank.get("spaces", {}).get(space_id)
    if existing:
        start_id = existing.get("next_id", len(existing.get("questions", [])) + 1)
        for q in questions:
            q["bank_id"] = start_id
            start_id += 1
            existing["questions"].append(q)
        existing["next_id"] = start_id
        if source_text:
            existing["source_text"] = source_text
    else:
        start_id = 1
        for q in questions:
            q["bank_id"] = start_id
            start_id += 1
        bank["spaces"][space_id] = {
            "name": space_name,
            "created_at": datetime.now().isoformat(),
            "questions": list(questions),
            "next_id": start_id,
            "source_text": source_text
        }

    _save_bank(bank)
    total = len(bank["spaces"][space_id]["questions"])
    return jsonify({
        "success": True,
        "space_id": space_id,
        "space_name": space_name,
        "added": len(questions),
        "total": total,
        "total_in_space": total
    })


@app.route('/api/bank/space/<space_id>', methods=['PUT'])
@login_required
def rename_space(space_id):
    """重命名文档空间"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请求数据为空"}), 400

    new_name = data.get('name', '').strip()
    if not new_name:
        return jsonify({"error": "空间名称不能为空"}), 400

    if _USE_SUPABASE:
        if not db.space_exists(space_id):
            return jsonify({"error": "空间未找到"}), 404
        db.update_space_name(space_id, new_name)
        return jsonify({"success": True, "space_id": space_id, "name": new_name})

    bank = _load_bank()
    space = bank.get("spaces", {}).get(space_id)
    if not space:
        return jsonify({"error": "空间未找到"}), 404
    space["name"] = new_name
    _save_bank(bank)
    return jsonify({"success": True, "space_id": space_id, "name": new_name})


@app.route('/api/bank/space/<space_id>', methods=['DELETE'])
@login_required
def delete_space(space_id):
    """删除整个文档空间"""
    if _USE_SUPABASE:
        if not db.space_exists(space_id):
            return jsonify({"error": "空间未找到"}), 404
        db.delete_space(space_id)
        return jsonify({"success": True})

    bank = _load_bank()
    if space_id not in bank.get("spaces", {}):
        return jsonify({"error": "空间未找到"}), 404
    del bank["spaces"][space_id]
    _save_bank(bank)
    return jsonify({"success": True})


# ==================== 空间错题本 API ====================
@app.route('/api/bank/space/<space_id>/wrong', methods=['POST'])
@login_required
def record_wrong_answers(space_id):
    """记录某个空间内答错的题目"""
    if _USE_SUPABASE:
        if not db.space_exists(space_id):
            return jsonify({"error": "空间未找到"}), 404
        data = request.get_json()
        wrong_list = data.get('wrong_questions', [])
        if not wrong_list:
            return jsonify({"error": "没有错题数据"}), 400
        db.record_wrong_answers(space_id, wrong_list)
        return jsonify({
            "success": True,
            "wrong_count": len(db.get_wrong_questions(space_id)),
            "space_id": space_id
        })

    bank = _load_bank()
    space = bank.get("spaces", {}).get(space_id)
    if not space:
        return jsonify({"error": "空间未找到"}), 404

    data = request.get_json()
    wrong_list = data.get('wrong_questions', [])
    if not wrong_list:
        return jsonify({"error": "没有错题数据"}), 400

    if "wrong_book" not in space:
        space["wrong_book"] = {}

    for wq in wrong_list:
        bank_id = str(wq.get("bank_id", ""))
        if bank_id:
            space["wrong_book"][bank_id] = {
                "bank_id": wq["bank_id"],
                "stem": wq.get("stem", ""),
                "type": wq.get("type", "choice"),
                "options": wq.get("options", []),
                "answer": wq.get("answer", ""),
                "analysis": wq.get("analysis", ""),
                "user_answer": wq.get("user_answer", ""),
                "recorded_at": datetime.now().isoformat()
            }

    _save_bank(bank)
    return jsonify({
        "success": True,
        "wrong_count": len(space["wrong_book"]),
        "space_id": space_id
    })


@app.route('/api/bank/space/<space_id>/wrong', methods=['GET'])
@login_required
def get_wrong_questions(space_id):
    """获取某空间的错题本"""
    if _USE_SUPABASE:
        space = db.get_space(space_id)
        if not space:
            return jsonify({"error": "空间未找到"}), 404
        wrong_list = db.get_wrong_questions(space_id)
        for q in wrong_list:
            if isinstance(q.get('options'), str):
                try:
                    q['options'] = json.loads(q['options'])
                except Exception:
                    q['options'] = []
        return jsonify({
            "space_id": space_id,
            "space_name": space.get("name", space_id),
            "wrong_questions": wrong_list,
            "wrong_count": len(wrong_list)
        })

    bank = _load_bank()
    space = bank.get("spaces", {}).get(space_id)
    if not space:
        return jsonify({"error": "空间未找到"}), 404

    wrong_book = space.get("wrong_book", {})
    wrong_list = list(wrong_book.values())
    wrong_list.sort(key=lambda x: x.get("recorded_at", ""), reverse=True)

    return jsonify({
        "space_id": space_id,
        "space_name": space.get("name", space_id),
        "wrong_questions": wrong_list,
        "wrong_count": len(wrong_list)
    })


@app.route('/api/bank/space/<space_id>/wrong/clear', methods=['POST'])
@login_required
def clear_wrong_questions(space_id):
    """清空某空间的错题本"""
    if _USE_SUPABASE:
        if not db.space_exists(space_id):
            return jsonify({"error": "空间未找到"}), 404
        db.clear_wrong_questions(space_id)
        return jsonify({"success": True, "space_id": space_id})

    bank = _load_bank()
    space = bank.get("spaces", {}).get(space_id)
    if not space:
        return jsonify({"error": "空间未找到"}), 404

    space["wrong_book"] = {}
    _save_bank(bank)
    return jsonify({"success": True, "space_id": space_id})


@app.route('/api/bank/<int:bank_id>', methods=['DELETE'])
@login_required
def delete_from_bank(bank_id):
    """从题库中删除指定题目"""
    if _USE_SUPABASE:
        space_id = db.delete_question_by_bank_id(bank_id)
        if not space_id:
            return jsonify({"error": "题目未找到"}), 404
        return jsonify({"success": True, "deleted_from": space_id})

    bank = _load_bank()
    for sid, sdata in bank.get("spaces", {}).items():
        before = len(sdata["questions"])
        sdata["questions"] = [q for q in sdata["questions"] if q.get("bank_id") != bank_id]
        if len(sdata["questions"]) < before:
            if not sdata["questions"]:
                del bank["spaces"][sid]
            _save_bank(bank)
            return jsonify({"success": True, "deleted_from": sid})
    return jsonify({"error": "题目未找到"}), 404


@app.route('/api/bank/clear', methods=['POST'])
@login_required
def clear_bank():
    """清空所有题库空间"""
    if _USE_SUPABASE:
        username = session.get('username', '')
        db.clear_all_spaces(username)
        return jsonify({"success": True, "total": 0})
    _save_bank(_new_bank())
    return jsonify({"success": True, "total": 0})


@app.route('/api/config', methods=['GET'])
@login_required
def get_config():
    return jsonify({
        "has_docx": HAS_DOCX,
        "has_pil": HAS_PIL,
        "has_tesseract": HAS_TESSERACT,
        "has_pdfplumber": HAS_PDFPLUMBER,
        "has_pypdf2": HAS_PYPDF2,
        "has_openpyxl": HAS_OPENPYXL,
        "has_pptx": HAS_PPTX,
        "has_openai_sdk": HAS_OPENAI_SDK,
        "has_gemini": HAS_GEMINI,
        "has_anthropic": HAS_ANTHROPIC,
        "has_ddg": HAS_DDG,
        "has_bs4": HAS_BS4,
        "has_requests": HAS_REQUESTS,
        "allowed_types": list(ALLOWED_EXTENSIONS),
        "max_file_size_mb": MAX_FILE_SIZE // (1024 * 1024),
    })


# ==================== 网络搜题 API ====================
def _extract_search_keywords(text, topn=8):
    """从文档中提取搜索关键词，返回搜索词组列表"""
    import re as _re
    if not text:
        return []

    # 清理文本：移除题型标记、题号、选项标号等
    cleaned = text
    # 移除 单选题/多选题/判断题/填空题/简答题 等
    cleaned = _re.sub(r'（?\s*(单选题|多选题|判断题|填空题|简答题|不定项|选择题|判断|单选|多选)\s*）?', ' ', cleaned)
    cleaned = _re.sub(r'[（(]\s*[A-Za-z]\s*[)）]', ' ', cleaned)
    cleaned = _re.sub(r'第\s*\d+\s*题', ' ', cleaned)
    cleaned = _re.sub(r'\d+[\.、．\)）]', ' ', cleaned)

    # 提取中文词组（2-10字），优先保留长词
    cn_words = _re.findall(r'[\u4e00-\u9fa5]{2,10}', cleaned)
    # 提取英文词组
    en_words = _re.findall(r'[a-zA-Z]{3,}', text)

    stopwords = {
        '这是', '不是', '一个', '可以', '进行', '使用', '没有', '我们', '他们', '你们',
        '这个', '那个', '什么', '因为', '所以', '而且', '但是', '如果', '就是', '只是',
        '已经', '还是', '或者', '不过', '然后', '其中', '因此', '然而', '虽然', '尽管',
        '这些', '那些', '一些', '一般', '一种', '可能', '需要', '应该', '能够', '以及',
        '关于', '对于', '由于', '根据', '按照', '随着', '通过', '经过', '属于', '成为',
        '表示', '说明', '认为', '指出', '提出', '强调', '要求', '坚持', '推进', '推动',
        '加强', '完善', '发展', '实现', '建设', '建立', '形成', '提高', '促进', '增强',
        '第一', '第二', '第三', '第四', '第五', '第六', '第七', '第八', '第九', '第十',
        '下列', '以下', '以下关于', '关于下列', '正确', '错误', '正确的是', '错误的是',
        '以上', '其中', '下列关于', '本题', '答案', '解析', '题干', '选项', '题目'
    }

    # 题型词过滤
    type_words = {'单选题', '多选题', '判断题', '填空题', '简答题', '选择题', '判断', '单选', '多选', '不定项', '单选多选'}

    freq = {}
    for w in cn_words + en_words:
        if w in type_words:
            continue
        if w in stopwords:
            continue
        freq[w] = freq.get(w, 0) + 1

    # 合并相近词：如果一个长词包含另一个短词，给长词加分
    sorted_words = sorted(freq.items(), key=lambda x: x[1], reverse=True)
    keywords = [w for w, _ in sorted_words[:topn]]

    # 优先保留长词：如果包含"习近平新时代中国特色社会主义思想"这样的长词，保留
    long_terms = []
    for m in _re.finditer(r'[\u4e00-\u9fa5]{8,20}', cleaned):
        term = m.group(0)
        if term not in stopwords and len(term) >= 8:
            long_terms.append(term)
    # 按出现频率取前几个长词
    long_freq = {}
    for t in long_terms:
        long_freq[t] = long_freq.get(t, 0) + 1
    top_long = sorted(long_freq.items(), key=lambda x: x[1], reverse=True)[:3]

    # 构建搜索查询：组合关键词 + 题型后缀
    queries = []
    # 把长专有名词直接加入关键词列表
    combined_keywords = [t for t, _ in top_long] + keywords
    combined_keywords = list(dict.fromkeys(combined_keywords))  # 去重保持顺序
    combined_keywords = combined_keywords[:6]  # 限制数量

    if len(combined_keywords) >= 3:
        queries.append('"' + '" "'.join(combined_keywords[:3]) + '" 考试题 题库')
    if len(combined_keywords) >= 2:
        queries.append('"' + '" "'.join(combined_keywords[:2]) + '" 试题 答案')
    if combined_keywords:
        queries.append('"' + combined_keywords[0] + '" 选择题 判断题')
    # 额外补充一个开放查询
    if len(combined_keywords) >= 2:
        queries.append(' '.join(combined_keywords[:3]) + ' 题库 网课')
    return queries


def _search_web(query, max_results=10):
    """使用 DuckDuckGo 搜索"""
    if not HAS_DDG:
        return []
    try:
        with DDGS() as ddgs:
            results = []
            for r in ddgs.text(query, max_results=max_results, timeout=10):
                results.append({
                    'title': r.get('title', ''),
                    'url': r.get('href', ''),
                    'snippet': r.get('body', '')
                })
            return results
    except Exception as e:
        print(f"[搜索异常] {query}: {e}")
        return []


def _fetch_page_text(url, timeout=5):
    """获取网页文本内容"""
    if not HAS_REQUESTS or not HAS_BS4:
        return ''
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        resp = http_requests.get(url, headers=headers, timeout=timeout)
        resp.encoding = resp.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')
        # 移除 script/style
        for tag in soup(['script', 'style', 'nav', 'footer', 'header']):
            tag.decompose()
        text = soup.get_text(separator='\n')
        # 清理多余空白
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        return '\n'.join(lines[:300])  # 限制长度
    except Exception:
        return ''


def _extract_questions_from_web_text(text, keyword_set):
    """从网页文本中尝试提取题目"""
    if not text:
        return []

    questions_found = []
    # 按题号分割
    blocks = re.split(
        r'(?:^|\n)\s*(?:\d+[\.、．\)）]\s*|'
        r'[(（]\s*\d+\s*[)）]\s*|'
        r'第\s*\d+\s*题\s*)',
        text
    )

    for block in blocks[1:]:
        block = block.strip()
        if len(block) < 10:
            continue

        # 检测是否有选项标记
        has_options = bool(re.search(r'[A-E][\.、．\)）]', block))

        # 检测是否包含关键词
        has_keyword = any(kw in block for kw in keyword_set if len(kw) >= 2)
        if not has_keyword and keyword_set:
            continue

        # 提取答案
        answer = ''
        ans_m = re.search(
            r'(?:答案|正确答案)[：:是为]?\s*([A-Ea-e]+|正确|错误|对|错)',
            block, re.IGNORECASE
        )
        if ans_m:
            answer = ans_m.group(1).upper()
            if answer in ('正确', '对', 'TRUE', 'T'):
                answer = '正确'
            elif answer in ('错误', '错', 'FALSE', 'F'):
                answer = '错误'

        # 提取选项
        options = []
        for m in re.finditer(r'([A-E])[\.、．\)）]\s*(.+?)(?=\s*[A-E][\.、．\)）]|\s*答案|\s*解析|\s*$)', block):
            options.append(f"{m.group(1)}. {m.group(2).strip()[:80]}")

        if has_options and options:
            qtype = 'multi' if len(options) >= 5 else 'choice'
        else:
            qtype = 'tf'
            if not options:
                options = ['正确. 正确', '错误. 错误']

        # 题干 = 第一个选项之前的内容
        if options and re.search(r'[A-E][\.、．\)）]', block):
            stem_end = block.find(options[0].split('. ')[0] if '. ' in options[0] else '')
            stem = block[:stem_end].strip() if stem_end > 0 else block[:60]
        else:
            stem = block[:80]

        # 清理题干中的答案标记
        stem = re.sub(r'(?:答案|正确答案)[：:是为]?\s*[A-Ea-e]+\s*', '', stem)
        stem = re.sub(r'[（(]\s*(?:正确|错误|对|错)\s*[)）]', '', stem)
        stem = re.sub(r'\s+', ' ', stem).strip()[:120]

        if len(stem) >= 6:
            questions_found.append({
                'stem': stem,
                'type': qtype,
                'options': options[:6],
                'answer': answer or ('A' if options else '正确'),
                'analysis': '来源：网络搜索结果'
            })

        if len(questions_found) >= 15:
            break

    return questions_found


@app.route('/api/search-questions', methods=['POST'])
@login_required
def search_questions():
    """网络搜题：根据文档内容搜索相关题目"""
    if not HAS_DDG:
        return jsonify({
            "error": "需要安装 duckduckgo-search: pip install duckduckgo-search"
        }), 400

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "请求数据为空"}), 400

    space_id = data.get('space_id', '')
    search_text = data.get('text', '')

    # 获取搜索文本：优先使用空间保存的原文，其次传入的 text，最后从题目题干提取
    if not search_text and space_id:
        bank = _load_bank()
        space = bank.get('spaces', {}).get(space_id)
        if space:
            search_text = space.get('source_text', '')
            if not search_text:
                questions = space.get('questions', [])
                stems = ' '.join(q.get('stem', '') for q in questions[:10])
                search_text = stems

    if not search_text or len(search_text) < 20:
        return jsonify({"error": "文本内容太少，无法提取关键词搜索。请重新上传文档并确保能识别出文字内容。"}), 400

    # 提取关键词并构建搜索查询
    queries = _extract_search_keywords(search_text)
    keywords_raw = _extract_keywords(search_text, 6)

    # 搜索
    all_results = []
    seen_urls = set()
    search_errors = []

    for query in queries[:3]:  # 最多3组搜索词
        try:
            results = _search_web(query, max_results=6)
            if results:
                for r in results:
                    if r['url'] not in seen_urls:
                        seen_urls.add(r['url'])
                        all_results.append(r)
            else:
                search_errors.append(f"{query} (无结果)")
        except Exception as e:
            search_errors.append(str(e)[:80])

    # 尝试从搜索结果页面提取题目
    extracted_questions = []
    fetched_pages = 0

    for result in all_results[:6]:  # 抓取前6个结果
        if fetched_pages >= 4:
            break
        try:
            page_text = _fetch_page_text(result['url'], timeout=5)
            if page_text:
                fetched_pages += 1
                qs = _extract_questions_from_web_text(page_text, set(keywords_raw))
                for q in qs:
                    q['source_url'] = result['url']
                    q['source_title'] = result['title'][:50]
                extracted_questions.extend(qs)
        except Exception:
            pass

    # 去重题干
    seen_stems = set()
    unique_questions = []
    for q in extracted_questions:
        if q['stem'] not in seen_stems:
            seen_stems.add(q['stem'])
            unique_questions.append(q)

    # 如果完全没有搜索结果，返回更明确的提示
    if not all_results and not unique_questions:
        error_msg = "未找到相关网页结果。"
        if search_errors:
            error_msg += f" 搜索异常：{'；'.join(search_errors[:3])}"
        return jsonify({
            "success": True,
            "search_queries": queries,
            "search_results": [],
            "extracted_questions": [],
            "total_web_results": 0,
            "total_extracted": 0,
            "searched_pages": 0,
            "mode": "网络搜索",
            "error": error_msg
        }), 200

    return jsonify({
        "success": True,
        "search_queries": queries,
        "search_results": all_results[:8],
        "extracted_questions": unique_questions[:30],
        "total_web_results": len(all_results),
        "total_extracted": len(unique_questions),
        "searched_pages": fetched_pages,
        "mode": "网络搜索"
    })


if __name__ == '__main__':
    _OS_INFO = f"Windows {_platform.release()}" if _IS_WINDOWS else _platform.platform()
    print("=" * 60)
    print("  文件上传 → 题目生成器 v2.0")
    print(f"  运行平台: {_OS_INFO}")
    print("  访问: http://127.0.0.1:5000")
    print("  默认账号: admin / admin123")
    print("  支持: OpenAI | DeepSeek | Qwen | 智谱 | Kimi")
    print("        百度文心 | Claude | Gemini | 自定义接口")
    print("  内置: 纯本地模式（无需任何API）")
    print("-" * 60)
    if _IS_WINDOWS:
        # Windows 特有功能状态
        _has_pywin32 = False
        try:
            import win32com.client
            _has_pywin32 = True
        except ImportError:
            pass
        _has_tesseract_exe = False
        try:
            import subprocess as _sp
            _r = _sp.run(['where', 'tesseract'], capture_output=True, text=True, timeout=5)
            _has_tesseract_exe = _r.returncode == 0
        except Exception:
            pass

        print(f"  .doc 旧版文件解析: {'可用' if _has_pywin32 else '不可用（需 pywin32 + Word/WPS）'}")
        print(f"  OCR 图片文字识别: {'可用' if HAS_TESSERACT and _has_tesseract_exe else '不可用（需 Tesseract OCR）'}")
        print(f"  图片上传处理:       {'可用' if HAS_PIL else '不可用（需 Pillow）'}")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)
