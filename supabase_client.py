# -*- coding: utf-8 -*-
"""
Supabase 数据库客户端
替代本地文件存储，支持 Netlify 无状态部署
"""

import os
import json
import hashlib
import secrets
from datetime import datetime
from typing import Optional, List, Dict, Any

try:
    from supabase import create_client, Client
    HAS_SUPABASE = True
except ImportError:
    HAS_SUPABASE = False


# ==================== 初始化 Supabase 客户端 ====================
_url = os.environ.get('SUPABASE_URL', '')
_key = os.environ.get('SUPABASE_ANON_KEY', '')
_service_key = os.environ.get('SUPABASE_SERVICE_KEY', '')

_db: Optional[Client] = None


def _get_client() -> Client:
    """懒加载 Supabase 客户端"""
    global _db
    if _db is None:
        if not HAS_SUPABASE:
            raise RuntimeError("supabase 包未安装，请执行: pip install supabase")
        if not _url or not _key:
            raise RuntimeError(
                "缺少 SUPABASE_URL 或 SUPABASE_ANON_KEY 环境变量。"
                "请在 Netlify 环境变量中设置这些值。"
            )
        _db = create_client(_url, _key)
    return _db


def is_configured() -> bool:
    """检查 Supabase 是否已配置"""
    return HAS_SUPABASE and bool(_url and _key)


# ==================== 用户管理 ====================

def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode('utf-8')).hexdigest()


def _generate_token() -> str:
    return secrets.token_hex(32)


def init_default_user():
    """确保存在默认 admin 用户"""
    client = _get_client()
    default_user = os.environ.get('QUIZ_USERNAME', 'admin')
    default_pw = os.environ.get('QUIZ_PASSWORD', 'admin123')

    result = client.table('users').select('username').eq('username', default_user).execute()
    if not result.data:
        user_data = {
            'username': default_user,
            'password_hash': _hash_pw(default_pw),
            'display_name': default_user,
            'remember_token': _generate_token(),
            'created_at': datetime.utcnow().isoformat()
        }
        client.table('users').insert(user_data).execute()


def get_user(username: str) -> Optional[Dict]:
    """根据用户名获取用户数据"""
    client = _get_client()
    result = client.table('users').select('*').eq('username', username).execute()
    return result.data[0] if result.data else None


def verify_password(username: str, password: str) -> bool:
    """验证用户密码"""
    user = get_user(username)
    if not user:
        return False
    return user['password_hash'] == _hash_pw(password)


def create_user(username: str, password: str, display_name: str) -> Dict:
    """创建新用户"""
    client = _get_client()
    token = _generate_token()
    user_data = {
        'username': username,
        'password_hash': _hash_pw(password),
        'display_name': display_name or username,
        'remember_token': token,
        'created_at': datetime.utcnow().isoformat()
    }
    client.table('users').insert(user_data).execute()
    return user_data


def user_exists(username: str) -> bool:
    """检查用户名是否存在"""
    client = _get_client()
    result = client.table('users').select('username').eq('username', username).execute()
    return len(result.data) > 0


def get_all_usernames() -> List[str]:
    """获取所有用户名（用于匹配 remember token）"""
    client = _get_client()
    result = client.table('users').select('username').execute()
    return [r['username'] for r in result.data]


def update_remember_token(username: str, token: Optional[str]):
    """更新 remember token"""
    client = _get_client()
    client.table('users').update({'remember_token': token}).eq('username', username).execute()


def find_by_remember_token(token: str) -> Optional[Dict]:
    """根据 remember token 查找用户"""
    client = _get_client()
    result = client.table('users').select('*').eq('remember_token', token).execute()
    return result.data[0] if result.data else None


# ==================== 题库空间管理 ====================

def _generate_space_id(name: str) -> str:
    """生成唯一 space_id"""
    import re
    base = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '_', name)[:30]
    ts = datetime.now().strftime('%Y%m%d%H%M%S')
    return f"{base}_{ts}"


def list_spaces(username: str) -> List[Dict]:
    """列出用户的所有题库空间"""
    client = _get_client()
    result = client.table('spaces').select('*').eq('username', username).order('created_at', desc=True).execute()
    return result.data or []


def get_space(space_id: str) -> Optional[Dict]:
    """获取指定空间"""
    client = _get_client()
    result = client.table('spaces').select('*').eq('id', space_id).execute()
    return result.data[0] if result.data else None


def get_space_questions(space_id: str) -> List[Dict]:
    """获取空间下所有题目"""
    client = _get_client()
    result = client.table('questions').select('*').eq('space_id', space_id).order('bank_id').execute()
    return result.data or []


def create_space(username: str, name: str, source_text: str = '') -> Dict:
    """创建新空间"""
    client = _get_client()
    space_id = _generate_space_id(name)
    now = datetime.utcnow().isoformat()
    space_data = {
        'id': space_id,
        'username': username,
        'name': name,
        'source_text': source_text,
        'created_at': now,
        'updated_at': now
    }
    client.table('spaces').insert(space_data).execute()
    return space_data


def update_space_name(space_id: str, new_name: str):
    """更新空间名称"""
    client = _get_client()
    client.table('spaces').update({
        'name': new_name,
        'updated_at': datetime.utcnow().isoformat()
    }).eq('id', space_id).execute()


def update_space_source(space_id: str, source_text: str):
    """更新空间源文本"""
    client = _get_client()
    client.table('spaces').update({
        'source_text': source_text,
        'updated_at': datetime.utcnow().isoformat()
    }).eq('id', space_id).execute()


def delete_space(space_id: str):
    """删除空间及其所有题目和错题（CASCADE）"""
    client = _get_client()
    client.table('spaces').delete().eq('id', space_id).execute()


def space_exists(space_id: str) -> bool:
    """检查空间是否存在"""
    return get_space(space_id) is not None


def get_spaces_summary(username: str) -> List[Dict]:
    """获取所有空间摘要（含题目数和错题数）"""
    client = _get_client()
    spaces = list_spaces(username)
    summary = []
    for s in spaces:
        sid = s['id']
        # 统计题目数
        q_result = client.table('questions').select('type').eq('space_id', sid).execute()
        questions = q_result.data or []
        q_count = len(questions)
        q_types = list(set(q.get('type', 'choice') for q in questions))
        # 统计错题数
        w_result = client.table('wrong_questions').select('id').eq('space_id', sid).execute()
        w_count = len(w_result.data or [])

        summary.append({
            'space_id': sid,
            'name': s.get('name', sid),
            'created_at': s.get('created_at', ''),
            'question_count': q_count,
            'question_types': q_types,
            'wrong_count': w_count
        })
    return summary


# ==================== 题目管理 ====================

def save_questions_to_space(space_id: str, questions: List[Dict]) -> int:
    """批量保存题目到指定空间"""
    client = _get_client()
    # 获取当前最大 bank_id
    result = client.table('questions').select('bank_id').eq('space_id', space_id).order('bank_id', desc=True).limit(1).execute()
    next_id = result.data[0]['bank_id'] + 1 if result.data else 1

    rows = []
    now = datetime.utcnow().isoformat()
    for q in questions:
        rows.append({
            'space_id': space_id,
            'bank_id': next_id,
            'stem': q.get('stem', ''),
            'type': q.get('type', 'choice'),
            'options': json.dumps(q.get('options', []), ensure_ascii=False),
            'answer': q.get('answer', ''),
            'analysis': q.get('analysis', ''),
            'saved_at': now
        })
        next_id += 1

    client.table('questions').insert(rows).execute()
    # 更新空间时间戳
    update_space_source(space_id, '')  # 只更新时间
    return len(rows)


def delete_question_by_bank_id(bank_id: int) -> Optional[str]:
    """根据 bank_id 删除题目，返回所属 space_id 或 None"""
    client = _get_client()
    # 先查找题目
    result = client.table('questions').select('*').eq('bank_id', bank_id).execute()
    if not result.data:
        return None
    space_id = result.data[0]['space_id']
    q_id = result.data[0]['id']
    client.table('questions').delete().eq('id', q_id).execute()
    return space_id


def clear_all_spaces(username: str):
    """清空用户所有空间和题目"""
    client = _get_client()
    # 获取所有空间 ID
    result = client.table('spaces').select('id').eq('username', username).execute()
    for s in (result.data or []):
        delete_space(s['id'])


# ==================== 错题本管理 ====================

def record_wrong_answers(space_id: str, wrong_list: List[Dict]) -> int:
    """记录答错的题目"""
    client = _get_client()
    now = datetime.utcnow().isoformat()
    rows = []
    for wq in wrong_list:
        rows.append({
            'space_id': space_id,
            'bank_id': wq.get('bank_id', 0),
            'stem': wq.get('stem', ''),
            'type': wq.get('type', 'choice'),
            'options': json.dumps(wq.get('options', []), ensure_ascii=False),
            'answer': wq.get('answer', ''),
            'analysis': wq.get('analysis', ''),
            'user_answer': wq.get('user_answer', ''),
            'recorded_at': now
        })
    if rows:
        client.table('wrong_questions').insert(rows).execute()
    return len(rows)


def get_wrong_questions(space_id: str) -> List[Dict]:
    """获取错题本列表"""
    client = _get_client()
    result = client.table('wrong_questions').select('*').eq('space_id', space_id).order('recorded_at', desc=True).execute()
    return result.data or []


def clear_wrong_questions(space_id: str):
    """清空错题本"""
    client = _get_client()
    client.table('wrong_questions').delete().eq('space_id', space_id).execute()


# ==================== 初始化 ====================
def initialize():
    """初始化：创建默认用户"""
    if not is_configured():
        return
    try:
        init_default_user()
    except Exception:
        pass  # 表可能还没创建，首次部署会报错但不影响
