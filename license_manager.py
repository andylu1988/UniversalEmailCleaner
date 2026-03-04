"""
License Key Management for UniversalEmailCleaner
=================================================
AES-256-GCM 加密许可证密钥系统，使用 PBKDF2-HMAC-SHA256 密钥派生。

密钥格式:
  - 4 字节随机 nonce + AES-256-GCM 加密载荷 (5 字节明文 → 21 字节密文+标签) = 25 字节
  - Base32 编码 → 40 字符
  - 显示格式: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX (8 组 × 5 字符)

加密载荷:
  - 版本号 (1 字节)
  - 有效期代码 (1 字节): 0=1天, 1=1周, 2=1月, 3=6月, 4=1年, 5=永久
  - 创建时间 (2 字节): 自 2024-01-01 起的天数偏移 (uint16)
  - 随机填充 (1 字节)
"""

import struct
import os
import json
import base64
import hashlib
import hmac
import time
from datetime import datetime, timedelta

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    _HAS_AESGCM = True
except ImportError:
    _HAS_AESGCM = False


# ===================== 常量 =====================

LICENSE_VERSION = 1
EPOCH_DATE = datetime(2024, 1, 1)

DURATION_MAP = {
    0: ("1天 (1 Day)", timedelta(days=1)),
    1: ("1周 (1 Week)", timedelta(weeks=1)),
    2: ("1个月 (1 Month)", timedelta(days=30)),
    3: ("6个月 (6 Months)", timedelta(days=180)),
    4: ("1年 (1 Year)", timedelta(days=365)),
    5: ("永久 (Permanent)", None),
}

# PBKDF2 密钥派生参数 (拆分以增加逆向难度)
_KD_PARTS = [
    b'\x55\x45\x43\x5f',  # UEC_
    b'\x41\x45\x53\x32',  # AES2
    b'\x35\x36\x5f\x47',  # 56_G
    b'\x43\x4d\x5f\x4c',  # CM_L
    b'\x49\x43\x5f\x56',  # IC_V
    b'\x31\x5f\x53\x45',  # 1_SE
    b'\x43\x52\x45\x54',  # CRET
    b'\x5f\x4b\x45\x59',  # _KEY
]
_KD_SALT = b'UEC_AES256GCM_LICENSE_SALT_2024_v1'
_KD_ITERATIONS = 100_000

# 缓存派生密钥
_cached_aes_key = None


# ===================== 内部函数 =====================

def _derive_key():
    """使用 PBKDF2-HMAC-SHA256 派生 AES-256 密钥 (32 字节)"""
    global _cached_aes_key
    if _cached_aes_key is not None:
        return _cached_aes_key
    passphrase = b''.join(_KD_PARTS)
    _cached_aes_key = hashlib.pbkdf2_hmac(
        'sha256', passphrase, _KD_SALT, _KD_ITERATIONS
    )
    return _cached_aes_key


def _derive_iv(aes_key: bytes, nonce: bytes) -> bytes:
    """从 AES 密钥和 nonce 确定性地派生 12 字节 GCM IV"""
    return hashlib.sha256(aes_key + nonce).digest()[:12]


def _format_key(raw_bytes: bytes) -> str:
    """将原始字节格式化为人类可读的许可证密钥"""
    encoded = base64.b32encode(raw_bytes).decode('ascii').rstrip('=')
    groups = [encoded[i:i + 5] for i in range(0, len(encoded), 5)]
    return '-'.join(groups)


def _parse_key(key_str: str) -> bytes:
    """将人类可读的许可证密钥解析回原始字节"""
    cleaned = key_str.replace('-', '').replace(' ', '').strip().upper()
    padding = (8 - len(cleaned) % 8) % 8
    cleaned += '=' * padding
    return base64.b32decode(cleaned)


# ===================== HMAC 回退 (无 cryptography 库时使用) =====================

def _hmac_generate(aes_key: bytes, plaintext: bytes, nonce: bytes) -> bytes:
    """使用 HMAC-SHA256 流密码加密 + MAC 认证 (回退方案)"""
    # 生成密钥流用于 XOR 加密
    keystream = hmac.new(aes_key, b'\x01' + nonce, hashlib.sha256).digest()[:len(plaintext)]
    encrypted = bytes(a ^ b for a, b in zip(plaintext, keystream))
    # Encrypt-then-MAC
    mac = hmac.new(aes_key, b'\x02' + nonce + encrypted, hashlib.sha256).digest()[:16]
    return nonce + encrypted + mac  # 4 + 5 + 16 = 25 bytes


def _hmac_decrypt(aes_key: bytes, raw: bytes) -> bytes:
    """解密 HMAC 回退方案的密钥"""
    nonce = raw[:4]
    encrypted = raw[4:-16]
    mac = raw[-16:]
    # 验证 MAC
    expected_mac = hmac.new(aes_key, b'\x02' + nonce + encrypted, hashlib.sha256).digest()[:16]
    if not hmac.compare_digest(mac, expected_mac):
        raise ValueError("MAC 验证失败")
    # 解密
    keystream = hmac.new(aes_key, b'\x01' + nonce, hashlib.sha256).digest()[:len(encrypted)]
    plaintext = bytes(a ^ b for a, b in zip(encrypted, keystream))
    return plaintext


# ===================== 公共 API =====================

def generate_license_key(duration_code: int) -> str:
    """
    生成加密的许可证密钥。

    Args:
        duration_code: 0=1天, 1=1周, 2=1月, 3=6月, 4=1年, 5=永久

    Returns:
        格式化的许可证密钥字符串 (40 字符, 8×5 分组)
    """
    if duration_code not in DURATION_MAP:
        raise ValueError(f"无效的有效期代码: {duration_code}")

    aes_key = _derive_key()
    nonce = os.urandom(4)

    # 自纪元以来的天数偏移
    day_offset = (datetime.now() - EPOCH_DATE).days
    if day_offset < 0:
        day_offset = 0
    random_byte = os.urandom(1)

    # 明文: version(1) + duration(1) + day_offset(2) + random(1) = 5 字节
    plaintext = struct.pack('>BBH', LICENSE_VERSION, duration_code, day_offset) + random_byte

    if _HAS_AESGCM:
        # AES-256-GCM 加密
        iv = _derive_iv(aes_key, nonce)
        aesgcm = AESGCM(aes_key)
        ct_tag = aesgcm.encrypt(iv, plaintext, nonce)  # 5 + 16 = 21 字节
        raw = nonce + ct_tag  # 4 + 21 = 25 字节
    else:
        # HMAC 回退方案
        raw = _hmac_generate(aes_key, plaintext, nonce)  # 25 字节

    return _format_key(raw)


def validate_license_key(key_str: str) -> dict:
    """
    验证并解码许可证密钥。

    Returns:
        dict: {
            valid: bool,
            duration_code: int or None,
            duration_name: str or None,
            created: datetime or None,
            expires: datetime or None,
            expired: bool,
            error: str or None,
        }
    """
    result = {
        'valid': False,
        'duration_code': None,
        'duration_name': None,
        'created': None,
        'expires': None,
        'expired': False,
        'error': None,
    }

    try:
        raw = _parse_key(key_str)
    except Exception as e:
        result['error'] = f"密钥格式无效: {e}"
        return result

    if len(raw) < 25:
        result['error'] = "密钥长度不正确"
        return result

    aes_key = _derive_key()
    nonce = raw[:4]

    # 尝试解密
    plaintext = None
    if _HAS_AESGCM:
        try:
            iv = _derive_iv(aes_key, nonce)
            ct_tag = raw[4:]
            aesgcm = AESGCM(aes_key)
            plaintext = aesgcm.decrypt(iv, ct_tag, nonce)
        except Exception:
            pass

    if plaintext is None:
        # 尝试 HMAC 回退
        try:
            plaintext = _hmac_decrypt(aes_key, raw)
        except Exception:
            result['error'] = "密钥无效或已损坏 (解密/验证失败)"
            return result

    if len(plaintext) < 5:
        result['error'] = "密钥数据损坏"
        return result

    version, duration_code, day_offset = struct.unpack('>BBH', plaintext[:4])

    if version != LICENSE_VERSION:
        result['error'] = f"不支持的密钥版本: {version}"
        return result

    if duration_code not in DURATION_MAP:
        result['error'] = f"无效的有效期代码: {duration_code}"
        return result

    duration_name, delta = DURATION_MAP[duration_code]
    created = EPOCH_DATE + timedelta(days=day_offset)

    if delta is not None:
        expires = created + delta
        expired = datetime.now() > expires
    else:
        expires = None
        expired = False

    result.update({
        'valid': True,
        'duration_code': duration_code,
        'duration_name': duration_name,
        'created': created,
        'expires': expires,
        'expired': expired,
    })

    return result


# ===================== 许可证文件管理 =====================

def get_license_dir() -> str:
    """获取许可证存储目录"""
    docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
    if not os.path.exists(docs_dir):
        os.makedirs(docs_dir)
    return docs_dir


def get_license_path() -> str:
    """获取许可证文件路径"""
    return os.path.join(get_license_dir(), "license.dat")


def save_license(key_str: str) -> None:
    """保存已激活的许可证密钥"""
    license_path = get_license_path()
    data = {
        'key': key_str.strip(),
        'activated_at': datetime.now().isoformat(),
    }
    with open(license_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_license() -> str | None:
    """加载已保存的许可证密钥。返回密钥字符串或 None"""
    license_path = get_license_path()
    if not os.path.exists(license_path):
        return None
    try:
        with open(license_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('key')
    except Exception:
        return None


def check_license() -> dict:
    """
    检查应用程序是否已正确授权。

    Returns:
        dict: {licensed: bool, info: dict or None, error: str or None}
    """
    key = load_license()
    if not key:
        return {'licensed': False, 'info': None, 'error': '未找到许可证'}

    info = validate_license_key(key)
    if not info['valid']:
        return {'licensed': False, 'info': info, 'error': info['error']}

    if info['expired']:
        exp_str = info['expires'].strftime('%Y-%m-%d') if info['expires'] else '未知'
        return {
            'licensed': False,
            'info': info,
            'error': f"许可证已过期 (到期时间: {exp_str})",
        }

    return {'licensed': True, 'info': info, 'error': None}


def activate_license(key_str: str) -> dict:
    """
    使用给定的许可证密钥激活。

    Returns:
        dict: {success: bool, info: dict, error: str or None}
    """
    info = validate_license_key(key_str)

    if not info['valid']:
        return {'success': False, 'info': info, 'error': info['error']}

    if info['expired']:
        exp_str = info['expires'].strftime('%Y-%m-%d') if info['expires'] else '未知'
        return {
            'success': False,
            'info': info,
            'error': f"该许可证已过期 (到期时间: {exp_str})",
        }

    save_license(key_str)
    return {'success': True, 'info': info, 'error': None}


def deactivate_license() -> None:
    """移除当前许可证"""
    license_path = get_license_path()
    if os.path.exists(license_path):
        os.remove(license_path)
