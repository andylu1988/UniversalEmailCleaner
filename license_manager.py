"""
License Key Management for UniversalEmailCleaner
=================================================
机器码绑定许可证系统。

密钥明文载荷:
  - version (1 byte)
  - duration_code (1 byte)
  - created_day_offset (2 bytes, from 2024-01-01)
  - machine_hash8 (8 bytes, SHA256(machine_code) 前 8 字节)
  - random_pad (1 byte)

密钥封装:
  - nonce 4 bytes + AES-GCM(ciphertext+tag)
  - Base32 可读分组格式
"""

import base64
from datetime import datetime, timedelta
import hashlib
import hmac
import json
import os
import platform
import struct
import uuid

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    _HAS_AESGCM = True
except ImportError:
    _HAS_AESGCM = False

try:
    import winreg
except Exception:
    winreg = None


LICENSE_VERSION = 2
EPOCH_DATE = datetime(2024, 1, 1)

DURATION_MAP = {
    0: ("1天 (1 Day)", timedelta(days=1)),
    1: ("1周 (1 Week)", timedelta(weeks=1)),
    2: ("1个月 (1 Month)", timedelta(days=30)),
    3: ("6个月 (6 Months)", timedelta(days=180)),
    4: ("1年 (1 Year)", timedelta(days=365)),
    5: ("永久 (Permanent)", None),
}

_KD_PARTS = [
    b'\x55\x45\x43\x5f',
    b'\x41\x45\x53\x32',
    b'\x35\x36\x5f\x47',
    b'\x43\x4d\x5f\x4c',
    b'\x49\x43\x5f\x56',
    b'\x31\x5f\x53\x45',
    b'\x43\x52\x45\x54',
    b'\x5f\x4b\x45\x59',
]
_KD_SALT = b'UEC_AES256GCM_LICENSE_SALT_2024_v2_MACHINE'
_KD_ITERATIONS = 100_000

_cached_aes_key = None
_cached_machine_code = None


def _derive_key() -> bytes:
    global _cached_aes_key
    if _cached_aes_key is not None:
        return _cached_aes_key
    passphrase = b''.join(_KD_PARTS)
    _cached_aes_key = hashlib.pbkdf2_hmac('sha256', passphrase, _KD_SALT, _KD_ITERATIONS)
    return _cached_aes_key


def _derive_iv(aes_key: bytes, nonce: bytes) -> bytes:
    return hashlib.sha256(aes_key + nonce).digest()[:12]


def _format_key(raw_bytes: bytes) -> str:
    encoded = base64.b32encode(raw_bytes).decode('ascii').rstrip('=')
    groups = [encoded[i:i + 5] for i in range(0, len(encoded), 5)]
    return '-'.join(groups)


def _parse_key(key_str: str) -> bytes:
    cleaned = key_str.replace('-', '').replace(' ', '').strip().upper()
    padding = (8 - len(cleaned) % 8) % 8
    cleaned += '=' * padding
    return base64.b32decode(cleaned)


def _get_machine_guid_win() -> str:
    if platform.system().lower() != 'windows' or winreg is None:
        return ''
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            val, _ = winreg.QueryValueEx(key, "MachineGuid")
            return str(val).strip()
    except Exception:
        return ''


def _normalize_machine_code(machine_code: str) -> str:
    return ''.join(ch for ch in machine_code.upper() if ch.isalnum())


def _format_machine_code(raw_hex: str) -> str:
    groups = [raw_hex[i:i + 4] for i in range(0, len(raw_hex), 4)]
    return '-'.join(groups)


def get_machine_code() -> str:
    global _cached_machine_code
    if _cached_machine_code:
        return _cached_machine_code

    parts = [
        platform.system(),
        platform.release(),
        platform.machine(),
        platform.node(),
        str(uuid.getnode()),
        _get_machine_guid_win(),
    ]
    raw = '|'.join(part.strip() for part in parts if part and str(part).strip())
    digest_hex = hashlib.sha256(raw.encode('utf-8')).hexdigest().upper()[:32]
    _cached_machine_code = _format_machine_code(digest_hex)
    return _cached_machine_code


def _machine_hash8(machine_code: str) -> bytes:
    normalized = _normalize_machine_code(machine_code)
    return hashlib.sha256(normalized.encode('utf-8')).digest()[:8]


def _hmac_generate(aes_key: bytes, plaintext: bytes, nonce: bytes) -> bytes:
    keystream = hmac.new(aes_key, b'\x01' + nonce, hashlib.sha256).digest()[:len(plaintext)]
    encrypted = bytes(a ^ b for a, b in zip(plaintext, keystream))
    mac = hmac.new(aes_key, b'\x02' + nonce + encrypted, hashlib.sha256).digest()[:16]
    return nonce + encrypted + mac


def _hmac_decrypt(aes_key: bytes, raw: bytes) -> bytes:
    nonce = raw[:4]
    encrypted = raw[4:-16]
    mac = raw[-16:]
    expected_mac = hmac.new(aes_key, b'\x02' + nonce + encrypted, hashlib.sha256).digest()[:16]
    if not hmac.compare_digest(mac, expected_mac):
        raise ValueError("MAC 验证失败")
    keystream = hmac.new(aes_key, b'\x01' + nonce, hashlib.sha256).digest()[:len(encrypted)]
    return bytes(a ^ b for a, b in zip(encrypted, keystream))


def _calc_remaining_days(expires: datetime | None) -> int | None:
    if expires is None:
        return None
    return max(0, (expires.date() - datetime.now().date()).days)


def generate_license_key(duration_code: int, machine_code: str | None = None) -> str:
    if duration_code not in DURATION_MAP:
        raise ValueError(f"无效的有效期代码: {duration_code}")

    machine_code = machine_code or get_machine_code()
    machine_hash = _machine_hash8(machine_code)

    aes_key = _derive_key()
    nonce = os.urandom(4)
    day_offset = max(0, (datetime.now() - EPOCH_DATE).days)
    random_byte = os.urandom(1)

    plaintext = struct.pack('>BBH', LICENSE_VERSION, duration_code, day_offset) + machine_hash + random_byte

    if _HAS_AESGCM:
        iv = _derive_iv(aes_key, nonce)
        aesgcm = AESGCM(aes_key)
        ct_tag = aesgcm.encrypt(iv, plaintext, nonce)
        raw = nonce + ct_tag
    else:
        raw = _hmac_generate(aes_key, plaintext, nonce)

    return _format_key(raw)


def validate_license_key(key_str: str, machine_code: str | None = None) -> dict:
    result = {
        'valid': False,
        'duration_code': None,
        'duration_name': None,
        'created': None,
        'expires': None,
        'expired': False,
        'days_remaining': None,
        'machine_bound': False,
        'machine_match': False,
        'error': None,
    }

    try:
        raw = _parse_key(key_str)
    except Exception as e:
        result['error'] = f"密钥格式无效: {e}"
        return result

    if len(raw) < 33:
        result['error'] = "密钥长度不正确"
        return result

    aes_key = _derive_key()
    nonce = raw[:4]
    plaintext = None

    if _HAS_AESGCM:
        try:
            iv = _derive_iv(aes_key, nonce)
            plaintext = AESGCM(aes_key).decrypt(iv, raw[4:], nonce)
        except Exception:
            pass

    if plaintext is None:
        try:
            plaintext = _hmac_decrypt(aes_key, raw)
        except Exception:
            result['error'] = "密钥无效或已损坏 (解密/验证失败)"
            return result

    if len(plaintext) < 13:
        result['error'] = "密钥数据损坏"
        return result

    version, duration_code, day_offset = struct.unpack('>BBH', plaintext[:4])
    machine_hash = plaintext[4:12]

    if version != LICENSE_VERSION:
        result['error'] = f"不支持的密钥版本: {version}"
        return result
    if duration_code not in DURATION_MAP:
        result['error'] = f"无效的有效期代码: {duration_code}"
        return result

    machine_code = machine_code or get_machine_code()
    expected_hash = _machine_hash8(machine_code)
    machine_match = hmac.compare_digest(machine_hash, expected_hash)

    if not machine_match:
        result['error'] = "该密钥不属于当前机器 (Machine Code 不匹配)"
        result['machine_bound'] = True
        result['machine_match'] = False
        return result

    duration_name, delta = DURATION_MAP[duration_code]
    created = EPOCH_DATE + timedelta(days=day_offset)
    expires = created + delta if delta is not None else None
    expired = datetime.now() > expires if expires else False

    result.update({
        'valid': True,
        'duration_code': duration_code,
        'duration_name': duration_name,
        'created': created,
        'expires': expires,
        'expired': expired,
        'days_remaining': _calc_remaining_days(expires),
        'machine_bound': True,
        'machine_match': True,
    })
    return result


def get_license_dir() -> str:
    docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
    if not os.path.exists(docs_dir):
        os.makedirs(docs_dir)
    return docs_dir


def get_license_path() -> str:
    return os.path.join(get_license_dir(), "license.dat")


def save_license(key_str: str, machine_code: str | None = None) -> None:
    data = {
        'key': key_str.strip(),
        'machine_code': machine_code or get_machine_code(),
        'activated_at': datetime.now().isoformat(),
    }
    with open(get_license_path(), 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_license() -> str | None:
    license_path = get_license_path()
    if not os.path.exists(license_path):
        return None
    try:
        with open(license_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('key')
    except Exception:
        return None


def check_license(machine_code: str | None = None) -> dict:
    key = load_license()
    if not key:
        return {'licensed': False, 'info': None, 'error': '未找到许可证'}

    info = validate_license_key(key, machine_code=machine_code)
    if not info['valid']:
        return {'licensed': False, 'info': info, 'error': info['error']}

    if info['expired']:
        exp_str = info['expires'].strftime('%Y-%m-%d') if info['expires'] else '未知'
        return {'licensed': False, 'info': info, 'error': f"许可证已过期 (到期时间: {exp_str})"}

    return {'licensed': True, 'info': info, 'error': None}


def activate_license(key_str: str, machine_code: str | None = None) -> dict:
    machine_code = machine_code or get_machine_code()
    info = validate_license_key(key_str, machine_code=machine_code)

    if not info['valid']:
        return {'success': False, 'info': info, 'error': info['error']}

    if info['expired']:
        exp_str = info['expires'].strftime('%Y-%m-%d') if info['expires'] else '未知'
        return {'success': False, 'info': info, 'error': f"该许可证已过期 (到期时间: {exp_str})"}

    save_license(key_str, machine_code=machine_code)
    return {'success': True, 'info': info, 'error': None}


def deactivate_license() -> None:
    license_path = get_license_path()
    if os.path.exists(license_path):
        os.remove(license_path)
