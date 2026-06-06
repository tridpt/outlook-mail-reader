"""Lưu/đọc MSAL token cache dưới dạng mã hóa.

MSAL quản lý refresh token cho NHIỀU tài khoản trong cùng một cache.
Ta serialize cache đó rồi mã hóa (Fernet/AES) trước khi ghi xuống đĩa,
nên refresh token không nằm ở dạng plaintext.
"""
from __future__ import annotations

import os
import stat

from cryptography.fernet import Fernet
from msal import SerializableTokenCache

from config import ENCRYPTION_KEY_PATH, TOKEN_CACHE_PATH


def _load_or_create_key() -> bytes:
    """Đọc khóa mã hóa, tạo mới nếu chưa có. File khóa đặt quyền chỉ chủ sở hữu đọc."""
    if ENCRYPTION_KEY_PATH.exists():
        return ENCRYPTION_KEY_PATH.read_bytes()

    key = Fernet.generate_key()
    ENCRYPTION_KEY_PATH.write_bytes(key)
    # Hạn chế quyền truy cập (hiệu lực trên hệ POSIX; Windows bỏ qua an toàn)
    try:
        os.chmod(ENCRYPTION_KEY_PATH, stat.S_IRUSR | stat.S_IWUSR)
    except (OSError, NotImplementedError):
        pass
    return key


def load_cache() -> SerializableTokenCache:
    """Tạo SerializableTokenCache, nạp dữ liệu đã giải mã (nếu có)."""
    cache = SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        try:
            fernet = Fernet(_load_or_create_key())
            data = fernet.decrypt(TOKEN_CACHE_PATH.read_bytes())
            cache.deserialize(data.decode("utf-8"))
        except Exception:
            # Cache hỏng/đổi khóa -> bỏ qua, người dùng đăng nhập lại
            pass
    return cache


def save_cache(cache: SerializableTokenCache) -> None:
    """Mã hóa và ghi cache xuống đĩa nếu có thay đổi."""
    if not cache.has_state_changed:
        return
    fernet = Fernet(_load_or_create_key())
    token = fernet.encrypt(cache.serialize().encode("utf-8"))
    TOKEN_CACHE_PATH.write_bytes(token)
    try:
        os.chmod(TOKEN_CACHE_PATH, stat.S_IRUSR | stat.S_IWUSR)
    except (OSError, NotImplementedError):
        pass
