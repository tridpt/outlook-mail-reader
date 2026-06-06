"""Cấu hình cho ứng dụng đọc hộp thư Outlook đa tài khoản.

Dùng Microsoft Graph + OAuth 2.0 (MSAL). KHÔNG lưu mật khẩu:
mỗi tài khoản đăng nhập 1 lần qua Microsoft, sau đó MSAL tự dùng
refresh token (được lưu mã hóa) để lấy mail mà không hỏi lại.
"""
from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
STORAGE_DIR = BASE_DIR / "storage"
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

# Nơi lưu token cache (đã mã hóa) và khóa mã hóa
TOKEN_CACHE_PATH = STORAGE_DIR / "token_cache.bin"
ENCRYPTION_KEY_PATH = STORAGE_DIR / "cache.key"
IMPORTED_ACCOUNTS_PATH = STORAGE_DIR / "imported_accounts.bin"

# ----- Cấu hình OAuth -----
# client_id của Azure App Registration (Public client / native).
# Đặt qua biến môi trường OUTLOOK_CLIENT_ID, hoặc sửa trực tiếp ở đây.
#
# Cách lấy: portal.azure.com -> App registrations -> New registration
#   - Supported account types: "Accounts in any org directory and personal Microsoft accounts"
#   - Authentication -> Advanced -> "Allow public client flows" = Yes
#   - Copy "Application (client) ID" và dán vào đây.
CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "").strip()

# "common" = hỗ trợ cả tài khoản cá nhân (outlook.com) lẫn tài khoản tổ chức.
AUTHORITY = "https://login.microsoftonline.com/common"

# Mail.ReadWrite: đọc mail + đánh dấu đã đọc/chưa đọc (không gửi mail).
# Nếu chỉ muốn đọc thuần, đổi lại thành "Mail.Read".
SCOPES = ["Mail.ReadWrite", "User.Read"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# OAuth/IMAP settings used when importing an existing refresh token.
TOKEN_ENDPOINT = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
IMAP_SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "offline_access",
]
IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993

# Số mail lấy về mỗi tài khoản khi tải hộp thư
DEFAULT_MAIL_COUNT = 15
