"""Đăng nhập Outlook đa tài khoản + đọc mail qua Microsoft Graph.

Cơ chế:
- Mỗi tài khoản đăng nhập 1 lần bằng *device code flow* (mở link, nhập code).
- MSAL lưu refresh token vào cache (đã mã hóa) -> các lần sau tự lấy
  access token mới mà KHÔNG cần đăng nhập lại.
- Không bao giờ xử lý/lưu mật khẩu.
"""
from __future__ import annotations

import threading
import uuid
from typing import Any

import msal
import requests

from config import (
    AUTHORITY,
    CLIENT_ID,
    DEFAULT_MAIL_COUNT,
    GRAPH_BASE,
    SCOPES,
)
from store import load_cache, save_cache


class OutlookEngine:
    """Quản lý nhiều tài khoản Outlook trong một MSAL token cache dùng chung."""

    def __init__(self) -> None:
        self._cache = load_cache()
        self._lock = threading.Lock()
        # Các phiên đăng nhập device-code đang chờ người dùng nhập code
        self._login_jobs: dict[str, dict[str, Any]] = {}

    # ---------- App ----------
    def _build_app(self) -> msal.PublicClientApplication:
        if not CLIENT_ID:
            raise RuntimeError(
                "Chưa cấu hình OUTLOOK_CLIENT_ID. Đăng ký một Azure App "
                "(public client) rồi đặt biến môi trường OUTLOOK_CLIENT_ID "
                "hoặc sửa trong config.py."
            )
        return msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=self._cache
        )

    def _persist(self) -> None:
        save_cache(self._cache)

    # ---------- Đăng nhập (device code flow) ----------
    def begin_login(self) -> dict[str, Any]:
        """Khởi tạo phiên đăng nhập, trả về code + link để người dùng nhập.

        Sau khi gọi, chạy nền chờ người dùng hoàn tất; client poll qua
        login_status(job_id).
        """
        app = self._build_app()
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Không tạo được device flow: {flow.get('error_description', flow)}"
            )

        job_id = uuid.uuid4().hex
        self._login_jobs[job_id] = {"status": "pending", "account": None, "error": None}

        def _wait() -> None:
            try:
                # Chặn cho tới khi người dùng nhập code xong hoặc hết hạn
                result = app.acquire_token_by_device_flow(flow)
                if "access_token" in result:
                    self._persist()
                    acct = result.get("id_token_claims", {})
                    self._login_jobs[job_id].update(
                        status="success",
                        account={
                            "username": acct.get("preferred_username")
                            or acct.get("email")
                            or "?",
                            "name": acct.get("name", ""),
                        },
                    )
                else:
                    self._login_jobs[job_id].update(
                        status="error",
                        error=result.get("error_description", "Đăng nhập thất bại."),
                    )
            except Exception as exc:  # noqa: BLE001
                self._login_jobs[job_id].update(status="error", error=str(exc))

        threading.Thread(target=_wait, daemon=True).start()

        return {
            "job_id": job_id,
            "user_code": flow["user_code"],
            "verification_uri": flow["verification_uri"],
            "message": flow.get("message", ""),
            "expires_in": flow.get("expires_in", 900),
        }

    def login_status(self, job_id: str) -> dict[str, Any]:
        job = self._login_jobs.get(job_id)
        if job is None:
            return {"status": "unknown"}
        return job

    # ---------- Quản lý tài khoản ----------
    def list_accounts(self) -> list[dict[str, str]]:
        app = self._build_app()
        out = []
        for a in app.get_accounts():
            out.append(
                {
                    "home_account_id": a["home_account_id"],
                    "username": a.get("username", "?"),
                }
            )
        return out

    def remove_account(self, home_account_id: str) -> bool:
        app = self._build_app()
        target = next(
            (a for a in app.get_accounts() if a["home_account_id"] == home_account_id),
            None,
        )
        if target is None:
            return False
        app.remove_account(target)
        self._persist()
        return True

    # ---------- Lấy token (tự refresh) ----------
    def _token_for(self, home_account_id: str) -> str:
        app = self._build_app()
        account = next(
            (a for a in app.get_accounts() if a["home_account_id"] == home_account_id),
            None,
        )
        if account is None:
            raise RuntimeError("Không tìm thấy tài khoản. Hãy đăng nhập lại.")

        result = app.acquire_token_silent(SCOPES, account=account)
        self._persist()
        if not result or "access_token" not in result:
            raise RuntimeError(
                "Phiên đã hết hạn cho tài khoản này, cần đăng nhập lại."
            )
        return result["access_token"]

    # ---------- Thư mục ----------
    def list_folders(self, home_account_id: str) -> list[dict[str, Any]]:
        """Danh sách thư mục mail + số mail chưa đọc của mỗi thư mục."""
        token = self._token_for(home_account_id)
        url = (
            f"{GRAPH_BASE}/me/mailFolders"
            "?$top=50&$select=id,displayName,unreadItemCount,totalItemCount"
        )
        resp = requests.get(
            url, headers={"Authorization": f"Bearer {token}"}, timeout=30
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}")
        return [
            {
                "id": f.get("id", ""),
                "name": f.get("displayName", ""),
                "unread": f.get("unreadItemCount", 0),
                "total": f.get("totalItemCount", 0),
            }
            for f in resp.json().get("value", [])
        ]

    def inbox_unread_count(self, home_account_id: str) -> int:
        """Số mail chưa đọc trong Inbox của 1 tài khoản (cho badge)."""
        token = self._token_for(home_account_id)
        resp = requests.get(
            f"{GRAPH_BASE}/me/mailFolders/inbox?$select=unreadItemCount",
            headers={"Authorization": f"Bearer {token}"},
            timeout=30,
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}")
        return resp.json().get("unreadItemCount", 0)

    def account_unread_counts(self) -> list[dict[str, Any]]:
        """Số mail chưa đọc của Inbox cho từng tài khoản."""
        out = []
        for acc in self.list_accounts():
            entry: dict[str, Any] = {"home_account_id": acc["home_account_id"]}
            try:
                entry["unread"] = self.inbox_unread_count(acc["home_account_id"])
            except Exception:  # noqa: BLE001
                entry["unread"] = None
            out.append(entry)
        return out

    # ---------- Đọc mail ----------
    @staticmethod
    def _build_filter(
        unread_only: bool,
        has_attachments: bool,
        date_from: str | None,
        date_to: str | None,
    ) -> str:
        """Ghép biểu thức $filter của Graph từ các tiêu chí."""
        parts: list[str] = []
        if unread_only:
            parts.append("isRead eq false")
        if has_attachments:
            parts.append("hasAttachments eq true")
        if date_from:
            parts.append(f"receivedDateTime ge {date_from}T00:00:00Z")
        if date_to:
            parts.append(f"receivedDateTime le {date_to}T23:59:59Z")
        return " and ".join(parts)

    def get_messages(
        self,
        home_account_id: str,
        count: int = DEFAULT_MAIL_COUNT,
        skip: int = 0,
        folder: str | None = None,
        unread_only: bool = False,
        has_attachments: bool = False,
        sender: str | None = None,
        date_from: str | None = None,
        date_to: str | None = None,
    ) -> dict[str, Any]:
        """Lấy mail (có phân trang, theo thư mục, có lọc).

        Trả về {"messages": [...], "has_more": bool}.
        Lọc theo người gửi (`sender`) làm phía client trên trang đã tải vì
        Graph $filter không hỗ trợ tìm gần đúng địa chỉ người gửi.
        """
        token = self._token_for(home_account_id)
        base = (
            f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
            if folder
            else f"{GRAPH_BASE}/me/messages"
        )
        url = (
            f"{base}?$top={count}&$skip={skip}&$orderby=receivedDateTime desc"
            "&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,"
            "hasAttachments,webLink"
        )
        flt = self._build_filter(unread_only, has_attachments, date_from, date_to)
        if flt:
            from urllib.parse import quote

            url += f"&$filter={quote(flt)}"

        resp = requests.get(
            url, headers={"Authorization": f"Bearer {token}"}, timeout=30
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}"
            )
        data = resp.json()

        messages = []
        for m in data.get("value", []):
            s = (m.get("from") or {}).get("emailAddress", {})
            messages.append(
                {
                    "id": m.get("id", ""),
                    "subject": m.get("subject", "(không tiêu đề)"),
                    "from_name": s.get("name", ""),
                    "from_address": s.get("address", ""),
                    "received": m.get("receivedDateTime", ""),
                    "preview": m.get("bodyPreview", ""),
                    "is_read": m.get("isRead", False),
                    "has_attachments": m.get("hasAttachments", False),
                    "web_link": m.get("webLink", ""),
                }
            )

        has_more = bool(data.get("@odata.nextLink"))

        if sender:
            q = sender.lower()
            messages = [
                m
                for m in messages
                if q in m["from_name"].lower() or q in m["from_address"].lower()
            ]

        return {"messages": messages, "has_more": has_more}

    def get_message_detail(
        self, home_account_id: str, message_id: str
    ) -> dict[str, Any]:
        """Lấy nội dung đầy đủ của 1 mail (kèm body HTML/text)."""
        token = self._token_for(home_account_id)
        url = (
            f"{GRAPH_BASE}/me/messages/{message_id}"
            "?$select=id,subject,from,toRecipients,ccRecipients,"
            "receivedDateTime,body,isRead,hasAttachments,webLink"
        )
        resp = requests.get(
            url, headers={"Authorization": f"Bearer {token}"}, timeout=30
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}"
            )
        m = resp.json()
        sender = (m.get("from") or {}).get("emailAddress", {})

        def _addrs(key: str) -> list[str]:
            out = []
            for r in m.get(key, []) or []:
                ea = r.get("emailAddress", {})
                out.append(ea.get("name") or ea.get("address", ""))
            return out

        body = m.get("body", {}) or {}
        return {
            "id": m.get("id", ""),
            "subject": m.get("subject", "(không tiêu đề)"),
            "from_name": sender.get("name", ""),
            "from_address": sender.get("address", ""),
            "to": _addrs("toRecipients"),
            "cc": _addrs("ccRecipients"),
            "received": m.get("receivedDateTime", ""),
            "body_type": body.get("contentType", "text"),
            "body": body.get("content", ""),
            "is_read": m.get("isRead", False),
            "has_attachments": m.get("hasAttachments", False),
            "web_link": m.get("webLink", ""),
        }

    # ---------- Đính kèm ----------
    def list_attachments(
        self, home_account_id: str, message_id: str
    ) -> list[dict[str, Any]]:
        """Danh sách file đính kèm của 1 mail (không tải nội dung)."""
        token = self._token_for(home_account_id)
        url = (
            f"{GRAPH_BASE}/me/messages/{message_id}/attachments"
            "?$select=id,name,contentType,size,isInline"
        )
        resp = requests.get(
            url, headers={"Authorization": f"Bearer {token}"}, timeout=30
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}")
        out = []
        for a in resp.json().get("value", []):
            out.append(
                {
                    "id": a.get("id", ""),
                    "name": a.get("name", "attachment"),
                    "content_type": a.get("contentType", "application/octet-stream"),
                    "size": a.get("size", 0),
                    "is_inline": a.get("isInline", False),
                }
            )
        return out

    def get_attachment(
        self, home_account_id: str, message_id: str, attachment_id: str
    ) -> dict[str, Any]:
        """Tải nội dung 1 file đính kèm. Trả về {name, content_type, bytes}."""
        import base64

        token = self._token_for(home_account_id)
        url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments/{attachment_id}"
        resp = requests.get(
            url, headers={"Authorization": f"Bearer {token}"}, timeout=60
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}")
        a = resp.json()
        content_b64 = a.get("contentBytes")
        if content_b64 is None:
            raise RuntimeError(
                "Đính kèm này không phải file tải được (có thể là mail nhúng)."
            )
        return {
            "name": a.get("name", "attachment"),
            "content_type": a.get("contentType", "application/octet-stream"),
            "bytes": base64.b64decode(content_b64),
        }

    def mark_read(
        self, home_account_id: str, message_id: str, is_read: bool = True
    ) -> bool:
        """Đánh dấu mail đã đọc / chưa đọc (cần quyền Mail.ReadWrite)."""
        token = self._token_for(home_account_id)
        resp = requests.patch(
            f"{GRAPH_BASE}/me/messages/{message_id}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"isRead": is_read},
            timeout=30,
        )
        if resp.status_code not in (200, 204):
            raise RuntimeError(
                f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}"
            )
        return True

    def search_messages(
        self, query: str, count: int = DEFAULT_MAIL_COUNT
    ) -> list[dict[str, Any]]:
        """Tìm mail theo từ khóa trên TẤT CẢ tài khoản đã đăng nhập."""
        result = []
        for acc in self.list_accounts():
            entry: dict[str, Any] = {
                "account": acc["username"],
                "home_account_id": acc["home_account_id"],
            }
            try:
                entry["messages"] = self._search_one(
                    acc["home_account_id"], query, count
                )
            except Exception as exc:  # noqa: BLE001
                entry["error"] = str(exc)
                entry["messages"] = []
            result.append(entry)
        return result

    def _search_one(
        self, home_account_id: str, query: str, count: int
    ) -> list[dict[str, Any]]:
        token = self._token_for(home_account_id)
        # $search dùng để tìm full-text; không kết hợp được $orderby.
        from urllib.parse import quote

        url = (
            f"{GRAPH_BASE}/me/messages"
            f'?$search="{quote(query)}"&$top={count}'
            "&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,webLink"
        )
        resp = requests.get(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "ConsistencyLevel": "eventual",
            },
            timeout=30,
        )
        if resp.status_code != 200:
            raise RuntimeError(
                f"Graph trả lỗi {resp.status_code}: {resp.text[:200]}"
            )
        out = []
        for m in resp.json().get("value", []):
            sender = (m.get("from") or {}).get("emailAddress", {})
            out.append(
                {
                    "id": m.get("id", ""),
                    "subject": m.get("subject", "(không tiêu đề)"),
                    "from_name": sender.get("name", ""),
                    "from_address": sender.get("address", ""),
                    "received": m.get("receivedDateTime", ""),
                    "preview": m.get("bodyPreview", ""),
                    "is_read": m.get("isRead", False),
                    "web_link": m.get("webLink", ""),
                }
            )
        return out

    def get_unified_inbox(
        self,
        count: int = DEFAULT_MAIL_COUNT,
        skip: int = 0,
        folder: str | None = None,
        unread_only: bool = False,
        has_attachments: bool = False,
        sender: str | None = None,
        date_from: str | None = None,
        date_to: str | None = None,
        only_account: str | None = None,
    ) -> list[dict[str, Any]]:
        """Gộp mail của các tài khoản đã đăng nhập (có thể lọc 1 tài khoản)."""
        result = []
        accounts = self.list_accounts()
        if only_account:
            accounts = [a for a in accounts if a["home_account_id"] == only_account]
        for acc in accounts:
            entry: dict[str, Any] = {
                "account": acc["username"],
                "home_account_id": acc["home_account_id"],
                "has_more": False,
            }
            try:
                res = self.get_messages(
                    acc["home_account_id"],
                    count=count,
                    skip=skip,
                    folder=folder,
                    unread_only=unread_only,
                    has_attachments=has_attachments,
                    sender=sender,
                    date_from=date_from,
                    date_to=date_to,
                )
                entry["messages"] = res["messages"]
                entry["has_more"] = res["has_more"]
            except Exception as exc:  # noqa: BLE001
                entry["error"] = str(exc)
                entry["messages"] = []
            result.append(entry)
        return result


engine = OutlookEngine()
