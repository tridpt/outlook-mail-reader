"""Đăng nhập Outlook đa tài khoản + đọc mail qua Microsoft Graph.

Cơ chế:
- Mỗi tài khoản đăng nhập 1 lần bằng *device code flow* (mở link, nhập code).
- MSAL lưu refresh token vào cache (đã mã hóa) -> các lần sau tự lấy
  access token mới mà KHÔNG cần đăng nhập lại.
- Không bao giờ xử lý/lưu mật khẩu.
"""
from __future__ import annotations

import base64
import email.utils
import hashlib
import html
import imaplib
import re
import threading
import uuid
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from queue import Empty, Queue
from typing import Any

import msal
import requests

from config import (
    AUTHORITY,
    CLIENT_ID,
    DEFAULT_MAIL_COUNT,
    GRAPH_BASE,
    IMAP_HOST,
    IMAP_PORT,
    IMAP_SCOPES,
    SCOPES,
    TOKEN_ENDPOINT,
)
from store import (
    load_cache,
    load_imported_accounts,
    save_cache,
    save_imported_accounts,
)


class OutlookEngine:
    """Quản lý nhiều tài khoản Outlook trong một MSAL token cache dùng chung."""

    def __init__(self) -> None:
        self._cache = load_cache()
        self._imported = load_imported_accounts()
        self._lock = threading.Lock()
        # Các phiên đăng nhập device-code đang chờ người dùng nhập code
        self._login_jobs: dict[str, dict[str, Any]] = {}
        self._import_jobs: dict[str, dict[str, Any]] = {}

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

    def _persist_imported(self) -> None:
        save_imported_accounts(self._imported)

    @staticmethod
    def _now_iso() -> str:
        return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    @staticmethod
    def _imported_id(email_address: str, client_id: str) -> str:
        digest = hashlib.sha256(
            f"{email_address.lower()}|{client_id.lower()}".encode("utf-8")
        ).hexdigest()[:20]
        return f"imported:{digest}"

    @staticmethod
    def _public_account(account_id: str, data: dict[str, Any]) -> dict[str, Any]:
        mode = data.get("mode", "imported")
        if mode == "graph":
            scope = data.get("graph_scope_label") or (
                "Mail.ReadWrite" if data.get("graph_can_write") else "Mail.Read"
            )
        elif mode == "imap":
            scope = "IMAP.AccessAsUser.All"
        else:
            scope = ""
        return {
            "home_account_id": account_id,
            "username": data.get("username", "?"),
            "source": mode,
            "scope": scope,
            "updated_at": data.get("updated_at", ""),
            "can_update_token": True,
            "can_write": bool(data.get("graph_can_write", mode == "imap")),
        }

    def _get_imported_account(self, home_account_id: str) -> dict[str, Any] | None:
        return self._imported.get("accounts", {}).get(home_account_id)

    @staticmethod
    def _parse_import_line(line: str) -> dict[str, str]:
        parts = [p.strip() for p in line.strip().split("|")]
        if len(parts) != 4:
            raise RuntimeError(
                "Dòng import phải có dạng email|password|refresh_token|client_id."
            )
        email_address, _password, refresh_token, client_id = parts
        if not email_address or "@" not in email_address:
            raise RuntimeError("Email trong dòng import không hợp lệ.")
        if not refresh_token:
            raise RuntimeError("Thiếu refresh_token trong dòng import.")
        if not client_id:
            raise RuntimeError("Thiếu client_id trong dòng import.")
        return {
            "email": email_address,
            "refresh_token": refresh_token,
            "client_id": client_id,
        }

    @staticmethod
    def _line_email(line: str) -> str:
        return line.split("|", 1)[0].strip()

    def _import_line_result(self, idx: int, line: str) -> dict[str, Any]:
        email_address = self._line_email(line)
        try:
            account = self.import_refresh_token_account(line)
            return {
                "line": idx,
                "email": account.get("username") or email_address,
                "ok": True,
                "source": account.get("source", ""),
                "scope": account.get("scope", ""),
                "account": account,
            }
        except RuntimeError as exc:
            return {
                "line": idx,
                "email": email_address,
                "ok": False,
                "error": str(exc),
            }

    @staticmethod
    def _token_error(data: dict[str, Any]) -> str:
        detail = data.get("error_description") or data.get("error") or "unknown error"
        return str(detail).replace("\r", " ").replace("\n", " ")[:350]

    @staticmethod
    def _graph_import_scope_options() -> list[tuple[list[str], bool, str]]:
        return [
            (["Mail.Read", "offline_access"], False, "Mail.Read"),
            (["Mail.ReadWrite", "offline_access"], True, "Mail.ReadWrite"),
            (
                ["https://graph.microsoft.com/Mail.Read", "offline_access"],
                False,
                "https://graph.microsoft.com/Mail.Read",
            ),
            (
                ["https://graph.microsoft.com/Mail.ReadWrite", "offline_access"],
                True,
                "https://graph.microsoft.com/Mail.ReadWrite",
            ),
            (["User.Read", "Mail.Read", "offline_access"], False, "User.Read Mail.Read"),
            (
                ["User.Read", "Mail.ReadWrite", "offline_access"],
                True,
                "User.Read Mail.ReadWrite",
            ),
        ]

    def _redeem_refresh_token(
        self, client_id: str, refresh_token: str, scopes: list[str]
    ) -> dict[str, Any]:
        resp = requests.post(
            TOKEN_ENDPOINT,
            data={
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": " ".join(scopes),
            },
            timeout=30,
        )
        try:
            data = resp.json()
        except ValueError:
            data = {"error_description": resp.text[:350]}
        if resp.status_code != 200 or "access_token" not in data:
            raise RuntimeError(self._token_error(data))
        return data

    def _refresh_imported_access_token(
        self, account: dict[str, Any], scopes: list[str]
    ) -> str:
        data = self._redeem_refresh_token(
            account["client_id"], account["refresh_token"], scopes
        )
        new_refresh_token = data.get("refresh_token")
        if new_refresh_token and new_refresh_token != account.get("refresh_token"):
            with self._lock:
                account["refresh_token"] = new_refresh_token
                account["updated_at"] = self._now_iso()
                self._persist_imported()
        return data["access_token"]

    def _try_graph_import(
        self, parsed: dict[str, str], refresh_token: str
    ) -> tuple[dict[str, Any], list[str], bool, str, str]:
        errors = []
        current_refresh_token = refresh_token
        for scopes, can_write, label in self._graph_import_scope_options():
            try:
                token_data = self._redeem_refresh_token(
                    parsed["client_id"], current_refresh_token, scopes
                )
                current_refresh_token = (
                    token_data.get("refresh_token") or current_refresh_token
                )
                resp = requests.get(
                    f"{GRAPH_BASE}/me/messages?$top=1&$select=id",
                    headers={"Authorization": f"Bearer {token_data['access_token']}"},
                    timeout=30,
                )
                if resp.status_code == 200:
                    return (
                        token_data,
                        scopes,
                        can_write,
                        label,
                        current_refresh_token,
                    )
                errors.append(f"{label}: Graph messages {resp.status_code}")
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{label}: {str(exc)[:140]}")
        raise RuntimeError(" | ".join(errors[:4]))

    def import_refresh_token_account(self, line: str) -> dict[str, str]:
        """Import one user-owned Outlook account from email|password|refresh_token|client_id."""
        parsed = self._parse_import_line(line)
        mode = "graph"
        name = ""
        refresh_token = parsed["refresh_token"]
        graph_scopes: list[str] = []
        graph_can_write = False
        graph_scope_label = ""

        try:
            (
                token_data,
                graph_scopes,
                graph_can_write,
                graph_scope_label,
                refresh_token,
            ) = self._try_graph_import(parsed, refresh_token)
            profile = requests.get(
                f"{GRAPH_BASE}/me?$select=displayName",
                headers={"Authorization": f"Bearer {token_data['access_token']}"},
                timeout=30,
            )
            if profile.status_code == 200:
                name = profile.json().get("displayName") or ""
        except Exception as graph_exc:  # noqa: BLE001
            mode = "imap"
            try:
                token_data = self._redeem_refresh_token(
                    parsed["client_id"], refresh_token, IMAP_SCOPES
                )
                refresh_token = token_data.get("refresh_token") or refresh_token
                self._imap_probe(parsed["email"], token_data["access_token"])
            except Exception as imap_exc:  # noqa: BLE001
                raise RuntimeError(
                    "Không dùng được refresh token này cho Graph hoặc IMAP. "
                    f"Graph: {str(graph_exc)[:180]} | IMAP: {str(imap_exc)[:180]}"
                ) from imap_exc

        account_id = self._imported_id(parsed["email"], parsed["client_id"])
        account = {
            "username": parsed["email"],
            "name": name,
            "client_id": parsed["client_id"],
            "refresh_token": refresh_token,
            "mode": mode,
            "graph_scopes": graph_scopes,
            "graph_can_write": graph_can_write,
            "graph_scope_label": graph_scope_label,
            "created_at": self._now_iso(),
            "updated_at": self._now_iso(),
        }
        with self._lock:
            existing = self._imported.setdefault("accounts", {}).get(account_id)
            if existing:
                account["created_at"] = existing.get("created_at", account["created_at"])
            self._imported["accounts"][account_id] = account
            self._persist_imported()
        return self._public_account(account_id, account)

    def update_imported_refresh_token(
        self, home_account_id: str, refresh_token: str
    ) -> dict[str, Any]:
        account = self._get_imported_account(home_account_id)
        if not account:
            raise RuntimeError(
                "Chỉ đổi token trực tiếp cho tài khoản import. Tài khoản đăng nhập Microsoft cần xóa rồi đăng nhập lại."
            )
        refresh_token = refresh_token.strip()
        if not refresh_token:
            raise RuntimeError("Thiếu refresh_token mới.")
        line = (
            f"{account.get('username', '')}||{refresh_token}|"
            f"{account.get('client_id', '')}"
        )
        return self.import_refresh_token_account(line)

    def import_refresh_token_lines(self, lines: list[str]) -> dict[str, Any]:
        results = [self._import_line_result(idx, line) for idx, line in enumerate(lines, 1)]
        added = sum(1 for item in results if item["ok"])
        failed = len(results) - added
        return {
            "ok": failed == 0,
            "added": added,
            "failed": failed,
            "results": results,
            "account": results[0].get("account") if len(results) == 1 and added else None,
        }

    def begin_import_refresh_token_job(
        self, lines: list[str], concurrency: int = 3
    ) -> dict[str, Any]:
        workers = min(max(concurrency, 1), 5)
        workers = min(workers, len(lines))
        job_id = uuid.uuid4().hex
        now = self._now_iso()
        work_queue: Queue[tuple[int, str]] = Queue()
        for item in enumerate(lines, 1):
            work_queue.put(item)
        job: dict[str, Any] = {
            "job_id": job_id,
            "status": "queued",
            "created_at": now,
            "updated_at": now,
            "finished_at": "",
            "total": len(lines),
            "processed": 0,
            "added": 0,
            "failed": 0,
            "current_line": None,
            "current_lines": [],
            "cancel_requested": False,
            "concurrency": workers,
            "active_workers": workers,
            "results": [],
        }
        with self._lock:
            self._import_jobs[job_id] = job

        def _worker(worker_id: int) -> None:
            with self._lock:
                job["status"] = "running"
                job["updated_at"] = self._now_iso()
            try:
                while True:
                    with self._lock:
                        if job["cancel_requested"]:
                            break
                    try:
                        idx, line = work_queue.get_nowait()
                    except Empty:
                        break

                    with self._lock:
                        job.setdefault("current_lines", []).append(idx)
                        job["current_lines"] = sorted(set(job["current_lines"]))
                        job["current_line"] = job["current_lines"][0]
                        job["updated_at"] = self._now_iso()

                    result = self._import_line_result(idx, line)

                    with self._lock:
                        job["current_lines"] = [
                            item for item in job["current_lines"] if item != idx
                        ]
                        job["current_line"] = (
                            job["current_lines"][0] if job["current_lines"] else None
                        )
                        job["results"].append(result)
                        job["processed"] += 1
                        if result["ok"]:
                            job["added"] += 1
                        else:
                            job["failed"] += 1
                        job["updated_at"] = self._now_iso()
                    work_queue.task_done()
            finally:
                with self._lock:
                    job["active_workers"] -= 1
                    if job["active_workers"] == 0:
                        job["current_line"] = None
                        job["current_lines"] = []
                        job["status"] = (
                            "cancelled" if job["cancel_requested"] else "done"
                        )
                        job["finished_at"] = self._now_iso()
                        job["updated_at"] = job["finished_at"]

        for worker_id in range(workers):
            threading.Thread(target=_worker, args=(worker_id,), daemon=True).start()
        return self.import_job_status(job_id)

    def import_job_status(
        self, job_id: str, offset: int = 0, limit: int = 50
    ) -> dict[str, Any]:
        with self._lock:
            job = self._import_jobs.get(job_id)
            if not job:
                raise RuntimeError("Không tìm thấy import job.")
            safe_offset = max(offset, 0)
            safe_limit = min(max(limit, 1), 200)
            results = sorted(job["results"], key=lambda item: item["line"])
            total_results = len(results)
            page = results[safe_offset : safe_offset + safe_limit]
            return {
                "job_id": job["job_id"],
                "status": job["status"],
                "created_at": job["created_at"],
                "updated_at": job["updated_at"],
                "finished_at": job["finished_at"],
                "total": job["total"],
                "processed": job["processed"],
                "added": job["added"],
                "failed": job["failed"],
                "current_line": job["current_line"],
                "current_lines": list(job.get("current_lines", [])),
                "cancel_requested": job["cancel_requested"],
                "concurrency": job.get("concurrency", 1),
                "active_workers": job.get("active_workers", 0),
                "result_count": total_results,
                "offset": safe_offset,
                "limit": safe_limit,
                "results": page,
            }

    def cancel_import_job(self, job_id: str) -> dict[str, Any]:
        with self._lock:
            job = self._import_jobs.get(job_id)
            if not job:
                raise RuntimeError("Không tìm thấy import job.")
            if job["status"] in {"queued", "running"}:
                job["cancel_requested"] = True
                job["status"] = "cancelling"
                job["updated_at"] = self._now_iso()
        return self.import_job_status(job_id)

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
        out = []
        if CLIENT_ID:
            app = self._build_app()
            for a in app.get_accounts():
                out.append(
                    {
                        "home_account_id": a["home_account_id"],
                        "username": a.get("username", "?"),
                        "source": "graph",
                        "scope": ", ".join(SCOPES),
                        "updated_at": "",
                        "can_update_token": False,
                        "can_write": True,
                    }
                )
        for account_id, data in self._imported.get("accounts", {}).items():
            out.append(self._public_account(account_id, data))
        return out

    def remove_account(self, home_account_id: str) -> bool:
        if home_account_id in self._imported.get("accounts", {}):
            with self._lock:
                self._imported["accounts"].pop(home_account_id, None)
                self._persist_imported()
            return True
        if not CLIENT_ID:
            return False
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
        imported = self._get_imported_account(home_account_id)
        if imported:
            if imported.get("mode") != "graph":
                raise RuntimeError("Tài khoản này dùng IMAP, không có Graph token.")
            scopes = imported.get("graph_scopes") or ["Mail.Read", "offline_access"]
            return self._refresh_imported_access_token(imported, scopes)
        if not CLIENT_ID:
            raise RuntimeError("Chưa cấu hình OUTLOOK_CLIENT_ID.")
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
    # ---------- Imported IMAP accounts ----------
    def _imap_account(self, home_account_id: str) -> dict[str, Any] | None:
        account = self._get_imported_account(home_account_id)
        if account and account.get("mode") == "imap":
            return account
        return None

    @staticmethod
    def _imap_authenticate(mail: imaplib.IMAP4_SSL, username: str, token: str) -> None:
        auth = f"user={username}\x01auth=Bearer {token}\x01\x01".encode("utf-8")
        mail.authenticate("XOAUTH2", lambda _challenge: auth)

    def _imap_probe(self, username: str, access_token: str) -> None:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        try:
            self._imap_authenticate(mail, username, access_token)
            status, _ = mail.select("INBOX", readonly=True)
            if status != "OK":
                raise RuntimeError("Không mở được INBOX qua IMAP.")
        finally:
            self._imap_logout(mail)

    def _imap_connect(self, account: dict[str, Any]) -> imaplib.IMAP4_SSL:
        token = self._refresh_imported_access_token(account, IMAP_SCOPES)
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        self._imap_authenticate(mail, account["username"], token)
        return mail

    @staticmethod
    def _imap_logout(mail: imaplib.IMAP4_SSL) -> None:
        try:
            mail.logout()
        except Exception:  # noqa: BLE001
            pass

    @staticmethod
    def _imap_mailbox_arg(name: str) -> str:
        if name.upper() == "INBOX":
            return "INBOX"
        escaped = name.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'

    @staticmethod
    def _imap_folder_candidates(folder: str | None) -> list[str]:
        mapping = {
            "inbox": ["INBOX"],
            "sentitems": ["Sent", "Sent Items"],
            "drafts": ["Drafts"],
            "junkemail": ["Junk", "Junk Email"],
            "deleteditems": ["Deleted", "Deleted Items", "Trash"],
            "archive": ["Archive"],
        }
        return mapping.get((folder or "inbox").lower(), [folder or "INBOX"])

    def _imap_select(
        self, mail: imaplib.IMAP4_SSL, folder: str | None, readonly: bool = True
    ) -> str:
        errors = []
        for candidate in self._imap_folder_candidates(folder):
            status, data = mail.select(
                self._imap_mailbox_arg(candidate), readonly=readonly
            )
            if status == "OK":
                return candidate
            errors.append(f"{candidate}: {data!r}")
        raise RuntimeError("Không mở được thư mục IMAP: " + "; ".join(errors[:3]))

    @staticmethod
    def _encode_imap_id(folder: str, uid: bytes | str) -> str:
        uid_text = uid.decode("ascii", errors="ignore") if isinstance(uid, bytes) else uid
        folder_token = base64.urlsafe_b64encode(folder.encode("utf-8")).decode("ascii")
        return f"imap:{folder_token.rstrip('=')}:{uid_text}"

    @staticmethod
    def _decode_imap_id(message_id: str) -> tuple[str, str]:
        if not message_id.startswith("imap:"):
            return "INBOX", message_id
        try:
            _prefix, folder_token, uid = message_id.split(":", 2)
            padding = "=" * (-len(folder_token) % 4)
            folder = base64.urlsafe_b64decode(folder_token + padding).decode("utf-8")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError("Mã mail IMAP không hợp lệ.") from exc
        if not re.fullmatch(r"\d+", uid):
            raise RuntimeError("UID mail IMAP không hợp lệ.")
        return folder, uid

    @staticmethod
    def _part_text(part: Any) -> str:
        try:
            content = part.get_content()
            return content if isinstance(content, str) else str(content)
        except Exception:  # noqa: BLE001
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")

    def _extract_body(self, msg: Any) -> tuple[str, str]:
        plain = ""
        html_body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.is_multipart():
                    continue
                ctype = part.get_content_type()
                disposition = (part.get_content_disposition() or "").lower()
                if disposition == "attachment":
                    continue
                if ctype == "text/html" and not html_body:
                    html_body = self._part_text(part)
                elif ctype == "text/plain" and not plain:
                    plain = self._part_text(part)
        else:
            ctype = msg.get_content_type()
            if ctype == "text/html":
                html_body = self._part_text(msg)
            else:
                plain = self._part_text(msg)
        if html_body:
            return "html", html_body
        return "text", plain

    @staticmethod
    def _preview_from_body(body_type: str, body: str) -> str:
        text = re.sub(r"<[^>]+>", " ", body) if body_type == "html" else body
        text = html.unescape(re.sub(r"\s+", " ", text)).strip()
        return text[:300]

    @staticmethod
    def _attachment_parts(msg: Any) -> list[tuple[int, Any]]:
        out = []
        idx = 0
        for part in msg.walk():
            if part.is_multipart():
                continue
            disposition = (part.get_content_disposition() or "").lower()
            filename = part.get_filename()
            if filename or disposition in {"attachment", "inline"}:
                out.append((idx, part))
            idx += 1
        return out

    def _imap_fetch_raw(
        self, mail: imaplib.IMAP4_SSL, uid: bytes | str
    ) -> tuple[bytes, str]:
        uid_text = uid.decode("ascii", errors="ignore") if isinstance(uid, bytes) else uid
        status, data = mail.uid("fetch", uid_text, "(RFC822 FLAGS)")
        if status != "OK":
            raise RuntimeError(f"IMAP FETCH lỗi: {data!r}")
        raw = None
        fetch_header = ""
        for item in data:
            if isinstance(item, tuple):
                fetch_header = item[0].decode("utf-8", errors="replace")
                raw = item[1]
                break
        if raw is None:
            raise RuntimeError("Không đọc được nội dung mail từ IMAP.")
        return raw, fetch_header

    def _imap_summary(
        self, uid: bytes | str, raw: bytes, fetch_header: str, folder: str
    ) -> dict[str, Any]:
        msg = BytesParser(policy=policy.default).parsebytes(raw)
        sender_name, sender_addr = email.utils.parseaddr(str(msg.get("from", "")))
        body_type, body = self._extract_body(msg)
        received = ""
        try:
            parsed_date = email.utils.parsedate_to_datetime(str(msg.get("date", "")))
            if parsed_date:
                received = parsed_date.isoformat()
        except Exception:  # noqa: BLE001
            received = ""
        return {
            "id": self._encode_imap_id(folder, uid),
            "subject": str(msg.get("subject") or "(không tiêu đề)"),
            "from_name": sender_name,
            "from_address": sender_addr,
            "received": received,
            "preview": self._preview_from_body(body_type, body),
            "is_read": "\\Seen" in fetch_header,
            "has_attachments": bool(self._attachment_parts(msg)),
            "web_link": "",
        }

    @staticmethod
    def _imap_quote(value: str) -> str:
        return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'

    @staticmethod
    def _imap_date(value: str, add_days: int = 0) -> str:
        d = datetime.strptime(value, "%Y-%m-%d") + timedelta(days=add_days)
        return d.strftime("%d-%b-%Y")

    def _imap_search_uids(
        self,
        mail: imaplib.IMAP4_SSL,
        unread_only: bool = False,
        sender: str | None = None,
        date_from: str | None = None,
        date_to: str | None = None,
        text: str | None = None,
    ) -> list[bytes]:
        criteria = ["ALL"]
        if unread_only:
            criteria.append("UNSEEN")
        if sender:
            criteria.extend(["FROM", self._imap_quote(sender)])
        if date_from:
            criteria.extend(["SINCE", self._imap_date(date_from)])
        if date_to:
            criteria.extend(["BEFORE", self._imap_date(date_to, add_days=1)])
        if text:
            criteria.extend(["TEXT", self._imap_quote(text)])
        status, data = mail.uid("search", None, *criteria)
        if status != "OK":
            raise RuntimeError(f"IMAP SEARCH lỗi: {data!r}")
        return data[0].split() if data and data[0] else []

    def _imap_get_messages(
        self,
        account: dict[str, Any],
        count: int,
        skip: int,
        folder: str | None,
        unread_only: bool,
        has_attachments: bool,
        sender: str | None,
        date_from: str | None,
        date_to: str | None,
    ) -> dict[str, Any]:
        mail = self._imap_connect(account)
        try:
            selected_folder = self._imap_select(mail, folder, readonly=True)
            uids = list(
                reversed(
                    self._imap_search_uids(
                        mail, unread_only, sender, date_from, date_to
                    )
                )
            )
            messages = []
            consumed = 0
            fetch_limit = max(count * (8 if has_attachments else 1), count)
            for uid in uids[skip:]:
                if consumed >= fetch_limit and len(messages) < count:
                    break
                consumed += 1
                raw, fetch_header = self._imap_fetch_raw(mail, uid)
                summary = self._imap_summary(uid, raw, fetch_header, selected_folder)
                if has_attachments and not summary["has_attachments"]:
                    continue
                messages.append(summary)
                if len(messages) >= count:
                    break
            return {"messages": messages, "has_more": skip + consumed < len(uids)}
        finally:
            self._imap_logout(mail)

    def _imap_list_folders(self, account: dict[str, Any]) -> list[dict[str, Any]]:
        mail = self._imap_connect(account)
        try:
            status, data = mail.list()
            if status != "OK":
                raise RuntimeError(f"IMAP LIST lỗi: {data!r}")
            folders = []
            for row in data or []:
                text = row.decode("utf-8", errors="replace")
                match = re.search(r'"([^"]+)"\s*$', text)
                name = match.group(1) if match else text.rsplit(" ", 1)[-1]
                folders.append({"id": name, "name": name, "unread": 0, "total": 0})
            return folders
        finally:
            self._imap_logout(mail)

    def _imap_unread_count(self, account: dict[str, Any]) -> int:
        mail = self._imap_connect(account)
        try:
            self._imap_select(mail, "inbox", readonly=True)
            return len(self._imap_search_uids(mail, unread_only=True))
        finally:
            self._imap_logout(mail)

    def _imap_message_detail(
        self, account: dict[str, Any], message_id: str
    ) -> dict[str, Any]:
        folder, uid = self._decode_imap_id(message_id)
        mail = self._imap_connect(account)
        try:
            selected_folder = self._imap_select(mail, folder, readonly=True)
            raw, fetch_header = self._imap_fetch_raw(mail, uid)
            msg = BytesParser(policy=policy.default).parsebytes(raw)
            sender_name, sender_addr = email.utils.parseaddr(str(msg.get("from", "")))
            body_type, body = self._extract_body(msg)

            def _addresses(header: str) -> list[str]:
                return [
                    name or addr
                    for name, addr in email.utils.getaddresses(
                        [str(msg.get(header, ""))]
                    )
                    if name or addr
                ]

            received = ""
            try:
                parsed_date = email.utils.parsedate_to_datetime(
                    str(msg.get("date", ""))
                )
                if parsed_date:
                    received = parsed_date.isoformat()
            except Exception:  # noqa: BLE001
                received = ""
            return {
                "id": self._encode_imap_id(selected_folder, uid),
                "subject": str(msg.get("subject") or "(không tiêu đề)"),
                "from_name": sender_name,
                "from_address": sender_addr,
                "to": _addresses("to"),
                "cc": _addresses("cc"),
                "received": received,
                "body_type": body_type,
                "body": body,
                "is_read": "\\Seen" in fetch_header,
                "has_attachments": bool(self._attachment_parts(msg)),
                "web_link": "",
            }
        finally:
            self._imap_logout(mail)

    def _imap_list_attachments(
        self, account: dict[str, Any], message_id: str
    ) -> list[dict[str, Any]]:
        folder, uid = self._decode_imap_id(message_id)
        mail = self._imap_connect(account)
        try:
            self._imap_select(mail, folder, readonly=True)
            raw, _ = self._imap_fetch_raw(mail, uid)
            msg = BytesParser(policy=policy.default).parsebytes(raw)
            out = []
            for idx, part in self._attachment_parts(msg):
                payload = part.get_payload(decode=True) or b""
                out.append(
                    {
                        "id": str(idx),
                        "name": part.get_filename() or f"attachment-{idx}",
                        "content_type": part.get_content_type(),
                        "size": len(payload),
                        "is_inline": (part.get_content_disposition() or "").lower()
                        == "inline",
                    }
                )
            return out
        finally:
            self._imap_logout(mail)

    def _imap_get_attachment(
        self, account: dict[str, Any], message_id: str, attachment_id: str
    ) -> dict[str, Any]:
        folder, uid = self._decode_imap_id(message_id)
        mail = self._imap_connect(account)
        try:
            self._imap_select(mail, folder, readonly=True)
            raw, _ = self._imap_fetch_raw(mail, uid)
            msg = BytesParser(policy=policy.default).parsebytes(raw)
            for idx, part in self._attachment_parts(msg):
                if str(idx) == attachment_id:
                    return {
                        "name": part.get_filename() or f"attachment-{idx}",
                        "content_type": part.get_content_type(),
                        "bytes": part.get_payload(decode=True) or b"",
                    }
            raise RuntimeError("Không tìm thấy file đính kèm.")
        finally:
            self._imap_logout(mail)

    def _imap_mark_read(
        self, account: dict[str, Any], message_id: str, is_read: bool
    ) -> bool:
        folder, uid = self._decode_imap_id(message_id)
        mail = self._imap_connect(account)
        try:
            self._imap_select(mail, folder, readonly=False)
            flag_op = "+FLAGS" if is_read else "-FLAGS"
            status, data = mail.uid("store", uid, flag_op, "(\\Seen)")
            if status != "OK":
                raise RuntimeError(f"IMAP STORE lỗi: {data!r}")
            return True
        finally:
            self._imap_logout(mail)

    def _imap_search_messages(
        self, account: dict[str, Any], query: str, count: int
    ) -> list[dict[str, Any]]:
        mail = self._imap_connect(account)
        try:
            selected_folder = self._imap_select(mail, "inbox", readonly=True)
            uids = list(reversed(self._imap_search_uids(mail, text=query)))
            out = []
            for uid in uids[:count]:
                raw, fetch_header = self._imap_fetch_raw(mail, uid)
                out.append(self._imap_summary(uid, raw, fetch_header, selected_folder))
            return out
        finally:
            self._imap_logout(mail)

    def list_folders(self, home_account_id: str) -> list[dict[str, Any]]:
        """Danh sách thư mục mail + số mail chưa đọc của mỗi thư mục."""
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_list_folders(imap_account)
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_unread_count(imap_account)
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_get_messages(
                imap_account,
                count=count,
                skip=skip,
                folder=folder,
                unread_only=unread_only,
                has_attachments=has_attachments,
                sender=sender,
                date_from=date_from,
                date_to=date_to,
            )
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_message_detail(imap_account, message_id)
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_list_attachments(imap_account, message_id)
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

        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_get_attachment(imap_account, message_id, attachment_id)
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_mark_read(imap_account, message_id, is_read)
        imported = self._get_imported_account(home_account_id)
        if imported and imported.get("mode") == "graph" and not imported.get(
            "graph_can_write", True
        ):
            raise RuntimeError(
                "Refresh token này chỉ có quyền Mail.Read; cần Mail.ReadWrite để đánh dấu đã đọc/chưa đọc."
            )
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
        imap_account = self._imap_account(home_account_id)
        if imap_account:
            return self._imap_search_messages(imap_account, query, count)
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
