"""Các route API cho hộp thư Outlook đa tài khoản."""
from __future__ import annotations

from fastapi import APIRouter, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel

from config import CLIENT_ID, DEFAULT_MAIL_COUNT

router = APIRouter(prefix="/api/outlook", tags=["outlook"])


class ImportRefreshTokenRequest(BaseModel):
    line: str
    concurrency: int = 3


class UpdateRefreshTokenRequest(BaseModel):
    refresh_token: str


class CheckEmailsRequest(BaseModel):
    emails: str


# Map tên thư mục thân thiện -> well-known folder của Graph
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sentitems",
    "drafts": "drafts",
    "junk": "junkemail",
    "deleted": "deleteditems",
    "archive": "archive",
}


@router.get("/status")
def status() -> dict:
    """Trạng thái cấu hình + số tài khoản đã đăng nhập."""
    configured = bool(CLIENT_ID)
    from engine import engine

    accounts = engine.list_accounts()
    return {"configured": configured, "accounts": accounts}


@router.post("/login/start")
def login_start() -> dict:
    """Bắt đầu đăng nhập 1 tài khoản (device code flow)."""
    from engine import engine

    try:
        return engine.begin_login()
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@router.get("/login/status/{job_id}")
def login_poll(job_id: str) -> dict:
    """Client poll để biết người dùng đã hoàn tất đăng nhập chưa."""
    from engine import engine

    return engine.login_status(job_id)


@router.get("/accounts")
def accounts() -> dict:
    from engine import engine

    return {"accounts": engine.list_accounts()}


@router.post("/check-emails")
def check_emails(payload: CheckEmailsRequest) -> dict:
    from engine import engine

    addresses = [line.strip() for line in payload.emails.splitlines() if line.strip()]
    if not addresses:
        raise HTTPException(status_code=400, detail="Thiếu địa chỉ email để kiểm tra.")
    return engine.check_email_addresses(addresses)


@router.post("/accounts/import-refresh-token")
def import_refresh_token(payload: ImportRefreshTokenRequest) -> dict:
    from engine import engine

    lines = [line.strip() for line in payload.line.splitlines() if line.strip()]
    if not lines:
        raise HTTPException(
            status_code=400,
            detail="Thiếu dòng import email|password|refresh_token|client_id.",
        )

    return engine.import_refresh_token_lines(lines)


@router.post("/accounts/import-refresh-token/job")
def start_import_refresh_token_job(payload: ImportRefreshTokenRequest) -> dict:
    from engine import engine

    lines = [line.strip() for line in payload.line.splitlines() if line.strip()]
    if not lines:
        raise HTTPException(
            status_code=400,
            detail="Thiếu dòng import email|password|refresh_token|client_id.",
        )
    return engine.begin_import_refresh_token_job(
        lines, concurrency=payload.concurrency
    )


@router.get("/accounts/import-jobs/{job_id}")
def import_job_status(job_id: str, offset: int = 0, limit: int = 50) -> dict:
    from engine import engine

    try:
        return engine.import_job_status(job_id, offset=offset, limit=limit)
    except RuntimeError as exc:
        raise HTTPException(status_code=404, detail=str(exc))


@router.post("/accounts/import-jobs/{job_id}/cancel")
def cancel_import_job(job_id: str) -> dict:
    from engine import engine

    try:
        return engine.cancel_import_job(job_id)
    except RuntimeError as exc:
        raise HTTPException(status_code=404, detail=str(exc))


@router.post("/accounts/{home_account_id}/refresh-token")
def update_refresh_token(
    home_account_id: str, payload: UpdateRefreshTokenRequest
) -> dict:
    from engine import engine

    try:
        account = engine.update_imported_refresh_token(
            home_account_id, payload.refresh_token
        )
        return {"ok": True, "account": account}
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@router.delete("/accounts/{home_account_id}")
def remove_account(home_account_id: str) -> dict:
    from engine import engine

    ok = engine.remove_account(home_account_id)
    if not ok:
        raise HTTPException(status_code=404, detail="Không tìm thấy tài khoản.")
    return {"ok": True}


@router.get("/unread-counts")
def unread_counts() -> dict:
    """Số mail chưa đọc trong Inbox của từng tài khoản (cho badge)."""
    from engine import engine

    return {"counts": engine.account_unread_counts()}


@router.get("/inbox")
def unified_inbox(
    count: int = DEFAULT_MAIL_COUNT,
    skip: int = 0,
    folder: str = "inbox",
    unread_only: bool = False,
    has_attachments: bool = False,
    sender: str = "",
    date_from: str = "",
    date_to: str = "",
    account: str = "",
) -> dict:
    """Hộp thư gộp (theo thư mục, có lọc & phân trang)."""
    from engine import engine

    if not engine.list_accounts():
        return {"inboxes": []}
    graph_folder = FOLDER_MAP.get(folder, "inbox")
    return {
        "inboxes": engine.get_unified_inbox(
            count=count,
            skip=skip,
            folder=graph_folder,
            unread_only=unread_only,
            has_attachments=has_attachments,
            sender=sender or None,
            date_from=date_from or None,
            date_to=date_to or None,
            only_account=account or None,
        )
    }


@router.get("/folders/{home_account_id}")
def folders(home_account_id: str) -> dict:
    """Danh sách thư mục mail của 1 tài khoản."""
    from engine import engine

    try:
        return {"folders": engine.list_folders(home_account_id)}
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@router.get("/search")
def search(q: str, count: int = DEFAULT_MAIL_COUNT) -> dict:
    """Tìm mail theo từ khóa trên tất cả tài khoản."""
    from engine import engine

    q = (q or "").strip()
    if not q:
        raise HTTPException(status_code=400, detail="Thiếu từ khóa tìm kiếm.")
    if not engine.list_accounts():
        return {"results": []}
    return {"results": engine.search_messages(q, count=count)}


@router.get("/message/{home_account_id}/{message_id}")
def message_detail(home_account_id: str, message_id: str) -> dict:
    """Nội dung đầy đủ của 1 mail."""
    from engine import engine

    try:
        return engine.get_message_detail(home_account_id, message_id)
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@router.get("/message/{home_account_id}/{message_id}/attachments")
def list_attachments(home_account_id: str, message_id: str) -> dict:
    """Danh sách file đính kèm của 1 mail."""
    from engine import engine

    try:
        return {"attachments": engine.list_attachments(home_account_id, message_id)}
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@router.get("/message/{home_account_id}/{message_id}/attachments/{attachment_id}")
def download_attachment(
    home_account_id: str, message_id: str, attachment_id: str
) -> Response:
    """Tải 1 file đính kèm về."""
    from urllib.parse import quote

    from engine import engine

    try:
        att = engine.get_attachment(home_account_id, message_id, attachment_id)
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    filename = quote(att["name"])
    return Response(
        content=att["bytes"],
        media_type=att["content_type"],
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename}"
        },
    )


@router.post("/message/{home_account_id}/{message_id}/read")
def mark_read(home_account_id: str, message_id: str, is_read: bool = True) -> dict:
    """Đánh dấu mail đã đọc / chưa đọc."""
    from engine import engine

    try:
        engine.mark_read(home_account_id, message_id, is_read=is_read)
        return {"ok": True, "is_read": is_read}
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
