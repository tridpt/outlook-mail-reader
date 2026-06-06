"""Các route API cho hộp thư Outlook đa tài khoản."""
from __future__ import annotations

from fastapi import APIRouter, HTTPException

from config import CLIENT_ID, DEFAULT_MAIL_COUNT

router = APIRouter(prefix="/api/outlook", tags=["outlook"])


@router.get("/status")
def status() -> dict:
    """Trạng thái cấu hình + số tài khoản đã đăng nhập."""
    configured = bool(CLIENT_ID)
    accounts: list = []
    if configured:
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


@router.delete("/accounts/{home_account_id}")
def remove_account(home_account_id: str) -> dict:
    from engine import engine

    ok = engine.remove_account(home_account_id)
    if not ok:
        raise HTTPException(status_code=404, detail="Không tìm thấy tài khoản.")
    return {"ok": True}


@router.get("/inbox")
def unified_inbox(count: int = DEFAULT_MAIL_COUNT, unread_only: bool = False) -> dict:
    """Hộp thư gộp của tất cả tài khoản đã đăng nhập."""
    from engine import engine

    if not engine.list_accounts():
        return {"inboxes": []}
    return {"inboxes": engine.get_unified_inbox(count=count, unread_only=unread_only)}


@router.get("/inbox/{home_account_id}")
def single_inbox(
    home_account_id: str, count: int = DEFAULT_MAIL_COUNT, unread_only: bool = False
) -> dict:
    from engine import engine

    try:
        return {
            "messages": engine.get_messages(
                home_account_id, count=count, unread_only=unread_only
            )
        }
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


@router.post("/message/{home_account_id}/{message_id}/read")
def mark_read(home_account_id: str, message_id: str, is_read: bool = True) -> dict:
    """Đánh dấu mail đã đọc / chưa đọc."""
    from engine import engine

    try:
        engine.mark_read(home_account_id, message_id, is_read=is_read)
        return {"ok": True, "is_read": is_read}
    except RuntimeError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
