"""FastAPI backend: đọc hộp thư Outlook đa tài khoản.

Đăng nhập 1 lần mỗi tài khoản (OAuth/Microsoft Graph), sau đó xem gộp mail
mà không phải đăng nhập lại. Không lưu mật khẩu.
"""
from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from config import CLIENT_ID
from routes import router as outlook_router

app = FastAPI(title="Outlook Mail Reader", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(outlook_router)


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "configured": bool(CLIENT_ID)}


# Phục vụ frontend tĩnh (đặt cuối để không che các route /api)
BASE_DIR = Path(__file__).resolve().parent.parent
frontend_dir = BASE_DIR / "frontend"
app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
