# 📬 Outlook Mail Reader — hộp thư đa tài khoản

Web app gộp mail của **nhiều tài khoản Outlook / Microsoft 365** vào một chỗ.
Mỗi tài khoản chỉ cần **đăng nhập 1 lần**, sau đó app tự lấy mail các lần sau
mà không phải đăng nhập lại.

> 🔐 **An toàn & hợp lệ:** Dùng OAuth 2.0 chuẩn của Microsoft (Microsoft Graph).
> **Không lưu mật khẩu.** Mỗi tài khoản vẫn phải qua màn đăng nhập Microsoft thật
> đúng 1 lần — nên chỉ dùng được với tài khoản bạn sở hữu hoặc có quyền truy cập.
> Quyền yêu cầu là `Mail.Read` (chỉ đọc). Token được lưu **mã hóa** dưới đĩa.

## Cấu trúc

```
outlook-mail-reader/
├── backend/
│   ├── config.py    # client_id, scope, đường dẫn
│   ├── store.py     # lưu token cache đã mã hóa (Fernet/AES)
│   ├── engine.py    # đăng nhập device-code, tự refresh, gọi Graph API
│   ├── routes.py    # API /api/outlook/*
│   └── main.py      # FastAPI app + phục vụ frontend
└── frontend/
    ├── index.html
    ├── style.css
    └── app.js
```

## Thiết lập Azure (làm 1 lần)

App đọc mail bắt buộc phải đăng ký với Microsoft để được cấp quyền OAuth.

1. Vào <https://portal.azure.com> → **App registrations** → **New registration**
   - Supported account types: *Accounts in any organizational directory and personal Microsoft accounts*
   - Bấm **Register**
2. Mở app vừa tạo → **Authentication** → **Advanced settings** →
   bật **Allow public client flows** = **Yes** → Save
3. Copy **Application (client) ID**

## Chạy ở máy local

```bash
cd outlook-mail-reader
python -m venv .venv

# Windows (cmd)
.venv\Scripts\activate
set OUTLOOK_CLIENT_ID=<client-id-cua-ban>

# PowerShell
# .venv\Scripts\Activate.ps1
# $env:OUTLOOK_CLIENT_ID="<client-id-cua-ban>"

# macOS / Linux
# source .venv/bin/activate
# export OUTLOOK_CLIENT_ID=<client-id-cua-ban>

pip install -r requirements.txt
cd backend
uvicorn main:app --reload --host 0.0.0.0 --port 8809
```

Mở trình duyệt: <http://localhost:8809>

## Dùng

1. Bấm **+ Thêm tài khoản** → mở link, nhập mã hiển thị, đăng nhập Microsoft
2. Lặp lại cho từng tài khoản (mỗi cái chỉ 1 lần)
3. Bấm **🔄 Tải mail** để xem hộp thư gộp của tất cả tài khoản

Token mã hóa lưu tại `storage/`. Xóa thư mục này nếu muốn đăng xuất hết tài khoản.

## API

| Method | Đường dẫn | Mô tả |
|--------|-----------|-------|
| GET | `/api/outlook/status` | Trạng thái cấu hình + danh sách tài khoản |
| POST | `/api/outlook/login/start` | Bắt đầu đăng nhập (trả về code + link) |
| GET | `/api/outlook/login/status/{job_id}` | Poll trạng thái đăng nhập |
| GET | `/api/outlook/accounts` | Danh sách tài khoản |
| DELETE | `/api/outlook/accounts/{id}` | Xóa tài khoản khỏi cache |
| GET | `/api/outlook/inbox` | Hộp thư gộp tất cả tài khoản |
| GET | `/api/outlook/inbox/{id}` | Hộp thư của 1 tài khoản |

## Hướng phát triển tiếp

- Xem nội dung đầy đủ của từng mail (hiện chỉ có preview)
- Đánh dấu đã đọc / lưu trữ
- Tìm kiếm mail trên nhiều tài khoản cùng lúc
- Thông báo khi có mail mới
