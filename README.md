# 📬 Outlook Mail Reader — hộp thư đa tài khoản

Web app gộp mail của **nhiều tài khoản Outlook / Microsoft 365** vào một chỗ.
Mỗi tài khoản chỉ cần **đăng nhập 1 lần**, sau đó app tự lấy mail các lần sau
mà không phải đăng nhập lại.

**Tính năng:**
- Gộp hộp thư nhiều tài khoản (unified inbox)
- Quản lý account: xem nguồn Graph/IMAP, scope, thời gian cập nhật token, đổi token, xóa account
- Xem đầy đủ nội dung từng mail (HTML render trong iframe sandbox an toàn)
- Tìm kiếm mail theo từ khóa trên tất cả tài khoản cùng lúc
- Đánh dấu đã đọc / chưa đọc, lọc riêng mail chưa đọc
- Import refresh token hàng loạt kèm log lỗi theo từng dòng
- Import queue cho danh sách lớn: chọn 1-5 luồng xử lý, progress, phân trang log, hủy job
- Copy hoặc retry riêng các dòng import lỗi
- Check email theo địa chỉ: kiểm tra format + DNS/MX + provider, không xác minh mailbox cụ thể

> 🔐 **An toàn & hợp lệ:** Dùng OAuth 2.0 chuẩn của Microsoft (Microsoft Graph).
> **Không lưu mật khẩu.** Mỗi tài khoản vẫn phải qua màn đăng nhập Microsoft thật
> đúng 1 lần — nên chỉ dùng được với tài khoản bạn sở hữu hoặc có quyền truy cập.
> Quyền yêu cầu là `Mail.ReadWrite` (đọc mail + đánh dấu đã đọc, không gửi/xóa).
> Token được lưu **mã hóa** dưới đĩa.

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

### Chạy nhanh trên Windows

```bat
start_outlook_reader.bat
```

File này tự tạo `.venv` nếu chưa có, cài `requirements.txt`, set
`OUTLOOK_CLIENT_ID`, rồi chạy server ở port `8809`.

## Dùng

1. Bấm **+ Thêm tài khoản** → mở link, nhập mã hiển thị, đăng nhập Microsoft
2. Lặp lại cho từng tài khoản (mỗi cái chỉ 1 lần)
3. Hoặc mở **Nhập refresh token** và dán một hay nhiều dòng
   `email|password|refresh_token|client_id`; app bỏ qua password và chỉ lưu token mã hóa
   - Chọn **Luồng** 1-5 để kiểm soát tốc độ import
   - Sau khi import xong có thể **Copy dòng lỗi** hoặc **Retry lỗi**
4. Bấm **🔄 Tải mail** để xem hộp thư gộp của tất cả tài khoản
5. Bấm vào một mail để xem nội dung đầy đủ; trong cửa sổ chi tiết có nút
   đánh dấu đã đọc / chưa đọc
6. Dùng ô tìm kiếm để tìm mail trên mọi tài khoản; tích "Chỉ chưa đọc" để lọc

Token mã hóa lưu tại `storage/`. Xóa thư mục này nếu muốn đăng xuất hết tài khoản.

## API

| Method | Đường dẫn | Mô tả |
|--------|-----------|-------|
| GET | `/api/outlook/status` | Trạng thái cấu hình + danh sách tài khoản |
| POST | `/api/outlook/login/start` | Bắt đầu đăng nhập (trả về code + link) |
| GET | `/api/outlook/login/status/{job_id}` | Poll trạng thái đăng nhập |
| GET | `/api/outlook/accounts` | Danh sách tài khoản |
| POST | `/api/outlook/check-emails` | Check nhiều email theo format + DNS/MX |
| POST | `/api/outlook/accounts/import-refresh-token` | Import một hoặc nhiều dòng refresh token |
| POST | `/api/outlook/accounts/import-refresh-token/job` | Tạo import job chạy nền (`concurrency` 1-5) |
| GET | `/api/outlook/accounts/import-jobs/{job_id}` | Poll progress + log phân trang (`offset`, `limit`) |
| POST | `/api/outlook/accounts/import-jobs/{job_id}/cancel` | Hủy import job |
| POST | `/api/outlook/accounts/{id}/refresh-token` | Đổi refresh token cho tài khoản import |
| DELETE | `/api/outlook/accounts/{id}` | Xóa tài khoản khỏi cache |
| GET | `/api/outlook/inbox` | Hộp thư gộp (tham số `unread_only`) |
| GET | `/api/outlook/inbox/{id}` | Hộp thư của 1 tài khoản |
| GET | `/api/outlook/search?q=` | Tìm mail trên tất cả tài khoản |
| GET | `/api/outlook/message/{id}/{msg}` | Nội dung đầy đủ 1 mail |
| POST | `/api/outlook/message/{id}/{msg}/read` | Đánh dấu đã đọc / chưa đọc |

## Hướng phát triển tiếp

- Đọc mail to bằng giọng AI (text-to-speech)
- Lưu trữ / xóa mail
- Thông báo khi có mail mới
- Soạn & gửi mail
