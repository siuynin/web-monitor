# Chạy Website Monitor 24/24 miễn phí trên GitHub Actions

GitHub sẽ tự kiểm tra các website trong `website serve.xlsx` khoảng **5–15 phút/lần**
(lịch của GitHub thường trễ vài phút) và gửi email khi có website lỗi hoặc hoạt động lại.

> Repo phải để **Public** thì mới chạy miễn phí không giới hạn. Ai cũng xem được danh sách
> domain, nhưng mật khẩu email được giấu trong *Secrets* nên không ai thấy.
> File `config.ini` (chứa mật khẩu) đã nằm trong `.gitignore`, sẽ không bị đưa lên.

## Bước 1 – Tạo mật khẩu ứng dụng Gmail
1. Bật xác minh 2 bước cho tài khoản Gmail dùng để gửi.
2. Vào https://myaccount.google.com/apppasswords → tạo mật khẩu (16 ký tự).

## Bước 2 – Tạo repo và đưa code lên
1. Vào https://github.com/new → đặt tên (vd `website-monitor`) → chọn **Public** → *Create repository*
   (không tick tạo README).
2. Mở terminal trong thư mục `monitor-web` rồi chạy (thay `TEN-BAN`):

```bash
git init -b main
git add .
git commit -m "Website monitor"
git remote add origin https://github.com/TEN-BAN/website-monitor.git
git push -u origin main
```

## Bước 3 – Khai báo Secrets
Repo → **Settings → Secrets and variables → Actions → New repository secret**, tạo:

| Tên            | Giá trị ví dụ                         | Bắt buộc |
|----------------|---------------------------------------|----------|
| `SMTP_USERNAME`| `ban@gmail.com`                       | ✅ |
| `SMTP_PASSWORD`| mật khẩu ứng dụng 16 ký tự            | ✅ |
| `ALERT_TO`     | `a@gmail.com, b@congty.vn`            | ✅ |
| `SMTP_HOST`    | `smtp.gmail.com` (mặc định)           | |
| `SMTP_PORT`    | `587` (mặc định)                      | |
| `SMTP_USE_SSL` | `true` nếu dùng port 465              | |
| `MAIL_FROM`    | địa chỉ người gửi, mặc định = username | |

## Bước 4 – Bật và chạy thử
1. Repo → tab **Actions** → nếu GitHub hỏi thì bấm bật workflows.
2. Chọn **Website Monitor** → **Run workflow** để chạy ngay.
3. Bấm vào lượt chạy để xem bảng trạng thái các website.

Từ đó GitHub tự chạy, không cần bật máy tính.

## Thêm / bớt website
Sửa `website serve.xlsx` trên máy rồi chạy:

```bash
git add "website serve.xlsx"
git commit -m "Cập nhật danh sách website"
git pull --rebase
git push
```

Hoặc trên GitHub: **Add file → Upload files** → tải file Excel mới lên (cùng tên) → *Commit changes*.

## Thay đổi cài đặt (chu kỳ nhắc, số email tối đa…)
Sửa `config.example.ini` (trên GitHub dùng file này, không dùng `config.ini`), rồi commit/push.
Muốn đổi tần suất kiểm tra thì sửa dòng `cron` trong `.github/workflows/monitor.yml`
(5 phút là nhanh nhất GitHub cho phép).

## Ghi chú
- Trạng thái (`state.json`) và lịch sử lỗi (`history.csv`) được bot tự commit về repo khi có thay đổi
  → giới hạn 5 email/sự cố và email "đã khôi phục" vẫn hoạt động đúng. Xem lịch sử lỗi ngay trên GitHub.
- GitHub tắt lịch chạy nếu repo 60 ngày không có hoạt động; workflow tự tạo 1 commit "Keepalive"
  sau 50 ngày để tránh việc này.
- Nếu một lượt chạy bị lỗi, GitHub sẽ gửi email báo cho chủ repo.
