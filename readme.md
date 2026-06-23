# 51LA UV/PV Scraper

Scrape UV/PV data from 51.la and send daily reports via Gmail.

## Quy tắc quan trọng

- **KHÔNG BAO GIỜ** chạy lệnh có `--allow-real-email` trừ khi đã xác nhận Excel đúng ngày
- Luôn test bằng `--dry-run` hoặc `--no-notification` trước
- Không chạy song song 2 instance — dùng PM2 để quản lý

---

## 0. Chạy từ đúng thư mục

**QUAN TRỌNG:** Chạy PM2 từ đúng thư mục chứa code mới nhất:

```
D:\Download\manager\updated-mail-send\merged_51la
```

Không phải:
```
D:\Administrator\Downloads\merged_51la\merged_51la   ← THƯ MỤC CŨ, KHÔNG DÙNG
```

Kiểm tra đường dẫn hiện tại:
```bash
cd D:\Download\manager\updated-mail-send\merged_51la
pwd
```

---

## 1. Chạy một lần (one-time)

### Test (không gửi email):
```bash
# Chạy với pre-created monthly workbook, ngày hôm qua
py -3 main.py --method gmail --lang zh --dry-run --precreated-monthly --no-notification

# Test với ngày cụ thể (ví dụ: 2026-06-01)
py -3 main.py --method gmail --lang zh --dry-run --precreated-monthly --report-date 2026-06-01 --no-notification

# Test simulation: giả lập ngày chạy là 2026-06-01 → computed report_date = 2026-05-31
py -3 main.py --method gmail --lang zh --dry-run --precreated-monthly --run-date 2026-06-01 --no-notification
```

### Chạy thật (gửi email):
```bash
# Gửi Gmail thật — chạy full flow scrape + Excel + email
py -3 main.py --method gmail --lang zh --precreated-monthly --allow-real-email
```

---

## 2. Chạy theo lịch với PM2 (recommended)

### 2.1 Cài đặt PM2 (nếu chưa có)
```bash
npm install -g pm2
```

### 2.2 Khởi động process

**Bước 1:** Di chuyển đến đúng thư mục:
```bash
cd D:\Download\manager\updated-mail-send\merged_51la
```

**Bước 2:** Chạy PM2:
```bash
# Chạy lịch hằng ngày lúc 8:50 AM (Vietnam time) — GỬI EMAIL THẬT
pm2 start py --name "51la-daily" -- main.py --schedule-mode --method gmail --lang zh --precreated-monthly --allow-real-email

# Chạy lịch hằng ngày nhưng KHÔNG gửi email thật (test trước)
pm2 start py --name "51la-daily-test" -- main.py --schedule-mode --method gmail --lang zh --precreated-monthly
```

**Ghi chú:** `--schedule-mode` là flag để chạy loop vô tận kiểm tra lịch mỗi 30 giây. KHÔNG phải `--daily` (vì PM2 có flag `--daily` built-in sẽ bị intercept).

### 2.3 Quản lý PM2
```bash
# Xem logs realtime
pm2 logs 51la-daily

# Xem status
pm2 status

# Restart (sau khi sửa code)
pm2 restart 51la-daily

# Stop (tạm dừng)
pm2 stop 51la-daily

# Xóa khỏi PM2
pm2 delete 51la-daily
```

### 2.4 Tự động khởi động lại khi reboot
```bash
pm2 startup
pm2 save
```

---

## 3. Chạy không liên tục (Windows Task Scheduler)

Nếu không muốn PM2 loop, dùng Windows Task Scheduler chạy một lần mỗi ngày:

```bash
# Chạy lúc 8:50 AM hằng ngày
py -3 main.py --method gmail --lang zh --precreated-monthly --allow-real-email
```

Cài đặt Task Scheduler:
```powershell
# Mở Task Scheduler (Win + R → taskschd.msc)
# Tạo Task với:
#   Trigger: Daily at 8:50 AM
#   Action: Start a program
#   Program: py.exe
#   Arguments: -3 main.py --method gmail --lang zh --precreated-monthly --allow-real-email
#   Working directory: D:\Download\manager\updated-mail-send\merged_51la
```

---

## 4. Các cờ quan trọng

| Cờ | Mô tả |
|----|-------|
| `--schedule-mode` | Loop vô tận kiểm tra lịch mỗi 30s (DÙNG flag này cho PM2) |
| `--precreated-monthly` | Dùng pre-created monthly workbook từ `reports/` |
| `--report-date YYYY-MM-DD` | Override ngày báo cáo (trực tiếp) |
| `--run-date YYYY-MM-DD` | Giả lập ngày chạy → report_date = run_date - 1 ngày |
| `--dry-run` | Không gửi email, chỉ hiển thị kế hoạch |
| `--no-notification` | Scrape + ghi Excel nhưng KHÔNG gửi email |
| `--no-excel-write` | Scrape nhưng KHÔNG ghi Excel |
| `--allow-real-email` | Cho phép gửi email thật (cần `ALLOW_REAL_EMAIL_SEND=true` trong .env) |
| `--legacy-excel` | Dùng legacy `51la.xlsx` thay vì pre-created monthly |

---

## 5. Các chế độ chạy

### 5.1 Pre-created Monthly (mặc định, KHUYÊN DÙNG)
- File Excel: `reports/51la_YYYY-MM.xlsx`
- Ngày báo cáo = ngày hôm qua (`yesterday` trong Asia/Ho_Chi_Minh)
- Nếu thiếu file hoặc thiếu dòng ngày → **BLOCKED**, không tự tạo

### 5.2 Legacy Mode
```bash
py -3 main.py --method gmail --lang zh --legacy-excel --allow-real-email
```
- File Excel: `51la.xlsx` (như cũ)
- Tự động tạo dòng ngày nếu thiếu

### 5.3 Auto-create Monthly (KHÔNG khuyến khích)
```bash
py -3 main.py --method gmail --lang zh --auto-create-monthly --allow-real-email
```
- Tự tạo monthly workbook từ template
- Có thể làm hỏng layout Excel tùy chỉnh

---

## 6. Xử lý sự cố

### Lỗi "unknown option '--daily'"
→ PM2 đang chạy từ thư mục CŨ. Di chuyển đến `D:\Download\manager\updated-mail-send\merged_51la` và chạy lại.

### Lỗi "Monthly workbook not found"
→ Tạo file `reports/51la_YYYY-MM.xlsx` theo template mẫu

### Lỗi "Date row not found"
→ Thêm dòng ngày vào đúng vị trí trong file Excel (column B)

### Lỗi "Another instance is already running"
```bash
# Xóa lock file thủ công
rm .running
```

### Không gửi được email
```bash
# Kiểm tra ALLOW_REAL_EMAIL_SEND trong .env
grep ALLOW_REAL_EMAIL_SEND .env
# Phải là: ALLOW_REAL_EMAIL_SEND=true
```

### Excel bị lỗi format
→ Dùng pre-created monthly mode (`--precreated-monthly`) thay vì auto-create

---

## 7. Bảo mật

- **KHÔNG BAO GIỜ** để `ALLOW_REAL_EMAIL_SEND=true` khi test
- **LUÔN** dùng `--dry-run` hoặc `--no-notification` khi test
- Nếu gửi nhầm email thật với dữ liệu sai → thông báo người nhận và xóa email đã gửi
- Password app Gmail: chỉ lưu trong file `.env`, không commit lên git