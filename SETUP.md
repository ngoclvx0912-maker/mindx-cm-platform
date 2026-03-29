# SETUP — MindX CM Report Platform

Hướng dẫn thiết lập nền tảng báo cáo nội bộ MindX (2 tab: Action Plan + Weekly Report).

---

## Yêu cầu

- Tài khoản Google (có quyền truy cập Google Sheets)
- Google Spreadsheet ID: `1beULgTt53o_mXun8ImVGoqVATMI2TJkpa2cASsPkoN8`
- 2 Worksheet đã tồn tại:
  - `ActionPlan` (ID: 294175968) — lưu dữ liệu Action Plan
  - `T3_WeeklyReport` (ID: 160877044) — lưu dữ liệu Weekly Report

---

## Bước 1: Mở Google Apps Script

1. Mở Google Spreadsheet theo ID trên
2. Vào **Extensions → Apps Script**
3. Xoá toàn bộ code mặc định trong editor

---

## Bước 2: Copy code Apps Script

1. Mở file `google-apps-script.js` trong thư mục dự án
2. Copy **toàn bộ nội dung** file
3. Paste vào Google Apps Script editor
4. Lưu lại (Ctrl+S hoặc nút Save)

---

## Bước 3: Khởi tạo Headers (chạy 1 lần)

1. Trong Apps Script editor, chọn hàm `setupSheetHeaders` trong dropdown
2. Nhấn **Run**
3. Cấp quyền khi được yêu cầu (Allow → Continue → Allow)
4. Kiểm tra log — sẽ thấy thông báo headers đã được set cho `ActionPlan` và `T3_WeeklyReport`

> **Lưu ý:** Nếu sheet đã có headers rồi, hàm sẽ bỏ qua và ghi log "Headers đã tồn tại".

---

## Bước 4: Deploy Web App

1. Nhấn nút **Deploy → New deployment**
2. Chọn Type: **Web app**
3. Cấu hình:
   - **Description:** MindX CM Report Platform v2
   - **Execute as:** Me (your account)
   - **Who has access:** Anyone
4. Nhấn **Deploy**
5. **Copy URL** hiện ra (dạng `https://script.google.com/macros/s/XXXX.../exec`)

---

## Bước 5: Cập nhật URL vào app.js

1. Mở file `app.js` trong thư mục dự án
2. Tìm dòng:
   ```js
   const APPS_SCRIPT_URL = '__APPS_SCRIPT_URL__';
   ```
3. Thay `'__APPS_SCRIPT_URL__'` bằng URL đã copy ở Bước 4:
   ```js
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/XXXX.../exec';
   ```
4. Lưu file

---

## Bước 6: Deploy lại website

Nếu bạn đang dùng S3 hosting, chạy lại lệnh deploy để cập nhật file `app.js` mới.

---

## Cấu trúc dữ liệu

### Tab Action Plan (`ActionPlan` worksheet)
| Cột | Mô tả |
|-----|-------|
| bu | Tên BU |
| month | Tháng (YYYY-MM) |
| week | Tuần (1-4) |
| func | Function (GROWTH / OPTIMIZE / OPS) |
| kpi | Tên KPI |
| target_prev | Target tuần trước |
| actual_prev | Actual tuần trước |
| root_cause | Nguyên nhân gốc rễ |
| action_item | Hành động thực hiện |
| priority | Ưu tiên (Cao / Trung bình / Thấp) |
| mo_ta | Mô tả chi tiết |
| target | Mục tiêu action |
| deadline | Deadline |
| owner | Người phụ trách |
| fm_support | Hỗ trợ cần từ FM |
| status | Trạng thái |
| saved_at | Thời gian lưu (ISO) |

### Tab Weekly Report (`T3_WeeklyReport` worksheet)
| Cột | Mô tả |
|-----|-------|
| bu | Tên BU |
| month | Tháng (YYYY-MM) |
| week | Tuần (1-4) |
| kpi | Tên KPI (cố định) |
| target | Target |
| actual | Actual |
| notes | Ghi chú |
| saved_at | Thời gian lưu (ISO) |

---

## Hệ thống chấm điểm (tổng 100 điểm)

### Action Plan (max 70 điểm)
- **Độ đầy đủ dữ liệu (max 20):** % dòng có KPI + target_prev + actual_prev + action_item
- **Chất lượng Root Cause (max 20):** % dòng có root_cause > 20 ký tự
- **Chất lượng Action (max 15):** % dòng có action_item VÀ (target HOẶC deadline)
- **Tiêu chí SMART (max 10):** % dòng có đủ target + deadline + owner
- **Đa dạng Function (max 5):** 3 function = 5đ, 2 function = 3đ, 1 function = 1đ

### Weekly Report (max 30 điểm)
- **Tỷ lệ điền KPI (max 15):** % dòng có cả target và actual
- **Ghi chú khi Gap lớn (max 15):** % dòng gap > 5% có ghi chú > 10 ký tự

### Xếp loại
- **≥ 80:** Tốt (xanh lá)
- **≥ 60:** Đầy đủ (xanh dương)
- **≥ 40:** Hời hợt (cam)
- **< 40:** Cần cải thiện (đỏ)

---

## Cấu trúc file dự án

```
cm-platform/
├── index.html              — Giao diện chính (Landing, CM, Dashboard)
├── style.css               — Stylesheet (MindX branding, Exo font)
├── app.js                  — Logic JS (state, tables, scoring, API calls)
├── google-apps-script.js   — Code Google Apps Script (copy vào GAS editor)
└── SETUP.md                — Tài liệu này
```

---

## Lỗi thường gặp

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| "Chưa kết nối Google Sheets" | APPS_SCRIPT_URL chưa thay | Xem Bước 5 |
| HTTP 401 / 403 | Apps Script chưa được phép truy cập | Re-deploy, chọn "Anyone" |
| Sheet not found | Worksheet chưa tồn tại | Kiểm tra tên sheet trong Spreadsheet |
| CORS error | Apps Script URL sai định dạng | Copy lại URL từ GAS Deploy |
