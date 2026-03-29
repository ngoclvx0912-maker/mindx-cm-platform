# Auto Period Logic - CM Report Platform

## Tuần trong tháng
- W1: ngày 1-7
- W2: ngày 8-14  
- W3: ngày 15-21
- W4: ngày 22-31

## CM nộp báo cáo: Chủ nhật + Thứ 2

### Quy tắc xác định kỳ báo cáo

**WR (Weekly Report)**: Báo cáo kết quả tuần VỪA QUA
**AP (Action Plan)**: Kế hoạch tuần KẾ TIẾP

Dựa trên ngày hiện tại → xác định tuần hiện tại → suy ra:
- WR period = tuần hiện tại (hoặc tuần vừa kết thúc)
- AP period = tuần kế tiếp

### Ví dụ cụ thể

| Hôm nay | Tuần hiện tại | WR cho | AP cho |
|---------|--------------|--------|--------|
| CN 8/3 (ngày 8) | W2 tháng 3 | W1 tháng 3 | W2 tháng 3 |
| T2 9/3 (ngày 9) | W2 tháng 3 | W1 tháng 3 | W2 tháng 3 |
| CN 15/3 (ngày 15) | W3 tháng 3 | W2 tháng 3 | W3 tháng 3 |
| T2 16/3 (ngày 16) | W3 tháng 3 | W2 tháng 3 | W3 tháng 3 |
| CN 22/3 (ngày 22) | W4 tháng 3 | W3 tháng 3 | W4 tháng 3 |
| CN 1/4 (ngày 1) | W1 tháng 4 | **W4 tháng 3** | W1 tháng 4 |
| T2 2/4 (ngày 2) | W1 tháng 4 | **W4 tháng 3** | W1 tháng 4 |
| CN 29/3 (ngày 29) | W4 tháng 3 | W3 tháng 3 | W4 tháng 3 |

### Logic chi tiết

```
today = ngày hiện tại
dayOfMonth = today.getDate()
currentWeek = dayOfMonth <= 7 ? 1 : dayOfMonth <= 14 ? 2 : dayOfMonth <= 21 ? 3 : 4

// WR = tuần trước đó
if currentWeek == 1:
  wrMonth = tháng trước
  wrWeek = 4
else:
  wrMonth = tháng hiện tại
  wrWeek = currentWeek - 1

// AP = tuần hiện tại
apMonth = tháng hiện tại
apWeek = currentWeek
```

### Deadline khóa

- Mở: Chủ nhật + Thứ 2 + Thứ 3 + Thứ 4 (CN, T2, T3, T4)
- **Khóa sau thứ 4** = từ thứ 5 trở đi (T5, T6, T7) → readonly
- dayOfWeek: 0=CN, 1=T2, 2=T3, 3=T4, 4=T5, 5=T6, 6=T7
- isLocked = dayOfWeek >= 4 (T5, T6, T7)
- isOpen = dayOfWeek <= 3 (CN, T2, T3, T4) — dayOfWeek 0,1,2,3

### UI Changes (CM View)
- Bỏ dropdown tháng/tuần → hiển thị text tĩnh
- Hiển thị rõ: "WR: Tháng 3 - Tuần 2" | "AP: Tháng 3 - Tuần 3"
- Nếu bị khóa: badge đỏ "Đã khóa — Hết hạn chỉnh sửa (sau thứ 4)"
- Nếu còn mở: badge xanh "Đang mở — Hạn chỉnh sửa: Thứ 4"
