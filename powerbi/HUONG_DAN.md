# Hướng dẫn mở dashboard Employee Attrition (`.pbip`) và lưu thành `.pbix`

## 0. Yêu cầu
- **Windows** có cài **Power BI Desktop bản mới**, nên cài từ Microsoft Store.
- Nếu không mở được file `.pbip`: vào **File → Options → Preview features**, bật **Power BI Project (.pbip)** và **PBIR**, rồi khởi động lại Power BI.

## 1. Giải nén (bắt buộc)
1. Chuột phải `Attrition_Dashboard.zip` → **Extract All…** → **Extract**. Giải nén ở đâu cũng được.
2. **Không** mở file `.pbip` trực tiếp từ bên trong file zip.

```
...\Attrition\
├── Attrition_Dashboard.pbip           ← BẤM ĐÚP FILE NÀY
├── Attrition_Dashboard.SemanticModel\ ← mô hình (dữ liệu đã NHÚNG SẴN)
├── Attrition_Dashboard.Report\        ← 7 trang + 1 trang drill-through
└── Attrition_Clean.xlsx               ← dữ liệu sạch để tham khảo / nộp kèm
```

## 2. Mở và nạp dữ liệu
1. Bấm đúp **`Attrition_Dashboard.pbip`**.
2. Lần đầu mở, biểu đồ sẽ **trống**. Đây là bình thường.
3. Bấm **Home → Refresh** và chờ 10–30 giây.

## 3. Lưu thành file `.pbix`
**File → Save as**, chọn **Power BI file (\*.pbix)**, đặt tên `MaLopBA_SoNhom.pbix`.

## 4. Kiểm tra nhanh: các con số phải khớp

| Trang | Visual | Giá trị đúng |
|---|---|---|
| 1 | Tổng nhân viên · Số người nghỉ · Tỷ lệ nghỉ | **1.470 · 237 · 16,1%** |
| 1 | Thu nhập trung vị: người nghỉ / người ở lại | **3.202 / 5.204** |
| 1 | Chi phí thay thế ước tính (mặc định 50%) | **6.807.246** |
| 3 | Ma trận Làm thêm giờ × Level 1 | **52,6%** |
| 6 | NV hiện tại rủi ro cao / trung bình | **19 / 557** |
| 6 | Ca nghỉ tránh được (giới hạn OT) | **83** (35,1%) |
| 6 | Tỷ lệ nghỉ nếu giới hạn OT | **10,5%** |
| 6 | Chi phí tiết kiệm được (50%) | **2.392.439** |
| 7 | Số dòng gốc / bị xoá / sạch | **1.470 / 0 / 1.470** |

## 5. Cấu trúc dashboard

| Trang | Câu hỏi | Nội dung chính |
|---|---|---|
| 1. Tổng quan | Vấn đề lớn cỡ nào? | 6 KPI, tỷ lệ và số người nghỉ theo phòng ban, tỷ lệ theo 9 vị trí, donut |
| 2. Chân dung | Ai nghỉ việc? | Tuổi, thâm niên, cấp bậc, hôn nhân, giới tính, học vấn, bảng so sánh người nghỉ và người ở lại |
| 3. Điều kiện làm việc | Đòn bẩy số 1 là gì? | **Heatmap Làm thêm giờ × Cấp bậc và × Thâm niên**, công tác, cân bằng công việc–cuộc sống, khoảng cách |
| 4. Lương & đãi ngộ | Có phải chỉ vì tiền? | Thu nhập, cổ phiếu, **heatmap Hôn nhân × Cổ phiếu (bẫy gây nhiễu)**, so sánh với cùng cấp, hiệu suất |
| 5. Hài lòng & gắn kết | Họ cảm thấy thế nào? | 5 thang hài lòng trên 1 biểu đồ, gắn kết, đổi quản lý, thăng chức, đào tạo |
| 6. Rủi ro & hành động | Giữ ai, giữ bằng cách nào? | KPI mô phỏng, **thanh what-if chi phí**, rủi ro cộng dồn, biểu đồ đòn bẩy, kế hoạch hành động, **danh sách NV đang rủi ro cao** |
| 7. Chất lượng dữ liệu | Vì sao các con số đáng tin? | Vấn đề chất lượng, nhật ký làm sạch, 2 bẫy gây nhiễu |
| Chi tiết nhân viên *(ẩn)* | Danh sách theo nhóm rủi ro | Ở trang 6, chuột phải vào cột "Rủi ro cộng dồn" → **Drill through** |

- **Màu tự động:** cột **đỏ** là nhóm có tỷ lệ nghỉ cao hơn mức chung của vùng đang lọc, cột **xám** là thấp hơn. Màu tự cập nhật khi bạn đổi bộ lọc.
- **Slicer đồng bộ** trên trang 1–6: phòng ban, cấp bậc, nhóm tuổi, làm thêm giờ.

## 6. Mô hình dữ liệu
```
Dim_Department ─┐
Dim_AgeGroup ───┼──► Employees (1 dòng/nhân viên) ◄── Satisfaction_Long (unpivot 5 thang hài lòng)
Dim_IncomeBand ─┤
Dim_Tenure ─────┘
Bảng độc lập: _Measures (34 measure) · Driver_Impact · Action_Plan · Replacement Cost (what-if) · Cleaning_Log · Quality_Flags
```

## 7. Nếu gặp lỗi
| Hiện tượng | Cách xử lý |
|---|---|
| Không mở được `.pbip` | Cập nhật Power BI Desktop; bật Preview features (mục 0) |
| Báo thiếu thư mục | Bạn đang mở từ bên trong zip → giải nén trước |
| Biểu đồ chỉ có khung, không có dữ liệu | Chưa bấm **Refresh** |
| Cột không đổi màu đỏ/xám | Chọn biểu đồ → Format → Columns → Color → **fx** → Field value → `Màu tỷ lệ nghỉ` |
| Lỗi khác | Chụp màn hình hộp thoại lỗi gửi lại |

## 8. Tạo lại từ đầu
```bash
python scripts/01_data_cleaning.py
python scripts/02_build_pbip.py
```

## 9. Bản 2 - theme "Navy & Coral" và căn chỉnh lại bố cục
- **Theme mới:** dải tiêu đề xanh navy `#1B2A4A`, màu chính xanh chàm `#2E4A8C`, **đỏ san hô `#E5484D` = nghỉ việc / rủi ro**, xanh lục `#12A37A` = kết quả tốt, hổ phách `#C77700` = cảnh báo, nền trang xám xanh nhạt `#F1F4F9`. Ô "SO WHAT?" nền xanh nhạt, viền xanh chàm.
- **Phông:** Segoe UI (hiển thị đủ dấu tiếng Việt). Tiêu đề visual và số KPI dùng Segoe UI Semibold.
- **Hàng slicer** cao 54px (trước 44px), có lề trong và cách nội dung bên dưới 12px. Ô chú thích màu đổi thành "CÁCH ĐỌC MÀU" gọn trên 2 dòng.
- **Lưới thống nhất:** lề 16px, khoảng cách 12px; 2 hàng biểu đồ cao bằng nhau (248px) trên trang 2-5.
- **Trang 2:** bảng so sánh người nghỉ/ở lại (chỉ 2 dòng) đổi thành ma trận 7 chỉ số × 2 nhóm, lấp đầy khung.
- **Trang 3 & 4:** các heatmap (Làm thêm giờ × Cấp bậc, × Thâm niên, Hôn nhân × Cổ phiếu) tự giãn ô cho kín khung, chữ 12pt, không còn khoảng trắng phía dưới.
- **Trang 6:** sửa biểu đồ "Đòn bẩy" - loại visual cũ `stackedBarChart` không tồn tại trong Power BI, đổi thành `barChart` (biểu đồ thanh chồng).
- **Trang 7:** ô "2 cái bẫy gây nhiễu" chữ lớn hơn (11-14pt), bảng vấn đề chất lượng cao hơn.
