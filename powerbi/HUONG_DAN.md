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
| 1 | Chi phí thay thế ước tính (mặc định 50%) | **$6.807.246** (thẻ hiện ≈ $6,81 triệu) |
| 3 | Ma trận Làm thêm giờ × Level 1 | **52,6%** |
| 6 | NV đang làm có hồ sơ rủi ro cao / TB | **19 / 557** |
| 6 | Ca nghỉ có thể giảm (ước tính) | **83** (35,1%) |
| 6 | Tỷ lệ nghỉ nếu giới hạn OT (ước tính) | **10,5%** |
| 6 | Chi phí có thể tiết kiệm (ước tính, 50%) | **$2.230.374** (thẻ hiện ≈ $2,23 triệu) |

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
Bảng độc lập (bản Final): _Measures (30 measure) · Driver_Impact · Replacement Cost (what-if)
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

## 10. Bản Final - thay bảng bằng biểu đồ (từ Final_2)
Giữ nguyên theme, bố cục và chỉnh sửa của nhóm trong Final_2 (6 trang; trang Chất lượng dữ liệu và trang drill-through đã bỏ). Chỉ thay các bảng/ma trận, mỗi biểu đồ mới nằm đúng trong khung cũ:

| Trang | Bảng cũ | Biểu đồ mới |
|---|---|---|
| 2 | Ma trận so sánh người nghỉ vs ở lại | Thanh ngang **% Level 1 · % làm thêm giờ · % độc thân** cho 2 nhóm (60% / 54% / 51% vs 32% / 23% / 28%). Tuổi, thâm niên, thu nhập hiện khi rê chuột |
| 3 | Heatmap Làm thêm giờ × Cấp bậc | Cột nhóm theo cấp bậc: **có OT (đỏ) vs không OT (xám)** - Level 1 + OT = 52,6% |
| 3 | Heatmap Làm thêm giờ × Thâm niên | Cột nhóm theo thâm niên: năm đầu + OT = 55,1% |
| 4 | Heatmap Hôn nhân × Cổ phiếu | Cột nhóm theo hôn nhân, 4 mức cổ phiếu (mức 0 = đỏ, mức 1-3 = xanh đậm dần). Độc thân chỉ có cột "mức 0" → thấy ngay bẫy gây nhiễu |
| 6 | Kế hoạch hành động ưu tiên | 5 thẻ ưu tiên (1-4 viền xanh chàm, mục "KHÔNG ưu tiên" màu xám) |
| 6 | Danh sách NV đang làm rủi ro cao | Thanh **số NV rủi ro cao theo vị trí** (giữ nguyên bộ lọc: đang làm + rủi ro 4-5 yếu tố) |

> Thư mục `powerbi/` giờ là bản Final đã chỉnh tay của nhóm. `scripts/02_build_pbip.py` vẫn tạo được bản gốc (chưa có các chỉnh sửa tay này).

## 11. Bản sửa sau rà soát (từ Final_Two)
Giữ nguyên theme và bố cục của nhóm. Mọi con số dưới đây đã được tính lại từ dữ liệu.

**Sửa kết luận sai**
- **Thăng chức (trang 5, 6):** nhóm "đã thăng chức trong 5 năm" chứa 215 NV mới (≤1 năm) vốn nghỉ nhiều, nên so chung thì trông như chậm thăng chức không có tác động (14,2% vs 16,5%). Biểu đồ giờ **chỉ xét NV làm ≥5 năm** (bộ lọc YearsAtCompany ≥ 5): chưa thăng chức 5 năm nghỉ **14,2% vs 9,4%** (p ≈ 0,03). Đã sửa tiêu đề trang 5, ô SO WHAT, thẻ hành động số 3 và 5, hệ số trong biểu đồ Đòn bẩy (0,86 → 1,52).
- **Chi phí tiết kiệm (trang 6):** mỗi ca tránh được giờ được tính theo lương năm của chính nhóm làm thêm giờ đã nghỉ ở từng cấp bậc, thay vì lương trung bình của mọi người nghỉ: **$2.230.374** (trước: $2.392.439). Thêm đơn vị **$** cho các measure chi phí.

**Sửa chữ / số**
- Trang 1: "quản lý & giám đốc chỉ 2,5–6,9%" (Manufacturing Director 6,9%, trước ghi 4,9%).
- Trang 2: "42% từ 30 tuổi trở xuống" (dưới 30 tuổi là 38,4%).
- Trang 4: "+2,1 điểm %" và ghi rõ định nghĩa "< 85% trung vị cùng cấp".
- Trang 5: "gần gấp đôi" (1,85 lần); "cao gấp 1,4–2,2 lần"; cân bằng CV-CS 31,3%; đào tạo "n = 54"; gộp "3–4 thang Low" (28 NV, 42,9%) vì nhóm 4 thang chỉ có 1 người.
- Trang 3 và 6: tách rõ **82** (số người nghỉ trong nhóm Level 1 làm thêm giờ) và **~83** (số ca ước tính tránh được); cách tính ghi ở phụ đề trang 6.
- Trang 6: tiêu đề biểu đồ rủi ro cộng dồn ghi rõ 5 yếu tố (OT, Level 1, độc thân, công tác thường xuyên, ≥1 thang hài lòng Thấp).

**Sửa hiển thị**
- Donut trang 1, thu nhập theo cấp bậc trang 4, các thẻ KPI: bỏ đơn vị "nghìn" tự động → hiện đủ số (237 / 1.233; 2.437 vs 2.719). Thẻ chi phí hiện theo triệu với 2 chữ số thập phân.
- Trang 2: biểu đồ giới tính, học vấn và "nghỉ vs ở lại" đổi sang biểu đồ cột → đủ nhãn Nam/Nữ, đủ 5 mức học vấn, nhãn không còn chồng nhau.
- Trang 6: biểu đồ NV rủi ro cao theo vị trí đổi sang biểu đồ cột → hiện đủ 6 vị trí (tổng 19, trước bị thanh cuộn che 3 người); 5 thẻ hành động cao thêm cho đủ 2 dòng chữ.

## 12. Bản sửa sau rà soát lần 2 - kết luận không vượt quá dữ liệu
Số liệu không đổi; bản này sửa cách **diễn đạt** để không khẳng định quan hệ nhân quả khi dữ liệu chỉ cho thấy tương quan.

**Ước tính, không phải chắc chắn**
- Trang 6: "Khoảng 1/3 số ca nghỉ có thể tránh được" → "**có thể giảm tới ~1/3 (ước tính)**". Phụ đề ghi rõ giả định: làm thêm giờ là nguyên nhân. Kiểm tra độ vững: tách theo cấp bậc, tuổi, thâm niên, phòng ban hay hôn nhân đều ra 81-84 ca. Các thẻ KPI ghi "(ước tính)".
- Trang 6: "4-5 yếu tố → 70,3% nghỉ" → "nhóm có 4-5 yếu tố **từng nghỉ** 70,3%; 19 NV đang làm **có hồ sơ giống nhóm này**" (70% là tỷ lệ đã xảy ra trong dữ liệu, không phải xác suất nghỉ).
- Trang 4: bỏ "chưa có cơ chế giữ người giỏi" → "**điểm hiệu suất (chỉ 2 mức) không dự báo nghỉ việc**" (mức 3: 16,1%, n=1.244; mức 4: 16,4%, n=226).

**Lập luận chặt hơn**
- Trang 4: biểu đồ "Thu nhập < 3K" giờ **chỉ xét Level 1** (96,5% người dưới 3K là Level 1). Trong Level 1: <3K nghỉ **29,4% (n=381)** vs 3-5K **19,1% (n=162)**, p ≈ 0,01 → đề xuất nâng sàn lương vẫn có cơ sở.
- Cổ phiếu: không NV độc thân nào có cổ phiếu, và ở nhóm không độc thân mức 3 lại nghỉ 17,6% (n=85) → đổi thành "**thử nghiệm (pilot)** cổ phiếu cho NV độc thân/Level 1, hiệu quả chưa được kiểm chứng".
- Biểu đồ Đòn bẩy: ghi chú "các yếu tố chồng lấn nên không cộng dồn" (cấp bậc ~ lương r = 0,95; cấp bậc ~ thâm niên r = 0,53).
- Thang hài lòng: ghi rõ "Số thang Low" và cờ rủi ro dùng **4 thang hài lòng**; biểu đồ đầu trang 5 dùng 5 thang (4 thang + mức gắn kết).

**Chú thích & dọn mô hình**
- Ô GHI CHÚ trên mọi trang: "chữ nhận định là số toàn công ty, không đổi theo bộ lọc".
- Biểu đồ đổi quản lý: cột "NV < 2 năm" chỉ để đủ nhóm; dữ liệu không cho biết ai rời đi trước.
- Xoá 4 measure của "7. Chất lượng dữ liệu" và 3 bảng không dùng (Cleaning_Log, Quality_Flags, Action_Plan). Nhật ký làm sạch vẫn có trong `Attrition_Clean.xlsx`.

> `scripts/02_build_pbip.py` tạo **bản gốc** (có trang 7 và các bảng trên); thư mục `powerbi/` là **bản Final** đã chỉnh tay và sửa theo 2 vòng rà soát.

## 13. Theme "Navy & Teal" (bản hiện tại)
Đổi toàn bộ 6 trang theo mẫu thiết kế mới; nội dung, số liệu và vị trí biểu đồ giữ nguyên.

| Thành phần | Thiết kế |
|---|---|
| Dải tiêu đề | Navy `#1E3A5F` bo góc, icon teal theo chủ đề từng trang (nhân viên, chân dung, công việc, lương, hài lòng, khiên bảo vệ), ô **GHI CHÚ** màu nằm bên phải dải |
| Màu biểu đồ tỷ lệ nghỉ | Tự tính theo bộ lọc: **đỏ `#E5484D`** = rất cao (≥ 1,8 lần mức chung) · **teal `#2A9D8F`** = cao hơn mức chung · **xanh nhạt `#9CC9E8`** = thấp hơn |
| Màu khác | Xanh thép `#3C8DBC` (người ở lại), navy `#1F4E79` (nhấn mạnh), nền trang `#F3F6FA` |
| Thẻ KPI | Icon teal bên trái, số in đậm (đỏ cho chỉ số nghỉ việc) |
| Donut trang 1 | Nhãn "237 (16,12%)" và số **16,1%** ở giữa |
| Ô SO WHAT? | Nền xanh bạc hà, icon bóng đèn, nhãn "SO WHAT?" màu teal và đường kẻ dọc; cao hơn (80px) cho đủ chữ |
| Ghi chú bộ lọc | "ⓘ Chữ nhận định là số toàn công ty, không đổi theo bộ lọc" ở hàng slicer |

Icon là ảnh PNG trong `StaticResources/RegisteredResources/` (ic_*.png, hd_*.png). Muốn đổi icon: chọn icon trong Power BI → Format → Image → thay ảnh.
