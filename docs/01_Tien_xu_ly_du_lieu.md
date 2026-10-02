# Bước 1: Tiền xử lý và làm sạch dữ liệu (tiêu chí 2, 15%)

> Tài liệu này dùng để viết phần **"Data Pre-processing & Cleaning"** trong báo cáo cuối kỳ VJP205.
> Code: `scripts/01_data_cleaning.py`. Dữ liệu sạch: `data/processed/Attrition_Clean.xlsx` và `.csv`.

## 1. Đánh giá tổng quan chất lượng dữ liệu gốc

| Hạng mục | Kết quả | Đánh giá |
|---|---|---|
| Kích thước | 1.470 nhân viên × 34 cột (26 cột số, 8 cột chữ) | Đủ cho kiểm định thống kê và hồi quy logistic |
| Giá trị thiếu | 0 | Không cần điền khuyết |
| Bản ghi trùng lặp | 0; `EmployeeNumber` là duy nhất | Có mã nhân viên, truy vết được từng người |
| Logic thâm niên | 0 mâu thuẫn (5 quy tắc) | Các cột năm nhất quán với nhau và với tuổi |
| **Cột hằng số** | `Over18` (luôn Y), `StandardHours` (luôn 80) | Không mang thông tin → **loại bỏ** |
| **Cột nghi là nhiễu** | `DailyRate`, `HourlyRate`, `MonthlyRate` có \|r\| ≤ 0,03 với `MonthlyIncome` | Tên cột gợi ý là lương nhưng không khớp thu nhập → **không dùng để phân tích lương** |
| **Mâu thuẫn ngữ nghĩa** | 197 người có `NumCompaniesWorked = 0` nhưng đã đi làm trước khi vào công ty | Ngữ nghĩa cột không rõ → **gắn cờ**, dùng biến thay thế |
| **Biến trùng thông tin** | `PerformanceRating` chỉ có 2 mức và trùng khớp với `PercentSalaryHike`; `MonthlyIncome` ~ `JobLevel` (r = 0,95) | Đa cộng tuyến → chỉ chọn 1 biến trong mỗi cặp khi mô hình hoá |
| **Mất cân bằng nhãn** | Chỉ **16,1%** nghỉ việc (237/1.470) | Thách thức lớn nhất cho mô hình (tiêu chí 5) |

> **Điểm khác với các bộ dữ liệu "bẩn" thông thường:** bộ này sạch về **hình thức** (không thiếu, không trùng) nhưng có vấn đề về **ngữ nghĩa**. Phần tiền xử lý vì vậy tập trung vào: biến nào **đáng tin**, biến nào **thừa**, biến nào **gây hiểu lầm**. Đây là điểm cần làm rõ trong báo cáo để ăn điểm tiêu chí 2.

## 2. Các bước làm sạch (Cleaning Log)

| Bước | Quy tắc | Số dòng | Xử lý và lý do |
|---|---|---|---|
| R1 | Missing, trùng lặp, ID trùng, thang 1–4, nhãn Yes/No | 0 lỗi | Cấu trúc nhất quán |
| R2 | 5 quy tắc logic thâm niên, ví dụ `YearsInCurrentRole ≤ YearsAtCompany ≤ TotalWorkingYears ≤ Age − 18` | 0 lỗi | Không có hồ sơ vô lý |
| R3 | `NumCompaniesWorked = 0` nhưng `TotalWorkingYears > YearsAtCompany` | **197 (13,4%)** | **Gắn cờ, giữ lại.** Xoá sẽ mất 13% dữ liệu vì một cột có ngữ nghĩa mơ hồ. Thay bằng `Prior_Experience_Years` là biến tính được chính xác |
| R4 | Bỏ `Over18`, `StandardHours` | −2 cột | Hằng số, phương sai bằng 0 |
| R5 | `DailyRate`, `HourlyRate`, `MonthlyRate` | Gắn cờ | Không khớp thu nhập thực → không dùng để kết luận về lương |
| R6 | `PerformanceRating` ↔ `PercentSalaryHike`; `MonthlyIncome` ↔ `JobLevel` | Ghi chú | Tránh đa cộng tuyến khi mô hình hoá |
| R7 | Ngoại lai IQR ở thu nhập và thâm niên | **Giữ lại** (485 người) | Quản lý cấp cao, nhân viên lâu năm là có thật. Xoá họ sẽ làm lệch kết luận về giữ chân nhân tài |
| R8 | Mã hoá và gắn nhãn | — | Nhị phân (Attrition, OverTime, Gender), thứ bậc (BusinessTravel), nhãn thang 1–4 theo chuẩn IBM |
| R9 | Feature engineering | +31 biến | Xem mục 3 |
| R10 | Kiểm tra cuối (assert) | ✓ | 0 missing, ID duy nhất, giữ đủ 1.470 dòng |

**Kết quả: giữ 100% nhân viên (1.470 dòng); từ 34 cột thành 63 cột.**

> **Vì sao không xoá dòng nào?** Bộ dữ liệu không có lỗi logic thật. Với nhân sự, mỗi người nghỉ việc (chỉ 237 người) đều quý cho mô hình. Xoá dòng một cách tuỳ tiện sẽ làm mất cân bằng nhãn nặng thêm.

## 3. Feature engineering (31 biến mới)

| Nhóm | Biến | Mục đích phân tích |
|---|---|---|
| Mã hoá | `Attrition_Flag`, `OverTime_Flag`, `Gender_Male`, `BusinessTravel_Ord` | Dùng cho tương quan và mô hình |
| Nhãn | `*_Label` cho Education, 4 thang hài lòng, JobInvolvement, Performance, BusinessTravel | Hiển thị trên dashboard |
| Nhóm hoá | `Age_Group`, `Income_Band`, `Tenure_Group`, `Distance_Group`, `Satisfaction_Group` | Slicer và so sánh tỷ lệ nghỉ việc |
| Sự nghiệp | `Prior_Experience_Years`, `Avg_Years_Per_Company` (độ "nhảy việc"), `Is_New_Hire`, `Role_Tenure_Ratio`, `Promotion_Stagnation` | Kiểm định giả thuyết về thăng tiến và gắn bó |
| Hài lòng | `Satisfaction_Index` (TB 4 thang), `Low_Satisfaction_Count` | Gộp 4 thang thành 1 chỉ số |
| **Công bằng lương** | `Income_vs_Level_Ratio`, `Underpaid_vs_Level` | Lương so với trung vị **cùng cấp**: bị trả thấp hơn đồng nghiệp có nghỉ nhiều hơn không? |
| Rủi ro | `Risk_Flags` (0–5), `Risk_Group` | Chỉ để **mô tả** trên dashboard. Không dùng làm biến mô hình, vì được tạo từ chính các quan sát |
| Mô hình | `Log_MonthlyIncome`, `Log_YearsAtCompany` | Giảm lệch phải |

## 4. Tín hiệu ban đầu (gợi ý giả thuyết)
Tỷ lệ nghỉ việc chung là **16,1%**.

| Yếu tố | Nhóm rủi ro cao | Nhóm rủi ro thấp | Gợi ý giả thuyết |
|---|---|---|---|
| Làm thêm giờ | Có: **30,5%** | Không: 10,4% | H1: OverTime làm tăng khả năng nghỉ việc |
| Thâm niên | 0–1 năm: **34,9%** | 11–20 năm: 6,7% | H2: Giai đoạn onboarding là thời điểm rủi ro nhất |
| Tuổi | 18–25: **35,8%** | 36–45: 9,2% | H3: Nhân viên trẻ dễ nghỉ hơn |
| Thu nhập | <3K: **28,6%** | 15K+: 3,8% | H4: Thu nhập thấp làm tăng khả năng nghỉ |
| Cổ phiếu | Level 0: **24,4%** | Level 1–2: 7,6–9,4% | H5: Quyền chọn cổ phiếu có tác dụng giữ chân |
| Hài lòng | Thấp: **25,1%** | Cao: 9,5% | H6: Mức hài lòng tỷ lệ nghịch với nghỉ việc |
| Công tác | Thường xuyên: **24,9%** | Không đi: 8,0% | H7: Đi công tác nhiều làm tăng khả năng nghỉ |
| Cộng dồn rủi ro | 4–5 yếu tố: **70,3%** | 0–1 yếu tố: 4,9% | Các yếu tố **cộng hưởng** với nhau |

**Phát hiện đi ngược trực giác (nên đưa vào báo cáo):**
- **"Lâu chưa thăng chức" không làm tăng tỷ lệ nghỉ:** 14,2% so với 16,5%.
- **"Bị trả thấp hơn đồng nghiệp cùng cấp" chỉ tăng nhẹ:** 17,8% so với 15,6%.

→ Nhân viên nghỉ việc chủ yếu vì **thu nhập tuyệt đối thấp, cấp bậc thấp, làm thêm giờ**, chứ không phải vì so sánh với đồng nghiệp hay vì chậm thăng chức. Cả hai giả thuyết này cần được **kiểm định thống kê** ở bước mô hình.

## 5. Cấu trúc file đầu ra `Attrition_Clean.xlsx`

| Sheet | Nội dung |
|---|---|
| `Data_Clean` | 1.470 nhân viên × 63 cột |
| `Data_Dictionary` | Mô tả từng biến |
| `Cleaning_Log` | Nhật ký 20 bước kiểm tra và xử lý |
| `Data_Quality_Flags` | 5 vấn đề chất lượng và cách xử lý |
| `Outlier_Check` | Ngưỡng IQR và độ lệch của từng biến |
| `Quality_Before_After` | Thống kê trước và sau |
