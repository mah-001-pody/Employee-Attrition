"""
VJP205 - Employee Attrition
Bước 1: Tiền xử lý & làm sạch dữ liệu (Data Pre-processing & Cleaning)

Chạy:   python scripts/01_data_cleaning.py
Input:  02_Employee Attrition.xlsx (sheet "Data", thư mục gốc repo)
Output: data/processed/Attrition_Clean.csv   -> nạp vào Power BI / mô hình
        data/processed/Attrition_Clean.xlsx  -> Data_Clean, Data_Dictionary, Cleaning_Log,
                                                Data_Quality_Flags, Outlier_Check,
                                                Quality_Before_After
"""
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "02_Employee Attrition.xlsx"
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 0. Đọc dữ liệu & chuẩn hoá tên cột (bỏ khoảng trắng -> PascalCase)
# ---------------------------------------------------------------------------
raw = pd.read_excel(RAW, sheet_name="Data")
df = raw.copy()
df.columns = [c.replace(" ", "") for c in df.columns]
df = df.rename(columns={"Over18": "Over18", "EmployeeNumber": "EmployeeID"})

SAT_COLS = ["EnvironmentSatisfaction", "JobSatisfaction", "RelationshipSatisfaction", "WorkLifeBalance"]
ORDINAL_1_4 = SAT_COLS + ["JobInvolvement"]
YEAR_COLS = ["TotalWorkingYears", "YearsAtCompany", "YearsInCurrentRole",
             "YearsSinceLastPromotion", "YearsWithCurrentManager"]

log = []


def note(step, action, n, reason, d):
    log.append({"Step": step, "Action": action, "Rows_Affected": int(n), "Reason": reason, "Rows_After": len(d)})


def profile(d, label):
    num = d.select_dtypes("number")
    p = num.describe().T[["count", "mean", "std", "min", "50%", "max"]]
    p.columns = [f"{label}_{c}" for c in p.columns]
    return p


before = profile(df, "Before")
note("R0", "Kiểm tra", len(df), "Dữ liệu gốc: 1.470 nhân viên × 34 cột", df)

# ---------------------------------------------------------------------------
# 1. Kiểm tra cấu trúc
# ---------------------------------------------------------------------------
note("R1", "Kiểm tra", df.isna().sum().sum(), "Giá trị thiếu (missing): không có", df)
note("R1", "Kiểm tra", df.duplicated().sum(), "Bản ghi trùng lặp toàn bộ cột: không có", df)
note("R1", "Kiểm tra", df["EmployeeID"].duplicated().sum(), "Mã nhân viên trùng: không có (ID duy nhất)", df)
bad_ord = int((~df[ORDINAL_1_4].isin([1, 2, 3, 4])).sum().sum())
note("R1", "Kiểm tra", bad_ord, "Thang đo hài lòng/gắn kết nằm ngoài 1-4", df)
note("R1", "Kiểm tra", int((~df["Attrition"].isin(["Yes", "No"])).sum()), "Nhãn Attrition ngoài Yes/No", df)

# ---------------------------------------------------------------------------
# 2. Kiểm tra logic thâm niên (quan hệ giữa các cột năm)
# ---------------------------------------------------------------------------
rules = {
    "YearsAtCompany > TotalWorkingYears": df["YearsAtCompany"] > df["TotalWorkingYears"],
    "YearsInCurrentRole > YearsAtCompany": df["YearsInCurrentRole"] > df["YearsAtCompany"],
    "YearsWithCurrentManager > YearsAtCompany": df["YearsWithCurrentManager"] > df["YearsAtCompany"],
    "YearsSinceLastPromotion > YearsAtCompany": df["YearsSinceLastPromotion"] > df["YearsAtCompany"],
    "TotalWorkingYears > Age - 18": df["TotalWorkingYears"] > df["Age"] - 18,
}
for k, m in rules.items():
    note("R2", "Kiểm tra", m.sum(), f"Mâu thuẫn thâm niên: {k}", df)

# NumCompaniesWorked = 0 nhưng đã có kinh nghiệm trước khi vào công ty -> mâu thuẫn ngữ nghĩa
m_nc = (df["NumCompaniesWorked"] == 0) & (df["TotalWorkingYears"] > df["YearsAtCompany"])
df["Flag_NumCompanies_Inconsistent"] = m_nc.astype(int)
note("R3", "Gắn cờ (giữ lại)", m_nc.sum(),
     "NumCompaniesWorked = 0 nhưng TotalWorkingYears > YearsAtCompany (đã làm ở nơi khác) → "
     "ngữ nghĩa cột không rõ; KHÔNG xoá (13,4% dữ liệu), gắn cờ và dùng biến Prior_Experience_Years thay thế", df)

# ---------------------------------------------------------------------------
# 3. Cột hằng số & cột nghi là nhiễu
# ---------------------------------------------------------------------------
const = [c for c in df.columns if df[c].nunique() == 1]
df = df.drop(columns=const)
note("R4", "Loại cột", 0, f"Bỏ cột hằng số không mang thông tin: {', '.join(const)}", df)

rate_corr = df[["DailyRate", "HourlyRate", "MonthlyRate"]].corrwith(df["MonthlyIncome"]).abs().max()
note("R5", "Gắn cờ (giữ lại)", 0,
     f"DailyRate/HourlyRate/MonthlyRate gần như không tương quan với MonthlyIncome (|r| ≤ {rate_corr:.2f}) "
     "→ không phản ánh thu nhập thực; giữ trong dữ liệu nhưng KHÔNG dùng để phân tích lương", df)

note("R6", "Ghi chú", 0,
     "PerformanceRating chỉ có 2 mức (3, 4) và trùng khớp hoàn toàn với PercentSalaryHike (≤19% ↔ 3, ≥20% ↔ 4) "
     "→ biến phương sai thấp, tránh đưa cả hai vào cùng một mô hình", df)
note("R6", "Ghi chú", 0,
     "MonthlyIncome tương quan 0,95 với JobLevel → đa cộng tuyến; khi mô hình hoá chỉ chọn một "
     "(hoặc dùng Income_vs_Level_Ratio)", df)

# ---------------------------------------------------------------------------
# 4. Ngoại lai - kiểm tra nhưng GIỮ LẠI
# ---------------------------------------------------------------------------
out_rows = []
for c in ["Age", "MonthlyIncome", "DistanceFromHome", "NumCompaniesWorked", "PercentSalaryHike",
          "TrainingTimesLastYear"] + YEAR_COLS:
    q1, q3 = df[c].quantile([0.25, 0.75])
    iqr = q3 - q1
    n = int(((df[c] < q1 - 1.5 * iqr) | (df[c] > q3 + 1.5 * iqr)).sum())
    out_rows.append({"Variable": c, "Q1": q1, "Q3": q3, "Lower": q1 - 1.5 * iqr, "Upper": q3 + 1.5 * iqr,
                     "IQR_Outliers": n, "Pct": round(n / len(df) * 100, 2), "Skewness": round(df[c].skew(), 2)})
outliers = pd.DataFrame(out_rows)
out_cols = ["MonthlyIncome"] + YEAR_COLS + ["NumCompaniesWorked", "TrainingTimesLastYear"]
q1s, q3s = df[out_cols].quantile(0.25), df[out_cols].quantile(0.75)
any_out = ((df[out_cols] < q1s - 1.5 * (q3s - q1s)) | (df[out_cols] > q3s + 1.5 * (q3s - q1s))).any(axis=1)
note("R7", "Giữ lại", int(any_out.sum()),
     "Số nhân viên có ít nhất 1 biến vượt ngưỡng IQR. Ngoại lai ở thu nhập/thâm niên là giá trị THẬT (quản lý cấp cao, nhân viên lâu năm) → không xoá; "
     "dùng log-transform hoặc nhóm (binning) khi mô hình hoá", df)

# ---------------------------------------------------------------------------
# 5. Mã hoá & gắn nhãn
# ---------------------------------------------------------------------------
df["Attrition_Flag"] = (df["Attrition"] == "Yes").astype(int)
df["OverTime_Flag"] = (df["OverTime"] == "Yes").astype(int)
df["Gender_Male"] = (df["Gender"] == "Male").astype(int)
TRAVEL = {"Non-Travel": 0, "Travel_Rarely": 1, "Travel_Frequently": 2}
df["BusinessTravel_Ord"] = df["BusinessTravel"].map(TRAVEL)
df["BusinessTravel_Label"] = df["BusinessTravel"].map(
    {"Non-Travel": "Không đi công tác", "Travel_Rarely": "Thỉnh thoảng", "Travel_Frequently": "Thường xuyên"})

df["Education_Label"] = df["Education"].map(
    {1: "1-Below College", 2: "2-College", 3: "3-Bachelor", 4: "4-Master", 5: "5-Doctor"})
SAT = {1: "1-Low", 2: "2-Medium", 3: "3-High", 4: "4-Very High"}
for c in ["EnvironmentSatisfaction", "JobSatisfaction", "RelationshipSatisfaction", "JobInvolvement"]:
    df[f"{c}_Label"] = df[c].map(SAT)
df["WorkLifeBalance_Label"] = df["WorkLifeBalance"].map({1: "1-Bad", 2: "2-Good", 3: "3-Better", 4: "4-Best"})
df["PerformanceRating_Label"] = df["PerformanceRating"].map({3: "3-Excellent", 4: "4-Outstanding"})
note("R8", "Mã hoá", len(df),
     "Mã hoá nhị phân (Attrition, OverTime, Gender), thứ bậc (BusinessTravel) và gắn nhãn cho các thang đo 1-4/1-5 "
     "theo mô tả chuẩn của bộ dữ liệu IBM HR", df)

# ---------------------------------------------------------------------------
# 6. Feature engineering
# ---------------------------------------------------------------------------
df["Age_Group"] = pd.cut(df["Age"], [17, 25, 35, 45, 55, 60],
                         labels=["18-25", "26-35", "36-45", "46-55", "56-60"])
df["Income_Band"] = pd.cut(df["MonthlyIncome"], [0, 3000, 5000, 10000, 15000, np.inf],
                           labels=["<3K", "3-5K", "5-10K", "10-15K", "15K+"], right=False)
df["Tenure_Group"] = pd.cut(df["YearsAtCompany"], [-1, 1, 5, 10, 20, 100],
                            labels=["0-1 năm", "2-5 năm", "6-10 năm", "11-20 năm", "20+ năm"])
df["Distance_Group"] = pd.cut(df["DistanceFromHome"], [0, 5, 15, 100],
                              labels=["Gần (≤5)", "Trung bình (6-15)", "Xa (>15)"])

df["Prior_Experience_Years"] = df["TotalWorkingYears"] - df["YearsAtCompany"]
df["Avg_Years_Per_Company"] = (df["TotalWorkingYears"] / df["NumCompaniesWorked"].clip(lower=1)).round(2)
df["Is_New_Hire"] = (df["YearsAtCompany"] <= 1).astype(int)
df["Role_Tenure_Ratio"] = np.where(df["YearsAtCompany"] > 0,
                                   df["YearsInCurrentRole"] / df["YearsAtCompany"], 0).round(3)
df["Promotion_Stagnation"] = (df["YearsSinceLastPromotion"] >= 5).astype(int)

df["Satisfaction_Index"] = df[SAT_COLS].mean(axis=1).round(2)
df["Low_Satisfaction_Count"] = (df[SAT_COLS] == 1).sum(axis=1)
df["Satisfaction_Group"] = pd.cut(df["Satisfaction_Index"], [0, 2.25, 2.75, 3.25, 4.01],
                                  labels=["Thấp", "Trung bình", "Khá", "Cao"])

# Lương so với mặt bằng cùng cấp (pay fairness): <1 = thấp hơn trung vị cùng Job Level
lvl_median = df.groupby("JobLevel")["MonthlyIncome"].transform("median")
df["Income_vs_Level_Ratio"] = (df["MonthlyIncome"] / lvl_median).round(3)
df["Underpaid_vs_Level"] = (df["Income_vs_Level_Ratio"] < 0.85).astype(int)

# Cờ rủi ro tổng hợp (dùng cho dashboard, KHÔNG phải mô hình)
df["Risk_Flags"] = (df["OverTime_Flag"] + (df["JobLevel"] == 1).astype(int)
                    + (df["MaritalStatus"] == "Single").astype(int)
                    + (df["BusinessTravel"] == "Travel_Frequently").astype(int)
                    + (df["Low_Satisfaction_Count"] >= 1).astype(int))
df["Risk_Group"] = pd.cut(df["Risk_Flags"], [-1, 1, 3, 5], labels=["Thấp (0-1)", "Trung bình (2-3)", "Cao (4-5)"])

df["Log_MonthlyIncome"] = np.log1p(df["MonthlyIncome"]).round(4)
df["Log_YearsAtCompany"] = np.log1p(df["YearsAtCompany"]).round(4)
n_new = len(df.columns) - (len(raw.columns) - len(const)) - 1  # -1: Flag_NumCompanies_Inconsistent da tinh o R3
note("R9", "Tạo biến", len(df),
     f"Feature engineering: {n_new} biến mới (nhãn, nhóm tuổi/thu nhập/thâm niên, kinh nghiệm trước, "
     "chỉ số hài lòng, lương so với cùng cấp, cờ rủi ro, log)", df)

# ---------------------------------------------------------------------------
# 7. Kiểm tra cuối
# ---------------------------------------------------------------------------
assert df.isna().sum().sum() == 0
assert df["EmployeeID"].is_unique
assert len(df) == len(raw)
note("R10", "Kiểm tra", len(df), "Dữ liệu sạch: 0 missing, 0 trùng lặp, giữ 100% nhân viên", df)
after = profile(df, "After")
quality = before.join(after, how="outer")

flags = pd.DataFrame([
    ("NumCompaniesWorked = 0 nhưng có kinh nghiệm trước", int(m_nc.sum()), "Gắn cờ, dùng Prior_Experience_Years"),
    ("DailyRate/HourlyRate/MonthlyRate không khớp MonthlyIncome", len(df), "Không dùng để phân tích lương"),
    ("PerformanceRating chỉ 2 mức, trùng với PercentSalaryHike", len(df), "Không đưa cả hai vào mô hình"),
    ("MonthlyIncome ~ JobLevel (r = 0,95)", len(df), "Chọn 1 biến khi mô hình hoá"),
    ("Mất cân bằng nhãn: Attrition = Yes chỉ 16,1%", int(df["Attrition_Flag"].sum()),
     "Stratified split, class_weight, đánh giá bằng Recall/AUC thay vì Accuracy"),
], columns=["Issue", "Rows", "Treatment"])

DICT = [
    ("EmployeeID", "Mã nhân viên (duy nhất)", "ID", "Gốc"),
    ("Attrition / Attrition_Flag", "Nghỉ việc: Yes/No · 1/0 (BIẾN MỤC TIÊU)", "Nhị phân", "Gốc / Tạo mới"),
    ("Age, Gender, MaritalStatus", "Tuổi, giới tính, tình trạng hôn nhân", "Số / Phân loại", "Gốc"),
    ("Department, JobRole, JobLevel", "Phòng ban, vị trí, cấp bậc (1-5)", "Phân loại", "Gốc"),
    ("BusinessTravel(_Ord/_Label)", "Tần suất công tác (0-2) + nhãn tiếng Việt", "Thứ bậc", "Gốc / Tạo mới"),
    ("OverTime / OverTime_Flag", "Làm thêm giờ Yes/No · 1/0", "Nhị phân", "Gốc / Tạo mới"),
    ("MonthlyIncome, PercentSalaryHike, StockOptionLevel", "Thu nhập tháng, % tăng lương, mức cổ phiếu", "Số", "Gốc"),
    ("DailyRate, HourlyRate, MonthlyRate", "Đơn giá (KHÔNG khớp thu nhập - không dùng phân tích lương)", "Số", "Gốc"),
    ("Education(_Label), EducationField", "Trình độ 1-5 + nhãn, ngành học", "Thứ bậc", "Gốc / Tạo mới"),
    ("*Satisfaction(_Label), WorkLifeBalance(_Label), JobInvolvement(_Label)", "Thang 1-4 + nhãn", "Thứ bậc", "Gốc / Tạo mới"),
    ("PerformanceRating(_Label)", "Đánh giá hiệu suất (3-Excellent / 4-Outstanding)", "Thứ bậc", "Gốc / Tạo mới"),
    ("TotalWorkingYears, YearsAtCompany, YearsInCurrentRole, YearsSinceLastPromotion, YearsWithCurrentManager",
     "Các mốc thâm niên (năm)", "Số", "Gốc"),
    ("NumCompaniesWorked, TrainingTimesLastYear, DistanceFromHome", "Số công ty đã làm, số lần đào tạo, khoảng cách (dặm)", "Số", "Gốc"),
    ("Flag_NumCompanies_Inconsistent", "1 = NumCompaniesWorked = 0 nhưng có kinh nghiệm trước", "Nhị phân", "Tạo mới"),
    ("Gender_Male", "1 = Nam", "Nhị phân", "Tạo mới"),
    ("Age_Group", "18-25 / 26-35 / 36-45 / 46-55 / 56-60", "Phân loại", "Tạo mới"),
    ("Income_Band", "<3K / 3-5K / 5-10K / 10-15K / 15K+", "Phân loại", "Tạo mới"),
    ("Tenure_Group", "0-1 / 2-5 / 6-10 / 11-20 / 20+ năm tại công ty", "Phân loại", "Tạo mới"),
    ("Distance_Group", "Gần ≤5 / Trung bình 6-15 / Xa >15 dặm", "Phân loại", "Tạo mới"),
    ("Prior_Experience_Years", "TotalWorkingYears − YearsAtCompany", "Số", "Tạo mới"),
    ("Avg_Years_Per_Company", "TotalWorkingYears / max(NumCompaniesWorked, 1) (độ 'nhảy việc')", "Số", "Tạo mới"),
    ("Is_New_Hire", "1 = làm ≤ 1 năm", "Nhị phân", "Tạo mới"),
    ("Role_Tenure_Ratio", "YearsInCurrentRole / YearsAtCompany", "Tỷ lệ", "Tạo mới"),
    ("Promotion_Stagnation", "1 = ≥ 5 năm chưa thăng chức", "Nhị phân", "Tạo mới"),
    ("Satisfaction_Index", "TB 4 thang: Environment, Job, Relationship, WorkLifeBalance", "Số", "Tạo mới"),
    ("Low_Satisfaction_Count", "Số thang đo bằng 1 (0-4)", "Số", "Tạo mới"),
    ("Satisfaction_Group", "Thấp / Trung bình / Khá / Cao theo Satisfaction_Index", "Phân loại", "Tạo mới"),
    ("Income_vs_Level_Ratio", "Thu nhập / trung vị thu nhập cùng JobLevel", "Tỷ lệ", "Tạo mới"),
    ("Underpaid_vs_Level", "1 = thấp hơn 85% trung vị cùng cấp", "Nhị phân", "Tạo mới"),
    ("Risk_Flags / Risk_Group", "Số yếu tố rủi ro: OT, Level 1, độc thân, công tác thường xuyên, có thang hài lòng = 1",
     "Số / Phân loại", "Tạo mới"),
    ("Log_MonthlyIncome, Log_YearsAtCompany", "log(1+x) cho mô hình", "Số", "Tạo mới"),
]
dictionary = pd.DataFrame(DICT, columns=["Column", "Description_VI", "Type", "Source"])

df.to_csv(OUT / "Attrition_Clean.csv", index=False, encoding="utf-8-sig")
with pd.ExcelWriter(OUT / "Attrition_Clean.xlsx", engine="openpyxl") as xw:
    df.to_excel(xw, sheet_name="Data_Clean", index=False)
    dictionary.to_excel(xw, sheet_name="Data_Dictionary", index=False)
    pd.DataFrame(log).to_excel(xw, sheet_name="Cleaning_Log", index=False)
    flags.to_excel(xw, sheet_name="Data_Quality_Flags", index=False)
    outliers.to_excel(xw, sheet_name="Outlier_Check", index=False)
    quality.to_excel(xw, sheet_name="Quality_Before_After")

print(pd.DataFrame(log)[["Step", "Action", "Rows_Affected", "Rows_After"]].to_string(index=False))
print(f"\nClean: {df.shape[0]} dòng × {df.shape[1]} cột -> {OUT}")
