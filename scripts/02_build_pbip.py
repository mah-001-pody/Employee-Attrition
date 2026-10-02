"""
VJP205 - Employee Attrition
Bước 2: Sinh dự án Power BI (.pbip) - Semantic model (TMDL) + Report (PBIR)

Chạy:   python scripts/02_build_pbip.py
Output: powerbi/Attrition_Dashboard.pbip
        powerbi/Attrition_Dashboard.SemanticModel/  (mô hình, Power Query, DAX - dữ liệu NHÚNG SẴN)
        powerbi/Attrition_Dashboard.Report/         (7 trang + 1 trang drill-through)

Mở bằng Power BI Desktop -> Refresh -> Save As .pbix (xem powerbi/HUONG_DAN.md)
"""
import base64
import hashlib
import json
import re
import shutil
import uuid
import zlib
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "powerbi"
NAME = "Attrition_Dashboard"
SM = OUT / f"{NAME}.SemanticModel"
RP = OUT / f"{NAME}.Report"
THEME_SRC = ROOT / "scripts" / "pbip_assets"
CLEAN_XLSX = ROOT / "data" / "processed" / "Attrition_Clean.xlsx"
MEASURES = "_Measures"


def lt(seed: str) -> str:
    """lineageTag / logicalId cố định (chạy lại script không đổi ID)."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"attrition/{seed}"))


def hid(seed: str) -> str:
    return hashlib.md5(("attrition/" + seed).encode("utf-8")).hexdigest()[:20]


def q(name: str) -> str:
    """Quote tên trong TMDL khi cần."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        return name
    return "'" + name.replace("'", "''") + "'"


# =============================================================================
# 1. SEMANTIC MODEL SPEC
# =============================================================================
I, D, S = ("int64", "Int64.Type"), ("double", "type number"), ("string", "type text")


def c(name, t, fmt=None, summ="none", sort=None, hidden=False):
    return dict(name=name, dt=t[0], mt=t[1], fmt=fmt, summ=summ, sort=sort, hidden=hidden)


EMP_SOURCE = [
    c("EmployeeID", I, "0"), c("Attrition", S), c("Attrition_Flag", I, "0", "sum"),
    c("Age", I, "0", "average"), c("Gender", S), c("MaritalStatus", S),
    c("Department", S), c("JobRole", S), c("JobLevel", I, "0"),
    c("BusinessTravel", S), c("BusinessTravel_Label", S, sort="BusinessTravel_Ord"),
    c("BusinessTravel_Ord", I, "0", hidden=True),
    c("OverTime", S), c("OverTime_Flag", I, "0", "sum"),
    c("DistanceFromHome", I, "0", "average"), c("Distance_Group", S),
    c("Education", I, "0"), c("Education_Label", S), c("EducationField", S),
    c("MonthlyIncome", I, "#,##0", "average"), c("PercentSalaryHike", I, "0", "average"),
    c("StockOptionLevel", I, "0"),
    c("EnvironmentSatisfaction", I, "0"), c("JobSatisfaction", I, "0"),
    c("RelationshipSatisfaction", I, "0"), c("WorkLifeBalance", I, "0"), c("JobInvolvement", I, "0"),
    c("EnvironmentSatisfaction_Label", S), c("JobSatisfaction_Label", S),
    c("RelationshipSatisfaction_Label", S), c("WorkLifeBalance_Label", S), c("JobInvolvement_Label", S),
    c("PerformanceRating", I, "0"), c("PerformanceRating_Label", S),
    c("TotalWorkingYears", I, "0", "average"), c("YearsAtCompany", I, "0", "average"),
    c("YearsInCurrentRole", I, "0", "average"), c("YearsSinceLastPromotion", I, "0", "average"),
    c("YearsWithCurrentManager", I, "0", "average"), c("NumCompaniesWorked", I, "0", "average"),
    c("TrainingTimesLastYear", I, "0", "average"),
    c("Age_Group", S), c("Income_Band", S), c("Tenure_Group", S),
    c("Prior_Experience_Years", I, "0", "average"), c("Is_New_Hire", I, "0", "sum"),
    c("Promotion_Stagnation", I, "0", "sum"),
    c("Satisfaction_Index", D, "0.00", "average"), c("Low_Satisfaction_Count", I, "0"),
    c("Income_vs_Level_Ratio", D, "0.00", "average"), c("Underpaid_vs_Level", I, "0", "sum"),
    c("Risk_Flags", I, "0"), c("Risk_Group", S, sort="Risk_Sort"),
    c("Flag_NumCompanies_Inconsistent", I, "0", "sum"),
]
EMP_ADDED = [
    c("Attrition_Label", S, sort="Attrition_Sort"), c("Attrition_Sort", I, "0", hidden=True),
    c("OverTime_Label", S), c("JobLevel_Label", S, sort="JobLevel"),
    c("Marital_Label", S, sort="Marital_Sort"), c("Marital_Sort", I, "0", hidden=True),
    c("Gender_Label", S),
    c("StockOption_Label", S, sort="StockOptionLevel"),
    c("Mgr_Change_Label", S, sort="Mgr_Change_Sort"), c("Mgr_Change_Sort", I, "0", hidden=True),
    c("Training_Label", S, sort="TrainingTimesLastYear"),
    c("Promotion_Label", S), c("Underpaid_Label", S),
    c("LowSat_Label", S, sort="Low_Satisfaction_Count"),
    c("Risk_Sort", I, "0", hidden=True),
    c("Annual_Income", I, "#,##0", "sum"),
]


def m_types(cols):
    return ", ".join(f'{{"{x["name"]}", {x["mt"]}}}' for x in cols)


def m_embedded(sheet: str, columns: list) -> str:
    """Nhúng dữ liệu vào M (giống 'Enter data' của Power BI): JSON -> raw deflate -> base64.
    Không phụ thuộc đường dẫn file trên máy người mở."""
    df = pd.read_excel(CLEAN_XLSX, sheet_name=sheet)[columns]
    rows = [[None if pd.isna(v) else (v.item() if hasattr(v, "item") else v) for v in r]
            for r in df.itertuples(index=False)]
    comp = zlib.compressobj(9, zlib.DEFLATED, -15)
    raw = comp.compress(json.dumps(rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8")) + comp.flush()
    b64 = base64.b64encode(raw).decode("ascii")
    names = ", ".join(f'"{x}"' for x in columns)
    return (f'Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("{b64}", '
            f'BinaryEncoding.Base64), Compression.Deflate)), {{{names}}})')


EMPLOYEES_M = f"""let
    {m_embedded("Data_Clean", [x["name"] for x in EMP_SOURCE])},
    Typed = Table.TransformColumnTypes(Source, {{{m_types(EMP_SOURCE)}}}),
    #"Added Attrition_Label" = Table.AddColumn(Typed, "Attrition_Label", each if [Attrition_Flag] = 1 then "Nghỉ việc" else "Ở lại", type text),
    #"Added Attrition_Sort" = Table.AddColumn(#"Added Attrition_Label", "Attrition_Sort", each if [Attrition_Flag] = 1 then 1 else 2, Int64.Type),
    #"Added OverTime_Label" = Table.AddColumn(#"Added Attrition_Sort", "OverTime_Label", each if [OverTime_Flag] = 1 then "Có làm thêm giờ" else "Không làm thêm giờ", type text),
    #"Added JobLevel_Label" = Table.AddColumn(#"Added OverTime_Label", "JobLevel_Label", each "Level " & Text.From([JobLevel]), type text),
    MaritalMap = [Single = {{"Độc thân", 1}}, Married = {{"Đã kết hôn", 2}}, Divorced = {{"Ly hôn", 3}}],
    #"Added Marital_Label" = Table.AddColumn(#"Added JobLevel_Label", "Marital_Label", each Record.Field(MaritalMap, [MaritalStatus]){{0}}, type text),
    #"Added Marital_Sort" = Table.AddColumn(#"Added Marital_Label", "Marital_Sort", each Record.Field(MaritalMap, [MaritalStatus]){{1}}, Int64.Type),
    #"Added Gender_Label" = Table.AddColumn(#"Added Marital_Sort", "Gender_Label", each if [Gender] = "Male" then "Nam" else "Nữ", type text),
    #"Added StockOption_Label" = Table.AddColumn(#"Added Gender_Label", "StockOption_Label", each "Cổ phiếu mức " & Text.From([StockOptionLevel]), type text),
    #"Added Mgr_Change_Label" = Table.AddColumn(#"Added StockOption_Label", "Mgr_Change_Label", each if [YearsAtCompany] < 2 then "NV < 2 năm" else if [YearsWithCurrentManager] = 0 then "Vừa đổi quản lý" else "Quản lý ổn định", type text),
    #"Added Mgr_Change_Sort" = Table.AddColumn(#"Added Mgr_Change_Label", "Mgr_Change_Sort", each if [YearsAtCompany] < 2 then 3 else if [YearsWithCurrentManager] = 0 then 1 else 2, Int64.Type),
    #"Added Training_Label" = Table.AddColumn(#"Added Mgr_Change_Sort", "Training_Label", each Text.From([TrainingTimesLastYear]) & " lần", type text),
    #"Added Promotion_Label" = Table.AddColumn(#"Added Training_Label", "Promotion_Label", each if [Promotion_Stagnation] = 1 then "≥5 năm chưa thăng chức" else "Thăng chức trong 5 năm", type text),
    #"Added Underpaid_Label" = Table.AddColumn(#"Added Promotion_Label", "Underpaid_Label", each if [Underpaid_vs_Level] = 1 then "Thấp hơn 85% cùng cấp" else "Ngang/cao hơn cùng cấp", type text),
    #"Added LowSat_Label" = Table.AddColumn(#"Added Underpaid_Label", "LowSat_Label", each Text.From([Low_Satisfaction_Count]) & " thang Low", type text),
    #"Added Risk_Sort" = Table.AddColumn(#"Added LowSat_Label", "Risk_Sort", each if [Risk_Flags] <= 1 then 1 else if [Risk_Flags] <= 3 then 2 else 3, Int64.Type),
    Result = Table.AddColumn(#"Added Risk_Sort", "Annual_Income", each [MonthlyIncome] * 12, Int64.Type)
in
    Result"""

SAT_M = """let
    Source = Table.SelectColumns(Employees, {"EmployeeID", "EnvironmentSatisfaction", "JobSatisfaction", "RelationshipSatisfaction", "WorkLifeBalance", "JobInvolvement"}),
    Unpivoted = Table.UnpivotOtherColumns(Source, {"EmployeeID"}, "Dimension_Code", "Score"),
    Map = [EnvironmentSatisfaction = {"Môi trường", 1}, JobSatisfaction = {"Công việc", 2}, RelationshipSatisfaction = {"Quan hệ", 3}, WorkLifeBalance = {"Cân bằng CV-CS", 4}, JobInvolvement = {"Mức gắn kết", 5}],
    #"Added Dimension" = Table.AddColumn(Unpivoted, "Dimension", each Record.Field(Map, [Dimension_Code]){0}, type text),
    #"Added Dimension_Sort" = Table.AddColumn(#"Added Dimension", "Dimension_Sort", each Record.Field(Map, [Dimension_Code]){1}, Int64.Type),
    #"Added Score_Label" = Table.AddColumn(#"Added Dimension_Sort", "Score_Label", each Text.From([Score]) & " - " & {"Thấp", "Trung bình", "Cao", "Rất cao"}{[Score] - 1}, type text),
    Typed = Table.TransformColumnTypes(#"Added Score_Label", {{"EmployeeID", Int64.Type}, {"Dimension_Code", type text}, {"Score", Int64.Type}})
in
    Typed"""

LOG_M = """let
    """ + m_embedded("Cleaning_Log", ["Step", "Action", "Rows_Affected", "Reason", "Rows_After"]) + """,
    Typed = Table.TransformColumnTypes(Source, {{"Step", type text}, {"Action", type text}, {"Rows_Affected", Int64.Type}, {"Reason", type text}, {"Rows_After", Int64.Type}}),
    Result = Table.AddIndexColumn(Typed, "Log_Order", 1, 1, Int64.Type)
in
    Result"""

FLAGS_M = """let
    """ + m_embedded("Data_Quality_Flags", ["Issue", "Rows", "Treatment"]) + """,
    Typed = Table.TransformColumnTypes(Source, {{"Issue", type text}, {"Rows", Int64.Type}, {"Treatment", type text}}),
    Result = Table.AddIndexColumn(Typed, "Flag_Order", 1, 1, Int64.Type)
in
    Result"""


def static_m(cols_types, rows):
    t = ", ".join(f"{n} = {mt}" for n, mt in cols_types)

    def mlit(v):
        if isinstance(v, (int, float)):
            return repr(v)
        return '"' + v.replace('"', '""') + '"'
    body = ",\n            ".join("{" + ", ".join(mlit(v) for v in r) + "}" for r in rows)
    return f"""let
    Source = #table(
        type table [{t}],
        {{
            {body}
        }})
in
    Source"""


# ---- Driver_Impact: tính sẵn từ dữ liệu sạch (bảng đòn bẩy cho trang 6) ----
def build_drivers():
    d = pd.read_excel(CLEAN_XLSX, sheet_name="Data_Clean")
    ns = d["MaritalStatus"] != "Single"
    y2 = d["YearsAtCompany"] >= 2
    K, P_, N = "Kiểm soát được", "Kiểm soát một phần", "Không kiểm soát được"
    specs = [
        ("Làm thêm giờ", d["OverTime_Flag"] == 1, None, "Có", "Không", K),
        ("Thu nhập < 3K", d["MonthlyIncome"] < 3000, None, "<3K", "≥3K", K),
        ("Không có cổ phiếu*", d["StockOptionLevel"] == 0, ns, "Mức 0", "Mức 1-3", K),
        ("Đi công tác thường xuyên", d["BusinessTravel"] == "Travel_Frequently", None, "Thường xuyên", "Ít/không", K),
        ("Cân bằng CV-CS = Bad", d["WorkLifeBalance"] == 1, None, "Bad", "Khác", K),
        ("Gắn kết công việc = Low", d["JobInvolvement"] == 1, None, "Low", "Khác", K),
        ("Môi trường = Low", d["EnvironmentSatisfaction"] == 1, None, "Low", "Khác", K),
        ("Vừa đổi quản lý*", d["YearsWithCurrentManager"] == 0, y2, "Vừa đổi", "Ổn định", K),
        ("Không được đào tạo", d["TrainingTimesLastYear"] == 0, None, "0 lần", "≥1 lần", K),
        ("Trả thấp hơn cùng cấp", d["Underpaid_vs_Level"] == 1, None, "<85%", "≥85%", K),
        ("Cấp bậc Level 1", d["JobLevel"] == 1, None, "Level 1", "Level 2-5", P_),
        ("≥5 năm chưa thăng chức", d["Promotion_Stagnation"] == 1, None, "≥5 năm", "<5 năm", P_),
        ("Thâm niên 0-1 năm", d["YearsAtCompany"] <= 1, None, "0-1 năm", ">1 năm", N),
        ("Tuổi 18-25", d["Age"] <= 25, None, "18-25", "26+", N),
        ("Độc thân", d["MaritalStatus"] == "Single", None, "Độc thân", "Khác", N),
        ("Nhà xa > 15 dặm", d["DistanceFromHome"] > 15, None, ">15", "≤15", N),
    ]
    order = {K: 1, P_: 2, N: 3}
    rows = []
    for name, risk, scope, g1, g0, ctrl in specs:
        sub = d if scope is None else d[scope]
        r = risk if scope is None else risk[scope]
        r1, r0 = sub.loc[r, "Attrition_Flag"].mean(), sub.loc[~r, "Attrition_Flag"].mean()
        rows.append((name, g1, round(float(r1), 4), int(r.sum()), g0, round(float(r0), 4),
                     round(float(r1 / r0), 2), ctrl, order[ctrl]))
    rows.sort(key=lambda x: -x[6])
    return [r + (i + 1,) for i, r in enumerate(rows)]


DRIVERS = build_drivers()

ACTIONS = [
    (1, "Giới hạn làm thêm giờ", "Level 1 & nhân viên < 1 năm",
     "Level 1 làm thêm giờ nghỉ 52,6%; tránh được ~35% số ca nghỉ việc"),
    (2, "Mở rộng quyền chọn cổ phiếu", "Nhân viên độc thân & Level 1",
     "Không cổ phiếu 21,1% vs 9,9% (đã loại yếu tố hôn nhân); 100% NV độc thân chưa có cổ phiếu"),
    (3, "Chương trình 90 ngày đầu + bàn giao khi đổi quản lý", "Nhân viên mới, người vừa đổi quản lý",
     "Năm đầu nghỉ 34,9%; vừa đổi quản lý 23,0% vs 12,4%"),
    (4, "Nâng mức sàn thu nhập, hạn chế công tác", "Nhóm < 3K, Sales Representative",
     "<3K nghỉ 28,6%; Sales Rep 39,8%; công tác thường xuyên 24,9%"),
    (5, "KHÔNG ưu tiên: cân bằng lương nội bộ, đẩy nhanh thăng chức", "-",
     "Trả thấp hơn cùng cấp chỉ +2,2 điểm %; chậm thăng chức không làm tăng nghỉ việc"),
]

DIMS = {
    "Dim_Department": ("Department", ["Research & Development", "Sales", "Human Resources"]),
    "Dim_AgeGroup": ("Age_Group", ["18-25", "26-35", "36-45", "46-55", "56-60"]),
    "Dim_IncomeBand": ("Income_Band", ["<3K", "3-5K", "5-10K", "10-15K", "15K+"]),
    "Dim_Tenure": ("Tenure_Group", ["0-1 năm", "2-5 năm", "6-10 năm", "11-20 năm", "20+ năm"]),
}

TABLES = {
    "Employees": dict(cols=EMP_SOURCE + EMP_ADDED, m=EMPLOYEES_M,
                      desc="Bảng sự kiện chính: 1 dòng = 1 nhân viên (dữ liệu đã làm sạch)"),
}
for tname, (col, vals) in DIMS.items():
    TABLES[tname] = dict(
        cols=[c(col, S, sort=f"{col}_Sort"), c(f"{col}_Sort", I, "0", hidden=True)],
        m=static_m([(col, "text"), (f"{col}_Sort", "Int64.Type")], [(v, i + 1) for i, v in enumerate(vals)]),
        desc=f"Dimension {col}")
TABLES.update({
    "Satisfaction_Long": dict(
        cols=[c("EmployeeID", I, "0"), c("Dimension_Code", S, hidden=True), c("Score", I, "0"),
              c("Dimension", S, sort="Dimension_Sort"), c("Dimension_Sort", I, "0", hidden=True),
              c("Score_Label", S, sort="Score")],
        m=SAT_M, desc="Unpivot 5 thang hài lòng/gắn kết (1-4)"),
    "Driver_Impact": dict(
        cols=[c("Factor", S, sort="Factor_Rank"), c("Risk_Group_Label", S), c("Rate_Risk", D, "0.0%", "sum"),
              c("N_Risk", I, "#,##0", "sum"), c("Base_Group_Label", S), c("Rate_Base", D, "0.0%", "sum"),
              c("Lift", D, "0.00", "sum"), c("Controllable", S, sort="Controllable_Sort"),
              c("Controllable_Sort", I, "0", hidden=True), c("Factor_Rank", I, "0", hidden=True)],
        m=static_m([("Factor", "text"), ("Risk_Group_Label", "text"), ("Rate_Risk", "number"),
                    ("N_Risk", "Int64.Type"), ("Base_Group_Label", "text"), ("Rate_Base", "number"),
                    ("Lift", "number"), ("Controllable", "text"), ("Controllable_Sort", "Int64.Type"),
                    ("Factor_Rank", "Int64.Type")], DRIVERS),
        desc="Bảng đòn bẩy: tỷ lệ nghỉ nhóm rủi ro vs nhóm còn lại (* = đã loại yếu tố gây nhiễu)"),
    "Action_Plan": dict(
        cols=[c("Priority", I, "0"), c("Lever", S), c("Target", S), c("Evidence", S)],
        m=static_m([("Priority", "Int64.Type"), ("Lever", "text"), ("Target", "text"), ("Evidence", "text")],
                   ACTIONS),
        desc="Kế hoạch hành động ưu tiên"),
    "Cleaning_Log": dict(
        cols=[c("Step", S), c("Action", S), c("Rows_Affected", I, "#,##0", "sum"), c("Reason", S),
              c("Rows_After", I, "#,##0"), c("Log_Order", I, "0")],
        m=LOG_M, desc="Nhật ký làm sạch dữ liệu"),
    "Quality_Flags": dict(
        cols=[c("Issue", S), c("Rows", I, "#,##0", "sum"), c("Treatment", S), c("Flag_Order", I, "0")],
        m=FLAGS_M, desc="Các vấn đề chất lượng dữ liệu và cách xử lý"),
})

RELATIONSHIPS = [
    ("Employees", "Department", "Dim_Department", "Department"),
    ("Employees", "Age_Group", "Dim_AgeGroup", "Age_Group"),
    ("Employees", "Income_Band", "Dim_IncomeBand", "Income_Band"),
    ("Employees", "Tenure_Group", "Dim_Tenure", "Tenure_Group"),
    ("Satisfaction_Long", "EmployeeID", "Employees", "EmployeeID"),
]

# Bảng màu (đã kiểm tra CVD/độ tương phản - xem scripts/pbip_assets/AttritionTheme.json)
BAND, INK_ON_DARK, INK_ON_DARK_2 = "#0F5C55", "#FFFFFF", "#D3EFEB"
INK, INK_2, MUTED = "#1F2328", "#52514E", "#898781"
SURFACE, BORDER = "#FFFFFF", "#E3E5E8"
BRAND, SERIES_1, CONTEXT, CONTEXT_2 = "#0F5C55", "#16867B", "#7CC7BD", "#D3EFEB"
NEUTRAL, NEUTRAL_DARK = "#C3C2B7", "#898781"
ACCENT, SOWHAT_BG = "#EB6834", "#FFF4EE"
CRITICAL, CRITICAL_2, CRITICAL_3, GOOD, WARN = "#D03B3B", "#EE8F8F", "#FAD4D4", "#0CA30C", "#EDA100"
M, W, G, TOP = 16, 1248, 12, 116

# (tên, DAX, format, folder)
F1, F2, F3, F4, F5, F6, F7 = ("1. Tổng quan", "2. Chân dung", "3. Công việc", "4. Lương & đãi ngộ",
                              "5. Hài lòng", "6. Hành động", "7. Chất lượng dữ liệu")
OT1 = "Employees[OverTime_Flag] = 1"
MEASURE_LIST = [
    ("Tổng nhân viên", "COUNTROWS(Employees)", "#,##0", F1),
    ("Số người nghỉ việc", "SUM(Employees[Attrition_Flag])", "#,##0", F1),
    ("Số người ở lại", "[Tổng nhân viên] - [Số người nghỉ việc]", "#,##0", F1),
    ("Tỷ lệ nghỉ việc", "DIVIDE([Số người nghỉ việc], [Tổng nhân viên])", "0.0%", F1),
    ("Tỷ lệ nghỉ (mức chung)", "CALCULATE([Tỷ lệ nghỉ việc], ALLSELECTED(Employees))", "0.0%", F1),
    ("Màu tỷ lệ nghỉ",
     f'IF([Tỷ lệ nghỉ việc] > [Tỷ lệ nghỉ (mức chung)], "{CRITICAL}", "{NEUTRAL}")', "General", F1),
    ("Màu heatmap", [
        "VAR _r = [Tỷ lệ nghỉ việc]",
        "RETURN",
        f'    SWITCH(TRUE(), ISBLANK(_r), BLANK(), _r >= 0.4, "{CRITICAL_2}", _r >= 0.25, "#F6B3B3", '
        f'_r >= 0.15, "{CRITICAL_3}", "#F2F8F7")'], "General", F1),
    ("% nhân viên", "DIVIDE([Tổng nhân viên], CALCULATE([Tổng nhân viên], ALLSELECTED(Employees)))", "0.0%", F1),
    ("% số người nghỉ", "DIVIDE([Số người nghỉ việc], CALCULATE([Số người nghỉ việc], ALLSELECTED(Employees)))",
     "0.0%", F1),
    ("Thu nhập trung vị", "MEDIAN(Employees[MonthlyIncome])", "#,##0", F1),
    ("Thu nhập TV người nghỉ", "CALCULATE([Thu nhập trung vị], Employees[Attrition_Flag] = 1)", "#,##0", F1),
    ("Thu nhập TV người ở lại", "CALCULATE([Thu nhập trung vị], Employees[Attrition_Flag] = 0)", "#,##0", F1),
    ("Tuổi TB", "AVERAGE(Employees[Age])", "0.0", F2),
    ("Thâm niên TB (năm)", "AVERAGE(Employees[YearsAtCompany])", "0.0", F2),
    ("% làm thêm giờ", "AVERAGE(Employees[OverTime_Flag])", "0.0%", F2),
    ("% Level 1", "DIVIDE(CALCULATE([Tổng nhân viên], Employees[JobLevel] = 1), [Tổng nhân viên])", "0.0%", F2),
    ("% độc thân",
     'DIVIDE(CALCULATE([Tổng nhân viên], Employees[MaritalStatus] = "Single"), [Tổng nhân viên])', "0.0%", F2),
    ("% có cổ phiếu",
     "DIVIDE(CALCULATE([Tổng nhân viên], Employees[StockOptionLevel] > 0), [Tổng nhân viên])", "0.0%", F4),
    ("Tỷ lệ nghỉ (thang hài lòng)",
     "AVERAGEX(Satisfaction_Long, RELATED(Employees[Attrition_Flag]))", "0.0%", F5),
    ("Tổng lương năm người nghỉ",
     "CALCULATE(SUM(Employees[Annual_Income]), Employees[Attrition_Flag] = 1)", "#,##0", F6),
    ("Tỷ lệ chi phí thay thế", "SELECTEDVALUE('Replacement Cost'[Replacement Cost], 0.5)", "0%", F6),
    ("Chi phí thay thế ước tính", "[Tổng lương năm người nghỉ] * [Tỷ lệ chi phí thay thế]", "#,##0", F6),
    ("Ca nghỉ tránh được (giới hạn OT)", [
        "SUMX(",
        "    VALUES(Employees[JobLevel]),",
        "    VAR _noRate = CALCULATE([Tỷ lệ nghỉ việc], Employees[OverTime_Flag] = 0)",
        f"    VAR _n = CALCULATE([Tổng nhân viên], {OT1})",
        f"    VAR _left = CALCULATE([Số người nghỉ việc], {OT1})",
        "    RETURN MAX(_left - _n * _noRate, 0)",
        ")"], "#,##0", F6),
    ("% ca nghỉ tránh được", "DIVIDE([Ca nghỉ tránh được (giới hạn OT)], [Số người nghỉ việc])", "0.0%", F6),
    ("Tỷ lệ nghỉ nếu giới hạn OT",
     "DIVIDE([Số người nghỉ việc] - [Ca nghỉ tránh được (giới hạn OT)], [Tổng nhân viên])", "0.0%", F6),
    ("Chi phí tiết kiệm được", [
        "[Ca nghỉ tránh được (giới hạn OT)]",
        "    * DIVIDE([Tổng lương năm người nghỉ], [Số người nghỉ việc])",
        "    * [Tỷ lệ chi phí thay thế]"], "#,##0", F6),
    ("NV hiện tại rủi ro cao",
     "CALCULATE([Tổng nhân viên], Employees[Attrition_Flag] = 0, Employees[Risk_Flags] >= 4)", "#,##0", F6),
    ("NV hiện tại rủi ro trung bình",
     "CALCULATE([Tổng nhân viên], Employees[Attrition_Flag] = 0, Employees[Risk_Flags] >= 2, "
     "Employees[Risk_Flags] <= 3)", "#,##0", F6),
    ("Lương năm NV rủi ro cao",
     "CALCULATE(SUM(Employees[Annual_Income]), Employees[Attrition_Flag] = 0, Employees[Risk_Flags] >= 4)",
     "#,##0", F6),
    ("Hệ số rủi ro", "SUM(Driver_Impact[Lift])", "0.00", F6),
    ("Số dòng gốc", 'CALCULATE(MAX(Cleaning_Log[Rows_After]), Cleaning_Log[Step] = "R0")', "#,##0", F7),
    ("Số dòng sạch", 'CALCULATE(MAX(Cleaning_Log[Rows_After]), Cleaning_Log[Step] = "R10")', "#,##0", F7),
    ("Số dòng bị xoá", "[Số dòng gốc] - [Số dòng sạch]", "#,##0", F7),
    ("Số vấn đề chất lượng", "COUNTROWS(Quality_Flags)", "#,##0", F7),
]
MEASURE_NAMES = {m_[0] for m_ in MEASURE_LIST}

COLS = {t: {x["name"] for x in spec["cols"]} for t, spec in TABLES.items()}
COLS["Replacement Cost"] = {"Replacement Cost"}


def check_dax():
    for name, expr, *_ in MEASURE_LIST:
        text = expr if isinstance(expr, str) else "\n".join(expr)
        for tbl, col in re.findall(r"('?[A-Za-z_][A-Za-z0-9_ ]*'?)\[([^\]]+)\]", text):
            tbl = tbl.strip("'")
            assert tbl in COLS and col in COLS[tbl], f"{name}: {tbl}[{col}]"
        bare = re.sub(r"'?[A-Za-z_][A-Za-z0-9_ ]*'?\[[^\]]+\]", "", text)
        for ref in re.findall(r"\[([^\]]+)\]", bare):
            assert ref in MEASURE_NAMES, f"{name}: [{ref}]"
    for spec in TABLES.values():
        for x in spec["cols"]:
            if x["sort"]:
                assert x["sort"] in {y["name"] for y in spec["cols"]}, x


# =============================================================================
# 2. TMDL WRITER
# =============================================================================
T = "\t"


def tmdl_column(tname, x):
    out = [f"{T}column {q(x['name'])}", f"{T*2}dataType: {x['dt']}"]
    if x["hidden"]:
        out.append(f"{T*2}isHidden")
    if x["fmt"]:
        out.append(f"{T*2}formatString: {x['fmt']}")
    out += [f"{T*2}lineageTag: {lt(tname + '/' + x['name'])}",
            f"{T*2}summarizeBy: {x['summ']}",
            f"{T*2}sourceColumn: {x['name']}"]
    if x["sort"]:
        out.append(f"{T*2}sortByColumn: {q(x['sort'])}")
    out += ["", f"{T*2}annotation SummarizationSetBy = Automatic", ""]
    return out


def tmdl_m_partition(pname, m_code):
    out = [f"{T}partition {q(pname)} = m", f"{T*2}mode: import", f"{T*2}source ="]
    out += [f"{T*4}{line}" if line else "" for line in m_code.splitlines()]
    return out + [""]


def tmdl_table(tname, spec):
    out = [f"/// {spec['desc']}", f"table {q(tname)}", f"{T}lineageTag: {lt(tname)}", ""]
    for x in spec["cols"]:
        out += tmdl_column(tname, x)
    out += tmdl_m_partition(tname, spec["m"])
    out += [f"{T}annotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def tmdl_measures():
    out = ["/// Bảng chứa toàn bộ measure DAX", f"table {MEASURES}", f"{T}lineageTag: {lt(MEASURES)}", ""]
    for name, expr, fmt, folder in MEASURE_LIST:
        if isinstance(expr, str):
            out.append(f"{T}measure {q(name)} = {expr}")
        else:
            out.append(f"{T}measure {q(name)} =")
            out += [f"{T*3}{line}" for line in expr]
        out += [f"{T*2}formatString: {fmt}", f"{T*2}displayFolder: {folder}",
                f"{T*2}lineageTag: {lt('m/' + name)}", ""]
    out += [f"{T}column Dummy", f"{T*2}dataType: int64", f"{T*2}isHidden", f"{T*2}formatString: 0",
            f"{T*2}lineageTag: {lt(MEASURES + '/Dummy')}", f"{T*2}summarizeBy: none",
            f"{T*2}sourceColumn: Dummy", "", f"{T*2}annotation SummarizationSetBy = Automatic", ""]
    out += tmdl_m_partition(MEASURES, "let\n    Source = #table(type table [Dummy = Int64.Type], {{1}})\nin\n    Source")
    out += [f"{T}annotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def tmdl_whatif():
    return "\n".join([
        "/// Tham số What-if: chi phí thay thế 1 nhân viên = % lương năm",
        "table 'Replacement Cost'", f"{T}lineageTag: {lt('Replacement Cost')}", "",
        f"{T}column 'Replacement Cost'", f"{T*2}formatString: 0%",
        f"{T*2}lineageTag: {lt('Replacement Cost/col')}", f"{T*2}summarizeBy: none",
        f"{T*2}sourceColumn: [Value]", "",
        f"{T*2}extendedProperty ParameterMetadata =", f"{T*4}{{", f'{T*4}  "version": 0', f"{T*4}}}", "",
        f"{T*2}annotation SummarizationSetBy = User", "",
        f"{T}partition 'Replacement Cost' = calculated", f"{T*2}mode: import",
        f"{T*2}source = GENERATESERIES(0.25, 1.5, 0.25)", "",
        f"{T}annotation PBI_Id = {hid('replacement')}", ""])


def write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def wjson(path: Path, obj):
    write(path, json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def build_model():
    d = SM / "definition"
    wjson(SM / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": lt("semanticmodel")}})
    wjson(SM / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.0", "settings": {}})
    wjson(SM / ".pbi" / "editorSettings.json", {
        "version": "1.0", "autodetectRelationships": False, "parallelQueryLoading": True,
        "typeDetectionEnabled": True, "relationshipImportEnabled": False,
        "shouldNotifyUserOfNameConflictResolution": True})
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n")
    names = list(TABLES) + [MEASURES, "Replacement Cost"]
    write(d / "model.tmdl", "\n".join([
        "model Model", f"{T}culture: en-US", f"{T}defaultPowerBIDataSourceVersion: powerBI_V3",
        f"{T}sourceQueryCulture: en-US", f"{T}dataAccessOptions", f"{T*2}legacyRedirects",
        f"{T*2}returnErrorValuesAsNull", "",
        f"annotation PBI_QueryOrder = {json.dumps(list(TABLES) + [MEASURES], ensure_ascii=False)}", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", ""]
        + [f"ref table {q(n)}" for n in names]) + "\n")
    rel = []
    for ft, fc, tt, tc in RELATIONSHIPS:
        rel += [f"relationship {lt(f'rel/{ft}.{fc}->{tt}.{tc}')}",
                f"{T}fromColumn: {q(ft)}.{q(fc)}", f"{T}toColumn: {q(tt)}.{q(tc)}", ""]
    write(d / "relationships.tmdl", "\n".join(rel))
    for tname, spec in TABLES.items():
        write(d / "tables" / f"{tname}.tmdl", tmdl_table(tname, spec))
    write(d / "tables" / f"{MEASURES}.tmdl", tmdl_measures())
    write(d / "tables" / "Replacement Cost.tmdl", tmdl_whatif())


# =============================================================================
# 3. REPORT (PBIR) HELPERS
# =============================================================================
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"


def lit(v):
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return f"{v}L"
    if isinstance(v, float):
        return f"{v}D"
    return "'" + str(v).replace("'", "''") + "'"


def p(v):
    return {"expr": {"Literal": {"Value": lit(v)}}}


def color(hexv):
    return {"solid": {"color": p(hexv)}}


def fcol(table, col):
    assert col in COLS[table], f"{table}[{col}]"
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": col}}


def fmeas(name):
    assert name in MEASURE_NAMES, name
    return {"Measure": {"Expression": {"SourceRef": {"Entity": MEASURES}}, "Property": name}}


def color_by_measure(name):
    return {"solid": {"color": {"expr": fmeas(name)}}}


AGG = {"sum": (0, "Sum"), "avg": (1, "Avg")}


def field(spec):
    """'Table.Col' | 'm:Measure' | 'sum:Table.Col'"""
    if spec.startswith("m:"):
        n = spec[2:]
        return fmeas(n), f"{MEASURES}.{n}"
    if ":" in spec:
        fn, rest = spec.split(":", 1)
        t, col = rest.split(".", 1)
        code, lbl = AGG[fn]
        return {"Aggregation": {"Expression": fcol(t, col), "Function": code}}, f"{lbl}({t}.{col})"
    t, col = spec.split(".", 1)
    return fcol(t, col), f"{t}.{col}"


def projections(specs):
    out = []
    for s in specs:
        disp = None
        if isinstance(s, tuple):
            s, disp = s
        f, ref = field(s)
        pr = {"field": f, "queryRef": ref}
        if disp:
            pr["displayName"] = disp
        out.append(pr)
    return out


def vfilter(spec, values):
    """Bộ lọc mức visual: Table.Col IN (values)."""
    t, col = spec.split(".", 1)
    return {"name": "Filter_" + hid(f"{spec}/{values}"), "field": fcol(t, col), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "x", "Entity": t, "Type": 0}],
                       "Where": [{"Condition": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "x"}},
                                                       "Property": col}}],
                           "Values": [[{"Literal": {"Value": lit(v)}}] for v in values]}}}]},
            "howCreated": "User"}


WILDCARD = {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}
ASC, DESC = "Ascending", "Descending"
RATE = "m:Tỷ lệ nghỉ việc"


class Page:
    """Lưới 1280x720: lề 16px, gutter 12px.
    Dải tiêu đề (0-56) · hàng slicer (64-108) · nội dung (116-640) · ô SO WHAT (648-712)."""

    SLICERS = [("Dim_Department.Department", "Phòng ban", "sync_dept"),
               ("Employees.JobLevel_Label", "Cấp bậc", "sync_level"),
               ("Dim_AgeGroup.Age_Group", "Nhóm tuổi", "sync_age"),
               ("Employees.OverTime_Label", "Làm thêm giờ", "sync_ot")]

    def __init__(self, key, display, title, subtitle, sowhat=None, slicers=True):
        self.key, self.name, self.display = key, hid("page/" + key), display
        self.visuals, self.z = [], 1000
        self.extra = {}
        self.top = TOP if slicers else 64
        self.textbox("header", 0, 0, 1280, 56, [
            [(title, {"fontSize": "16pt", "fontWeight": "bold", "color": INK_ON_DARK})],
            [(subtitle, {"fontSize": "9pt", "color": INK_ON_DARK_2})]], bg=BAND, border=None)
        if slicers:
            for i, (fld, disp, grp) in enumerate(self.SLICERS):
                self.slicer(f"slicer_{grp}", M + i * 212, 64, 200, 44, fld, disp, grp)
            self.textbox("legend_note", M + 4 * 212, 64, W - 4 * 212, 44, [
                [("■ ", {"fontSize": "11pt", "color": CRITICAL}),
                 ("Đỏ = tỷ lệ nghỉ cao hơn mức chung của vùng đang lọc   ", {"fontSize": "8pt", "color": INK_2}),
                 ("■ ", {"fontSize": "11pt", "color": NEUTRAL}),
                 ("Xám = thấp hơn", {"fontSize": "8pt", "color": INK_2})]])
        if sowhat:
            self.textbox("sowhat", M, 648, W, 64, [
                [("SO WHAT?   ", {"fontSize": "10pt", "fontWeight": "bold", "color": ACCENT}),
                 (sowhat[0], {"fontSize": "10pt", "fontWeight": "bold", "color": INK})],
                [(sowhat[1], {"fontSize": "9pt", "color": INK_2})]],
                bg=SOWHAT_BG, border=ACCENT)

    def add(self, key, x, y, w, h, visual, filters=None):
        self.z += 100
        v = {"$schema": f"{SCHEMA}/visualContainer/2.0.0/schema.json",
             "name": hid(f"{self.key}/{key}"),
             "position": {"x": x, "y": y, "z": self.z, "height": h, "width": w, "tabOrder": self.z},
             "visual": visual}
        if filters:
            v["filterConfig"] = {"filters": filters}
        self.visuals.append(v)

    def chart(self, key, vtype, x, y, w, h, roles, title=None, sort=None, objects=None, labels=False,
              fill=None, highlight=None, cf_fill=None, hide_value_axis=False, legend_top=False, filters=None):
        """fill: màu mặc định · highlight: (field, {giá trị: màu}) · cf_fill: measure trả về mã màu."""
        qs = {r: {"projections": projections(f)} for r, f in roles.items()}
        v = {"visualType": vtype, "query": {"queryState": qs}, "drillFilterOtherVisuals": True}
        if sort:
            fspec, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field(fspec)[0], "direction": direction}],
                                            "isDefaultSort": False}
        obj = dict(objects or {})
        if labels:
            obj.setdefault("labels", [{"properties": {"show": p(True), "color": color(INK_2)}}])
        points = []
        if fill:
            points.append({"properties": {"fill": color(fill)}})
        if highlight:
            hfield, mapping = highlight
            for val, hexv in mapping.items():
                points.append({"properties": {"fill": color(hexv)},
                               "selector": {"data": [{"scopeId": {"Comparison": {
                                   "ComparisonKind": 0, "Left": field(hfield)[0],
                                   "Right": {"Literal": {"Value": lit(val)}}}}}]}})
        if cf_fill:
            points.append({"properties": {"fill": color_by_measure(cf_fill)}, "selector": WILDCARD})
        if points:
            obj["dataPoint"] = points
        if hide_value_axis:
            obj["valueAxis"] = [{"properties": {"show": p(False)}}]
        if legend_top:
            obj["legend"] = [{"properties": {"show": p(True), "position": p("Top")}}]
        if obj:
            v["objects"] = obj
        if title:
            v["visualContainerObjects"] = {"title": [{"properties": {"show": p(True), "text": p(title)}}]}
        self.add(key, x, y, w, h, v, filters)

    def rate_bars(self, key, vtype, x, y, w, h, cat, title, sort=True, measure=RATE, **kw):
        """Tỷ lệ nghỉ theo nhóm: nhãn số, ẩn trục, đỏ khi cao hơn mức chung (màu theo measure)."""
        cat_field = cat[0] if isinstance(cat, tuple) else cat
        self.chart(key, vtype, x, y, w, h, {"Category": [cat], "Y": [(measure, "Tỷ lệ nghỉ việc")]}, title,
                   sort=(cat_field, ASC) if sort is True else sort, labels=True, hide_value_axis=True,
                   fill=NEUTRAL, cf_fill="Màu tỷ lệ nghỉ", **kw)

    def heatmap(self, key, x, y, w, h, rows, cols, title):
        """Ma trận tỷ lệ nghỉ với màu nền theo measure 'Màu heatmap'."""
        obj = {"values": [{"properties": {"backColor": color_by_measure("Màu heatmap")},
                           "selector": {**WILDCARD, "metadata": f"{MEASURES}.Tỷ lệ nghỉ việc"}}],
               "subTotals": [{"properties": {"rowSubtotals": p(False), "columnSubtotals": p(False)}}]}
        self.chart(key, "pivotTable", x, y, w, h,
                   {"Rows": [rows], "Columns": [cols], "Values": [(RATE, "Tỷ lệ nghỉ việc")]}, title,
                   objects=obj)

    def card(self, key, x, y, w, h, measure, label=None, value_color=None):
        obj = {"labels": [{"properties": {"color": color(value_color or BRAND), "fontSize": p(22.0)}}],
               "categoryLabels": [{"properties": {"show": p(True), "color": color(INK_2), "fontSize": p(9.0)}}]}
        v = {"visualType": "card",
             "query": {"queryState": {"Values": {"projections": projections(
                 [(f"m:{measure}", label) if label else f"m:{measure}"])}}},
             "objects": obj, "drillFilterOtherVisuals": True,
             "visualContainerObjects": {
                 "title": [{"properties": {"show": p(False)}}],
                 "border": [{"properties": {"show": p(True), "color": color(BORDER), "radius": p(10.0)}}]}}
        self.add(key, x, y, w, h, v)

    def slicer(self, key, x, y, w, h, fld, disp, group, dropdown=True, single=False):
        obj = {}
        if dropdown:
            obj["data"] = [{"properties": {"mode": p("Dropdown")}}]
        if single:
            obj["selection"] = [{"properties": {"singleSelect": p(True)}}]
        v = {"visualType": "slicer",
             "query": {"queryState": {"Values": {"projections": projections([(fld, disp)])}}},
             "objects": obj, "drillFilterOtherVisuals": True,
             "visualContainerObjects": {"title": [{"properties": {"show": p(False)}}]}}
        if group:
            v["syncGroup"] = {"groupName": group, "fieldChanges": True, "filterChanges": True}
        self.add(key, x, y, w, h, v)

    def textbox(self, key, x, y, w, h, paragraphs, bg=SURFACE, border=BORDER):
        paras = [{"textRuns": [{"value": t, "textStyle": {"fontFamily": "Segoe UI", **st}} for t, st in para]}
                 for para in paragraphs]
        v = {"visualType": "textbox",
             "objects": {"general": [{"properties": {"paragraphs": paras}}]},
             "drillFilterOtherVisuals": True,
             "visualContainerObjects": {
                 "title": [{"properties": {"show": p(False)}}],
                 "background": [{"properties": {"show": p(True), "color": color(bg), "transparency": p(0.0)}}],
                 "border": ([{"properties": {"show": p(True), "color": color(border), "radius": p(10.0)}}]
                            if border else [{"properties": {"show": p(False)}}])}}
        self.add(key, x, y, w, h, v)

    def json(self):
        page = {"$schema": f"{SCHEMA}/page/1.4.0/schema.json", "name": self.name,
                "displayName": self.display, "displayOption": "FitToPage",
                "height": 720, "width": 1280}
        page.update(self.extra)
        return page


def grid(n, x0=None, width=None, gap=G):
    """Chia đều `width` thành n cột -> [(x, w), ...]."""
    x0 = M if x0 is None else x0
    width = W if width is None else width
    w = (width - gap * (n - 1)) / n
    return [(round(x0 + i * (w + gap)), round(w)) for i in range(n)]


def para(text, size="9pt", col=INK_2, bold=False):
    st = {"fontSize": size, "color": col}
    if bold:
        st["fontWeight"] = "bold"
    return [(text, st)]


LEFT_STAY = {"Nghỉ việc": CRITICAL, "Ở lại": CONTEXT}


# =============================================================================
# 4. PAGES
# =============================================================================
def build_pages():
    pages = []
    BOT = 640
    C3 = grid(3)
    ROW1_H = 256

    # ---- Trang 1: Tổng quan --------------------------------------------------
    pg = Page("overview", "1. Tổng quan",
              "Cứ 6 nhân viên có 1 người rời đi - và họ rời đi từ tuyến đầu",
              "Tổng quan nghỉ việc  ·  1.470 nhân viên  ·  tỷ lệ chung 16,1%",
              ("Sales Representative mất 39,8% nhân sự; quản lý & giám đốc chỉ 2,5-4,9%. "
               "Người nghỉ có lương trung vị thấp hơn 38% so với người ở lại.",
               "Vấn đề không nằm ở cấp lãnh đạo mà ở vị trí tuyến đầu. Lưu ý: R&D có tỷ lệ thấp nhất nhưng "
               "số người nghỉ nhiều nhất (133) - luôn đọc tỷ lệ cùng với số lượng."))
    y, h = pg.top, 88
    for i, ((x, w), (m_, col)) in enumerate(zip(grid(6), [
            ("Tổng nhân viên", None), ("Số người nghỉ việc", CRITICAL), ("Tỷ lệ nghỉ việc", CRITICAL),
            ("Thu nhập TV người nghỉ", CRITICAL), ("Thu nhập TV người ở lại", None),
            ("Chi phí thay thế ước tính", None)])):
        pg.card(f"kpi{i}", x, y, w, h, m_, value_color=col)
    y2 = y + h + G
    hh = (BOT - y2 - G) // 2
    pg.rate_bars("dept_rate", "clusteredColumnChart", M, y2, 304, hh,
                 ("Dim_Department.Department", "Phòng ban"), "Tỷ lệ nghỉ theo phòng ban")
    pg.chart("dept_count", "clusteredColumnChart", M, y2 + hh + G, 304, BOT - (y2 + hh + G),
             {"Category": [("Dim_Department.Department", "Phòng ban")],
              "Y": [("m:Số người nghỉ việc", "Số người nghỉ")]},
             "...nhưng R&D có số người nghỉ nhiều nhất", sort=("Dim_Department.Department", ASC),
             labels=True, hide_value_axis=True, fill=BRAND)
    pg.rate_bars("role_rate", "clusteredBarChart", M + 304 + G, y2, 616, BOT - y2,
                 ("Employees.JobRole", "Vị trí"), "Tỷ lệ nghỉ theo vị trí: tuyến đầu cao gấp nhiều lần quản lý",
                 sort=(RATE, DESC))
    pg.chart("donut", "donutChart", M + 304 + G + 616 + G, y2, W - 304 - 616 - 2 * G, BOT - y2,
             {"Category": [("Employees.Attrition_Label", "Trạng thái")], "Y": [("m:Tổng nhân viên", "Nhân viên")]},
             "Cơ cấu nhân viên: ở lại vs nghỉ việc", labels=True, legend_top=True,
             highlight=("Employees.Attrition_Label", LEFT_STAY), sort=("Employees.Attrition_Label", ASC))
    pages.append(pg)

    # ---- Trang 2: Chân dung --------------------------------------------------
    pg = Page("profile", "2. Chân dung người nghỉ việc",
              "Người rời đi: trẻ, mới vào, cấp thấp, độc thân",
              "Ai nghỉ việc? (các yếu tố KHÔNG kiểm soát được - dùng để xác định ai cần bảo vệ)",
              ("18-25 tuổi nghỉ 35,8%; năm đầu 34,9%; Level 1 là 26,3%; độc thân 25,5%. "
               "60% người nghỉ là Level 1, 42% dưới 30 tuổi.",
               "Giới tính và học vấn gần như không tạo khác biệt. Công ty không đổi được tuổi hay hôn nhân, "
               "nhưng biết chính xác nhóm nào cần ưu tiên giữ chân."))
    y = pg.top
    pg.rate_bars("age", "clusteredColumnChart", C3[0][0], y, C3[0][1], ROW1_H,
                 ("Dim_AgeGroup.Age_Group", "Nhóm tuổi"), "Tỷ lệ nghỉ theo tuổi: 18-25 cao gấp 4 lần 36-45")
    pg.rate_bars("tenure", "clusteredColumnChart", C3[1][0], y, C3[1][1], ROW1_H,
                 ("Dim_Tenure.Tenure_Group", "Thâm niên"), "Theo thâm niên: năm đầu là nguy hiểm nhất")
    pg.rate_bars("level", "clusteredColumnChart", C3[2][0], y, C3[2][1], ROW1_H,
                 ("Employees.JobLevel_Label", "Cấp bậc"), "Theo cấp bậc: Level 1 cao nhất")
    y2 = y + ROW1_H + G
    pg.rate_bars("marital", "clusteredColumnChart", C3[0][0], y2, C3[0][1], BOT - y2,
                 ("Employees.Marital_Label", "Hôn nhân"), "Theo hôn nhân: độc thân gấp đôi")
    hh = (BOT - y2 - G) // 2
    pg.rate_bars("gender", "clusteredBarChart", C3[1][0], y2, C3[1][1], hh,
                 ("Employees.Gender_Label", "Giới tính"), "Giới tính: chênh lệch nhỏ")
    pg.rate_bars("edu", "clusteredBarChart", C3[1][0], y2 + hh + G, C3[1][1], BOT - (y2 + hh + G),
                 ("Employees.Education_Label", "Học vấn"), "Học vấn: chênh lệch nhỏ")
    pg.chart("profile_table", "tableEx", C3[2][0], y2, C3[2][1], BOT - y2,
             {"Values": [("Employees.Attrition_Label", "Nhóm"), ("m:Tổng nhân viên", "Số NV"),
                         ("m:Tuổi TB", "Tuổi TB"), ("m:Thâm niên TB (năm)", "Thâm niên"),
                         ("m:Thu nhập trung vị", "Lương TV"), ("m:% Level 1", "% Level 1"),
                         ("m:% làm thêm giờ", "% OT"), ("m:% độc thân", "% độc thân")]},
             "So sánh người nghỉ vs người ở lại", sort=("Employees.Attrition_Label", ASC))
    pages.append(pg)

    # ---- Trang 3: Công việc --------------------------------------------------
    pg = Page("work", "3. Điều kiện làm việc",
              "Làm thêm giờ là yếu tố nguy hiểm nhất - nhất là với người mới",
              "Công việc & điều kiện làm việc (các đòn bẩy KIỂM SOÁT ĐƯỢC)",
              ("Level 1 làm thêm giờ nghỉ 52,6%; nhân viên năm đầu làm thêm giờ nghỉ 55,1%. "
               "156 NV Level 1 làm thêm giờ chiếm 82 người nghỉ = 35% tổng số người nghỉ.",
               "Đòn bẩy số 1: giới hạn làm thêm giờ cho Level 1 và nhân viên mới - nhanh, gần như không tốn tiền, "
               "mô phỏng có thể đưa tỷ lệ nghỉ từ 16,1% xuống ~10,5%."))
    y = pg.top
    pg.rate_bars("ot", "clusteredColumnChart", M, y, 304, ROW1_H,
                 ("Employees.OverTime_Label", "Làm thêm giờ"), "Làm thêm giờ: gấp 3 lần",
                 sort=(RATE, DESC))
    pg.heatmap("ot_level", M + 304 + G, y, 616, ROW1_H, ("Employees.OverTime_Label", "Làm thêm giờ"),
               ("Employees.JobLevel_Label", "Cấp bậc"), "Làm thêm giờ × Cấp bậc: Level 1 + OT = 52,6%")
    pg.rate_bars("travel", "clusteredColumnChart", M + 304 + G + 616 + G, y, W - 304 - 616 - 2 * G, ROW1_H,
                 ("Employees.BusinessTravel_Label", "Công tác"), "Đi công tác thường xuyên: 24,9%")
    y2 = y + ROW1_H + G
    pg.heatmap("ot_tenure", M, y2, 616, BOT - y2, ("Employees.OverTime_Label", "Làm thêm giờ"),
               ("Dim_Tenure.Tenure_Group", "Thâm niên"), "Làm thêm giờ × Thâm niên: năm đầu + OT = 55,1%")
    (xa, wa), (xb, wb) = grid(2, M + 616 + G, W - 616 - G)
    pg.rate_bars("wlb", "clusteredColumnChart", xa, y2, wa, BOT - y2,
                 ("Employees.WorkLifeBalance_Label", "Cân bằng CV-CS"), "Cân bằng công việc-cuộc sống")
    pg.rate_bars("distance", "clusteredColumnChart", xb, y2, wb, BOT - y2,
                 ("Employees.Distance_Group", "Khoảng cách (dặm)"), "Khoảng cách từ nhà")
    pages.append(pg)

    # ---- Trang 4: Lương & đãi ngộ --------------------------------------------
    pg = Page("pay", "4. Lương & đãi ngộ",
              "Tiền có quan trọng - nhưng không theo cách nhiều người nghĩ",
              "Thu nhập tuyệt đối, cổ phiếu và công bằng lương (chỉ dùng MonthlyIncome - các cột Rate là dữ liệu giả)",
              ("Thu nhập <3K nghỉ 28,6% vs 3,8% ở nhóm 15K+. Không có cổ phiếu: 21,1% vs 9,9% (đã loại yếu tố hôn nhân). "
               "Bị trả thấp hơn đồng nghiệp cùng cấp chỉ +2 điểm %.",
               "Nâng mức sàn thu nhập & mở rộng cổ phiếu cho NV độc thân/Level 1 (100% NV độc thân chưa có cổ phiếu). "
               "Cân bằng lương nội bộ có tác động thấp. Người hiệu suất cao nghỉ ngang người khác → chưa có cơ chế giữ người giỏi."))
    y = pg.top
    pg.rate_bars("income", "clusteredColumnChart", C3[0][0], y, C3[0][1], ROW1_H,
                 ("Dim_IncomeBand.Income_Band", "Thu nhập/tháng"), "Thu nhập < 3K: nghỉ 28,6%")
    pg.rate_bars("stock", "clusteredColumnChart", C3[1][0], y, C3[1][1], ROW1_H,
                 ("Employees.StockOption_Label", "Cổ phiếu"), "Không có cổ phiếu: 24,4%")
    pg.heatmap("stock_marital", C3[2][0], y, C3[2][1], ROW1_H, ("Employees.Marital_Label", "Hôn nhân"),
               ("Employees.StockOption_Label", "Cổ phiếu"),
               "Bẫy gây nhiễu: 100% NV độc thân KHÔNG có cổ phiếu")
    y2 = y + ROW1_H + G
    pg.chart("income_level", "clusteredColumnChart", M, y2, 616, BOT - y2,
             {"Category": [("Employees.JobLevel_Label", "Cấp bậc")],
              "Series": [("Employees.Attrition_Label", "Trạng thái")],
              "Y": [("m:Thu nhập trung vị", "Thu nhập trung vị")]},
             "Thu nhập trung vị theo cấp bậc: người nghỉ vs người ở lại",
             sort=("Employees.JobLevel_Label", ASC), labels=True, hide_value_axis=True, legend_top=True,
             highlight=("Employees.Attrition_Label", LEFT_STAY))
    (xa, wa), (xb, wb) = grid(2, M + 616 + G, W - 616 - G)
    pg.rate_bars("underpaid", "clusteredColumnChart", xa, y2, wa, BOT - y2,
                 ("Employees.Underpaid_Label", "So với cùng cấp"), "Trả thấp hơn cùng cấp: tác động nhỏ")
    pg.rate_bars("perf", "clusteredColumnChart", xb, y2, wb, BOT - y2,
                 ("Employees.PerformanceRating_Label", "Hiệu suất"), "Người giỏi nghỉ ngang người khác")
    pages.append(pg)

    # ---- Trang 5: Hài lòng & gắn kết ------------------------------------------
    pg = Page("satisfaction", "5. Hài lòng & gắn kết",
              "Đổi người quản lý làm tỷ lệ nghỉ tăng gấp đôi; chậm thăng chức không phải lý do",
              "Mức hài lòng, gắn kết, quản lý trực tiếp, thăng chức và đào tạo",
              ("Ở mức 'Thấp', tỷ lệ nghỉ cao gấp 1,5-2 lần (cân bằng CV-CS 31,2%; gắn kết 33,7%). "
               "NV ≥2 năm vừa đổi quản lý nghỉ 23,0% vs 12,4%. ≥5 năm chưa thăng chức chỉ 14,2%.",
               "Thời điểm nguy hiểm là NĂM ĐẦU và lúc ĐỔI QUẢN LÝ, không phải lúc chờ thăng chức → "
               "chương trình đồng hành 90 ngày, quy trình bàn giao khi đổi quản lý, ≥1 lần đào tạo/năm."))
    y = pg.top
    pg.chart("sat_long", "clusteredColumnChart", M, y, 616, ROW1_H,
             {"Category": [("Satisfaction_Long.Dimension", "Thang đo")],
              "Series": [("Satisfaction_Long.Score_Label", "Mức")],
              "Y": [("m:Tỷ lệ nghỉ (thang hài lòng)", "Tỷ lệ nghỉ việc")]},
             "Tỷ lệ nghỉ theo mức hài lòng của 5 thang đo (đỏ = mức Thấp)",
             sort=("Satisfaction_Long.Dimension", ASC), legend_top=True, hide_value_axis=True, labels=True,
             highlight=("Satisfaction_Long.Score_Label",
                        {"1 - Thấp": CRITICAL, "2 - Trung bình": CRITICAL_2, "3 - Cao": CONTEXT,
                         "4 - Rất cao": BRAND}))
    (xa, wa), (xb, wb) = grid(2, M + 616 + G, W - 616 - G)
    pg.rate_bars("involve", "clusteredColumnChart", xa, y, wa, ROW1_H,
                 ("Employees.JobInvolvement_Label", "Gắn kết"), "Mức gắn kết công việc")
    pg.rate_bars("lowsat", "clusteredColumnChart", xb, y, wb, ROW1_H,
                 ("Employees.LowSat_Label", "Số thang Low"), "Càng nhiều thang 'Low', càng dễ nghỉ")
    y2 = y + ROW1_H + G
    pg.rate_bars("mgr", "clusteredColumnChart", C3[0][0], y2, C3[0][1], BOT - y2,
                 ("Employees.Mgr_Change_Label", "Quản lý"), "Vừa đổi quản lý (NV ≥2 năm): 23,0%")
    pg.rate_bars("promo", "clusteredColumnChart", C3[1][0], y2, C3[1][1], BOT - y2,
                 ("Employees.Promotion_Label", "Thăng chức"), "Chậm thăng chức KHÔNG làm tăng nghỉ việc")
    pg.rate_bars("training", "clusteredColumnChart", C3[2][0], y2, C3[2][1], BOT - y2,
                 ("Employees.Training_Label", "Đào tạo/năm"), "Không được đào tạo: 27,8%")
    pages.append(pg)

    # ---- Trang 6: Rủi ro & hành động -------------------------------------------
    pg = Page("action", "6. Rủi ro & hành động",
              "Khoảng 1/3 số ca nghỉ việc có thể tránh được",
              "Điểm rủi ro cộng dồn, đòn bẩy kiểm soát được, mô phỏng chi phí và kế hoạch hành động",
              ("4-5 yếu tố rủi ro → 70,3% nghỉ việc (0-1 yếu tố: 4,9%). 19 NV hiện tại ở nhóm rủi ro cao, "
               "557 ở mức trung bình. Giới hạn làm thêm giờ tránh được ~83 ca (35%).",
               "Ưu tiên: (1) giới hạn OT cho Level 1 & NV mới, (2) cổ phiếu cho NV độc thân/Level 1, "
               "(3) chương trình 90 ngày + bàn giao quản lý, (4) nâng mức sàn thu nhập. Không ưu tiên cân bằng lương nội bộ."))
    y, h = pg.top, 88
    cards = grid(5, M, 980)
    for i, ((x, w), (m_, col)) in enumerate(zip(cards, [
            ("NV hiện tại rủi ro cao", CRITICAL), ("NV hiện tại rủi ro trung bình", WARN),
            ("Ca nghỉ tránh được (giới hạn OT)", GOOD), ("Tỷ lệ nghỉ nếu giới hạn OT", GOOD),
            ("Chi phí tiết kiệm được", GOOD)])):
        pg.card(f"kpi{i}", x, y, w, h, m_, value_color=col)
    pg.slicer("cost", M + 980 + G, y, W - 980 - G, h, "Replacement Cost.Replacement Cost",
              "Chi phí thay thế = % lương năm (mặc định 50%)", None, single=True)
    y2 = y + h + G
    pg.rate_bars("risk", "clusteredColumnChart", M, y2, 304, BOT - y2,
                 ("Employees.Risk_Group", "Số yếu tố rủi ro"), "Rủi ro cộng dồn: 4-5 yếu tố = 70%")
    pg.chart("drivers", "stackedBarChart", M + 304 + G, y2, 452, BOT - y2,
             {"Category": [("Driver_Impact.Factor", "Yếu tố")],
              "Series": [("Driver_Impact.Controllable", "Loại")],
              "Y": [("m:Hệ số rủi ro", "Hệ số (nhóm rủi ro / nhóm còn lại)")]},
             "Đòn bẩy: hệ số tỷ lệ nghỉ (* đã loại yếu tố gây nhiễu)",
             sort=("Driver_Impact.Factor", ASC), labels=True, hide_value_axis=True, legend_top=True,
             highlight=("Driver_Impact.Controllable",
                        {"Kiểm soát được": BRAND, "Kiểm soát một phần": WARN, "Không kiểm soát được": NEUTRAL}))
    xr, wr = M + 304 + G + 452 + G, W - 304 - 452 - 2 * G
    pg.chart("actions", "tableEx", xr, y2, wr, 250,
             {"Values": [("Action_Plan.Priority", "#"), ("Action_Plan.Lever", "Đòn bẩy"),
                         ("Action_Plan.Target", "Nhóm mục tiêu"), ("Action_Plan.Evidence", "Bằng chứng")]},
             "Kế hoạch hành động ưu tiên", sort=("Action_Plan.Priority", ASC))
    pg.chart("at_risk", "tableEx", xr, y2 + 250 + G, wr, BOT - (y2 + 250 + G),
             {"Values": [("Employees.EmployeeID", "Mã NV"), ("Employees.JobRole", "Vị trí"),
                         ("Employees.JobLevel_Label", "Cấp"), ("sum:Employees.Age", "Tuổi"),
                         ("sum:Employees.YearsAtCompany", "Thâm niên"), ("Employees.OverTime_Label", "OT"),
                         ("sum:Employees.Risk_Flags", "Điểm rủi ro")]},
             "Nhân viên ĐANG LÀM có rủi ro cao - cần liên hệ ngay",
             sort=("sum:Employees.Risk_Flags", DESC),
             filters=[vfilter("Employees.Attrition_Label", ["Ở lại"]),
                      vfilter("Employees.Risk_Group", ["Cao (4-5)"])])
    pages.append(pg)

    # ---- Trang 7: Chất lượng dữ liệu -------------------------------------------
    pg = Page("quality", "7. Chất lượng dữ liệu",
              "Vì sao các con số này đáng tin",
              "Tiền xử lý: giữ 100% dữ liệu, xử lý 4 vấn đề nội dung, tránh 2 bẫy gây nhiễu",
              ("Không xoá dòng nào (1.470 → 1.470). Loại 2 cột hằng số, không dùng 3 cột 'Rate' giả, "
               "gắn cờ 197 hồ sơ NumCompaniesWorked mâu thuẫn.",
               "Hạn chế: dữ liệu dạng snapshot (không có ngày nghỉ), không phân biệt nghỉ tự nguyện/bị cho nghỉ, "
               "chi phí thay thế là giả định, một số nhóm nhỏ (HR 63 NV, WLB 'Bad' 80 NV)."),
              slicers=False)
    y, h = pg.top, 88
    for i, ((x, w), (m_, col)) in enumerate(zip(grid(4), [
            ("Số dòng gốc", None), ("Số dòng bị xoá", GOOD), ("Số dòng sạch", None),
            ("Số vấn đề chất lượng", WARN)])):
        pg.card(f"kpi{i}", x, y, w, h, m_, value_color=col)
    y2 = y + h + G
    pg.chart("flags", "tableEx", M, y2, 616, 200,
             {"Values": [("Quality_Flags.Flag_Order", "#"), ("Quality_Flags.Issue", "Vấn đề"),
                         ("sum:Quality_Flags.Rows", "Số dòng"), ("Quality_Flags.Treatment", "Cách xử lý")]},
             "Các vấn đề chất lượng dữ liệu", sort=("Quality_Flags.Flag_Order", ASC))
    pg.chart("log", "tableEx", M, y2 + 200 + G, 616, BOT - (y2 + 200 + G),
             {"Values": [("Cleaning_Log.Log_Order", "#"), ("Cleaning_Log.Step", "Bước"),
                         ("Cleaning_Log.Action", "Hành động"),
                         ("sum:Cleaning_Log.Rows_Affected", "Số dòng"), ("Cleaning_Log.Reason", "Lý do")]},
             "Nhật ký làm sạch (Cleaning Log)", sort=("Cleaning_Log.Log_Order", ASC))
    xr, wr = M + 616 + G, W - 616 - G
    pg.textbox("traps", xr, y2, wr, BOT - y2, [
        para("2 cái bẫy gây nhiễu đã được tách riêng", "11pt", BRAND, True),
        para(" "),
        para("1. Cổ phiếu ↔ Hôn nhân", "10pt", INK, True),
        para("100% nhân viên độc thân có StockOptionLevel = 0. Nhìn nhanh, 'không cổ phiếu' nghỉ 24,4% - "
             "nhưng một phần là do độc thân. Chỉ xét NV đã kết hôn/ly hôn: không cổ phiếu 21,1% vs có 9,9% "
             "→ hiệu ứng cổ phiếu VẪN có thật."),
        para(" "),
        para("2. Đổi quản lý ↔ Thâm niên", "10pt", INK, True),
        para("202/263 người có YearsWithCurrentManager = 0 là nhân viên mới. Chỉ xét NV ≥ 2 năm: vừa đổi "
             "quản lý 23,0% vs 12,4% → hiệu ứng có thật (mẫu nhỏ: 61 người)."),
        para(" "),
        para("Vì sao không xoá ngoại lai?", "10pt", INK, True),
        para("485 NV có thu nhập/thâm niên vượt ngưỡng IQR - đó là quản lý cấp cao và NV lâu năm có thật. "
             "Chỉ có 237 người nghỉ việc, mỗi bản ghi đều quý cho mô hình."),
        para(" "),
        para("Cột không dùng", "10pt", INK, True),
        para("DailyRate/HourlyRate/MonthlyRate: |r| ≤ 0,03 với MonthlyIncome → không phản ánh lương thật."),
    ])
    pages.append(pg)

    # ---- Trang drill-through ---------------------------------------------------
    pg = Page("detail", "Chi tiết nhân viên",
              "Danh sách nhân viên theo nhóm rủi ro",
              "Drill-through: ở trang 6, chuột phải vào một cột 'Rủi ro cộng dồn' → Drill through → Chi tiết nhân viên",
              slicers=False)
    flt = "Filter_" + hid("dt_risk")
    pg.extra = {
        "filterConfig": {"filters": [{"name": flt, "field": fcol("Employees", "Risk_Group"),
                                      "type": "Categorical", "howCreated": "Drillthrough"}]},
        "pageBinding": {"name": hid("binding_detail"), "type": "Drillthrough",
                        "parameters": [{"name": "Param_" + flt, "boundFilter": flt,
                                        "fieldExpr": fcol("Employees", "Risk_Group")}]},
        "visibility": "HiddenInViewMode"}
    pg.chart("emp_table", "tableEx", M, pg.top, W, BOT + 72 - pg.top,
             {"Values": [("Employees.EmployeeID", "Mã NV"), ("Employees.Attrition_Label", "Trạng thái"),
                         ("Employees.Department", "Phòng ban"), ("Employees.JobRole", "Vị trí"),
                         ("Employees.JobLevel_Label", "Cấp"), ("sum:Employees.Age", "Tuổi"),
                         ("Employees.Marital_Label", "Hôn nhân"), ("sum:Employees.MonthlyIncome", "Thu nhập"),
                         ("sum:Employees.YearsAtCompany", "Thâm niên"), ("Employees.OverTime_Label", "OT"),
                         ("Employees.StockOption_Label", "Cổ phiếu"), ("sum:Employees.Risk_Flags", "Điểm rủi ro")]},
             "Danh sách nhân viên (sắp xếp theo điểm rủi ro)", sort=("sum:Employees.Risk_Flags", DESC))
    pages.append(pg)
    return pages


# =============================================================================
# 5. REPORT WRITER
# =============================================================================
THEME_NAME = "AttritionTheme.json"
BASE_THEME = "CY24SU10"


def build_report():
    d = RP / "definition"
    wjson(RP / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": lt("report")}})
    wjson(RP / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    wjson(d / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    wjson(d / "report.json", {
        "$schema": f"{SCHEMA}/report/1.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": BASE_THEME, "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": THEME_NAME, "reportVersionAtImport": "5.61", "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": THEME_NAME, "path": THEME_NAME, "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True}})
    for src, dst in [(f"{BASE_THEME}.json", f"SharedResources/BaseThemes/{BASE_THEME}.json"),
                     (THEME_NAME, f"RegisteredResources/{THEME_NAME}")]:
        target = RP / "StaticResources" / dst
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(THEME_SRC / src, target)

    pages = build_pages()
    wjson(d / "pages" / "pages.json", {
        "$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [pg.name for pg in pages], "activePageName": pages[0].name})
    for pg in pages:
        wjson(d / "pages" / pg.name / "page.json", pg.json())
        for v in pg.visuals:
            wjson(d / "pages" / pg.name / "visuals" / v["name"] / "visual.json", v)
    return pages


def main():
    check_dax()
    for sub in (SM, RP):
        shutil.rmtree(sub, ignore_errors=True)
    OUT.mkdir(parents=True, exist_ok=True)
    build_model()
    pages = build_report()
    wjson(OUT / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True}})
    write(OUT / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    n_vis = sum(len(pg.visuals) for pg in pages)
    print(f"OK: {len(TABLES) + 2} bảng, {len(MEASURE_LIST)} measure, {len(RELATIONSHIPS)} quan hệ, "
          f"{len(pages)} trang, {n_vis} visual -> {OUT}")


if __name__ == "__main__":
    main()
