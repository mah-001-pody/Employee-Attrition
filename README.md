# Employee Attrition Analysis - VJP205 Business Analytics

Phân tích nghỉ việc của nhân viên (1.470 nhân viên, bộ dữ liệu IBM HR Attrition): **yếu tố nào giữ chân nhân viên, yếu tố nào khiến họ rời đi, và công ty có thể thay đổi điều gì.**

| Thư mục | Nội dung |
|---|---|
| `02_Employee Attrition.xlsx` | Dữ liệu gốc |
| `data/processed/` | Dữ liệu sạch `Attrition_Clean.xlsx` / `.csv` |
| `scripts/01_data_cleaning.py` | Tiền xử lý và làm sạch |
| `docs/01_Tien_xu_ly_du_lieu.md` | Bài viết phần tiền xử lý cho báo cáo |
| `scripts/02_build_pbip.py` | Sinh dự án Power BI (.pbip) |
| `powerbi/` | Dự án Power BI: `Attrition_Dashboard.pbip` + hướng dẫn `HUONG_DAN.md` |
| `dist/Attrition_Dashboard.zip` | **Gói tải về để mở trên Windows** |
| `VJP205 Guideline ... .docx` | Đề bài |

```bash
pip install pandas openpyxl numpy
python scripts/01_data_cleaning.py
python scripts/02_build_pbip.py
```

**Tải dashboard:** https://github.com/mah-001-pody/Employee-Attrition/raw/claude/serene-pasteur-rcxanv/dist/Attrition_Dashboard.zip
