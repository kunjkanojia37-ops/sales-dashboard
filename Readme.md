# 📊 Sales Dashboard (GST + HSN Analysis)

This project is a **Streamlit-based Sales Dashboard** that automatically processes raw Excel data and generates:

- 📈 Sales analysis
- 🧾 GST (HSN-wise reports)
- 🏢 B2B & B2C summaries
- 📊 Graphs (Top items, Best days, Payment mode)
- 📥 Exportable Excel reports

---

## 🚀 Features

- Auto-detects header from messy Excel files
- Cleans raw data (ignores metadata rows)
- Generates:
  - HSN summary (5%, 40%, Non-taxable)
  - B2B and B2C reports
- Visual dashboards using Streamlit
- Export reports to Excel

---

## 📁 Project Structure
project/
│
│
├──assets/
│  └── Dashboard Preview
│
├── graphical presentation/
│ └── generated graphs 
│
│
├── output/
│ └── exported reports
│
├── raw date/
│ └── excel_file.xlsx
│  
├── main.py
│
└── README.md
│
├── requirements.txt