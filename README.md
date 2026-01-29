# 📊 Automated Sales Reporting & Data Cleaning Pipeline
**Status:** Production Ready | **Tech:** Python, Pandas, XlsxWriter | **Type:** ETL Automation

![Check Before After Demo]

## 💼 The Business Challenge
A mid-sized marketing agency was spending **4+ hours weekly** manually aggregating sales logs. Their raw data contained:
* **Inconsistent Formatting:** Currency mixed with text (`$1,000` vs `1000 USD`).
* **Duplication:** System glitches caused repeat transaction entries.
* **Dirty Categorization:** Regions entered as "north", "North", and "  North ".

This prevented them from uploading data to Salesforce and generating accurate revenue reports.

## 🛠️ The Solution
I engineered a **Python-based ETL (Extract, Transform, Load) Pipeline** that automates the entire reporting workflow.

**Key Features:**
* **Smart Cleaning:** Uses Regex to strip non-numeric characters from financial data.
* **Normalization:** Automatically standardizes text fields (Title Case, Trim) for consistent database entry.
* **Logic-Based De-duplication:** Identifies duplicates based on composite keys (ID + Date).
* **Automated Reporting:** Generates a **Management Pivot Table** showing "Revenue by Region" automatically.
* **Professional Styling:** Uses `XlsxWriter` to apply corporate formatting (Brand colors, Currency format) programmatically.

## 🚀 Performance
* **Manual Time:** 4 Hours/Week
* **Automation Time:** < 1 Second
* **Accuracy:** 100%

## 💻 How to Run
```bash
# 1. Install dependencies
pip install pandas xlsxwriter openpyxl

# 2. Run the pipeline
python main.py
