# Fraud-Detection-and-Reporting-Excel-
🛡️ Excel Fraud Detection Dashboard 
This project showcases a complete fraud analysis system built entirely in Excel. From data cleaning to a dynamic dashboard with slicers and automation, this file demonstrates how Excel can be transformed into a mini BI tool — without Power BI or external tools.

📁 Files Included
File	Description
Fraud_Analysis_Clean.xlsx	Processed and cleaned data with dashboards and visuals (no macros).
Fraud_Analysis_Dashboard.xlsm	Macro-enabled dashboard with automation (click-to-generate summary).

✨ Features
✅ Data Cleaning & Preparation
Removed blank, duplicate, and invalid entries.

Applied:

PROPER() for use_chip classification

UPPER() for Gender


🧩 Feature Engineering
Created new columns to enrich data analysis:

Age (from DOB)

Age Group (buckets like 18–25, 26–35...)

Income Category (Low, Middle, High)

Is_Fraud (binary fraud flag)

Credit Category (based on credit score)

Chip Classification (Chip-Based / Non-Chip-Based)

Time Category (Morning / Afternoon / Evening / Late Night)

📊 Interactive Dashboard (No Coding Required)
🚀 KPIs
Total Transactions

Total Amount

Fraud Transactions

Fraud Amount

Detection Rate

📈 Visualizations
Fraud vs Non-Fraud

Chip-Based Fraud Share

Fraud by Gender

Monthly Fraud Trend

Fraud by Income

Client-wise Fraud Cases

Fraud by Credit Score

Fraud by Error Type

Fraud by States

Card-based Fraud Trend

Fraud by Time of Day

🧮 Slicers
Year

Gender

🧠 VBA Macro – Auto Fraud Summary
A macro button titled "Fraud Summary (Monthly)" auto-generates a Pivot Table showing:

| Year | Month | Sum of Amount | Count of Fraud Transactions |

No need to manually insert filters or pivot tables – click and done ✅

# ⚙️ How to Use
# For Fraud_Analysis_Clean.xlsx
Open normally to explore cleaned data and dashboard.

No macros are involved.

# For Fraud_Analysis_Dashboard.xlsm
Right-click the file → Properties

Tick Unblock (if shown) and click OK

Open the file in Excel

Click Enable Content when prompted

Use the macro button to generate monthly summaries

Play with slicers to filter dashboard by year or gender

📌 Requires macros to be enabled for full functionality

🛠 Tools Used
Microsoft Excel

Data Cleaning

Pivot Tables

Conditional Columns

Dashboard Design

VBA (Visual Basic for Applications)

Automated summary generation using button click


🙋 About
This project demonstrates how Excel can evolve from a spreadsheet tool to a mini data product. From data wrangling to visualization to automation — all in one .xlsm.

Made with 💻 + 📊 + ❤️ by a data enthusiast.
