# Sales Performance Automation Dashboard (Excel + VBA)

This project is a comprehensive Excel-based dashboard designed to analyze key metrics related to sales performance, customer behavior, product profitability, and operational efficiency. It uses advanced Excel features including Power Query, Pivot Tables, DAX Measures, VBA Macros, and custom visuals.


## 🧩 Key Features

- **Interactive Filter Panel with VBA**: Users can show/hide slicers dynamically to enhance usability and save screen space.
- **Customer Analysis**: Insights by gender and age group with waffle charts and custom metrics.
- **Product Performance**: Return/refund rate breakdown, most profitable products, and high-selling categories.
- **Profitability Over Time**: Month-over-month trends, weekday patterns, and seasonal sales fluctuations.
- **Advanced Formatting**: Custom value formats (`+0.0%;-0.0%`), conditional formatting, gradients, and Zebra BI tables.
- **Power Query Editor**:
  - Removed duplicates and nulls
  - Added new calculated columns
  - Created a custom Date Table
- **Form Controls & Navigation**: Smooth tab switching with VBA and shape-based interactivity.
- **DAX and Excel Measures**: Centralized measure table ("Calculations") to perform all required KPI calculations.
- **Used Functions**:
  - `SEQUENCE`, `TEXTJOIN`, `IF`, `SWITCH`, `CALCULATE`, `FILTER`, etc.

## 📌 Requirements Answered
The dashboard addresses all business queries regarding:
- Monthly and weekday profitability
- Store performance vs targets
- High revenue products and cities
- Budget vs actual variance
- Customer segment insights

## 📐 Measures using DAX
- Created a separate `Calculations` table to centralize KPIs
- Profitability, COGS, Return Rate, Refund Rate, MoM Growth

 ## 📋 Slicers & Interactivity
- Month and Category filters with **custom styling**
- VBA-powered **toggle button** to show/hide filter panel

## 📁 File Info
- File type: Excel Macro-Enabled Workbook (.xlsm)
Macros: Enabled and safe (used for UI automation and formatting)

## ⚙️ Tech Stack
- **Excel Power Query**
- **Pivot Tables**
- **DAX (basic to intermediate)**
- **VBA Macros**
- **Zebra BI (non-native charting)**
- **Form Controls**

## 📷 Dashboard Preview
**Profit View** 
![Image](https://github.com/user-attachments/assets/4860b4c2-97e0-417f-a195-5609b405bffb)

**Store**  
![Image](https://github.com/user-attachments/assets/dc0ff5a5-31b8-4526-a687-e25d43d6458c)

**TimeFrame**  
![Image](https://github.com/user-attachments/assets/1de64db8-fc83-4277-8535-efe6a97c3fb2)

## 💡 Learnings

- Used **TextJoin, Sequence**, and custom % formats
- Learned difference between **return and refund rate**
- Created VBA macro for dynamic UI interactivity
- Structured calculations for reuse and easier maintenance
---

## 📣 Future Work
Planning to enhance this with Power BI or Tableau version with real-time data sources and integration.




👤 **Created By:** Ritu Thakur  
📬 [LinkedIn Profile](https://linkedin.com/in/rituthakur-29)  
📫 ritut452@gmail.com
