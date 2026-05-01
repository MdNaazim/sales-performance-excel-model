# 📊 Sales Performance & Market Sizing — Excel Model

> A dynamic Excel-based business analytics model tracking 500+ SKU-level sales across 6 regions, uncovering a ₹12 Lakh revenue opportunity through market gap analysis — with findings packaged in an executive-ready PowerPoint deck.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-Advanced-green?style=flat-square&logo=microsoftexcel)
![PowerPoint](https://img.shields.io/badge/PowerPoint-Executive%20Report-orange?style=flat-square&logo=microsoftpowerpoint)
![Business Analysis](https://img.shields.io/badge/Business-Market%20Sizing-blue?style=flat-square)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=flat-square)

---

## 📌 Project Overview

Retail businesses with hundreds of products across multiple regions often lack a clear view of where revenue is being lost. This project builds a **dynamic Excel analytics model** that:

- Tracks **500+ SKU-level sales** across **6 regions** for a simulated retail client
- Uses **Pivot Tables, VLOOKUP/XLOOKUP, and conditional formatting** for interactive analysis
- Performs **market gap analysis** to identify under-served regions and underperforming products
- Quantifies a **₹12 Lakh revenue opportunity** that the business can act on immediately
- Delivers findings in a **structured executive PowerPoint report** with visualizations and recommendations

**Business impact:** Decision-makers can instantly see which products are underperforming in which regions, and how much revenue is being left on the table — without needing any coding or BI tools.

---

## 🗂️ Repository Structure

```
sales-performance-excel-model/
│
├── data/
│   └── sales_data_raw.xlsx          # Raw sales dataset (500+ SKUs × 6 regions)
│
├── model/
│   └── sales_analysis_model.xlsx    # Final dynamic Excel model with all analysis
│
├── presentation/
│   └── executive_report.pptx        # Executive-ready slide deck with insights
│
├── screenshots/
│   ├── pivot_table_view.png         # Pivot table screenshot
│   ├── regional_heatmap.png         # Regional performance heatmap
│   ├── gap_analysis.png             # Market gap analysis chart
│   └── dashboard_overview.png       # Full Excel dashboard view
│
└── README.md
```

---

## 🔍 Problem Statement

| | |
|---|---|
| **Domain** | Retail Sales / Business Analytics |
| **Problem** | Identify underperforming SKUs and revenue gaps across regions |
| **Data** | 500+ SKU-level transactions across 6 regions |
| **Tools** | MS Excel (Pivot Tables, VLOOKUP, XLOOKUP, Conditional Formatting) |
| **Output** | Dynamic Excel model + Executive PowerPoint report |
| **Key Finding** | ₹12 Lakh revenue opportunity identified through gap analysis |

---

## 📋 Dataset Description

The dataset is a simulated retail sales dataset with the following structure:

| Column | Description |
|---|---|
| `SKU_ID` | Unique product identifier |
| `Product_Name` | Name of the product |
| `Category` | Product category (Electronics, Apparel, FMCG, etc.) |
| `Region` | One of 6 sales regions (North, South, East, West, Central, Northeast) |
| `Units_Sold` | Number of units sold in the period |
| `Unit_Price` | Selling price per unit (₹) |
| `Revenue` | Calculated: Units Sold × Unit Price |
| `Target_Revenue` | Expected revenue for that SKU in that region |
| `Revenue_Gap` | Calculated: Target − Actual Revenue |
| `Approver` | Sales manager responsible for the region |

---

## 🛠️ Tools & Techniques Used

| Tool / Feature | Purpose |
|---|---|
| **Pivot Tables** | Summarize revenue by region, category, and SKU |
| **VLOOKUP / XLOOKUP** | Link SKU master data to sales records |
| **Conditional Formatting** | Highlight underperforming SKUs in red, top performers in green |
| **Slicers** | Interactive filters for region, category, and time period |
| **Excel Charts** | Bar charts, waterfall chart, pie chart by region |
| **Data Validation** | Dropdown lists for region and category filters |
| **SUM / SUMIF / AVERAGEIF** | Aggregate metrics by region and category |
| **PowerPoint** | Executive report with charts and recommendations |

---

## ⚙️ How to Use This Model

### Requirements

- Microsoft Excel 2016 or later (or Microsoft 365)
- Microsoft PowerPoint (to view the executive report)
- No Python or coding knowledge required

### Steps

**1. Download the files**

Click the green **Code** button → **Download ZIP** → Extract the folder.

**2. Open the raw dataset**

Open `data/sales_data_raw.xlsx` to explore the original 500+ row dataset.

**3. Open the analysis model**

Open `model/sales_analysis_model.xlsx`. The workbook has the following sheets:

| Sheet Name | Contents |
|---|---|
| `Raw Data` | Original cleaned dataset |
| `SKU Master` | Product reference table (linked via XLOOKUP) |
| `Pivot Analysis` | 3 pivot tables with slicers |
| `Gap Analysis` | Revenue gap calculations per region and SKU |
| `Dashboard` | Summary charts and KPI cards |
| `Recommendations` | Written insights based on findings |

**4. Interact with the model**

- Use the **slicers** on the Pivot Analysis sheet to filter by Region, Category, or Month
- The Dashboard sheet updates automatically as you change filters
- Red cells = underperforming (below target by > 20%), Green = on track or above

**5. View the executive report**

Open `presentation/executive_report.pptx` for the 6-slide summary deck.

---

## 🔬 Methodology

### Step 1 — Data Preparation

- Created the raw sales dataset with 500+ SKU rows across 6 regions
- Cleaned inconsistent region names (e.g. "north" → "North"), removed duplicates
- Added calculated columns:
  - `Revenue = Units_Sold × Unit_Price`
  - `Revenue_Gap = Target_Revenue − Revenue`
  - `Gap_Percentage = Revenue_Gap / Target_Revenue × 100`

### Step 2 — Pivot Table Analysis

Built **3 core pivot tables:**

1. **Revenue by Region** — Which region contributes the most / least revenue
2. **Top 20 SKUs by Revenue** — Best and worst performing products overall
3. **Region × Category Matrix** — Cross-tab of revenue by region and product category

Added **slicers** connected to all 3 pivot tables for interactive filtering.

### Step 3 — XLOOKUP Integration

- Created a `SKU Master` reference sheet with: SKU ID, Product Name, Category, Unit Price, Margin %
- Used `XLOOKUP(SKU_ID, master_range, return_range)` to pull category and margin data into the sales sheet
- Applied `SUMIF` formulas to calculate total revenue and gap per category per region

### Step 4 — Market Gap Analysis

- Calculated `Revenue_Gap` for each SKU–Region combination
- Summed gaps by region to find total missed revenue per region
- **Top 3 under-served regions:** Region 4, Region 5, Region 6
- **Total gap across all regions = ₹12,00,000 (₹12 Lakhs)**
- Applied conditional formatting rules:
  - Gap > 20%: Red fill — critical underperformance
  - Gap 10–20%: Amber fill — needs attention
  - Gap < 10%: Green fill — on track

### Step 5 — Dashboard & Reporting

Built the following visuals on the Dashboard sheet:
- **KPI cards:** Total Revenue, Total Target, Overall Gap %, Top Region, Bottom Region
- **Bar chart:** Revenue vs Target by Region (side-by-side)
- **Waterfall chart:** Revenue gap breakdown by region
- **Pie chart:** Revenue share by product category
- **Top 10 SKUs table:** Sorted by gap amount descending

### Step 6 — Executive PowerPoint Deck

| Slide | Content |
|---|---|
| 1 | Title + Project Scope |
| 2 | Summary KPIs — Total Revenue, Gap, Regions Analyzed |
| 3 | Regional Performance Heatmap |
| 4 | Top Underperforming SKUs |
| 5 | Market Gap Analysis — ₹12L opportunity breakdown |
| 6 | Recommendations for business action |

---

## 📈 Results & Key Findings

| Metric | Value |
|---|---|
| Total SKUs Analyzed | 500+ |
| Regions Covered | 6 |
| Total Revenue Tracked | ₹68.4 Lakhs (simulated) |
| Total Revenue Gap Found | ₹12 Lakhs (17.5% of target) |
| Top Performing Region | Region 2 (North) |
| Most Under-served Region | Region 5 (Northeast) |
| Highest Gap SKU Category | Electronics — ₹4.2L gap |
| Pivot Tables Built | 3 |
| Charts Created | 5 |
| PowerPoint Slides | 6 |

### Key Business Recommendations

1. **Region 5 (Northeast)** has the largest revenue gap (₹3.8L) — investigate distribution and pricing strategy
2. **Electronics category** is the highest-value gap segment — prioritize stock availability and promotional campaigns
3. **Top 20 underperforming SKUs** account for 68% of the total ₹12L gap — focus interventions here first
4. **Regions 4 and 6** are consistently below target across all categories — consider assigning dedicated sales resources

---

## 🖼️ Screenshots

### Dashboard Overview
> *(Add screenshot: `screenshots/dashboard_overview.png`)*

### Regional Performance Heatmap
> *(Add screenshot: `screenshots/regional_heatmap.png`)*

### Gap Analysis Chart
> *(Add screenshot: `screenshots/gap_analysis.png`)*

---

## 👤 Author

**Mohammed Naazim Pasha Z**
Business Analyst | Data Analytics | Customer Strategy

- 📧 mohammednaazim77@gmail.com
- 💼 [LinkedIn](https://linkedin.com/in/mohammed-naazim-pasha)
- 🐙 [GitHub](https://github.com/MdNaazim)

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

## 🙏 Acknowledgements

- Business analysis methodology inspired by Deloitte Business Data Analysis Simulation (Forage)
- Market sizing framework based on standard retail analytics best practices

---

> ⭐ If this project helped you understand Excel-based business analytics, please give it a star!
