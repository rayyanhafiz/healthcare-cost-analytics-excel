# healthcare-cost-analytics-excel
End-to-end healthcare insurance and hospitalization cost analysis in Microsoft Excel featuring data cleaning, multi-sheet XLOOKUP integration, risk segmentation, pivot tables, and an interactive slicer-driven dashboard across 2,335 patient records.
# Healthcare Cost & Risk Analytics Dashboard (Microsoft Excel)

An end-to-end exploratory data analysis (EDA) and business intelligence project built in Microsoft Excel. This project cleans, transforms, and consolidates 2,335 patient records across demographic, clinical, and hospitalization datasets to uncover key cost drivers in healthcare insurance charges.

---

## 📋 Table of Contents
- [Project Overview](#project-overview)
- [Dataset Architecture](#dataset-architecture)
- [Data Cleaning & Transformation Workflow](#data-cleaning--transformation-workflow)
- [Key Analytical Findings](#key-analytical-findings)
- [Dashboard & Visualizations](#dashboard--visualizations)
- [Repository Structure](#repository-structure)
- [How to Use](#how-to-use)

---

## 🔍 Project Overview
Healthcare insurers and hospital networks face significant variance in patient medical charges. The objective of this project is to:
1. Audit and standardize raw relational tables (`Customer Names`, `Medical Examinations`, and `Hospitalisation Details`).
2. Engineer clinical risk categories (BMI Weight Status and HbA1c Diabetes Status) using Excel logical formulas.
3. Merge all cleaned records into a unified analytical master table (`Healthcare`).
4. Build Pivot Tables, statistical scatter plots, and an interactive slicer-driven dashboard to evaluate how lifestyle factors, chronic conditions, and facility tiers impact medical costs.

---

## 🗂️ Dataset Architecture

The workbook (`healthcare.xlsx`) consists of **2,335 unique patient records** (`Id1` through `Id2335`) spanning **17 regional State IDs**, organized across raw, master, and presentation tabs:

| Worksheet | Type | Description |
| :--- | :--- | :--- |
| **`Customer Names`** | Raw / Transformed | Patient IDs alongside parsed `Title`, `First Name`, and `Last Name` fields. |
| **`Medical Examinations`** | Raw / Transformed | Clinical metrics (`BMI`, `HBA1C`, `Heart Issues`, `Any Transplants`, `Cancer history`, `NumberOfMajorSurgeries`, `smoker`) and derived risk classifications. |
| **`Hospitalisation Details`** | Raw / Transformed | Admission dates, calculated `Date of Birth`, `Age`, `children`, `charges`, `Hospital tier`, `City tier`, and `State ID`. |
| **`Healthcare`** | Master Dataset | Consolidated 17-column analytical table joining all three source sheets via `Customer ID`. |
| **`Pie Chart`** | Pivot Analysis | Customer distribution by Smoking & Cancer history; total major surgeries by Transplant status. |
| **`Column and Bar`** | Pivot Analysis | Average hospitalization charges segmented by Weight/Diabetes status and State ID/Hospital Tier. |
| **`Scatter Chart 1–3`** | Chartsheets | Correlation views comparing patient charges against Age, BMI, and HbA1c levels. |
| **`Dashboard`** | Executive View | Interactive visual dashboard equipped with dynamic slicers for real-time filtering. |
| **`ANS`** | Audit Log | Step-by-step documentation of data quality checks, cell-level imputations, and transformations. |

---

## 🧹 Data Cleaning & Transformation Workflow

All cleaning steps are documented in the `ANS` tab and visually highlighted in the source sheets:

1. **Missing & Invalid Value Detection:**
   - Audited placeholder `"?"` values across all columns using `COUNTIF` formulas.
   - Standardized missing `State ID` entries by imputing `"Unknown"`.
   - Resolved missing or inconsistent `year`, `month`, `smoker`, and `NumberOfMajorSurgeries` entries using statistical imputation (e.g., averaged year values) and consistency rules.
2. **String Parsing & Text Standardization:**
   - Split full customer names (`Last, Title First`) into discrete `Title`, `First Name`, and `Last Name` columns.
3. **Feature Engineering (Clinical Risk Stratification):**
   - **Weight Status (`IF` logic on BMI):** Categorized patients into *Underweight*, *Normal Weight*, *Overweight*, and *Obesity* (mean BMI: `30.97`).
   - **Diabetes Status (`IF` logic on HbA1c):** Classified glycemic levels into *Normal*, *Prediabetes*, and *Diabetes* (mean HbA1c: `6.58`).
4. **Date & Financial Formatting:**
   - Constructed standardized `Date of Birth` fields (`DD-MMM-YYYY`) and dynamically calculated patient `Age` (ranging from `18` to `65`, mean `39.2` years) using `INT`, `YEARFRAC`, and `DATE` functions.
   - Formatted `Charges` into standard currency format (`$`).
5. **Data Consolidation:**
   - Unified all 17 attributes into the master `Healthcare` sheet using lookup functions (`XLOOKUP` / `VLOOKUP`) keyed on `Customer ID`.

---

## 📊 Key Analytical Findings

Across **$31.59M** in total cumulative hospitalization charges (overall mean: **$13,529.92**; median: **$9,630.91**; max: **$63,770.43**):

- **Facility Tier Impact:** Hospital tier is one of the strongest structural cost drivers. **Tier 1** hospitals average **$30,129.20** per patient—more than **2.5×** the average cost of **Tier 2** (**$11,865.27**) and **3.1×** that of **Tier 3** (**$9,462.27**) facilities.
- **Compounding Clinical Risks (BMI × Diabetes):**
  - Patients with both **Obesity and Diabetes** incur the highest average charges at **$19,103.40**, compared to **$7,897.08** for **Normal Weight, Non-Diabetic** patients.
  - Across all weight categories, diabetic patients average **$15,973.54** vs. **$12,090.72** for patients with normal HbA1c levels.
- **Lifestyle & Comorbidity Breakdown:**
  - **21.0%** of the cohort (`490` of `2,335` customers) are smokers, with `84` smokers also reporting a history of cancer.
  - Patients have undergone a cumulative **1,579 major surgeries** (`1,417` among non-transplant patients and `162` among transplant recipients).

---

## 📈 Dashboard & Visualizations

The **`Dashboard`** worksheet provides an interactive overview powered by Pivot Charts and Slicers:
- **KPI Summary & Breakdowns:** Customer segmentation by smoking status, cancer history, and major surgical volume.
- **Multi-Dimensional Cost Comparisons:** Clustered column and bar charts comparing average charges across Weight Status, Diabetes Status, Hospital Tier, and State ID.
- **Continuous Variable Correlations (`Scatter Chart 1–3`):** Multi-series scatter plots illustrating cost clustering across Age, BMI, and HbA1c thresholds.

---

## 📁 Repository Structure

```text
├── healthcare.xlsx      # Complete Excel workbook (Raw data, Master table, Pivot tables, Charts & Dashboard)
└── README.md            # Project documentation, methodology, and analytical summary
