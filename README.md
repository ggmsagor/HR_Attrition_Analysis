# Employee Attrition Analysis
### Tools Used: Microsoft Excel · Microsoft Power BI

---

<img width="1899" height="771" alt="image" src="https://github.com/user-attachments/assets/42ddea3a-63b4-4f50-958d-5bef152c6250" />
<img width="1905" height="687" alt="image" src="https://github.com/user-attachments/assets/b4199eae-8ec1-42f2-9b72-55bcefe736cd" />

---

## Project Overview

This project explores the factors driving employee attrition at a mid-to-large organisation using the IBM HR Analytics dataset. The goal was to move beyond a raw attrition rate and identify **which employee segments are most at risk, why they leave, and where the company should act first**.

The dataset contains records for **1,470 employees** across 35 variables — covering demographics, job satisfaction, compensation, work patterns, and career history.

All analysis was done entirely in **Microsoft Excel** (pivot tables, calculated fields, summary tables) and **Microsoft Power BI** (interactive dashboard). No coding or statistical software was used.

---

## Files in This Repository

| File | Description |
|------|-------------|
| `HR_Attrition_Analysis.xlsx` | Excel workbook with all summary tables, breakdowns, and analysis sheets |
| `HR_Attrition_Dashboard.pbix` | Power BI dashboard file (interactive) |
| `WA_Fn-UseC_-HR-Employee-Attrition.csv` | Raw dataset (source data) |
| `Data_Dictionary.md` | Field definitions and scale descriptions |
| `PowerBI_Build_Guide.md` | Step-by-step guide for recreating the dashboard |

---

## Key Findings

**Overall attrition rate: 16.1%** (237 of 1,470 employees left)

### 1. Job Role is the strongest predictor of attrition
Sales Representatives have a **39.8%** attrition rate — the highest across all roles. Laboratory Technicians (23.9%) and HR staff (23.1%) are next. In contrast, Research Directors leave at just 2.5%.

### 2. Overtime doubles the attrition rate
Employees who work overtime leave at **30.5%** vs **10.4%** for those who don't. This is the single largest binary split in the dataset.

### 3. Young employees leave at a disproportionately high rate
The 18–25 age group has a **34.8%** attrition rate. This drops to 9.2% for the 36–45 bracket. Early career experiences — compensation, growth opportunities, manager quality — appear to play a central role.

### 4. Compensation gap is significant
Employees who left earned an average of **$4,787/month** vs **$6,833** for those who stayed — a 30% gap. While some of this is explained by role and seniority, it points to compensation as a retention lever.

### 5. Single employees leave at 2.5× the rate of divorced employees
Single employees have a **25.5%** attrition rate vs 10.1% for divorced colleagues. This likely reflects different levels of financial commitments and mobility.

### 6. Stock options are a strong retention tool
Employees with no stock options leave at **24.4%**. Those at Level 1 drop to just **9.4%**. The relationship is not perfectly linear (Level 3 rises to 17.6%) but the signal is clear at the entry level.

### 7. Frequent business travel elevates risk
Employees who travel frequently leave at **24.9%** — three times the rate of non-travellers (8.0%). This compounds with job satisfaction and work-life balance scores.

### 8. Job satisfaction predicts attrition, but the gap is moderate
Low satisfaction employees (score 1/4) leave at **22.8%** vs **11.3%** for the most satisfied. Satisfaction alone is not the whole story — structural factors appear equally important.

---

## Analysis Structure (Excel Workbook)

**Sheet 1 – Executive Summary**
High-level KPIs and a summary of the eight key findings. Suitable for a one-page management read.

**Sheet 2 – By Department & Role**
Full breakdown of attrition counts and rates for all three departments and all nine job roles, with risk-level classifications.

**Sheet 3 – Workforce Demographics**
Six analysis tables covering age group, gender, marital status, business travel, overtime, and work-life balance.

**Sheet 4 – Job Factors**
Attrition breakdowns by job satisfaction, environment satisfaction, job involvement, stock option level, income comparison, and tenure.

**Sheet 5 – Raw Data (Sample)**
A 50-row sample of the original dataset for reference. The full CSV is provided separately.

---

## Power BI Dashboard Pages

1. **Overview** — Attrition rate KPIs, trend breakdown, headline filters
2. **Demographics** — Age, gender, marital status, business travel slicers
3. **Job & Compensation** — Role-level attrition, salary distribution, stock options
4. **Risk Heatmap** — Cross-tab of department × overtime × satisfaction

---

## Data Source

IBM HR Analytics Employee Attrition & Performance dataset, made publicly available on Kaggle. The data is fictional and was created by IBM data scientists for learning purposes.

**Dataset size:** 1,470 rows × 35 columns
**No missing values** in the original dataset.

---

## Methodology Notes

- Attrition rate is calculated as: `(Employees Left / Total Employees) × 100`
- Satisfaction scores (1–4) represent: 1 = Low, 2 = Medium, 3 = High, 4 = Very High
- Work-Life Balance scores (1–4): 1 = Bad, 2 = Good, 3 = Better, 4 = Best
- Job Involvement scores (1–4): 1 = Low, 2 = Medium, 3 = High, 4 = Very High
- Age groups were manually bucketed: 18–25, 26–35, 36–45, 46–60
- All figures are based on the full dataset of 1,470 records

---

## How to Use This Project

1. Open `HR_Attrition_Analysis.xlsx` in Excel — all summary tables are pre-built and formatted
2. Open `HR_Attrition_Dashboard.pbix` in Power BI Desktop — use slicers to explore segments interactively
3. To refresh or extend the analysis, load `WA_Fn-UseC_-HR-Employee-Attrition.csv` as the data source in both tools
4. Refer to `Data_Dictionary.md` for field definitions

---

*This project was completed as part of a data analytics portfolio. Feedback and questions are welcome via GitHub Issues.*
