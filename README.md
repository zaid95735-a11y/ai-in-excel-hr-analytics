# AI in Excel - HR Analytics Dashboard | End-to-End Data Cleaning & Insights

> Built for HR Decision Making using Excel's Native AI Features - No Power BI Required

**Dataset:** 14k+ Employee Records | **Tools:** Excel 365 (Analyze Data AI, Flash Fill, Python in Excel) | **Location:** Mumbai, India

### Business Problem
The HR team had no visibility on department-wise headcount distribution and the impact of training on performance. The raw dataset had 4 major data quality issues - typos like `Marketir`, `region_26`, `Bach/Mast`, `m/f`.

### My Approach

**1. AI-Powered Data Cleaning (90% Faster)**
- **Problem:** Inconsistent values like `Sales & Marketir`, `Techn`, `region_26`, `Bach`
- **Solution:** Used `Flash Fill (Ctrl+E)` - Excel's AI pattern detection for auto-correction + `SUBSTITUTE()` formulas
- **Formulas Used:** `=SUBSTITUTE(B2,"Marketir","Marketing")`, `=SUBSTITUTE(C2,"_"," ")`
- **Impact:** Cleaned 14k rows in 5 minutes, saved 2 hours of manual effort

**2. AI Insights - Analyze Data Feature**
Used `Data > Analyze Data` - Excel's built-in AI Copilot generated automatic insights:
- Which department has maximum headcount? -> Technology (32%)
- Correlation between trainings and rating? -> Employees with rating 5 take avg 1.2 trainings
- Education impact? -> Masters employees have 0.8 higher avg rating

**3. Python in Excel for Advanced Analysis**
```python
=PY("import pandas as pd; df=xl('A1:H100'); df.groupby('department')['previous_year_rating'].mean().sort_values(ascending=False)")
