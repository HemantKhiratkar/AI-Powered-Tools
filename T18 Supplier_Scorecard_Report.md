# T18: Supplier_Scorecard_Report

# Task T18: Procurement Supplier Scorecard – Final Management Report

Prepared by: Hemant Khiratkar
Program: MBA Data Science & Data Analytics, SCIT
PRN:## 26030242022

## 1. Executive Summary

This report details the development and implementation of an automated Supplier Scorecard application built to evaluate vendor performance across a portfolio of 40 suppliers and 600 purchase orders. The resulting Streamlit application provides procurement managers with a dynamic, data-driven dashboard to adjust Key Performance Indicator (KPI) weights, classify suppliers into strategic tiers, and automatically generate corrective feedback for underperforming vendors.

## 2. Technical Architecture & Verification

The application was built as a monolithic Python script utilizing `pandas` for data manipulation, `streamlit` for the interactive user interface, and `openpyxl` for exporting auditable reports.

### Data Processing & Scoring Logic

- **KPI Aggregation:** Order-level data was rolled up to the supplier level to calculate four core metrics: On-Time Delivery Rate, Rejection Rate, Average Price Change, and Responsiveness.
- **Normalization:** Higher-is-better metrics (On-Time Delivery, Responsiveness) were linearly scaled from 0 to 100. Lower-is-better metrics (Rejection Rate, Price Change) were inverted.
- **Dynamic Weighting:** A largest-remainder rounding algorithm was implemented in the user interface, allowing managers to adjust KPI weights via sliders while strictly enforcing a 100% total sum.
- **Verification:** To ensure correctness, an automated baseline extraction script (`finalize_submission.py`) was engineered to recalculate all unweighted aggregations and write them to a styled `Baseline Metrics` sheet in the final Excel workbook. Manual spot-checks confirmed zero variance between the application’s internal dataframe and the raw synthetic data.

## 3. Business Insights & Procurement Recommendations

Based on the default weighted scoring model (On-Time Delivery: 25%, Rejection Rate: 25%, Price Change: 10%, Responsiveness: 40%), the portfolio exhibits an overall healthy supply chain with an average on-time delivery rate of 93.5% and a mean rejection rate of 6.03%. However, performance variance necessitates targeted procurement strategies.

### Strategic Tier Strategies

1. **Strategic & Maintain Tiers (Scores >= 65):**
    - **Recommendation:** Consolidate future PO volume toward these top-tier vendors. Lock in long-term pricing contracts with suppliers in the “Strategic” band to hedge against future price volatility, leveraging their proven reliability.
2. **Develop Tier (Scores 50 – 64.9):**
    - **Recommendation:** Implement mandatory 30-day corrective action plans. The application explicitly flags suppliers falling into this tier and generates rule-based recovery requests. Procurement buyers must mandate root-cause analysis for any supplier crossing the 5% rejection threshold before awarding new contracts.
3. **Exit Tier (Scores < 50):**
    - **Recommendation:** Initiate immediate offboarding protocols. Buyers should source secondary vendors for categories currently serviced by Exit-tier suppliers to mitigate supply chain disruption risks.

## 4. Model Limitations & Responsible AI

In compliance with responsible AI development principles, no real-world personal or confidential corporate data was utilized; all inputs were sourced from the provided synthetic teaching dataset.

### System Limitations

- **Relative Price Normalization:** The algorithm scores price changes relatively across the dataset (where the highest price increase maps to 0 and the lowest to 100). If a new outlier supplier is introduced with an extreme price hike, it will abruptly skew the scores of all other suppliers. In a production environment, price changes should be benchmarked against absolute inflation targets rather than dataset minimums/maximums.
- **Volume Agnosticism:** The current weighting model treats all purchase orders equally regardless of total spend. A late delivery on a highly critical, high-value component impacts the score exactly the same as a late delivery on low-value packaging materials.

## 5. Personal Reflection

Transitioning from writing rigorous ETL testing frameworks and data pipeline validations to utilizing AI for rapid application prototyping requires a fundamental shift in workflow. In traditional data quality engineering, the focus is on deterministic validation and Medallion architecture rigidity. With LLM-assisted development, the challenge shifts toward steering the model, managing context windows, and strictly verifying generative outputs.

The 15-prompt iteration process highlighted the necessity of breaking complex logic into modular steps. For example, asking the AI to immediately merge tables and calculate weighted scores resulted in duplicated PO rows. By steering the LLM to first execute a `groupby` aggregation before the left join, the data integrity was preserved.

Furthermore, leveraging AI to construct the Streamlit UI and the largest-remainder slider logic dramatically reduced frontend development time, allowing for a stronger focus on the underlying mathematical transformations and business logic. The experience reinforced that LLMs are exceptional accelerators for code generation, provided the human-in-the-loop maintains strict oversight over mathematical accuracy and architectural design.