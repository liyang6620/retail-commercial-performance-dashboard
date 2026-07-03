# Retail Commercial Performance Dashboard (Excel)

## Executive Summary

This project presents an interactive commercial performance dashboard developed in Microsoft Excel to support executive reporting and business decision-making in a retail environment.

The dashboard consolidates sales, profitability, regional performance, discount effectiveness and margin monitoring into a single reporting interface. By combining dynamic KPI cards, PivotTables, PivotCharts, Power Query and interactive slicers, the solution enables management to quickly monitor business performance and identify commercial improvement opportunities.

The project simulates the type of executive reporting commonly used by Commercial Analysts, Business Intelligence Analysts and Category Managers within retail organisations.

---

# Business Objectives

The dashboard was designed to answer the following business questions:

- How is the business performing in terms of sales and profitability?
- Which product categories generate the highest commercial value?
- Which regions contribute the strongest business performance?
- Are current discount strategies improving profitability?
- Which products generate the majority of revenue?
- How many transactions fail to achieve the expected profit margin?
- How does business performance vary across different years and regions?

---

# Executive Dashboard

The dashboard provides an executive overview through eight dynamic KPI indicators.

| KPI | Business Purpose |
|------|------------------|
| Total Sales | Monitor overall business revenue |
| Total Profit | Evaluate business profitability |
| Profit Margin | Measure overall operational efficiency |
| Total Orders | Monitor business activity |
| Average Discount | Evaluate pricing strategy |
| Average Order Value | Measure customer purchasing behaviour |
| Below Target Orders | Identify low-margin transactions |
| Below Target Sales | Quantify revenue generated below target margin |

All KPI indicators dynamically update according to Region and Year slicer selections.

---

# Business Analysis

## Sales & Profit Trend Analysis

Monthly sales and profit trends enable management to monitor business performance over time and identify seasonal demand patterns.

Business Value

- Monitor monthly revenue growth
- Identify seasonal fluctuations
- Compare profitability across different periods
- Detect unusual business performance

---

## Product Category Performance

Sales and profit are analysed across product categories to identify the strongest commercial contributors.

Business Value

- Compare category performance
- Identify high-value product groups
- Support merchandising decisions
- Optimise category management

---

## Regional Performance

Regional analysis compares commercial performance across different geographical markets.

Business Value

- Evaluate regional profitability
- Identify high-performing regions
- Support regional resource allocation
- Improve sales planning

---

## Discount Effectiveness

Commercial performance is analysed under different discount levels to evaluate pricing effectiveness.

Business Value

- Measure promotional effectiveness
- Compare sales and profit across discount bands
- Support pricing strategy optimisation
- Protect business profitability

---

## Pareto Analysis

Pareto Analysis applies the 80/20 principle to identify products contributing the largest share of total sales.

Business Value

- Identify key revenue-driving products
- Support inventory prioritisation
- Improve product assortment decisions
- Focus commercial resources on high-value products

---

## Margin Performance

Transactions are classified according to margin performance to monitor pricing quality.

Business Value

- Identify low-margin transactions
- Quantify commercial risk
- Monitor pricing effectiveness
- Support margin improvement initiatives

---

# Key Business Findings

### Finding 1 — Regional Performance Difference

The West region achieved the strongest commercial performance, generating approximately **$4.22M** in sales and **$1.08M** in profit while maintaining a profit margin of **25.58%**.

By comparison, the South region generated approximately **$2.40M** in sales with a lower profit margin of **24.76%**.

**Recommendation**

Investigate pricing strategies, product mix and promotional activities implemented in the West region and evaluate whether similar approaches can be replicated in lower-performing regions.

---

### Finding 2 — Discount Strategy

The dashboard shows that **Medium Discount** generated both the highest sales and the highest profit among all discount bands.

This indicates that moderate discounting achieved a better balance between revenue generation and profitability than aggressive discount strategies.

**Recommendation**

Prioritise moderate promotional campaigns and regularly evaluate the return on investment of high-discount activities.

---

### Finding 3 — Margin Risk

The dashboard identified **5,062 out of 9,994 orders (50.65%)** as falling below the target profit margin.

These transactions generated approximately **$7.57M** in sales.

**Recommendation**

Review pricing policies for products consistently generating low margins and strengthen discount approval controls to improve overall profitability.

---

### Finding 4 — Product Concentration

Pareto Analysis indicates that a relatively small number of product sub-categories contribute the majority of overall sales.

This concentration suggests that business performance depends heavily on a limited number of high-performing products.

**Recommendation**

Prioritise inventory availability, promotional investment and supply planning for high-contributing products while reviewing the commercial value of lower-performing items.

---

### Finding 5 — Profit Margin Stability Across Regions

Despite noticeable differences in regional sales performance, profit margins remained relatively stable across regions. For example, the West region achieved a profit margin of **25.58%**, while the South region recorded **24.76%**, both remaining close to the overall business average of **25.05%**.

This indicates that the current pricing framework is generally consistent across regions. However, the high proportion of below-target transactions suggests that profitability challenges are more likely driven by product-level pricing and discount decisions rather than regional performance alone.

**Recommendation**

Maintain the existing regional pricing strategy while prioritising the review of low-margin products and discount policies. Improving transaction-level margin performance is likely to deliver greater profitability gains than implementing broad regional pricing changes.
---

# Technical Implementation

## Data Preparation

The raw dataset was transformed using Power Query.

Data preparation included:

- Data cleaning
- Data type validation
- Date transformation
- Profit Margin calculation
- Target Margin lookup
- Margin Review classification
- Discount Band grouping
- Category grouping

---

## Dashboard Development

The dashboard was developed entirely using native Microsoft Excel functionality without VBA.

Technologies used include:

- Power Query
- Pivot Tables
- Pivot Charts
- Pivot Cache Connections
- Slicers
- GETPIVOTDATA
- SUMIFS
- COUNTIF
- XLOOKUP
- IF
- Conditional Formatting
- Custom Number Formatting

---

# Business Value

Compared with traditional spreadsheet reports, this dashboard provides:

- Interactive executive reporting
- Real-time KPI monitoring
- Dynamic commercial analysis
- Improved visibility of pricing and profitability
- Faster business decision-making
- Better identification of commercial risks

The project demonstrates how Microsoft Excel can be used to build an executive-level business intelligence solution without requiring VBA or external BI software.

---

# Skills Demonstrated

### Business Intelligence

- Executive Dashboard Design
- KPI Development
- Commercial Analytics
- Interactive Business Reporting

### Microsoft Excel

- Power Query
- Pivot Tables
- Pivot Charts
- Dynamic Slicers
- GETPIVOTDATA
- Conditional Formatting
- Advanced Excel Functions

### Business Analysis

- Sales Performance Analysis
- Profitability Analysis
- Category Analysis
- Regional Analysis
- Discount Analysis
- Pareto Analysis
- Executive Reporting

---

# Dataset

This project uses the publicly available **Supermart Grocery Sales – Retail Analytics Dataset** from Kaggle.

Dataset:

https://www.kaggle.com/datasets/mohamedharris/supermart-grocery-sales-retail-analytics-dataset

The dataset contains retail transaction records including:

- Order Date
- Region
- Category
- Sub Category
- Sales
- Profit
- Discount
- Customer Information

---

# Repository Structure

```
excel-retail-commercial-dashboard
│
├── Dashboard.xlsx
└── README.md
```

---

# Future Enhancements

Potential future improvements include:

- Power BI implementation
- SQL database integration
- Automated data refresh
- Sales forecasting
- Customer segmentation
- Product profitability analysis
- Drill-through reporting

---

# Acknowledgements

This project uses the publicly available **Supermart Grocery Sales – Retail Analytics Dataset** from Kaggle for educational and portfolio purposes.

Dataset:

https://www.kaggle.com/datasets/mohamedharris/supermart-grocery-sales-retail-analytics-dataset

Special thanks to the dataset author for making the dataset publicly available to the data analytics community.

---

# Author

**Yang Li**

Master of Information Technology  
University of Auckland

Business Intelligence | Data Analytics | Commercial Analytics
