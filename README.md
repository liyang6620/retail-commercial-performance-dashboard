# Retail Commercial Performance Dashboard
Microsoft Excel | Power Query | Commercial Analytics | Interactive Dashboard

## Executive Summary

Retail businesses generate large volumes of transactional data every day, making it difficult for managers to monitor sales performance, profitability and pricing effectiveness using traditional spreadsheet reports.

This project addresses that challenge by developing an interactive executive dashboard that consolidates key commercial metrics into a single reporting interface. This project presents an interactive commercial performance dashboard developed in Microsoft Excel to support data-driven commercial decision-making in a retail environment.

The dashboard consolidates sales, profitability, regional performance, discount effectiveness and margin monitoring into a single reporting interface. By combining dynamic KPI cards, PivotTables, PivotCharts, Power Query and interactive slicers, the solution enables management to quickly monitor business performance and identify commercial improvement opportunities.

The dashboard demonstrates how Microsoft Excel can be used to deliver an executive reporting solution comparable to entry-level Business Intelligence applications.

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

All KPI indicators are dynamically linked to PivotTables through GETPIVOTDATA and respond instantly to Region and Year slicer selections. The dashboard provides an executive overview through eight dynamic KPI indicators.

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

## Finding 1 — Regional Performance Difference

### Evidence

The West region achieved the strongest commercial performance, generating approximately **$4.22M** in sales and **$1.08M** in profit while maintaining a **25.58%** profit margin.

By comparison, the South region generated approximately **$2.40M** in sales with a lower profit margin of **24.76%**.

### Business Impact

The significant difference in sales and profitability suggests that commercial performance is influenced by regional market conditions, pricing strategies and product mix. Replicating successful business practices from high-performing regions could improve overall company performance.

### Recommendation

Investigate the pricing strategy, promotional activities and product assortment implemented in the West region, and evaluate whether similar commercial approaches can be adopted in lower-performing regions to improve profitability and revenue growth.

---

## Finding 2 — Discount Strategy Effectiveness

### Evidence

The dashboard shows that the **Medium Discount** band generated both the highest sales and the highest profit among all discount levels.

Higher discount levels did not deliver proportional improvements in profitability.

### Business Impact

The results indicate that moderate discounting provides the best balance between revenue generation and profit protection. Increasing discount depth beyond this level may stimulate sales volume but reduces overall commercial returns.

### Recommendation

Prioritise moderate promotional campaigns and regularly evaluate the return on investment (ROI) of discount activities before implementing aggressive pricing strategies. Future promotions should focus on maintaining profitability rather than maximising sales volume alone.

---

## Finding 3 — Margin Risk

### Evidence

The dashboard identified **5,062 out of 9,994 orders (50.65%)** as falling below the target profit margin.

These transactions generated approximately **$7.57M** in sales.

### Business Impact

Although overall sales performance remains strong, more than half of all transactions failed to achieve the expected profit margin. This indicates that pricing and discount decisions are limiting overall profitability and increasing commercial risk.

### Recommendation

Review pricing policies for products consistently generating low margins, strengthen discount approval processes and monitor below-target transactions more closely to improve transaction-level profitability.

---

## Finding 4 — Revenue Concentration in Key Product Categories

### Evidence

Pareto Analysis demonstrates that a relatively small number of product sub-categories contribute the majority of overall sales, while many remaining products contribute only a small proportion of total revenue.

### Business Impact

Business performance is highly dependent on a limited number of high-contributing products. Stock shortages, pricing issues or supply disruptions affecting these key products could have a disproportionate impact on total revenue and profitability.

### Recommendation

Prioritise inventory availability, supplier collaboration and promotional investment for high-contributing product categories. At the same time, review low-performing products to determine whether pricing, merchandising or product rationalisation could improve overall category performance and operational efficiency.

---

## Finding 5 — Profit Margin Stability Across Regions

### Evidence

Despite noticeable differences in regional sales performance, profit margins remained relatively consistent across regions.

For example:

- **West:** 25.58%
- **South:** 24.76%
- **Overall Business:** 25.05%

### Business Impact

The relatively stable regional margins suggest that the overall pricing framework is generally effective across geographical markets. However, the large proportion of below-target transactions indicates that profitability challenges are more likely driven by product-level pricing and discount decisions rather than regional pricing differences.

### Recommendation

Maintain the current regional pricing strategy while prioritising the optimisation of product pricing, discount policies and low-margin transactions. Improving product-level margin performance is likely to deliver greater profitability gains than implementing broad regional pricing changes.

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

Business Intelligence | Data Analytics | Commercial Analytics
