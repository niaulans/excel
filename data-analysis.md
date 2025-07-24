# 📊 Sales Data Analysis Workflow – Complete Version

| Stage | Step Name               | Goal                               | Activities (Step-by-Step) | Tools / Examples |
| ----- | ----------------------- | ---------------------------------- | ------------------------- | ---------------- |
| 1     | Understand the Business | Align analysis with business needs |

1. Identify the key business questions
2. Understand the sales goals (e.g. increase revenue, reduce churn)
3. Define KPIs (e.g. total sales, AOV, conversion rate)
4. Know the end users (e.g. sales manager, C-level) | Notion, Docs, Business Brief |
   | 2 | Data Collection | Gather all relevant data sources |
5. Collect sales transactions
6. Include product, customer, and promotion data
7. Pull from database, CSV, API, or Excel
8. Ensure data ranges match the analysis period | SQL, Excel, Python (pandas), API |
   | 3 | Data Cleaning | Prepare clean, reliable, and analysis-ready data |
9. Remove duplicate records (e.g. based on transaction ID)
10. Handle missing values:  
      - Drop irrelevant columns  
      - Fill using mean/median/mode  
      - Forward-fill or backward-fill for time series  
      - Flag incomplete rows
11. Fix inconsistent formats:  
      - Standardize dates (`YYYY-MM-DD`)  
      - Ensure numeric columns are correct type
12. Normalize text data:  
      - Convert to lowercase  
      - Trim whitespace  
      - Remove special characters if needed
13. Correct categorical values (fix typos, merge variations)
14. Check and handle outliers (using IQR, z-score)
15. Validate logical relationships:  
      - `total_price = qty × unit_price`  
      - Ensure no future transaction dates
16. Create new features:  
      - `total_price`, `discount_rate`, `days_since_last_purchase`  
      - Categorize into tiers or buckets
17. Check for invalid numeric values (e.g. negative price or quantity)
18. Rename and format columns consistently (e.g. snake_case)
19. Ensure uniqueness of key columns (e.g. `order_id`)
20. Merge with other datasets (products, customers) safely
21. Recheck and convert data types (datetime, float, string)
22. Export the cleaned dataset (new file or table)
23. Document all cleaning steps (log or README) |
    Python (pandas): `dropna()`, `fillna()`, `str.strip()`<br>Excel: `TRIM`, `IFERROR`<br>Power Query: clean, merge, transform |
    | 4 | Exploratory Data Analysis (EDA) | Explore patterns and trends |
24. Generate descriptive statistics
25. Identify top-selling products and customers
26. Analyze trends over time (daily/monthly/yearly)
27. Segment by region, product category, customer group
28. Visualize anomalies and outliers
29. Check distribution of numeric features | Excel charts, seaborn, matplotlib, Power BI |
    | 5 | Insight Extraction | Derive business-relevant findings |
30. Highlight key patterns (e.g. sales drop in Q2)
31. Correlate with promotions, campaigns, or seasonality
32. Explain root causes of performance shifts
33. Prioritize actionable insights | Jupyter Notebook, Markdown, Docs |
    | 6 | Visualization / Dashboard | Communicate insights visually |
34. Build KPI trackers (sales, margin, growth)
35. Design interactive dashboards
36. Use filters for time, region, category
37. Apply effective color & chart types | Power BI, Tableau, Looker Studio |
    | 7 | Recommendation | Suggest actions based on insights |
38. Provide actionable business suggestions
39. Recommend focus areas (e.g. region or product)
40. Identify opportunities for optimization (pricing, discount, inventory)
41. Forecast possible outcomes | Business Memo, Report, Strategy Brief |
    | 8 | Communication | Present insights clearly to stakeholders |
42. Summarize insights into slides or executive brief
43. Emphasize impact and clear call-to-action
44. Present story (problem → analysis → solution)
45. Anticipate and answer business questions | Google Slides, PowerPoint, PDF Report |
