# 📊 Sales Analysis & KPI Dashboard --- Excel Project

## 📌 Project Overview

This project is an **Excel-based Sales Analysis and KPI Dashboard**
created to analyze transaction-level sales data and present important
business insights using formulas, KPI summaries, PivotTable-style
summaries, and dashboard components.

The workbook contains **1,000 sales transactions** and includes branch,
city, customer type, gender, product line, unit price, quantity, VAT,
total revenue, date/time, payment method, and customer rating
information.

> **Note:** The main worksheet is named `amazon_sales_excel`, but the
> data structure is a generic retail/supermarket sales dataset. This
> README therefore describes it as a sales-analysis project without
> assuming a specific company or original data source.

------------------------------------------------------------------------

## 📁 Workbook Structure

  -----------------------------------------------------------------------
  Sheet                               Purpose
  ----------------------------------- -----------------------------------
  `amazon_sales_excel`                Main transaction-level sales
                                      dataset

  `Sheet1`                            Supporting calculations and
                                      unique-value/helper data

  `Dashboard`                         Dashboard/presentation area

  `KPI`                               KPI summaries, branch revenue,
                                      product-line revenue, quantity and
                                      rating metrics
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 🗃️ Dataset Details

### Dataset Size

-   **Total transactions/orders:** 1,000
-   **Total columns:** 14
-   **Total quantity sold:** 5,510 units
-   **Total revenue:** 322,966.749
-   **Average customer rating:** 6.9727 / 10

### Main Dataset Columns

  Column             Description
  ------------------ --------------------------------------------------
  `invoice_id`       Unique invoice/transaction identifier
  `branch`           Branch identifier
  `city`             Transaction city
  `customer_type`    Customer classification such as Member or Normal
  `gender`           Customer gender
  `product_line`     Product category purchased
  `unit_price`       Price of one unit
  `quantity`         Number of units purchased
  `vat`              Value Added Tax associated with the transaction
  `total`            Total transaction value
  `date`             Transaction date
  `time`             Transaction time
  `payment_method`   Payment method used
  `rating`           Customer rating

------------------------------------------------------------------------

# 📈 Key Performance Indicators

  KPI                                   Value
  ------------------------- -----------------
  **Total Orders**                  **1,000**
  **Total Revenue**           **322,966.749**
  **Total Quantity Sold**           **5,510**
  **Average Rating**          **6.9727 / 10**

These KPIs provide a high-level view of overall sales performance and
customer feedback.

------------------------------------------------------------------------

# 🏢 Revenue by Branch

  Branch                       Revenue
  ----------------- ------------------
  A                        39,383.1165
  B                        68,516.8785
  C                        69,656.1390
  D                    **85,349.5965**
  E                        60,061.0185
  **Grand Total**     **322,966.7490**

### 🔎 Branch Insights

-   **Branch D** generates the highest revenue.
-   **Branch C** is the second-highest revenue contributor.
-   **Branch A** has the lowest revenue among the listed branches.
-   Branch comparison can help identify strong locations and areas
    requiring improvement.

------------------------------------------------------------------------

# 🛍️ Revenue & Quantity by Product Line

  Product Line                        Revenue    Quantity
  ------------------------ ------------------ -----------
  Health and beauty               49,193.7390         854
  Fashion accessories             54,305.8950         902
  Home and lifestyle              53,861.9130         911
  Sports and travel               55,122.8265         920
  Food and beverages          **56,144.8440**         952
  Electronic accessories          54,337.5315     **971**
  **Grand Total**            **322,966.7490**   **5,510**

### 🔎 Product-Line Insights

-   **Food and beverages** has the highest revenue.
-   **Electronic accessories** has the highest quantity sold.
-   **Health and beauty** has the lowest quantity in this summary.
-   Revenue is reasonably distributed across the six product lines.

------------------------------------------------------------------------

# 🧮 Supporting Sheet --- `Sheet1`

`Sheet1` acts as a helper/calculation sheet.

It contains:

-   Total row calculation
-   Total column calculation
-   Unique city values
-   Unique branch values
-   Formula-based helper calculations

Example formula used:

``` excel
=COUNTA(amazon_sales_excel!A2:A1001)
```

This counts the populated invoice IDs and returns the total number of
transactions.

------------------------------------------------------------------------

# 📊 KPI Sheet --- `KPI`

The `KPI` worksheet brings together the main analytical summaries:

### KPI Summary

-   Total Orders: **1,000**
-   Revenue: **322,966.749**
-   Quantity: **5,510**
-   Average Rating: **6.9727**

### Branch Analysis

Revenue is grouped by branches A--E.

### Product Analysis

Revenue and quantity are grouped by product line.

### Rating Analysis

The overall average customer rating is approximately **6.97/10**.

------------------------------------------------------------------------

# 📊 Dashboard

The workbook contains a dedicated `Dashboard` sheet intended as the
visual presentation layer.

A complete dashboard can include:

-   💰 Total Revenue
-   🧾 Total Orders
-   📦 Total Quantity Sold
-   ⭐ Average Rating
-   🏢 Revenue by Branch
-   🛍️ Revenue by Product Line
-   📦 Quantity by Product Line
-   💳 Payment Method Distribution
-   👥 Customer Type Analysis
-   🚻 Gender-wise Sales Analysis
-   📅 Sales Trend by Date
-   🏙️ City-wise Revenue

The `KPI` sheet provides the supporting metrics for dashboard reporting.

------------------------------------------------------------------------

# 🔍 Business Questions This Project Can Answer

1.  What is the total revenue?
2.  How many orders were placed?
3.  How many products were sold?
4.  Which branch generates the highest revenue?
5.  Which branch performs the weakest?
6.  Which product line generates the highest revenue?
7.  Which product line sells the highest quantity?
8.  What is the average customer rating?
9.  Which cities contribute most to sales?
10. How do Member and Normal customers compare?
11. How do male and female customers contribute to sales?
12. Which payment methods are most frequently used?
13. How do sales change over time?
14. Which product categories may need additional marketing attention?

------------------------------------------------------------------------

# 🛠️ Excel Features Used

This project demonstrates:

-   Tabular data organization
-   Excel formulas
-   `COUNTA` calculations
-   Unique-value/helper calculations
-   KPI summaries
-   Branch-wise aggregation
-   Product-line aggregation
-   Revenue analysis
-   Quantity analysis
-   Customer-rating analysis
-   PivotTable-style summaries
-   Dashboard-oriented reporting

------------------------------------------------------------------------

# 🎯 Project Objectives

The main objectives are:

-   Organize raw sales transaction data.
-   Calculate important business KPIs.
-   Analyze branch performance.
-   Analyze product-line performance.
-   Measure revenue and quantity.
-   Evaluate customer ratings.
-   Prepare data for dashboard visualization.
-   Support business decisions using Excel.

------------------------------------------------------------------------

# 🚀 How to Use the Workbook

### Step 1 --- Open the Excel file

Open the `.xlsx` workbook in Microsoft Excel or a compatible spreadsheet
application.

### Step 2 --- Explore the raw data

Open:

`amazon_sales_excel`

This is the main transaction-level dataset.

### Step 3 --- Review helper calculations

Open:

`Sheet1`

This contains supporting calculations and unique-value information.

### Step 4 --- Review KPIs

Open:

`KPI`

Review total orders, revenue, quantity, rating, branch revenue, and
product-line performance.

### Step 5 --- Review the dashboard

Open:

`Dashboard`

Use it as the visual reporting layer and extend it with charts, KPI
cards, slicers, and filters as required.

------------------------------------------------------------------------

# 💡 Possible Future Improvements

The project can be extended with:

-   Interactive PivotCharts
-   Slicers for branch, city, and product line
-   Date filters
-   Monthly sales trends
-   Revenue contribution percentages
-   Customer-type comparison
-   Gender-wise sales analysis
-   Payment-method analysis
-   Top-performing categories
-   Conditional formatting
-   Interactive KPI cards
-   Automated dashboard refresh
-   Profit/margin analysis when cost data is available

------------------------------------------------------------------------

# 📊 Analytical Summary

### Overall Performance

-   **1,000 orders** were recorded.
-   **5,510 units** were sold.
-   Total revenue was approximately **322.97K**.
-   Average customer rating was approximately **6.97/10**.

### Branch Performance

**Branch D** is the strongest revenue-generating branch with
approximately **85,349.60** in revenue.

### Product Performance

**Food and beverages** generates the highest product-line revenue at
approximately **56,144.84**.

**Electronic accessories** has the highest quantity sold with **971
units**.

### Customer Experience

The average rating of approximately **6.97/10** provides a useful
baseline for customer-experience analysis.

------------------------------------------------------------------------

# 📂 Recommended GitHub Repository Structure

``` text
sales-analysis-excel/
│
├── README.md
├── sales_analysis.xlsx
├── screenshots/
│   ├── dashboard.png
│   └── kpi.png
│
└── documentation/
    └── project_notes.md
```

------------------------------------------------------------------------

# 🧑‍💻 Project Information

**Project Type:** Data Analysis / Business Intelligence

**Tool:** Microsoft Excel

**Focus:** Sales Analytics, KPI Reporting & Dashboard

**Dataset Size:** 1,000 transactions

**Primary Output:** Excel-based KPI and sales analysis

------------------------------------------------------------------------

# ⭐ Key Takeaway

This project transforms transaction-level sales data into meaningful
business information through **KPI calculations, branch analysis,
product-line analysis, quantity analysis, and customer-rating
analysis**.

The analysis reports approximately **322.97K total revenue**, with
**Branch D** leading branch revenue, **Food and beverages** leading
product-line revenue, and **Electronic accessories** leading product
quantity.

------------------------------------------------------------------------

## 📜 License

This project is intended for **educational, portfolio, and data-analysis
demonstration purposes**.

If the dataset is redistributed, verify and follow the license or terms
of the original dataset source.

------------------------------------------------------------------------

## 👤 Author

**Sales Analysis Excel Project**

Built as an Excel-based data analytics and KPI dashboard project.
