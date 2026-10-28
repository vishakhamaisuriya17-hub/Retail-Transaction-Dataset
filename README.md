# Retail Sales Analysis Dashboard Using Microsoft Excel

## Project Overview

The **Retail Sales Analysis Dashboard** is an Excel-based data analytics project designed to transform retail transaction data into meaningful business insights through data cleaning, data preparation, advanced formulas, PivotTables, data visualization, and interactive dashboard development.

The project demonstrates practical Microsoft Excel skills used in data analysis and reporting. It focuses on analyzing retail sales performance across different cities, payment methods, store types, customer categories, seasons, and promotional activities.

The dashboard combines key performance indicators (KPIs), charts, slicers, and a timeline to help users explore transaction data, compare sales performance, and identify patterns in retail operations.

This project is designed to demonstrate beginner-to-intermediate Excel skills and provide a practical example of how Excel can be used for business intelligence and data-driven decision-making.

---

## Table of Contents

1. [Project Objectives](#project-objectives)
2. [Project Details](#project-details)
3. [Tools and Technologies](#tools-and-technologies)
4. [Dataset Description](#dataset-description)
5. [Project Workflow](#project-workflow)
6. [Data Cleaning and Preparation](#data-cleaning-and-preparation)
7. [PivotTables and Data Analysis](#pivottables-and-data-analysis)
8. [Advanced Excel Formulas](#advanced-excel-formulas)
9. [Data Visualization](#data-visualization)
10. [Interactive Dashboard](#interactive-dashboard)
11. [Key Performance Indicators](#key-performance-indicators)
12. [Dashboard Components](#dashboard-components)
13. [Workbook Structure](#workbook-structure)
14. [How to Use the Project](#how-to-use-the-project)
15. [Skills Demonstrated](#skills-demonstrated)
16. [Business Applications](#business-applications)
17. [Project Limitations](#project-limitations)
18. [Future Improvements](#future-improvements)
19. [Conclusion](#conclusion)
20. [Author](#author)

---

## Project Objectives

The main objective of this project is to develop an interactive retail sales dashboard using Microsoft Excel and demonstrate the complete data analysis workflow.

The specific objectives are:

* Import retail transaction data into Excel.
* Inspect and clean the dataset.
* Identify duplicate records and missing values.
* Validate column data types and apply appropriate formatting.
* Convert the dataset into a structured Excel Table.
* Create PivotTables to summarize retail transaction data.
* Apply advanced Excel formulas for calculations and lookups.
* Analyze sales performance across cities, payment methods, and store types.
* Compare sales across seasons, customer categories, and promotional activities.
* Create charts and apply conditional formatting.
* Develop KPI cards for important business metrics.
* Build an interactive dashboard using slicers and a timeline.
* Organize the workbook and document the analysis process.

---

## Project Details

| Attribute            | Description                                                           |
| -------------------- | --------------------------------------------------------------------- |
| Project Name         | Retail Sales Analysis Dashboard                                       |
| Repository Name      | Retail-Transaction-Dataset                                            |
| Author               | Vishakha Maisuriya                                                    |
| Primary Tool         | Microsoft Excel                                                       |
| Project Category     | Data Analysis and Visualization                                       |
| Dataset Type         | Retail Transaction Data                                               |
| Analysis Methods     | Data Cleaning, PivotTables, Formula-Based Analysis, and Visualization |
| Dashboard Components | KPI Cards, Charts, Slicers, and Timeline                              |
| Output               | Excel Workbook and Interactive Dashboard                              |

---

## Tools and Technologies

### Microsoft Excel

Microsoft Excel is the primary tool used throughout this project.

The project applies the following Excel features:

* Excel Tables
* Data Cleaning Tools
* Remove Duplicates
* Text to Columns
* Cell Formatting
* Conditional Formatting
* PivotTables
* PivotCharts
* VLOOKUP
* INDEX-MATCH
* IF and Nested IF Functions
* Date and Time Functions
* KPI Cards
* Slicers
* Timelines
* Dashboard Formatting

These features help organize raw data, perform calculations, summarize information, and communicate findings through visual reports.

---

## Dataset Description

The project uses a retail transaction dataset containing information about transactions, customers, products, transaction costs, payment methods, locations, store types, discounts, seasons, and promotions.

The dataset provides a foundation for exploring retail sales patterns and comparing different transaction categories.

### Dataset Columns

| Column Name         | Description                                         | Data Type            |
| ------------------- | --------------------------------------------------- | -------------------- |
| `Transaction_ID`    | Unique identifier for a transaction                 | Text or Whole Number |
| `Date`              | Date and time of the transaction                    | Date/Time            |
| `Customer_Name`     | Name of the customer                                | Text                 |
| `Product`           | Product or products associated with the transaction | Text                 |
| `Total_Items`       | Number of items recorded for the transaction        | Whole Number         |
| `Total_Cost`        | Total recorded transaction cost                     | Number               |
| `Payment_Method`    | Payment method used for the transaction             | Text                 |
| `City`              | City associated with the transaction                | Text                 |
| `Store_Type`        | Type of retail store                                | Text                 |
| `Discount_Applied`  | Indicates whether a discount was applied            | Boolean or Text      |
| `Customer_Category` | Customer category assigned to the transaction       | Text                 |
| `Season`            | Season associated with the transaction              | Text                 |
| `Promotion`         | Promotion associated with the transaction           | Text                 |

**Note:** The table above describes the expected fields used in this project. The actual number of records, date range, category values, and data quality issues should be determined from the dataset being analyzed.

---

## Project Workflow

The project follows a structured seven-task workflow.

| Task                             | Description                                  | Key Excel Features                          |
| -------------------------------- | -------------------------------------------- | ------------------------------------------- |
| 1. Dataset Selection and Import  | Import the transaction data into Excel       | Data Import, Power Query                    |
| 2. Data Cleaning and Preparation | Inspect and clean the dataset                | Remove Duplicates, Filters, Data Formatting |
| 3. Create PivotTables            | Summarize transaction data                   | PivotTables, Aggregation                    |
| 4. Advanced Formulas             | Perform lookups and conditional calculations | VLOOKUP, INDEX-MATCH, IF, YEAR, TEXT        |
| 5. Data Visualization            | Present findings using charts                | PivotCharts, Conditional Formatting         |
| 6. Interactive Dashboard         | Build an interactive reporting interface     | KPI Cards, Slicers, Timeline                |
| 7. Documentation                 | Organize and document the workbook           | Worksheet Organization, Naming Conventions  |

Each task contributes to the final dashboard and helps demonstrate a different aspect of Excel-based data analysis.

---

## Data Cleaning and Preparation

Data cleaning is an essential step because incorrect or inconsistent data can affect calculations and analytical results.

The dataset is inspected and prepared before creating PivotTables and charts.

### 1. Importing the Dataset

The transaction dataset is imported into Microsoft Excel.

The imported data is checked to ensure that each field is placed in the correct column and that the column headers are clearly defined.

### 2. Creating an Excel Table

The dataset is converted into a structured Excel Table.

The table is named:

`tblTransactions`

Using a structured table makes it easier to reference columns in formulas and manage the dataset.

### 3. Identifying Duplicate Records

Duplicate records are inspected using Excel's Remove Duplicates feature.

Transaction identifiers are checked to identify potential duplicate transactions.

Records are removed only when they are confirmed to be duplicates. Multiple transactions from the same customer are not automatically treated as duplicates.

### 4. Handling Missing Values

The dataset is inspected for missing values in important columns, including:

* Transaction_ID
* Date
* Customer_Name
* Product
* Total_Items
* Total_Cost
* Payment_Method
* City

Missing values are investigated before deciding whether records should be corrected, retained, or excluded.

Unknown values are not automatically replaced with zero because zero and missing data have different meanings.

### 5. Formatting Date and Numeric Columns

Appropriate formats are applied to improve consistency and readability.

Examples include:

* Date columns formatted as date and time.
* Transaction costs formatted as numbers with two decimal places.
* Item counts formatted as whole numbers.
* Text fields standardized where appropriate.

### 6. Validating Data Types

Each column is checked to ensure that its values use an appropriate data type.

Dates should be recognized as actual Excel date values, while numeric fields should contain valid numeric data.

### 7. Creating a Cleaned Dataset

The original imported data is preserved separately from the cleaned dataset.

This makes it easier to compare the original data with the prepared data and maintain a clear analytical workflow.

---

## PivotTables and Data Analysis

PivotTables are used to summarize transaction data and compare sales performance across different categories.

The following analytical summaries are included in the project workflow.

### 1. Total Sales by City

**Rows:** City

**Values:** Sum of Total_Cost

Purpose:

* Compare recorded transaction costs across cities.
* Identify cities with relatively high or low sales totals.
* Support location-based sales analysis.

### 2. Sales by Payment Method

**Rows:** Payment_Method

**Values:** Sum of Total_Cost

Purpose:

* Compare sales across payment methods.
* Understand how transaction value is distributed among payment categories.
* Support payment-method reporting.

### 3. Store Type Analysis

**Rows:** Store_Type

**Values:**

* Sum of Total_Cost
* Count of Transaction_ID

Purpose:

* Compare recorded transaction costs across store types.
* Examine transaction counts by store type.
* Understand differences in transaction activity.

### 4. Seasonal Sales Analysis

**Rows:** Season

**Values:** Sum of Total_Cost

Purpose:

* Compare transaction costs across seasons.
* Explore potential seasonal variations in sales.
* Support seasonal reporting.

### 5. Customer Category Analysis

**Rows:** Customer_Category

**Values:**

* Sum of Total_Cost
* Count of Transaction_ID

Purpose:

* Compare recorded sales across customer categories.
* Examine transaction distribution across customer groups.
* Support customer-category analysis.

### 6. Promotion Analysis

**Rows:** Promotion

**Values:**

* Sum of Total_Cost
* Count of Transaction_ID

Purpose:

* Compare transactions associated with different promotions.
* Examine recorded transaction costs by promotion.
* Support promotional performance analysis.

The actual findings depend on the values and quality of the dataset. No sales ranking or business conclusion should be assumed before analyzing the results.

---

## Advanced Excel Formulas

Advanced Excel formulas are used to extract information, classify transactions, and retrieve values based on transaction identifiers.

### 1. YEAR Function

The YEAR function extracts the year from a transaction date.

Formula:

```excel
=YEAR([@Date])
```

Purpose:

* Extract the transaction year.
* Support year-based analysis.
* Create a year field for reporting.

### 2. TEXT Function

The TEXT function converts a date value into a formatted month name.

Formula:

```excel
=TEXT([@Date],"mmm")
```

Example output:

| Date        | Month |
| ----------- | ----- |
| 15-Jan-2022 | Jan   |
| 18-Feb-2022 | Feb   |
| 21-Mar-2022 | Mar   |

Purpose:

* Extract readable month labels.
* Support month-based reporting.
* Improve the presentation of date-related information.

For chronological month analysis, a real date field or a separate month-number field can be used to ensure that months appear in calendar order.

### 3. Nested IF Function

The nested IF function categorizes transaction costs into Low, Medium, and High groups.

Example formula:

```excel
=IF([@Total_Cost]<25,"Low",IF([@Total_Cost]<60,"Medium","High"))
```

Classification rules:

| Condition                 | Category |
| ------------------------- | -------- |
| Total_Cost < 25           | Low      |
| Total_Cost >= 25 and < 60 | Medium   |
| Total_Cost >= 60          | High     |

These thresholds are illustrative project rules and can be adjusted according to the analytical requirements.

Purpose:

* Classify transaction costs.
* Apply conditional logic.
* Support category-based summaries.

### 4. IF Function for Discount Status

The IF function identifies whether a discount was applied.

For a column containing actual Boolean TRUE/FALSE values:

```excel
=IF([@Discount_Applied]=TRUE,"Discount Applied","No Discount")
```

If the column contains text values instead, the formula must be adjusted to match the actual data format.

Purpose:

* Create readable discount categories.
* Simplify discount-related reporting.
* Support the calculation of discounted transaction counts.

### 5. VLOOKUP Function

VLOOKUP searches for a transaction identifier and returns the corresponding customer name.

Example formula:

```excel
=VLOOKUP(A2,tblTransactions[[Transaction_ID]:[Customer_Name]],2,FALSE)
```

Where:

* `A2` contains the transaction identifier to search for.
* `tblTransactions` is the structured Excel Table.
* `2` specifies the return column within the selected table range.
* `FALSE` requests an exact match.

Purpose:

* Retrieve customer information using transaction identifiers.
* Practice exact-match lookups.
* Demonstrate lookup-based data retrieval.

### 6. INDEX-MATCH Combination

INDEX-MATCH retrieves the customer name associated with a transaction identifier.

Formula:

```excel
=INDEX(tblTransactions[Customer_Name],MATCH(A2,tblTransactions[Transaction_ID],0))
```

How it works:

* `MATCH` identifies the position of the transaction identifier.
* `INDEX` returns the customer name at that position.
* `0` requests an exact match.

Purpose:

* Demonstrate flexible lookup techniques.
* Retrieve values using matching identifiers.
* Compare INDEX-MATCH with VLOOKUP.

**Important:** These formulas assume that the structured table and column names match the workbook. Formula results depend on the underlying data being correctly formatted.

---

## Data Visualization

Data visualization converts analytical summaries into charts that are easier to interpret.

The project uses PivotCharts and conditional formatting to present transaction patterns visually.

### 1. Sales by City Chart

A column chart is used to compare total transaction costs across cities.

This visualization supports location-based sales comparisons.

### 2. Sales by Payment Method Chart

A bar chart is used to compare recorded transaction costs across payment methods.

The chart helps users compare payment categories.

### 3. Seasonal Sales Chart

A column chart is used to compare sales totals across seasons.

It provides a visual overview of seasonal transaction patterns.

### 4. Sales by Store Type Chart

A bar chart is used to compare transaction costs across store types.

This helps users explore differences between store categories.

### 5. Conditional Formatting

Conditional formatting is applied to selected fields to make important values easier to identify.

Examples include:

* Color scales for the Total_Cost column.
* Formatting rules for Low, Medium, and High sales categories.
* Highlighting values that require attention.

Conditional formatting improves readability without changing the underlying transaction data.

---

## Interactive Dashboard

The interactive dashboard combines KPI cards, charts, slicers, and a timeline into a single reporting interface.

The dashboard is designed to make transaction analysis more accessible and reduce the need to inspect individual records manually.

### Dashboard Components

The dashboard includes the following components:

* Dashboard title
* Total Sales KPI card
* Total Transactions KPI card
* Average Transaction Value KPI card
* Discounted Transactions KPI card
* Sales by City chart
* Sales by Payment Method chart
* Seasonal Sales chart
* Sales by Store Type chart
* City slicer
* Payment Method slicer
* Store Type slicer
* Season slicer
* Date Timeline

The final workbook should be tested to confirm that the charts and filtering controls behave as intended.

### Dashboard Layout

The recommended layout is:

| Dashboard Section | Components                                           |
| ----------------- | ---------------------------------------------------- |
| Header            | Dashboard title                                      |
| KPI Section       | Four KPI cards                                       |
| Chart Section 1   | Sales by City and Sales by Payment Method            |
| Chart Section 2   | Seasonal Sales and Sales by Store Type               |
| Filter Section    | City, Payment Method, Store Type, and Season slicers |
| Date Filter       | Timeline                                             |

The layout can be adjusted to suit the screen size and available worksheet space.

### Dashboard Interactivity

Slicers allow users to filter connected PivotTables and PivotCharts using selected categories.

The timeline allows users to filter date-based PivotTables by a selected time period.

For consistent filtering across the dashboard, relevant slicers and timelines should be connected to compatible PivotTables through Excel's Report Connections or PivotTable Connections settings.

**Note:** Formula-based KPI cards using SUM, COUNTA, and AVERAGE calculate values from the full table. They do not automatically respond to PivotTable slicers. For fully interactive KPI cards, the displayed values must be linked to appropriately filtered PivotTable summaries or another filtering-aware calculation method.

---

## Key Performance Indicators

The dashboard includes four primary KPI cards.

### 1. Total Sales

Formula:

```excel
=SUM(tblTransactions[Total_Cost])
```

Purpose:

Calculates the sum of the numeric values in the Total_Cost column.

This KPI provides an overview of the total recorded transaction cost in the dataset.

### 2. Total Transactions

Formula:

```excel
=COUNTA(tblTransactions[Transaction_ID])
```

Purpose:

Counts the nonblank transaction identifiers.

The result represents the number of records with a populated transaction identifier. Unique transaction counts require the identifiers to be validated and duplicates handled appropriately.

### 3. Average Transaction Value

Formula:

```excel
=AVERAGE(tblTransactions[Total_Cost])
```

Purpose:

Calculates the average numeric transaction cost.

This KPI provides a measure of the typical recorded transaction value.

### 4. Discounted Transactions

For an Excel column containing actual Boolean TRUE/FALSE values:

```excel
=COUNTIF(tblTransactions[Discount_Applied],TRUE)
```

Purpose:

Counts records where a discount was applied.

If the column contains text instead of Boolean values, the formula must be adapted to the actual cell contents.

### KPI Interpretation

| KPI                       | Business Purpose                                        |
| ------------------------- | ------------------------------------------------------- |
| Total Sales               | Understand total recorded transaction costs             |
| Total Transactions        | Measure transaction-record volume                       |
| Average Transaction Value | Understand average transaction cost                     |
| Discounted Transactions   | Measure the number of transactions marked as discounted |

These metrics are intended to support reporting and exploration. Their interpretation depends on the dataset's definitions and data quality.

---

## Workbook Structure

The Excel workbook is organized into separate worksheets to maintain a clear workflow.

| Worksheet          | Purpose                                                                         |
| ------------------ | ------------------------------------------------------------------------------- |
| `Dashboard`        | Displays KPI cards, charts, slicers, and the timeline                           |
| `Raw_Data`         | Preserves the imported source dataset                                           |
| `Cleaned_Data`     | Contains the prepared dataset used for analysis                                 |
| `Pivot_Analysis`   | Stores PivotTables used to summarize transaction data                           |
| `Formula_Practice` | Demonstrates lookup formulas and related calculations                           |
| `Charts`           | Organizes charts before they are placed on the dashboard                        |
| `Documentation`    | Records project information, worksheet descriptions, and dashboard instructions |

The actual workbook may contain additional sheets or use different names depending on the final implementation.

---

## How to Use the Project

Follow these steps to explore the workbook.

### Step 1: Open the Excel Workbook

Download or open the project's Excel workbook from the repository, if the workbook has been uploaded.

Use a compatible desktop version of Microsoft Excel for the best support of PivotTables, PivotCharts, slicers, and timelines.

### Step 2: Review the Raw Data

Open the `reatail_dataset` worksheet to inspect the imported transaction records and original column structure.

### Step 3: Review the Cleaned Data

Open `Retail_dataset` to examine the prepared data and any additional calculated columns.

### Step 4: Explore the PivotTables

Open `Pivot_Analysis` to review the summaries used for city, payment method, store type, season, customer category, and promotion analysis.

### Step 5: Review the Formulas

Open `Formula` to inspect the VLOOKUP, INDEX-MATCH, IF, and other formulas used in the project.

### Step 6: Open the Dashboard

Navigate to the `Dashboard` worksheet.

Review the KPI cards and charts to understand the overall transaction data.

### Step 7: Apply Filters

Use the available slicers to select cities, payment methods, store types, or seasons.

Use the timeline to filter date-based summaries.

Only connected PivotTables and charts will respond to a particular filter.

### Step 8: Refresh the Analysis

If the underlying data has changed, refresh the relevant PivotTables.

Depending on the workbook setup, use:

`Data → Refresh All`

Then verify that the charts, PivotTables, and dashboard values have updated correctly.

### Step 9: Save the Workbook

Save changes after reviewing the results.

Keep a copy of the original dataset when making substantial changes to the analysis.

---

## Skills Demonstrated

This project demonstrates practical skills in the following areas.

### Data Preparation

* Importing data into Excel
* Converting ranges into structured tables
* Identifying duplicate records
* Inspecting missing values
* Validating data types
* Formatting date and numeric fields

### Excel Formulas

* SUM
* COUNTA
* AVERAGE
* COUNTIF
* YEAR
* TEXT
* IF
* Nested IF
* VLOOKUP
* INDEX
* MATCH

### Data Analysis

* Creating PivotTables
* Summarizing numeric values
* Counting transaction records
* Comparing categories
* Analyzing transaction costs
* Examining seasonal patterns

### Data Visualization

* Creating column charts
* Creating bar charts
* Creating PivotCharts
* Applying conditional formatting
* Designing KPI cards
* Organizing dashboard components

### Dashboard Development

* Creating an interactive Excel dashboard
* Adding slicers
* Adding a date timeline
* Connecting filters to compatible PivotTables
* Organizing worksheet layouts
* Improving report readability

### Documentation

* Applying consistent naming conventions
* Organizing workbook sheets
* Documenting formulas and analytical steps
* Preparing dashboard usage instructions

---

## Business Applications

The techniques demonstrated in this project can be applied to several retail reporting scenarios.

### Sales Performance Monitoring

Businesses can summarize transaction costs across cities and store types to understand how sales are distributed across locations and channels.

### Payment Method Analysis

Transaction summaries can help compare the recorded value and volume of transactions across payment methods.

### Seasonal Reporting

Season-based summaries can help explore whether transaction activity varies across different seasons.

### Customer Category Analysis

Customer-category summaries can help identify how transaction value and volume are distributed across customer groups.

### Promotional Reporting

Promotion-based summaries can help compare the transaction costs and record counts associated with different promotional categories.

### Management Reporting

An interactive dashboard provides a consolidated view of selected performance indicators and analytical summaries.

**Important:** These are potential applications of the analysis. Actual business conclusions must be based on verified calculations and observed results from the dataset.

---

## Project Limitations

The project has several considerations that should be kept in mind when interpreting the results.

1. **Data quality:** Duplicate records, missing values, and incorrect data types can affect analytical accuracy.

2. **Transaction identifiers:** Counting nonblank transaction IDs does not necessarily equal counting unique transactions if duplicate identifiers exist.

3. **Transaction cost interpretation:** Total_Cost is treated as the recorded transaction amount. Its precise business definition and currency should be confirmed from the dataset documentation.

4. **Product information:** If multiple product names are stored in one cell, detailed product-level analysis may require additional data preparation.

5. **Discount information:** Discounted transaction counts depend on the correct interpretation of the Discount_Applied column.

6. **Dashboard filtering:** KPI formulas based directly on the complete table will not automatically respond to slicers unless they are implemented using a filtering-aware approach.

7. **Data refresh:** PivotTables may need to be refreshed when the underlying dataset changes.

8. **Business conclusions:** The dashboard presents descriptive analysis and does not independently establish causation or guarantee future sales performance.

---

## Future Improvements

The project can be extended with additional Excel features and analytical capabilities.

Potential improvements include:

* Adding more advanced KPI metrics.
* Creating month-over-month transaction cost comparisons.
* Developing additional date-based PivotTables.
* Adding top-performing city or store summaries.
* Expanding product-level analysis by restructuring product information.
* Improving interactive KPI calculations.
* Adding data validation rules.
* Automating data preparation with Power Query.
* Creating additional conditional formatting rules.
* Improving dashboard layout and usability.
* Adding a detailed summary of verified analytical findings.
* Integrating additional data sources for broader reporting.

These enhancements can build on the existing Excel workflow and expand the project's analytical capabilities.

---

## Conclusion

The Retail Sales Analysis Dashboard demonstrates how Microsoft Excel can be used to organize, prepare, analyze, and visualize retail transaction data.

By combining data cleaning, structured Excel Tables, advanced formulas, PivotTables, charts, KPI cards, slicers, and a timeline, the project provides a practical example of an end-to-end Excel analytics workflow.

The project also reinforces the importance of data quality, accurate calculations, clear documentation, and effective data visualization.

Overall, it serves as a practical learning project for developing Excel-based data analysis and dashboard-building skills.

---

## Author

**Vishakha Maisuriya**

GitHub: [vishakhamaisuriya17-hub](https://github.com/vishakhamaisuriya17-hub)

Project Repository: [Retail-Transaction-Dataset](https://github.com/vishakhamaisuriya17-hub/Retail-Transaction-Dataset)

---

## Acknowledgment

This project was developed for practical learning and portfolio development in Microsoft Excel, data analysis, and dashboard visualization.

The project focuses on demonstrating Excel functionality through a structured retail transaction analysis workflow.
