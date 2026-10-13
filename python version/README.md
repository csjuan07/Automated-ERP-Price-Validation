# Automated ERP Price Validation (Python Version)

## Overview

This repository presents a Python-based alternative implementation for the automated monthly validation of supplier invoice prices against ERP reference prices.

While the original architecture relies on Microsoft Power Automate and Office Scripts (TypeScript), this standalone Python pipeline processes transaction records in memory, performs cross-table reference lookups, converts currencies, calculates spend deviations, and dynamically populates the centralized Excel validation model.

---

## Architecture Comparison

### Original Workflow (Power Automate + Office Scripts)
Monthly Report ---> Outlook Email ---> Power Automate ---> Office Scripts ---> Excel Validation Model ---> Power BI

### Python Alternative Workflow
Local / Cloud Source Files (.xlsx)
|
v
Python Data Engine (Pandas \& OpenPyXL)
|-- 1. Data Ingestion \& Array Normalization
|-- 2. ERP Price \& Currency Reference Lookups
|-- 3. Unit Cost \& Currency Conversion Adjustments
|-- 4. Spend \& Variance Analysis Calculations
`-- 5. Dynamic Formatting \& Table Append (OpenPyXL)
|
v
Excel Validation Model (priceDivergence.xlsx)
|
v
Power BI Dashboard

\---

### Technical Highlights \& Advantages

* High-Speed Processing: Replaces sequential Office Scripts execution and cloud batch inserting with vectorized pandas operations in memory.
* Robust Table Appending: Uses openpyxl to detect existing Excel table ranges (table.ref), preserve cell styles, borders, and number formats from previous rows, and dynamically update the table boundaries.
* Simplified Workflow: Consolidates data extraction, matching, currency conversion, category enrichment, and variance calculations into a single maintainable Python script (ERP\_Price\_Validation.py).

---

## Workflow \& Logic Breakdown

### 1. Data Ingestion
The script ingests four primary data sources via pandas:

* Invoices / NFs: Invoice details (Part Numbers, Supplier IDs, Prices, Volumes, Currencies).
* ERP Reference Prices: Net prices, units of measure (UM), currencies, and supplier IDs.
* Exchange Rates: Currency conversion rates for BRL calculations.
* Category \& Regional Reference: Category Manager (CM) and Regional/Global (R/G) mappings.

### 2. Key Generation \& Matching
Creates composite lookup keys (PartNumber|SupplierID) for precise matching between invoice lines and ERP master data.

### 3. Price Adjustments \& Currency Conversion

* Converts unit prices based on ERP measure units (e.g., MI divided by 1000, CT divided by 100).
* Converts invoice and ERP prices into base currency (BRL) using exchange rate mappings.

### 4. Calculations \& Classification
Calculates the following metrics per line:

* Actual Spend (Original \& BRL)
* ERP Spend (Original \& BRL)
* Absolute Price Difference (BRL)
* Percentage Variance (dif\_pct)
* Location Classifier (L / I / O): Based on currency and supplier attributes.

### 5. Dynamic Excel Output
Inserts the processed records into priceDivergence.xlsx inside PriceDivergenceTable, preserving formatting and expanding the table range automatically.

---

## Script Mapping (Office Scripts vs. Python)

|Original Office Script (.ts)|Python Equivalent (ERP\_Price\_Validation.py)|Function / Responsibility|
|-|-|-|
|copyActual.ts|pandas.read\_excel() \& Series extraction|Reads source sheets and converts columns into data structures.|
|findEmptyCell.ts|openpyxl table bounds inspection|Identifies table boundaries and target rows for insertion.|
|pastePDLotes.ts|add\_df\_data\_to\_existing\_excel\_table()|Appends batch records to the target Excel table with style cloning.|
|updatePriceDivergence.ts|Core processing loop \& dictionary lookups|Performs ERP matching, currency conversions, and variance calculations.|

---

## Technologies

* Python 3.x
* Pandas: Data manipulation, vectorization, and lookups.
* OpenPyXL: Excel workbook manipulation and table/style formatting.
* Power BI: Analytical dashboard \& KPI visualization.

---

## Sample Data Disclaimer
All sample workbooks included in this repository contain dummy data generated exclusively for demonstration purposes. The data does not contain real supplier names, prices, buyers, or business transactions.

---

## Author
Juan Santos
Sourcing Intern | Procurement Automation | Python | Power Automate | Power BI | Data Analytics

