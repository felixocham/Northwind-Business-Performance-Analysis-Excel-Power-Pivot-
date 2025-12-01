# Northwind Business Performance Analysis (Excel/Power Pivot)
## 1. Project Overview
This project provides a comprehensive, interactive Business Intelligence solution built within Microsoft Excel, leveraging the Power Pivot data model and advanced DAX measures. It utilises the classic Northwind Traders database schema to analyse sales, profit, and operational performance metrics.

The solution is designed to support real-time decision-making through dynamic analysis tables, strategic customer/product segmentation, and a highly interactive, VBA-driven dashboard that toggles between Sales, Orders, and Profit views.
## 2. Data Sources and Model Structure
### 2.1 Data Sources
The project uses tabular data, with each primary entity stored in separate Excel sheets:

*  Transactions (Fact Table)

*  Shipper (Dimension Table)

*  Customer (Dimension Table)

*  Employee (Dimension Table)

*  Product (Dimension Table)

*  Category (Dimension Table)

### 2.2 Data Model (Power Pivot)
The data is loaded into the Excel Data Model, establishing a Star Schema architecture.

**Key Components:**

1. Relationships: Defined between the central Transactions fact table and all dimension tables (Customer, Product, Employee, etc.).

2. Calendar Dimension: A dedicated Calendar dimension table was created.

 *  Fiscal Year Definition: The fiscal year begins on July 1st (01-07).

 *  Time Intelligence: Calculations for Fiscal Year and Fiscal Quarter were added to the Calendar table to support time-based analysis.
  

  <img width="1544" height="961" alt="image" src="https://github.com/user-attachments/assets/0eb386ba-e205-44bb-bc98-71cda1246d1b" />
  
## 3. Core Calculated Measures (DAX)
The following core measures form the foundation of the analysis, providing key performance indicators (KPIs) and enabling year-over-year comparisons.
| Measure Name    | Description                            | Time Intelligence                      |
| :---            | :---                                   | :---:                                  |
| Total Sales     | Sum of transaction sales amount        | Total Sales LY (Same period Last Year) |
| Total Orders    | Distinct count of orders               | Total Orders LY                        |
| Total Profit    | Sum [Sales] - [COGs]                   | Total Profit LY                        |
| Total Quantity  | Sum of units sold                      | N/A                                    |
| Total COGs      | Sum of Cost of Goods Sold              | N/A                                    |
| Total Freight   | Sum of Shipping expenses               | N/A                                    |
|Percentage Profit| (Total Sales - Total COGs)/ Total Sales| N/A                                    |

## 4. Advanced Segmentation and Operational Metrics

Custom metrics were developed to classify customers and highlight product stock status, critical for strategic planning.

### 4.1 Customer Segmentation

Customers are classified into three mutually exclusive groups based on sales performance and engagement (lifespan).

| Segment          | Criteria (Sales & Lifespan)                    | DAX Logic                                         |
| :---             | :---                                           | :---                                              |
| High Value       | Sales > 50,000 AND Lifespan > 15 Months        | Conditional IF logic based on calculated metrics. |
| Growth Customers | Sales > 25,000 AND Lifespan > 10 Months        | Conditional IF logic based on calculated metrics. |
| New Customers    | All others who do not meet the above criteria. | Default category.                                 |

**Supporting Customer Metrics (Calculated in Power Pivot):**

*  Average Monthly Orders

*  Average Order Value

*  First Order Date / Last Order Date

*  Customer Lifespan (Months): Calculated from First Order Date to a reference date.
