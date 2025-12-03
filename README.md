# Northwind Business Performance Analysis (Excel/Power Pivot)
## <img width="30" height="30" alt="18275319" src="https://github.com/user-attachments/assets/dffc81ec-3e0c-4127-bc54-84bd5962e2e9" />  1. Project Overview 

This project provides a comprehensive, interactive Business Intelligence solution built within Microsoft Excel, leveraging the Power Pivot data model and advanced DAX measures. It utilises the classic Northwind Traders database schema to analyse sales, profit, and operational performance metrics.

The solution is designed to support real-time decision-making through dynamic analysis tables, strategic customer/product segmentation, and a highly interactive, VBA-driven dashboard that toggles between Sales, Orders, and Profit views.

## <img width="30" height="30" alt="2833509" src="https://github.com/user-attachments/assets/6ed132fa-e298-40e4-9ddb-589c7834625c" /> 2. Prerequisites and Dependencies 


To interact fully with the Power Pivot model, analyse the data, and utilise the interactive dashboard functionality, the following prerequisites are required:

### 2.1. Software Requirements

*  Microsoft Excel (2016 or newer / Microsoft 365): The file relies heavily on native Excel features.

*  Power Pivot Add-in: Must be enabled within Excel to view and modify the Data Model and DAX measures.

* VBA (Visual Basic for Applications): Must be enabled. The dashboard's interactive view switching (Sales, Orders, Profit buttons) and sheet navigation depend entirely on custom VBA macros. Macro security settings must permit running macros.

### 2.2. Skill Requirements for Maintenance

*  DAX (Data Analysis Expressions): Necessary for modifying, troubleshooting, or enhancing the calculated measures (KPIs, time intelligence, segmentation logic).

*  Power Pivot Data Modelling: Required for understanding and adjusting table relationships, column properties, and the Calendar table logic.

*  VBA: Required for maintaining or expanding the dashboard's interactive button logic and navigation macros.

*  Excel Pivot Tables/Charts: Proficiency is needed to adjust the presentation of the analysis tables and charts.

## <img width="30" height="30" alt="3979345" src="https://github.com/user-attachments/assets/512f38d1-f6d2-4b30-ab47-a1528bc15fd0" /> 3. Data Sources and Model Structure 

### 3.1 Data Sources
The project uses tabular data, with each primary entity stored in separate Excel sheets:

*  Transactions (Fact Table)

*  Shipper (Dimension Table)

*  Customer (Dimension Table)

*  Employee (Dimension Table)

*  Product (Dimension Table)

*  Category (Dimension Table)

### 3.2 Data Model (Power Pivot)
The data is loaded into the Excel Data Model, establishing a Star Schema architecture.

**Key Components:**

1. Relationships: Defined between the central Transactions fact table and all dimension tables (Customer, Product, Employee, etc.).

2. Calendar Dimension: A dedicated Calendar dimension table was created.

 *  Fiscal Year Definition: The fiscal year begins on July 1st (01-07).

 *  Time Intelligence: Calculations for Fiscal Year and Fiscal Quarter were added to the Calendar table to support time-based analysis.
  

  <img width="1544" height="961" alt="image" src="https://github.com/user-attachments/assets/0eb386ba-e205-44bb-bc98-71cda1246d1b" />
  
## <img width="30" height="30" alt="7870594" src="https://github.com/user-attachments/assets/6e4cdf58-f2da-46e6-b48b-7162edff9e76" /> 4. Core Calculated Measures (DAX) 

The following core measures form the foundation of the analysis, providing key performance indicators (KPIs) and enabling year-over-year comparisons.
| Measure Name    | Description                            | Time Intelligence                      |
| :---            | :---                                   | :---                                   |
| Total Sales     | Sum of transaction sales amount        | Total Sales LY (Same period Last Year) |
| Total Orders    | Distinct count of orders               | Total Orders LY                        |
| Total Profit    | Sum [Sales] - [COGs]                   | Total Profit LY                        |
| Total Quantity  | Sum of units sold                      | N/A                                    |
| Total COGs      | Sum of Cost of Goods Sold              | N/A                                    |
| Total Freight   | Sum of Shipping expenses               | N/A                                    |
|Percentage Profit| (Total Sales - Total COGs)/ Total Sales| N/A                                    |

## <img width="30" height="30" alt="8922128" src="https://github.com/user-attachments/assets/a65db67b-4e93-4c3c-93b6-5d29e4a61d66" /> 5. Advanced Segmentation and Operational Metrics 


Custom metrics were developed to classify customers and highlight product stock status, critical for strategic planning.

### 5.1 Customer Segmentation

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

### 5.2 Product Stock Status
A critical measure was created to assess the inventory health of each product by comparing stock levels against average consumption.

 | Measure Name | Inputs | Logic |
 | :---         | :---   | :---  |
 | Stock Highlight | Units in Stock, Reorder Level, Units on Order, Average Monthly Sales Quantity | Conditional logic to determine status (e.g., 'Immediate Reorder', 'Stock Healthy', 'Order in Progress'). |

 ## <img width="30" height="30" alt="10857082" src="https://github.com/user-attachments/assets/8b07a7e7-b332-473e-a7d4-2dc943fbfba3" /> 6. Analysis Worksheet (Analysis) 

 This worksheet contains a collection of interactive pivot tables designed for detailed slice-and-dice analysis.

 <img width="2924" height="1822" alt="image" src="https://github.com/user-attachments/assets/000b6f9f-4db7-4a63-ac29-c08b2c38c396" />


### 6.1. Key Pivot Tables

The following analysis tables are available, each utilising the core calculated measures (Total Sales, Total Profit, etc.):

*  Analysis by Product Category: Performance of all product groups.

*  Top 10 Products: Ranked list based on Total Sales.

*  Time-Based Analysis: Drilldown functionality on rows for Fiscal Year > Fiscal Quarter > Month.

*  Geographic Analysis: Performance broken down by Employee/Sales Office location.

*  Customer Performance: Top 10 Customers based on Sales.

*  Specialised Customer Analysis: Uses advanced customer metrics (AOV, Lifespan, First/Last Order Date) against the defined Customer Segmentation.

*  Specialised Product Analysis: Displays the Stock Highlight measure alongside operational metrics (Avg. Monthly Sales Qty, Last Order Date, Number of Orders).
  <img width="684" height="301" alt="image" src="https://github.com/user-attachments/assets/4cd3bc63-9069-46c4-bf10-d8aa1071c77f" />


### 6.2. Interactivity and Formatting

*  Slicers: Added for filtering by Region, Fiscal Year, and Employee Office.

*  Visual Aids: Conditional Data Bars and Conditional Formatting rules are applied across all tables for immediate visual identification of high/low performance, especially for the Stock Highlight measure.
  <img width="822" height="292" alt="image" src="https://github.com/user-attachments/assets/ae4d3d87-9f29-4a7d-98f4-b14b476b8c90" />

## <img width="30" height="30" alt="10989830" src="https://github.com/user-attachments/assets/4a6b9acf-dfd6-462a-8abc-33afb78c9a32" /> 7. Dashboard Worksheet (Dashboard) 


The dashboard is the user-facing summary that provides key insights and interactive visualisations.

<img width="1774" height="1218" alt="image" src="https://github.com/user-attachments/assets/4abdcd4e-739e-4ad9-8045-d30fb4835261" />



<img width="1773" height="1220" alt="image" src="https://github.com/user-attachments/assets/98200814-0f12-4574-bb1e-d8f946c21cb8" />

### 7.1. Key Performance Indicators (KPIs)

The following KPIs are prominently displayed:

*  Total Sales

*  Total Profit (Calculated as Sales - COGs)

*  Number of Orders (Nr. Orders)

*  Total Quantity

*  Total COGs

* Total Freight Cost

**Visual Context:** Each KPI includes a Sparkline that shows the monthly trend for the selected period.

### 7.2. VBA-Driven Interactive Views

The core functionality is an interactive view switcher built with **VBA (Visual Basic for Applications)**. Buttons allow users to instantly switch the displayed charts based on the analytic focus:

| View Button        | Chart Set Displayed  |
| :---               | :---                 |
| Sales              | Sales by Category, Top Products by Sales, Sales by Region, Top 10 Customers by Sales, Top 5 Countries by Sales, Top 10 Cities by Sales, Sales by Sales Rep, Monthly Sales Trend (Previous vs. Current Year), Total Freight Cost. |
| Orders             | Corresponding charts visualising Total Orders (e.g., Orders by Category, Top Products by Orders, etc.). |
| Profits            | Corresponding charts visualising Total Profit (e.g., Profit by Category, Top Customers by Profit, etc.). |

### 7.3. Navigation

*  **Dashboard Button:** A VBA button is placed on the Analysis sheet for quick navigation back to the Dashboard.

*  **Analysis Button:** A VBA button is placed on the Dashboard sheet for quick navigation to the Analysis sheet.

 ## <img width="35" height="35" alt="7891893" src="https://github.com/user-attachments/assets/a8dbab38-b7ef-4b4d-9598-813e97417353" /> 8. Future Roadmap & Improvements 

To ensure the longevity and scalability of this BI solution, the following enhancements are proposed for future iterations:

* **Migration to Power BI:** Transitioning the current Excel-based model to Microsoft Power BI. This will enable:

  1.  Cloud-based sharing and collaboration via Power BI Service.

  2.  Mobile accessibility for stakeholders.

  3.  Implementation of Row-Level Security (RLS) to restrict data views based on user roles (e.g., Regional Managers only seeing their region's data).

* **Automated Data Pipelines:** Replacing the manual upload of static Excel sheets with direct Power Query connections to live data sources (SQL Server, SharePoint, or APIs) to ensure real-time data freshness without manual intervention.

* **Predictive Analytics:** Introducing forecasting algorithms to predict future sales trends and inventory requirements based on historical patterns.

* **Advanced Customer Churn Analysis:** Developing deeper insights into customer retention to flag "at-risk" customers before they churn, utilising the existing "Last Order Date" metrics more aggressively.

* **Inventory Optimisation Model:** Enhancing the current "Stock Highlight" logic to include lead time variability and seasonal demand spikes for more accurate reorder recommendations.

---
*Created by Felix Ocham  |  [Linkedin Profile](https://www.linkedin.com/in/felix-o-703987a7/)*




