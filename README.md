# Excel-coffee-dashboard-project
-----

# ☕ Coffee Sales Dashboard in Excel

This project walks you through the end-to-end creation of a dynamic and interactive coffee sales dashboard using Microsoft Excel. From gathering raw data spread across multiple tables to transforming, cleaning, and visualizing it with pivot tables and charts, this dashboard provides key insights into coffee sales performance.

-----

## ✨ Features

The final dashboard provides interactive elements to explore coffee sales data:

  * **Total Sales Over Time Line Chart:** Visualizes sales trends, segmented by coffee type (Arabica, Excelsa, Liberica, Robusta).
  * **Sales by Country Bar Chart:** Shows sales distribution across different countries (U.S., Ireland, UK).
  * **Top 5 Customers Bar Chart:** Highlights your most valuable customers.
  * **Interactive Timeline:** Filter all visuals by date periods (years and months).
  * **Slicers:**
      * **Roast Type:** Filter by Dark, Light, or Medium roasts.
      * **Size:** Filter by coffee package size (e.g., 0.2 kilo, 0.5 kilo, 1 kilo, 2.5 kilo).
      * **Loyalty Card:** Filter customers based on whether they have a loyalty card.

-----

## 🛠️ Project Steps

Follow these steps to recreate the dashboard:

### 1\. Data Acquisition & Transformation

The initial dataset (Orders) contains `Order ID`, `Order Date`, `Customer ID`, `Product ID`, and `Quantity`. Additional details for customers and products are sourced from separate `Customers` and `Products` tables.

  * **Gathering Customer Data (XLOOKUP):**
    Use `XLOOKUP` to bring `Customer Name`, `Email`, and `Country` from the `Customers` table into your `Orders` table.

      * **Customer Name (e.g., in `F2`):**
        ```excel
        =XLOOKUP(C2,Customers!A:A,Customers!B:B,,0)
        ```
      * **Email (e.g., in `G2` - handles blanks):**
        ```excel
        =IF(XLOOKUP(C2,Customers!A:A,Customers!C:C,,0)=0,"",XLOOKUP(C2,Customers!A:A,Customers!C:C,,0))
        ```
      * **Country (e.g., in `H2`):**
        ```excel
        =XLOOKUP(C2,Customers!A:A,Customers!G:G,,0)
        ```

  * **Gathering Product Data (INDEX/MATCH):**
    Use the dynamic `INDEX/MATCH` combination to pull `Coffee Type`, `Roast Type`, `Size`, and `Unit Price` from the `Products` table efficiently. This single formula can be dragged across columns and down rows.

      * **Formula (e.g., in `I2` for "Coffee Type"):**
        ```excel
        =INDEX(Products!$A$1:$F$245,MATCH($D2,Products!$A:$A,0),MATCH(I$1,Products!$1:$1,0))
        ```
          * `$D2`: Locks the `Product ID` column for row matching.
          * `I$1`: Locks the header row for column matching.

  * **Calculating Sales:**
    Add a `Sales` column, calculating `Sales` as `Unit Price * Quantity`.

      * **Formula (e.g., in `M2`):**
        ```excel
        =L2*E2
        ```

  * **Creating User-Friendly Names:**
    Add new columns for `Coffee Type Name` and `Roast Type Name` to convert abbreviations into full names for better readability in the dashboard.

      * **Coffee Type Name (e.g., in `N2`):**
        ```excel
        =IF(I2="ROB","Robusta",IF(I2="EXE","Excelsa",IF(I2="ARA","Arabica",IF(I2="LIB","Liberica",""))))
        ```
      * **Roast Type Name (e.g., in `O2`):**
        ```excel
        =IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))
        ```

### 2\. Data Cleaning & Formatting

Prepare your data for optimal analysis and presentation.

  * **Format Order Date:**
    Select the `Order Date` column, press `Ctrl + 1`, choose **Custom**, and use the format `dd-mmm-yyyy` to display dates clearly (e.g., `05-Sep-2025`).
  * **Format Size:**
    Select the `Size` column, press `Ctrl + 1`, choose **Custom**, and use `0.0 "kilo"` to show units (e.g., `1.0 kilo`).
  * **Format Unit Price & Sales:**
    Select these columns, go to the `Home` tab \> `Number` group, choose **Currency**, and set to **US Dollars** with `0` decimal places and `Use 1000 Separator`.
  * **Remove Duplicates:**
    Select your entire data range (`Ctrl + A`), go to `Data` tab \> `Data Tools` \> `Remove Duplicates`, and confirm to ensure no redundant records.
  * **Convert to Excel Table:**
    Click anywhere in your data, press `Ctrl + T`, confirm headers, and click `OK`. **Rename** the table (e.g., `Orders_Table`) in the `Table Design` tab for easier referencing and automatic expansion.

### 3\. Dashboard Creation (Pivot Tables & Charts)

Now, let's build the interactive components of the dashboard.

  * **Insert Pivot Table:**
    Select any cell within your `Orders_Table`, go to `Insert` tab \> `PivotTable` or use the shortcut **Alt + N + V + T**, then `Enter`. Place it on a `New Worksheet` (rename it e.g., "Total Sales").

  * **Configure "Total Sales Over Time" Pivot Table:**

    1.  **Rename Pivot Table:** In `PivotTable Analyze` tab, rename the pivot table (e.g., `TotalSales_PT`).
    2.  **Fields:**
          * Drag `Order Date` to **Rows**.
          * **Group Dates:** Right-click a date cell in the pivot table \> `Group` \> select `Years` and `Months`.
          * Drag `Coffee Type Name` to **Columns**.
          * Drag `Sales` to **Values**.
          * **Format Sales:** Right-click a sales value \> `Value Field Settings` \> `Number Format` \> `Number` (0 decimals, 1000 separator).
    3.  **Layout:** In `Design` tab \> `Report Layout` \> `Show in Tabular Form`. Remove `Grand Totals` and `Subtotals`.

  * **Create & Format "Total Sales Over Time" Line Chart:**

    1.  **Insert Chart:** Select your pivot table. Go to `Insert` tab \> `Charts` group \> select a **Line Chart**.
    2.  **Clean Up:** Right-click on any field button on the chart \> `Hide All Field Buttons on Chart`.
    3.  **Styling:**
          * **Chart Area Fill:** Double-click chart area \> `Format Chart Area` \> `Solid fill` (e.g., RGB 60, 20, 100 for purple).
          * **Font Color:** Change all chart text/font colors to match the chart area (e.g., purple or white for contrast).
          * **Axis Lines:** Format axis lines (e.g., to white for visibility).
          * **Add Axis Title:** `Design` tab \> `Add Chart Element` \> `Axis Titles` \> `Primary Vertical` (type "USD").
          * **Add Chart Title:** `Design` tab \> `Add Chart Element` \> `Chart Title` \> `Above Chart` (type "Total Sales Over Time").
          * **Line Colors:** Click each data series line individually and change its color for differentiation (e.g., yellow, brown, blue, red).

  * **Insert & Customize Timeline Slicer:**

    1.  **Insert Timeline:** Select the pivot table or chart. Go to `PivotChart Analyze` tab \> `Filter` group \> `Insert Timeline`. Select `Order Date`.
    2.  **Custom Style:**
          * Select the timeline. Go to `Timeline Tools` \> `Options` tab \> `Styles` group \> `New Timeline Style`.
          * **Name:** "Purple Timeline Style".
          * **Format Elements:** Customize various elements (e.g., `Whole Timeline`, `Header`, `Selection Label`, `Time Level`, `Period Labels`, `Selected Time Block`, `Unselected Time Block`) to match your dashboard's color scheme (e.g., purple fills, white text/borders).

-----

## ▶️ Usage

Once all components are set up, use the timeline and slicers to dynamically filter and analyze your sales data. Select different time periods, coffee roast types, sizes, or loyalty card statuses to observe how sales trends and distributions change across your charts.

-----

## 💻 Technologies Used

  * **Microsoft Excel** (preferably Excel 365 or a version supporting `XLOOKUP`).

-----
