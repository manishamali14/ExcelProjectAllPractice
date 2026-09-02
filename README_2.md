# MS Excel Practice Project

A single macro-enabled Excel workbook (`MSExcelProject1.xlsm`) built as a hands-on practice project covering core-to-intermediate Excel skills — formulas, formatting, sorting/filtering, dates, logical functions, data cleaning, pivot tables, lookups, conditional aggregation, conditional formatting, report consolidation, and VBA macros — all applied to a Superstore-style sales dataset.

## Contents

The workbook has 26 sheets. The `Sec_*` sheets are the practice exercises, each focused on one skill; the remaining sheets hold the source data and supporting tables used throughout the exercises.

### Skill sections

| Sheet | Topic |
|---|---|
| `Sec_1` | Basic arithmetic operations — `SUM`, `MAX`, `MIN`, `SMALL`, `LARGE`, `MEDIAN` |
| `Sec_2Format` | Ways to format cells |
| `Sec_3Sortnfilter` | Sorting numbers and dates in different ways |
| `Sec_4Dates` | Working with dates |
| `Sec_5Logical` | Logical functions |
| `Sec_6DataCleaning` | Practical data cleaning |
| `Sec_7Pivot` | Pivot tables A–Z (workbook includes 37 pivot tables in total) |
| `Sec_8Lookups` | Lookup formulas (`VLOOKUP`, `INDEX`/`MATCH`, etc.) |
| `Sec_9ConditionalIFS` | Conditional aggregation — `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| `Sec_10ConditionalFormatting` | Conditional formatting |
| `Sec_11Report` | Report consolidation |
| `Sec_112` | Excel's Consolidate feature |
| `Sec_12` | Basics of macros (VBA) |

### Data sheets

| Sheet | Description |
|---|---|
| `Orders` | The main dataset — ~9,994 Superstore-style sales orders (order/ship dates, customer, segment, region, product, sales, profit, etc.) |
| `Orders$FilterDatabase` | Filtered view of the Orders table used for the sort/filter exercises |
| `Returns` | Returned order IDs |
| `People` | Region-to-person (sales rep) mapping |
| `WestRegion`, `South`, `East` | Region-specific breakouts of the order data |
| `Jan_data`, `Feb_Data`, `March_Data` | Monthly data used in the report consolidation exercises |

Two named ranges (`Category`, `Product_table`) and several Excel Tables back the lookup and pivot exercises. A handful of additional sheets are kept hidden as scratch/helper sheets for some of the pivot and lookup examples.

### Macros (VBA)

The workbook is macro-enabled (`.xlsm`) and includes a few small VBA routines used in the `Sec_12` macros exercise:

- **`Format_Table`** — applies borders, fill, and column formatting to a table; runs automatically on `Workbook_Open`.
- **`QuickTotal`** — sums the Sales column of the `Orders` sheet and displays the total in a message box.
- **`FlagLowMargin`** — highlights any order row in `Orders` with a negative profit value.

## Getting started

1. Download `MSExcelProject1.xlsm`.
2. Open it in Microsoft Excel (or another spreadsheet app that supports `.xlsm`, such as LibreOffice Calc).
3. When prompted, **enable macros/content** to see the VBA-driven formatting and to run the macros in `Sec_12`.

> **Note:** This file contains VBA macros. Only enable macros in files you trust — GitHub does not scan or execute the macro code, it's just hosting the file.

## Requirements

- Microsoft Excel 2016 or later (recommended, for full pivot table and formatting fidelity)
- Macros enabled to use the `Sec_12` VBA examples

## Author

Manisha Mali
