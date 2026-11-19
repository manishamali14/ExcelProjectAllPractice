# MS Excel Practical Assignment (Superstore Sales)

This is my Excel practical assignment from the Jobaaj program. I worked through 12 modules of Excel, starting with basic formulas and going all the way to Pivot Tables, lookups, report consolidation and VBA macros. Every task from the assignment brief is answered in the workbook, and I wrote the task above each answer so it is easy to follow.

## Dataset

I used the Sample Superstore dataset (Tableau community). It has 9,994 orders from 2014 to 2017 with details like order date, region, category, sales, quantity, discount and profit. Total sales come to $2,297,200.86. I loaded the data with Power Query so it can be refreshed.

## What is inside

The workbook has one sheet for each section:

1. Basics of Excel: SUM, MAX, MIN, SMALL, LARGE, MEDIAN, TRANSPOSE, COUNT functions, status bar and YoY growth with CAGR
2. Formatting: currency and custom number formats, conditional formatting, wrap text, alignment and Format Painter
3. Sort and Filter: sorting numbers, dates and text, sorting by colour and multi level sorting
4. Dates: durations with timestamps, DATEDIF and date functions
5. Logical functions: IF, AND, OR and IFERROR
6. Data cleaning: removing duplicates and errors, text functions, SEARCH vs FIND, Go To Special, Text to Columns, CONCATENATE and Find & Replace
7. Pivot Tables: grouping by date, handling blanks, custom groups and conditional formatting
8. Lookups: VLOOKUP (exact and approximate match), named ranges and HLOOKUP
9. Conditional aggregation: COUNTIF, SUMIF, AVERAGEIF, COUNTIFS and SUMIFS
10. Conditional formatting: highlight rules, colour scales, data bars and icon sets
11. Report consolidation: Consolidate across monthly and regional sheets, and SUBTOTAL with filters
12. Macros: recording, editing and running a macro (button, shortcut and on open)

## Data cleaning

The Orders data had no missing values, so I did not need to fill anything in. I fixed 449 postal codes that had lost their leading zero, cleaned hidden spaces from 273 product names and replaced a wrong revenue column with unit price. I did not delete any rows. One duplicate record and 32 product IDs with two names were flagged and kept, since I could not be sure they were errors.

## A few findings

2017 was the best year with $733,215 in sales, and November was the strongest month. Sales grew at about 14.8% a year between 2014 and 2017. Technology brings in the most sales, while Tables sell well but run at a loss because of heavy discounts.

## How to open

Download MSExcelProject1.xlsm and open it in Excel on Windows. Click Enable Content when Excel asks, otherwise the macros will not run. Each sheet named Sec_1 to Sec_12 matches one section of the assignment.

## Note

The YoY task asks for 5 years of revenue, but the dataset only covers 4 years (2014 to 2017), so I used all 4 instead of making up a year.

## Author

Manisha Mali

LinkedIn: [www.linkedin.com/in/manisha-mali-970b44129](https://www.linkedin.com/in/manisha-mali-970b44129)
