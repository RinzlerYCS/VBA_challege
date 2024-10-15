# VBA_challege

# Excel VBA Script for Quarterly Stock Data Analysis

## Description

This Excel VBA script is designed to automate the analysis of quarterly stock data for multiple sheets in an Excel workbook. The script processes each sheet in the workbook, calculating the quarterly changes, percent changes, and total stock volumes for tickers, and highlights the largest increase, largest decrease, and greatest total volume for each sheet.

## Features

- Processes each sheet in the workbook (`Q1`, `Q2`, `Q3`, and `Q4`) automatically.
- Calculates:
  - **Quarterly change**: The difference between the open and close prices.
  - **Percent change**: The percentage change based on the open price.
  - **Total stock volume**: The total volume of stocks traded.
- Highlights positive changes in green and negative changes in red.
- Identifies and displays:
  - The ticker with the greatest percentage increase.
  - The ticker with the greatest percentage decrease.
  - The ticker with the greatest total stock volume.
