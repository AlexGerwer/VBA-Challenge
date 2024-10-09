# VBA_Challenge
Module 2 Challenge
Process Stock Data
Explanation:
1. ProcessStockData Subroutine:
•	Initialization: Declares variables for workbook, worksheet, last row, ticker, loop counter, unique tickers, various calculations (percent change, quarterly change, total volume), and row numbers for maximum and minimum values.
•	Worksheet Loop: Iterates through each worksheet in the active workbook.
•	Within Each Worksheet:
o	Headers: Writes column headers in columns I through L and O through Q. Autofits these columns to the header width.
o	Finding Last Row: Determines the last row containing data in column A.
o	Data Check: Checks if data exists (lastRow > 1).
o	Unique Tickers: Extracts unique ticker symbols from column A using WorksheetFunction.Unique.
o	Store First Data Cell: Stores the first cell in column J to use later for conditional formatting.
o	Ticker Loop: Iterates through each unique ticker symbol.
	Calculate Values: Calculates quarterlyChange, percentChange, and totalVolume for the current ticker. GetEarliestDate and GetLatestDate helper functions (explained below) are used to retrieve the appropriate row numbers for these calculations. Handles division by zero errors when calculating percentChange.
	Write Results: Writes the ticker, quarterlyChange, percentChange, and totalVolume to the next available row in columns I through L. Formats percentChange as a percentage.
o	Conditional Formatting: Applies conditional formatting to column J (Quarterly Change). Negative values are formatted red, and positive values are formatted green. This formatting is applied after all data is written.
o	Find Max/Min: Finds the row numbers for the maximum and minimum percent change and maximum volume using WorksheetFunction.Match (more efficient than looping). Stores these values. Handles potential errors if the relevant columns are empty.
o	Write Max/Min: Writes the ticker and the maximum/minimum values to cells O2:Q4. Formats the percentage values correctly.
o	Autofit Columns (Again): Autofits columns I:Q again to account for the new data written (especially the max/min section). Includes logic to explicitly check header lengths against column widths and adjust if necessary.
2. Helper Functions:
•	GetEarliestDate(ws, ticker): Returns the row number containing the earliest date for a given ticker symbol in the specified worksheet. It iterates through the rows and keeps track of the smallest date value and its corresponding row number.
•	GetLatestDate(ws, ticker): Returns the row number containing the latest date for a given ticker symbol in the specified worksheet. Similar logic to GetEarliestDate, but it tracks the largest date value.
Data Layout Assumptions:
•	Column A: Stock ticker symbols.
•	Column B: Dates (numeric representation).
•	Column C: Starting values (used for quarterly change calculation).
•	Column F: Ending values (used for percent change calculation).
•	Column G: Stock volume (used for total volume calculation).
•	Columns I:L: Output columns for the calculated metrics.
•	Columns O:Q: Output columns for the greatest increase/decrease in percentage and greatest total volume.
Key Improvements and Error Handling:
•	Efficiency: Uses WorksheetFunction.Match for finding max/min values, which is considerably faster than looping.
•	Error Handling: Includes error handling for division by zero and potential errors when finding max/min values in empty columns.
•	Conditional Formatting: Applies conditional formatting after data is written, ensuring it applies correctly to the entire data range.
•	Clearer Variable Names: Uses more descriptive variable names than in the initial version, improving readability.
•	Autofitting: Implements more robust autofitting logic that considers both data and header widths.
