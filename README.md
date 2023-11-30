# VBA_Chalenge
Functionality
AssignmentForAllSheets: This subroutine iterates through all the worksheets in an Excel workbook. It calls the Assignment2 subroutine for each sheet, ensuring comprehensive data analysis across multiple datasets.

Assignment2: This subroutine, which takes a worksheet as an argument, performs detailed stock analysis. Key functionalities include:

Calculation of Yearly Change: For each stock (ticker), the script calculates the yearly change in value.
Percentage Change Computation: It computes the percentage change in stock value throughout the year.
Total Stock Volume Assessment: The script aggregates the total stock volume for each ticker.
Identification of Extremes: The subroutine identifies stocks with the greatest percentage increase, greatest percentage decrease, and the greatest total volume.
Data Visualization Enhancement: Conditional formatting is applied to highlight positive or negative yearly changes in stock values.
Implementation Details
The script initializes necessary variables for tracking and comparison purposes (like MaxIncrease, MaxDecrease, and Total_Volume).
A loop processes each row in a worksheet, calculating the required values and checking against the established maxima and minima.
The final results, including the identified extremes, are outputted to specific cells within the worksheet.
Usage
Place the VBA script in your Excel environment.
Ensure your workbook contains sheets with stock data in a compatible format.
Run AssignmentForAllSheets to execute the analysis across all sheets.
Review the outputs in the respective columns and the summary of extremes at the specified range.
Note
This script is tailored for specific data layouts (tickers, yearly opening and closing values, etc.). Ensure your data aligns with these expectations for optimal performance.
The script can handle large datasets across multiple sheets, making it suitable for extensive stock market analysis.
