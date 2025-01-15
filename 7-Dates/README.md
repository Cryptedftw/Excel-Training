Exercise 7: Calculating the Number of Working Days Remaining in a Year
Overview
This exercise focused on applying date functions and calculations in Microsoft Excel to generate project timeline information. The goal was to calculate key date-related metrics, including calendar and working days, for project deadlines.

Purpose
The main objectives of this exercise were to:

Determine the total calendar days between the current date and project deadlines.
Calculate the number of working days, excluding weekends and federal holidays.
Extract and display the month and year from the deadline dates for better reporting and categorization.
Tasks Performed
Calculate Current Date:
Used the TODAY function in cell B1 to dynamically display the current date.
Determine Calendar Days:
Subtracted the current date from project deadline dates to calculate the total number of calendar days remaining.
Calculate Working Days:
Used the NETWORKDAYS function to determine the number of working days, excluding weekends and federal holidays (using a holiday list provided in the worksheet).
Extract Month and Year:
Applied the MONTH and YEAR functions to extract and display the month and year of the project deadlines in separate columns.
Key Formulas Used
Current Date: =TODAY()
Total Calendar Days: =D5-$B$1 (where D5 is the deadline date)
Total Working Days: =NETWORKDAYS($B$1,D5,$J$5:$J$26) (where J5:J26 contains the list of federal holidays)
Extract Month: =MONTH(D5)
Extract Year: =YEAR(D5)
Process Summary
Added the current date dynamically in cell B1 using TODAY.
Calculated the total calendar and working days for each project.
Extracted month and year components from deadline dates for better organization.
Used Autofill to replicate formulas across rows for all project entries.
Verified and validated results by comparing outputs with sample data.
Example Results
Calendar Days: 54 days between the current date (05/09/23) and the deadline date in D5.
Working Days: 37 days, excluding weekends and federal holidays.
Month: Extracted as 7 from a deadline date of 07/01/23.
Year: Extracted as 2023 from a deadline date of 07/01/23.
Files
7-Working-Days-Begin.xlsx: Initial file with raw project and date data.
7-Working-Days-End.xlsx: Final file with calculated calendar and working days, and extracted month/year information.
Tools Used
Microsoft Excel 2021
Date functions: TODAY, NETWORKDAYS, MONTH, YEAR
Data calculations and timeline management techniques.