Exercise 6: Standardizing Text-Based Data
Overview
This exercise focused on cleaning and standardizing text-based data in an Excel worksheet using various text functions. The goal was to enhance the readability and consistency of the data, ensuring it was ready for further analysis.

Purpose
The purpose of this exercise was to:

Remove unnecessary spaces in text entries.
Adjust the casing of text to meet standard formatting requirements.
Extract and combine portions of text entries for improved data organization.
Use Excel’s text functions to create polished and clean data for presentation or analysis.
Tasks Performed
Removed Extra Spaces:
Used the TRIM function to eliminate unnecessary spaces from text entries while keeping spaces between words intact.
Changed Text Case:
Used the PROPER function to capitalize the first letter of each word in city names.
Extracted Text Portions:
Used the LEFT, RIGHT, and MID functions to extract specific portions of text strings from cells.
Combined Text Entries:
Used the CONCAT function to join text entries from different columns while ensuring proper spacing between words.
Converted Text to Uppercase:
Used the UPPER function to transform mixed-case text into uppercase.
Key Formulas Used
Remove Extra Spaces: =TRIM(B2)
Capitalize Text: =PROPER(D2)
Extract Left Portion: =LEFT(H2,6)
Extract Right Portion: =RIGHT(H2,8)
Extract Middle Portion: =MID(H2,8,3)
Combine Text: =CONCAT(G2, " ", I2)
Convert to Uppercase: =UPPER(L2)
Example Results
Removed spaces from " The Bicycle Accessories Company" to "The Bicycle Accessories Company."
Capitalized "ALHAMBRA" to "Alhambra."
Extracted "States" from a text string using LEFT.
Extracted "New York" from a text string using RIGHT.
Combined "United" and "States" into "United States."
Converted "uSA" to "USA."
Process Summary
Used text functions to clean and reformat the data.
Applied the Autofill feature to extend formulas across rows.
Replaced formulas with values to finalize the cleaned data.
Deleted unnecessary columns to achieve a polished worksheet.
Files
6-Standardizing-Text-Begin.xlsx: Initial file with raw, unformatted text data.
6-Standardizing-Text-End.xlsx: Final file with cleaned and standardized text data.
Tools Used
Microsoft Excel 2021
Text functions: TRIM, PROPER, LEFT, RIGHT, MID, CONCAT, UPPER
Formatting for presentation-ready data.