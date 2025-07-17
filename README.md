# Analysis-short-formats
These are just scripts ran internally to make repetitive work done using excel much easier.


Night Driving Report Cleanup & Summary Script
Overview
This Python script processes a monthly night driving report exported from MiX Telematics. It prepares the data for analysis by cleaning, standardizing, and summarizing vehicle night driving occurrences.

What the Script Does
1. Load and Clean the Excel File
Opens the raw Excel file containing the detailed event report.

Removes embedded images (e.g., logos).

Unmerges merged cells to ensure data is structured properly.

Clears all formatting (fonts, colors, borders, etc.).

Applies proper number formats to:

Start Date (Column F): yyyy-mm-dd

Start Time (Column G): hh:mm:ss

End Time (Column H): hh:mm:ss

Deletes the first 6 rows, which typically contain headers, merged cells, or branding info.

2. Save a Cleaned Version
The cleaned version of the file is saved separately to preserve the original.

3. Process with pandas (DataFrame Logic)
Loads the cleaned Excel data into a DataFrame (structured table).

Converts Start Date to a date-only format (removes time).

Removes duplicate entries, keeping only the first night driving occurrence per vehicle per day.

Sorts the data by date and registration number.

4. Generate Monthly Summary
Groups the data by Registration Number.

Counts how many unique days each vehicle was involved in night driving during the month.

Saves this as a summary report in Excel.

Output Files
Night driving occurrences morl.xlsx
→ Cleaned, deduplicated record of night driving events (one row per vehicle per day).

Monthly Occurrence Summary - morl.xlsx
→ Summary table showing how many days each vehicle drove at night.

🛠 Requirements
This script uses:

openpyxl – to manipulate Excel formatting and structure

pandas – to handle data cleaning, grouping, and summarizing

✅ Benefits
Removes redundant or cluttered rows and formatting.

Ensures consistency in date and time formats.

Reduces noise by keeping only one entry per vehicle per day.

Provides a clear monthly summary for reporting and decision-making.


