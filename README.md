# Python-Excel-Workbook-Unique-Sheet-Generator

This Python script uses the openpyxl library to automate the process of creating new sheets in an Excel workbook based on unique values in a specified column. This is particularly useful for organizing and segregating data for better readability and management.

Features
Load an Existing Workbook: The script opens an existing Excel workbook.
Identify Unique Values: It identifies unique items from a specified column in a given sheet.
Sanitize Sheet Names: It ensures that new sheet names are valid by removing invalid characters and limiting the length to 31 characters.
Create New Sheets: For each unique item, it creates a new sheet (or uses an existing one if already created).
Copy Data: It copies the header row and the rows corresponding to each unique item into the respective new sheets.
Save the Workbook: Finally, it saves the updated workbook with the newly created sheets.
