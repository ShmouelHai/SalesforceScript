# Script 3: Update Excel with Missing Elements

## Overview

This Python script is designed to update an existing Excel file generated from the metadata audit script. It takes an Excel file and a Salesforce org's package XML, identifies missing metadata members, and adds them to the Excel file with a status of "added."

## Prerequisites

- Python 3.x
- pandas library (`pip install pandas`)
- openpyxl library (`pip install openpyxl`)

## Usage

1. **Enter Excel File Path:**
   Enter the path of the Excel file when prompted.

2. **Select the Sheet with Most Rows:**
   The script automatically identifies the sheet with the most rows in the Excel file.

3. **Enter XML Package Path:**
   Enter the path of the Salesforce org's package XML when prompted.

4. **Run the Script:**
   Execute the script by running the following command in your terminal:

   ```bash
   python3 update_excel_with_xml_data.py
   ```

5. **Review Console Output:**
   The script will output information about the added elements and the success of the update.

## Configuration

- The script includes a list of metadata members to exclude (`membersToRemove`). Modify this list based on your specific requirements.

## Notes

- The script compares the metadata members in the Excel file with those in the Salesforce org's package XML.
- For each metadata member in the XML package not present in the Excel file, the script adds a new row with a "added" status.


---

## Author

Shmouel Illouz
