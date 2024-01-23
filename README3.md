# SalesforceScript
some Tools to improve Hyperforce Projects

---


# Script 1: Metadata Audit Script

## Overview

This Python script is designed to perform an audit of metadata within an XML file. It analyzes the provided XML file, extracts relevant information, and generates an Excel report with details on the data types and their status in the context of SFOA (Salesforce Objects Architecture).

## Prerequisites

- Python 3.x
- pandas library (`pip install pandas`)

## Usage

### Close Excel File

Before launching the script, ensure that any existing Excel file generated by the script is closed.

### Copy the path of the package (Package XML)

Right-click in VS Code and "Copy path."

### Run the Script

Execute the script by running the following command in your terminal:

```bash
python3 createExcelFromPackageXML.py
```

### Input XML File

Enter the path of the XML file when prompted.

### Output Excel File

The script will generate an Excel file named `Audit.xlsx` on your desktop. If the file already exists, a new sheet with a timestamp will be added.

### Review Report

Open the generated Excel file to review the audit report. The report includes information about metadata members, their types, and their SFOA status.

## Configuration

- The script includes a list of metadata members to exclude (`membersToRemove`). Modify this list based on your specific requirements.

## Notes

- The status for each metadata member is determined based on predefined rules.
  - `Expected OK`: The metadata member is supported in SFOA.
  - `Not supported in SFOA`: The metadata member is not supported in SFOA.
  - `Data Scope`: Applicable only for `EmailTemplate` type.
  - `Scope To Define`: Applicable for `Report` and `Dashboard` types.

---

# Script 2: Metadata Comparison Script

## Overview

This Python script allows you to compare metadata between two Salesforce orgs by analyzing two XML packages. It identifies the differences in metadata members and generates a detailed Excel report outlining what is missing in the package being compared.

## Prerequisites

- Python 3.x
- pandas library (`pip install pandas`)

## Usage

1. **Close Excel File:**
   Before launching the script, ensure that any existing Excel file generated by the script is closed.

2. **Package to be Compared:**
   The package to be compared should have more metadata than the comparing package, and the script checks for missing components.

3. **Run the Script:**
   Execute the script by running the following command in your terminal:

   ```bash
   python3 ComparePackagesXML.py
   ```

   Input Package Paths:
   - Enter the path of the package to be compared (ROW) when prompted.
   - Enter the path of the comparing package (SFOA).

4. **Review Excel Report:**
   Open the generated Excel file named `CompareXML.xlsx` on your desktop. The report provides a comprehensive overview of metadata members, their types, SFOA status, and deployment status.

## Configuration

- The script includes a list of metadata members to exclude (`membersToRemove`). Modify this list based on your specific requirements.

## Notes

- The script compares the SFOA status with the SFOA package to determine deployment status.
  - **Additional Note:**
    - For each metadata member in the comparing package (ROW), the script checks if it is present in the SFOA package. The result is recorded in the "SFOA package" column as either "In Package" or "Not in Package."

- **Deployment Status:**
  - `OK`: The metadata member is expected and in the SFOA package.
  - `NOT OK`: The metadata member is either missing or not supported as per SFOA.

---


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

