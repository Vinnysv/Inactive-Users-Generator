# Inactive-Users-Generator

This PowerShell script sorts users based on their last login date across multiple services (Teams, SharePoint, OneDrive, and Exchange) by analyzing a given input spreadsheet. The script will output an Excel file with users who have not logged in to any of the services since the specified date.

## Requirements

- PowerShell (Windows)
- [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel) module

## Installation

1. Open PowerShell as administrator.
2. Install the ImportExcel module by running the following command:

Install-Module ImportExcel

##Usage
1. Save the script to a local directory on your computer.
2. Open PowerShell as administrator.
3. Navigate to the directory containing the script.
4. Run the script by typing .\script_name.ps1, replacing script_name with the name of the script file. If the script file is in a different directory, provide the full path to the script.
5. When prompted, provide the path to the input spreadsheet (CSV or XLSX format) containing the user data.
6. The script will process the data and generate an output Excel file in the same directory as the script, named Users(YYYY-MM-DD).xlsx, where YYYY-MM-DD is the current date.

#Input Spreadsheet Format
The input spreadsheet should have the following column headers:

Display Name
Teams Last Activity Date
SharePoint Last Activity Date
OneDrive Last Activity Date
Exchange Last Activity Date
Output

The output file will contain the following columns:

User
Last Teams Login Date
Last OneDrive Login Date
Last Exchange Login Date
Last Sharepoint Login Date

The output file will only contain users who have not logged in to any of the services since the specified date.
