# Googlesheets, Asana & Salesforce integrated reporting 

## 1. Google Sheets Custom Function to make an Asana API call 

![image](https://user-images.githubusercontent.com/9472095/87558145-aa2f6d00-c6b0-11ea-899c-4156f2d15ef7.png)

Quick and cheesy fix to automate bringing Asana Updates into a Google Sheets Cell using a gscript function

## Prerequisites
- Get an Asana Personal Access Token (PAT) [here](https://developers.asana.com/docs/authentication).
- Create a new Google Sheets Doc
- Select Tools->Script Editor from the menu
- Replace the entire placeholder script with the contents of [Asana_Project_Update.gs](https://github.com/allthingsclowd/GoogleSheetCellWithAsanaFunction/blob/master/Asana_Project_Update.gs)
- Configure your PAT token retrieved earlier : `var token = "INSERT ASANA BEARER TOKEN HERE";`


## What does the script do?
It creates a new custom googlesheets FUNCTION called `GET_ASANA_UPDATE(CELL)`
The input should contain the Asana Project ID - `gid` which is usually displayed in the browser URL when viewing the project through a browser.

The function then reads the latest progress update for the given project id from Asana and drops it inplace of the function.

## 2. Google Sheets UI script to retrieve all the updates from a portfolio of projects in Asana
file - `Retrieve_Latest_Update_From_Protfolio.gs`

## 3. Google Sheets UI script to retrieve custom fields of interest from a portfolio of projects in Asana
file - `retrieve_custom_fields_from_projects.gs`

## 4. Script to read in portfolio data and product a formatted Summary & PIE Chart
file - `Eligible_Accounts_AllInOne.gs`
![image](https://user-images.githubusercontent.com/9472095/88834739-2c00b980-d1cc-11ea-8c6a-580e44b86592.png)

## 5. Putting it all together in a single automated spreadsheet
file - `Weekly_Reports.gs` this is extremely untidy and not recommended for anyone really - this is a work in progress
![image](https://user-images.githubusercontent.com/9472095/91070963-76d8ea00-e62f-11ea-8190-f142b5e84ea3.png)

