# Asana API Call as a Function in Google Sheets

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



