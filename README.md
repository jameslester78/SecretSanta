# SecretSanta

## Set up your config file

- outputFile - where do you want the output file to reside
- sheetsJson = location for credentials to access Google Sheets
- SheetName = The name of the Google Sheet where your participant data is stored
- WorksheetName = The name of the worksheet where your participant data is stored
- SigName = What name do you want to appear on the signature
- SigEmail = What email do you want to appear on the signature
- SigJobTitle = What job title do you want to appear on the signature
- EmailPickle = The credential file to access your Gmail account
- SendMail = Do you actually want to send emails? True or False

Spreadsheet should contain the following column headers

- name	
- email_address	
- avoid_gifting_to	- a comma seperated list of names
- gift_to

values in avoid_gifting_to and gift_to should exist in the name column

## Set the mode to Dev or Prod

At the top of script.py set the mode to either "Dev" or "Prod"

## Generate Sheets credentials

- https://console.developers.google.com/
- Add Google Sheets and Google Drive API
- Create service account
- Create json key for service account (will be auto downloaded)
- Grant permission on the spreadsheet to the service account that you just created (use email shown on account screen or in key file)
