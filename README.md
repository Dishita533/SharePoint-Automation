# SharePoint-Automation (For Migration Purpose)
This is an automation tool to compare LHN, counts, settings of migrated SPO sites with respective OnPrem sites. It run completely based on Powershell script where version 5 for OnPrem connection and version 7 required for SPO connection.

By default PowerShell 5 will be present in the System, Install PowerShell 7 before proceeding further

Before running the code
PowerShell 7 (Go to code  folder path) - Run -> Install-Module PnP.PowerShell -Scope CurrentUser
PowerShell 5 (Go to code folder path) - Run -> Install-Module SharePointPnPPowerShell2019 -Scope CurrentUser

Then run that specific folder's "Compare-xxxxx.ps1" file as example " .\Compare-Main.ps1 " in Powershell 7 under that specific folder.

OUTPUT FILE:
csv file (Output file) - present at "C:\Dishita\Reports\SPCompare" --> If there is no such folder, it will be created for the first time and will be reused to append for all other sites.

IN THE CODE - Add your own 
1. SPO and OnPrem URL
2. Client ID for SPO site connection
3. OnPrem Username and Password Credentials (All these are present in the "Compare-Main.ps1" file and similar compare files)

Purpose is to Reduce Manual effort and time by using this Automation tool
