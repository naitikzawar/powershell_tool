```


  /■■■■■■  /■■■■■■■■/■■■■■■■                                             /■■■■■■  /■■                 /■■ /■■
 /■■__  ■■| ■■_____/ ■■__  ■■                                           /■■__  ■■| ■■                | ■■| ■■
| ■■  \ ■■| ■■     | ■■  \ ■■/■■■■■■  /■■  /■■  /■■  /■■■■■■   /■■■■■■ | ■■  \__/| ■■■■■■■   /■■■■■■ | ■■| ■■
| ■■  | ■■| ■■■■■  | ■■■■■■■/■■__  ■■| ■■ | ■■ | ■■ /■■__  ■■ /■■__  ■■|  ■■■■■■ | ■■__  ■■ /■■__  ■■| ■■| ■■
| ■■  | ■■| ■■__/  | ■■____/ ■■  \ ■■| ■■ | ■■ | ■■| ■■■■■■■■| ■■  \__/ \____  ■■| ■■  \ ■■| ■■■■■■■■| ■■| ■■
| ■■/■■ ■■| ■■     | ■■    | ■■  | ■■| ■■ | ■■ | ■■| ■■_____/| ■■       /■■  \ ■■| ■■  | ■■| ■■_____/| ■■| ■■
|  ■■■■■■/| ■■     | ■■    |  ■■■■■■/|  ■■■■■/■■■■/|  ■■■■■■■| ■■      |  ■■■■■■/| ■■  | ■■|  ■■■■■■■| ■■| ■■
 \____ ■■■|__/     |__/     \______/  \_____/\___/  \_______/|__/       \______/ |__/  |__/ \_______/|__/|__/
      \__/                                                                                                   


```

# Table of Contents
---
* [Introduction](#introduction)
    * [Features](#features)
    * [Compatibility](#compatibility)
    * [Requirements](#Requirements)
* [Getting Started](#getting-started)
    * [Quick Start](#quick-start)
    * [Advanced Setup](#advanced-setup)
        * [Auto update using Git](#auto-update-using-git)
        * [Automatically loading the module when PowerShell starts](#automatically-loading-the-module-when-powershell-starts)
* [Instructions](#instructions)
    * [New-QFTicket](#new-qfticket)
        * [New-QFTicket - AutoMode](#new-qfticket---automode)
    * [Export-QFExcel](#export-qfexcel)
    * [Format-QFPlayCheck](#format-qfplaycheck)
    * [Get-QFAAMSStatus](#get-qfaamsstatus)
    * [Get-QFAudit](#get-qfaudit)
    * [Get-QFBetSettingProc](#get-qfbetsettingproc)    
    * [Get-QFETIProviderInfo](#get-qfetiproviderinfo)
    * [Get-QFHelixAccessToken](#get-qfhelixaccesstoken)
    * [Get-QFHelixDefaultHeader](#get-qfhelixdefaultheader)
    * [Get-QFHelixIncident](#get-qfhelixincident)
    * [Get-QFHelixIncidentWorkInfo](#get-qfhelixincidentworkinfo)
    * [Get-QFGameBlocking](#get-qfgameblocking)
    * [Get-QFGameStats](#get-qfgamestats)
    * [Get-QFOktaToken](#get-qfoktatoken)
    * [Get-QFOperatorAPIKeys](#get-qfoperatorapikeys)
    * [Get-QFOperatorToken](#get-qfoperatortoken)
    * [Get-QFPlaycheck](#get-qfplaycheck)
    * [Get-QFSQLServerSAPassword](#get-qfsqlserversapassword)
    * [Get-QFUser](#get-qfuser)
    * [Invoke-QFAI](#invoke-qfai)
    * [Invoke-QFAutoTicket](#invoke-qfautoticket)
    * [Invoke-QFMenu](#invoke-qfmenu)
    * [Invoke-QFPortalRequest](#invoke-qfportalrequest)
    * [Invoke-QFReconAPIRequest](#invoke-qfreconapirequest)
    * [New-QFHelixIncident](#new-qfhelixincident)
    * [New-QFHelixIncidentWorkInfo](#new-qfhelixincidentworkinfo)
    * [Search-QFHelixIncident](#search-qfhelixincident)
    * [Search-QFServerDetails](#search-qfserverdetails)
    * [Search-QFUser](#search-qfuser)
    * [Start-QFGame](#start-qfgame)
    * [Update-QFHelixIncident](#update-qfhelixincident)
    * [Update-QFPowerShell](#update-qfpowershell)
* [Known Issues and Troubleshooting](#known-issues-and-troubleshooting)
* [Changelog](#changelog)
* [Future Goals](#future-goals)
* [Contribute](#contribute)
* [Contact and Help](#contact-and-help)



# Introduction 
---
This PowerShell module is intended for MIGS-IT Customer Solutions team members, to automate and streamline some of our common tasks.


## Features
---
* Generate player Transaction and Financial audits via Help Desk Express API, and export into an Excel file with a standard Games Global template.
* Generate multiple Play check and Game Statistics reports, and open in your web browser, or save as a PDF.
* Reformat Play Check reports, so that all pages are condensed on to a single page.
* Automatically create a folder named for a REQ number from a Canvas/Remedy ticket and zip up its contents into a password-protected ZIP archive.
* Interact with Reconciliation API (Check Rollback and Commit queues, check round status and unlock rounds).
* Lookup Game and Casino info from Quickfire/Casino portal.
* Interactive menu to access multiple functions.
* SA password function, to retrieve SA passwords for SQL servers.
* Server Descriptions function, to search and display Server Details page in your browser.
* All PowerShell cmdlets have comprehensive help included.
* All PowerShell cmdlets support the `-Verbose` parameter if you want detailed information about what the cmdlet is doing.
* Many more features planned!

## Compatibility
---
This module has been tested and developed with PowerShell 5.1 and PowerShell 7.
**The new Transaction Audit and API features in version 1.5 of this module require [PowerShell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3)**.
The module will still work with PowerShell 5.1 but the Transaction Audit and API features will not be available.

PowerShell 5.1 is currently the default version on Windows 10 and 11.

The module has not been tested with PowerShell Core on non-windows platforms. Support may be added in the future.


## Requirements
---
This module requires that [7-Zip](https://www.7-zip.org/download.html) is installed on your system. Microsoft Edge browser is also used to generate and convert PDF's.

The API features (e.g. transaction and financial audits) require [PowerShell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3)

The Excel features of this module require the third-party module [Import-Excel](https://www.powershellgallery.com/packages/ImportExcel/). The module will attempt to download and install it automatically if it isn't installed already.

The automatic update feature of this module requires [Git](https://git-scm.com/downloads) to be installed.

The module should generally load and run from any location, however some functions will not be accessible outside of the Derivco private network due to IP address restrictions. (.e.g. API features, Play checks, Game Statistics Reports).

The SQL functions in this module currently only work if run from a PowerShell session on Citrix. This is strictly for testing only and should not be used with live data.


# Getting Started
---

A detailed guide with screenshots is available at:
https://confluence.derivco.co.za/display/MIGS/Internal+-+Customer+Solutions+QFPowerShell+Module


## Quick Start
---
1. Download the files and folders from this repo, and save them somewhere easily accessible. **I very strongly recommend cloning this repository using Git instead of just downloading the files, as detailed in [Advanced Setup - auto update using git](#auto-update-using-git) below!**
2. Open a PowerShell window. (Not as administrator, normal user is fine)
3. Run the command: `Set-ExecutionPolicy Bypass -Scope CurrentUser` - this changes the security settings to allow you to run unsigned scripts. Otherwise, you will get an error trying to load the module.
4. Run the command: `import-module Quickfire.psd1` (You will need to include the full path to the downloaded file).
5. Run the command: `zz` and paste in a REQ number to get going!

If you copy a *Transaction_Audit.xlsx* file into the same folder as the *Quickfire.psd1* file is located, you will get the option to copy this file into the ticket folder *(Version 1.5 includes new Transaction Audit features that can generate this file automatically when using PowerShell 7).*

Ticket folders will be created under **Documents\Tickets** inside your OneDrive folder by default.
If you want to change the path where ticket folders and ZIP files are created, or where your *Transaction Audit.xlsx* is located, check out the [New-QFTicket Instructions](#new-qfticket)


## Advanced Setup
---

Using the Quick Start above means that you need to import the module every time you restart PowerShell. 
This section will describe how you can set up PowerShell to automatically load the module on startup.
In addition, you can optionally set up Git to pull from the main branch automatically, so this script will automatically update itself! Nifty!

### Auto update using Git
---
This step is optional, but it will ensure you can always get the latest version of the module from the Main branch.
1. First you need to [install Git](https://git-scm.com/downloads) on your PC if you don't have it already. The default installation options are fine.
2. Open a new PowerShell window, and change directory into a path where you want to save the module files. For example, you might want to just save into your user profile folder such as `c:\users\MyUsername`. It is recommended to store somewhere on your local PC rather than a network drive for performance reasons.
3. Ensure you are in the **Main** branch of the repo and click the **Clone** button in the upper right of the page.
4. Copy the HTTPS URI 
5. In your Powershell window, enter `git clone` followed by a space and paste the URI. Example: `git clone https://Derivco@dev.azure.com/Derivco/Software/_git/MIGS-IT-QFPowershell`
6. The module files will be downloaded into the *MIGS-IT-QFPowershell* folder.
7. In the future when you run `git pull` from inside this folder, it will automatically download any updated files from this repository. 

For example: 
``` 
PowerShell 7.3.4

PS C:\Users\MyUsername> cd .\MIGS-IT-QFPowershell\
PS C:\Users\MyUsername\MIGS-IT-QFPowershell> git pull
Already up to date.
```

As of version **1.3** an *Update-QFPowerShell* function has been added that will attempt to auto-update the module periodically, by running `git pull` automatically. So you don't need to worry about running `git pull` yourself.


### Automatically loading the module when PowerShell starts
---
You can add commands to your [PowerShell Profile file](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_profiles?view=powershell-5.1) so they run automatically when you open a PowerShell window. Note the folders are slightly different between PowerShell 5 (the default version included with Windows 10/11) and PowerShell 7 (which you would normally download and install manually.)

1. Create a **profile.ps1** file
- For your local user profile, open your **Documents** folder, and create a folder name **WindowsPowerShell** if it doesn't exist already. (For PowerShell 7, the folder is named **PowerShell** instead.)Then create a new text file named **profile.ps1** if it doesn't exist already.
NOTE: Sometimes OneDrive folder redirection messes with the local user profile, and PowerShell won't find the *profile.ps1* file. In this case try using the system profile instead.
- For the system profile (will affect all user accounts on your PC), open the folder **C:\Windows\System32\WindowsPowerShell\v1.0**. (For PowerShell 7, the folder is your PowerShell installation folder, **C:\Program Files\PowerShell\7** by default.) Then create a new text file named **profile.ps1** if it doesn't exist already. 
NOTE: Windows UAC might not let you create or edit this file. Instead, try saving it on your desktop and then drag and dropping the file into the above folder, then click Yes when asked if you want to make changes to this device.
2. Open the **profile.ps1** file in a text editor.
3. Add the following line to import the module:
```
Import-Module C:\Users\MyUsername\MIGS-IT-QFPowershell\Quickfire.psd1 
```
Substituting the path in the command example above, with the path of your MIGS-IT-QFPowershell folder.
Ensure you import the **PSD1** file rather than a **PSM1** file.

4. Save the file, open PowerShell and test! Try entering `gcm -module quickfire` and if the module loaded successfully, you should see several commands listed.


# Instructions
---
This section has brief instructions on how to use the functions in this module. 

A detailed guide with screenshots is available at:
https://confluence.derivco.co.za/display/MIGS/QF-Powershell+Guide 

Remember that each function has help built in. You can access it by using `help *function-name* -full`
For example: `help New-QFTicket -full`

Aliases are configured for most modules - you can simply enter the alias into PowerShell to run the function instead of the full function name. Some of the aliases also set specific parameters for the function.


## New-QFTicket
---
Alias: **zz** 

Alias: **za** (Disables the Out-GridView file selector, adds all files in the folder to the zip archive. Same as -NoGridView parameter)

Alias: **zx** (Does not display any menus to user; just zips up all files in the REQ ticket folder. Useful if you are using this cmdlet from an automated function as it doesn't require user input. Same as the -NoMenu parameter)

This cmdlet will ask the user to enter an REQ number corresponding to a support ticket from the Canvas/Remedy system.
It will then create a new folder named with the REQ number under a preconfigured path if the folder doesn't exist already.
The user will be given several options. The user can run the function to generate play checks and Game Statistics reports from this cmdlet.
If this function is used to save a playcheck to PDF, it then check the playcheck HTML data, to find the game name, MID and CID, and ask if you want to run a Game Statistics report for these games, and save that to PDF as well.

By default, tickets are created under your user profile, in the folder path "\OneDrive - Derivco (Pty) Limited\Documents\Tickets"
If you wish to change this, run the cmdlet with the argument **-TicketPath** specifying the full path where you would like the ticket folders to be created. It will then save this path in your user registry, and use the same folder path every time you run the command. For example: `zz -TicketPath 'C:\Users\MyUserName\Tickets\'`

If a *Transaction_Audit.xlsx* file exists in the same folder as this *Quickfire.psd1* file, (In the **MIGS-IT-QFPowershell** folder if you used Git), it can copy the spreadsheet into the REQ ticket folder and open it in Excel.
If the file already exists in the REQ ticket folder, it will ask if you want to overwrite it with a new blank copy or not.
If you have the Transaction Audit file saved somewhere else, run the cmdlet with the argument **-TransAuditFile** specifying the full path to the file. It will then save this path in your user registry, and use the same file every time you run the command. For example: `zz -TransAuditFile 'C:\Users\MyUserName\Documents\Transaction Audit.xlsx'`
If you are running **PowerShell 7**, version 1.5 of this module includes new API features that can pull the transaction audit data from Help Desk Express API, and automatically populate the Excel file. The TransAuditFile is not used when running audits in this manner, however if you select **(O) Open Audit File** and a transaction audit file doesn't already exist, the blank TransAuditFile will be copied into the current folder and opened.

Once all the required files are present in the folder it will compress them into a ZIP file, using the REQ number as a password, which can be attached to the ticket and sent to the customer.
The default ZIP file name is *GameData.zip*; this can be changed with the `-ZipFile` parameter. If you want to change the ZIP file name permanently, specify the `-ZipFile` parameter and the desired filename, and also specify `-ZipFileDefault` which will save this setting into your user registry. The new ZIP file name will be used every time this function is run in the future.

The new API Transaction/Financial Audit features available in version 1.5 of the module with **PowerShell 7** also support some additional parameters:

**-AuditStartDate** - The number of days ago for the starting date of Transaction and Financial Audits.
    The start date is generated by subtracting the specified number of days from todays date.
    By default this parameter is set to 14, so the start date for generating audits will be 14 days ago from today.
    e.g. if today's date is 2023-01-15 then the default start date for generating audits will be 2023-01-01.

**-AuditStartDateDefault** - The number of days ago for the starting date of Transaction and Financial Audits.
    As per **AuditStartDate** parameter, however this setting is saved in the user registry and remembered when this cmdlet is run again in the future.

**-AuditEndDate** - The number of days ago for the end date of Transaction and Financial Audits.
    The end date is generated by subtracting the specified number of days from todays date.
    By default this paramater is set to 0, so the end date for generating audits will be today.
    e.g. if today's date is 2023-01-15 and this parameter is set to 1, then the default end date for generating audits will be 2023-01-14.

**-AuditEndDateDefault** - The number of days ago for the end date of Transaction and Financial Audits.
    As per **AuditEndDate** parameter, however this setting is saved in the user registry and remembered when this cmdlet is run again in the future.

**-FinancialAuditDefault** - This parameter sets whether Financial Audits will be generated by default in the Transaction Audit menu.
    Financial Audits are enabled by default, until you set this parameter to False.
    This setting is saved in the user registry and remembered when this cmdlet is run again in the future.

**-SortOrder** -
    Sets the sorting order of Transaction and Financial audits.
    Valid settings for this parameter are "ASC" (ascending order) and "DESC" (descending order).
    Audits are sorted on the TransactionTime field.
    By default, audits are sorted in Descending order, so the newest transactions appear at the top of the report.
    This setting is saved in the user registry and remembered when this cmdlet is run again in the future.

**-TransAuditDefault** - This parameter sets whether Transaction Audits will be generated by default in the Transaction Audit menu.
    Transaction Audits are enabled by default, until you set this parameter to *False*.
    This setting is saved in the user registry and remembered when this cmdlet is run again in the future.

Several more parameters are available to customize the cmdlet's behaviour. See the full help page for further details.
Default parameters are saved in the user registry under **HKEY_CURRENT_USER\Software\QFPowershell**

### New-QFTicket - AutoMode
Starting from version 1.6 of the module, *New-QFTicket* supports a fully automated mode of operation.
Simply run ``New-QFTicket`` and specify the *REQNumber*, player *Login*, *CasinoID*, and optionally *TransactionIDs* parameters.
The cmdlet will generate a Transaction and Financial audit for the specified player automatically. 
If the *TransactionIDs* parameter is specified, this cmdlet will attempt to generate Play Checks for these game rounds and Game Statistics Reports for those games.
Multiple TransactionIDs can be specified, seperated by commas without spaces.
The cmdlet will then exit and output a pipeline object containing details of the player and ZIP file.

You must provide either a *UserID* or player *Login*, plus at least one *CasinoID*, or a *CasinoName*, or an *OperatorID* using the appropriate parameters.
If a player *Login* is provided without a *UserID*, this cmdlet will search the provided Casinos to locate the matching player and retrieve their *UserID*.

If multiple *CasinoID*'s are specified, this cmdlet will check each one for the matching Player.
If a *CasinoName* parameter is specified, this cmdlet will search for any Casino's that contain this string in their name, and search each one for the matching Player.
If an *OperatorID* parameter is specified, this cmdlet will retrieve a list of all *CasinoID*'s linked to the specified *OperatorID*, and search each one for the matching Player.

This cmdlet is configured with positional parameters, so you don't have to specify the parameter names, as long as you provide the *REQNumber*, *Login*, *CasinoID* and *TransactionIDs* parameters, in this order. 
e.g.
``zz REQ123456 GuyIncognito 12345 11,12,13``

Please refer to the cmdlet help text for further details.



## Export-QFExcel
---
This cmdlet exports data into an Excel spreadsheet.
It is tailored for working with QuickFire / Games Global transaction audits by default, but can be used for any Excel spreadsheet.

This cmdlet requires the Import-Excel third party module. It will attempt to download and install this module automatically.
The module can be downloaded manually via the command:
Install-Module -Scope CurrentUser ImportExcel
See the module website for details: https://www.powershellgallery.com/packages/ImportExcel/

By default, this cmdlet will copy a worksheet from a source file, into a new or existing Excel file.
The cmdlet will then populate the new worksheet with data provided in the $ExcelData parameter.
An *'*Audit.xlsx*'* file is included in the QFPowerShell repository under the *template* folder.
This Excel source file is configured with the standard Quickfire / Games Global transaction audit header.
If the source file cannot be found, a new empty Excel file will be created (or a blank worksheet in an existing Excel target file).

This cmdlet can also set date or number formatting on specified cell ranges, or change text and background fill colours.
A number of default parameter values are configured, which can be overwritten by specifying parameters for this function. See the full help
text for details.



## Format-QFPlayCheck 
---
Alias: **fpc**

*Since version 1.2, This is now performed automatically by Get-QFPlaycheck when saving as a PDF.*

Converts a Play Check with multiple pages into a single page.
Play Checks for some games are created with multiple pages, that can be selected from a control element at the bottom of the page.
This makes it cumbersome to save to PDF as you must select and save each individual page.
This function will modify a play check saved as MHTML to remove the page selector, expand any collapsed/minimised sections and display all content on one page.
It will also convert the file to PDF.

To use it:
-Generate the play check normally in your web browser
-Save the play check as 'Web Page - Single File (*.mhtml)'
-Run this function and specify the file name(s) - For example: `fpc *.mhtml` - will reformat and convert all MHTML files in the current directory



## Get-QFAAMSStatus
---
Alias: **aams**

Retrieves AAMS Participation status from the ADM Italy site www.adm.gov.it
You must specify an AAMS Participation code (A 16 digit code beginning with the letter N).

Details of the round returned include the Date, Bet Amount, Remote Session ID and Status.
The Status can be either 'Riscosso' (Round is completed and closed) or 'Registrato' (Round is open and needs to be processed).
NOTE - this function may require an Italy VPN to work, as the AAMS website appears to be blocked from IP addresses outside Italy.



## Get-QFAudit
---
Alias: **qfaudit**

Retrieves transaction and financial audit data from the Back Office Help Desk Express API.

You must provide an operator API Key, a player UserID and a CasinoID (aka ProductID/ServerID).
You must also provide a HostingSiteID, this can be retrieved via ```Invoke-QFPortalRequest``` with a *CasinoID* parameter.

By default this cmdlet will use the API endpoint api.valueactive.eu - this can be adjusted by changing the APIHost parameter.

API documentation is available at:
https://reviewdocs.gameassists.co.uk/internal/document/BackOffice/Help%20Desk%20Express%20API/1/Resources/FinancialAudits



## Get-QFBetSettingProc
---
Alias: **bsp**

Create bet setting procs for 1 to 1 and currency multiplier variations.

This cmdlet is used to create bet settings procs that CasinoPortal will not generate. Casino Portal and Sherlock do not contain the logic for, or allow settings to be changed for table games, and video bingo games.
This cmdlet does not contain the pre-existing logic that sherlock has for functional bet value options.
It will therefore not find the closest viable bet settings options to the users input. Instead this needs to be tested calculated before this cmdlet is used.
The cmdlet will take the players input OperatorID, MID-CID combos, Currencies, SettingsIDs and values and pre-fill all required SQL procs to be run on the DB.



## Get-QFETIProviderInfo
---
Alias: **qfeti**

Looks up support contact information for Quickfire ETI providers, from the file 'ETIProviders.csv'
This CSV file must be present in the same folder as this PowerShell Module file.

The information returned for each ETI provider includes a support email address, and a support portal URI and login credentials.
If the provider does not have any of this information available in the CSV file, an empty value will be returned.

You may specify either an ETI Provider Name, or their ID number.
The ETI Provider ID number can found in the Master Games List excel spreadsheet, downloaded from the Games Global Client Zone.


## Get-QFHelixAccessToken
---
Generates the authentication tokens required for the Helix ITSM API using hardcoded credentials.
By default, the service account 'RS-INT-MIGS-Automation-PowerShell' will be used, with a hardcoded password.
The token will be output to pipeline, and also set in the script-scoped object $script:ITSMtoken
This allows the Token object to persist after the function completes.
If $script:ITSMtoken is already set and the Token has not yet expired, a new Token will not be generated.
**This function is currently only to be used for testing purposes.**



## Get-QFHelixDefaultHeader
---
Generates a Helix Access Token and returns a hash table, which can be passed as a request header to Helix ITSM API.
This function is generally called internally from other functions before calling the API.
**This function is currently only to be used for testing purposes.**



## Get-QFHelixIncident
---
Retrieves an Incident and associated data from the Helix ITSM system.
The Incident Number of a valid must be provided as the IncidentNumber parameter.

The 'Detailed Description' field of the ticket will be parsed, and each field will be split into a hashtable as a key:value pair.
This hashtable will be included in the pipeline output as a member named 'DescriptionFields'.
**This function is currently only to be used for testing purposes.**



## Get-QFHelixIncidentWorkInfo
---
Retrieves all Work Info from the specified Incident on the Helix ITSM system.
The Incident Number of a valid must be provided as the IncidentNumber parameter.
This cmdlet will output all Work Info on the specified Incident as an array of PSCustomObjects.
**This function is currently only to be used for testing purposes.**



## Get-QFGameBlocking
---
Alias: **qfblock**

Requests game blocking information from gw2.mgsops.net for the specified Casino, Country and Game.

You may specify multiple CountryID, ModuleID and ClientID parameters.
Game Blocking will be checked for each combination of provided values.

A list of Countries is presented to the user via Out-GridView if the CountryId parameter is not specified.

When Game Blocking details are retrieved, a brief summary of each game's blocking status is displayed to the user.
The full details of game blocking are then output to pipeline.



## Get-QFGameStats
---
Alias: **gs**

Alias: **gsw** (Opens the Game Statistics Reports in the default web browser. Equivalent to 'Get-QFGameStats -OpenBrowser')

Alias: **gsp** (Use Edge's Print To PDF feature to save the Game Statistics Reports as a PDF in the current directory and opens the PDF file in the default PDF Viewer. Equivalent to 'Get-QFGameStats -SavePDF')

Alias: **gss** (Use Edge's Print To PDF feature to save the Game Statistics Reports as a PDF in the current directory, but doesn't open the PDF file. Equivalent to 'Get-QFGameStats -SavePDF -NoViewPDF')

This cmdlet is used to generate a list Game Statistics Reports links for the specified Player Login.
You may optionally specify a Game Name, ModuleID or ClientID to filter the list.
By default this cmdlet will simply output an object that contains a list of  Game Statistics Reports including the Game Name, ModuleID, ClientID and the URI to view the report.

Specifying the '-OpenBrowser' parameter will each Game Statistics Report in the default web browser.
Specifying the '-SavePDF' switch parameter will automatically save the generated Game Statistics Reports to a PDF file using Edge's 'Save as PDF' printer, then open in the default PDF Viewer program.
Specifying the '-NoViewPDF' switch parameter will not open the PDF file, and just silently save the generated PDF without further user interaction.



## Get-QFOktaToken
---
Alias: **okta**

This cmdlet requests an OKTA Bearer Token, for use with EpicAPI or the Casino Portal API.

This cmdlet requires the *Epic.WebApi.Client.DLL* file, which is included in the QFPowerShell repository under the 'lib' folder.
Please visit the [EpicAPI site](https://epicapi-v4.mgsops.net/) for further information.
This cmdlet also requires PowerShell Core or PowerShell 7 as this is a requirement for the DLL to function.

Bearer tokens generally have a lifetime of two hours. 
The default behaviour is to remember tokens as they are generated, and if this cmdlet is run again while the token is still valid, the existing token will be output to pipeline instead of requesting a new one.
The -Force parameter skips this check and will always generate a new token.

To use a token with Invoke-RestMethod or Invoke-WebRequest, pass the .Token member of the output object as an Authorization header.
For example, if you stored the output of this cmdlet into a $Token object, use the below value for the -Headers parameter:
```@{ Authorization = "Bearer " + $Token.Token }```



## Get-QFOperatorAPIKeys
---
Alias: **qfk**

Retrieves an operator API key from the Operator Security website: https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login

It will first look up the Operator Security Credentials for the specified *OperatorID* using the Casino Portal API.
If you specify a *CasinoID* parameter instead of an OperatorId, this cmdlet will first look up the OperatorID for that CasinoID using the Casino Portal API.



## Get-QFOperatorToken
---
Alias: **qfot**

Generates an operator API token using the provided API Key, via the Operator Security API.

By default this cmdlet will use the API endpoint *operatorsecurity.valueactive.eu* - this can be adjusted by changing the *APIHost* parameter.
The token will be output to pipeline along with the timestamp of issue (UTC) and expiry duration in seconds, plus the token expiry timestamp in local time.
The token value will be in the output member *'*AccessToken*

API documentation is available at the [Operator Security API Docs page](https://reviewdocs.gameassists.co.uk/internal/document/System/Operator%20Security%20API/1/Resources/OperatorTokens/3EFA1721EA)



## Get-QFPlaycheck
---
Alias: **pc**

Alias: **pcp** (Use Edge's Print To PDF feature to save the play check as a PDF in the current directory and opens the PDF file in the default PDF Viewer. Equivalent to 'Get-QFPlaycheck -SavePDF')

Alias: **pcs** (Use Edge's Print To PDF feature to save the play check as a PDF in the current directory, but doesn't open the PDF file. Equivalent to 'Get-QFPlaycheck -SavePDF -NoViewPDF')

This cmdlet is used to generate a Play Check for the specified Player Login, CasinoID and TransactionID.
You may optionally pass multiple TransactionIDs in a list and a Play Check will be generated for each one.
Your web browser will open to display each Play Check report. This cmdlet generates no other output.
You may optionally pass the -SavePDF switch parameter to automatically save the generated Play Check to a PDF file using Edge's 'Save as PDF' printer, then open in the default PDF Viewer program.
Specifying the '-NoViewPDF' switch parameter will not open the PDF file, and just silently save the generated PDF without further user interaction.



## Get-QFSQLServerSAPassword
---
Alias: **sa**

This cmdlet will retrieve the SA password for the specified SQL Server. 
You may pass a SQL Server host name using the -ServerName parameter; multiple server names can be provided either on the pipeline in an array, or by separating each server name with a comma.
If you do not specify a server name, you will be shown a GridView where you can select the required SQL servers from a list.
The GridView is populated by a CSV file in the 'src' folder. You can *Control-click* to select multiple servers.
You must also provide a Reason for requesting the SA password; if this function is invoked from inside New-QFTicket, the current REQ ticket number will be supplied automatically.


## Get-QFUser
---
Alias: **qfuser**

Checks that the specified Player Login exists on the specified CasinoID, and returns the matching UserID.
You must provide an operator API Key, a player Login and a CasinoID (aka ProductID/ServerID) using the appropriate parameters.
You must also provide a HostingSiteID, this can be retrieved via *Invoke-QFPortalRequest* with a CasinoID parameter.

You must provide a player Login matching the exact value from the Casino database, otherwise the player will not be found.
You may optionally provide the Casino Login Prefix (2 characters followed by an underscore) but this is not required.
Wildcards are not supported due to a limitation of the Account API.

By default this cmdlet will use the API endpoint api.valueactive.eu - this can be adjusted by changing the APIHost parameter.
API documentation is available at: https://reviewdocs.gameassists.co.uk/internal/document/Account/Account%20API/1/Resources/Accounts/01C3E50E29



## Invoke-QFAI
---
This cmdlet will take the provided support incident information in the Query parameter, and process it via OpenAI.
The AI should detect Win Verification and Game Play Analysis tickets, and provide a JSON format object containing parameters for New-QFTicket.
This is currently in development and further functions are planned to be added.

When the AutoMode parameter is set, the function call arguments generated by the AI will be passed to New-QFTicket automatically.
This will perform a transaction audit, playchecks and game statistics reports based on the parameters received from the AI.
By default, the output from AI is simply output to pipeline.
You must also provide a ticket REQ number via the REQNumber parameter.

The text to be processed should be provided in the Query parameter.
Since this text can span multiple lines, You may need to enclose this text in a 'here-string' and assign to an object to pass to this cmdlet.
e.g.:
```
$Query = @'
This is a here-string
it will escape all the newlines and symbols in this string
'@
```

If the AI decides the UserMessage is a Win Verification ticket, it will provide a JSON object containing parameters that can be passed directly to New-QFTicket.



## Invoke-QFAutoTicket
---
Automates processing of Game Play Analysis Incidents on the Helix ITSM system.
This includes identifying relevant Incidents for processing, retrieving Incident data, generating Play Checks and updating/closing the ticket.

This cmdlet searches for any Incidents that may be Game Play Analysis tickets, based on the default QueryString criteria in the cmdlet 'Search-QFHelixIncident'.
It then checks for a player Login, CasinoID, and TransactionID fields in matching Incidents.
If this information is found, this cmdlet will invoke New-QFTicket to generate transaction and financial audits for the specified Player. 
It will then generate Play Checks for the specified TransactionIDs, and Game Statistics Reports for any Games identified in the Play Checks.
**This function is currently only to be used for testing purposes.**


## Invoke-QFMenu
---
Alias: **qfm**

Presents an interactive menu for requesting information from Casino Portal or Reconciliation API.
For example, you can search for a Quickfire Casino via CasinoID or Casino Name, or list all casinos belonging to an OperatorID.
You can search for Quickfire Games by ModuleID/ClientID or Game Name.
You can also perform Reconciliation API functions such as checking for queued transactions and unlocking them.



## Invoke-QFPortalRequest
---
Alias: **qfp**

This cmdlet retrieves information from the Casino Portal API.

Documentation for the Casino Portal API is available at https://casinoportal.gameassists.co.uk/api/swagger/index.html
This cmdlet does not implement all functions of the API.
Data returned from the API will be output to pipeline. If no data is returned from the API, e.g. a non-existent CasinoID was specified, there will be no pipeline output.



## Invoke-QFReconAPIRequest
---
Alias: **qfr**

Invokes Reconciliation API functions such as managing Commit and Rollback queues for QuickFire operators.
An Operator API Token is required for authentication, and must be provided using the 'Token' parameter.

Documentation for the Reconciliation API is available at https://reviewdocs.gameassists.co.uk/internal/document/ExternalOperators/Reconciliation%20API/1
This cmdlet does not implement all functions of the API.



## New-QFHelixIncident
---
Creates a new Incident in the Helix ITSM system.
Currently, parameter values are hard coded.
**This function is currently only to be used for testing purposes.**


## New-QFHelixIncidentWorkInfo
---
Creates a new Work Info on the specified Incident on the Helix ITSM system.
A hash table of Field Names and corresponding Values must be provided, otherwise the Incident will not be updated.
The Work Info can bet set to to Public or Internal visibility using the "View Access" update field.
Note that this cmdlet cannot change the status of an incident, such as closing a request or setting the status reason to 'Client Action Required'.
**This function is currently only to be used for testing purposes.**



## Search-QFHelixIncident
---
Searches the Helix ITSM system for Incidents matching specified criteria and returns basic information for any matching Incidents.
The search criteria is specified using the 'QueryField' parameter. This must be a valid Helix ITSM query string.
If no QueryField parameter is set, a default value will be used.
Each Incident identified that matches the Query String will be output to pipeline, as an array containing a member object for each matching Incident.
**This function is currently only to be used for testing purposes.**



## Search-QFServerDetails
---
Alias: **sd**

This cmdlet is used to search the Server Details web site for the specified Server Name and opens the results in the default browser.
This function is useful if you need to request the SA password for a particular SQL server or look at other server details.
If you enter server name containing a backslash (\\) the cmdlet assumes you are looking for a SQL Instance name and will remove everything before the slash.
This command accepts multiple server names separated by commas, or simply run the command with no parameters and you will be prompted to enter multiple server names on separate lines. 
Press Enter on a blank line to begin the search.



## Search-QFUser
---
Alias: **qffind**

Searches multiple Casinos for a player with the specified Login.
If a player was found, return the matching UserID plus details of the Casino where the player was located.
You must specify a player Login and another parameter to search with - either an OperatorID, CasinoID or Casino Name.

You must provide a player Login matching the exact value from the Casino database, otherwise the player will not be found.
You may optionally provide the Casino Login Prefix (2 characters followed by an underscore) but this is not required.
Wildcards are not supported due to a limitation of the Account API.
This cmdlet can take a long time to complete if there are a large number of casinos to search through.

This cmdlet will automatically retrieve an Operator API Token for any OperatorID's found that match the specified search option.
An Operator API Key for 'All Products' must exist on the Operator Security site.
By default all UAT casinos will be excluded from the search. UAT search may be added in the future if such a requirement arises.



## Start-QFGame
---
Alias: **qflaunch**

Launches a Quickfire Game in the default web browser via CasinoID 21699 and OperatorId 47600 (Quickfire FakeAPI)
You must specify either MID and CID parameters, or a UGL Launch Code, for the desired Game.
A launch token is generated by making a request to gameshub.gameassists.co.uk - game launch will fail if this site is unreachable.

You may optionally specify a Currency or Language for the game using the associated parameters.
Otherwise the game will launch using Euros in English language.
Your account balance will automatically be set to 10,000.00 in 1:1 currencies (USD, GBP, EUR etc) or equivalent for other currencies.

You may also specify a ServerID parameter if you wish to launch a game using a ServerID/CasinoID that is supported by the Games Hub.
For example, the various Showcase casinos on each site. Refer to the Games Hub to identify which CasinoID's are supported.



## Update-QFHelixIncident
---
Updates the specified Incident on the Helix ITSM system.
This cmdlet can be used to change the status of an Incident, such as closing a request, or setting the Status Reason to 'Customer Action Required'.
A hash table of Incident Field Names and corresponding Values must be provided, otherwise the Incident will not be updated.
**This function is currently only to be used for testing purposes.**



## Update-QFPowershell
---
Automatically updates this module, by performing a 'git pull' from the Azure Devops repository. By default, this cmdlet willcheck for updates every week.
This cmdlet is called automatically by **New-QFTicket**.

This cmdlet requires Git to be installed on your machine, and also requires the QFPowershell module files to be cloned from the Azure Devops repository using the 'git clone' command. Refer to [Advanced Setup - auto update using git](#auto-update-using-git) for instructions.

Specifying the '-DisableUpdateCheck' parameter will disable the update checks permanently. 
Specifying the '-UpdateInterval' parameter allows you to set the duration in days between update checks. By default this is set to 7, so will only check for updates each week.
Specifying the '-UpdateNow' parameter forces an update check immediately, regardless of the Update Interval, and even if updates have previously been disabled.




# Known Issues and Troubleshooting
---

* **Get-QFGameStats doesn't work!**

The Games Monitoring site has recently implemented OKTA authentication. This module currently does not work with this authentication method. 
We are currently working to implement this into module, or find an alternative solution. 

Unfortunately, until this is completed,  the Get-QFGameStats function will not work. 

The Game Statistics reports will also not be generated by *New-QFTicket* in AutoMode until this is resolved. The Get-QFGameStats function can still be run manually or from the *New-QFTicket* menu, but it will fail with an error.


* **I keep receiving this error: Get-QFUser: Account API request failed - please check that you have a valid Operator Token and that it has permission for the specified CasinoId.**

This seems to be an issue with API Keys not correctly mapped between OperatorID's and CasinoID's. 
This module attempts to use API keys for 'All Products' for the given OperatorID rather than for specific CasinoID's wherever possible. However, in some cases 'All Products' API keys don't seem to actually work for all CasinoID's linked to a given Operator.

You may need to reach out to MIGS IT - CORE to assist with rectifying this issue. It could simply be a case of a new CasinoID still in the process of getting set up for this Operator, in which case you can just ignore these errors.



* **I keep receiving registry errors -e.g. "Failed to save LastUpdate date into registry"**

The module saves certain settings and info into your user registry hive, under the key **HKEY_CURRENT_USER\Software\QFPowershell**
If there is a permissions issue, or something is blocking writing to registry (e.g. antivirus software on your PC) these registry writes can fail.

You may be able to work around this by manually creating this key:
-Run **regedit.exe**
-Expand **HKEY_CURRENT_USER** 
-Right click on **Software** and select **New Key**
-Enter the key name **QFPowershell**
-Try running the cmdlet again and see if this resolves the issue.



* **Playcheck Game MIDs aren't correct** 

When saving a Play Check to PDF, the script will look inside the HTML code for the game name, MID and CID. However, this is not always correct. Often the MID will be for the 'standard' variant of the game (with the highest RTP%) even though the player was actually playing a lower RTP% variant. This has been observed with Thunderstruck Stormchaser, and Jane Blonde Max Volume, and will possibly occour with other games. Players who are playing a V92 or V94 variant (e.g. Thunderstruck Stormchaser V92 MID = 10992) will get picked up as a MID of the standard variant (MID = 10956) by this script.

This does not appear to be a bug in this PowerShell script; the Play Check HTML data retrieved from the server contains the incorrect game MID. I'm unsure why this is happening, and why it doesn't happen on all games. It just seems to be an idiosyncrasy of the play check system. This is only really an issue when you want to generate Game Statistics Reports for affected games. You should specify the game name in this case, rather than the game MID. If the player only has game history in a lower RTP% game variant, specifying the standard version game MID for Game Statistics Reports will not provide any results.


* **Can't generate Game Statistics Reports for any players**

The Game Statistics report requires the player's full LoginName, including our prefix. The prefix consists of two characters followed by an underscore.

That is, use the LoginName provided by the Casino database, and not the one provided by Vanguard State database, as the latter does not include this prefix. PlayChecks seems not to care about this prefix (except for SSO accounts like PokerStars) but I'd suggest including the prefix where possible either way.



* **Can't generate Playchecks or Game Statistics Reports for PokerStars**

As described above, the Game Statistics Reports won't work with the SSO login for PokerStars players that you get from the Vanguard State database. e.g. `2-12345678-GameNameMGS` - you must provide the LoginName from the Casino database including the prefix. e.g. `YA_2-12345678-GameNameMGS`

If you use the *Find Player Account - SSO* SQL Query in the Casino database, you can also use the **Mapped Account Name in tb_UserAccount**, for example, `00087654321` 
This can sometimes be inconsistent, for reasons unknown. If the LoginName doesn't work, try the Mapped Account Name, or vice versa.



* **I get an error when I try to run a Transaction or Financial Audit**

The audits make use of the [Help Desk Express API](https://reviewdocs.gameassists.co.uk/internal/document/BackOffice/Help%20Desk%20Express%20API/1/Resources/FinancialAudits) which requires an Operator API Key. The API keys are retrieved from the [Operator Security site](https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login) which requires credentials that are stored in [Casino Portal](https://casinoportal.gameassists.co.uk/Information) under the Information tab. If these credentials are incorrect or missing from *Casino Portal* for an operator, the script cannot retrieve an API key and the transaction/financial audits will fail. 

Another issue is that the credentials are correct but there are no API keys loaded for this operator, in this case you can simply login to the Operator Security Site using the operator's credentials, and generate a new API Key for *all products*.



* **I don't know where this is creating my ticket folders, and it doesn't give me a Transaction Audit option**
[Read The Fine Manual :) New-QFTicket Instructions](#new-qfticket)



* **How do I disable the auto update feature? I don't use git and its just giving me errors.**

You *really should* use git, it is a very useful tool, and it's part of the Dreadnought training program. But if you really don't want to, just run **Update-QFPowerShell -DisableUpdateCheck** and it will stop bugging you.



# Changelog
---
## Version: 1.6.4
* Get-QFPlayCheck - bug fix for broken stylesheets when saving as PDF
* Get-QFPlayCheck - added support for Jelly ETI provider, will embed link to the provider's video playcheck in the PDF
* Updated ETIProviders.csv


## Version: 1.6.3
* Get-QFPlayCheck - bug fix for broken images, added user-agent parameter to Edge call when saving as PDF


## Version: 1.6.2
* New-QFTicket - disabled Game Statistics reports in AutoMode. These functions are currently not working due to new OKTA authentication.
* Get-QFSQLServerSAPassword - new Username parameter, also added logic to check for Username column in the SQLServers.CSV file, to specify the SQL account username to retrieve passwords for.
* Updated SQLServers.CSV with Traditional Casino DB's, provided by Alex Bowker - default Username for these casinos is DBReadOnly_PI
* New Remedy integration functions, provided by Bernhard Heije
* Updates to ETIProviders.CSV - Added Jelly Entertainment and changed OnAir to Proxy Live
* Get-QFPlayCheck - bug fix for card images in 'Aces and Eights' game
* README.MD updates - removed warnings about PowerShell 7.4, added note about Get-QFGameStats, added links to Confluence pages for Getting Started and Instructions


## Version: 1.6.1
* New modules - Quickfire-AI.psm1 - functions for integration with OpenAI ChatGPT
* Get-QFGameStats and Get-QFPlaycheck - bug fixes, replaced deprecated MS Edge command line argument
* Updated SQLServers.CSV for Get-QFSQLServerSAPassword function
* Updates for Helix integration functions, provided by Bernhard Heije
* Start-QFGame updated for new Games Hub; added Showcase, QFGames, Language, ServerID and Balance parameters
* Get-QFAudit now gives detailed error messages from HDE API
* New-QFTicket - bug fix for Transaction menu option (F) 'Open Audit File'
* Bundled **Epic.WebApi.Client.dll** file updated to 4.0.0.756
* Get-QFPlayCheck - New parameters 'OpenExplorer' and 'FileName' - submitted by Alex Bowker

## Version: 1.6
* New modules - Quickfire-Helix-Base.psm1 and Quickfire-Helix-GPA.psm1- Helix integration functions - contributed by Bernard Heije. Currently for testing purposes only!
* New modules - Quickfire-BetSettingsProc.psm1 - Bet settings proc builder for table games and video bingo - contributed by Harley Osgood.
* New function - Get-QFUser - Checks that the specified Player Login exists on the specified CasinoID, and returns the matching UserID.
* New function - Search-QFUser - Searches multiple Casinos for a player with the specified Login.
* New function - Get-QFETIProviderInfo - lists ETI provider contact information.
* New function - Invoke-QFReconAPIRequest - performs actions against Reconciliation API (List contents of Vanguard Commit and Rollback queues, unlock transactions etc)
* New function - Invoke-QFMenu - interactive menu for performing various API functions
* New function - Start-QFGame - launches a Quickfire game via qfgames.gameassists.co.uk (CasinoID 18226)
* New function - Get-QFGameBlocking - retrieves game blocking details from gw2.mgsops.net for the specified Casino, Game and Country.
* New function - Get-QFAAMSStatus - retrieves AAMS game session details for specified Participation Code, for Italian casinos.
* Removed Get-QFPlayer and Get-QFSQLServers - only worked on citrix, no longer used
* New-QFTicket: New AutoMode feature - specify player Login, CasinoID and optionally TransactionID parameters, to automatically create audits and playchecks. Also checks Reconciliation API round status, and queued transactions in Vanguard Commit/Rollback queues for specified player. Results will be output to pipeline.
* New-QFTicket: AuditStartDate and AuditEndDate parameters to override default settings
* New-QFTicket: Entering **O** or **C** for REQ number automatically opens folder, or copies folder path, for previous REQ number
* New-QFTicket: New 'API Functions' menu option
* Invoke-QFPortalRequest: Now uses Casino Portal instead of Quickfire Portal
* Get-QFOktaToken: Remembers tokens generated in current session and will output existing token if not yet expired, instead of always requesting a new token. Specify -Force parameter to override and always request a new token.
* Get-QFOperatorToken: Remembers tokens generated in current session and will output existing token for specified Operator if not yet expired, instead of always requesting a new token. Specify -Force parameter to override and always request a new token.
* Get-QFPlaycheck: Now identifies ETI games and displays contact info, some error handling for specific ETI providers; outputs game details to pipeline.
* Get-QFGameStats: Outputs report results to pipeline.
* Get-QFSQLServerSAPassword: Skip SSL certificate verification for Powershell 6 or newer. Corrected some SQL server hostnames e.g. Goldfishka
* Bundled **Epic.WebApi.Client.dll** file updated to 4.0.0.476

## Version: 1.5
* New function - Get-QFOktaToken - generate an OKTA Bearer token using bundled Epic API DLL (requires PowerShell 7)
* New functions - Get-QFAudit, Get-QFOperatorAPIKeys, Get-QFOperatorToken, Invoke-QFPortalRequest - interacting with Help Desk Express, Operator Security and QuickFire Portal APIs
* New function - Export-QFExcel - creates Excel spreadsheet from template, for Transaction and Financial audits
* New-QFTicket: Refactored to remove some script scoped variables, moved Playcheck and Gamestats menu items into local functions
* New-QFTicket: New Transaction Audit menu for automated Transaction/Financial audits via new API functions. Only available with PowerShell 7. Old menu retained for previous versions.
* New-QFTicket: now looks up Login Prefix from QuickFire Portal after running Transaction Audit, and offers this value as default when running Game Stats or Playcheck
* New-QFTicket: Added logic to retry Get-QFGameStats with userID, casinoID and GamingSystemID if Login doesn't work
* Get-QFGameStats: added parameters to support UserID plus CasinoID and GamingSystemId, as alternative to specfying a player Login
* Get-QFPlayCheck: Sort-Unique transaction numbers (prevents running duplicate playchecks), temp files now uniquely named, misc bug fixes
* Get-QFSQLServersSAPassword: Added GIC2 and Betway Africa/Goldfishka servers to SQLServers.CSV file

## Version: 1.4
* CrushFTP functions removed. No longer working since OTP email codes are enforced.
* Split Quickfire.psm1 into multiple module files to make the functions easier to read and maintain. Quickfire.psd1 defitinion file updated to automatically load each module.
* Changed some variables from script scoped to global scope in Get-QFPlaycheck, Get-QFGameStats, and New-QFTicket to resolve issues caused by splitting these functions into their own module files.

## Version: 1.3
* New Update-QFPowerShell function, auto updates the module from the Azure Devops repository every week by default.
* New feature in New-QFTicket - will now remember the last used REQNumber and if you run it again without specifying an REQNumber parameter, it will offer this as a default option. 
* Bug fix - Get-QFPlaycheck: Fix for 'Session Expired' error, changed Edge parameters and added a check/retry loop
* Bug fix - Get-QFPlaycheck and Get-QFGameStats: URL Encode player login names in the URI, fixes issues with login names containing symbols.
* Bug fix - New-QFTicket: don't show the Grid View file selector if only one file is present in the current folder.

## Version: 1.2
* New Get-QFGameStats function to generate Game Monitor/Game Statistics reports for player's game history; and open in web browser or save as PDF. Can filter by game name, MID and/or CID.
* New Get-QFSQLServersSAPassword function to retrieve SQL server SA passwords.
* Get-QFPlaycheck: Now Format-QFPlaycheck is integrated into this function when saving as a PDF. This makes Format-QFPlaycheck mostly redundant but it is still available for use with MHTML files.
* Get-QFPlaycheck: Play Checks are no longer hard-coded to open in Chrome and will now open in your default web browser.
* Get-QFPlaycheck: Added -Hostname parameter in case the address changes in the future
* Format-QFPlaycheck: now removes the + Expand icon for extra rounds (e.g. Reel Positions 2 in Thunderstruck Stormchaser) - they were already expanded automatically so this image is redundant
* New-QFTicket: Added menu item for Get-QFGameStats and Get-QFSQLServersSAPassword
* New-QFTicket: When saving a playcheck to PDF, will look in the HTML data to find game name, MID and CID and ask user if they want to generate a Game Stats report for these games
* New-QFTicket: Will save Login and CasinoID into a file in the current folder, and offer them as default options when running a play check or game stats report again
* New-QFTicket: Will rename existing zip output file instead of just adding to the existing 
* New-QFTicket: Default 7zip compression level set to 7 (Maximum)
* New-QFTicket: Now copies the ZIP file path to the clipboard upon completion, ready to attach to a ticket. On a Work Info, click *Attachments > Choose Files*, then paste the path into the *File Name* box and click **Open**. The ZIP file will be attached to the work info.
* Readme updates: added table of contents; clarified use of Format-QFPlaycheck, and setup paths in New-QFTicket
* Now uses Microsoft Edge to generate and convert PDF's instead of Chrome; seems to fix some formatting issues in playchecks, and is installed by default with Windows, unlike Chrome. Web sites will open in whatever browser you have set to default on your system.
* Changes to the repo configuration (.engineeringinsights, .editorconfig, .gitattributes files) to improve the maturity level of the repo.
* PowerShell 7 compatibility fixes
* Formatting fixes, extra comments

## Version: 1.0.1
* Initial commit to Azure Devops repository
* Added New-QFTicket function and aliases zz (select files using Out-GridView) and za (Zip All files without Out-GridView)
* Added new Search-QFServerDetails function and alias sd to search the Server Details page and open the results in the default web browser.
* Bug fixes for Format-QFPlaycheck function and added alias fpc
* Added aliases pc, pcp and pcs for Get-QFPlaycheck function and new -SavePDF and -NoViewPDF parameters. Also added check/retry if PDF file is below a certain size (likely blank)
* Updates to function definitions and help text
* Formatting fixes, made code easier to read and more consistent styling


# Future Goals
---
* ~~Automate Vanguard State and Financial audits into Excel files~~ DONE!
* Look up Commit/Rollback queues, generate reports and clear transactions.
* Graphical User Interface for enhanced usability
* Link into Canvas/Remedy and automatically search for player UserIDs, and update Team Notes on the ticket with info
* Look up stuck rounds or VGS queues and update Team Notes on ticket with info
* Automatically upload zip files into Canvas/Remedy and attach to tickets
* ~~Automatically save Play Checks as MHTML, run Format-Playcheck, and convert to PDF in one go (rather than the manual process now, as Chrome cannot automatically save as MHTML)~~ DONE!
* ~~Automate Game Monitor reports and save as PDF~~ DONE!
* Move into XSOAR to run automatically (or run as a scheduled task on a trusted Windows server)


# Contribute
---
Contributions are welcome! Please refer to the **CONTRIBUTING.md** file inside this repository.


# Contact and Help
---
Any questions, comments, bug reports, ideas or suggestions for improvement are welcome!
Contact Christopher Byrne on Teams or via email: christopher.byrne@derivco.com.au
Please also reach out if you require any assistance with this module.
Or visit the [MIGS-IT-CustomerSolutions General channel on Teams](https://teams.microsoft.com/l/channel/19%3a88e8e42686ec4cc0ab3aae2add1834fd%40thread.skype/General?groupId=9c9bac3c-b436-4d22-a819-94255dcea0b6&tenantId=72aa0d83-624a-4ebf-a683-1b9b45548610) for further assistance or general queries.