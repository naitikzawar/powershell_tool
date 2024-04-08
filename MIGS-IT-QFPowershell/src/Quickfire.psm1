###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                                Main Module                                  #
#                                   v1.6.3                                    #
#                                                                             #
###############################################################################

#Author: Chris Byrne - christopher.byrne@derivco.com.au


function New-QFTicket {
    <#
    .SYNOPSIS
        Automates process of generating player audits and playchecks, then combining them into a password protected ZIP file for Canvas/Remedy support tickets.

    .DESCRIPTION
        This cmdlet will ask the user to enter an REQ number corresponding to a support ticket from the Canvas/Remedy system.
        It will then create a new folder named with the REQ number under a preconfigured path if the folder doesn't exist already.
        The user will be given several options. if a Transaction Audit file exists in the same location as the module file, it can copy this into the REQ ticket folder and open it in Excel.
        The user can also select a menu item to generate play checks and Game Statistics reports, and various other functions, from this cmdlet.
        Once all the required files are present in the folder it will compress them into a ZIP file, using the REQ number as a password, which can be attached to the ticket and sent to the customer.
        
        Several parameters are available to customize the cmdlet's behaviour. See the full help page for further details.
        Default settings for this cmdlet are stored in the user Registry, under the key 'HKEY_CURRENT_USER\Software\QFPowershell'
        Deleting this key will reset all configured settings.

        AUTOMODE
        ========
        This cmdlet also supports a fully automated mode without any user interaction, when 'AutoMode' parameters are set.
        e.g. running the cmdlet with a UserID and CasinoID, or a LoginID and a CasinoName parameter, will enable the AutoMode features.
        AutoMode features require PowerShell Core or PowerShell 7.

        This mode is intended to be used by other PowerShell scripts to automate transaction audits, playchecks, and creation of the ZIP file.
        Menus will not be displayed.
        The cmdlet will generate a Transaction and Financial audit for the specified player automatically. 
        If any TransactionID's are specified, this cmdlet will attempt to generate Play Checks for these game rounds and Game Statistics Reports for those games.
        Vanguard Commit and Rollback queues will also be checked for any queued transactions.
        The cmdlet will then exit.
        If the -PipelineOutput parameter was specified, the cmdlet will output a pipeline object containing details of the player and ZIP file, plus results of the game statistics and Vanguard queues checks - see OUTPUT help section for more information.

        You must provide either a UserID or player Login, plus at least one CasinoID, or a CasinoName, or an OperatorID using the appropriate parameters.
        If a player Login is provided without a UserID, this cmdlet will search the provided Casinos to locate the matching player and retrieve their UserID.

        If multiple CasinoID's are specified, this cmdlet will check each one for the matching Player.
        If a CasinoName parameter is specified, this cmdlet will search for any Casino's that contain this string in their name, and search each one for the matching Player.
        If an OperatorID parameter is specified, this cmdlet will retrieve a list of all CasinoID's linked to the specified OperatorID, and search each one for the matching Player.

        Any Default settings that are saved in the user registry will also apply in this mode. (e.g. ZipFileDefault, FinancialAuditDefault, AuditStartDateDefault etc.)

    .EXAMPLE
        New-QFTicket REQ123456
            Creates a folder named REQ123456 under the TicketPath folder and prompts the user for further action.

    .EXAMPLE    
        New-QFTicket -REQNumber REQ123456 -Login ABCXYZ -CasinoID 54321 -TransactionIDs 123 -PipelineOutput
            Enables AutoMode for transaction audits and play checks without any user interaction.

            Creates a folder named REQ123456 under the TicketPath folder.
            Then looks up the Login on the specified CasinoID and retrieves the matching UserID.
            Then will retrieve a Transaction and Financial audit for this player and save into an Excel audit file.
            Then will generate a Play Check for TransactionID 123, and then generate a Game Statistics Report for any games located in the Play Check data.
            Finally, will output a pipeline object containing details of the player and ZIP file if the process was successful.
            If the player cannot be located, will fail with an error and provide no pipeline output.

    .EXAMPLE    
        New-QFTicket -REQNumber REQ123456 -Login ABCXYZ -CasinoID 54321,65432,76543 -TransactionIDs 123,124,125 -PipelineOutput
            Enables AutoMode for transaction audits and play checks without any user interaction.
            
            Creates a folder named REQ123456 under the TicketPath folder.
            Then searches each of the specified CasinoID's for a player with a matching Login, and retrieves the matching UserID.
            Then will retrieve a Transaction and Financial audit for this player and save into an Excel audit file.
            Then will generate Play Checks for TransactionIDs 123, 124 and 125, and then generate a Game Statistics Report for any games located in the Play Check data.
            Finally, will output a pipeline object containing details of the player and ZIP file if the process was successful.
            If the player cannot be located, or if multiple matching players are located, the cmdlet will fail with an error and provide no pipeline output.

    .EXAMPLE    
        New-QFTicket -REQNumber REQ123456 -UserID 987654321 -CasinoID 54321 -TransactionIDs 123 -PipelineOutput
            Enables AutoMode for transaction audits and play checks without any user interaction.

            Creates a folder named REQ123456 under the TicketPath folder.
            Then will retrieve a Transaction and Financial audit for this player and save into an Excel audit file.
            Then will generate a Play Check for TransactionID 123, and then generate a Game Statistics Report for any games located in the Play Check data.
            Finally, will output a pipeline object containing details of the player and ZIP file if the process was successful.
            If the player cannot be located, this cmdlet will fail with an error and provide no pipeline output.

    .EXAMPLE    
        New-QFTicket -REQNumber REQ123456 -Login ABCXYZ -CasinoName Betting -TransactionIDs 123 -PipelineOutput
            Enables AutoMode for transaction audits and play checks without any user interaction.

            Creates a folder named REQ123456 under the TicketPath folder.
            Then searches for any casinos with "Betting" in their name, for a player with a matching Login, and retrieves the matching UserID.
            Then will retrieve a Transaction and Financial audit for this player and save into an Excel audit file.
            Then will generate a Play Check for TransactionID 123, and then generate a Game Statistics Report for any games located in the Play Check data.
            Finally, will output a pipeline object containing details of the player and ZIP file if the process was successful.
            If the player cannot be located, or if multiple matching players are located, the cmdlet will fail with an error and provide no pipeline output.

    .EXAMPLE    
        New-QFTicket -REQNumber REQ123456 -Login ABCXYZ -CasinoName Betting -TransactionIDs 123 -AuditStartDate 2023-01-01 -AuditEndDate "2023-01-11 23:00" -PipelineOutput
            Enables AutoMode for transaction audits and play checks without any user interaction.

            Creates a folder named REQ123456 under the TicketPath folder.
            Then searches for any casinos with "Betting" in their name, for a player with a matching Login, and retrieves the matching UserID.
            Then will retrieve a Transaction and Financial audit for this player and save into an Excel audit file.
            These Audits will retrieve transactions dated between 12:00 AM Jan 1st 2023, and 11PM Jan 11th 2023.
            Then will generate a Play Check for TransactionID 123, and then generate a Game Statistics Report for any games located in the Play Check data.
            Finally, will output a pipeline object containing details of the player and ZIP file if the process was successful.
            If the player cannot be located, or if multiple matching players are located, the cmdlet will fail with an error and provide no pipeline output.

    .EXAMPLE
        New-QFTicket REQ123456 -ZipFile TicketInfo
            Creates a folder named REQ123456 under the TicketPath folder and prompts the user for further action.
            Once the appropriate menu item is selected, a password protected Zip file will be created with the name 'TicketInfo.zip'

    .EXAMPLE
        New-QFTicket REQ123456 -ZipFile TicketInfo -ZipFileDefault
            Creates a folder named REQ123456 under the TicketPath folder and prompts the user for further action.
            Once the appropriate menu item is selected, a password protected Zip file will be created with the name 'TicketInfo.zip'
            'TicketInfo.zip' will be saved in the user registry as the default Zip file name when this command is run in the future.

    .EXAMPLE
        New-QFTicket -TransAuditFile "C:\Users\Myname\Documents\Transaction Audit.xlsx"
            Specifies the full path and filename of the Transacton Audit file and saves this setting in the user registry.

    .EXAMPLE
        New-QFTicket -TicketPath "C:\Users\Myname\Tickets\"
            Specifies the full path where you wish to create folders for each REQ ticket and saves this setting in the user registry.

    .EXAMPLE
        New-QFTicket REQ123456 -NoMenu
            Creates a folder named REQ123456 under the TicketPath folder if it doesn't exist then exits.
            if the folder does exist it will simply add all the files in the folder into a compressed ZIP archive using the specified REQNumber as a password.
            Will not prompt user for any further action.
            Using the alias 'zx' will automatically set this parameter for you.

    .EXAMPLE
        New-QFTicket -FinancialAuditDefault $false
            Disables generating Financial Audits by default.
            This setting is saved in the user registry, and will be remembered when this cmdlet is run again in the future.
            Set the FinancialAuditDefault parameter to $true to re-enable Financial Audits.
            You can also switch Financial and Transaction audits on/off inside the Transaction Audit menu, but the setting is not saved.

            You can toggle Transaction Audits in a similar manner with the TransAuditDefault parameter.
            The default setting is to generate both Transaction and Financial audits.

    .EXAMPLE
        New-QFTicket -SortOrder ASC
            Sets Transaction and Financial audits to sort in Ascending order on the TransactionTime field.
            (i.e. the oldest transactions are at the top of the report.)

            By default, the audits are sorted in Descending order on the TransactionTime field.
            (i.e. the newest transactions are at the top of the report.)
            To set back to Descending order, set the SortOrder parameter to "DESC"

            This setting is saved in the user registry, and will be remembered when this cmdlet is run again in the future.
            You can also toggle the sort order inside the Transaction Audit menu, but the setting is not saved.

    .EXAMPLE
        New-QFTicket -AuditStartDateDefault 7
            Sets Transaction and Financial audits with a starting date of 7 days ago from today.
            e.g. If today's date is 2023-01-08 then the default start date for generating audits will be 2023-01-01.

            The default setting is 14 days ago from today.
            This setting is saved in the user registry, and will be remembered when this cmdlet is run again in the future.
            You can configure the audit End dates in a similar manner with the AuditEndDateDefault parameter. (By default, the end date is set to the current date.)
            
    .PARAMETER AuditStartDate
        The date and time of the earliest transactions to retrieve when generating Transaction and Financial audits.
        Transactions older than the specified EndDate will be excluded from the audit.
        All dates and times must be in UTC time zone.

        You must provide a value DateTime value.
        e.g. 
        "2023-01-10"
        or
        "2023-02-28 11:00"
        If you don't specify a time, this will default to 12:00:00 AM on the specified date.

        By default, the AuditStartDateDefault value will be used, if configured.
        Please refer to the help for that Parameter for further information.
        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER AuditStartDateDefault
        The number of days ago for the starting date of Transaction and Financial Audits.
        The start date is generated by subtracting the specified number of days from todays date.
        By default this parameter is set to 14, so the start date for generating audits will be 14 days ago from today.
        e.g. if today's date is 2023-01-15 then the default start date for generating audits will be 2023-01-01.

        This setting is saved in the user registry and remembered when this cmdlet is run again in the future.
        If you  wish to adjust the date range for an audit without changing the default setting, 
        you can use the AuditStartDate parameter, or change the date from the Transaction Audit Menu.
        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER AuditEndDate
        The date and time of the most recent transactions to retrieve when generating Transaction and Financial audits.
        Transactions newer than the specified EndDate will be excluded from the audit.
        All dates and times must be in UTC time zone.

        You must provide a value DateTime value.
        e.g. 
        "2023-01-10"
        or
        "2023-02-28 11:00"
        If you don't specify a time, this will default to 12:00:00 AM on the specified date.

        By default, the AuditEndDateDefault value will be used, if configured.
        Please refer to the help for that Parameter for further information.
        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER AuditEndDateDefault
        The number of days ago for the end date of Transaction and Financial Audits.
        The end date is generated by subtracting the specified number of days from todays date.
        By default this parameter is set to 0, so the end date for generating audits will be today.
        e.g. if today's date is 2023-01-15 and this parameter is set to 1, then the default end date for generating audits will be 2023-01-14.

        If you  wish to adjust the date range for an audit without changing the default setting,
        you can use the AuditEndDate parameter, or change the date from the Transaction Audit Menu.
        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER AutoGameStats
        When running a Play Check, this cmdlet can identify the games in the play check data.
        The default behaviour, when this parameter is not set, is to ask the user if they want to generate the Game Statistics Reports for these games.
        If this parameter is set, this cmdlet will automatically generate Game Statistics Reports for any identified games without prompting the user.

        This parameter has no effect in AutoMode. Game Statistics reports will automatically be generated after a Play Check in AutoMode.

    .PARAMETER CasinoID
        Specify the CasinoID that the player belongs to for AutoMode.
        You must specify this parameter in conjunction with a UserID or Login parameter.

        You may specify multiple CasinoID's seperated by commas, if you also specified a player Login.
        This cmdlet will check for the specified player Login on each CasinoID.
        If a UserId parameter is specified instead of a Login, you must specify only a single CasinoID - there is no API to check for a UserID on a particular CasinoID.

        The cmdlet will fail with an error, and provide no pipeline output, if the player Login cannot be located on the specified CasinoIDs.

    .PARAMETER CasinoName
        Specify the name of the Casino that the player belongs to for AutoMode.
        You must specify this parameter in conjunction with a Login parameter.

        This cmdlet will search for a list of Casino's that match the provided CasinoName, via Casino Portal API.
        Each of these Casinos will be checked for a player that matches the provided Login.

        Any Casinos that include the provided CasinoName parameter anywhere in their Name will be included in the search.
        e.g. Specifying -CasinoName "Bort" will match any of these Casino Names:
        "Bort's Casino"
        "No Borts Allowed"
        "My Casino Is Also Named Bort"

        Wildcards are not permitted.

        The cmdlet will fail with an error, and provide no pipeline output, if the player cannot be located on any of the matching Casinos.

    .PARAMETER FinancialAuditDefault
        This parameter sets whether Financial Audits will be generated by default in the Transaction Audit menu.
        Financial Audits are enabled by default until you set this parameter to False.
        This setting is saved in the user registry and remembered when this cmdlet is run again in the future.

        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER IDDQD
        Degreelessness mode on

    .PARAMETER Login
        Specifies the player Login for Transaction and Financial audits in AutoMode.
        You must also specify a CasinoID, CasinoName or OperatorID parameter.
        You may optionally specify a UserID parameter, but this is not required.
        
        You may optionally include the login prefix from the Casino DB but this is not required. 
        This cmdlet will look up the login prefix from the Casino Portal API and add it to the provided Login value automatically, if it is not included already.

        The cmdlet will fail with an error, and provide no pipeline output, if the player Login cannot be located on the specified Casinos.

    .PARAMETER NoAutoAudit
        In AutoMode, will not retrieve Transaction or Financial audits from Help Desk Express API for the specified player.

        If you wish to set this parameter by default every time the cmdlet is run, see the help text for parameter 'TransAuditDefault'.
        
    .PARAMETER NoAutoPlaycheck
        In AutoMode, will not generate Play Checks for the specified player Login and TransactionIDs.

    .PARAMETER NoAutoReconCheck
        In AutoMode, will not perform any Recon API checks for the specified player and TransactionIDs. 
        This includes Round Status checks for each Transaction ID's and checking Vanguard Queues for the Player.
        
        If you wish to set this parameter by default every time the cmdlet is run, see the help text for parameter 'ReconAPICheckDefault'.

    .PARAMETER NoCopyFilePath
        Disables copying the output ZIP file path to the clipboard.
        Useful if this cmdlet is called from another cmdlet or function, or when using this cmdlet in an automated manner.
        Copying the path to clipboard serves no purpose in this case, except to annoy anyone else using the computer while the cmdlet runs in the background.

    .PARAMETER NoGridView
        Disables use of the Out-GridView function to select files to be added to the ZIP archive.
        By default, Out-Gridview is used to allow the user to select specific files in the current REQ ticket folder, to be added to the ZIP archive.
        if this parameter is set, Out-Gridview is not used, instead every file inside the current REQ ticket folder is added to the ZIP without asking the user.
        Running this cmdlet using the alias 'za' (short for Zip All) will automatically set this parameter for you.

        This parameter has no effect in AutoMode.

    .PARAMETER NoMenu
        Does not display any menus to the user. Simply creates the folder if it doesn't exist, and if it does exist will zip up all files in the folder into a password protected archive.
        The NoGridView parameter will also be set.
        Running this cmdlet using the alias 'zx' will automatically set this parameter for you.

        This parameter has no effect in AutoMode.

    .PARAMETER OperatorID
        Specify the the OperatorID that is linked to the Casino, that the player belongs to for AutoMode.
        You must specify this parameter in conjunction with a Login parameter.

        This cmdlet will search for a list of Casino's that are linked to the specified OperatorID, via Casino Portal API.
        Each of these Casinos will be checked for a player that matches the provided Login.

        Only a single OperatorID is allowed. Wildcards are not permitted.

        The cmdlet will fail with an error, and provide no pipeline output, if the player cannot be located on any of the Casinos that are linked to the provided OperatorID.

    .PARAMETER PipelineOutput
        In AutoMode, will output results of the play check, game statistics and Vanguard Queues checks to pipeline.
        This is useful if you are calling this cmdlet from another function and want to act on its output.

    .PARAMETER ReconAPICheckDefault
        This parameter sets whether Reconciliation API checks will be performed by default in AutoMode.
        This includes Round Status checks for each Transaction ID's and checking Vanguard Queues for the Player.
        These checks will be performed by default until you set this parameter to False - e.g.
        -ReconAPICheckDefault $false

        To re-enable these checks, set this parameter to True - e.g.
        -ReconAPICheckDefault $true
        
        This setting is saved in the user registry and remembered when this cmdlet is run again in the future.
        If you don't want to save and remember this setting, you can instead set the parameter -NoAutoAudit which will only disable Reconciliation API checks for the current invocation of this cmdlet.

        The Reconciliation API features require PowerShell Core or PowerShell 7.

    .PARAMETER REQNumber
        The REQ number corresponding to a support ticket from the Canvas/Remedy system. This will be used to create a folder under the TicketPath and also used as the password for the ZIP archive.
        if you only enter a number, this function will automatically add the 'REQ' prefix for you.

        If you do not set this parameter, you will be asked if you want to re-open the previous REQNumber, or enter a new one.

    .PARAMETER SortOrder
        Sets the sorting order of Transaction and Financial audits.
        Valid settings for this parameter are "ASC" (ascending order) and "DESC" (descending order).
        Audits are sorted on the TransactionTime field.
        By default, audits are sorted in Descending order, so the newest transactions appear at the top of the report.

        This setting is saved in the user registry, so it will be remembered the next time you run this cmdlet. You don't need to specify it every time.

    .PARAMETER TicketPath
        Sets the full path where you wish to create the REQ ticket folders. When the appropriate menu option is selected, this file will be copied into the current REQ ticket folder, and automatically opened in its default program.
        By default, Powershell will create these folders under your profile, in the path '\OneDrive - Derivco (Pty) Limited\Documents\Tickets\'
        This parameter can be used to specify a custom path

        Once set, this will be saved in the user registry under "HKEY_CURRENT_USER\Software\QFPowershell\TicketPath" and will load this setting automatically whenever this function is invoked.
        There is no need to specify this parameter again once the correct path is set unless you need to change it.
        Deleting the above reg key will reset this parameter to default.

    .PARAMETER TransactionIDs
        Specify a list of TransactionID's for AutoMode.
        A Play Check will be generated for each TransactionID. A Game Statistics Report will also be generated for any games found in the play checks.
        You must also specify a Login or UserID parameter, plus a CasinoID, CasinoName or OperatorID parameter.

        You can specify a single TransactionID, or multiple seperated by commas (without any spaces.)
        You can also specify a range of transactionIDs using the Range operator syntax: (x..y)
        e.g. to generate play checks for every TransactionID between 10 to 15 inclusive, specify: (10..15)
    
    .PARAMETER TransAuditFile
        Sets the full path and filename of the Transaction Audit file. When the appropriate menu option is selected, this file will be copied into the current REQ ticket folder, and automatically opened in its default program.
        By default, Powershell will look for a file named 'Transaction_Audit.xlsx' in the same folder where this module file is loaded from; it will then open the copied file in Excel.
        This parameter can be used to specify a custom path and filename.

        Once set, this will be saved in the user registry under "HKEY_CURRENT_USER\Software\QFPowershell\TransAuditFile" and will load this setting automatically whenever this function is invoked.
        There is no need to specify this parameter again once the correct path and filename is set unless you need to change it.
        Deleting the above reg key will reset this parameter to default.

    .PARAMETER TransAuditDefault
        This parameter sets whether Transaction Audits will be generated by default in the Transaction Audit menu.
        Transaction Audits are enabled by default until you set this parameter to False. - e.g.
        -TransAuditDefault $false

        To re-enable these audits, set this parameter to True - e.g.
        -TransAuditDefault $true

        This setting is saved in the user registry and remembered when this cmdlet is run again in the future.
        If you don't want to save and remember this setting, you can instead set the parameter -NoAutoAudit which will only disable audit features for the current invocation of this cmdlet.

        Automated transaction and financial audit features require PowerShell Core or PowerShell 7.

    .PARAMETER UserID
        Specifies the UserID for Transaction and Financial audits in AutoMode.
        You must also specify a CasinoID, CasinoName or OperatorID parameter.
        You may optionally specify a Login parameter, but this is not required.
        
        Note that UserID's cannot be used for Play Checks. 
        This cmdlet will attempt to retrieve the player Login from the Transaction Audit data automatically, if a Login parameter is not provided.
        However, if the Transaction Audit fails, or retrieves no data, AutoMode will not generate any Play Checks in this case.

    .PARAMETER ZipFile
        The file name of the ZIP file you wish to create. Don't include the full path, just a file name. It will be created under the TicketPath folder.
        if not set, the zip file will be named 'GameData.zip' by default.
        if your filename doesn't include a .zip extension, it will be added automatically.
        This parameter is only effective for the current instance of the cmdlet. if you want to save this setting so it is effective every time you run this cmdlet, use ZipFileDefault parameter.
        if there is a saved default ZipFile setting, specifying this paramenter will override the default for the current instance of the cmdlet. Next time you run the cmdlet without specifying this parameter, it will revert to the saved default setting for the ZipFile name.

    .PARAMETER ZipFileDefault
        Will save the specified 'ZipFile' parameter in the user registry, under HKEY_CURRENT_USER\Software\QFPowershell\ZipFileDefault"
        This will remember the ZipFile setting every time you run this cmdlet without having to specify the ZipFile parameter. if you set the ZipFile parameter in the future, this will take precedence over the saved default file name but only for the current instance of the cmdlet.
        Deleting the above reg key will reset the ZipFile parameter to the default 'GameData.zip'.

    .INPUTS
        System.String
            You can pipe a string to this cmdlet that contains a Request Number corresponding to a support ticket from the Canvas/Remedy system, and a valid name for the Zip file to be created under the TicketPath folder.
            In AutoMode you can additionally pipe an object that contains properties for UserId, Login, CasinoID, CasinoName, and/or OperatorID.

    .OUTPUTS
        In Standard mode, this cmdlet does not provide any pipeline output.

        In AutoMode, (i.e. when Login, UserID, CasinoID, CasinoName, and/or OperatorID parameters are specified), if the PipelineOutput parameter is specified, this cmdlet will provide the following pipeline output:

        System.Management.Automation.PSCustomObject
        Name                MemberType      Definition
        ----                ----------      ----------
        ZipFile             NoteProperty    string
        Contents            NoteProperty    string[]
        Player              NoteProperty    PSCustomObject
        GameStatistics      NoteProperty    PSCustomObject
        QueueInfo           NoteProperty    PSCustomObject
        RoundInfo           NoteProperty    PSCustomObject

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding(DefaultParameterSetName = "Standard")]
    [alias("zz","za","zx")]
    param(
        [Parameter(Position = 5, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [datetime]$AuditEndDate,

        [Parameter(ParameterSetName="Standard")]
        [int]$AuditEndDateDefault,

        [Parameter(Position = 4, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [datetime]$AuditStartDate,

        [Parameter(ParameterSetName="Standard")]
        [int]$AuditStartDateDefault,


        [Parameter(ParameterSetName="Standard")]
        [switch]$AutoGameStats,

        <#
        AutoMode switch parameter is not actually required as setting any of the AutoMode parameters eg userid, casinoid, will enable it automatically
        [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName="AutoMode")]
        [switch]$AutoMode,
        #>

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=2,ParameterSetName="AutoMode")]
        [int[]]$CasinoID,

        [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName="AutoMode")]
        [string]$CasinoName,

        [Parameter(ParameterSetName="Standard")]
        [bool]$FinancialAuditDefault = $true,

        [Parameter(ParameterSetName="Standard")]
        [switch]$IDDQD,

        [Parameter(ValueFromPipelineByPropertyName=$true,Position=1,ParameterSetName="AutoMode")]
        [ValidateNotNullOrEmpty()]
        [string]$Login,

        [Parameter(ParameterSetName="AutoMode")]
        [switch]$NoAutoAudit,

        [Parameter(ParameterSetName="AutoMode")]
        [switch]$NoAutoPlaycheck,

        [Parameter(ParameterSetName="AutoMode")]
        [switch]$NoAutoReconCheck,

        [Parameter()]
        [switch]$NoCopyFilePath,

        [Parameter(ParameterSetName="Standard")]
        [switch]$NoGridView,

        [Parameter(ParameterSetName="Standard")]
        [switch]$NoMenu,

        [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName="AutoMode")]
        [int]$OperatorID,

        [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName="AutoMode")]
        [switch]$PipelineOutput,

        [Parameter(ParameterSetName="Standard")]
        [bool]$ReconAPICheckDefault = $true,

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            # Ensure there are no invalid characters for our folder name
            if ($_ -match '([\\/:"*?<>|]+|^ +$)') {
                Throw "$_ is not a valid folder name."
            }
            else {
                $true
            }
        }
        )]
        [string]$REQNumber,

        [Parameter()]
        [ValidateSet("ASC","DESC")]
        [string]$SortOrder,

        [Parameter(ParameterSetName="Standard")]
        [string]$TicketPath,

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=3,ParameterSetName="AutoMode")]
        [int[]]$TransactionIDs,

        [Parameter(ParameterSetName="Standard")]
        [string]$TransAuditFile,

        [Parameter(ParameterSetName="Standard")]
        [bool]$TransAuditDefault = $true,

        [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName="AutoMode")]
        [int]$UserID,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
        # Ensure there are no invalid characters for our zip file name
            if ($_ -match '([\\/:"*?<>|]+|^ +$)') {
                Throw "$_ is not a valid file name."
            }
            else {
                $true
            }
        }
        )]
        [string]$ZipFile,

        [Parameter(ParameterSetName="Standard")]
        [ValidateScript({
            # Ensure there are no invalid characters for our zip file name
            if ($_ -match '([\\/:"*?<>|]+|^ +$)') {
                Throw "$_ is not a valid file name."
            }
            else {
                $true
            }
        }
        )]
        [switch]$ZipFileDefault

    )
    begin {

        # Declare local functions that are internal to this cmdlet
        function Read-QFREQDataFile {
            # Local function to check if we have a settings file in this folder, and try to read the saved settings and output them to pipeline
            param(
                [switch]$GetLoginID, # if this is set we will ask user for the player LoginID
                [switch]$GetCasinoID, # if this is set we will ask user for the CasinoID
                [switch]$GetUserID # if this is set we will ask user for the player UserID
            )

            # try to read the saved settings into $REQFileData.
            # Then set $Login to the saved Login or prompt the user to enter it.
            if (Test-Path ".REQdata") {
                $REQFileData = $()
                $REQFileData = Get-Content .REQData -encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-JSON
                Write-Verbose ("[$(Get-Date)] REQ Settings File Data: $REQFileData")
            }

            $Output = @{}
            # Get the LoginID
            Try {
                if ($GetLoginID.IsPresent) {
                    if ($AutoMode) {
                        # AutoMode - don't prompt user at all, just give saved value
                        $Login = $REQFileData.Login
                    } elseif ($null -eq $REQFileData.Login -or $REQFileData.Login -eq "") {
                        # if we didn't get any data for login from the settings file, prompt the user
                        [string]$Login = Read-Host -Prompt "Please enter the player's full Login name"
                    } else {
                        # if we got data from the REQ settings file for login, offer that as a default option
                        [string]$Login = Read-Host -Prompt "Please enter the player's full Login name (ENTER for $($REQFileData.Login.Trim()))"
                        if ($Login -eq "") {
                            [string]$Login  = $REQFileData.Login.Trim()
                        } else {
                            # LoginID changed so clear out any saved UserID
                            $REQFileData.UserID = $null
                        }
                    }
                    Write-Verbose ("[$(Get-Date)] Login: $Login")
                    $Output = @{"Login" = $Login}
                }
            } catch {
                Continue MainMenu
            }

            # get the CasinoID
            try {
                if ($GetCasinoID.IsPresent) {
                    if ($AutoMode) {
                        # In Auto Mode just give the existing saved value without interaction
                        $CasinoID = $REQFileData.CasinoID
                    } elseif ($null -eq $REQFileData.CasinoID) {
                        # if we didn't get any data for casinoid from the settings file, prompt the user
                        [int]$CasinoID = Read-Host -Prompt "Please enter the CasinoID/ServerID"
                    } else {
                        # if we got data from the REQ settings file for casinoID, offer that as a default option
                        [int]$CasinoID = Read-Host -Prompt "Please enter the CasinoID/ServerID (ENTER for $($REQFileData.CasinoID))"
                        if ($CasinoID -eq 0) {
                            $CasinoID = $REQFileData.CasinoID
                        } else {
                            # CasinoID changed so clear out any saved GamingServerID
                            $REQFileData.GamingServerID = $null
                        }
                    }
                    # finally add the CasinoId to our output object
                    Write-Verbose ("[$(Get-Date)] CasinoID: $CasinoID")
                    $Output += @{ "CasinoID" = $CasinoID }
                } elseif ($REQFileData.CasinoID -gt 0) {
                    # Give saved value as output
                    $Output += @{"CasinoID" = $REQFileData.CasinoID}
                }
            } catch {
                Continue MainMenu
            }

            # Get the UserID, needed for API actions eg transaction audits
            try {
                if ($GetUserID.IsPresent) {
                    if ($AutoMode) {
                        # In AutoMode just give the existing saved value without interaction
                        $UserID = $REQFileData.UserID
                    } elseif ($null -eq $REQFileData.UserID -or $REQFileData.UserID -eq 0) {
                        # if we didn't get any data for UserID from the settings file, prompt the user
                        [int]$UserID = Read-Host -Prompt "Please enter the player's UserID"
                    } else {
                        # if we got data from the REQ settings file for UserID, offer that as a default option
                        [int]$UserID = Read-Host -Prompt "Please enter the player's UserID (ENTER for $($REQFileData.UserID))"
                        if ($UserID -eq 0) {$UserID = $REQFileData.UserID}
                    }
                    # finally add the UserID to our output object
                    Write-Verbose ("[$(Get-Date)] UserID: $UserID")
                    $Output += @{"UserID" = $UserID}
                } elseif ($null -ne $REQFileData.UserID -and $REQFileData.UserID -gt 0) {
                    # Give saved value as output
                    $Output += @{ "UserID" = $REQFileData.UserID }
                }
            } catch {
                Continue MainMenu
            }

            # If we got a GamingServerID, OperatorID or HostingSiteID from the REQ file add that to the output too
            $Output += @{ "GamingServerID" = $REQFileData.GamingServerID}
            $Output += @{ "HostingSiteID" = $REQFileData.HostingSiteID}
            $Output += @{ "OperatorID" = $REQFileData.OperatorID}

            # Output the details as a PSCustomObject to pipeline
            [PSCustomObject]$Output
        }


        function Save-QFREQDataFile {
            # Local function to save the Login and CasinoID to the settings file
            $REQFileData = $()
            if (Test-Path ".REQdata") {
                # Read the data from the file
                $REQFileData = Get-Content .REQData -Encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-JSON
                Write-Verbose ("[$(Get-Date)] REQ Settings File Data: $REQFileData")
                # Update the data in the existing settings file if a value is present in the $QFScriptLogin or $QFScriptCasinoID objects
                if ($null -ne $QFScriptLogin -and $QFScriptLogin -ne "") { $REQFileData | Add-Member -Name Login -Value $QFScriptLogin.Trim() -MemberType NoteProperty -force }
                if ($null -ne $QFScriptCasinoID -and $QFScriptCasinoID -gt 0) { $REQFileData | Add-Member -Name CasinoID -Value $QFScriptCasinoID -MemberType NoteProperty -force }
                # Don't check if these two are null, as we want to clear any saved values if user inputs new login/casinoID in Read-QFReqDataFile function
                $REQFileData | Add-Member -Name UserID -Value $QFScriptUserID -MemberType NoteProperty -force
                $REQFileData | Add-Member -Name GamingServerID -Value $QFScriptGamingServerID -MemberType NoteProperty -force
                $REQFileData | Add-Member -Name HostingSiteID -Value $QFScriptHostingSiteID -MemberType NoteProperty -force
                $REQFileData | Add-Member -Name OperatorId -Value $QFScriptOpID -MemberType NoteProperty -force
                <# some other properties we might want to save in the future....
                if ($null -ne $QFScriptSiteCode -and $QFScriptSiteCode -ne "") { $REQFileData | Add-Member -Name SiteCode -Value $QFScriptSiteCode -MemberType NoteProperty -force }
                if ($null -ne $QFScriptCasinoName -and $QFScriptCasinoName -ne "") { $REQFileData | Add-Member -Name CasinoName -Value $QFScriptCasinoName.Trim() -MemberType NoteProperty -force }
                if ($null -ne $QFScriptOpName -and $QFScriptOpName -ne "") { $REQFileData | Add-Member -Name OpName -Value $QFScriptOpName.Trim() -MemberType NoteProperty -force }
                #>
            } else {
                # if REQData file doesnt exist, initialise a new array object that we will save into the file
                $REQFileData = New-Object -TypeName PSObject -Property @{
                    Login = $QFScriptLogin
                    CasinoID = $QFScriptCasinoID
                    UserID = $QFScriptUserID
                    GamingServerID = $QFScriptGamingServerID
                    HostingSiteID = $QFScriptHostingSiteID
                    OperatorId = $QFScriptOpID
                }
            }

            # Save the settings file and set it to hidden
            Write-Verbose ("[$(Get-Date)] Updated REQ Settings File Data: $REQFileData")
            try {
                $REQFileData | ConvertTo-JSON | Set-Content .REQData -Encoding UTF8 -ErrorAction Stop
                (Get-ChildItem .REQData -Hidden -ErrorAction SilentlyContinue).Attributes = "Hidden"
            }
            catch {
                Write-Warning "Could not save the REQData file: $Infilepath.REQData"
            }
        }


        function Invoke-QFPlayCheck {
            # Local function to set up parameters for Get-QFPlayCheck, and automatically run Game Stats reports for any games found
            param(
                [switch]$SavePDF,
                [switch]$NoViewPDF
            )

            # Get the player ID data from .REQData file or ask user to input them
            $REQDataFile = Read-QFREQDataFile -GetCasinoID -GetLoginID
            $QFScriptLogin = $REQDataFile.Login
            $QFScriptCasinoID = $REQDataFile.CasinoID
            $QFScriptUserID = $REQDataFile.UserID
            $QFScriptGamingServerID = $REQDataFile.GamingServerID

            # This hash table will be parameters for splatting to Get-QFPlaycheck
            $PCArgs =
            @{
                Login = $QFScriptLogin
                CasinoID = $QFScriptCasinoID
                SavePDF = ($SavePDF.IsPresent)
                NoViewPDF = ($NoViewPDF.IsPresent)
            }

            Write-Verbose ("[$(Get-Date)] PCArgs Array:")
            foreach($k in $PCArgs.Keys) { Write-Verbose "$k $($PCArgs[$k])" }

            # Confirm player Login and casinoID was provided, otherwise return to the menu
            if ($QFScriptLogin -eq "" -or $null -eq $QFScriptLogin) {Continue MainMenu}
            if ($QFScriptCasinoID -eq 0 -or $null -eq $QFScriptCasinoID) {Continue MainMenu}
            # Handle first and last transactions for Range mode, or add TransactionID parameter to the PCArgs hash table for AutoMode
            If ($AutoMode) {
                $PCArgs.Add("TransID",($TransactionIDs))
            } elseif ($script:PlayCheckRangeMode) {
                Write-Host -ForegroundColor DarkGreen "Range mode - All transaction ID's between the two specified numbers will be play checked."
                Try {
                    [int]$TransIDStart = Read-Host "Enter FIRST Transaction ID of the range to play check"
                    [int]$TransIDEnd = Read-Host "Enter LAST Transaction ID of the range to play check"
                    $PCArgs.Add("TransID",($TransIDStart..$TransIDEnd))
                } catch {
                    # If non numeric value entered, exit the play check function
                    continue MainMenu
                }
            } else {
                # Otherwise just let the Get-QFPlaycheck function ask for the TransactionID's
                Write-Host "Enter Transaction ID's to play check. Press ENTER on an empty line to start generating the play checks."
            }
            try {
                # Return object from Get-QFPlayCheck will be stored in $PlayCheckGameArray, this will be a list of all games found in the playcheck data
                $PlayCheckGameArray = Get-QFPlayCheck @PCArgs

                Write-Verbose ("[$(Get-Date)] PlayCheck Game Array record count: $(@($PlayCheckGameArray).count)")
                if (@($PlayCheckGameArray).count -gt 0) {
                    Save-QFREQDataFile
                    # ask the user if they'd like to run game stats reports for these games too
                    Write-Host ""
                    Write-Host "Found these games in the play check data:"
                    foreach ($Game in $PlayCheckGameArray) { 
                        If ($Game.ETI) {Write-Host -ForegroundColor Yellow -NoNewline "ETI "}
                        Write-Host -ForegroundColor White $Game.GameName " MID: " $Game.MID " CID: " $Game.CID
                    }

                    <# 
                    # GameStats currently not working due to OKTA authentication - disabled as of version 1.6.2 17/1/2024 
                    # Uncomment this section to re-enable

                    If (!($AutoMode -or $AutoGameStats.IsPresent)) {
                        Write-Host "Hit SPACEBAR or ENTER to run Game Statistics reports for these games now, anything else to cancel:"
                        $Waitkey = [System.Console]::ReadKey()
                    }
                    Write-Host ""
                    if ($Waitkey.Key -eq "Enter" -or $Waitkey.Key -eq "Spacebar" -or $AutoGameStats.IsPresent -or $AutoMode) {
                        $GameStatResultsOutput = @()
                        Write-Host "Running Game Statistics reports, one moment please..."
                        $PlayCheckGameArray | ForEach-Object {
                            # Try to match the MID and CID for game stats reports, if that fails try GameName and CID
                            # Build a hash table of parameters for splatting to Invoke-QFGameStats - MID and CID
                            $GSArgs =
                            @{
                                Login = $QFScriptLogin
                                UserID = $QFScriptUserID
                                CasinoID = $QFScriptCasinoID
                                GamingSystemID = $QFScriptGamingServerID
                                MID = $_.MID
                                CID = $_.CID
                                SavePDF = ($SavePDF.IsPresent)
                                NoViewPDF = ($NoViewPDF.IsPresent)
                            }
                            Write-Verbose ("[$(Get-Date)] GSArgs Array (MID/CID):")
                            foreach($k in $GSArgs.Keys) { Write-Verbose "$k $($GSArgs[$k])" }
                            # Invoke-QFGameStats should return empty object if no reports found for player/game
                            $GameStatResults = Invoke-QFGameStats @GSArgs
                            If ($null -eq $GameStatResults -or @($GameStatResults).Count -eq 0) {
                                Write-Verbose ("[$(Get-Date)] Did not find any Game Stats matching specified MID... trying GameName instead")
                                # Build a hash table of parameters for splatting to Invoke-QFGameStats - GameName
                                $GSArgs =
                                @{
                                    Login = $QFScriptLogin
                                    UserID = $QFScriptUserID
                                    CasinoID = $QFScriptCasinoID
                                    GamingSystemID = $QFScriptGamingServerID
                                    GameName = $_.GameName
                                    SavePDF = ($SavePDF.IsPresent)
                                    NoViewPDF = ($NoViewPDF.IsPresent)
                                }
                                Write-Verbose ("[$(Get-Date)] GSArgs Array (GameName):")
                                foreach($k in $GSArgs.Keys) { Write-Verbose "$k $($GSArgs[$k])" }
                                # Invoke-QFGameStats should return empty object if no reports found for player/game
                                $GameStatResults = Invoke-QFGameStats @GSArgs
                            }
                            # Check again if we have anything in GameStatResults, if so add it to our output object
                            If ($null -eq $GameStatResults -or @($GameStatResults).Count -eq 0) {
                                Write-Host ""
                                Write-Host ("Failed to generate any Game Statistics reports for " + $_.GameName + " MID: " + $_.MID + " CID: " + $_.CID + "; Please run these manually.")
                            } else {
                                # Add ETI bool value to our output object
                                Add-Member -InputObject $GameStatResults -MemberType NoteProperty -Name 'ETI' -Value $_.ETI
                                $GameStatResultsOutput += $GameStatResults
                            }
                        }
                    }

                    # End of Game Stats section
                    #>
                }
            } catch {
                Write-Error $_.Exception.Message
                Continue MainMenu
            }
            Write-Host ""

            # Lookup support info for ETI games
            $ETIProviders = @()
            Foreach ($ETIGame in ($PlayCheckGameArray|Where-Object {$_.ETI})) {
                $ETIGameInfo = $null
                $ETIProviderInfo = $null

                # Lookup game info from the Casino Portal. We mainly need to check if this game has an ETI product ID number
                # First check by MID/CID then try by game name if we didn't find the MID
                $ETIGameInfo = Invoke-QFPortalRequest -MID $ETIGame.MID -CID $ETIGame.CID -ErrorAction SilentlyContinue
                If ($null -eq $ETIGameInfo) {
                    $ETIGameInfo = (Invoke-QFPortalRequest -GameName $ETIGame.GameName -ErrorAction SilentlyContinue) | Select-Object -First 1
                }
                If ($null -ne $ETIGameInfo.etiProductId) {
                    # Lookup the ETI ID number
                    $ETIProviderInfo = Get-QFETIProviderInfo -Id $ETIGameInfo.etiProductId -ErrorAction SilentlyContinue
                    If ($Null -eq $ETIProviderInfo) {
                        # Didn't find any ETI provider with that ID number, try searching by name
                        $ETIProviderInfo = Get-QFETIProviderInfo -Name $ETIGameInfo.provider -ErrorAction SilentlyContinue
                    }
                    # If we still didn't get any ETI info, don't output anything
                    If ($null -ne $ETIProviderInfo) {
                        Write-Host ""
                        Write-Host ("$([char]27)[36mETI Game: $([char]27)[0m" + $ETIGame.GameName + " $([char]27)[36mMID: $([char]27)[0m" + $ETIGame.MID + "$([char]27)[36m CID: $([char]27)[0m" + $ETIGame.CID)
                        Write-Host ("$([char]27)[36mETI Provider: $([char]27)[0m" + $ETIProviderInfo.ETIProvider)
                        If ($null -ne $ETIProviderInfo.Email -and $ETIProviderInfo.Email -ne "") {Write-Host ("$([char]27)[36mSupport Email: $([char]27)[0m" + $ETIProviderInfo.Email)}
                        If ($null -ne $ETIProviderInfo.PortalURI -and $ETIProviderInfo.PortalURI -ne "") {Write-Host ("$([char]27)[36mSupport Portal: $([char]27)[0m" + $ETIProviderInfo.PortalURI)}
                        If ($null -ne $ETIProviderInfo.PortalUsername -and $ETIProviderInfo.PortalUsername -ne "") {Write-Host ("$([char]27)[36mUsername: $([char]27)[0m" + $ETIProviderInfo.PortalUsername + "$([char]27)[36m Password: $([char]27)[0m" + $ETIProviderInfo.PortalPassword)}
                        Write-Host ""
                        $ETIProviderInfo | Add-Member -Name ETIGame -Value $ETIGame.GameName -MemberType NoteProperty
                        $ETIProviders += $ETIProviderInfo
                    }
                }
            }
            # Output QFGameStatsResult and ETIProviders objects to pipeline in AutoMode
            If ($AutoMode) {
                @{
                    GameStatResults = $GameStatResultsOutput
                    ETIProviders = $($ETIProviders | Sort-Object -Unique -Property ETIGame)
                }
            }
        }

        function Invoke-QFGameStats {
            # Local function to prompt user for Get-QFGameStats parameters into arrray $GSARGS, and then splat this to the Get-QFGameStats function.
            param(
                [string]$Login,
                [int]$MID,
                [int]$CID,
                [string]$GameName,
                [int]$UserID,
                [int]$CasinoID,
                [int]$GamingSystemID,
                [switch]$SavePDF,
                [switch]$OpenBrowser,
                [switch]$NoViewPDF
            )
            
            # If this function was called from Invoke-QFPlaycheck, login should be set already
            If ($null -eq $Login -or $Login -eq "") {
                # Call Read-QFReqDataFile function to read saved player info or prompt user
                $REQFileData = Read-QFREQDataFile -GetLoginID
                $QFScriptLogin = $REQFileData.Login
                $QFScriptUserID = $REQFileData.UserID
                $QFScriptCasinoID = $REQFileData.CasinoID
                $QFScriptGamingServerID = $REQFileData.GamingServerID

                # Gamename, MID and CID aren't saved in the settings file so just prompt the user for these
                [string]$GameName = Read-Host -Prompt "Please enter Game Name to search for (ENTER for none)"
                try {
                    [int]$MID = Read-Host -Prompt "Please enter Game ModuleID to generate a report for (ENTER for none)"
                    [int]$CID = Read-Host -Prompt "Please enter Game ClientID to generate a report for (ENTER for none)"
                } catch {
                    Continue MainMenu
                }
            } else {
                # Set these objects to the provided function parameter values. These objects are used by Save-QFReqData
                $QFScriptLogin = $Login
                $QFScriptUserID = $UserID
                $QFScriptCasinoID = $CasinoID
                # sorry about the inconsistent naming... QFPortal uses GamingServerID but Game Monitoring page uses GamingSystemID for the same value.. blame TechOps?
                $QFScriptGamingServerID = $GamingSystemID
            }

            # Build the array of Get-QFGameStats parameters. Try using Login parameter first.
            $GSArgs =
            @{
                Login = $QFScriptLogin
                GameName = $(if (($null -ne $GameName) -or ($GameName -ne "")) {$GameName.trim()})
                MID = $(if (($null -ne $MID) -or ($MID -ne 0)) {$MID})
                CID = $(if (($null -ne $CID) -or ($CID -ne 0)) {$CID})
                SavePDF = ($SavePDF.IsPresent)
                NoViewPDF = ($NoViewPDF.IsPresent)
                OpenBrowser = ($OpenBrowser.IsPresent)
            }

            if ($null -ne $QFScriptGamingServerID -and $QFScriptGamingServerID -ne 0) {
                $GSArgs.Add('GamingSystemID',$QFScriptGamingServerID)
            }

            Write-Verbose ("[$(Get-Date)] GSArgs Array:")
            foreach($k in $GSArgs.Keys) { Write-Verbose "$k $($GSArgs[$k])" }

            # call QFGameStats and splat the parameter values in the GSArgs array
            try {
                $GameStatsData = Get-QFGameStats @GSArgs
            } catch {
                # If we get a No Gaming Systems found error, try using the UserID, CasinoID and GamingServerID if we have them, instead of Login
                if ($_.exception.message -match "No Gaming Systems found*") {
                    Write-Verbose ("[$(Get-Date)] No Gaming Systems found for this player Login.")
                } else {
                    # A different error occured than "No Gaming Systems found", so print the error message and go back to main menu
                    Write-Error $_.Exception.Message
                    Continue MainMenu
                }
            }

            # If no game stats reports returned, Check we have userid/casinoid/gamingsystemid and if so try game stats function again
            if ($null -eq $GameStatsData -and ($null -ne $QFScriptCasinoID -and $QFScriptCasinoID -gt 0) -and
            ($null -ne $QFScriptUserID -and $QFScriptUserID -gt 0) -and
            ($null -ne $QFScriptGamingServerID -and $QFScriptGamingServerID -gt 0)) {
                Write-Verbose ("[$(Get-Date)] Trying Game Stats function again with UserID, CasinoID and GamingSystemID...")
                $GSArgs.Remove('Login')
                $GSArgs.Add('UserID',$QFScriptUserID)
                $GSArgs.Add('CasinoID',$QFScriptCasinoID)

                Write-Verbose ("[$(Get-Date)] GSArgs Array:")
                foreach($k in $GSArgs.Keys) { Write-Verbose "$k $($GSArgs[$k])" }

                # Try the Game Stats function again
                try {
                    $GameStatsData = Get-QFGameStats @GSArgs
                } catch {
                    Write-Error $_.Exception.Message
                    Continue MainMenu
                }
            }

            If ($null -eq $GameStatsData) {
                Write-Verbose ("[$(Get-Date)] No game statistics reports found for the specified player and/or games.")
            } else {
                # Returned array should contain casinoID and GamingSystemID
                If ($null -ne $($GameStatsData.CasinoID |Sort-Object -Unique) -and $($GameStatsData.CasinoID |Sort-Object -Unique) -gt 0) {
                    [int]$QFScriptCasinoID = $GameStatsData.CasinoID | Sort-Object -Unique
                }
                If ($null -ne $($GameStatsData.GamingSystemID |Sort-Object -Unique) -and $($GameStatsData.GamingSystemID |Sort-Object -Unique) -gt 0) {
                    [int]$QFScriptGamingServerID = $GameStatsData.GamingSystemID | Sort-Object -Unique
                }
                # Save-QFREQDataFile function will inherit the QFScriptCasinoID and QFScriptLogin objects from this function, and save their values into .REQData file
                Save-QFREQDataFile
                # Output the Get-QFGameStats results to pipeline. This should be null if no reports successfully created
                $GameStatsData
            }
        }

        function Invoke-QFTransAuditMenu {
            # Local function to present a menu of options for generating transaction/financial audits. In AutoMode will not show menus, will just run the audits without interactions.

            If (!$AutoMode) {
                $TransAuditMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Run &Audit',"Generate audit and export into Excel using specified options"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Start Date',"Adjust the Audit Start Date - only transactions older than this date will be included in the audit"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&End Date',"Adjust the Audit End Date - only transactions newer than this date will be included in the audit"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Audit &Type',"Choose to run a Transaction audit, Financial audit, or both."))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&ModuleID',"Filter Transaction audit by the specified game ModuleID. No effect on Financial audits"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Sort &Order',"Toggle between Ascending and Descending sort order on the TransactionTime field."))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Open Audit &File',"Open the audit spreadsheet file if it exists, otherwise copy a blank file if TransAuditFile parameter was set"))
                $TransAuditMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exits without any further action. Re-run this function with the same REQ number to resume"))
            }

            do {
                If ($AutoMode) {
                    # in AutoMode set the menu option to run the audit without interaction.
                    $TransAuditMenuChoice = 1
                } else {
                    # Display configured audit options and menu if not in AutoMode
                    Write-Host ""
                    Write-Host -ForegroundColor Green -NoNewline "Start Date: "; Write-Host -ForegroundColor White -NoNewline "$(Get-Date $script:AuditStartDate -Format 'yyyy-MM-dd HH:mm:ss')`t"
                    Write-Host -ForegroundColor Green -NoNewline "End Date: "; Write-Host -ForegroundColor White "$(Get-Date $script:AuditEndDate -Format 'yyyy-MM-dd HH:mm:ss')"
                    Write-Host -ForegroundColor Green -NoNewline "Audit Type: "
                    If ($script:DoFinAudit -and $script:DoTransAudit) {
                        Write-Host -ForegroundColor White "Both Transaction and Financial Audits"
                    } elseif ($script:DoFinAudit -and !($script:DoTransAudit)) {
                        Write-Host -ForegroundColor White "Financial Audit Only"
                    } elseif (!($script:DoFinAudit) -and $script:DoTransAudit) {
                        Write-Host -ForegroundColor White "Transaction Audit Only"
                    } else { Write-Host -ForegroundColor White "NO Audits!!! WTF mate" }
                    Write-Host -ForegroundColor Green -NoNewline "Sort Order: "
                    If ($script:SortAscendingToggle) {
                        Write-Host -ForegroundColor White -NoNewline "Ascending`t`t"
                    } else  {
                        Write-Host -ForegroundColor White -NoNewline "Descending`t`t"
                    }
                    If ($script:FilterModuleID -gt 0 -and $script:DoTransAudit) {
                        Write-Host -ForegroundColor Green -NoNewline "Filter By ModuleID: "
                        Write-Host -ForegroundColor White -NoNewline $script:FilterModuleID
                    }
                    Write-Host ""
                    $TransAuditMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$("$REQNumber - Audit")$([char]27)[0m", "Hit ENTER to return to the previous menu. Press Q to quit.", $TransAuditMenu, 0)
                    Write-Host ""
                }
                
                switch ($TransAuditMenuChoice) {
                    # Case statement for handling the Audit submenu options. In AutoMode this should always be set to 1
                    0 { return } # Back to previous menu
                    1 { # Run the transaction audit with specified options
                        # Get the userid and casinoID
                        Try {
                            $REQData = Read-QFREQDataFile -GetUserID -GetCasinoID
                        } catch {
                            Write-Error "Failed to retrieve UserID and CasinoID - cannot continue."
                            Write-Error $_.Exception.Message
                            return
                        }
                        $QFScriptCasinoID = $REQData.CasinoID
                        If ($null -eq $REQData.UserID -or $REQData.UserID -eq "") {
                            return
                        } else {
                            $QFScriptUserID = $REQData.UserID
                        }
                        try {
                            Write-Host "Please wait while retrieving audit data from the API..."
                            # Get the casino/operator details
                            $CasinoData = Invoke-QFPortalRequest -CasinoID $QFScriptCasinoID
                            # Record the gaming serverID, required for playchecks with UserID
                            [int]$QFScriptGamingServerID = $CasinoData.GamingServerID
                            [int]$QFScriptOpId = $CasinoData.OperatorID
                            [int]$QFScriptHostingSiteId = $CasinoData.HostingSiteId
                            Write-Verbose ("[$(Get-Date)] Gaming System ID: $QFScriptGamingServerID")
                            Write-Verbose ("[$(Get-Date)] Hosting Site ID: $QFScriptHostingSiteId")
                            # Record the casino username prefix
                            [string]$CasinoPrefix = $CasinoData.productSettings | Where-Object {$_.name.trim() -like "Register - SGI JIT Account Creation Prefix"}| Select-Object -expandProperty StringValue
                            $CasinoPrefix = $CasinoPrefix.Trim()
                            # Get the API key for this casino/operator. Script scoped so parent functions can read it too
                            Write-Verbose ("[$(Get-Date)] OperatorID: $QFScriptOpId")
                            $script:QFAPIKey = $null
                            $script:QFAPIKey = (Get-QFOperatorAPIKeys -OperatorID $QFScriptOpId).APIKey | Select-Object -First 1
                        } catch {
                            Write-Error "Failed to retrieve operator API key - cannot continue."
                            Write-Error $_.Exception.Message
                            return
                        }
                        # Check we actually got an API key
                        If ($null -eq $script:QFAPIKey) {
                            Write-Error "Couldn't get an API Key for CasinoID $QFScriptCasinoID - check credentials for Operator Security Site are valid, and the operator has generated a key."
                            Return
                        }
                        Write-Verbose ("[$(Get-Date)] API Key: $script:QFAPIKey")
                        # Now get an Operator Token
                        Try {
                            $APIToken = (Get-QFOperatorToken -APIKey $script:QFAPIKey).AccessToken
                        } catch {
                            Write-Error "Failed to generate an operator API Token - cannot continue."
                            Write-Error $_.Exception.Message
                            return
                        }
                        # Make the API request for the trans audit data
                        If ($script:DoTransAudit) {
                            Write-Verbose ("[$(Get-Date)] Beginning transaction audit...")
                            Try {
                                $TransAuditData = $null
                                # Build hashtable of QFAudit Parameters
                                $TransAuditParams = @{
                                    Token               = $APIToken
                                    HostingSiteID       = $CasinoData.HostingSiteID
                                    UserID              = $QFScriptUserID
                                    CasinoID            = $QFScriptCasinoID
                                    TransactionAudit    = $true
                                    StartDate           = $script:AuditStartDate 
                                    EndDate             = $script:AuditEndDate
                                    SortAscending       = $script:SortAscendingToggle
                                }

                                # Add a parameter for ModuleID if set
                                If ($script:FilterModuleID -gt 0) {
                                    $TransAuditParams.Add("ModuleID",$script:FilterModuleID)
                                }

                                # Call the Help Desk Express API and splat TransAuditParams object
                                $TransAuditData = Get-QFAudit @TransAuditParams
                            } catch {
                                Write-Warning "Failed to retrieve Transaction Audit data from Reconciliation API. Please ensure you have network connectivity."
                                Write-Warning $_.exception.message
                            }
                            # Confirm we actually got some data back from the API
                            Write-Verbose ("[$(Get-Date)] Retrieved $(@($TransAuditData).Count) records from the API.")
                            If ($null -eq $TransAuditData) {
                                Write-Warning "Did not retrieve any records for this Transaction Audit - please confirm player/casino details are correct, or adjust the ModuleID filter and date range."
                            } else {
                                try {
                                    Export-QFExcel -ExcelData $TransAuditData -ExcelSourceWorksheetName "Transaction Audit" -ExcelDestWorksheetName "Transaction Audit" -ColourRange "A7:R7" -DateFormatRange "I:J" -NumberFormatRange "A:A","E:E","G:G","O:O"
                                    # Try to get the Login aka UserName from the Transaction Audit
                                    [string]$QFScriptLogin = $CasinoPrefix + ($TransAuditData | Select-Object -ExpandProperty userName -First 1)
                                    $QFScriptLogin = $QFScriptLogin.trim()
                                    Write-Verbose ("[$(Get-Date)] Found Player Login: $QFScriptLogin ")
                                } catch {
                                    Write-Warning "Failed to export Transaction Audit data into an Excel file. Please ensure you don't have an existing Excel file open."
                                    Write-Warning $_.exception.message
                                }
                            }
                        }
                        Start-Sleep 1 # 1 sec delay to ensure the file save completes

                        If ($script:DoFinAudit) {
                            Write-Verbose ("[$(Get-Date)] Beginning financial audit...")
                            try {
                                $FinAuditData = $null
                                # Build hashtable of QFAudit Parameters
                                $FinAuditParams = @{
                                    Token               = $APIToken
                                    HostingSiteID       = $CasinoData.HostingSiteID
                                    UserID              = $QFScriptUserID
                                    CasinoID            = $QFScriptCasinoID
                                    FinancialAudit      = $true
                                    StartDate           = $script:AuditStartDate 
                                    EndDate             = $script:AuditEndDate
                                    SortAscending       = $script:SortAscendingToggle
                                }
                                
                                # Call the Help Desk Express API and splat the FinAuditParams object
                                $FinAuditData = Get-QFAudit @FinAuditParams
                            } catch {
                                Write-Warning "Failed to retrieve Financial Audit data from Reconciliation API. Please ensure you have network connectivity."
                                Write-Warning $_.exception.message
                            }
                            # Confirm we actually got some data back from the API
                            Write-Verbose ("[$(Get-Date)] Retrieved $(@($FinAuditData).Count) records from the API.")
                            If ($null -eq $FinAuditData) {
                                Write-Warning "Did not retrieve any records for this Financial Audit - please confirm player/casino details are correct, or adjust the date range."
                            } else {
                                try {
                                    Export-QFExcel -ExcelData $FinAuditData -ExcelSourceWorksheetName "Financial Audit" -ExcelDestWorksheetName "Financial Audit" -ColourRange "A7:I7" -DateFormatRange "A:A" -NumberFormatRange "C:C","F:G"
                                } catch {
                                    Write-Warning "Failed to export Financial Audit data into an Excel file. Please ensure you don't have an existing Excel file open."
                                    Write-Warning $_.exception.message
                                }
                            }
                        }
                        Start-Sleep 1 # 1 sec delay to ensure the file save completes

                        # open the spreadsheet in excel if it exists
                        If (!$AutoMode -and (Test-Path -Path ".\Transaction_Audit.xlsx" -PathType Leaf)) {
                            Start-Process ".\Transaction_Audit.xlsx"
                        }
                        # Finally save the user/casino data
                        Save-QFREQDataFile
                    }
                    2 {
                        # Start Date option
                        Try {
                            [datetime]$script:AuditStartDate = Read-Host "Please enter an Audit Start Date & Time in the format YYYY-MM-DD HH:MM:SS`n(Time is optional, will default to 00:00:00 if not specified)"
                        } catch {
                            Write-Warning "Not a valid date - Audit Start Date unchanged from $(Get-date $script:AuditStartDate -Format 'yyyy-MM-dd HH:mm:ss')"
                        }
                    }
                    3 {
                        # End Date option
                        Try {
                            [datetime]$script:AuditEndDate = Read-Host "Please enter an Audit End Date & Time in the format YYYY-MM-DD HH:MM:SS`n(Time is optional, will default to 23:59:59 if not specified)"
                            # Set the time to 11:59:59 if only date was specified
                            If ((Get-date $script:AuditEndDate -Format 'HH:mm:ss') -eq "00:00:00") {[datetime]$script:AuditEndDate = "$(Get-Date $script:AuditEndDate -format 'yyyy-MM-dd') 23:59:59"}
                        } catch {
                            Write-Warning "Not a valid date - Audit End Date unchanged from $(Get-date $script:AuditEndDate -Format 'yyyy-MM-dd HH:mm:ss')"
                        }
                    }
                    4 {
                        # Audit Type option
                        If ($script:DoFinAudit -and $script:DoTransAudit) {
                            # set Trans Audit Only
                            $script:DoFinAudit = $false
                        } elseif ($script:DoFinAudit -and !($script:DoTransAudit)) {
                            # set Both audits
                            $script:DoTransAudit = $true
                        } elseif (!($script:DoFinAudit) -and $script:DoTransAudit) {
                            # set Fin audit only
                            $script:DoFinAudit = $true
                            $script:DoTransAudit = $false
                        } else {
                            # both options were false... reset to true; should never get here
                            $script:DoTransAudit = $true
                            $script:DoFinAudit = $true
                        }
                    }
                    5 {
                        # ModuleID filter option
                        Try {
                            [int]$script:FilterModuleID = Read-Host -Prompt "Enter the ModuleID to filter Transaction Audit by (enter 0 or blank to disable filter)"
                        } catch {
                            # non-integer value entered so just set to 0 which disables the ModuleID filter
                            $script:FilterModuleID = 0
                        }
                    }
                    6 {
                        # Sort order option
                        If ($script:SortAscendingToggle) {
                            $script:SortAscendingToggle = $false 
                        } else {
                            $script:SortAscendingToggle = $true
                        }
                    }
                    7 { # open the spreadsheet in excel if it exists
                        If ((!(Test-Path -Path ".\Transaction_Audit.xlsx" -PathType Leaf)) -and (Test-Path -Path $TransAuditFile -PathType Leaf)) {
                            # copy blank audit spreadsheet file if one doesn't exist
                            Write-Verbose ("[$(Get-Date)] Transaction Audit File doesn't exist - will copy a new file over and open it.")
                            Copy-Item $TransAuditFile "Transaction_Audit.xlsx"
                        }
                        if (Test-Path -Path ".\Transaction_Audit.xlsx" -PathType Leaf) {
                            # Open the transaction audit file
                            Start-Process ".\Transaction_Audit.xlsx" -ErrorAction SilentlyContinue
                        } else {
                            Write-Host "Transaction audit file doesn't exist... try running an audit first!"
                        }
                    }
                    8 { return "Q" } # Quit the module entirely.
                }
            } until ($TransAuditMenuChoice -eq 0 -or $AutoMode)
        }


        function Invoke-AutoMode {
            # Local function for Auto Mode. Intended to fully automate the play check/transaction audit process, without interaction.
            # User and transaction details will be passed as parameters, instead of displaying menus or asking for these details.

            # First, if we don't have the UserID and CasinoID we need to find them
            # Can specify multiple CasinoIDs, so this object could be an array or a single int
            If ([int]$UserId -eq 0 -or $CasinoID.count -ne 1) {
                Write-Host "Searching for the player..."
                # If we don't have a Login either we will just exit. Not enough info to work with!
                If ($null -eq $Login -or $Login.trim() -eq "") {
                    Throw {"Please provide a player Login, or both a UserID and a single CasinoID to run an automated Play Check / Transaction Audit."}
                }


                # Set up hashtable of parameters for Search-QFUser, based on provided parameters when New-QFTicket was run
                $SearchParams = @{
                    Login = $Login.trim()
                }
                # We also need either an OperatorID, CasinoID or CasinoName to search for a player. Add to the QFUser parameters hashtable
                If ($null -ne $CasinoName -and $CasinoName.trim() -ne "") {
                    $SearchParams.Add("CasinoName",$CasinoName.trim())
                } elseif ($null -ne $CasinoID) {
                    $SearchParams.Add("CasinoId",$CasinoID)
                } elseif ([int]$OperatorID -ne 0) {
                    $SearchParams.Add("OperatorId",$OperatorID)
                } else {
                    Throw "Not enough player information provided to run an automated Play Check / Transaction Audit. Please provide a player Login, plus either a CasinoID, CasinoName or OperatorID."
                }

                # Kick off the player search
                Write-Verbose ("[$(Get-Date)] Search-QFUser Parameters: Login: $($SearchParams["Login"]) CasinoID: $($SearchParams["CasinoID"]) CasinoName: $($SearchParams["CasinoName"]) OpID: $($SearchParams["OperatorID"])")
                Try {
                    $PlayerData = Search-QFUser @SearchParams
                } Catch {
                    Write-Warning $_.Exception.Message
                }
                Write-Verbose ("[$(Get-Date)] Search-QFUser Results: $($PlayerData)")
                # Check we only got one player
                If (@($PlayerData).Count -eq 0) {
                    Throw "Unable to locate any players using the provided information."
                } elseif (@($PlayerData).Count -gt 1) {
                    Throw "Found multiple players matching the provided information. Please try again and specify the correct player's UserID and CasinoID."
                }
                Write-Host -ForegroundColor White "Found player - $([char]27)[36mUserID:$([char]27)[0m $($PlayerData.UserID) $([char]27)[36mCasinoID:$([char]27)[0m $($PlayerData.CasinoID) - $($PlayerData.CasinoName) $([char]27)[36mSite:$([char]27)[0m $($PlayerData.GamingSystem)"
            }

            # At this point we should have enough player info to proceed, store this in the objects used by Save-QFREQDataFile and other functions in this cmdlet
            If ($null -ne $Login -and $Login.trim() -ne "") {
                $QFScriptLogin = $Login.trim()
                # If we have a Login, check if it includes the Casino Login Prefix
                if ($null -ne $PlayerData) {
                    If ([string]$PlayerData.LoginPrefix.trim() -ne "") {
                        # if not, add the prefix to the login
                        If ($QFScriptLogin.Substring(0,$(($PlayerData.LoginPrefix.trim()).Length)) -ne $PlayerData.LoginPrefix.trim()) {
                            $QFScriptLogin = $PlayerData.LoginPrefix.trim() + $Login.trim()
                        }
                    }
                }
            }

            If ([int]$PlayerData.UserID -ne 0) {
                $QFScriptUserID = [int]$PlayerData.UserID
            } elseif ([int]$UserID -ne 0) {
                $QFScriptUserID = [int]$UserID
            }
            
            If ([int]$PlayerData.CasinoID -ne 0) {
                $QFScriptCasinoID = [int]$PlayerData.CasinoID
            } elseif ($CasinoID.count -le 1) {
                # If a single CasinoID parameter was provided, use that to proceed. This will also pick up if no CasinoID was specified ($casinoID.count will be 0)
                If ($null -eq $CasinoID) {
                    # No CasinoID parameter was provided
                    Throw "No CasinoID found for the specified player. Please try specifying the CasinoID parameter."
                } else {
                    # One CasinoID was provided
                    $QFScriptCasinoID = [int]$CasinoID[0]
                }
            } else {
                # If an array of CasinoID's was specified and we didn't find the player in any of them, we won't proceed
                Throw "Could not locate this player on any of the specified CasinoIDs. Please specify a single CasinoID or try searching by Casino Name."
            }

            # VSCode might whinge that this is assigned but never used, in fact it is read by child functions called from this function, such as Save-QFReqDataFile
            If ([int]$PlayerData.GamingServerID -ne 0) {[int]$QFScriptGamingServerID = [int]$PlayerData.GamingServerID}
            If ([int]$PlayerData.OperatorID -ne 0) {[int]$QFScriptOpID = [int]$PlayerData.OperatorID}
            If ([int]$PlayerData.HostingSiteID -ne 0) {[int]$QFScriptHostingSiteID = [int]$PlayerData.HostingSiteID}

            # Update the .REQdata file with our player info - the function can read the $QFScript objects from this parent function
            Save-QFREQDataFile
           
            # If we have a UserID and CasinoID try to run a transaction/financial audit
            If (($null -ne $QFScriptCasinoID -and $null -ne $QFScriptUserID) -and !$NoAutoAudit) {
                Write-Verbose ("[$(Get-Date)] Beginning Transaction Audits - UserID: $QFScriptUserID CasinoID: $QFScriptCasinoID")
                Invoke-QFTransAuditMenu 
            }

            # If we have TransactionID's and a Login, try a playcheck and gamestats
            $PlayerData = Read-QFREQDataFile -GetLoginID
            If (($null -ne $TransactionIDs -and $null -ne $PlayerData.Login -and $PlayerData.Login -ne "") -and !$NoAutoPlaycheck) {
                # If game stats are generated, the report results should go straight out to pipeline. The parent function can save this in an object
                Invoke-QFPlayCheck -SavePDF -NoViewPDF
            }

            # If we got an API key from the transaction audit, check for queued transactions for this player
            if ($NoAutoReconCheck -or $DoReconAPICheck -eq $false) {
                Write-Verbose ("[$(Get-Date)] Auto Recon Checks disabled - will skip Recon API checks.")
            } Elseif ($null -eq $script:QFAPIKey -or [int]$PlayerData.HostingSiteID -eq 0) {
                Write-Verbose ("[$(Get-Date)] No Operator API key or Hosting Site ID - will skip Recon API checks.")
            } else {
                if ($null -ne $TransactionIDs) {
                    Write-Verbose ("[$(Get-Date)] Checking Round Status for UserID $QFScriptUserID CasinoID $QFScriptCasinoID HostingSiteID $($PlayerData.HostingSiteID)")
                    try {
                        $APIToken = Get-QFOperatorToken -APIKey $script:QFAPIKey
                        Write-Host "Checking Round Status for each TransactionID..."
                        $RoundInfo = Invoke-QFReconAPIRequest -Token $APIToken.AccessToken -HostingSiteID $PlayerData.HostingSiteID -UserID $QFScriptUserID -CasinoID $QFScriptCasinoID -RoundInfo -TransactionIDs $TransactionIDs
                        If ($RoundInfo.Count -le 0) {
                            Throw "Did not receive any Round Status data for this Player's Transactions."
                        } else {
                            # Output round info to pipeline and display on screen
                            $RoundInfo | ForEach-Object {
                                If ($null -ne $_.transactionInfo) {
                                    # numberOfEvents is how many records are in the transactionInfo object, eg free spins can have many wagers and winnings in one Transaction.
                                    $_ | Add-Member -Name 'numberOfEvents' -value $_.transactionInfo.count -MemberType NoteProperty
                                } else {
                                    # If nothing in Transaction Info (e.g. Unknown status for the round) add NumberOfEvents member with value 0
                                    $_ | Add-Member -Name 'numberOfEvents' -value 0 -MemberType NoteProperty
                                }
                            }
                            $RoundInfo | Format-List -Property TransactionNumber,RoundStatusName,WinAmount,Currency,numberOfEvents | Out-Host
                            @{RoundInfo = $RoundInfo | Select-Object -Property TransactionNumber,RoundStatusName,WinAmount,Currency,numberOfEvents}
                        }
                    } catch {
                        Write-Warning $_.Exception.Message
                    }
                }
                Write-Verbose ("[$(Get-Date)] Checking Vanguard queues for UserID $QFScriptUserID CasinoID $QFScriptCasinoID HostingSiteID $($PlayerData.HostingSiteID)")
                try {
                    $APIToken = Get-QFOperatorToken -APIKey $script:QFAPIKey
                    Write-Host "Checking Vanguard Commit and Rollback queues..."
                    $QueueInfo = Invoke-QFReconAPIRequest -Token $APIToken.AccessToken -HostingSiteID $PlayerData.HostingSiteID -UserID $QFScriptUserID -CasinoID $QFScriptCasinoID -QueueInfo
                    If ($QueueInfo.CommitCount -gt 0) {
                        Write-Host -ForegroundColor DarkYellow "Player has $($QueueInfo.CommitCount) transaction(s) in the Vanguard Commit queue!"
                        $QueueInfo.CommitQueue | Select-Object TransactionNumber,DateCreated,gameName | Out-Host
                    } else {
                        Write-Host "Player has no transactions in the Vanguard Commit queue."
                    }
                    If ($QueueInfo.RollbackCount -gt 0) {
                        Write-Host -ForegroundColor DarkYellow "Player has $($QueueInfo.rollbackCount) transaction(s) in the Vanguard Rollback queue!"
                        $QueueInfo.RollbackQueue | Select-Object TransactionNumber,DateCreated,gameName | Out-Host
                    } else {
                        Write-Host "Player has no transactions in the Vanguard Rollback queue."
                    }
                    # Output queue info to pipeline
                    @{QueueInfo = $QueueInfo}
                } catch {
                    Write-Warning "Unable to check Vanguard Queues for this player."
                    Write-Warning $_.Exception.Message
                }
            }
        }

        # End of local functions

        # Check registry key exists in registry
        try {
            If (!(Test-Path -Path HKCU:\Software\QFPowershell -PathType Container)) {
                New-Item -Path HKCU:\Software\ -Name QFPowershell -Force | Out-Null
            }
        } catch {
            Write-Warning "Unable to create registry key HKCU:\Software\QFPowershell - Please check for registry permissions issues or try creating this key manually."
        }

        # Auto update check
        Update-QFPowerShell

        # Check if AutoMode parameters were passed
        If ($psCmdlet.ParameterSetName -eq "AutoMode") {
            $AutoMode = $true
        }

        # Check if the REQ number was passed on the command line
        if (!($PSBoundParameters.ContainsKey('REQNumber'))) {
            If ($AutoMode) {
                # In Auto Mode we need the REQ number parameter to be set
                Throw "REQNumber parameter is not set."
            }
            # No REQ Number parameter set so see if we can read the last value from registry
            Write-Verbose ("[$(Get-Date)] No REQNumber parameter set, so attempting to read LastREQNumber from registry...")
            [string]$REQNumber = (Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name LastREQNumber -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastREQNumber)
            If ($null -eq $REQNumber -or $REQNumber.Trim() -eq "") {
                # If no value retrieved for REQNumber, prompt the user for the value
                [string]$REQNumber = Read-Host -Prompt "Please enter the REQ ticket number"
            } else {
                # We do have a value for REQnumber so offer it as a default
                Write-Verbose ("[$(Get-Date)] Retrieved REQ Number from registry: $REQNumber")
                $Prompt = "Please enter the REQ ticket number (Or press ENTER to reuse previous value - " + "$([char]27)[36m$($REQNumber)$([char]27)[0m)"
                [string]$REQNumberTemp = Read-Host -Prompt $Prompt
                # If user just hit ENTER and gave an empty value, we'll leave REQNumber as it is, otherwise we'll set it to the value they just entered
                If (!($null -eq $REQNumberTemp -or $REQNumberTemp.Trim() -eq "" -or $REQNumberTemp.Trim().ToLower() -eq "o" -or $REQNumberTemp.Trim().ToLower() -eq "c")) {
                    $REQNumber = $REQNumberTemp
                }
            }
        }

        # Sanitize the REQNumber parameter
        $REQNumber = $REQNumber.ToUpper().Trim()
        # if only numbers were entered, add the REQ prefix
        if ($REQNumber -match "^[0-9]+$") {
            $REQNumber = "REQ" + $REQNumber
        }
        # Finally check that the REQNumber parameter is valid
        if (($REQNumber -match '([\\/:"*?<>|]+|^ +$)') -or $REQNumber -eq "") {
            Throw "The REQNumber parameter: $REQNumber is not a valid folder name name."
        }

        Write-Verbose ("[$(Get-Date)] REQ Number: $REQNumber")

        # Path to the empty Transaction Audit XLSX file
        if ($PSBoundParameters.ContainsKey('TransAuditFile')) {
            if (Test-Path -Path $TransAuditFile -PathType Leaf) {
                # Check the file exists, then save the parameter into the registry
                Write-Verbose ("[$(Get-Date)] Saving TransAuditFile into registry: $TransAuditFile")
                Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name TransAuditFile -Value $TransAuditFile.Trim()
            }
            else {
                Throw "TransAuditFile parameter: $TransAuditFile is not valid, please confirm path is correct and file exists."
            }
        }
        # Read the trans audit file path from registry if it exists
        try {
            $TransAuditFile = (Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name TransAuditFile -ErrorAction Stop | Select-Object -ExpandProperty TransAuditFile).Trim()
        }
        catch {
            # Otherwise just set it to the default file path
            $TransAuditFile = $($PSScriptRoot -replace "\\src$","") + "\Transaction_Audit.xlsx"
        }
        Write-Verbose ("[$(Get-Date)] TransAuditFile: $TransAuditFile")

        # Parent folder where you want to create the REQ ticket folders.
        if ($PSBoundParameters.ContainsKey('TicketPath')) {
            # Check the folder exists, then save the parameter into the registry
            if (Test-Path -Path $TicketPath -PathType Container) {
                # Check for a trailing slash and add if its missing
                if (!($TicketPath.Trim() -match "\\$")) {
                    $TicketPath = $TicketPath.Trim() + "\"
                }
                Write-Verbose ("[$(Get-Date)] Saving TicketPath into registry: $TicketPath")
                try {
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name TicketPath -Value $TicketPath.Trim() -ErrorAction Stop | Out-Null
                }
                catch {
                   Write-Warning "Unable to save the Ticket Path into the registry!"
                }

            }
            else {
                Throw "TicketPath parameter: $TicketPath is not valid, please confirm path is correct."
            }
        }
        # Read the ticket path from registry if it exists
        try {
            $TicketPath = (Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name TicketPath -ErrorAction Stop | Select-Object -ExpandProperty TicketPath).Trim()
        }
        catch {
            # Otherwise just set it to the default path setting
            $TicketPath = $Env:UserProfile + "\OneDrive - Derivco (Pty) Limited\Documents\Tickets\"
        }
        Write-Verbose ("[$(Get-Date)] TicketPath: $TicketPath")

        # Get the path for the 7zip executable file from registry
        try {
            $7z = (Get-ItemProperty -Path HKLM:\Software\7-zip -name Path -ErrorAction Stop | Select-Object -ExpandProperty Path) + "7z.exe"
            # Check 7z exe exists
            Test-Path -Path $7z -PathType Leaf -ErrorAction Stop | Out-Null
        }
        catch {
            Throw "Cannot find 7zip executable! Please ensure it is installed correctly."
        }
        Write-Verbose ("[$(Get-Date)] 7zip path: $7z")

        Invoke-Expression $([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("aWYgKC" +
        "RJRERRRC5Jc1ByZXNlbnQpIHtJbnZva2UtRXhwcmVzc2lvbiAkKFtTeXN0ZW0uVGV4dC5FbmNvZGluZ106OlVURjguR2V0U3RyaW" +
        "5nKFtTeXN0ZW0uQ29udmVydF06OkZyb21CYXNlNjRTdHJpbmcoIkppQnpkR0Z5ZENBbmFIUjBjSE02THk5M2QzY3VlVzkxZEhWaV" +
        "pTNWpiMjB2ZDJGMFkyZy9kajE0ZGtaYWFtODFVR2RITUNjPSIpKSl9")))

        # Check the ZIP file parameter and ensure its a valid filename
        if ($PSBoundParameters.ContainsKey('ZipFile')) {
            Write-Verbose ("[$(Get-Date)] Zip file parameter specified: $ZipFile")
            if ($ZipFile.trim() -notmatch "^[\w\-. ]+$") {
                Throw "$ZipFile is not a valid output file name."
            }
            else {
                # Check if .zip extension is present in the specified filename, append it if not
                if ($ZipFile.toLower().trim() -notmatch "\.zip$") {
                    $ZipFile = $ZipFile.trim() + ".zip"
                }
                # Path and file name of the zip file to create
                $OutFile = $TicketPath + $REQNumber + "\" + $ZipFile.trim()

                # if ZipFileDefault parameter is set, save the ZipFile value for future reference
                if ($PSBoundParameters.ContainsKey('ZipFileDefault')) {
                    try {
                        Write-Verbose ("[$(Get-Date)] Saving default ZipFile setting into registry: $ZipFile")
                        Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name ZipFile -ErrorAction Stop -Value $ZipFile.Trim()
                    }
                    catch {
                        Write-Warning "Unable to save the Default Zip File Name into the registry!"
                    }

                }
            }
        }
        else {
             # Read the ZipFile setting from registry if it exists
            try {
                $ZipFile = (Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name ZipFile -ErrorAction Stop | Select-Object -ExpandProperty ZipFile).Trim()
            }
            catch {
                # Otherwise just Set the zip filename to the default value
                $ZipFile = "GameData.zip"
            }
        }

        # Path and file name of the zip file to create
        $OutFile = $TicketPath + $REQNumber + "\" + $ZipFile
        Write-Verbose ("[$(Get-Date)] Output Zip Filename: $OutFile")

        # Path and file name of the source files to add to the zip archive
        # by default it will look in the same directory as the outfile path, you can change it here if you desire
        # Include a trailing slash
        $InFilePath = $TicketPath + $REQNumber + "\"

        # Catch user selecting open or copy menu option instead of entering a REQ number
        If ($null -ne $REQNumberTemp) {
            If ($REQNumberTemp.Trim().ToLower() -eq "o") {
                Start-Process $InFilePath
            } elseif ($REQNumberTemp.Trim().ToLower() -eq "c") {
                Set-Clipboard $InFilePath
            }
        }

        # Get the basename of the output file
        $ZipFile -match '(^.+)\.zip$' | Out-Null
        $ZipFileBase = $Matches[1]


        # Transaction Audit default settings - start/end date in days; audit type
        # If PowerShell 5 skip this section as its not compatible
        If ($PSVersionTable.PSVersion.Major -gt 5) {
            # Reconciliation API enable/disable default setting
            if ($PSBoundParameters.ContainsKey('ReconAPICheckDefault')) {
                Write-Verbose ("[$(Get-Date)] Saving ReconAPICheckDefault into registry: $ReconAPICheckDefault")
                try {
                    [bool]$DoReconAPICheck = $ReconAPICheckDefault
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name ReconAPICheckDefault -Value ([int]$ReconAPICheckDefault) -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Unable to save the ReconAPICheckDefault setting into the registry!"
                }
            } else {
                    # Read the ReconAPICheckDefault setting from registry if it exists. have to convert bool value to int to save, and back again to load
                try {
                    [int]$ReconAPICheckDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name ReconAPICheckDefault -ErrorAction Stop | Select-Object -ExpandProperty ReconAPICheckDefault
                    [bool]$DoReconAPICheck = $ReconAPICheckDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $DoReconAPICheck = $true
                }
            }
            Write-Verbose ("[$(Get-Date)] ReconAPICheckDefault: $([bool]$ReconAPICheckDefault)")

            if ($PSBoundParameters.ContainsKey('AuditStartDateDefault')) {
                Write-Verbose ("[$(Get-Date)] Saving AuditStartDateDefault into registry: $AuditStartDateDefault")
                    try {
                        Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditStartDateDefault -Value $AuditStartDateDefault -ErrorAction Stop | Out-Null
                    }
                    catch {
                    Write-Warning "Unable to save the AuditStartDateDefault setting into the registry!"
                    }
            } else {
                # Read the audit start date setting from registry if it exists
                try {
                    $AuditStartDateDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditStartDateDefault -ErrorAction Stop | Select-Object -ExpandProperty AuditStartDateDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $AuditStartDateDefault = 14
                }
            }
            Write-Verbose ("[$(Get-Date)] AuditStartDateDefault: $AuditStartDateDefault")

            if ($PSBoundParameters.ContainsKey('AuditEndDateDefault')) {
                Write-Verbose ("[$(Get-Date)] Saving AuditEndDateDefault into registry: $AuditEndDateDefault")
                try {
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditEndDateDefault -Value $AuditEndDateDefault -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Unable to save the AuditEndDateDefault setting into the registry!"
                }
            } else {
                    # Read the audit end date setting from registry if it exists
                try {
                    $AuditEndDateDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditEndDateDefault -ErrorAction Stop | Select-Object -ExpandProperty AuditEndDateDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $AuditEndDateDefault = 0
                }
            }
            Write-Verbose ("[$(Get-Date)] AuditEndDateDefault: $AuditEndDateDefault")

            # Transaction Audit enable/disable default setting
            if ($PSBoundParameters.ContainsKey('TransAuditDefault')) {
                Write-Verbose ("[$(Get-Date)] Saving TransAuditDefault into registry: $TransAuditDefault")
                try {
                    [bool]$script:DoTransAudit = $TransAuditDefault
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name TransAuditDefault -Value ([int]$TransAuditDefault) -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Unable to save the TransAuditDefault setting into the registry!"
                }
            } else {
                    # Read the audit end date setting from registry if it exists. have to convert bool value to int to save, and back again to load
                try {
                    [int]$TransAuditDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name TransAuditDefault -ErrorAction Stop | Select-Object -ExpandProperty TransAuditDefault
                    [bool]$script:DoTransAudit = $TransAuditDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $script:DoTransAudit = $true
                }
            }
            Write-Verbose ("[$(Get-Date)] TransAuditDefault: $([bool]$DoTransAudit)")

            # Financial Audit enable/disable default setting
            if ($PSBoundParameters.ContainsKey('FinancialAuditDefault')) {
                Write-Verbose ("[$(Get-Date)] Saving FinancialAuditDefault into registry: $FinancialAuditDefault")
                try {
                    [bool]$script:DoFinAudit = $FinancialAuditDefault
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name FinancialAuditDefault -Value ([int]$FinancialAuditDefault) -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Unable to save the FinancialAuditDefault setting into the registry!"
                }
            } else {
                    # Read the audit end date setting from registry if it exists.  have to convert bool value to int to save, and back again to load
                try {
                    [int]$FinancialAuditDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name FinancialAuditDefault -ErrorAction Stop | Select-Object -ExpandProperty FinancialAuditDefault
                    [bool]$script:DoFinAudit = $FinancialAuditDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $script:DoFinAudit = $true
                }
            }
            Write-Verbose ("[$(Get-Date)] FinancialAuditDefault: $([bool]$DoFinAudit)")

            # Audit sort order setting
            if ($PSBoundParameters.ContainsKey('SortOrder')) {
                Write-Verbose ("[$(Get-Date)] SortOrder parameter set, saving into registry...")
                If ($SortOrder -like "ASC") {
                    $script:SortAscendingToggle = $true
                } else {
                    $script:SortAscendingToggle = $false
                }
                try {
                    Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditSortAscending -Value ([int]$script:SortAscendingToggle) -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Unable to save the SortOrder setting into the registry!"
                }
            } else {
                    # Read the audit end date setting from registry if it exists. have to convert bool value to int to save, and back again to load
                try {
                    [int]$SortOrderDefault = Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name AuditSortAscending -ErrorAction Stop | Select-Object -ExpandProperty AuditSortAscending
                    [bool]$script:SortAscendingToggle = $SortOrderDefault
                }
                catch {
                    # Otherwise just set it to the default
                    $script:SortAscendingToggle = $false
                }
            }

            # Set up default audit options. Use data provided as parameters, otherwise check the saved values in the registry
            # Script scoped so the options are remembered if you invoke the transaction audit menu function, exit and return
            If ($null -ne $AuditStartDate) {
                $script:AuditStartDate = $AuditStartDate
            } else {
                $script:AuditStartDate = (Get-Date (Get-Date).AddDays(-$AuditStartDateDefault) -Format "yyyy-MM-dd")
            }
            if ($null -ne $AuditEndDate) {
                $script:AuditEndDate = $AuditEndDate
            } else {
                $script:AuditEndDate = "$(Get-Date (Get-Date).AddDays(-$AuditEndDateDefault) -format 'yyyy-MM-dd') 23:59:59"
            }
            Write-Verbose ("[$(Get-Date)] Audit date range: $script:AuditStartDate to $script:AuditEndDate")
            # If trans and financial audit are both disabled, it is futile to run the audit function!
            If (!($script:DoTransAudit) -and !($script:DoFinAudit) -and $PSVersionTable.PSVersion.Major -gt 5) {
                Write-Warning ('Both Transaction and Financial audits are disabled! Run this cmdlet again and set TransAuditDefault and/or FinancialAuditDefault parameters to $true to re-enable audits via API.')
            }

            # Clear the ModuleID filter for transaction audits, so it doesn't remember the setting after exiting and re-running the cmdlet
            $script:FilterModuleID =0
        }

        # Read the ticket path from registry if it exists
        try {
            $TicketPath = (Get-ItemProperty -Path HKCU:\Software\QFPowershell -Name TicketPath -ErrorAction Stop | Select-Object -ExpandProperty TicketPath).Trim()
        }
        catch {
            # Otherwise just set it to the default path setting
            $TicketPath = $Env:UserProfile + "\OneDrive - Derivco (Pty) Limited\Documents\Tickets\"
        }
        Write-Verbose ("[$(Get-Date)] TicketPath: $TicketPath")
    }

    
    process {
        # Try-Finally block is used to put us back at the same folder location where we ran the command from
        try {

            # Check the source folder exists.
            if (!(Test-Path -Path $InFilePath -PathType Container)) {
                # Create the source folder if it doesn't exist
                try {
                    New-Item $InFilePath -ItemType Directory -ErrorAction Stop | Out-Null
                }
                catch {
                    Throw "Unable to create a folder named with the specified REQ Number. Please ensure the TicketPath exists and you have write permission to it."
                }

                Write-Host ""
                Write-Host "Created New Folder: $InFilePath"
            }
            else {
                Write-Host ""
                Write-Host "Folder already exists: $InFilePath"
            }

            # Save the current REQNumber into the registry
            Try {
                Set-ItemProperty -Path HKCU:\Software\QFPowershell -Name LastREQNumber -Value $REQNumber.Trim() -ErrorAction Stop | Out-Null
            } Catch {
                Write-Warning "Failed to save the current REQNumber into the registry at HKCU:\Software\QFPowershell\LastREQNumber"
            }

            # in AutoMode, instead of displaying menus etc the player details are provided as parameters.
            If ($AutoMode) {
                Push-Location $InFilePath
                $AutoModeGameStatResults = Invoke-AutoMode
            }

            # Skip the menus for AutoMode, if the NoMenu option is present or the cmdlet was run with the zx alias
            if (!($AutoMode) -and (!($NoMenu.IsPresent)) -and ($($psCmdlet.myinvocation.line) -notmatch "^zx")) {
                # Sets up the PromptForChoice method for the main menu, provides user a list of options
                $MainMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Zip',"Proceed to add files in the source directory to the ZIP file"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit', "Exits without creating ZIP file. Re-run this function with the same REQ number to resume"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Open folder',"Open the folder in Windows Explorer to view files"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Copy path',"Copies the current REQ Folder path to the clipboard"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&PlayCheck',"Generate a new Play Check report"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Game Stats',"Generate a new Game Monitor / Game Statistics report"))
                $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&SA Password',"Request the SA password for a SQL server host. This ticket number ($REQNumber) will be provided as the reason."))


                # Transaction Audit Menu item
                # If Powershell core or 7 give the full menu with API transactions if at least one audit type enabled
                # Otherwise just add the option for a new transaction audit if the XLSX file is present
                if (($PSVersionTable.PSVersion.Major -gt 5 -and ($script:DoTransAudit -or $script:DoFinAudit)) -or (Test-Path -Path $TransAuditFile -PathType Leaf)) {
                    $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Transaction Audit',"Copies the Transaction Audit XLSX file into the folder and opens it in Excel"))
                }

                # API functions menu - only available in powershell core/7
                if ($PSVersionTable.PSVersion.Major -gt 5) {
                    $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&API Functions',"Perform requests to Reconciliation API or Casino Portal API"))
                }

                # Sets up the PromptForChoice method for the PlayCheck submenu
                $PlaycheckMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Open in Browser',"Generates a new Play Check report, and opens in your default browser"))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&PDF',"Generates a new Play Check report, and saves it as a PDF, then opens in the default PDF viewer"))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Silent PDF',"Generates a new Play Check report, and saves it as a PDF without opening in a PDF viewer"))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Format PlayCheck',"Reformats all MHTML PlayCheck files in the current folder, so all content is on a single page. This is performed automatically when saving playchecks to PDF."))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Range Mode',"Range Mode generates play checks for all Transaction ID's between two provided numbers. The default mode lets you specify non-sequential Transaction ID's."))
                $PlaycheckMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exits without creating ZIP file. Re-run this function with the same REQ number to resume"))

                # Sets up the PromptForChoice method for the GameStats submenu
                $GameStatsMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $GameStatsMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
                $GameStatsMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Open in Browser',"Generates a new Game Statistics report, and opens in your default browser"))
                $GameStatsMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&PDF',"Generates a new Game Statistics report, and saves it as a PDF, then opens in the default PDF viewer"))
                $GameStatsMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Silent PDF',"Generates a new Game Statistics  report, and saves it as a PDF without opening in a PDF viewer"))
                $GameStatsMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exits without creating ZIP file. Re-run this function with the same REQ number to resume"))

                # Ask if the user is ready to continue, and create the ZIP file. They have the option to quit the script, or open the source folder in Windows Explorer.
                $MainMenuChoice = $null
                $PlaycheckMenuChoice = $null
                $GameStatsMenuChoice = $null
                :MainMenu do {
                    $MainMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$($REQNumber)$([char]27)[0m", "Hit ENTER when the files in the source folder are ready to add to the ZIP file. Press Q to quit.", $MainMenu, 0)
                    Write-Host ""
                    switch ($MainMenuChoice) {
                        # case statement for handling the menu options
                        1 { return } #  quit the function
                        2 { Start-Process $InFilePath } # Opens folder window in Explorer
                        3 { Set-Clipboard $InFilePath } # Copies folder path to clipboard
                        4 {
                            # Opens the PlayCheck submenu
                            Push-Location $InFilePath
                            Do {
                                $PlaycheckMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$($REQNumber) - Play Checks$(if ($script:PlayCheckRangeMode) {"$([char]27)[32m - Range Mode Enabled"}) $([char]27)[0m", 
                                "Please select a Play Check option. Press B to go back or Q to quit.", $PlaycheckMenu, 0)
                                Write-Host ""
                                Switch ($PlaycheckMenuChoice) {
                                    # Case statement for handling the Playcheck submenu options
                                    # 0 - do nothing, go back to the last menu
                                    1 { Invoke-QFPlayCheck } # Open playcheck in web browser
                                    2 { Invoke-QFPlayCheck -SavePDF } # Save playcheck as PDF
                                    3 { Invoke-QFPlayCheck -SavePDF -NoViewPDF } # Save playcheck as PDF, don't open in PDF viewer
                                    4 { Format-QFPlayCheck *.mhtml } # Format Playcheck - this is now done automatically by Get-QFPlaycheck
                                    5 {
                                        # Range Mode Toggle
                                        if ($script:PlayCheckRangeMode) {
                                            $script:PlayCheckRangeMode = $false
                                        } else {
                                            $script:PlayCheckRangeMode = $true
                                        }
                                    }
                                    6 {
                                        #  quit the function completely
                                        Pop-Location
                                        return
                                    }
                                }
                            } while (
                                # This keeps us in the playcheck menu if range mode toggle is selected
                                $PlaycheckMenuChoice -eq 5
                                )
                            Pop-Location
                        }
                        5 {
                            # Opens the Game Stats Report submenu
                            Push-Location $InFilePath
                            $GameStatsMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$($REQNumber) - Game Statistics$([char]27)[0m", "Please select a Game Statistics Report option. Press B to go back or Q to quit.", $GameStatsMenu, 0)
                            Write-Host ""
                            switch ($GameStatsMenuChoice) {
                                # Case statement for handling the Playcheck submenu options
                                # 0 - Do nothing, go back to the last menu
                                1 { $GameStatResults = Invoke-QFGameStats -OpenBrowser} # Open game stats report in web browser.
                                2 { $GameStatResults = Invoke-QFGameStats -SavePDF } # Save game stats report as PDF.
                                3 { $GameStatResults = Invoke-QFGameStats -SavePDF -NoViewPDF }# Save game stats report as PDF, don't open in PDF viewer.
                                4 {
                                    #  quit the function completely
                                    Pop-Location
                                    return
                                }
                            }
                            Pop-Location
                        }
                        6 {
                            # Request the SA password for a SQL server
                            Push-Location $InFilePath
                            Write-Host "Requesting SA Password..."
                            Try{
                                Get-QFSQLServerSAPassword -Reason $REQNumber
                            }
                            Catch {
                                $Exception = $_.Exception
                                Write-Warning "An error occured attempting to request the SQL server SA Password."
                                Write-Warning $Exception.Message
                            }
                            Write-Host ""
                            Pop-Location
                        }
                        7 { # Transaction Audit Menu
                            if ($PSVersionTable.PSVersion.Major -gt 5 -and ($script:DoTransAudit -or $script:DoFinAudit)) {
                                # If powershell core or 7, give full trans audit menu using API
                                Push-Location $InFilePath
                                $ReturnVal = Invoke-QFTransAuditMenu
                                Pop-Location
                                # If user picked Quit option then exit this main function as well
                                If ($ReturnVal -eq "Q") { return }
                            } elseif (Test-Path -Path $TransAuditFile -PathType Leaf) {
                                # If only powershell 5, just give basic menu to to copy blank audit spreadsheet file
                                # Strip the path from the Transaction Audit file name
                                $TAFileName = ($TransAuditFile -split "\\")[-1]
                                if (Test-Path -Path "$InFilePath\$TAFileName" -PathType Leaf) {
                                    Write-Host "A Transaction Audit file already exists in the current folder!"
                                    Write-Host "Do you want to overwrite it with a blank copy? (Y to confirm)"
                                    $Waitkey = [System.Console]::ReadKey()
                                    Write-Host ""
                                    if ($Waitkey.Key -eq "Y") {
                                        Copy-Item $TransAuditFile $InFilePath -Force
                                    }
                                }
                                else {
                                Copy-Item $TransAuditFile $InFilePath
                                }
                                # Open the transaction audit file
                                Start-Process "$InFilePath\$TAFileName"
                            }
                        }
                        8 { # API Functions menu
                            # If we already have some player data, splat these parameters to Invoke-QFMenu function for use as default values
                            Push-Location $InFilePath                          
                            $QFMenuData = @{
                                REQNumber = $REQNumber
                            }
                            if (Test-Path ".REQdata") {
                                $PlayerData = $()
                                $PlayerData = Get-Content .REQData -encoding UTF8 -ErrorAction SilentlyContinue | ConvertFrom-JSON
                            
                                If ($null -ne $PlayerData.Login) {$QFMenuData.Add("Login",$PlayerData.Login)}
                                If ($null -ne $PlayerData.CasinoId) {$QFMenuData.Add("CasinoId",$PlayerData.CasinoId)}
                                If ($null -ne $PlayerData.UserId) {$QFMenuData.Add("UserId",$PlayerData.UserId)}
                                If ($null -ne $PlayerData.HostingSiteId) {$QFMenuData.Add("HostingSiteId",$PlayerData.HostingSiteId)}
                                If ($null -ne $PlayerData.OperatorID) {$QFMenuData.Add("OperatorID",$PlayerData.OperatorID)}
                            }
                            
                            Write-Verbose "[$(Get-Date)] Invoke-QFMenu Parameters:"
                            foreach($k in $QFMenuData.Keys) { Write-Verbose "$k $($QFMenuData[$k])" }

                            Invoke-QFMenu @QFMenuData
                            Pop-Location
                            if ($global:QFMenuExitFlag) {Return}
                        }
                    }
                }
                until
                (
                    $MainMenuChoice -eq 0
                )

            }
            # if the NoGridView parameter is set, or the cmdlet was called using the 'za' or 'zx' alias, just zip up everything in the folder without showing the GridView file selection dialog to the user
            if (($NoGridView.IsPresent) -or ($NoMenu.IsPresent) -or ($AutoMode) -or ($($psCmdlet.myinvocation.line) -match "^z[ax]")) {
                Write-Verbose "[$(Get-Date)] Zipping up all files without GridView..."
                # Don't use gridview to select source files, just get every file in the source directory, exclude the output ZIP file and the REQData settings file
                $InFilesList = Get-ChildItem -File -Path $InFilePath -Exclude @($ZipFile, $($ZipFileBase + "*.zip"),".REQData") -Recurse | Select-Object -ExpandProperty Name
            }
            else {
                # Use gridview to select source files, exclude the output ZIP file and the REQData settings file
                Write-Verbose "[$(Get-Date)] Zipping up files using selection from GridView..."
                $InFilesListTemp =  Get-ChildItem -File -Path $InFilePath -Exclude @($ZipFile, $($ZipFileBase + "*.zip"),".REQData") -Recurse | Select-Object Name, LastWriteTime, CreationTime | Sort-Object CreationTime -Descending

                # if more than one file in the source folder, use the GridView to select a list of files
                if ($null -eq $InFilesListTemp.Count -or $InFilesListTemp.Count -le 1) {
                    $InFilesList = $InFilesListTemp | Select-Object -ExpandProperty Name
                }
                else {
                    $InFilesList = $InFilesListTemp | Out-GridView -OutputMode Multiple -Title "Please select files to add to the zip archive" | Select-Object -ExpandProperty Name
                }
            }

            # Check we actually have some source files to zip up, or if the user clicked Cancel on the gridview
            if ($InFilesList.count -eq 0) {
                Write-Host "No files to add to ZIP archive."
                return
            }

            # Check if an output zip file already exists
            If (Test-Path $OutFile -PathType Leaf) {
                Write-Verbose "[$(Get-Date)] Output file already exists - renaming and keeping up to 9 old copies..."
                # If so we will rename the existing file with an underscore and a number up to 8
                $Filecount = @()
                Foreach($File in $(Get-ChildItem -Path "$InFilePath" -File -Filter $($ZipFileBase + "_?.zip"))) {
                    Write-Verbose "[$(Get-Date)] Found existing file: $File"
                    $File -match '.+_([1-8])\.zip' | Out-Null
                    $Filecount += $Matches[1]
                }
                If (@($Filecount).Count -gt 0) {
                    [int]$FileCountMax = ($Filecount | Measure-Object -Maximum).Maximum
                    Write-Verbose "[$(Get-Date)] FileCountMax: $FileCountMax"
                    # Loop through the files, increment the number on end of each one. Delete any older file if already existing.
                    Do {
                        Write-Verbose "[$(Get-Date)] Remaining files to rename: $FileCountMax"
                        If (Test-Path $("$InFilePath" + $ZipFileBase + "_" + ([int]$FileCountMax + 1 ) + ".zip") -PathType Leaf) {
                            Write-Warning "Old zip file $("$InFilePath" + $ZipFileBase + "_" + ([int]$FileCountMax + 1 ) + ".zip") already exists and will be overwritten."
                            Remove-Item -Path $("$InFilePath" + $ZipFileBase + "_" + ([int]$FileCountMax + 1 ) + ".zip") -Force -ErrorAction SilentlyContinue
                        }
                        Write-Verbose "[$(Get-Date)] Renaming: $("$InFilePath" + $ZipFileBase + "_" + $FileCountMax + ".zip") to $("$InFilePath" + $ZipFileBase + "_" + ([int]$FileCountMax + 1 ) + ".zip")"
                        Rename-Item -Path $("$InFilePath" + $ZipFileBase + "_" + $FileCountMax + ".zip")  -NewName $("$InFilePath" + $ZipFileBase + "_" + ([int]$FileCountMax + 1 ) + ".zip") -Force -ErrorAction SilentlyContinue
                        [int]$FileCountMax = $FileCountMax - 1
                    }
                    Until (
                        $FileCountMax -eq 0
                    )
                }
                # Rename the GameData file to GameData_1.zip
                Write-Verbose "[$(Get-Date)] Renaming: $OutFile to $($InFilePath + $ZipFileBase + "_1" + ".zip")"
                Rename-Item -Path "$OutFile" -NewName $($InFilePath + $ZipFileBase + "_1" + ".zip") -Force -ErrorAction SilentlyContinue
            }

            # This will be an array of strings containing file names successfully added to zip
            $ZipFileContents = @()

            # Loop through the input files list and try to add them to the zip archive
            foreach ($FileName in $InFilesList) {
                $InFile = $InFilePath + $FileName
                Write-Verbose ("[$(Get-Date)] Input Filename: $InFile")
                # Check the file isn't locked which will prevent it getting added to the zip
                try {
                    $FileStream = [System.IO.File]::Open($InFile,'Open','Write')
                    $FileStream.Close()
                    $FileStream.Dispose()
                }
                catch {
                    Write-Warning "Unable to add $FileName to zip archive! Ensure it is not open in another program."
                    continue
                }

                # Zip up the files using the REQ number as the password
                try {
                    Write-Host ""
                    Write-Host "Adding file $FileName to ZIP archive..."
                    & $7z a -mx7 -sse -p"$REQNumber" "$OutFile" "$InFile" | Out-Null
                    # If we got this far the file should be in the ZIP, add it to our list of successfully added files
                    $ZipFileContents += $FileName
                }
                catch {
                    Throw "Failed to create the ZIP file $OutFile"
                }
            }

            if (Test-Path -Path $OutFile -PathType Leaf) {
                if ($AutoMode -and $PipelineOutput) {
                    # If running AutoMode we will output the ZIP file name and list of contents to pipeline, plus player and game stats info
                    $QFScriptOutput = @{
                        ZipFile = $OutFile
                        Contents = $ZipFileContents
                        Player = Read-QFREQDataFile -GetLoginID -GetCasinoID -GetUserID
                    }
                    If ($null -ne $AutoModeGameStatResults.GameStatResults) {$QFScriptOutput.Add("GameStatistics",$AutoModeGameStatResults.GameStatResults)}
                    If ($null -ne $AutoModeGameStatResults.ETIProviders) {$QFScriptOutput.Add("ETIProviders",$AutoModeGameStatResults.ETIProviders)}
                    If ($null -ne $AutoModeGameStatResults.QueueInfo) {$QFScriptOutput.Add("QueueInfo",$AutoModeGameStatResults.QueueInfo)}
                    If ($null -ne $AutoModeGameStatResults.RoundInfo) {$QFScriptOutput.Add("RoundInfo",$AutoModeGameStatResults.RoundInfo)}

                    [PSCustomObject]$QFScriptOutput
                } else {
                    $FileSize = $(Get-ChildItem $Outfile).Length
                    Write-Verbose ("[$(Get-Date)] Output file size: $FileSize")
                    # If the file size is over 9MB warn user (Remedy has attachment size limit of 10MB, MIME encoding can bloat this out)
                    if ($FileSize -gt "9437184" -and (!($NoMenu.IsPresent)) -and (!($AutoMode)) -and (!($($psCmdlet.myinvocation.line) -match "^zx"))) {
                        Write-Host ""
                        Write-Host "This Zip file is rather large - $([math]::Round($FileSize / 1048576,2)) MB"
                        Write-Host "You may not be able to attach it to Remedy."
                        Write-Host "You can upload it to CrushFTP https://crush.gameassists.co.uk/ and share with the customer."
                    }
                }

                Write-Host ""
                Write-Host "Created ZIP file: $OutFile"
                Write-Host "Password: $REQNumber"

                If (!($NoCopyFilePath.IsPresent)) {
                    Set-Clipboard -Value $OutFile
                    Write-Host "The file path has been copied to the clipboard."
                } 
                
            } elseif ($AutoMode) {
                # If running AutoMode throw an error if the ZIP file wasn't created
                Throw "An error occured, could not create an archive of the Play Check data for this REQNumber."
            }
        }
        Finally {
            # This puts us back in the folder we original ran the script from
            Pop-Location
        }
    }
}


function Update-QFPowerShell {
    <#
    .SYNOPSIS
        Automatically updates the QFPowerShell module using Git.

    .DESCRIPTION
        Automatically updates the QFPowerShell module. This function requires Git to be installed,
        and the module files must have been cloned from a Git repository.
        This function runs a 'git pull' to download the latest files from the current repository.

        Please refer to the Readme.md file of this module - section 'Auto update using git', or visit this link:
        https://dev.azure.com/Derivco/Software/_git/MIGS-IT-QFPowershell?anchor=auto-update-using-git

        This function will attempt an update every week by default. It will store the date of the most recent
        update attempt in the user's registry, under 'HKEY_CURRENT_USER\Software\QFPowershell\LastUpdate'

    .EXAMPLE
        Update-QFPowerShell
            Attempts to download the latest version of the QFPowerShell module using 'git pull'.
            Will check for the date of the last successful update, and if less than UpdateInterval days have passed
            (7 by default), will skip the update.

    .EXAMPLE
        Update-QFPowerShell -UpdateNow
            Attempts to download the latest version of the QFPowerShell module using 'git pull'.
            Does not check for the date of the last successful update, and ignores the 'DisableUpdateCheck' setting.

    .EXAMPLE
        Update-QFPowerShell -UpdateInterval 14
            Attempts to download the latest version of the QFPowerShell module using 'git pull'.
            Will check for the date of the last successful update, and if less than UpdateInterval days have passed
            will skip the update. The UpdateInterval setting of 14 days will be saved in the registry, and this value
            will be used for the UpdateInterval setting when this command is run again.

    .EXAMPLE
        Update-QFPowerShell -DisableUpdateCheck
            Disables automatically checking for updates.

    .PARAMETER DisableUpdateCheck
        Don't ever automatically check for updates at all.
        This setting will be stored in the registry under 'HKEY_CURRENT_USER\Software\QFPowershell\DisableUpdateCheck'
        If you want to re-enable updates you will need to delete this value from the registry,
        or manually run this cmdlet with the UpdateNow parameter.

    .PARAMETER UpdateInterval
        Number of days to wait before checking for updates again.
        By default this is set to 7, so update checks will occur weekly.
        This setting will be stored in the registry under 'HKEY_CURRENT_USER\Software\QFPowershell\UpdateInterval'

    .PARAMETER UpdateNow
        Forces an update check to run now, regardless of when the last update check was performed.
        This also ignores the 'DisableUpdateCheck' setting.
        Note that this parameter does not allow you to set any other parameters when it is used.

    .INPUTS
        This cmdlet does not accept any pipeline input.

    .OUTPUTS
        This cmdlet does not produce any pipeline output.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding(DefaultParameterSetName="Default")]
    param(
        [Parameter(Position=1,ParameterSetName="Default")]
        [int]$UpdateInterval,

        [Parameter(ParameterSetName="UpdateNow")]
        [switch]$UpdateNow,

        [Parameter(ParameterSetName="Default")]
        [switch]$DisableUpdateCheck
    )

    Process {

        # Skip this entire section if UpdateNow parameter is set
        If (!($UpdateNow.IsPresent)) {
            # Check if updates have been disabled
            If (($DisableUpdateCheck.IsPresent) -or ((Get-ItemProperty -Path HKCU:\Software\QFPowerShell -name DisableUpdateCheck -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisableUpdateCheck)) -ge 1) {
                Write-Verbose ("[$(Get-Date)] Updates have been disabled, skipping update check")
                # Try to save the DisableUpdateCheck setting into registry
                Try {
                    Set-ItemProperty -Path HKCU:\Software\QFPowerShell -name DisableUpdateCheck -Value 1 -ErrorAction Stop
                } Catch {
                    Write-Warning "Failed to save DisableUpdateCheck setting into registry."
                }
                # Nothing further to do so exit this function
                Return
            }

            # Check the update interval was set on command line, or saved in registry
            If ($PSBoundParameters.ContainsKey('UpdateInterval') -and (!(($null -eq $UpdateInterval) -or ($UpdateInterval -le 0)))) {
                # Try to save the new UpdateInterval into registry
                Try {
                    Set-ItemProperty -Path HKCU:\Software\QFPowerShell -name UpdateInterval -Value $UpdateInterval -ErrorAction Stop
                    Write-Verbose ("[$(Get-Date)] Saved new UpdateInterval setting of $UpdateInterval days into registry.")
                } Catch {
                    Write-Warning "Failed to save UpdateInterval setting of $UpdateInterval days into registry. Reverting to default setting of 7 days on next run."
                }
            } else {
                # Try to read the UpdateInterval from registry if it exists; otherwise just set it to the default of 7
                [int]$UpdateInterval = (Get-ItemProperty -Path HKCU:\Software\QFPowerShell -name UpdateInterval -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UpdateInterval)
                If (($null -eq $UpdateInterval) -or ($UpdateInterval -le 0)) {
                    $UpdateInterval = 7
                } else {
                    Write-Verbose ("[$(Get-Date)] Retrieved UpdateInterval setting of $UpdateInterval days from registry.")
                }
            }

            # Read the last update attempt date from registry
            $LastUpdate = (Get-ItemProperty -Path HKCU:\Software\QFPowerShell -name LastUpdate -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastUpdate)
            Try {
                $LastUpdate = Get-Date $LastUpdate -ErrorAction Stop
                If ((Get-Date) -lt ((Get-Date $LastUpdate).AddDays($UpdateInterval))) {
                    Write-Verbose ("[$(Get-Date)] Skipping update check. Last update was $LastUpdate")
                    Return
                } else {
                    Write-Verbose ("[$(Get-Date)] Last update was $LastUpdate - proceeding with update check...")
                }
            } Catch {
                Write-Verbose ("[$(Get-Date)] Couldn't get LastUpdate date from registry, will proceed with update now.")
            }
        }

        # Check the HOME environment variable is set, otherwise Git can be very slow.
        If (($null -eq $env:HOME) -or ($env:HOME -eq "")) {
            # Set HOME variable to the user's home directory.
            Write-Verbose ("[$(Get-Date)] HOME environment variable not set. Setting it to $env:USERPROFILE")
            # Setx saves variables into user registry so they persist when powershell is restarted.
            setx HOME "%USERPROFILE%" | Out-Null
            # this only affects the variable in the current session.
            $env:HOME = $env:USERPROFILE
        }

        # Check if Git is installed in the system PATH
        Try {
            git -v|Out-Null
            Write-Verbose ("[$(Get-Date)] Git executable found in the path.")
        } Catch {
            # See if we can get the path to git.exe from the registry, and add to the PATH environment variable for this session
            Try {
                $GitPath = (Get-ItemProperty -Path HKLM:\Software\GitForWindows -name InstallPath -ErrorAction Stop | Select-Object -ExpandProperty InstallPath) + "\bin;"
                $env:Path = $env:Path + $GitPath
            } catch {
                Write-Host "Git doesn't appear to be installed or is not in the PATH. Git is required for the auto-update function to work."
                Write-Host "Please check the Readme for this module for further information on setting up Git:"
                Write-Host "https://dev.azure.com/Derivco/Software/_git/MIGS-IT-QFPowershell?anchor=auto-update-using-git"
                Return
            }
            Write-Verbose ("[$(Get-Date)] Git executable path found in registry: $GitPath")
        }

        # Check if the module was cloned from Git, and run 'git pull' if so
        $ModuleRootPath = ($PSScriptRoot -replace "\\src$","")
        If (Test-Path $ModuleRootPath\.git -PathType Container) {
            Write-Host "Checking for updates to the QFPowerShell module, please wait..."
            Push-Location
            Set-Location $ModuleRootPath
            # Call Git to do a pull and update the module
            git pull
            Pop-Location
            Try {
                # Finally save the LastUpdatedate in registry
                Set-ItemProperty -Path HKCU:\Software\QFPowerShell -name LastUpdate -Value $(Get-Date) -ErrorAction Stop
            } Catch {
                Write-Warning "Failed to save the LastUpdate date into the registry."
            }
        } else {
            Write-Host "It appears this module was downloaded manually, and not cloned from a Git repository."
            Write-Host "You must clone the module's repository using Git for the auto-update function to work."
            Write-Host "Please check the Readme for this module for further information on setting up Git:"
            Write-Host "https://dev.azure.com/Derivco/Software/_git/MIGS-IT-QFPowershell?anchor=auto-update-using-git"
        }

    }
}