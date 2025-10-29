###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                              Helix GPA Functions                            #
#                                     v1.6.4                                  #
#                                                                             #
###############################################################################

# Author: Bernard Heije - bernhard.heije@derivco.es and Chris Byrne - christopher.byrne@derivco.com.au

$GPASuccessPrefix = "[GPA-SUCCESS] "
$GPAFailedPrefix = "[GPA-FAILED] "


function Invoke-QFHelixAutoTicket {
    <#
    .SYNOPSIS
        Automates processing of Game Play Analysis Incidents on the Helix ITSM system.

    .DESCRIPTION
        Automates processing of Game Play Analysis Incidents on the Helix ITSM system.
        This includes identifying relevant Incidents for processing, retrieving Incident data, generating Play Checks and updating/closing the ticket.

        This cmdlet searches for any Incidents that may be Game Play Analysis tickets. It then checks for a player Login, CasinoID, and TransactionID fields in the Incident.
        If this information is found, this cmdlet will invoke New-QFTicket to generate transaction and financial audits for the specified Player. 
        It will then generate Play Checks for the specified TransactionIDs, and Game Statistics Reports for any Games identified in the Play Checks.

        If this process is successful, the resulting ZIP archive from New-QFTicket will be attached to the Incident, and the incident marked as Resolved with a standardised response to the customer.

        If this process fails, the incident will be updated with an internal Working Log detailing the cause of the failure. This incident will need to be manually reviewed.


    .PARAMETER AIMode
        When AIMode parameter is specified, this cmdlet will process incidents via ChatGPT.
        
        The default behaviour is to search for incidents that are assign to 'MIGS - Customer Solutions' and lodged using the "Gameplay Assessment" SRD.
        This requires operators to lodge incidents using the Helix web page and to fill out the SRD correctly.

        In AIMode, all open incidents assigned to 'MIGS - Customer Solutions' will be processed via the 'Invoke-QFAI' function.
        If the AI determines that the incident is a Gameplay Assessment request, it will attempt to generate playchecks and transaction audits, and update the incident.
        Notably, this removes the requirement of incidents to be lodged using the SRD on the Helix website. 
        Incidents that are lodged via email will also be processed.

        This is an experimental feature and may not work consistently.

    .INPUTS
        This cmdlet does not accept any pipeline input.

    .OUTPUTS   
        This cmdlet does not generate any pipeline input.
                

    #>
    # Set up parameters that can be passed to the function
    [CmdletBinding()]
    param(
        [switch]$AIMode
    )

    function Get-GPA-Incidents {
        #Local function to retrieve tickets from Helix relevant for GPA
        #Returns $GPATicketList: List of GPA tickets
        
        if ($AIMode) {
            Write-Verbose "AI Mode enabled, checking all open tickets assigned to MIGS - Customer Solutions..."
            $SearchHelixIncidentResult = Search-QFHelixIncident -QueryField ("`'Assigned Group`'=`"MIGS - Customer Solutions`" AND `'Status`'=`"Assigned`" AND `'Submit Date`'>`"" + `
            $(Get-Date (get-date).AddDays(-2) -format 'yyyy-MM-dd') + '"')
        } else {
            $SearchHelixIncidentResult = Search-QFHelixIncident
        }
        if ($SearchHelixIncidentResult.Success -eq $false) {
            Write-Error ((Get-LogPrefix) + "Error on Search-QFHelixIncident, unable to proceed")
            return $false
        }

        $GPATicketList = $SearchHelixIncidentResult.Result

        # Check we got any results
        If ($GPATicketList.Count -lt 1) {
            Write-Host((Get-LogPrefix) + "No Game Play Analysis tickets found.")
            return $false
        }
        Write-Host ((Get-LogPrefix) + "Found " + $GPATicketList.Count + " GPA tickets to process:")
        foreach ($ticket in $GPATicketList) { 
            Write-Host ((Get-LogPrefix) + $ticket.sRID + " - " + $ticket.'Incident Number')
        }
        return $GPATicketList
    }

    function Validate-GPA-Incident {
        #Local function that gets the full incident from Helix and validates the Login, Casino, and transaction number values
        #Returns GPAValidationResult

        $GetHelixIncidentResult = Get-QFHelixIncident $GPATicket.'Incident Number'
        if ($GetHelixIncidentResult.Success -eq $false) {
            Write-Error (Get-LogPrefix $GPATicket) + "Error on Get-QFHelixIncident, proceding to next ticket"
            return $false
        }

        $GPATicketDetail = $GetHelixIncidentResult.Result

        [PSCustomObject]$GPAValidationResult = [PSCustomObject]@{
            Success     = $false
            ResultText = $null
            NewQFTicketArgs = $null
        }

        if ($AIMode) {
            # AIMode, Process tickets via Invoke-QFAI
            $AIResult = Invoke-QFAI -REQNumber $GPATicket.sRID -Query $GPATicketDetail.IssueDescription
            Write-Verbose ("AI Determined Category: " + $AIResult.category)
            
            # Check the Category determined by AI and that New-QFTicket function parameters were returned
            If ($AIResult.category -ne "winLegitimacy" -or $AIResult.functions.function_name -ne "New-QFTicket" -or $null -eq $AIResult.functions.args) {
                # Not a Gameplay Assessement ticket or the AI didn't return any arguments for New-QFTicket, skip this ticket
                $GPAValidationResult.ResultText = ("AI Determined Category: " + $AIResult.category)
                return $GPAValidationResult
            } else {
                # Process arguments returned from AI
                $NewQFTicketArgs = $AIResult.functions.args
                $NewQFTicketArgs.Add('NoCopyFilePath',$true)
                $NewQFTicketArgs.Add('PipelineOutput',$true)

                $GPAValidationResult.Success = $true
                $GPAValidationResult.NewQFTicketArgs = $NewQFTicketArgs
                $GPAValidationResult.ResultText = "OK"
                return $GPAValidationResult
            }
        }

        # Make sure we have the minimum required info - CasinoID, and login, and optionally transactionIDs.
        # These field names may change, depends on the final SRD
        # Remove duplicates and empty values
        [string[]]$GPALogin = $GPATicketDetail.DescriptionFields.'Player ID / MGS Login Name' | Select-Object -Unique | Where-Object { $_ -ne "" }
        Write-Verbose ((Get-LogPrefix $GPATicket) + "Login: $GPALogin")
        [string[]]$GPACasinoID = $GPATicketDetail.DescriptionFields.'Casino ID / Server ID / Product ID' | Select-Object -Unique | Where-Object { $_ -ne "" }
        Write-Verbose ((Get-LogPrefix $GPATicket) + "CasinoID: $GPACasinoID")
        [string[]]$GPATransactionIDs = $GPATicketDetail.DescriptionFields.'Game Round / Transaction IDs' | Select-Object -Unique | Where-Object { $_ -ne "" }
        Write-Verbose ((Get-LogPrefix $GPATicket) + "TransactionIDs: $GPATransactionIDs")
        
        # We will only handle one Login or CasinoID per ticket. Ops need to load a new ticket for each player.
        If ($GPALogin.Count -gt 1 -or $GPACasinoID.count -gt 1) {
            $GPAValidationResult.ResultText = "GPA Automation Failed: This ticket has multiple player Logins or CasinoIDs - skipping..." 
            return $GPAValidationResult
        }

        if ($null -eq $GPALogin -or $GPALogin.trim() -eq "") {
            $GPAValidationResult.ResultText = "GPA Automation Failed: No Player Login specified on this ticket - skipping..."
            return $GPAValidationResult
        }

        if ($null -eq $GPACasinoID -or $GPACasinoID.trim() -eq "") {
            $GPAValidationResult.ResultText = "GPA Automation Failed: No CasinoID specified on this ticket - skipping..."
            return $GPAValidationResult
        }
        else {
            # Check there's a numeric CasinoID
            try {
                [int]$GPACasinoID = $GPACasinoID[0]
            }
            catch {
                $GPAValidationResult.ResultText = "GPA Automation Failed: Invalid CasinoID specified on this ticket: $GPACasinoID"
                return $GPAValidationResult
            }
        }

        # Process TransactionIDs into an array, if we have multiple, split them on commas semicolons or spaces. Remove duplicates and empty values
        If ($null -ne $GPATransactionIDs -and $GPATransactionIDs.trim() -ne "") {
            try {
                [int[]]$GPATransactionIDs = $GPATransactionIDs -split "," -split ";" -split " " | Select-Object -Unique | Where-Object { $_ -ne "" }
                Write-Verbose ((Get-LogPrefix $GPATicket) + "Parsed these numeric TransactionIDs: $GPATransactionIDs")
            }
            catch {
                Write-Host ((Get-LogPrefix $GPATicket) + "Could not process TransactionIDs for " + $GPATicket.sRID + " " + $GPATicket.'Incident Number' + " - will not generate PlayChecks for these Transactions")
                $GPAValidationResult.ResultText = "GPA Automation Failed: Could not process TransactionIDs for " + $GPATicket.sRID  + " " + $GPATicket.'Incident Number'
                return $GPAValidationResult
            }
        }
        # We should now have enough information to process an audit and play checks for this ticket. Create a hashtable of arguments and splat it to New-QFTicket
        $NewQFTicketArgs = @{
            #INCNumber = $GPATicket.'Incident Number'
            REQNumber = $GPATicket.SRID
            Login     = $GPALogin[0]
            CasinoID  = $GPACasinoID[0]
            NoCopyFilePath = $true  
            PipelineOutput = $true
        }
        If ($null -ne $GPATransactionIDs -and $GPATransactionIDs -gt 0) {
            $NewQFTicketArgs.Add('TransactionIDs', $GPATransactionIDs)
        }

        $GPAValidationResult.Success = $true
        $GPAValidationResult.NewQFTicketArgs = $NewQFTicketArgs
        $GPAValidationResult.ResultText = "OK"
        return $GPAValidationResult
    }

    function Process-GPA-Incident {
        #Local function that processes GPA incident by executing New-QFTicket & validates New-QFTicket results
        #Returns $GPADataResult

        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [PSCustomObject]$NewQFTicketArgs
        )

        Write-Host ((Get-LogPrefix $GPATicket) + "Invoking New-QFTicket, parameters: ")
        foreach ($k in $NewQFTicketArgs.Keys) { Write-Host ((Get-LogPrefix $GPATicket) + "$k $($NewQFTicketArgs[$k])") }
        
        [PSCustomObject]$GPADataResult = [PSCustomObject]@{
            Success     = $null
            ResultText = $null
            GPAData = $null
        }
        Try {
            #TODO: Helix works with INC number, the REQ number is not visible anywhere in the ticket. With the REQ number the engineer is not able to ref to the ticket
            #And the customer is not able to see the INC numbers...
            Write-Host ("`n-------------------------------------------------------`n")
            $GPAData = New-QFTicket @NewQFTicketArgs
            $GPADataResult.GPAData = $GPAData
            Write-Host ("`n-------------------------------------------------------`n")
        }
        Catch {
            Write-Host ("`n-------------------------------------------------------`n")
            #Add-Workinfo-GPA-Result -DetailedDescription ("GPA Automation Failed: An error occured in New-QFTicket while generating audits and play checks - Error: " + $_.Exception.Message) | Out-Null
            $GPADataResult.Success = $false
            $GPADataResult.GPAData = $GPAData
            $GPADataResult.ResultText = "GPA Automation Failed: An error occured in New-QFTicket while generating audits and play checks - Error: " + $_.Exception.Message
            return $GPADataResult
        }
        
        # Check our GPA results
        If ($Null -ne $GPAData) {
            # Check the Game Statistics, make sure no ERROR results
            If ($GPAData.GameStatistics.Result -eq "ERROR") {
                #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Game statistics report shows RTP% outside expected ranges! This issue must be escalated to Game Intelligence or the ETI provider."  | Out-Null
                $GPADataResult.Success = $false
                $GPADataResult.ResultText = "GPA Automation Failed: Game statistics report shows RTP% outside expected ranges! This issue must be escalated to Game Intelligence or the ETI provider." 
                return $GPADataResult
            }
            # Check for ETI games, we will flag these for manual review until we know they work properly
            If ($GPAData.GameStatistics.ETI) {
                #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Ticket has playchecks for ETI games - Please review."  | Out-Null
                $GPADataResult.Success = $false
                $GPADataResult.ResultText = "GPA Automation Failed: Ticket has playchecks for ETI games - Please review." 
                return $GPADataResult
            }
            # Check the output ZIP file exists
            If ($null -eq $GPAData.ZipFile -or !(Test-Path -PathType Leaf $GPAData.ZipFile -ErrorAction SilentlyContinue)) {
                #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Failed to create a ZIP archive with audits and play check data - Please review."  | Out-Null
                $GPADataResult.Success = $false
                $GPADataResult.ResultText = "GPA Automation Failed: Failed to create a ZIP archive with audits and play check data - Please review." 
                return $GPADataResult
            }
            # Check we have a transaction audit file
            If (!($GPAData.Contents -eq 'Transaction_Audit.xlsx')) {
                #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Failed to create a Transaction Audit excel file - Please review."  | Out-Null
                $GPADataResult.Success = $false
                $GPADataResult.ResultText = "GPA Automation Failed: Failed to create a Transaction Audit excel file - Please review."
                return $GPADataResult
            }
            # Check we created play checks for all the requested transactions
            If ($null -ne $NewQFTicketArgs.TransactionIDs) {
                # record which TransactionIDs failed in an array
                $GPATransError = @()
                Foreach ($TransactionID in $NewQFTicketArgs.TransactionIDs) {
                    If (!($GPAData.Contents -eq "Playcheck $TransactionID.pdf")) {
                        [int[]]$GPATransError += $TransactionID
                    }
                }
                if ($GPATransError.Count -gt 0) {
                    #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Failed to create a Play Check files for transaction(s):`n $GPATransError"  | Out-Null
                    $GPADataResult.Success = $false
                    $GPADataResult.ResultText = "GPA Automation Failed: Failed to create a Play Check files for transaction(s):`n $GPATransError" 
                    return $GPADataResult
                }
            }
        }
        else {
            #Add-Workinfo-GPA-Result -DetailedDescription "GPA Automation Failed: Failed to generate playchecks for this request. No data returned from New-QFTicket."  | Out-Null
            $GPADataResult.Success = $false
            $GPADataResult.ResultText = "GPA Automation Failed: No data returned from New-QFTicket."
            return $GPADataResult
        }
        $GPADataResult.Success = $true
        $GPADataResult.ResultText = "GPA Automation Success"
        
        return $GPADataResult
    }
    function Add-Workinfo-GPA-Result {
        #Local functions that adds an internal workinfo with the GPA results
        #Returns true/false for success

        param(
            #GPAData not mandatory - we do not have GPAData if newQFTicket failed, or if there is a failure in validation
            [Parameter(Mandatory = $false, Position = 0)]
            [PSCustomObject]$GPAData,
        
            [Parameter(Mandatory = $true, Position = 1)]
            [ValidateNotNullOrEmpty()]
            [string]$DetailedDescription
        )

        # update the ticket with an internal work info
        # If there is player info for this ticket, add them
        If ($null -ne $GPAData.Player) {
            $DetailedDescription = $DetailedDescription + "`nGamingSystem: " + $GPAdata.Player.GamingServerID + "`nCasinoID: " + $GPAData.Player.CasinoID + "`nUserID: " + $GPAData.Player.UserID + "`nLogin: " + $GPAData.Player.Login 
        }
        # If there are game statistics results for this ticket, add them
        If ($null -ne $GPAData.GameStatistics) {
            $DetailedDescription = $DetailedDescription + "`nGame Statistics:"
            $GPAData.GameStatistics | ForEach-Object {
                $DetailedDescription = $DetailedDescription + "`nGameName: " + $_.GameName + " MID: " + $_.MID + " CID: " + $_.CID + " ETI: " + $_.ETI +
                "`nResult: " + $_.ResultText + "`n"
            }
        }

        # If there are ETI providers for this ticket, add them
        If ($null -ne $GPAData.ETIProviders) {
            $DetailedDescription = $DetailedDescription + "`nETI contact info:"
            $GPAData.ETIProviders | ForEach-Object {
                $ETIOutput = @()
                $ETIOutput += ("`n`nETI Game: " + $_.ETIGame)
                $ETIOutput += ("`nETI Provider Name: " + $_.ETIProvider)
                $ETIOutput += ("`nETI Provider ID: " + $_.ETIProviderId)
                If ($null -ne $_.Email -and $_.Email -ne "") { $ETIOutput += "`nSupport Email: " + $_.Email }
                If ($null -ne $_.PortalURI -and $_.PortalURI -ne "") { $ETIOutput += "`nSupport Portal: " + $_.PortalURI }
                If ($null -ne $_.PortalUsername -and $_.PortalUsername -ne "") { $ETIOutput += "`nUsername: " + $_.PortalUsername + "`nPassword: " + $_.PortalPassword }

                $DetailedDescription = $DetailedDescription + $ETIOutput
            }
        }

        If ($null -ne $GPAData.ZipFile) {
            $DetailedDescription = $DetailedDescription + "`n`n ZIP file path: " + $GPAData.ZipFile
        }

        $WorkInfoFields = @{
            "Work Log Type"        = "Working Log"
            "View Access"          = "Internal"
            #"Description"= + $Description
            "Detailed Description" = $DetailedDescription
        }
        # Now call the function to create the work info
        $NewHelixIncidentWorkInfoResult = New-QFHelixIncidentWorkInfo -IncidentNumber $GPATicket.'Incident Number' -WorkInfoFields $WorkInfoFields
        if ($NewHelixIncidentWorkInfoResult.Success -eq $true) {
            Write-Host ((Get-LogPrefix $GPATicket) + "Created internal Working log with GPA results")
            return $true
        }
        else {
            Write-Error ((Get-LogPrefix $GPATicket) + "Failed to create internal Working log with GPA results. Error: " + $NewHelixIncidentWorkInfoResult.'Result')
            return $false
        }        
    }
    function Add-Workinfo-GPA-Attachment {
        #Local functions that adds an public workinfo with the attachment
        #Returns true/false for success

        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [ValidateNotNullOrEmpty()]
            [String]$ZipFile
        )

        $WorkInfoFields = @{
            "Work Log Type"        = "Working Log"
            "View Access"          = "Public"
            "Detailed Description" = "Please see attached file for this round gameplay data.`n`nThis file has been compressed and password protected using the REQ number of this request as the password. You can use the REQ number of this request to access the file (eg: REQ1234567)"
        }
        # Now call the function to create the work info
        [String[]] $Files = $ZipFile
        $NewHelixIncidentWorkInfoResult = New-QFHelixIncidentWorkInfo -IncidentNumber $GPATicket.'Incident Number' -WorkInfoFields $WorkInfoFields -Files $Files
        if ($NewHelixIncidentWorkInfoResult.Success -eq $true) {
            Write-Host ((Get-LogPrefix $GPATicket) + "Created new Public working Log with attachment")
            return $true;
        }
        else {
            Write-Error ((Get-LogPrefix $GPATicket) + "Failed to create Public working log with attachment. Error: " + $NewHelixIncidentWorkInfoResult.'Result')
            return $false;
        }        
    }

    function Update-Incident-GPA-MarkProcessed {
        #Local functions that updates the incident title with a prefix, used to mark if an incident has already been processed
        #Returns true/false for success

        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [ValidateNotNullOrEmpty()]
            [string]$Prefix
        )
        # Also mark the ticket as processed
        $newDescription = $Prefix + $GPATicket.'Description'
        $UpdateFields = @{
            "Description" = $newDescription
        }
        $UpdateIncidentResult = Update-QFHelixIncident -RequestID $GPATicket.'Request ID' -UpdateFields $UpdateFields
        if ($UpdateIncidentResult.Success -eq $true) {
            Write-Host ((Get-LogPrefix $GPATicket) + "Updated ticket description with $Prefix")
            return $true;
        }
        else {
            Write-Error ((Get-LogPrefix $GPATicket) + "Failed to update the ticket description with $Prefix. Error: " + $UpdateIncidentResult.'Result')
            #TODO: In case there is a failure with marking the ticket. If tickets are left unmarked they will be processed again and again. This is a critical error.
            #How to handle?
            return $false;
        }
    }

    function Update-Incident-GPA-Resolved {
        #Local functions that updates the incident with status resolved and status_reason no further action required
        #Also updates the resolution text
        #Returns true/false for success

        # Now update the ticket with customer response and close it
        $Resolution = "Hi Support`n`nThank you for your request.`n`nWe have had a look at the account provided and the rounds in question.`n"
        # Work out the text of the response based on the game stats results
        if ($GPAData.GameStatistics.Result -eq "WARNING") {
            $Resolution = $Resolution + "It appears the player just got lucky. Removing their max payout will bring payout % to be within expected ranges. The game is paying out fairly and no suspicious activity was detected.`n`n" 
        }
        elseif ($GPAData.GameStatistics.Result -eq "OK") {
            $Resolution = $Resolution + "It appears the player's payout % is within volatility of the game, and no suspicious activity was detected.`n`n"
        }

        # We should always have a ZIP file containing a transaction audit, this is checked earlier in the script.
        If ($GPAData.Contents -like "PlayCheck*.pdf" -and $GPAData.Contents -like "Game Statistics*.pdf") {
            $Resolution = $Resolution + "I have attached a play check and a game monitor report for this game, plus a transaction audit of the player's recent activity.`n`n"
        }
        elseif ($GPAData.Contents -like "PlayCheck*.pdf") {
            $Resolution = $Resolution + "I have attached a play check report for this game, plus a transaction audit of the player's recent activity.`n`n"
        }
        elseif ($GPAData.Contents -like "Game Statistics*.pdf") {
            $Resolution = $Resolution + "I have attached a game monitor report for this game, plus a transaction audit of the player's recent activity.`n`n"
        }
        else {
            $Resolution = $Resolution + "I have attached a transaction audit of the player's recent activity.`n`n"
        }

        # ZIP file message
        #$Resolution = $Resolution + "Please note, this file has been compressed and password protected using the REQ number of this request as the password. You can use the REQ number of this request to access the file (eg: REQ1234567)`n`n"
        
        # Signature
        $Resolution = $Resolution + "We are now resolving this call.`nHowever, please do not hesitate to contact us should you require any further assistance with this.`n`nKind regards`nGames Global Support Team"                  


        $UpdateFields = @{
            
            "Status"                     = "Resolved"
            "Status_Reason"              = "No Further Action Required"
            "Resolution Category"        = "Quickfire - Audit"
            "Resolution Category Tier 2" = "Win verification"
            "Resolution Category Tier 3" = "N/A"
            "Time Log Duration (Min)"    = "1"
            "Resolution"                 = $Resolution
        }

        $UpdateIncidentResult = Update-QFHelixIncident -RequestID $GPATicket.'Request ID' -UpdateFields $UpdateFields
        if ($UpdateIncidentResult.Success -eq $true) {
            Write-Host ((Get-LogPrefix $GPATicket) + "Updated ticket with status Resolved and reason No Further Action Required")
            return $true            
        }
        else {
            Write-Error ((Get-LogPrefix $GPATicket) + "Failed to close the ticket with a customer response. Error: " + $UpdateIncidentResult.Result)
            return $false
        }
    }
    

    Write-Host ("`n")
       
    #0. First identify any relevant GPA tickets for action
    $GPATicketList = Get-GPA-Incidents
    if ($GPATicketList -eq $false)
    {
        return
    }

    # Process each returned ticket. 
    Foreach ($GPATicket in $GPATicketList) {
        Write-Host ((Get-LogPrefix $GPATicket) + "Processing Ticket " + $GPATicket.sRID + "-" + $GPATicket.'Incident Number')

        #1. Validate SRD input and save in $NewQFTicketArgs
        $GPAValidationResult = Validate-GPA-Incident
        
        if (($GPAValidationResult.Success) -eq $false) {
            #1a. GPA validation failed: Add working log and mark ticket processed
            Add-Workinfo-GPA-Result -DetailedDescription $GPAValidationResult.ResultText | Out-Null
            Update-Incident-GPA-MarkProcessed -Prefix $GPAFailedPrefix | Out-Null
            Continue
        }
        #2. Process GPAIncident - execute newQFTicket
        $GPADataResult = Process-GPA-Incident -NewQFTicketArgs $GPAValidationResult.NewQFTicketArgs
        

        if (($GPADataResult.Success) -eq $false) {
            #2a. GPA failed: Add working log and mark ticket processed
            Add-Workinfo-GPA-Result -GPAData $GPADataResult.GPAData -DetailedDescription $GPADataResult.ResultText | Out-Null
            Update-Incident-GPA-MarkProcessed -Prefix $GPAFailedPrefix | Out-Null
            Continue
        }

        #3. Add working log (internal) with GPA results
        if ((Add-Workinfo-GPA-Result -GPAData $GPADataResult.GPAData -DetailedDescription $GPADataResult.ResultText) -eq $false) {
            #3a. Add working log failed: mark processed
            Update-Incident-GPA-MarkProcessed -Prefix $GPAFailedPrefix | Out-Null
            Continue
        }

        #4. Add working log (public) with the attachment
        if ((Add-Workinfo-GPA-Attachment -ZipFile $GPADataResult.GPAData.ZipFile) -eq $false) {
            #4a. Add working log failed: mark processed
            Update-Incident-GPA-MarkProcessed -Prefix $GPAFailedPrefix | Out-Null
            Continue
        }

        #5. Update (resolve) incident with resolution text
        if ((Update-Incident-GPA-Resolved) -eq $false) {
            #5a. Update (resolve) incident failed: mark processed
            Update-Incident-GPA-MarkProcessed -Prefix $GPAFailedPrefix | Out-Null
            Continue
        }
        
        #6. Update incident - mark processed with [GPA-SUCCESS]
        Update-Incident-GPA-MarkProcessed -Prefix $GPASuccessPrefix | Out-Null
    } 
}
