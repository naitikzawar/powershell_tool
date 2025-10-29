
#Client Action Required RSS Feed
$CARRemedyRSSFeedURL = $null

#Pending Closure RSS Feed
$PCRemedyRSSFeedURL = $null

#Close ticket no feedback resolution method
$NoFeedbackResolutionMethod = $null

#Close ticket no feedback detailed root cause
$NoFeedbackDetailedRootCause = $null

#Close ticket no feedback service categories
$NoFeedbackServiceCategory = $null
$NoFeedbackServiceCategoryTier1 = $null
$NoFeedbackServiceCategoryTier2 = $null

$NoFeedbackProduct = $null
$NoFeedbackMarket = $null
$NoFeedbackSite = $null

#List of engineers for which the automation is activated. If $null, the automation activated for everybody
$ActiveEngineer = $null

#When automation is activated for the first time, we can not blindly sent chasers on all CAR tickets
#Ticket created date from which tickets are picked up by automation.  If $null, ticket created date will not be checked
$IncidentCreatedFrom = $null

$CustomerSolutionGroupName = $null
$ExternalGamingGroupName = $null


#RemedyEnvironment variable: Set this in Quickfire-Remedy-Base.psm1, it is set to DEV by default
switch ($Global:RemedyEnvironment) {
    #------------------   DEV configuration  ------------------
    "DEV" {
        #"Not available for DEV, dummy function will be used to retrieve test incident"
        $CARRemedyRSSFeedURL = $null

        #Not available for DEV, dummy function will be used to retrieve test incident"
        $PCRemedyRSSFeedURL = $null

        $NoFeedbackDetailedRootCause = "Operator - Insufficient Feedback Received"
        $NoFeedbackResolutionMethod = "Remedy"

        #On DEV the insufficient feedback category is not available, using other service category
        $NoFeedbackServiceCategory = "Quickfire"
        $NoFeedbackServiceCategoryTier1 = "Operator - Knowledge"
        $NoFeedbackServiceCategoryTier2 = "Non-issue"

        $NoFeedbackProduct = "Quickfire"
        $NoFeedbackMarket = "N/A"
        $NoFeedbackSite = "MIT Quickfire"

        #$ActiveEngineer = "bernhardh", "remedyusername2", "remedyusername3"
        $ActiveEngineer = $null

        $IncidentCreatedFrom = $null

        $CustomerSolutionGroupName = 'MIGS - Customer Solutions'
        $ExternalGamingGroupName = 'MIGS - IT - External Gaming'
    }
    #------------------   PROD configuration  ------------------
    "PROD"{
        $CARRemedyRSSFeedURL = "http://quickfirerss.mgsops.net/rss/incidents/clientactionrequired"
        $PCRemedyRSSFeedURL = "http://quickfirerss.mgsops.net/rss/incidents/pendingclosure"

        $NoFeedbackDetailedRootCause = "Operator - Insufficient Feedback Received"
        $NoFeedbackResolutionMethod = "Remedy"

        $NoFeedbackServiceCategory = "Quickfire"
        $NoFeedbackServiceCategoryTier1 = "Insufficient Feedback"
        $NoFeedbackServiceCategoryTier2 = "Insufficient Feedback"

        $NoFeedbackProduct = "Quickfire"
        $NoFeedbackMarket = "N/A"
        $NoFeedbackSite = "Derivco Malaga"

        $ActiveEngineer = $null
        #$ActiveEngineer = "bernhardh", "jennifern", "jeffm", "harleyo", "simonk", "wouters"

        $IncidentCreatedFrom = $null
        #$IncidentCreatedFrom = (Get-Date -Year 2024 -Month 2 -Day 19 -Hour 0 -Minute 0 -Second 0)
        
        $CustomerSolutionGroupName = 'MIGS - Customer Solutions'
        $ExternalGamingGroupName = 'MIGS - IT - External Gaming'
    }
}

#------------------   GENERAL configuration  ------------------
$CustomerSolutionsFooter = "

Kind regards,

Games Global Support"

$ExternalGamingFooter = "

Kind regards,

MIGS IT External Gaming"

$CARChase1Text = "Good day,

We are hopeful that you've had a chance to review the latest comment from our support team. We are expecting further information from you to proceed with this ticket.

Could you provide feedback on our latest message?

If you'd like to provide an update, or require more time to work through our latest message, simply reply to this email and let us know." 



$CARChase2Text = "Good day,

We are hopeful that you've had a chance to review the latest comment from our support team. We are expecting further information from you to proceed with this ticket.

If you'd like to provide an update, or require more time to work through our latest message, simply reply to this email and let us know. 

If we don't hear back from you within the next 24 hours, we’ll assume this issue is solved."


$insufficientFeedbackResolutionText = "Good day,

We are closing this call as we are unable continue the investigation without further information from you.

If you'd like to provide an update, or require more time to work through our latest message, simply reply to this email and let us know.

We look forward to hearing back from you."

$CloseErrorMessage = "MIGS Customer Solution automation failed to close this incident.
    
If there is a task pending on this incident. You can check if the task is indeed cancelled/finished, close the task, and go ahead and manually close this ticket.

In any other case, please review why the automation was not able to close this incident. On calling the ResolveIncident method, the Remedy Integration API returned the following error:
"   

$CloseNotificationText = "Hello Engineer, 
This ticket has been in Pending Closure state for more then 24 hours. It has been moved to the assigned queue. You can close the ticket."


function New-WorkinfoCARChase1 {
        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [PSCustomObject]$Incident
        )
    
    $Footer = switch ($Incident.AssignedGroup) {
        $CustomerSolutionGroupName {$CustomerSolutionsFooter}
        $ExternalGamingGroupName {$ExternalGamingFooter}
    }

    $Notes = $CARChase1Text + $Footer

    # Now call the function to create the work info
    $NewRemedyIncidentWorkInfoResult = New-QFRemedyIncidentWorkInfo -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername -WorkInfoType 'Status Update' -ViewAccess 'Public' -Summary 'CC#1 - Customer feedback required' -Notes $Notes

    if ($NewRemedyIncidentWorkInfoResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Created public working log chase #1")
        return $true
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to create public work log chase #1. Error: " + $NewRemedyIncidentWorkInfoResult.'Result')
        return $false
    }        
}

function New-WorkinfoCARChase2 {
    param(
            [Parameter(Mandatory = $true, Position = 0)]
            [PSCustomObject]$Incident
        )

    $Footer = switch ($Incident.AssignedGroup) {
        $CustomerSolutionGroupName {$CustomerSolutionsFooter}
        $ExternalGamingGroupName {$ExternalGamingFooter}
    }
    $Notes = $CARChase2Text + $Footer

    # Now call the function to create the work info
    $NewRemedyIncidentWorkInfoResult = New-QFRemedyIncidentWorkInfo -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername -WorkInfoType 'Status Update' -ViewAccess 'Public' -Summary 'PC#1 - Customer feedback required' -Notes $Notes

    if ($NewRemedyIncidentWorkInfoResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Created public working log chase #2")
        return $true
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to create public work log chase #2.  Error: " + $NewRemedyIncidentWorkInfoResult.'Result')
        return $false
    }        
}

function Update-IncidentPendingClosure {

    param(
            [Parameter(Mandatory = $true, Position = 0)]
            [PSCustomObject]$Incident
        )

    $UpdateIncidentResult = Update-QFRemedyIncident -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername -Status 'Pending' -StatusReason 'Pending Closure'
    if ($UpdateIncidentResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Updated ticket status to Pending Closure")
        return $true;
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to update ticket status to Pending Closure. Error: " + $UpdateIncidentResult.'Result')
        return $false;
    }
}


#--- Functions used for tickets in status Pending Closure

function Update-IncidentResolve{
    <#
  .SYNOPSIS
      Resolve Remedy incident with insufficient feedback fields

  .DESCRIPTION
       Resolve Remedy incident with insufficient feedback fields


  .PARAMETER Fields
      $[object] ToCloseticket: Ticket to cloes
      $[string] resolutionText: Resolution text

  .OUTPUTS
      Boolean succes or failed
   #>
   param(
    [Parameter(Mandatory = $true, Position = 0)]
    [PSCustomObject]$Incident
  )

    $Footer = switch ($Incident.AssignedGroup) {
        $CustomerSolutionGroupName {$CustomerSolutionsFooter}
        $ExternalGamingGroupName {$ExternalGamingFooter}
    } 
    $Footer = $Footer + "
AC#1"

    $ResolutionText = $InsufficientFeedbackResolutionText + $Footer

  $UpdateIncidentResult = Resolve-QFRemedyIncident -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername -Status 'Resolved' -StatusReason 'Customer Follow-Up Required' `
  -ResolutionMethod $NoFeedbackResolutionMethod -DetailedRootCause $NoFeedbackDetailedRootCause `
  -ServiceCategory $NoFeedbackServiceCategory -ServiceCategoryTier1 $NoFeedbackServiceCategoryTier1 -ServiceCategoryTier2 $NoFeedbackServiceCategoryTier2 `
  -Product $NoFeedbackProduct -Market $NoFeedbackMarket -Site $NoFeedbackSite -ResolutionText $ResolutionText
  
  #$UpdateIncidentResult =  Resolve-QFRemedyIncident -IncidentNumber $Incident.Id -UpdateFields $UpdateFields
  if ($UpdateIncidentResult.Success -eq $true) {
      Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Updated ticket to Resolved-Customer follow up required")
      return $UpdateIncidentResult            
  }
  else {
      Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to update ticket to Resolved-Customer follow up required. Error: " + $UpdateIncidentResult.Result)
      return $UpdateIncidentResult   
  }
}

function Update-IncidentAssigned {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [PSCustomObject]$Incident
      )

    $UpdateIncidentResult = Update-QFRemedyIncident -IncidentNumber $Incident.Id -Status 'Assigned' -RemedyUsername $Incident.AssigneeUsername
    if ($UpdateIncidentResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Updated ticket status to Assigned")
        return $true;
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to update ticket status to Assigned. Error: " + $UpdateIncidentResult.'Result')
        return $false;
    }
}

function New-WorkinfoCloseNotificiation{
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [PSCustomObject]$Incident  
      )


    # Now call the function to create the work info
    $NewRemedyIncidentWorkInfoResult = New-QFRemedyIncidentWorkInfo -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername `
    -WorkInfoType 'Working Log' -ViewAccess 'Internal' -Summary 'No customer response - close ticket' -Notes $CloseNotificationText

    if ($NewRemedyIncidentWorkInfoResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Created internal working log with close notification")
        return $true
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to create internal working log with close notification.  Error: " + $NewRemedyIncidentWorkInfoResult.'Result')
        return $false
    }        
}


function New-WorkinfoCloseError{
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [PSCustomObject]$Incident,

        [Parameter(Mandatory = $true, Position = 1)]
        [PSCustomObject]$APIErrorMessage
         )

    $CloseErrorMessage = $CloseErrorMessage + $APIErrorMessage

    $NewRemedyIncidentWorkInfoResult = New-QFRemedyIncidentWorkInfo -IncidentNumber $Incident.Id -RemedyUsername $Incident.AssigneeUsername `
    -WorkInfoType 'Working Log' -ViewAccess 'Internal' -Summary 'Error occurred on closing incident - review' -Notes $CloseErrorMessage

    if ($NewRemedyIncidentWorkInfoResult.Success -eq $true) {
        Write-Host ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Created internal working log with close error")
        return $true
    }
    else {
        Write-Error ((Get-RemedyLogPrefix $Incident.RequestId $Incident.Id) + "Failed to create internal working log with close error.  Error: " + $NewRemedyIncidentWorkInfoResult.'Result')
        return $false
    }        
}






function Invoke-ProcessCARTickets {

    #First identify any relevant CAR tickets for action
    $CARTicketList = $null

    switch ($Global:RemedyEnvironment) {
        "DEV" {
            $CARTicketList = Search-QFRemedyCARIncidentsDummy
        }
        "PROD" {
            $CARTicketList = Search-QFRemedyCARIncidents
        }
    }

    if ($CARTickets.Success -eq $false)
    {
        return
    }

    $ticketCount = $CARTicketList.Result.Count
    Write-Host (((Get-LogPrefix)) + "Found $ticketCount Client Action Required tickets to process")
    $BasicCheckVerificationCount = 0
    $ActiveEngineerVerificationCount = 0
    $TicketCreationDateVerificationCount = 0
    $FirstChaserCount = 0
    $SecondChaserCount = 0
    $NoActionCount = 0
    $ErrorCount = 0


    # Process each returned ticket. 
    Foreach ($CARTicket in $CARTicketList.Result) {

      
        #First retrieve the full incident information, do checks
        $Incident = $null
        $GetIncidentResult = Get-QFRemedyIncident $CARTicket.IncidentNumber
        if ($GetIncidentResult.Success -eq $false) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Unable to retrieve incident. Error: " + $GetIncidentResult.'Result')
            $ErrorCount++
            continue
        }

        $Incident = $GetIncidentResult.Result

        #This should already be covered by the RSS feed query, we do it again for verification
        if (($Incident.Status -notin ('In Progress','Pending')) -or
            ($Incident.StatusReason -ne 'Client Action Required') -or
            ($Incident.AssignedGroup -notin ($CustomerSolutionGroupName,$ExternalGamingGroupName)) -or
            ($null -eq $Incident.AssigneeUsername)){
            
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Incident did not pass basic checks for proceding; investigate")
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Status= " + $incident.Status)
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "StatusReason= " + $Incident.StatusReason)
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "AssignedGroup= " + $Incident.AssignedGroup)
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "AssigneeUsername= " + $Incident.AssigneeUsername)

            $BasicCheckVerificationCount++
            continue
        }

        
        #Check if incident Engineer is within active engineers list
        if ($null -ne $ActiveEngineer -and $Incident.AssigneeUsername -notin $ActiveEngineer) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "AssigneeUsername= " + $Incident.AssigneeUsername + " is not present in ActiveEngineer list, skipping")
            $ActiveEngineerVerificationCount++
            continue
        }

        #Check if incident is created before automation activation date
        if ($null -ne $IncidentCreatedFrom -and $incident.Created -lt $IncidentCreatedFrom) {
           Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Incident is created before automation activation date, skipping")
           $TicketCreationDateVerificationCount++
           continue
        }

        $ticketWorkinfosResult = get-QFRemedyIncidentWorkInfo $CARTicket.IncidentNumber
        if ($ticketWorkinfosResult.Success -eq $false) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Unable to get workinfo's for incident. Error: " + $ticketWorkinfosResult.'Result')
            $ErrorCount++
            continue
        }

        #Ticket is in status Client Action Required

        #First find the last Public workinfo
        $lastPublicWorkinfo = $ticketWorkinfosResult.Result | 
        Sort-Object 'Created' -descending | 
        Where-Object {
            ($_.'View_Access' -eq 'Public')
        } | 
        Select-Object -First 1

        #Check if there are any internal workinfo's containing the nochase tage, that are created only after last Public workinfo
        $InternalWorkinfoNoChase = $null
        $InternalWorkinfoNoChase = $ticketWorkinfosResult.Result | 
        Sort-Object 'Created' -descending | 
        Where-Object {
            ($_.'WorkLogType' -eq 'Working Log') -and 
            ($_.'View_Access' -eq 'Internal') -and 
            ($_.'Created' -gt $lastPublicWorkinfo.Created) -and
            ($_.'Summary' -like '*NC#1*')
        } | 
        Select-Object -First 1

        #If we found internal workinfo containing nochase tag, cancel automation
        if ($null -ne $InternalWorkinfoNoChase) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Found internal workinfo with No chase tage, skipping")
            $NoActionCount++
            continue
        }


        #First scenario: No chaser has been sent
        #Work Log Type = 'Status update' or 'General Information'
        #Title not contain:  NC#1, CC#1
        #Work Log Submit Date = Older then 2 days
        $workInfoMatch = $null
        $workInfoMatch = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'WorkLogType' -eq 'Status Update') -or ($_.'WorkLogType' -eq 'General Information')) -and
            ($_.'Summary' -notlike '*NC#1*') -and
            ($_.'Summary' -notlike '*CC#1*') -and
            ((AddWorkingDays $_.'Created' 2) -lt (Get-Date))
        }
        
        #If we found a public workinfo wihin conditions, we can send the first chaser
        if ($null -ne $workInfoMatch) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Found ticket suitable for first chase")
            New-WorkinfoCARChase1 $Incident | Out-Null
            $FirstChaserCount++
            continue
        }

       
        #Second scenario: The first chaser has already been sent
        #Work Log Type = 'Status update' or 'General Information'
        #Title contain:  CC#1
        #Work Log Submit Date = Older then 2 days
        $workInfoMatch = $null
        $workInfoMatch = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'WorkLogType' -eq 'Status Update') -or ($_.'WorkLogType' -eq 'General Information')) -and
            ($_.'Summary' -like '*CC#1*') -and
            ((AddWorkingDays $_.'Created' 2) -lt (Get-Date))
        }

        #If we found a public workinfo wihin conditions, we can do the second chase and set the status to Pending Closure
        if ($null -ne $workInfoMatch) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Found ticket suitable for second chase")
            Update-IncidentPendingClosure $Incident | Out-Null
            New-WorkinfoCARChase2 $Incident | Out-Null
            $SecondChaserCount++
            continue
        }

        Write-Host ((Get-RemedyLogPrefix $CARTicket.RequestId $CARTicket.IncidentNumber) + "No action on ticket")
        $NoActionCount++

    }
    Write-Host ((Get-LogPrefix) + "Results: 
    Total ticket count: $ticketCount
    Ticket did not pass basic check: $BasicCheckVerificationCount
    Ticket assignee not in active engineer list: $ActiveEngineerVerificationCount
    Ticket create date before automation start date: $TicketCreationDateVerificationCount
    Error occurred: $ErrorCount
    First chaser sent: $FirstChaserCount
    Second chaser sent: $SecondChaserCount
    No action: $NoActionCount")
}

function Invoke-ProcessPCTickets {

    #First identify any relevant PC tickets for action
    $PCTicketList = $null
    switch ($Global:RemedyEnvironment) {
        "DEV" {
            $PCTicketList = Search-QFRemedyPCIncidentsDummy
        }
        "PROD" {
            $PCTicketList = Search-QFRemedyPCIncidents
        }
    }
    
    

    if ($PCTickets.Success -eq $false)
    {
        return
    }

    $ticketCount = $PCTicketList.Result.Count
    Write-Host (((Get-LogPrefix)) + "Found $ticketCount Pending Closure tickets to process")
    $BasicCheckVerificationCount = 0
    $ActiveEngineerVerificationCount = 0
    $TicketCreationDateVerificationCount = 0
    $ErrorCount = 0
    $CloseTicketFeedbackCount = 0
    $PushAssignedCount = 0
    $NoActionCount = 0


    # Process each returned ticket. 
    Foreach ($PCTicket in $PCTicketList.Result) {
        
        #First retrieve the full incident information, do checks
        $Incident = $null
        $GetIncidentResult = Get-QFRemedyIncident $PCTicket.IncidentNumber
        if ($GetIncidentResult.Success -eq $false) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Unable to retrieve incident. Error: " + $GetIncidentResult.'Result')
            $ErrorCount++
            continue
        }

        $Incident = $GetIncidentResult.Result

        #This should already be covered by the RSS feed query, we do it again for verification
        if (($Incident.Status -notin ('In Progress','Pending')) -or
            ($Incident.StatusReason -ne 'Pending Closure') -or
            ($Incident.AssignedGroup -notin ($CustomerSolutionGroupName,$ExternalGamingGroupName)) -or 
            ($null -eq $Incident.AssigneeUsername)){
            
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Incident did not pass basic checks for proceding; investigate")
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Status= " + $incident.Status)
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "StatusReason= " + $Incident.StatusReason)
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "AssignedGroup= " + $Incident.AssignedGroup)
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "AssigneeUsername= " + $Incident.AssigneeUsername)
            $BasicCheckVerificationCount++
            continue
        }

        #If $ActiveEngineer variable contains engineers, check if incident Assignee is in the list
        if ($null -ne $ActiveEngineer -and $Incident.AssigneeUsername -notin $ActiveEngineer) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "AssigneeUsername= " + $Incident.AssigneeUsername + " is not present in ActiveEngineer list, skipping")
            $ActiveEngineerVerificationCount++
            continue
        }

        #Check if incident is created before automation activation date
        if ($null -ne $IncidentCreatedFrom -and $incident.Created -lt $IncidentCreatedFrom) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Incident is created before automation activation date, skipping")
            $TicketCreationDateVerificationCount++
            continue
        }


        $ticketWorkinfosResult = get-QFRemedyIncidentWorkInfo $PCTicket.IncidentNumber
        if ($ticketWorkinfosResult.Success -eq $false) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Unable to get workinfo's for incident. Error: " + $ticketWorkinfosResult.'Result')
            $ErrorCount++
            continue
        }


         #First find the last Public workinfo
         $lastPublicWorkinfo = $ticketWorkinfosResult.Result | 
         Sort-Object 'Created' -descending | 
         Where-Object {
             ($_.'View_Access' -eq 'Public')
         } | 
         Select-Object -First 1

        #Check if there are any internal workinfo's containing the nochase tage, that are created only after last Public workinfo
        $InternalWorkinfoNoChase = $null
        $InternalWorkinfoNoChase = $ticketWorkinfosResult.Result | 
        Sort-Object 'Created' -descending | 
        Where-Object {
            ($_.'WorkLogType' -eq 'Working Log') -and 
            ($_.'View_Access' -eq 'Internal') -and 
            ($_.'Created' -gt $lastPublicWorkinfo.Created) -and
            ($_.'Summary' -like '*NC#1*')
        } | 
        Select-Object -First 1

        #If we found internal workinfo containing nochase tag, cancel automation
        if ($null -ne $InternalWorkinfoNoChase) {
            Write-Host ((Get-RemedyLogPrefix  $CARTicket.RequestId $CARTicket.IncidentNumber) + "Found internal workinfo with No chase tage, skipping")
            $NoActionCount++
            continue
        }


        #First scenario: The ticket has been set to Pending Closure with the PC#1 tag. Either by the automation, or manual by engineer
        #Work Log Type = 'Status update' or 'General Information'
        #Title contain:  PC#1
        #Work Log Submit Date = Older then 1 day
        $workInfoMatch = $null
        $workInfoMatch = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'WorkLogType' -eq 'Status Update') -or ($_.'WorkLogType' -eq 'General Information')) -and
            ($_.'Summary' -like '*PC#1*') -and
            ((AddWorkingDays $_.'Created' 1) -lt (Get-Date))
        }
        
        #If we found a public workinfo wihin conditions, we can close ticket with status customer feedback required
        if ($null -ne $workInfoMatch) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Found ticket suitable for closing")
            $UpdateIncidentResult = Update-IncidentResolve $Incident
            if ($UpdateIncidentResult.Success -eq $true) {
                $CloseTicketFeedbackCount++
            } else {
                Update-IncidentAssigned $Incident | Out-Null
                New-WorkinfoCloseError $Incident $UpdateIncidentResult.Result | Out-Null
                $ErrorCount++
            }
            continue
        }


        #Second scenario: The ticket has been set to Pending Closure without the PC#1 tag. This would be done manual by engineer.
        #Work Log Type = 'Status update' or 'General Information'
        #Title not contain:  PC#1
        #Work Log Submit Date = Older then 1 days
        $workInfoMatch = $null
        $workInfoMatch = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'WorkLogType' -eq 'Status Update') -or ($_.'WorkLogType' -eq 'General Information')) -and
            ($_.'Summary' -notlike '*PC#1*') -and
            ((AddWorkingDays $_.'Created' 1) -lt (Get-Date))
        }
        
        #If we found a public workinfo wihin conditions, we can move the ticket to the assigned queue, hereby notifying the engineer
        if ($null -ne $workInfoMatch) {
            Write-Host ((Get-RemedyLogPrefix  $PCTicket.RequestId $PCTicket.IncidentNumber) + "Found ticket suitable for closing, push to assigned queue")
            Update-IncidentAssigned $Incident | Out-Null
            New-WorkinfoCloseNotificiation $Incident | Out-Null
            $PushAssignedCount++
            continue
        }

        Write-Host ((Get-RemedyLogPrefix $PCTicket.RequestId $PCTicket.IncidentNumber) + "No action on ticket")
        $NoActionCount++
    }

    Write-Host ((Get-LogPrefix) + "Results: 
    Total ticket count: $ticketCount
    Ticket did not pass basic check: $BasicCheckVerificationCount
    Ticket assignee not in active engineer list: $ActiveEngineerVerificationCount
    Ticket create date before automation start date: $TicketCreationDateVerificationCount
    Error occurred: $ErrorCount
    Ticket closed Customer Feedback Required: $CloseTicketFeedbackCount
    Ticket pushed to Assigned: $PushAssignedCount
    No action: $NoActionCount")
}


function Search-QFRemedyCARIncidentsDummy {

    $incidents = @()
    $incident = @{
        'IncidentNumber' = 'INC1240188'
        'RequestId' = 'REQ1212083'
    }
    $incidents += $incident
        
    $CustomResponse = [PSCustomObject]@{
        Success = $true
        Result  = $incidents
    }
    $CustomResponse
}


function Search-QFRemedyPCIncidentsDummy {

    $incidents = @()
    
    $incident = @{
        'IncidentNumber' = 'INC1240247'
        'RequestId' = 'REQ1212184'
    }

    $incidents += $incident
        
    $CustomResponse = [PSCustomObject]@{
        Success = $true
        Result  = $incidents
    }
    $CustomResponse
}


function Search-QFRemedyCARIncidents {
    <#
    .SYNOPSIS
        Requests the RSS feed for Incidents matching specified criteria and returns basic information for any matching Incidents.

    .DESCRIPTION
        Retrieves an Incident and associated data from the Client Action Required RSS feed


    .EXAMPLE 
        Search-QFRemedyCARIncidents
            Requests the data from the RSS feed any Incidents matching the default QueryField parameter.
            

    .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request
        
    #>
    

    #Ticket list to be retrieved from Remedy DB:
    #Assigned Group: MIGS Customer Solutions
    #Assignee: Not null
    #Status: Pending
    #Status Reason: Client Action Required
    #Fields we want returned: INC number, REQ number

    try {
        $Response = Invoke-RestMethod $CARRemedyRSSFeedURL -Method 'GET' -SkipCertificateCheck
        #TODO: test behaviour in case not tickets are returned


        $incidents = $Response | ForEach-Object {
            @{
               
                'IncidentNumber' = $_.title.split(' || ')[0]
                'RequestId' = $_.title.split(' || ')[1]
            }
        }     
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $incidents
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on Search-QFRemedyCARIncidents feed")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    $CustomResponse
}




function Search-QFRemedyPCIncidents {
    <#
    .SYNOPSIS
        Requests the RSS feed for Incidents matching specified criteria and returns basic information for any matching Incidents.

    .DESCRIPTION
        Retrieves an Incident and associated data from the Pending Closure RSS feed


    .EXAMPLE 
        Search-QFRemedyPCIncidents
            Requests the data from the RSS feed any Incidents matching the default QueryField parameter.
            

    .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request
        
    #>
    

    #Ticket list to be retrieved from Remedy DB:
    #Assigned Group: MIGS Customer Solutions
    #Assignee: Not null
    #Status: Pending
    #Status Reason: Pending Closure
    #Fields we want returned: INC number, REQ number

    try {
        $Response = Invoke-RestMethod $PCRemedyRSSFeedURL -Method 'GET' -SkipCertificateCheck
        #TODO: test behaviour in case not tickets are returned

        $incidents = $Response | ForEach-Object {
            @{
               
                'IncidentNumber' = $_.title.split(' || ')[0]
                'RequestId' = $_.title.split(' || ')[1]
            }
        }     
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $incidents
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on Search-QFRemedyPCIncidents feed")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    $CustomResponse
}
