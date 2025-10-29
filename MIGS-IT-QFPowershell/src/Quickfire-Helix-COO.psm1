###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                      Helix COO (coordinator) Functions                      #
#                                  v1.6.4                                     #
#                                                                             #
###############################################################################
$24HourWarningText = "*If we don't hear back from you within the next 24 hours, we’ll assume this issue is solved.*" 
$24HourResolutionText = "Good day,

We are closing this call as we are unable continue the investigation without further information from you.

If you'd like to provide an update, or require more time to work through our latest message, simply reply to this email and let us know.

We look forward to hearing back from you.

Kind Regards,

Games Global Support"

$closeQuestionText = "*Please let us know whether you are happy for us to close the ticket?*"
$closeQuestionResolutionText = "Good day!

We are hopeful that you've had a chance to review the latest comment from our support team. We believe that there is nothing more to investigate with this ticket and hope the resolution was sufficient.

As it's been some time since we have heard back from you, we’ll assume this issue is solved.

If you'd like to provide an update, or require more time to work through our latest message, simply reply to this email within the next 4 days and let us know.

Kind Regards,

Games Global Support"


function Invoke-QFClosetickets {
    <#
    .SYNOPSIS
        Automates closing of tickets on which the operator is not responding.

    .DESCRIPTION
        Automates closing of tickets on which the operator is not responding.

        It reads through MIGS-Customer Solutions tickets
        - Assigned group = MIGS-Customer Solutions 
        - Status = pending 
        - Status reason = Client Action Required
        - Last modified = within 6 days

        It checks these ticket's latest Public workinfo. If this workinfo contains a certain text, and is older then a defined amount of working days (so this will exclude weekend days!), 
        The ticket will be closed 

        .INPUTS
            This function does not accept any pipeline input.

        .OUTPUTS   
            This function does not generate any pipeline input.
                

    #>

    #Get tickets all tickets with status: pending, status reason: client action required, within 4 days last modifid
    [String[]]$Fields = @("Incident Number", "Request ID","Assigned Group", "Status", "Status_Reason", "Submit Date", "Last Modified Date","Categorization Tier 3","SRID","Description", "Resolution Category", "Resolution Category Tier 2", "Resolution Category Tier 3")
    [string]$QueryField = "`'Assigned Group`'=`"MIGS - Customer Solutions`" AND `'Status`'=`"Pending`" AND `'Status_Reason`'=`"Client Action Required`" AND `'Last Modified Date`'>`"" + $(Get-Date (get-date).AddDays(-18) -format 'yyyy-MM-dd') + '"'

    $SearchHelixIncidentResult = Search-QFHelixIncident -Fields $Fields -QueryField $QueryField
    if ($SearchHelixIncidentResult.Success -eq $false) {
        Write-Error ((Get-LogPrefix) + "Error on Search-QFHelixIncident, unable to proceed")
        return
    }

    $ToCloseTicketList = $SearchHelixIncidentResult.Result

    # Check we got any results
    If ($ToCloseTicketList.Count -lt 1) {
        Write-Host((Get-LogPrefix) + "No tickets found for closing in initial lookup")
        return
    }

   
    #Loop through tickets & resolve if within closing conditions
    foreach ($ToCloseticket in $ToCloseTicketList)
    {       
        #Retrieve ticket Workinfo's
        $ticketWorkinfosResult = Get-QFHelixIncidentWorkInfo -IncidentNumber $ToCloseticket.'Incident Number'
        if ($ticketWorkinfos.Success -eq $false)
        {
            Write-Host ((Get-LogPrefix $ToCloseticket) + "Unable to get workinfo's for incident " + $ToCloseTicket.'Incident Number' + ". Error: " + $ticketWorkinfosResult.'Result')
            continue
        }
               
       
        #Find the newest public work info, we will do validations only on this workinfo: 
        $lastPublicWorkinfo = $ticketWorkinfosResult.Result | 
        Sort-Object 'Work Log Submit Date' -descending | 
        Where-Object {
            ($_.'View Access' -eq 'Public')
        } | 
        Select-Object -First 1 

        #Work Log Type = 'Status update' or 'General Information'
        #Work Log Submit Date = Older then 1 day (24 hours)
        #Detailed Description contains $24HourWarningText
        $Close24HourWorkInfo = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'Work Log Type' -eq 'Customer Status Update') -or ($_.'Work Log Type' -eq 'General Information')) -and
            ((AddWorkingDays $_.'Work Log Submit Date' 1) -lt (Get-Date)) -and 
            ($_.'Detailed Description' -like $24HourWarningText)
        } 
        
        #If we found a public workinfo wihin conditions -> Resolve ticket
        if ($null -ne $Close24HourWorkInfo)
        {
            Write-Host ((Get-LogPrefix $ToCloseticket) + "Closing ticket after 24-hour warning")
            ResolveTicket $ToCloseticket $24HourResolutionText
            continue
        }

        #Work Log Type = 'Status update' or 'General Information'
        #Work Log Submit Date = Older then 3 days
        #Detailed Description contains $closeQuestionText
        $CloseQuestionWorkInfo = $lastPublicWorkinfo | 
        Where-Object {
            (($_.'Work Log Type' -eq 'Customer Status Update') -or ($_.'Work Log Type' -eq 'General Information')) -and
            ((AddWorkingDays $_.'Work Log Submit Date' 3) -lt (Get-Date)) -and 
            ($_.'Detailed Description' -like $closeQuestionText)
        } 

        #If we found a public workinfo wihin conditions -> Resolve ticket
        if ($null -ne $CloseQuestionWorkInfo)
        {
            Write-Host ((Get-LogPrefix $ToCloseticket) + "Closing ticket after happy to close ticket question")
            ResolveTicket $ToCloseticket $closeQuestionResolutionText
            continue
        }
   
        Write-Host ((Get-LogPrefix $ToCloseticket) + "Ticket not within closing condition")         
    }
}

function ResolveTicket{
      <#
    .SYNOPSIS
        Resolve Helix ticket with Status_Reason Customer Follow-Up Required

    .DESCRIPTION
        Resolve Helix ticket with Status_Reason Customer Follow-Up Required


    .PARAMETER Fields
        $[object] ToCloseticket: Ticket to cloes
        $[string] resolutionText: Resolution text

    .OUTPUTS
        Boolean succes or failed
     #>
    param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [object]$ToCloseticket,
    
    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$resolutionText
    )

    $resolutionCategory1 = $ToCloseTicket.'Resolution Category'
    $resolutionCategory2 = $ToCloseTicket.'Resolution Category Tier 2'
    $resolutionCategory3 = $ToCloseTicket.'Resolution Category Tier 3'

    #If resolution categories are not yet filled, make use of the insufficient Feedback Category
    if (($null -eq $resolutionCategory1) -or ($null -eq $resolutionCategory2) -or ($null -eq $resolutionCategory3))
    {
        #TODO - Change to Insufficient Feedback Category when available in Helix DEV
        $resolutionCategory1 = "ARM IT Core - Bonus System"
        $resolutionCategory2 = "DB"
        $resolutionCategory3 = "Insufficient Feedback Received"
        Write-Host ((Get-LogPrefix $ToCloseticket) + "Not all resolution categories filled, using insufficient feedback category." )
    }

    $UpdateFields = @{
        "Status"                     = "Resolved"
        "Status_Reason"              = "Customer Follow-Up Required"
        #Time logging seems to be non-functioning in Helix DEV. After filling time log, incident time logging is always null.
        "Time Log Duration (Min)"    = "1"
        "Resolution Category"        = $resolutionCategory1
        "Resolution Category Tier 2" = $resolutionCategory2
        "Resolution Category Tier 3" = $resolutionCategory3
        "Resolution"                 = $resolutionText
    }

    $UpdateIncidentResult = Update-QFHelixIncident -RequestID $ToCloseticket.'Request ID' -UpdateFields $UpdateFields
    if ($UpdateIncidentResult.Success -eq $true) {
        Write-Host ((Get-LogPrefix $ToCloseticket) + "Updated ticket with status Resolved and reason Customer Follow-Up required")
        return $true            
    }
    else {
        Write-Error ((Get-LogPrefix $ToCloseticket) + "Failed to close the ticket with a customer response. Error: " + $UpdateIncidentResult.Result)
        return $false
    }
}



function AddWorkingDays{
      <#
    .SYNOPSIS
        Adds an amount of days to the given date, excluding weekend days

    .DESCRIPTION
        Adds an amount of days to the given date, excluding weekend days. 
        Examples: 
        - Add 3 days to Tuesday returns the next Friday
        - Add 1 day to Friday returns the next Monday
        - Add 2 days to Friday returns the next Tuesday
        - Add 5 days to Wednesday returns the next Wednesday
        - Add 7 days to Wednesday returns the next Friday (9 days later)


    .PARAMETER Fields
        $date: input date
        $daysToAdd: Amount of days to add

    .OUTPUTS
        Date object with working days added
     #>
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [datetime]$date,
        
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$daysToAdd
    )
    
    $dateOutput = $date
    if($daysToAdd -ne 0)
    {
        for ($i=0; $i -lt $daysToAdd; $i++){
            if (($date.AddDays($i).DayOfWeek -match "Saturday") -or ($date.AddDays($i).DayOfWeek -match "Sunday"))
            {
                $dateOutput = $dateOutput.AddDays(2)
            }
            else {
                $dateOutput = $dateOutput.AddDays(1)
            }
        }
    }
    
    #Finally if result date ends on a Saturday or Sunday -> Change to next Monday
    if ($dateOutput.DayOfWeek -match "Saturday"){
        $dateOutput = $dateOutput.AddDays(2)
    }
    if ($dateOutput.DayOfWeek -match "Sunday"){
        $dateOutput = $dateOutput.AddDays(1)
    }
    $dateOutput
}