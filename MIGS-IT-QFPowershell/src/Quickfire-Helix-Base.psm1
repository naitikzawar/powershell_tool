###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                             Helix Base Functions                            #
#                                    v1.6.3                                   #
#                                                                             #
###############################################################################

# Author: Bernard Heije - bernhard.heije@derivco.es and Chris Byrne - christopher.byrne@derivco.com.au



#TODO-Password encryption on token request

# Create Helix ticket: https://derivco-dev-dwp.onbmc.com/dwp/app/#/srm/profile/SRHAAHKLFD2VYAOOR4P6E2VUCK5KFG/srm
# View/edit Helix ticket: https://derivco-dev-smartit.onbmc.com/smartit/app/#/ticket-console

# Initialize Module - set these values automatically when module loads
$ITSMtoken = [PSCustomObject]@{}
$AccessTokenURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Authorization/GenerateAccessToken"
$GetIncidentURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Incident/GetIncident"   
$CreateIncidentURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Incident/CreateIncident"
$UpdateIncidentURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Incident/UpdateIncident"    
$CreateIncidentWorkInfoURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/WorkInfo/CreateIncidentWorkInfo"
$CreateIncidentWorkInfoWithAttachmentURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/WorkInfo/CreateIncidentWorkInfoWithAttachment"
$GetIncidentWorkInfoURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/WorkInfo/GetIncidentWorkInfo"
$CreateTaskURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Task/CreateTask"
$GetTaskURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/Task/GetTask"
$CreateChangeRequestURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/ChangeRequest/CreateChangeRequest"
$GetChangeRequestURL = "https://dev.itsmcommunication.mgsops.net/ITSMCommunication/api/ChangeRequest/GetChangeRequest"


function Get-QFHelixDefaultHeader {
    <#
    .SYNOPSIS
        Generates the default request header for interaction with the Helix ITSM API.

    .DESCRIPTION
        This function generates a Helix Access Token and returns a hash table, which can be passed as a request header to Helix ITSM API.
        This function is generally called internally from other functions before calling the Helix ITSM API.
    
    .INPUTS
        This function takes no pipeline input.

    .OUTPUTS
        A System.Collections.Hashtable with the following members:
            x-api-version
            Authorization
            Content-Type

    #>
    $token = Get-QFHelixAccessToken
    $Header = @{
        "x-api-version" = "2.0"
        "Content-Type"  = "application/json"
        "Authorization" = "$token"
        "Cookie"        = "AR-JWT=$token"
    }
    $Header
}

function Get-QFHelixAccessToken {
    <#
    .SYNOPSIS
        Generates the authentication tokens required for the Helix ITSM API.

    .DESCRIPTION
        Generates the authentication tokens required for the Helix ITSM API using hardcoded credentials.
        By default, the service account 'RS-INT-MIGS-Automation-PowerShell' will be used, with a hardcoded password.
        The token will be output to pipeline, and also set in the script-scoped object $ITSMtoken
        This allows the Token object to persist after the function completes. 
        If $ITSMtoken is already set and the Token has not yet expired, a new Token will not be generated.
    
    .INPUTS
        This function takes no pipeline input.

    .OUTPUTS
        A String object containing a Helix ITSM Access Token.

    #>

    <#
    Not working for now
    #.PARAMETER TokenCredentials
        A PSCredential Object containing valid credentials for the Helix ITSM API. This will be used to generate an Access Token.
        If not specified, a default set of hardcoded credentials will be used.

    param (
        The username/password that you will use to request an ITSM Helix API Token
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PSCredential]$TokenCredentials
    )
    #>

    #If token not exists or is older then 10 minutes
    if ([string]::IsNullOrEmpty($ITSMtoken.TokenValue) -or (Get-Date) -gt $ITSMtoken.TokenDate.AddMinutes(10)) {
        If ($null -ne $TokenCredentials) {
            $TokenUserName = $TokenCredentials.UserName
            $TokenPassword = $TokenCredentials.Password
        }
        else {
            Write-Verbose ((Get-LogPrefix) + "Using default RS-INT-MIGS-Automation-PowerShell service account")
            $TokenUserName = "RS-INT-MIGS-Automation-PowerShell"
            $TokenPassword = "M1gsPowerSh3ll" 
        }
        
        $Headers = @{
            "x-api-version" = "2.0"
            "Content-Type"  = "application/json"
        }
    
        $Body = @{
            "username" = $TokenUserName
            "password" = $TokenPassword 
        } | ConvertTo-Json

        try {
            $response = Invoke-RestMethod $AccessTokenURL -Method 'POST' -Headers $Headers -Body $Body -SkipCertificateCheck
        }
        catch {
            Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixAccessToken")
            Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
            Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
            #Do not rethrow exception here, if obtaining token failed, this will be noticed in the parent function
        }
       
        $ITSMtoken | Add-Member -Name "TokenValue" -MemberType NoteProperty -Value $response -Force
        $ITSMtoken | Add-Member -Name "TokenDate" -MemberType NoteProperty -Value $(Get-Date) -Force
    }
    $ITSMtoken.TokenValue.ToString()
}


#For now no incident parameter on this function
#Just for testing - to create a incident to work with
function New-QFHelixIncident {
    <#
    .SYNOPSIS
        Creates a new Incident in the Helix ITSM system.

    # todo.... fill this out, once parameters etc are added. currently all values are hardcoded
    #>
    [CmdletBinding()]

    $headers = Get-QFHelixDefaultHeader
    #For now just a default ticket - to test with
    #Transaction ID new: 3643
    #Transaction ID old: 654
    $Body = @{
        "values" = @{
            "Description"                   = "CUST-REF 123456 (Powershell test ticket)"
            "Service_Type"                  = "User Service Request"
            "Impact"                        = "4-Minor/Localized"
            "Categorization Tier 1"         = "Markets Integrations and Gaming Services"
            "Categorization Tier 2"         = "MIGS IT - Customer Solutions"
            "Categorization Tier 3"         = "Gameplay Assessment"
            "Company"                       = "Derivco"
            "Assigned Group"                = "MIGS - Customer Solutions"
            "Assigned Group ID"             = "SGP000000000515"
            "Assigned Support Organization" = "Customer Service Desk"
            "Assigned Support Company"      = "Derivco"
            "Assignee"                      = "Christopher Byrne"
            "Status"                        = "New"
            "Reported Source"               = "Direct Input"
            "First_Name"                    = "Christopher"
            "Last_Name"                     = "Byrne"
            "z1D_Action"                    = "CREATE"
            "SRID"                          = "REQ-DUMMY"
            "Detailed_Decription"           = 
            "Op Cat Tier1: Markets Integrations and Gaming Services
Op Cat Tier2: MIGS IT - Customer Solutions
How Many Users Affected?: One
Brand?: Derivco
Your Reference?: CUST-REF 123456
Affected Market?: .com
Urgency?: 3-Medium
Date of Occurence?: 22/08/2023
Is this a potential regulated market breach?: Yes
Casino ID / Server ID / Product ID: 36793
Player ID / MGS Login Name: 2418~13569681
Game Round / Transaction IDs: 46963"
        }
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod $CreateIncidentURL -Method 'POST' -Headers $headers -Body $Body -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response.values[0].'Incident Number' 
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on New-QFHelixIncident")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
    }
    
    $CustomResponse
}


function Get-QFHelixIncident {
    <#
    .SYNOPSIS
        Retrieves an Incident and associated data from the Helix ITSM system, and parses text fields into a PSCustomObject.

    .DESCRIPTION
        Retrieves an Incident and associated data from the Helix ITSM system.
        The 'Detailed Description' field of the ticket will be parsed, and each field will be split into a hashtable as a key:value pair.
        This hashtable will be included in the pipeline output as a member named 'DescriptionFields'.

    .PARAMETER IncidentNumber
        The Incident Number of the ticket you wish to retrieve from the Helix ITSM system. This parameter should in the format 'INCxxxx'
        e.g. INC1234

    .EXAMPLE 
        Get-QFHelixIncident -IncidentNumber INC1234
            Requests all data for Incident INC1234 from the Helix ITSM system.

    .INPUTS
        This parameter will accept a String object on the pipeline, containing a ticket Incident Number.

    .OUTPUTS
        A PSCustomObject array, containing multiple NoteProperty members with the data from the retrieved Incident.
        The DescriptionFields member is a Hashtable containing the parsed output of the 'Detailed Description' field of the Incident.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$IncidentNumber
    )
    
    $headers = Get-QFHelixDefaultHeader
    $GetIncidentURL = $GetIncidentURL + "?fieldName=Incident Number&fieldValue=$IncidentNumber"

    try {
        $Response = Invoke-RestMethod $GetIncidentURL -Method 'GET' -Headers $Headers -SkipCertificateCheck

    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixIncident '$IncidentNumber'")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
        return $customRespons
    }

    # Process the Detailed Decription field and make it into a hashtable
    # also confirm Decription is correct spelling? is this a typo on the SRD? 
    $DescriptionFields = $Response.entries[0].values.'Detailed Decription' 
    
    # Regex to pull out Issue Description, this is a free-text multi line field so process seperately.
    If ($DescriptionFields -match "(?s)Issue Description\?: (.*)") {
        # We have an Issue Description field, get its contents into an object
        $IssueDescription = $Matches[1]
        # Remove everything after Issue Description to the end of the whole DescriptionFields string
        $DescriptionFields = $DescriptionFields -replace "(?s)Issue Description\?:.*$", ""

        # Check if there is any following fields. Assume any following field is seperated by "?: " and each field ends with a line feed. loop through until all fields are gone
        While ($IssueDescription -match "(?s).*\n(.*\?: .*$)") {
            $IssueDescription = $IssueDescription -replace '(?s)(^.*\n)(.*\?: .*)', '$1'
            $DescriptionFields += $Matches[1].trim()
        }
    }

    # Convert remaining fields into hashtable, splitting on newlines. Add each hashtable to a custom object (allows duplicate keys)
    $DescriptionFields = $DescriptionFields.trim() -split '\n'
    $DescriptionFieldsParsed = @()
    Foreach ($Field in $DescriptionFields) {
        # assuming field name and data is seperated by ': ', replace the first occurence of this pattern with ASCII 254
        $FieldData = $Field -replace '(^.+?): (.*$)', '$1■$2'
        # Use a regex to split field name and data on ASCII 254 symbol
        # if field has no data in it i.e. line ends after the first colon, it wont match the regex and will be ignored
        If ($FieldData -match '^(.+?)■(.+?)$') {
            $DescriptionFieldsParsed += @{$Matches[1].trim() = $Matches[2].trim() }
        }
    }
    
    # Finally output the entire response and our parsed description fields to pipeline
    $Output = $Response.entries[0].values
    $Output | Add-Member -MemberType NoteProperty -Name "DescriptionFields" -Value $([PSCustomObject]$DescriptionFieldsParsed) -force
    If ($null -ne $IssueDescription) {
        $Output | Add-Member -MemberType NoteProperty -Name "IssueDescription" -Value $IssueDescription.trim()
    }

    $CustomResponse = [PSCustomObject]@{
        Success = $true
        Result  = $Output 
    }
    $CustomResponse
}


function Search-QFHelixIncident {
    <#
    .SYNOPSIS
        Searches the Helix ITSM system for Incidents matching specified criteria and returns basic information for any matching Incidents.

    .DESCRIPTION
        Retrieves an Incident and associated data from the Helix ITSM system.
        The 'Detailed Description' field of the ticket will be parsed, and each field will be split into a hashtable as a key:value pair.
        This hashtable will be included in the pipeline output as a member named 'DescriptionFields'.

    .PARAMETER Fields
        Sets the fields that will be retrieved from the ITSM system, and output to pipeline for any Incidents that match the search query.

        This parameter is a multi-valued String object.
        Each member String should match the name of a field in a Helix Incident.
        A full list of Fields for a particular Incident can be retrieved using the cmdlet 'Get-QFHelixIncident'.

        If this parameter is not specified, a default list of Fields will be returned for any matching Incidents. The default list of fields is:
        "Incident Number", "Request ID","Assigned Group", "Status", "Submit Date", "Categorization Tier 3","SRID"

    .PARAMETER QueryField
        The Query String used to search for Incidents on the Helix ITSM system.
        This parameter is a single String object. It must contain a valid Helix ITSM query string.

        If this parameter is not specified, a default Query String is used. The default string is:
        Assigned Group="MIGS - Customer Solutions" AND Status="Assigned" AND Categorization Tier 3="Gameplay Error" AND Submit Date > Yesterday
        (The Submit Date will be dynamically set to be today's date minus 1 day. A specific date can be provided in the format YYYY-MM-DD.)

    .EXAMPLE 
        Search-QFHelixIncident
            Searches the Helix ITSM system for any Incidents matching the default QueryField parameter.

    .EXAMPLE 
        Search-QFHelixIncident -Fields "Incident Number","Request ID","Assigned Group"
            Searches the Helix ITSM system for any Incidents matching the default QueryField parameter.
            The contents of the fields Incident Number, Request ID, and Assigned Group, will be retrieved for each matching Incident and output to pipeline.

    .EXAMPLE 
        Search-QFHelixIncident -QueryField "`'Assigned Group`'=`"MIGS - Customer Solutions`" AND `'Status`'=`"Assigned`""
            Searches the Helix ITSM system for any Incidents matching the specified QueryField parameter.
            The provided example will return all open tickets currently assigned to the group "MIGS - Customer Solutions".
            
    .INPUTS
        This parameter will accept a multi-valued String object on the pipeline containing valid Field parameter values.
        A single String object can also be provided via pipeline for the QueryField value, containing a valid Query String for the Helix ITSM system.

    .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request
        
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Fields = @("Incident Number", "Request ID", "Assigned Group", "Status", "Submit Date", "Categorization Tier 3", "SRID", "Description"),

        [Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        # Limited query to Categorization Tier 3 = Gameplay Error, will need to adjust when SRD is finalized
        [string]$QueryField = "`'Assigned Group`'=`"MIGS - Customer Solutions`" AND `'Status`'=`"Assigned`" AND `'Categorization Tier 3`'=`"Gameplay Assessment`" AND `'Submit Date`'>`"" + $(Get-Date (get-date).AddDays(-2) -format 'yyyy-MM-dd') + '"'
    )
    
    $FieldsString = $Fields -join ","
    $GetIncidentURL = $GetIncidentURL + "?fields=$FieldsString&queryField=$QueryField"

    Write-Verbose "GetIncidentURL: $GetIncidentURL"
    $headers = Get-QFHelixDefaultHeader

    try {
        $Response = Invoke-RestMethod $GetIncidentURL -Method 'GET' -Headers $Headers -SkipCertificateCheck

        #Build a new array with incidents that have not been processed yet (no [GPA- in description)
        $incidents = @()
        foreach ($value in $Response.entries.values) {
                       
            if ((-not [string]::IsNullOrEmpty($value.'Description')) -AND $value.'Description' -cNotLike '`[GPA-*') {
                $incidents += $value
            }
        }

        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $incidents
        }
    }
    catch {

        $message = $_.ErrorDetails.Message | ConvertFrom-Json | Select-Object ErrorMessage
        # If there are no entries returned, return empty
        if ($message.ErrorMessage -eq "No entries were found") {
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = @()
            }
        }
        else {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
            Write-Error ((Get-LogPrefix) + "An error occured on Search-QFHelixIncident")
            Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
            Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
        }
    }
    $CustomResponse
}



function Update-QFHelixIncident {
    <#
    .SYNOPSIS
        Updates the specified Incident on the Helix ITSM system.

    .DESCRIPTION
        Updates the specified Incident on the Helix ITSM system. 
        This cmdlet can be used to change the status of an Incident, such as closing a request, or setting the Status Reason to 'Customer Action Required'.
        A hash table of Incident Field Names and corresponding Values must be provided, otherwise the Incident will not be updated.
    
    .PARAMETER RequestID
        The RequestID of the Helix Incident to be updated.
        Note that this is a different value then the Incident Number or Request Number. 
        This can be retrieved via Get-QFHelixIncident.

    .PARAMETER UpdateFields 
        A hashtable containing Incident Field names to be updated, and their new values.
        A list of field names can be retrieved via Get-QFHelixIncident.
        This will overwrite any values that are already present in these fields.

   .EXAMPLE 
        Update-QFHelixIncident -RequestID 'INC000000001234|INC000000001234' -UpdateFields @{
            "Status" = "Pending"
            "Status_Reason" = "Client Action Required"
        }

        Changes the Status of the specified Incident, to 'Pending', with a Status Reason of 'Client Action Required'.

    .EXAMPLE 
        Update-QFHelixIncident -RequestID 'INC000000001234|INC000000001234' -UpdateFields @{
            "Status" = "Resolved"
            "Status_Reason" = "No Further Action Required"
            "Resolution Category" = "Quickfire - Audit"
            "Resolution Category Tier 2" = "Win verification"
            "Resolution Category Tier 3" = "N/A"
            "Time Log Duration (Min)" = "5"
            "Resolution" = "Example Resolution"
        }

            Closes the specified Incident. 
            The Status Reason is set to 'No Further Action Required'.
            5 minutes of work time is logged against the Incident.
            The text of the resolution update will be set to 'Example Resolution'. This will be emailed to the customer and also visible in their Help Desk.

    .INPUTS
        This cmdlet will accept a string object via pipeline containing a RequestID of an Incident from the Helix ITSM system.
        This cmdlet will also accept a hashtable object containing Incident Field Names to be updated on the specified RequestID, and their corresponding new Values.

    .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request

    #>
    [CmdletBinding()]
    param (

        ##Attention: Request ID is used to update an incident. This is a different value then the Incident Number or Request Number
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RequestID,

        [Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$UpdateFields
       
    )

    $UpdateIncidentURL = $UpdateIncidentURL + "?fieldValue=$RequestID"

    $headers = Get-QFHelixDefaultHeader

    $Body = @{
        "values" = $UpdateFields
    } | ConvertTo-Json
    

    try {
        $Response = Invoke-RestMethod $UpdateIncidentURL -Method 'PUT' -Headers $Headers -Body $Body
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on Update-QFHelixIncident for incident $RequestID")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    
    $CustomResponse
}

function New-QFHelixIncidentWorkInfo {
    <#
        .SYNOPSIS
            Creates a new Work Info on the specified Incident on the Helix ITSM system.

        .DESCRIPTION
            Creates a new Work Info on the specified Incident on the Helix ITSM system.
            A hash table of Field Names and corresponding Values must be provided, otherwise the Workinfo will not be created
            The Work Info can bet set to to Public or Internal visibility using the "View Access" update field.
            
        
        .PARAMETER IncidentNumber
            The Incident Number of the Helix Incident to be updated. e.g. INC123456
            This can be retrieved via Get-QFHelixIncident.

        .PARAMETER WorkInfoFields 
            A hashtable containing Incident Field names to be updated, and their new values.
            A list of field names can be retrieved via Get-QFHelixIncident.
            This will overwrite any values that are already present in these fields.
            $WorkInfoFields = @{
                "Work Log Type"        = "Working Log"
                "View Access"          = "Public"
                "Detailed Description" = "Workinfo description"
            }

        .PARAMETER files 
            A string[] with the full file paths (maximum 3 files)

        .EXAMPLE 
            New-QFHelixIncidentWorkInfo -IncidentNumber $GPATicket.'Incident Number' -WorkInfoFields $WorkInfoFields -Files $Files
        


        .INPUTS
            This cmdlet will accept a string object via pipeline containing an Incident Number of an Incident from the Helix ITSM system.
            This cmdlet will also accept a hashtable object containing the Field Names to be updated on the new Work Info, and their corresponding new Values.

        .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$IncidentNumber,

        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$WorkInfoFields,

        [Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Files
    )

    $headers = Get-QFHelixDefaultHeader 
    $WorkInfoFields.Add("Incident Number", $IncidentNumber)

    if ([string]::IsNullOrEmpty($Files) -or $Files.Count -eq 0) {
        #No files to attached - call normal $CreateIncidentWorkInfo without attachment

        $Body = @{
            "values" = $WorkInfoFields
        } | ConvertTo-Json

        try {
            $CreateWorkInfoResponse = Invoke-RestMethod $CreateIncidentWorkInfoURL -Method 'POST' -Headers $headers -Body $body
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = $CreateWorkInfoResponse[0].values
            }
        }
        catch {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
            Write-Error ((Get-LogPrefix) + "An error occured on New-QFHelixIncidentWorkInfo for incident $IncidentNumber")
            Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
            Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
        }
    }
    else {
        #There are attachments to be added, use alternate endpoint: $CreateIncidentWorkInfoWithAttachmentURL
        #For this endpoint the message body and the attachment are passed in a form

        #Loop through files; build fileitem array and add attachments to workInfoFields
        $attachmentPrefix = 'z2AF Work Log0'
        $fileItems = @()
        for ($i = 0; $i -lt $Files.Count; $i++) {
            $fileItems += (Get-Item -path $Files[$i])
            $WorkInfoFields.Add($attachmentPrefix + ($i + 1), $fileItems.Get($i).Name.ToString())
        }

        $Body = @{
            "values" = $WorkInfoFields
        } | ConvertTo-Json

        $Form = @{
            jsonObject  = $Body
            attachments = $fileItems
        }
        try {
            $CreateWorkInfoResponse = Invoke-RestMethod $CreateIncidentWorkInfoWithAttachmentURL -Method 'POST' -Headers $headers -Form $Form
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = $CreateWorkInfoResponse[0].values
            }
        }
        catch {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
            Write-Error ((Get-LogPrefix) + "An error occured on New-QFHelixIncidentWorkInfo for incident $IncidentNumber")
            Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
            Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
        }
    }
    $CustomResponse
}

function Get-QFHelixIncidentWorkInfo {
    <#
        .SYNOPSIS
            Retrieves all Work Info from the specified Incident on the Helix ITSM system.

        .DESCRIPTION
            Retrieves all Work Info from the specified Incident on the Helix ITSM system.
            This cmdlet will output all Work Info on the specified Incident as an array of PSCustomObjects.

        .EXAMPLE
            Get-QFHelixIncidentWorkInfo -IncidentNumber INC123456
                Retrieves all Work Info  from the specified Incident and outputs to pipeline.

        .EXAMPLE     
            Get-QFHelixIncidentWorkInfo -IncidentNumber INC123456 | Sort-Object 'Submit Date' -Descending | Select-Object -First 1
                Retrieves only the most recently creeated Work Info from the specified Incident and outputs to pipeline.

        .INPUTS
            A String object containing an Incident Number can be piped to this cmdlet.

        .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$IncidentNumber,

        [Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Fields = @("Incident Number", "Work Log ID", "Detailed Description", "Work Log Type", "View Access", "Work Log Submit Date")
    )

    $headers = Get-QFHelixDefaultHeader

    $FieldsString = $Fields -join ","
    $GetIncidentWorkInfoURL = $GetIncidentWorkInfoURL + "?fieldName=Incident Number&fieldValue=$IncidentNumber&fields=$FieldsString" 
    
    try {
        $Response = Invoke-RestMethod $GetIncidentWorkInfoURL -Method 'GET' -Headers $Headers -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response.entries.values
        }
    }
    catch {
        $message = $_.ErrorDetails.Message | ConvertFrom-Json | Select-Object ErrorMessage
        # If there are no entries returned, return empty
        if ($message.ErrorMessage -eq "No entries were found") {
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = @()
            }
        }
        else {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
            Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixIncidentWorkInfo for incident $IncidentNumber")
            Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
            Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
        }
    }
    $CustomResponse
}


function New-QFHelixTask {
    <#
        .SYNOPSIS
            Creates a new Task on the specified Incident on the Helix ITSM system.

        .DESCRIPTION
            Creates a new Task on the specified Incident on the Helix ITSM system.
            A hash table of Field Names and corresponding Values must be provided, otherwise the task will not be created
            
        
        .PARAMETER RootRequestID
            The Parent Requests ID eg. Incident Number for Incident (INC000000001234), Work Order ID for Work Orders(WO0000000001234) etc.
            This can be retrieved via Get-QFHelixIncident.

        .PARAMETER InstanceID
            The Instance Id of the Request (WO/INC/CRQ) you want to create the Task under eg. if you want to create a Task for a Work Order you would find its Instance Id

        .PARAMETER UpdateFields 
            A hashtable containing Task Field names and their new values.
            [Hashtable]$UpdateFields = @{
            "Parent Type"           = "Root Request"
            "TaskName"              = "Default MIGS PS test task - TaskName"
            "TaskSummary"           = "Default MIGS PS test task - TaskSummary"
            "TaskType"              = "Manual"
            "Status"                = "Staged"
            "Location Company"      = "Derivco"
            "RootRequestMode"       = "0"
            "Company"               = "Derivco"
            "First Name"            = "Bernhard"
            "Last Name"             = "Heije"
            "Customer Company"      = "Derivco"
            "Customer First Name"   = "Bernhard"
            "Customer Last Name"    = "Heije"
            }

        .PARAMETER RootRequestFormName
            For different type of parent a Specific RootRequestFormName is required. By default HPD:Help Desk (incident) is set
            For Work Order: WOI:WorkOrder
            For Incident: HPD:Help Desk
            For Change Request: CHG:Infrastructure Change


        .EXAMPLE 
            TODO 

        .INPUTS
            IncidentNumber: The incident number on which to create a task
            TaskFields: 

        .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootRequestID,

        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootInstanceID,

        [Parameter(Mandatory = $false, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$UpdateFields = @{
            "Parent Type"         = "Root Request"
            "TaskName"            = "Default MIGS PS test task - TaskName"
            "Summary"             = "Default MIGS PS test task - TaskSummary"
            "TaskType"            = "Manual"
            "Status"              = "Staged"
            "Location Company"    = "Derivco"
            "RootRequestMode"     = "0"
            "Company"             = "Derivco"
            "First Name"          = "Bernhard"
            "Last Name"           = "Heije"
            "Customer Company"    = "Derivco"
            "Customer First Name" = "Bernhard"
            "Customer Last Name"  = "Heije"
        },

        [Parameter(Mandatory = $false, Position = 3, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootRequestFormName = 'HPD:Help Desk'

    )
    #TODO: Having issue with assigning it to a group.. API returns "Failed to Create Entity, please verify accuracy of input values. ERROR(51001):  The Support Group \"MIGS - Customer Solutions\" is not configured for assignment to either \"Derivco\" or \"Derivco\". Please contact your Administrator.
    $headers = Get-QFHelixDefaultHeader

    $UpdateFields.Add("RootRequestName", $RootRequestID)
    $UpdateFields.Add("RootRequestID", $RootRequestID)
    $UpdateFields.Add("ParentID", $RootInstanceID)
    $UpdateFields.Add("RootRequestInstanceID", $RootInstanceID)
    $UpdateFields.Add("RootRequestFormName", $RootRequestFormName)

    $Body = @{
        "values" = $UpdateFields
    } | ConvertTo-Json

    try {
        $CreateTaskResponse = Invoke-RestMethod $CreateTaskURL -Method 'POST' -Headers $headers -Body $body
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $CreateTaskResponse[0].values
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on New-QFHelixTask on root $RootRequestID")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    
    $CustomResponse
}

function Get-QFHelixTask {
    <#
    .SYNOPSIS
        Retrieves a Task and associated data from the Helix ITSM system

    .DESCRIPTION
        Retrieves an Task and associated data from the Helix ITSM system.

    .PARAMETER IncidentNumber
        The Task ID of the ticket you wish to retrieve from the Helix ITSM system. This parameter should in the format 'TASxxxx'
        e.g. TAS000000002218

    .EXAMPLE 
        Get-QFHelixTask  -TaskID TAS000000002218
            Requests all data for Task TAS000000002218 from the Helix ITSM system.

    .INPUTS
        This parameter will accept a String object on the pipeline, containing a Task Number.

    .OUTPUTS
        A PSCustomObject array, containing multiple members with the data from the retrieved Task.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskID
    )
    
    $headers = Get-QFHelixDefaultHeader
    $GetTaskURL = $GetTaskURL + "?fieldName=Task ID&fieldValue=$TaskID"

    try {
        $Response = Invoke-RestMethod $GetTaskURL -Method 'GET' -Headers $Headers -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response[0].entries.values
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixTask '$TaskID'")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
        
    }
    return $CustomResponse
}


function Get-QFHelixTasksByRootRequest {
    <#
    .SYNOPSIS
        Returns all tasks related to the given root Request

    .DESCRIPTION
        Retrieves all tasks and associated data for the give root Request from the Helix ITSM system.

    .PARAMETER RootRequestID
        RootRequestID for which to get the tasks from. For example INC000000007208

    .PARAMETER Fields
        It is optional to pass a list of fields that should be returned for each task. If parameter is not passed, all task fields will be returned


    .EXAMPLE 
        Get-QFHelixTasksByRootRequest 'INC000000007208'
           Returns all tasks for incident INC000000007208

            
    .OUTPUTS
            A PSCustom Object containing 
             - Success: Succeeded Yes / No
             - Result: Response object(s) / Error message
        
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootRequestID,

        [Parameter(Mandatory = $false, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Fields = $null
    )
    

    $headers = Get-QFHelixDefaultHeader
    $GetTaskURL = $GetTaskURL + "?fieldName=RootRequestID&fieldValue=$RootRequestID"

    if ($null -ne $Fields) {    
        $FieldsString = $Fields -join ","
        $GetTaskURL += "&fields=$FieldsString"
    }
    
    try {
        $Response = Invoke-RestMethod $GetTaskURL -Method 'GET' -Headers $Headers -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response.entries.values
        }
    }
    catch {
        $message = $_.ErrorDetails.Message | ConvertFrom-Json | Select-Object ErrorMessage
        # If there are no entries returned, return empty
        if ($message.ErrorMessage -eq "No entries were found") {
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = @()
            }
        }
        else {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
        
            Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixTask '$TaskID'")
            Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
            Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
        }
    }
    return $CustomResponse

}

function New-QFHelixChangeRequest {
        <#
    .SYNOPSIS
        Creates a new CRQ in the Helix ITSM system.

    .DESCRIPTION
        Creates a new CRQ in the Helix ITSM system.

    .PARAMETER UpdateFields
        A hashtable containing CRQ Field names, and their  values. If this parameter is not passed, a default test CRQ will be created
        
        Example:
        [Hashtable]$UpdateFields = @{
            "Last Name" = "Heije"
            "First Name" = "Bernhard"
            "Submitter" = "heijeb"
            "Description" = "ITSM Communication API - Testing CRQ Creation"
            "Location Company" = "Derivco"
            "Company" = "Derivco"
            "Company3" = "Derivco"
            "Support Organization" = "Customer Service Desk"
            "Support Group Name" = "MIGS - Customer Solutions"
            }

    .EXAMPLE 
        New-QFHelixCRQ
        New-QFHelixCRQ $UpdateFields

    .OUTPUTS
            A PSCustom Object containing 
             - Success: Yes / No
             - Result: CRQ number / Error message
        
    #>
    
    [CmdletBinding()]
    param (
    [Parameter(Mandatory = $false, Position = 0, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [Hashtable]$UpdateFields = @{
            "Last Name" = "Heije"
            "First Name" = "Bernhard"
            "Submitter" = "heijeb"
            "Description" = "ITSM Communication API - Testing CRQ Creation"
            "Location Company" = "Derivco"
            "Company" = "Derivco"
            "Company3" = "Derivco"
            "Support Organization" = "Customer Service Desk"
            "Support Group Name" = "MIGS - Customer Solutions"
            }
    )

    $headers = Get-QFHelixDefaultHeader

    $Body = @{
        "values" = $UpdateFields
        } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod $CreateChangeRequestURL -Method 'POST' -Headers $headers -Body $Body -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response.values[0].'Infrastructure Change Id' 
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on New-QFHelixCRQ")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
    }
    
    $CustomResponse
}

function Get-QFHelixChangeRequest {
    <#
    .SYNOPSIS
        Retrieves a ChangeRequest and associated data from the Helix ITSM system

    .DESCRIPTION
        Retrieves an ChangeRequest and associated data from the Helix ITSM system.

    .PARAMETER ChangeRequestID
        The 'Infrastructure Change ID' of the CRQ you wish to retrieve from the Helix ITSM system. This parameter should in the format 'CRQxxxx'
        e.g. CRQ14614

    .EXAMPLE 
        Get-QFHelixTask -ChangeRequestID CRQ14614
            Requests all data for CRQ CRQ14614 from the Helix ITSM system.

    .INPUTS
        This parameter will accept a String object on the pipeline, containing a CRQ Number.

    .OUTPUTS
        A PSCustom Object containing 
             - Success: Yes / No
             - Result: CRQ data / Error message

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ChangeRequestID
    )
    
    $headers = Get-QFHelixDefaultHeader
    $GetChangeRequestURL = $GetChangeRequestURL + "?fieldName=Infrastructure Change ID&fieldValue=$ChangeRequestID"

    try {
        $Response = Invoke-RestMethod $GetChangeRequestURL -Method 'GET' -Headers $Headers -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response[0].entries.values
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on Get-QFHelixCRQ  '$ChangeRequestID'")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
        
    }
    return $CustomResponse
}

function Get-QFHelixCRQsByRequestID {
    #TODO 
    #Ideally we want to lookup all CRQ's linked to a certain INC.
    #I do not know if this is possible
    #At the moment I do not have CRQ creation/viewing rights on Helix DEV, which makes this a bit tedious to develop.
}


function Get-LogPrefix {
    <#
    .SYNOPSIS
        Returns a log prefix using datetime and the REQ + INC number

    .DESCRIPTION
        Returns a log prefix using datetime and the REQ + INC number
        Format: [dd/MM/yyyy hh:mm:ss] [REQ-INC]

    .INPUTS
        GPATicket object (optional), from which the REQ and INC numbers are read

    .OUTPUTS   
        Returns the log prefix string
        Format: [dd/MM/yyyy hh:mm:ss] [REQ-INC]
                
    #>
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [object]$Ticket
    )

    if ($null -eq $Ticket) {
        "[$(Get-Date -Format "dd/MM/yyyy HH:mm:ss")] "
    }
    else {
        "[$(Get-Date -Format "dd/MM/yyyy HH:mm:ss")] [" + $Ticket.sRID + "-" + $Ticket.'Incident Number' + "] "
    }
    
}

