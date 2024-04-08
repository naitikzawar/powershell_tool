
#Define Remedy environment to be used here. DEV and PROD are supported. This is the only location where DEV/PROD switch has to be set
#$Global:RemedyEnvironment = "PROD"
$Global:RemedyEnvironment = "DEV"

$OktaTokenURL = $null
$OktaUsername = $null
$OktaPassword = $null
$OktaAuthorization = $null
$LogIncidentURL = $null
$GetIncidentURL = $null
$UpdateIncidentURL = $null
$ResolveIncidentURL = $null
$CreateIncidentWorkInfoURL = $null
$GetIncidentWorkInfoURL = $null


switch ($Global:RemedyEnvironment) {
    #------------------   DEV configuration  ------------------
    #View tickets Remedy DEV
    #https://dev.queues.canvas.mgsops.net/dashboard/

    #Create tickets Remedy DEV
    #http://der2431:9000/dwp/app/#/srm/profile/SRHAAHKLFD2VYAOOR4P6E2VUCK5KFG/srm
    "DEV" {
        $OktaTokenURL = "https://derivco.oktapreview.com/oauth2/default/v1/token"
        $OktaUsername = "ok-IntAppsTest@derivcoservice.com"
        $OktaPassword = "8D]pQWa!"
        $OktaAuthorization = "Basic  MG9hZGJrcW1udlFMSDNkU1AwaDc6clFkdGFzRC1mY3NGZWZFenR5ZDEwWEZzOUZVbVZrLV9jb3d6dzFUZA=="
        $LogIncidentURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/logIncident"
        $GetIncidentURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/GetIncident"
        $UpdateIncidentURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/UpdateIncident"
        $ResolveIncidentURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/ResolveIncident"
        $CreateIncidentWorkInfoURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/CreateWorkInfo"
        $GetIncidentWorkInfoURL = "https://dev.remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/WorkInfos"
    }
    #------------------   PROD configuration  ------------------
    "PROD" {
        $OktaTokenURL = "https://derivco.okta-emea.com/oauth2/default/v1/token"
        $OktaUsername = "ok-remedyintegrationapi-migs-cs@derivcoservice.com"
        $OktaPassword = "Jaf935&AFk1!£agkHGA24rf"
        $OktaAuthorization = "Basic  MG9hMWl4eWxuZEdDV1FiVFEwaTc6RmhnamppaVMyNkpsN05XR1U5UjR2YTI4Q2ZabGhVMkd1QUtHTTVvbQ=="
        $LogIncidentURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/logIncident"
        $GetIncidentURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/GetIncident"
        $UpdateIncidentURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/UpdateIncident"
        $ResolveIncidentURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/ResolveIncident"
        $CreateIncidentWorkInfoURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/CreateWorkInfo"
        $GetIncidentWorkInfoURL = "https://remedyintegrationapi.mgsops.net/api/GenericIncidentRequest/WorkInfos"
        $GPARemedyRSSFeedURL = "http://quickfirerss/rss/incidents/v2"
    }
}

#------------------   GENERAL configuration  ------------------
$OktaToken = [PSCustomObject]@{}




function Get-QFRemedyDefaultHeader {
    <#
    .SYNOPSIS
        Generates the default request header for interaction with the Remedy Integration API.

    .DESCRIPTION
        This function generates a Okta Token and returns a hash table, which can be passed as a request header to Remedy Integration API.
        This function is generally called internally from other functions before calling the Remedy Integration API.
    
    .INPUTS
        This function takes no pipeline input.

    .OUTPUTS
        A System.Collections.Hashtable with the following members:
            x-api-version
            Authorization
            Content-Type

    #>
    $token = Get-QFRemedyOktaToken
    $Header = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $token"
    }
    $Header
}


function Get-QFRemedyOktaToken {
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

    #If token not exists or is older then 10 minutes
    if ([string]::IsNullOrEmpty($OktaToken.TokenValue) -or (Get-Date) -gt $OktaToken.TokenDate.AddMinutes(10)) {

        Write-Host ((Get-LogPrefix) + "Obtaining token for $OktaUsername")
      
        $Headers = @{
            #"x-api-version" = "2.0"
            "Accept"        = "application/json"
            "Authorization" = $OktaAuthorization
            "Content-Type"  = "application/x-www-form-urlencoded"
            #"Cookie"        = "DT=DI1QZ2HdiMHQjmStFO11bUekw; JSESSIONID=FE8355991B05D7B475BB5334948EC85E; Okta_Verify_Autopush_-1212278007=false; enduser_version=2"
            #"Cookie"        = "JSESSIONID=B98A754F603728A9FF55B21F073338E1"
            
        }

        $Form = @{
            "username" = $OktaUsername
            "password" = $OktaPassword
            "grant_type" = "password"
            "scope" = "openid roles"
        }

        try {
            $response = Invoke-RestMethod $OktaTokenURL -Method 'POST' -Headers $Headers -Body $Form -SkipCertificateCheck
        }
        catch {
            Write-Error ((Get-LogPrefix) + "An error occured on Get-QFRemedyOktaToken ")
            Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
            Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
            #Do not rethrow exception here, if obtaining token failed, this will be noticed in the parent function
        }
       
        $OktaToken | Add-Member -Name "TokenValue" -MemberType NoteProperty -Value $response.access_token -Force
        $OktaToken | Add-Member -Name "TokenDate" -MemberType NoteProperty -Value $(Get-Date) -Force
    }
    $OktaToken.TokenValue.ToString()
}



#For now no incident parameter on this function
#Just for testing - to create a incident to work with
function New-QFRemedyIncident {
    <#
    .SYNOPSIS
        Creates a new Incident in the Helix ITSM system.

    # todo.... fill this out, once parameters etc are added. currently all values are hardcoded
    #>
    [CmdletBinding()]

    $headers = Get-QFRemedyDefaultHeader
    #For now just a default ticket - to test with
    <#
Op Cat Tier1: Markets Integrations and Gaming Services
Op Cat Tier2: MIGS IT - Customer Solutions
How Many Users Affected?: One
Brand?: Derivco
Your Reference?: CUST-REF 123456
Affected Market?: .com
Urgency?: 3-Medium
Date of Occurence?: 22/08/2023
Is this a potential regulated market breach?: Yes
#>
    $Body = @{
        "RequestedFor_FirstName"= "Bernhard"
        "RequestedFor_LastName"= "Heije"
        "Requested_By_FirstName"= "Bernhard"
        "Requested_By_LastName"= "Heije"
        "Summary"= "MIGS CS Test ticket"
        "Status"= "New"
        "Urgency"= "Medium"
        "Site"= "Derivco Durban  FP 1"
        "Support_Organization"= "Customer Service Desk"
        "Support_Group"= "MIGS - Customer Solutions"
        "SupportCompany"= "Derivco"
        "Channel"= "Quickfire"
        "RemedyUsername"= "bernhardh"  
        "OperatorId"="56718"
        "Notes"= 
        "Casino Gameplay Related To?: Gameplay Assessment
Casino ID / Server ID / Product ID: 2512
Player ID / MGS Login Name: HQ_97286SZ
Game Round / Transaction IDs: 86957"
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod $LogIncidentURL -Method 'POST' -Headers $headers -Body $Body -SkipCertificateCheck
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $Response.'IncidentNumber' 
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


function Get-QFRemedyIncident {
    <#
    .SYNOPSIS
        Retrieves an Incident and associated data from the Remedy system, and parses text fields into a PSCustomObject.

    .DESCRIPTION
        Retrieves an Incident and associated data from the Remedy system.
        The 'Detailed Description' field of the ticket will be parsed, and each field will be split into a hashtable as a key:value pair.
        This hashtable will be included in the pipeline output as a member named 'DescriptionFields'.

    .PARAMETER IncidentNumber
        The Incident Number of the ticket you wish to retrieve from the Remedy system. This parameter should in the format 'INCxxxx'
        e.g. INC1234

    .EXAMPLE 
        Get-QFRemedyIncident -IncidentNumber INC1234
            Requests all data for Incident INC1234 from the Remedy system.

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
    
    $headers = Get-QFRemedyDefaultHeader
    $headers["Content-Type"] = "application/x-www-form-urlencoded"
    $body = @{
        "incidentID"  = $IncidentNumber
    }

    try {
        $Response = Invoke-RestMethod $GetIncidentURL -Method 'POST' -Body $body -Headers $Headers -SkipCertificateCheck
        if ($Response -eq 'The Incident is not found with provided IncidentNumber/RequestId'){
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $Response 
            }
            return $CustomResponse
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        
        Write-Error ((Get-LogPrefix) + "An error occured on Get-QFRemedyIncident '$IncidentNumber'")
        Write-Error ((Get-LogPrefix) + "$_.Exception.Response.StatusCode.value__")
        Write-Error ((Get-LogPrefix) + "$_.ErrorDetails.Message")
        return $customRespons
    }

    # Process the Notes field and make it into a hashtable
    $DescriptionFields = $Response.Notes
    
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
    $Output = $Response
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

function Search-QFRemedyIncidentsDummy {

    $incidents = @()
    $incident = @{
        'IncidentNumber' = 'INC1240261'
        'RequestId' = 'REQ1423761'
    }
    $incidents += $incident
        
    $CustomResponse = [PSCustomObject]@{
        Success = $true
        Result  = $incidents
    }
    $CustomResponse
}


function Search-QFRemedyGPAIncidents {
    <#
    .SYNOPSIS
        Requests the RSS feed for Incidents matching specified criteria and returns basic information for any matching Incidents.

    .DESCRIPTION
        Retrieves an Incident and associated data from the RSS feed

    .EXAMPLE 
        Search-QFRemedyGPAIncidents
            Requests the data from the RSS feed any Incidents matching the default QueryField parameter.
            

    .OUTPUTS
            A PSCustom Object containing 
             - Boolean: Request succeded Yes / No
             - String: Request response / Request
        
    #>
    
   

    #Ticket list to be retrieved from Remedy DB:
    #Assigned Group = 'MIGS - Customer Solutions' 
    #Status = 'New'
    #Notes like '%Casino Gameplay Related To?: Gameplay Assessment%'
    #Created < 2 days ago
    #Teamnotes not like '%[GPA-SUCCESS]%'
    #Teamnotes not like '%[GPA-FAILED]%'


    #Fields we want returned: INC number, REQ number
    try {
        $Response = Invoke-RestMethod $GPARemedyRSSFeedURL -Method 'GET' -SkipCertificateCheck
        #TODO: test behaviour in case not tickets are returned

        #Parse response. For each returned ticket add REQ and INC to incidents
        $incidents = @()
        foreach ($item in $Response)
        {
            $titleSplitted = $item.title.split(' || ')
            $incident = @{
                'IncidentNumber' = $titleSplitted[0]
                'RequestId' = $titleSplitted[1]
            }

            $incidents += $incident
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
        Write-Error ((Get-LogPrefix) + "An error occured on Search-QFRemedyIncident feed")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    $CustomResponse
}

function Update-QFRemedyIncident {
    <#
    .SYNOPSIS
        Updates the specified Incident on the Remedy system.

    .DESCRIPTION
        Updates the specified Incident on the Remedy system. 
        Note that for resolving an incident, a different function (and URL) is used 
    
    .PARAMETER RequestID
        The IncidentNumber of the Remedy Incident to be updated.


    .PARAMETER UpdateFields 
        A hashtable containing Incident Field names to be updated, and their new values.
        A list of field names can be retrieved via Get-QFRemedyIncident.
        This will overwrite any values that are already present in these fields.

   .EXAMPLE 
        Update-QFRemedyIncident -IncidentNumber 'INC1240178' $UpdateFields {
            "TeamNotes" = "Teamnotes update"
        }

    .INPUTS
        This cmdlet will accept a string object via pipeline containing a IncidentNumber of an Incident from the Remedy system.
        This cmdlet will also accept a hashtable object containing Incident Field Names to be updated on the specified Incident number, and their corresponding new Values.

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
        [Hashtable]$UpdateFields
       
    )


    $headers = Get-QFRemedyDefaultHeader
    $UpdateFields.Add("IncidentNumber", $IncidentNumber)
    $Body = $UpdateFields | ConvertTo-Json
    

    try {
        $Response = Invoke-RestMethod $UpdateIncidentURL -Method 'POST' -Headers $Headers -Body $Body
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
        Write-Error ((Get-LogPrefix) + "An error occured on Update-QFRemedyIncident for incident $IncidentNumber")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    
    $CustomResponse
}


function Resolve-QFRemedyIncident {
    <#
    .SYNOPSIS
        Resolves the specified Incident on the Remedy system.

    .DESCRIPTION
        This cmdlet can be used to resolve an Incident
        A hash table of Incident Field Names and corresponding Values must be provided, otherwise the Incident will not be updated.
    
    .PARAMETER Incident
        The IncidentNumber of the Remedy Incident to be updated.


    .PARAMETER UpdateFields 
        A hashtable containing Incident Field names to be updated, and their new values.
        A list of field names can be retrieved via Get-QFRemedyIncident.
        This will overwrite any values that are already present in these fields.

    .EXAMPLE 
        Update-QFRemedyIncident -IncidentNumber 'INC1240178' -UpdateFields @{
	        "StatusReason"="No Further Action Required"
	        "ResolutionText"="test close ticket"
	        "ResolutionMethod"="Remedy"
	        "DetailedRootCause"="Operator – Skills and Knowledge (Operator)"
	        "ServiceCategory"="Quickfire"
            "ServiceCategoryTier1"= "Operator - Audit"
            "ServiceCategoryTier2"= "Win Legitimacy"
	        "RemedyUsername"="bernhardh"
        }

            Closes the specified Incident. 
            The Status Reason is set to 'No Further Action Required'.

    .INPUTS
        This cmdlet will accept a string object via pipeline containing a IncidentNumber of an Incident from the Remedy system.
        This cmdlet will also accept a hashtable object containing Incident Field Names to be updated on the specified Incident number, and their corresponding new Values.

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
        [Hashtable]$UpdateFields
       
    )


    $headers = Get-QFRemedyDefaultHeader
    $UpdateFields.Add("IncidentNumber", $IncidentNumber)
   # $Body = $UpdateFields | ConvertTo-Json
    $Body = ([System.Text.Encoding]::UTF8.GetBytes(($UpdateFields | ConvertTo-Json)))

    try {
        #$Response = Invoke-RestMethod $ResolveIncidentURL -Method 'POST' -Headers $Headers -Body $Body
        $Response = Invoke-RestMethod $ResolveIncidentURL -Method 'POST' -Headers $Headers -Body $Body -ContentType 'application/json; charset=utf8'
        if ($Response -eq 'Incident Updated Successfully') {
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = $Response
            }
        } else {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $Response
            }
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on Resolve-QFRemedyIncident for incident $IncidentNumber")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }
    
    $CustomResponse
}


function New-QFRemedyIncidentWorkInfo {
    <#
        .SYNOPSIS
            Creates a new Work Info on the specified Incident on the Remedy system.

        .DESCRIPTION
            Creates a new Work Info on the specified Incident on the Remedy system.
            A hash table of Field Names and corresponding Values must be provided, otherwise the Workinfo will not be created
            The Work Info can bet set to to Public or Internal visibility using the "View Access" update field.
            
        
        .PARAMETER IncidentNumber
            The Incident Number of the Remedy Incident to be updated. e.g. INC123456
            This can be retrieved via Get-QFRemedyIncident.

        .PARAMETER WorkInfoFields 
            A hashtable containing Incident Field names to be updated, and their new values.
            A list of field names can be retrieved via Get-QFHelixIncident.
            This will overwrite any values that are already present in these fields.
            $WorkInfoFields = @{
                "WorkInfoType"      = "Status Update"
                "ViewAccess"       = "Public"
                "Summary"           = "Update"
                "Notes"             = "Workinfo description"
            }

        .PARAMETER files 
            A string[] with the full file paths (maximum 3 files)
            [string[]] $files = 'C:\test1.zip', 'C:\test2.zip'  

        .EXAMPLE 
            New-RemedyIncidentWorkInfo -IncidentNumber 'INC1240183' -WorkInfoFields $WorkInfoFields -Files $Files
        


        .INPUTS
            This cmdlet will accept a string object via pipeline containing an Incident Number of an Incident from the Remedy system.
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

    $headers = Get-QFRemedyDefaultHeader 
    $WorkInfoFields.Add("Id", $IncidentNumber)

    #Add files to body
    $attachmentPrefix = 'Attachment'
    if ($null -ne $Files -and $Files.Count -gt 0){
        for ($i = 0; $i -lt $Files.Count; $i++) {
            $fileItem = (Get-Item -path $Files[$i])
            
            #Get filename
            $fileName = $fileItem.Name.ToString()
            
            #Get Base64Encoded file content
            $base64string = [Convert]::ToBase64String([IO.File]::ReadAllBytes($fileItem))   
            
            $Attachment = 
            @{
                "FileName"= $fileName
                "FileBytes"= $base64string
            }
            $WorkInfoFields.Add($attachmentPrefix + ($i + 1), $Attachment)
        }
    }

    $Body = $WorkInfoFields | ConvertTo-Json


        try {
            $CreateWorkInfoResponse = Invoke-RestMethod $CreateIncidentWorkInfoURL -Method 'POST' -Headers $headers -Body $Body
            $CustomResponse = [PSCustomObject]@{
                Success = $true
                Result  = $CreateWorkInfoResponse
            }
        }
        catch {
            $CustomResponse = [PSCustomObject]@{
                Success = $false
                Result  = $_.ErrorDetails.Message
            }
            Write-Error ((Get-LogPrefix) + "An error occured on New-RemedyIncidentWorkInfo for incident $IncidentNumber")
            Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
            Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
        }
    
    $CustomResponse
}


function Get-QFRemedyIncidentWorkInfo {
    <#
        .SYNOPSIS
            Retrieves all Work Info from the specified Incident on the Remedy system.

        .DESCRIPTION
            Retrieves all Work Info from the specified Incident on the Remedy system.
            This cmdlet will output all Work Info on the specified Incident as an array of PSCustomObjects.

        .EXAMPLE
            Get-QFRemedyIncidentWorkInfo -IncidentNumber INC123456
                Retrieves all Work Info  from the specified Incident and outputs to pipeline.

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
        [string]$IncidentNumber
    )

    $headers = Get-QFRemedyDefaultHeader
    $body = @{
        "Id"  = $IncidentNumber
    } | ConvertTo-Json
    
    try {
        $GetWorkInfoResponse = Invoke-RestMethod $GetIncidentWorkInfoURL -Method 'GET' -Headers $headers -Body $Body
        $CustomResponse = [PSCustomObject]@{
            Success = $true
            Result  = $GetWorkInfoResponse
        }
    }
    catch {
        $CustomResponse = [PSCustomObject]@{
            Success = $false
            Result  = $_.ErrorDetails.Message
        }
        Write-Error ((Get-LogPrefix) + "An error occured on Get-RemedyIncidentWorkInfo for incident $IncidentNumber")
        Write-Error ((Get-LogPrefix) + $_.Exception.Response.StatusCode.value__)
        Write-Error ((Get-LogPrefix) + $_.ErrorDetails.Message)
    }

    $CustomResponse
}



function Get-RemedyLogPrefix {
    <#
    .SYNOPSIS
        Returns a log prefix using datetime and the REQ + INC number

    .DESCRIPTION
        Returns a log prefix using datetime and the REQ + INC number
        Format: [dd/MM/yyyy hh:mm:ss] [REQ-INC]

    .INPUTS
        REQ
        INC

    .OUTPUTS   
        Returns the log prefix string
        Format: [dd/MM/yyyy hh:mm:ss] [REQ-INC]
                
    #>
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$REQ,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$INC

    )

    if ($null -eq $REQ -or $null -eq $INC) {
        "[$(Get-Date -Format "dd/MM/yyyy HH:mm:ss")] "
    }
    else {
        "[$(Get-Date -Format "dd/MM/yyyy HH:mm:ss")] [" + $REQ + "-" + $INC + "] "
    }
}