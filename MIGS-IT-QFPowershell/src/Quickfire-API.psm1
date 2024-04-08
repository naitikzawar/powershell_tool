###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                               API Functions                                 #
#                                   v1.6.2                                    #
#                                                                             #
###############################################################################

# Author: Chris Byrne - christopher.byrne@derivco.com.au


function Get-QFOktaToken {
    <#
    .SYNOPSIS
        Requests an OKTA Bearer Token, for use with EpicAPI.

    .DESCRIPTION
        This cmdlet requests an OKTA Bearer Token, for use with EpicAPI.

        This cmdlet requires the Epic.WebApi.Client.DLL file, which is included in the QFPowerShell repository under the 'lib' folder.
        Please visit https://epicapi-v4.mgsops.net/ for further information.

        This cmdlet also requires PowerShell Core or PowerShell 7 as this is a requirement for the DLL to function.

        This cmdlet will output an object with two members:
            Token - the Bearer token
            Expiry - the date and time (in local time) that the token will no longer be valid and a new token will need to be requested.

        Bearer tokens generally have a lifetime of two hours. This cmdlet will remember tokens generated in the current PowerShell session,
        and if the most recent token is still valid, it will return the existing token rather than requesting a new one.
        Specifying the 'Force' parameter, or the OktaCredentials parameter, will override this behaviour and a new token will be requested.

        To use a token with Invoke-RestMethod or Invoke-WebRequest, pass the .Token member of the output object as an Authorization header.
        For example, if you stored the output of this cmdlet into a $Token object, use the below value for the -Headers parameter:
        @{ Authorization = "Bearer " + $Token.Token }

        Note that the 'Bearer ' method name required for the HTTP Authorization header is not included in the token object,
        so you will need to add it as demonstrated above.

    .PARAMETER Force
        Requests a new OKTA Bearer Token, without checking if a valid token already exists.
        The default behaviour is to remember tokens as they are generated, and if this cmdlet is run again while the token is still valid,
        the existing token will be output to pipeline instead of requesting a new one.
        The Force parameter skips this check and will always generate a new token.

        Tokens are only remembered in the current session; closing down PowerShell will clear any remembered tokens.

    .PARAMETER OktaClientId
        The OKTA Client ID. A default value from Tech Ops team is provided. You should not need to adjust this unless the Client ID changes.

    .PARAMETER OktaCredentials
        A PSCredential object, for the account that you will use to request an OKTA bearer token.
        The specified account must be a member of the EPIC group in OKTA. Please reach out to Security team for assistance with this.

        If this parameter is not specified, the service account "ok-quickfireapiepic" will be used with a pre-configured password.

        To create a PSCredential Object use "Get-Credential".
        You can pipe the output of Get-Credential directly to this cmdlet or store it in an object, and pass the object as the value for this parameter.

        If this parameter is specified, a new token will be generated each time (effectively enabling the Force parameter).

    .PARAMETER OktaHost
        The address of the Okta API host. You should not need to adjust this unless the Okta host name changes.
        The default value is "derivco.okta-emea.com"

    .EXAMPLE
        $Token = Get-QFOktaToken

        Requests an OKTA Bearer token, and stores the output in the $Token object. This object has two members, Token and Expiry.
        To use a token with Invoke-RestMethod or Invoke-WebRequest, pass the .Token member of the output object as an Authorization header.
        For example, if you stored the output of this cmdlet into a $Token object, use the below value for the -Headers parameter:
        @{ Authorization = "Bearer " + $Token.Token }

        Note that the 'Bearer ' method name required for the HTTP Authorization header is not included in the token object,
        so you will need to add it as demonstrated above.

    .EXAMPLE
        $Token = Get-Credential | Get-QFOktaToken

        Requests an OKTA Bearer token, and stores the output in the $Token object.
        You will be prompted to enter a username and password for the account which will request the token.
        This account must be a member of the EPIC group in OKTA. Please reach out to Security team for assistance with this.

    .EXAMPLE
        $Token = Get-QFOktaToken -OktaCredential $Creds

        Requests an OKTA Bearer token, and stores the output in the $Token object.
        The $Creds object is a PSCredential object created with the Get-Credential cmdlet.
        e.g. "$Creds = Get-Credential"
        This account must be a member of the EPIC group in OKTA. Please reach out to Security team for assistance with this.

        By creating a Credential object in this manner, you can re-use the same credentials next time this cmdlet is run, without
        having to enter your username and password every time.


    .INPUTS
        This cmdlet accepts pipeline input for the OktaCredentials, OktaClientID and OktaHost parameters.

    .OUTPUTS
        An object consisting of an OKTA Bearer token string, and an expiry timestamp (in local time) will be output to pipeline.

            System.Management.Automation.PSCustomObject

                Name            MemberType      Definition
                ----            ----------      ----------
                Token           NoteProperty    string
                Expiry          NoteProperty    datetime

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://epicapi-v4.mgsops.net/

    #>

    # Set up parameters for this function
    [CmdletBinding()]
    [alias("okta")]
    param (

        # Force a new token, don't just give an already existing valid token
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [switch]$Force,

        # The username that you will use to request an OKTA Bearer Token. User must be a member of the EPIC group in OKTA
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PSCredential]$OktaCredentials,

        # The OKTA Client ID. We are just using the default Epic client ID.
        [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $OktaClientId = "0oa1l6jqqgsJbUHyu0i7",

        # The address of the Okta API host. You should only need to change this if the host name changes.
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $OktaHost = "derivco.okta-emea.com"
    )

    # Check we are using PowerShell 7
    If ($PSVersionTable.PSVersion.Major -lt 6) {
        Throw "This cmdlet requires PowerShell Core or PowerShell 7. Please install the latest version of PowerShell before running this cmdlet again."
    }

    # Try to import the EpicAPI Client DLL from the lib folder
    $EpicDLLPath = $($PSScriptRoot -replace "\\src$","\lib\").trim()
    Try {
        Add-Type -Path ($EpicDLLPath + "Epic.WebApi.Client.dll") -ErrorAction Stop
    } Catch {
        Write-Error "Unable to load the Epic.WebApi.Client.dll file. Please ensure the file exists at $($PSScriptRoot -replace "\\src$","\lib\Epic.WebApi.Client.dll")"
        Throw $_.Exception.Message
    }

    # The address of the Okta API host. Strip http/s if provided and any additional path after the host name
    $OktaHost = $OktaHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
    Write-Verbose ("[$(Get-Date)] Okta API address: $OktaHost")

    # progressPreference set to silently continue, so test-netconnection, invoke-webrequest etc doesn't show the progress bar. Big speedup in some cases
    $global:progressPreference = 'silentlyContinue'

    # Confirm we have connectivity to the Okta API host
    try {
        Test-NetConnection $OktaHost -port 443 -WarningAction Stop | Out-Null
    }
    catch {
        Throw "Unable to connect to the Okta API host: $OktaHost - Please ensure you have internet connectivity."
    }

    # If a PSCredential object was piped to this function, use those credentials otherwise use the default service account
    If ($null -ne $OktaCredentials) {
        $OktaUserName =  $OktaCredentials.UserName
        $OktaPassword = $OktaCredentials.Password
    } else {
        Write-Verbose ("[$(Get-Date)] Using Quickfire API Epic service account")
        $OktaUserName = "ok-quickfireapiepic"
        Invoke-Expression $([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(
        "W2J5dGVbXV0kSGFzaCA9IChHZXQtRmlsZWhhc2ggKCRFcGljRExMUGF0aCArICIubWV0YWRhdGEiKSB8IFNlbGVjdC1PYmplY" +
        "3QgLUV4cGFuZFByb3BlcnR5IEhhc2gpWzAuLjMxXQokT2t0YVBhc3N3b3JkID0gR2V0LUNvbnRlbnQgKCRFcGljRExMUGF0aC" +
        "ArICIuc3BlYyIpIC1FbmNvZGluZyBVVEY4IHwgQ29udmVydFRvLVNlY3VyZVN0cmluZyAtS2V5ICRIYXNo")
        )
    )
    }

    # How to change the saved password...
    #$OktaPassword = "PASSWORD"|ConvertTo-SecureString -AsPlainText -Force
    #[byte[]]$Hash = (Get-Filehash .\lib\.metadata|Select -ExpandProperty Hash)[0..31]
    #$OktaPassword|Convertfrom-SecureString -key $Hash |Out-File .\lib\.spec -Encoding utf8

    # Check if we already have a valid Token with 2 minutes until expiry.
    # If not, or Force or OKTACredentials parameters were set, request a new token
    If ($Force.IsPresent -or $null -ne $OktaCredentials -or $null -eq $script:QFScriptOktaToken.Token -or `
    ($script:QFScriptOktaToken.Expiry -lt $(Get-Date).AddMinutes(-2))) {
        # Call the OKTA API and request a token
        Write-Verbose ("[$(Get-Date)] Requesting a new OKTA Token...")
        Try {
            # Attempt to get a token 3 times.
            $i = 0
            do {
                $TokenHandler = new-object Epic.WebApi.Client.AccessTokenManager(("https://" + $OktaHost), $OktaUsername, ($OktaPassword|ConvertFrom-SecureString -AsPlainText), $OktaClientId)
                $Token = $TokenHandler.GetToken().GetAwaiter().GetResult()
                $i += 1
                if ($i -ge 3) {Throw "Could not retrieve a token after 3 attempts."}
            } until ($Null -ne $Token)
        } catch {
            Write-Error "Failed to retrieve an OKTA bearer token. Please ensure credentials are correct and the OKTA server is online."
            Throw $_.Exception.Message
        }
        # Bearer tokens expire in 2 hours
        $Expiry = $(Get-Date).AddMinutes(120)
        # Output the token and expiry to pipeline
        $script:QFScriptOktaToken = [pscustomobject]@{ Token = $Token ; Expiry = $Expiry}
        $script:QFScriptOktaToken
    } else {
        Write-Verbose ("[$(Get-Date)] Existing token still valid, will not request a new one.")
        $script:QFScriptOktaToken
    }
}


function Get-QFPortalToken {
    <#
    .SYNOPSIS
        Requests a Bearer Token for use with the Casino Portal API.

    .DESCRIPTION
        This cmdlet requests a Bearer Token for use with the Casino Portal API.

        This cmdlet will output an object with two members:
            Token - the Bearer token
            Expiry - the date and time (in local time) that the token will no longer be valid and a new token will need to be requested.

        Bearer tokens have a lifetime of one hour. This cmdlet will remember tokens generated in the current PowerShell session,
        and if the most recent token is still valid, it will return the existing token rather than requesting a new one.
        Specifying the 'Force' parameter will override this behaviour and a new token will be requested.

        To use a token with Invoke-RestMethod or Invoke-WebRequest, pass the .Token member of the output object as an Authorization header.
        For example, if you stored the output of this cmdlet into a $Token object, use the below value for the -Headers parameter:
        @{ Authorization = $Token.Token }

    .PARAMETER Force
        Requests a new Bearer Token, without checking if a valid token already exists.
        The default behaviour is to remember tokens as they are generated, and if this cmdlet is run again while the token is still valid,
        the existing token will be output to pipeline instead of requesting a new one.
        The Force parameter skips this check and will always generate a new token.

        Tokens are only remembered in the current session; closing down PowerShell will clear any remembered tokens.

    .PARAMETER PortalHost
        The address of the Casino Portal API host. You should not need to adjust this unless the host name changes.
        The default value is "casinoportal.gameassists.co.uk"

    .EXAMPLE
        $Token = Get-QFPortalToken

        Requests a Bearer token, and stores the output in the $Token object. This object has two members, Token and Expiry.
        The token will be remembered in the current PowerShell session, and if it is still valid, the existing token will be output to pipeline,
        instead of requesting a new one. If the token has expired a new one will be requested automatically.

        To use a token with Invoke-RestMethod or Invoke-WebRequest, pass the .Token member of the output object as an Authorization header.
        For example, if you stored the output of this cmdlet into a $Token object, use the below value for the -Headers parameter:
        @{ Authorization = $Token.Token }

    .EXAMPLE
        $Token = Get-QFPortalToken -Force

        Requests a Bearer token, and stores the output in the $Token object. This object has two members, Token and Expiry.
        The -Force parameter will generate a new Bearer token, regardless if a token has already been requested and is still valid.


    .INPUTS
        This cmdlet  does not accept pipeline input.

    .OUTPUTS
        An object consisting of a Casino Portal Bearer token string, and an expiry timestamp (in local time) will be output to pipeline.
            System.Management.Automation.PSCustomObject
                Name            MemberType      Definition
                ----            ----------      ----------
                Token           NoteProperty    string
                Expiry          NoteProperty    datetime

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://casinoportal.gameassists.co.uk/api/swagger/index.html

    #>

    # Set up parameters for this function
    [CmdletBinding()]
    param (

        # Force a new token, don't just give an already existing valid token
        [Parameter(Position = 0, ValueFromPipelineByPropertyName = $true)]
        [switch]$Force,

        # The address of the Casino Portal API host. You should only need to change this if the host name changes.
        [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $PortalHost = "casinoportal.gameassists.co.uk"
    )

    # The address of the Casino Portal API host. Strip http/s if provided and any additional path after the host name
    $PortalHost = $PortalHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""

    $KeyPath = $($PSScriptRoot -replace "\\src$","\lib\").trim()
    
    Invoke-Expression $([System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String('WwBiAHkAdABlAFsAXQBdACQASABhAHMAaAAgAD0AIAAoA' +
    'EcAZQB0AC0ARgBpAGwAZQBoAGEAcwBoACAAKAAkAEsAZQB5AFAAYQB0AGgAIAArACAAIgAuAG0AZQB0AGEAZABhAHQAYQAiACkAIAB8ACAAUwBlAGwAZQBjAHQALQBPAGIAagBlAGMAdAA' +
    'gAC0ARQB4AHAAYQBuAGQAUAByAG8AcABlAHIAdAB5ACAASABhAHMAaAApAFsAMAAuAC4AMwAxAF0ACgAkAFEARgBQAG8AcgB0AGEAbABLAGUAeQAgAD0AIABHAGUAdAAtAEMAbwBuAHQAZ' +
    'QBuAHQAIAAoACQASwBlAHkAUABhAHQAaAAgACsAIAAiAC4AZABhAHQAIgApACAALQBFAG4AYwBvAGQAaQBuAGcAIABVAFQARgA4ACAAfAAgAEMAbwBuAHYAZQByAHQAVABvAC0AUwBlAGM' +
    'AdQByAGUAUwB0AHIAaQBuAGcAIAAtAEsAZQB5ACAAJABIAGEAcwBoAA==')
        )
    )

    
    # How to change the saved key...
    #$Key = "KEY"|ConvertTo-SecureString -AsPlainText -Force
    #[byte[]]$Hash = (Get-Filehash .\lib\.metadata|Select -ExpandProperty Hash)[0..31]
    #$Key|Convertfrom-SecureString -key $Hash |Out-File .\lib\.dat -Encoding utf8

    if ($null -eq $QFPortalKey) {Throw "Unable to read API key, cannot continue."}

    # Check if we already have a valid Token with 2 minutes until expiry.
    # If not, or Force parameters were set, request a new token
    If ($Force.IsPresent -or $null -eq $script:QFScriptPortalToken.Token -or `
    ($script:QFScriptPortalToken.Expiry -lt $(Get-Date).AddMinutes(-2))) {
        # Call the Casino Portal API and request a token
        Write-Verbose ("[$(Get-Date)] Requesting a new Casino Portal Token...")
        Try {
            $Token = Invoke-RESTMethod -Uri ("https://" + $PortalHost + '/api/Security/' + $($QFPortalKey|ConvertFrom-SecureString -AsPlainText)) -Method POST
        } catch {
            Write-Error "Failed to retrieve a Casino Portal bearer token. Please ensure credentials are correct and the Casino Portal site is online and reachable."
            Throw $_.Exception.Message
        }
        # Bearer tokens expire in 1 hours
        $Expiry = $(Get-Date).AddMinutes(60)
        # Output the token and expiry to pipeline
        $script:QFScriptPortalToken = [pscustomobject]@{ Token = $Token ; Expiry = $Expiry}
        $script:QFScriptPortalToken
    } else {
        Write-Verbose ("[$(Get-Date)] Existing token still valid, will not request a new one.")
        $script:QFScriptPortalToken
    }
}




function Invoke-QFPortalRequest {
    <#
    .SYNOPSIS
        Retrieves information from the Casino Portal API.

    .DESCRIPTION
        This cmdlet retrieves information from the Casino Portal API.
        You can use the alias 'qfp' to run this cmdlet.

        Documentation for the Casino Portal API is available at https://casinoportal.gameassists.co.uk/api/swagger/index.html
        This cmdlet does not implement all functions of the API.

        Data returned from the API will be output to pipeline. If no data is returned from the API, e.g. a non-existent CasinoID was specified, there will be no pipeline output.

        Due to the large amount of data returned from the API, a number of properties are hidden from the console display by default.
        To show all properties, you can pipe the output through 'Format-List -Property *' - e.g. 
        Invoke-QFPortalRequest -CasinoID 12345 | Format-List -Property *

    .PARAMETER AllCasinos
        Returns the details of ALL QuickFire Casinos. No other parameters are required.

    .PARAMETER AllGames
        Returns the details of ALL QuickFire Games. No other parameters are required.

    .PARAMETER AllOperators
        Returns the details of ALL QuickFire Operators. No other parameters are required.

    .PARAMETER CasinoID
        A list of CasinoID's as integers, for which to request information from the Casino Portal API.
        A single CasinoID may be specified, or you can provide a comma seperated list of multiple CasinoID's to this parameter.
        You may also provide an object named CasinoID containing multiple CasinoID values via the pipeline.

    .PARAMETER CasinoName
        The name of a Casino to search for. Enclose the casino name in quotes if it contains spaces or special characters.
        Full details of any matching QuickFire Casinos will be returned.

    .PARAMETER CasinosForOperatorID
        A list of OperatorID's as integers, for which to request information from Casino Portal API.
        This mode will retrieve detailed information for all Casinos that are linked to the specified OperatorId.

        A single OperatorID may be specified, or you can provide a comma seperated list of multiple OperatorID's to this parameter.
        You may also provide an object named OperatorID containing multiple OperatorID values via the pipeline.

    .PARAMETER CID
        The Client ID of a game to search for. You may seperate multiple ClientID's with commas.
        Full details of any matching games will be returned.
        You may also combine with the MID parameter to further restrict the search to games with specific ModuleID's.

    .PARAMETER Currency
        Returns a list of supported Currencies and their multiplier values.
        Specifying an OperatorID parameter is optional. If you don't specify one, the list of currencies for 41662 (QF Showcase UAT) will be returned.
        This appears to include most, if not all, valid currencies.

    .PARAMETER GameName
        The name of a game to search for. Enclose the game name in quotes if it contains spaces or special characters.
        Full details of any matching games will be returned.

    .PARAMETER MID
        The Module ID of a game to search for. You may seperate multiple ModuleID's with commas.
        Full details of any matching games will be returned.
        YOu may also combine with the CID parameter to further restrict the search to games with specific ClientID's.

    .PARAMETER OperatorID
        A list of OperatorID's as integers, for which to request information from the Casino Portal API.
        A single OperatorID may be specified, or you can provide a comma seperated list of multiple OperatorID's to this parameter.
        You may also provide an object named OperatorID containing multiple OperatorID values via the pipeline.

        If the -Currency parameter is set, specifying an OperatorID parameter will return the list of supported Currencies for the specified OperatorID.

    .PARAMETER OpSecID
        A list of OperatorID's as integers, for which to retrieve credentials for the Operator Security site.
        A single OperatorID may be specified, or you can provide a comma seperated list of multiple OperatorID's to this parameter.
        You may also provide an object named OperatorID containing multiple OperatorID values via the pipeline.

        These credentials can be used at https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login which allows you to create and retrieve operator API keys.

        If you wish to retrieve UAT operator security credentials, you must also specify the -UAT parameter. Otherwise, only Production operator security credentials will be retrieved.

    .PARAMETER UAT
        When this parameter is specified along with the OpSecID parameter, the credentials retrieved will be for the UAT Operator Security site.

        These credentials can be used at https://operatorsecurityuat.valueactive.eu/system/operatorsecurityweb/v1/#/login which allows you to create and retrieve operator API keys for the UAT environment.

    .PARAMETER QFPortalHost
        The hostname of the Casino Portal server.
        You should not need to adjust this value unless the server name changes.
        The default value is "casinoportal.gameassists.co.uk"

    .PARAMETER WebsiteCasinoID
        A list of CasinoID's as integers, for which to retrieve website credentials.
        A single CasinoID may be specified, or you can provide a comma seperated list of multiple CasinoID's to this parameter.
        You may also provide an object named CasinoID containing multiple CasinoID values via the pipeline.

        These credentials can be used to login to the Casino's website for testing purposes.

    .EXAMPLE
        Invoke-QFPortalRequest -OperatorID 12345

        Retrieves information about the QuickFire Operator ID '12345' and outputs to pipeline.

    .EXAMPLE
        Invoke-QFPortalRequest -CasinoID 54321,98765,67890

        Retrieves information about the QuickFire Casino ID's 54321, 98765, and 67890 and outputs to pipeline.

    .EXAMPLE
        Invoke-QFPortalRequest -GameName "Reel Thunder"

        Retrieves information about any QuickFire games with a name containing the text "Reel Thunder".
            
    .EXAMPLE
        Invoke-QFPortalRequest -MID 10991

        Retrieves information about any QuickFire games with a ModuleID of 10991.

    .EXAMPLE
        Invoke-QFPortalRequest -MID 10991 -CID 50300

        Retrieves information about any QuickFire games with a ModuleID of 10991 and a ClientID of 50300.

    .EXAMPLE
        Invoke-QFPortalRequest -MID 10991,10992 -CID 50300,40300

        Retrieves information about any QuickFire games with a ModuleID of either 10991 or 10992 and a ClientID of either 50300 or 40300.

    .EXAMPLE
        Invoke-QFPortalRequest -Currency -OperatorId 54321

        Retrieves a list of supported currencies for the specified OperatorID (54321).

    .EXAMPLE
        $Games = Invoke-QFPortalRequest -AllGames

        Retrieves information about ALL QuickFire games and stores in an object called $Games
        Due to the size of the returned data, this may take some time.

        You can then get a list of game ModuleIDs, ClientIDs, and Game Names using a Select statement:
        $Games | Select ModuleId, ClientId, GameName

        You could also search for a specific ModuleID in the output like so:
        $Games | Select ModuleId, ClientId, GameName | Where {$_.ModuleId -eq '11020'}


    .INPUTS
        This cmdlet accepts pipeline input for the various parameters such as CasinoID or OperatorID.

    .OUTPUTS
        A PSCustomObject consisting of the output from the Casino Portal API.
        This output will vary depending on the parameters provided to this cmdlet.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://casinoportal.gameassists.co.uk/api/swagger/index.html

    #>

    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [alias("qfp")]
    param (

    # Returns details of ALL QuickFire casinos.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="AllCasinos")]
    [switch]$AllCasinos,

    # Returns details of ALL QuickFire games.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="AllGames")]
    [switch]$AllGames,

    # Returns details of ALL QuickFire Operators.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="AllOperators")]
    [switch]$AllOperators,

    # The CasinoIDs you wish to request information for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Casino")]
    [ValidateNotNullOrEmpty()]
    [int[]]$CasinoID,

    # The Casino Type for the API request. Defaults to 'GGL' for Quickfire casinos.
    [Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('All','Ggl','Traditional')]
    [string]$CasinoType = 'Ggl',

    # CID of games to search for
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="MIDCID")]
    [ValidateNotNullOrEmpty()]
    [int[]]$CID,

    # A game name to search for
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="CasinoName")]
    [ValidateNotNullOrEmpty()]
    [string]$CasinoName,

    # The OperatorIDs you wish to request detailed Casino information for - will show all Casinos linked to this Operator
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $True, ParameterSetName ="Currency")]
    [switch]$Currency,

    # A game name to search for
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="GameName")]
    [ValidateNotNullOrEmpty()]
    [String]$GameName,

    # The OperatorIDs you wish to request information for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Operator")]
    [Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Currency")]
    [ValidateNotNullOrEmpty()]
    [int[]]$OperatorID,

    # The OperatorIDs you wish to request detailed Casino information for - will show all Casinos linked to this Operator
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="CasinosForOperator")]
    [ValidateNotNullOrEmpty()]
    [int[]]$CasinosForOperatorID,

    # MID of games to search for
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="MIDCID")]
    [ValidateNotNullOrEmpty()]
    [int[]]$MID,

    # The OperatorIDs to retrieve Operator Security logins for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="OpSec")]
    [ValidateNotNullOrEmpty()]
    [int[]]$OpSecID,

    # Retrieves Operator Security logins for UAT if set. Default is to retrieve Production logins.
    [Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="OpSec")]
    [ValidateNotNullOrEmpty()]
    [switch]$UAT,

    # The CasinoIDs to retrieve website logins for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Website")]
    [ValidateNotNullOrEmpty()]
    [int[]]$WebsiteCasinoID,

    # The address of the Casino Portal host. You should only need to change this if the host name changes.
    [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    $QFPortalHost = "casinoportal.gameassists.co.uk"

    )
    Begin {

        # The address of the Casino Portal host.. Strip http/s if provided and any additional path after the host name
        $QFPortalHost = $QFPortalHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
        Write-Verbose ("[$(Get-Date)] Casino Portal address: $QFPortalHost")

        # Get the Portal Token
        $Token = (Get-QFPortalToken).Token
        If ($null -eq $Token) {Throw "Unable to retrieve a Casino Portal Authentication Token - cannot continue."}

        # Specify properties that are visible by default, different properties depending on the operation
        $CasinoVisible = 'productId','productName','OperatorId','OperatorName','GamingSystemName','HostingSiteId','marketType','Environment','BepType','GamingServerId','VanguardApiUrl','loginPrefix'
        $GameVisible = 'ModuleId','ClientId','GameName','uglGameId','externalGameName','gametype','provider','etiProductId','platform','rtp'

    }
    Process {

        Write-Verbose ("[$(Get-Date)] Requested Operation: $($PSCMDLet.ParameterSetName)")

        # Set up different variables for the request based on the Parameter Set
        switch ($PSCMDLet.ParameterSetName) {
            "AllCasinos" {
                $OperationIDs = ""
                $RequestPath = "/api/Casino/$CasinoType/"
                [string[]]$visible = $CasinoVisible
            }
            "AllGames" {
                $OperationIDs = ""
                $RequestPath = "/api/Games/List/"
                [string[]]$visible = $GameVisible
            }
            "AllOperators" {
                $OperationIDs = ""
                $RequestPath = "/api/Operator/$CasinoType/"
            }
            "Casino" {
                $OperationIDs = $CasinoID
                $RequestPath = "/api/Casino/$CasinoType/"
                [string[]]$visible = $CasinoVisible
            }
            "CasinoName" {
                $OperationIDs = [uri]::EscapeDataString($CasinoName)
                $RequestPath = "/api/Casino/$CasinoType/Name/Contains/"
                [string[]]$visible = $CasinoVisible
            }
            "CasinosForOperator" {
                $OperationIDs = $CasinosForOperatorID
                $RequestPath = "/api/Casino/$CasinoType/OperatorID/"
                [string[]]$visible = $CasinoVisible
            }
            "Currency" {
                If ($null -eq $OperatorID -or $OperatorID -le 0) {
                    $OperatorID = 41662
                }
                $OperationIDs = $OperatorID
                $RequestPath = "/api/BetSettings/"
                $QFEnv = "/Currencies"
            }
            "GameName" {
                $OperationIDs = [uri]::EscapeDataString($GameName)
                $RequestPath = "/api/Games/List/Name/"
                [string[]]$visible = $GameVisible
            }
            "MIDCID" {
                # This is a special case as there is no API method to lookup games via MID/CID.
                # We will get the full games list and filter it based on the MID/CID parameters.

                # To save time on repeated requests, store the full games list in a script scoped object and add a timestamp.
                # Script scoped means it will remain in memory after this function exits, until the powershell instance is closed.
                # Check for an existing script scoped object, and compare the timestamp - request a new list every 24 hours
                If (($null -eq $script:GamesList.DateRetrieved) -or `
                    ($script:GamesList.DateRetrieved | Select-Object -First 1) -lt (Get-date -Format FileDateTimeUniversal -Date (Get-Date).AddDays(-1))) {
                    # Request a new games list
                    Try {
                        Write-Verbose ("[$(Get-Date)] Request URI: $("https://" + $QFPortalHost + "/api/Games/List/")")
                        $script:GamesList = Invoke-RestMethod -URI ("https://" + $QFPortalHost + "/api/Games/List/") -Headers @{ Authorization = $Token }
                        # Add a timestamp of when the list was retrieved
                        $script:GamesList | Add-Member -MemberType NoteProperty -Name DateRetrieved -Value (Get-date -Format FileDateTimeUniversal)
                    } Catch {
                        Write-Error "Failed to retrieve games list from the Casino Portal API."
                        Throw $_.Exception.Message
                    }
                } else {
                    Write-Verbose ("[$(Get-Date)] Already have a games list object less than 24 hours old - won't request an updated list.")
                }
                $GameResults = @()
                # Now filter games list - first check if both MID and CID parameter set
                If ($null -ne $MID -and $null -ne $CID) {
                    $GameResultsMID = @()
                    Foreach ($ModuleID in $MID) {
                        $GameResultsMID += $script:GamesList | Where-Object {$_.ModuleId -eq $ModuleId}
                    }
                    Foreach ($ClientID in $CID) {
                        $GameResults += $GameResultsMID | Where-Object {$_.ClientId -eq $ClientId}
                    }
                } elseif ($null -ne $MID) {
                    # Only MID set
                    Foreach ($ModuleID in $MID) {
                        $GameResults += $script:GamesList | Where-Object {$_.ModuleId -eq $ModuleId}
                    }
                } elseif ($null -ne $CID) {
                    # Only CID set
                    Foreach ($ClientID in $CID) {
                        $GameResults += $script:GamesList | Where-Object {$_.ClientId -eq $ClientId}
                    }
                }
                # Output results to pipeline
                Write-Verbose ("[$(Get-Date)] " + $GameResults.Count + " games found matching specified criteria.")
                [Management.Automation.PSMemberInfo[]]$visibleProperties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet',[string[]]$GameVisible)
                $GameResults | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $visibleProperties -PassThru -Force
                
            }
            "Operator" {
                $OperationIDs = $OperatorID
                $RequestPath = "/api/Operator/$CasinoType/"
                
            }
            "OpSec" {
                $OperationIDs = $OpSecID
                $RequestPath = "/api/Passwords/$CasinoType/BackOffice/"
                If ($UAT) {
                    $QFEnv = "/UAT"
                } else {
                    $QFEnv = "/Production"
                }
            }
            "Website" {
                $OperationIDs = $WebsiteCasinoID
                $RequestPath = "/api/Passwords/$CasinoType/Website/"
            }
            Default {
                Throw "Unable to determine desired operation - please provide additional parameters for this request."
            }
        }

        Foreach ($OperationID in $OperationIDs) {
            Write-Verbose ("[$(Get-Date)] Request URI: $("https://" + $QFPortalHost + $RequestPath + $OperationID + $QFEnv)")
            Try {
                $Output = Invoke-RestMethod -URI ("https://" + $QFPortalHost + $RequestPath + $OperationID + $QFEnv) -Headers @{ Authorization = $Token }
            } Catch {
                Write-Error "Failed to retrieve requested information from the Casino Portal API."
                Throw $_.Exception.Message
            }

            # If Currency operation, expand the Available member for output to pipeline
            If ($PSCMDLet.ParameterSetName -eq "Currency" -and $null -ne $Output.Available) {
                $Output = $Output.Available
            }

            # if Casino object, add the loginPrefix property
            if ($Output.productId -gt 0 -and ($Output | Get-Member -MemberType NoteProperty -Name productSettings)) {
                $Output | ForEach-Object {
                    Add-Member -InputObject $_ -Name 'loginPrefix' -MemberType NoteProperty -Force -Value `
                    $($_.productSettings | Where-Object {$_.name -like "*Register - SGI JIT Account Creation Prefix*"} | Select-Object -expandProperty StringValue)
                }
            }

            # output object, set member properties to hidden if applicable
            If ($null -ne $visible) {
                [Management.Automation.PSMemberInfo[]]$visibleProperties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet',[string[]]$visible)
                $Output | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $visibleProperties -PassThru
            } else {
                $Output
            }
        }
    }
}



function Get-QFOperatorAPIKeys {
    <#
    .SYNOPSIS
        Retrieves an operator API key from the Operator Security website: https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login

    .DESCRIPTION
        This cmdlet will retrieve an operator API key from the Operator Security website. If there are no API keys created for the specified operator, this cmdlet will attempt to create one.

        It will first look up the Operator Security Credentials for the specified OperatorID using the Casino Portal API.
        If you specify a CasinoID parameter instead of an OperatorId, this cmdlet will first look up the OperatorID for that CasinoID using the Casino Portal API.
        This is done using the Invoke-QFPortalRequest function.

        Details regarding all available Operator API keys will be output to pipeline.
        Any deleted or disabled API keys will not be output to pipeline.

        If you specify a CasinoID parameter instead of an OperatorId, only API Keys for this Operator that are mapped for all products, plus this specific CasinoID will be output to pipeline.
        Any API Keys that are not mapped to the specified CasinoID will not be output to pipeline.

        If an API Key is configured for all products belonging to this OperatorID, the 'AllProducts' member in the output object will be set to True.
        The 'ProductIDs' member will be Null.

        If an API key is configured for a specific list of products belonging to this OperatorID, the 'AllProducts' member in the output object will be set to False.
        In addition, the 'ProductIDs' member will be an array of integer ProductIds that this API key is configured for.

    .PARAMETER CasinoID
        The CasinoID (aka ServerID, ProductID) that you wish to retrieve an API key for.

        If you specify a CasinoID parameter instead of an OperatorID, this cmdlet will look up the CasinoID's associated OperatorID automatically via the Casino Portal API.
        Only API Keys that are mapped for all products belonging to this Operator, plus this specific CasinoID will be output to pipeline.
        Any API Keys that are not mapped to the specified CasinoID will not be output to pipeline.

        You cannot specify both a CasinoID and an OperatorID at the same time.

        Note that API Keys for UAT and Production environments are not the same. 
        If you specify a UAT CasinoID, the API key returned will only work in the UAT environment, and will not function on the Production environment for the same OperatorID.

    .PARAMETER OperatorID
        The OperatorID that you wish to retrieve an API key for.
        All currently valid API Keys for the specified OperatorID will be output to pipeline.

        You cannot specify both a CasinoID and an OperatorID at the same time.
        
    .PARAMETER OpSecHost
        The hostname of the Operator Security Site server.
        You should not need to adjust this value unless the server name changes.

        For Production casinos, the default value is "operatorsecurity.valueactive.eu" 
        For UAT casinos, the value is "operatorsecurityuat.valueactive.eu"

        These default values will be used automatically for Production or UAT casinos respectively.
        If you specify this parameter, the provided hostname will be used for all casinos.

    .PARAMETER UAT
        If this switch is set, the cmdlet will return an API key for the UAT environment.
        API Keys for UAT and Production environments are not the same. 
        An API key for the UAT environment will not function on the Production environment for the same OperatorID, and vice versa.

    .EXAMPLE
        Get-QFOperatorAPIKeys -OperatorID 12345

        Retrieves API Keys for the specified Operator ID and outputs to pipeline.

    .EXAMPLE
        Get-QFOperatorAPIKeys -CasinoID 54321

        Retrieves API Keys for the specified CasinoID's associated Operator.
        Only API keys that are either configured for 'All Products' or specifically mapped to the specified CasinoID will be returned.


    .EXAMPLE
        Get-QFOperatorAPIKeys -OperatorID 12345 -UAT

        Retrieves API Keys for the UAT environment, for the specified Operator ID and outputs to pipeline.
        
        By default the API keys provided by this cmdlet are only valid for the Production environment. 
        Specifying the -UAT parameter will request keys for the UAT environment.
        An API key for the UAT environment will not function on the Production environment for the same OperatorID, and vice versa.
        

    .INPUTS
        This cmdlet accepts pipeline input for the OperatorID or CasinoID.

    .OUTPUTS
        A PSCustomObject consisting of the details of API keys from the Operator Security page.

            System.Management.Automation.PSCustomObject
                Name            MemberType      Definition
                ----            ----------      ----------
                AllProducts     NoteProperty    bool
                APIKey          NoteProperty    string
                Name            NoteProperty    string
                OperatorID      NoteProperty    long
                ProductIDs      NoteProperty    PSCustomObject

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login

    #>

    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = 'Operator')]
    [alias("qfk")]
    param (

    # The OperatorID you wish to request API keys for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Operator")]
    [ValidateNotNullOrEmpty()]
    [int]$OperatorID,

    # The CasinoID you wish to request API keys for.
    [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName ="Casino")]
    [ValidateNotNullOrEmpty()]
    [int]$CasinoID,

    # The address of the Operator Security host. You should only need to change this if the host name changes.
    [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OpSecHost = "operatorsecurity.valueactive.eu",

    # Request UAT API Keys
    [Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [switch]$UAT
    )

    Begin {
        # The address of the Operator Security host. Strip http/s if provided and any additional path after the host name
        $OpSecHost = $OpSecHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
        Write-Verbose ("[$(Get-Date)] Operator Security address: $OpSecHost")

    }

    Process {

        # If we were given a Casino ID, look up its OperatorID first
        If ($PSCmdlet.ParameterSetName -eq "Casino") {
            Write-Verbose ("[$(Get-Date)] Looking up the OperatorID for CasinoID $CasinoID...")
            Try {
                $CasinoData = Invoke-QFPortalRequest -CasinoID $CasinoID
            } Catch {
                Write-Error "Failed to retrieve Casino details from QFPortal API."
                Throw $_.Exception.Message
            }

            If ($null -eq $CasinoData) {Throw "No details found for CasinoID $CasinoID from Casino Portal. Please ensure you have entered the CasinoID correctly."}

            $OperatorID = $CasinoData.operatorID
            Write-Verbose ("[$(Get-Date)] CasinoID: $CasinoID CasinoName:  $($CasinoData.productName)")
            Write-Verbose ("[$(Get-Date)] OperatorID: $OperatorID OperatorName:  $($CasinoData.operatorName)")
            # Check for UAT Casino
            if ($CasinoData.Environment -eq "UAT") {$UAT = $true}
        }

        # Parameters for splatting to QFPortal request
        $OpsecParams = @{
            OpSecID = $OperatorID
        }

        # Check for UAT Casino
        If ($UAT) {
            Write-Verbose ("[$(Get-Date)] UAT mode, will request UAT environment API keys.")
            $OpsecParams.Add("UAT",$True)
            # Set the OpSecHost hostname, unless this was specified when cmdlet was run
            If (!($PSBoundParameters.ContainsKey('OpSecHost'))) {
                $OpSecHost = "operatorsecurityuat.valueactive.eu"
            }
        }

        Write-Verbose ("[$(Get-Date)] Requesting Operator Security Credentials from QFPortal API...")
        Try {
            $OpSec = Invoke-QFPortalRequest @OpsecParams
        } Catch {
            Write-Error "Failed to retrieve Operator Security Credentials from QFPortal API."
            Throw $_.Exception.Message
        }

        If ($null -eq $OpSec.username -or $null -eq $OpSec.password) {Throw "No credentials found for OperatorID $OperatorID from Casino Portal. Please ensure you have entered the OperatorID correctly."}

        Write-Verbose ("[$(Get-Date)] Requesting an Operator Security Bearer Token - Request URI: $("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/token")")
        Try {
            $OpSecToken = Invoke-RestMethod -URI ("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/token") -Method POST -Body `
            ("grant_type=password&username=" + $OpSec.username.trim() +  "&password=" + $OpSec.Password.trim()) -ErrorAction Stop
        } Catch {
            Write-Error "Failed to retrieve an Operator Security Bearer Token."
            Throw $_.Exception.Message
        }

        Write-Verbose ("[$(Get-Date)] Requesting list of Operator API Keys - Request URI: $("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys")")
        Try {
            $APIKeyData = Invoke-RestMethod -URI ("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys") -Headers @{ Authorization = "Bearer $($OPSecToken.access_token)" } -ErrorAction Stop
        } Catch {
            Write-Error "Failed to retrieve Operator API Keys."
            Throw $_.Exception.Message
        }

        Write-Verbose ("[$(Get-Date)] API keys retrieved: $(@($APIKeyData).Count)")
        # If we didn't get any API keys, OR none for all products and/or our specified CasinoID, try to create one for ALL products
        If (@($APIKeyData).Count -lt 1 `
        -or ($PSCmdlet.ParameterSetName -eq "Casino" -and (($APIKeyData.ProductIDs -eq $CasinoID).Count -eq 0 -and ($APIKeyData | Where-Object {$_.AllProducts -eq $true}).Count -eq 0)) `
        -or ($PSCmdlet.ParameterSetName -ne "Casino" -and ($APIKeyData | Where-Object {$_.AllProducts -eq $true}).Count -eq 0)) {
            # First we need to request a valid key GUID string
            Try {
                $NewAPIKeyGUID = Invoke-RestMethod -URI ("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys/generate") -Headers @{ Authorization = "Bearer $($OPSecToken.access_token)" } -ErrorAction Stop
            } Catch {
                Write-Error "No Operator API keys found, and failed to generate a new Operator API Key GUID."
                Throw $_.Exception.Message
            }

            Write-Verbose ("[$(Get-Date)] No All Products Operator API keys found, Attempting to create a new API Key for All Products using GUID: $NewAPIKeyGUID")

            # Create a hashtable for the body of the new API key request using the GUID we just got
            $NewAPIKeyBody = @{
                name = "DoNotDelete-QFCS" 
                apiKeyGuid = $NewAPIKeyGUID
                allProducts = $true
                products = @()
            }

            # make a POST request to create a new API key 
            Try {
                Invoke-RestMethod -URI ("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys") -Headers @{ Authorization = "Bearer $($OPSecToken.access_token)" } -Method `
                POST -Body $NewAPIKeyBody -ErrorAction Stop | Out-Null
            } Catch {
                Write-Error "No Operator API keys found, and failed to create a new operator API key."
                Throw $_.Exception.Message
            }

            # Now attempt to retrieve API keys again... this time one should be there!
            Write-Verbose ("[$(Get-Date)] Requesting list of Operator API Keys again - Request URI: $("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys")")
            Start-Sleep 2
            Try {
                $APIKeyData = Invoke-RestMethod -URI ("https://" + $OpSecHost + "/system/operatorsecurityweb/v1/api/apikeys") -Headers @{ Authorization = "Bearer $($OPSecToken.access_token)" } -ErrorAction Stop
            } Catch {
                Write-Error "Failed to retrieve Operator API Keys."
                Throw $_.Exception.Message
            }
        }

        # check and confirm that we now have a valid API key, and exit if not
        If (@($APIKeyData).Count -lt 1) {
            Write-Host "No API keys found for the specified OperatorID/CasinoID."
            Return}

        $APIKeys = @() # Output object for valid API keys

        # Get all keys that are set for AllProducts and not deleted
        $APIKeys += $APIKeyData | Where-Object {$_.IsDeleted -eq $false -and $_.AllProducts -eq $True} | Select-Object @{Name = 'APIKey'; Expression = {$_.APIKeyGuid}}, Name, OperatorID, AllProducts, @{Name = 'ProductIDs'; Expression = {$_.APIKeyProductMaps|Select-Object -ExpandProperty ProductID}}

        If ($PSCmdlet.ParameterSetName -eq "Casino") {
            # If function was run with a CasinoID parameter, look check for any keys that are mapped to that OperatorID
            $APIKeys += $APIKeyData | Where-Object {$_.IsDeleted -eq $false -and $_.AllProducts -eq $false -and $_.Apikeyproductmaps.ProductID -eq $CasinoID} | Select-Object @{Name = 'APIKey'; Expression = {$_.APIKeyGuid}}, Name, OperatorID, AllProducts, @{Name = 'ProductIDs'; Expression = {$_.APIKeyProductMaps|Select-Object -ExpandProperty ProductID}}
        } else {
            # If function was run with OperatorID parameter, look for any keys that are not deleted but not set to all products.
            $APIKeys += $APIKeyData | Where-Object {$_.IsDeleted -eq $false -and $_.AllProducts -eq $false} | Select-Object @{Name = 'APIKey'; Expression = {$_.APIKeyGuid}}, Name, OperatorID, AllProducts, @{Name = 'ProductIDs'; Expression = {$_.APIKeyProductMaps|Select-Object -ExpandProperty ProductID}}
        }

        # finally output all keys to pipeline
        $APIKeys
    }
}


function Get-QFOperatorToken {
    <#
    .SYNOPSIS
        Generates an operator API Token using the provided API Key.

    .DESCRIPTION
        Generates an operator API token using the provided API Key via the Operator Security API.

        By default this cmdlet will use the API endpoint operatorsecurity.valueactive.eu - this can be adjusted by changing the APIHost parameter.

        The token will be output to pipeline along with the timestamp of issue (UTC) and expiry duration in seconds, plus the token expiry timestamp in local time.
        The token value will be in the output member 'AccessToken'.

        This cmdlet will retain any tokens created during the current PowerShell session, and automatically re-use any tokens for the specified API key that are still within the validity period.
        It is possible to generate multiple Tokens for the same Operator using different API Keys.
        If you wish to override this behaviour and request a new Token from the Operator Security API, specify the 'Force' parameter.
        Any retained Tokens for the specified API Key will be purged and a new Token will be requested.

        API documentation is available at:
        https://reviewdocs.gameassists.co.uk/internal/document/System/Operator%20Security%20API/1/Resources/OperatorTokens/3EFA1721EA

    .PARAMETER APIKey
        The operator API Key that will be used to request a Token.
        API keys are in the format xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

    .PARAMETER APIHost
        The hostname of the API endpoint host server.
        You should not need to adjust this value unless the server name changes.
        The default value is "operatorsecurity.valueactive.eu"

        If you are requesting a token for a UAT Casino, you will need to set this parameter with the value "operatorsecurityuat.valueactive.eu"

    .PARAMETER Force
        Ignores any stored, previously generated Tokens for the specified API key, and forces a new Token to be generated.

    .EXAMPLE
        $OpToken = Get-QFOperatorToken -APIKey abcdefgh-ijkl-mnop-qrst-uvwxyz012345

        Generates an API Operator Token using the provided API Key.
        The generated Token value will be in the object $OpToken.AccessToken
        The generated Token will be retained, and if another Token is requested for the same API Key while the original Token is still valid,
        the original Token will be output to pipeline instead of requesting a new one.

    .EXAMPLE
        $OpToken = Get-QFOperatorToken -APIKey abcdefgh-ijkl-mnop-qrst-uvwxyz012345

        Generates an API Operator Token using the provided API Key.
        The generated Token value will be in the object $OpToken.AccessToken
        Any previously generated Tokens will be ignored and a new Token will be requested from the API.

    .INPUTS
        This cmdlet accepts pipeline input for the Operator API Key.

    .OUTPUTS
        A PSCustomObject with the response from the Operator Security API,
        containing the generated API Operator Token, API key used to request the Token, issue timestamp (in UTC), expiry timestamp (local time) and expiry duration in seconds.

            System.Management.Automation.PSCustomObject
                Name            MemberType      Definition
                ----            ----------      ----------
                APIKey          NoteProperty    string
                AccessToken     NoteProperty    string
                ExpiryInSeconds NoteProperty    long
                IssuedAtUTC     NoteProperty    datetime
                Expiry          NoteProperty    datetime

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://reviewdocs.gameassists.co.uk/internal/document/System/Operator%20Security%20API/1/Resources/OperatorTokens/3EFA1721EA

    #>
    # Set up parameters for this function
    [CmdletBinding()]
    [alias("qfot")]
    param (
    # The Operators API Key
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$APIKey,

    # The address of the API endpoint host. You should only need to change this if the host name changes.
    [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$APIHost = "operatorsecurity.valueactive.eu",

    # Ignore any previously stored tokens and generate a new one
    [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
    [switch]$Force
    )

    # The address of the API host. Strip http/s if provided and any additional path after the host name
    $APIHost = $APIHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
    Write-Verbose ("[$(Get-Date)] API Endpoint address: $APIHost")

    # Check we have a valid API Key
    if (!($APIKey -match '[a-zA-Z0-9]{8}-(?:[a-zA-Z0-9]{4}-){3}[a-zA-Z0-9]{12}')) {
        Throw "$APIKey does not appear to be a valid operator API Key."
    }

    # Array of previously requested tokens. Script scoped so it will be retained after function exits
    If ($null -eq $script:QFOperatorTokens) {$script:QFOperatorTokens = @()}
    Write-Verbose ("[$(Get-Date)] $($script:QFOperatorTokens.count) previously generated operator token(s)")

    # If Force parameter not set, see if we can reuse an existing token
    If ($Force.IsPresent) {
        Write-Verbose ("[$(Get-Date)] FORCE parameter set, will request a fresh token and ignore any existing tokens.")
        # Remove existing token from the array object
        [PSCustomObject[]]$script:QFOperatorTokens = $script:QFOperatorTokens | Where-Object {$_.APIKey -ne $APIKey.trim()}
    } else {
        If ($script:QFOperatorTokens | Where-Object {$_.APIKey -eq $APIKey.trim()}) {
            # Check token expiry - if less than 2 minutes until expiry, request a new one. Otherwise just return the existing one
            $OpToken = $script:QFOperatorTokens | Where-Object {$_.APIKey -eq $APIKey.trim()}
            If ($OpToken.Expiry -gt $(Get-Date).AddMinutes(2)) {
                Write-Verbose ("[$(Get-Date)] Found existing valid token for specified API key - will not request a new one.")
                Return $OpToken
            } else {
                # Remove the expired token from the array object
                [PSCustomObject[]]$script:QFOperatorTokens = $script:QFOperatorTokens | Where-Object {$_.APIKey -ne $APIKey.trim()}
            }
        }
    }

    # Request the Operator token from the API endpoint
    Try {
        Write-Verbose ("[$(Get-Date)] Requesting new API token...")
        $OpToken = Invoke-RESTMethod -Uri ("https://" + $APIHost + "/System/OperatorSecurity/v1/operatortokens") -Body `
        $(@{'APIKey' = $APIKey.trim()}|ConvertTo-JSON) -ContentType 'application/json' -Method Post
    } Catch {
        Write-Error "Failed to retrieve an operator API token."
        Throw $_.Exception.Message
    }

    # Get the token expiry timestamp in local time and add to the output object. Also add API key used to generate the token
    $OpToken | Add-Member -Name "Expiry" -MemberType NoteProperty -Value $(Get-Date ($OpToken.issuedatutc).ToLocalTime()).AddSeconds($OpToken.ExpiryInSeconds)
    $OpToken | Add-Member -Name "APIKey" -MemberType NoteProperty -Value $APIKey.trim()
    # Add the token to script scoped array and output the token to pipeline
    $script:QFOperatorTokens += $OpToken
    $OpToken
}


function Get-QFAudit {
<#
    .SYNOPSIS
        Retrieves transaction and financial audits from the Back Office Help Desk Express API.

    .DESCRIPTION
        Retrieves transaction and financial audits from the Back Office Help Desk Express API.

        You must provide an operator API Key, a player UserID and a CasinoID (aka ProductID/ServerID).
        You must also provide a HostingSiteID, this can be retrieved via Invoke-QFPortalRequest with a CasinoID parameter.

        By default this cmdlet will use the API endpoint api.valueactive.eu - this can be adjusted by changing the APIHost parameter.

        API documentation is available at:
        https://reviewdocs.gameassists.co.uk/internal/document/BackOffice/Help%20Desk%20Express%20API/1/Resources/FinancialAudits

    .PARAMETER APIHost
        The hostname of the API endpoint host server.
        You should not need to adjust this value unless the server name changes.
        The default value is "api.valueactive.eu"

    .PARAMETER CasinoID
        The CasinoID of the player you wish to generate an audit for.

    .PARAMETER EndDate
        The date and time of the most recent transactions to show in the output data.
        Transactions newer than the specified EndDate will be excluded from the audit.
        All dates and times must be in UTC time zone.

        You can use Get-Date to generate a valid date and time value for this parameter.
        e.g.
        -EndDate (Get-Date).AddDays(-30)
        Will set the EndDate parameter to 30 days ago.

        By default there is no end date set so all transactions between the StartDate, and the current time and date will be included in the audit.

    .PARAMETER FinancialAudit
        Generates a CasinoArchive Financial Audit.
        The default setting is TransactionAudit.

    .PARAMETER HostingSiteID
        The ID Number of the Hosting Site for the specified CasinoID.
        This can be retrieved via Invoke-QFPortalRequest with a CasinoID parameter.

        Quickfire Site ID's are:
            2   Malta (MAL)
            3   Canada (MIT)
            9   Gibralta (GIC)
           25   Croatia (CIL)
           29   IOA Staging Environment

        Note that there is no distinction between different systems at each site. i.e. MAL1, MAL2, and MAL3 systems all have the same HostingSiteID of 2.

    .PARAMETER ModuleID
        The ModuleID number to filter the transaction audit data.
        By default all transaction history for the specified player will be returned in the audit data. 
        This parameter allows you to request transaction audit data only for a specific ModuleID.
        This parameter is not available for Financial audits as the API doesn't support it.

    .PARAMETER NoCurrencyConversion
        By default, currency values are converted into whole numbers and displayed as cents in the output.
        This matches how the data is stored in the SQL databases.
        Setting this parameter shows currency values to two decimal places.

    .PARAMETER SortAscending
        Sorts the audit data by TransactionTime column in ascending order, so older transactions are output first.
        By default the data is sorted in descending order, so newer transactions are output first.

    .PARAMETER StartDate
        The date and time of the oldest transactions to show in the output data.
        Transactions older than the specified StartDate will be excluded from the audit.
        All dates and times must be in UTC time zone.

        You can use Get-Date to generate a valid date and time value for this parameter.
        e.g.
        -StartDate (Get-Date).AddDays(-30)
        Will set the StartDate parameter to 30 days ago.

        By default the StartDate parameter is set to 40 days ago, so all transactions from the current date and time up to 40 days old will be included in the audit.

    .PARAMETER Token
        The Operator Bearer Token that will be used to authenticate to the API endpoint.
        You can use Get-QFOperatorToken to retrieve this token using an API key.
        The token will be in the member 'AccessToken' of the object returned from Get-QFOperatorToken
        The token must be valid for the Operator that owns the specified CasinoID.

    .PARAMETER TransactionAudit
        Generates a VanguardState Transaction Audit.
        This is the default setting.

    .PARAMETER UserId
        The UserID number of the player you wish to generate an audit for.

    .EXAMPLE
        Get-QFAudit -Token $OpToken.accesstoken -HostingSiteID 3 -UserID 12345678 -CasinoID 98765

            Generates a Transaction Audit for PlayerID 12345678 on CasinoID 98765
            HostingSiteID is set to 3 for Malta.
            Token parameter is set to OpToken.accesstoken which was retrieved from Get-QFOperatorToken.

    .EXAMPLE
        Get-QFAudit -Token $OpToken.accesstoken -HostingSiteID 3 -UserID 12345678 -CasinoID 98765 -FinancialAudit

            Generates a Financial Audit for PlayerID 12345678 on CasinoID 98765
            HostingSiteID is set to 3 for Malta.
            Token parameter is set to OpToken.accesstoken which was retrieved from Get-QFOperatorToken.

    .EXAMPLE
        Get-QFAudit -Token $OpToken.accesstoken -HostingSiteID 3 -UserID 12345678 -CasinoID 98765 -FinancialAudit -NoCurrencyConversion

            Generates a Financial Audit for PlayerID 12345678 on CasinoID 98765
            Currency values will be displayed to 2 decimal places and not converted to whole numbers/cents.
            HostingSiteID is set to 3 for Malta.
            Token parameter is set to OpToken.accesstoken which was retrieved from Get-QFOperatorToken.

    .EXAMPLE
        Get-QFAudit -Token $OpToken.accesstoken -HostingSiteID 3 -UserID 12345678 -CasinoID 98765 -StartDate (get-date).AddDays(-180) -EndDate "2023-03-01"

            Generates a Transaction Audit for PlayerID 12345678 on CasinoID 98765
            Transactions from 6 months ago (180 days), up to 1st March 2023 will be included in the Audit.
            This example shows how you can specify an explicit date or use Get-Date to get a relative date, for StartData and EndDate parameters.

    .INPUTS
        This cmdlet accepts pipeline input for the Operator API Token.

    .OUTPUTS
        A PSCustomObject with the response from the Back Office Help Desk Express API.

        For Transaction Audits:
        System.Management.Automation.PSCustomObject
            Name                MemberType    Definition
            ----                ----------    ----------
            actionTime          NoteProperty    datetime
            amount              NoteProperty    double
            clientId            NoteProperty    long
            currencyCode        NoteProperty    string
            externalActionId    NoteProperty    long
            externalGameName    NoteProperty    string
            externalReference   NoteProperty    string
            moduleId            NoteProperty    long
            numberOfAttempts    NoteProperty    long
            productId           NoteProperty    long
            source              NoteProperty    string
            statusDescription   NoteProperty    string
            statusId            NoteProperty    long
            transactionNumber   NoteProperty    long
            transactionTime     NoteProperty    datetime
            transactionType     NoteProperty    string
            userId              NoteProperty    long
            userName            NoteProperty    string

        For Financial Audits:
        System.Management.Automation.PSCustomObject
            Name                MemberType    Definition
            ----                ----------    ----------
            changeAmount        NoteProperty  double
            clientId            NoteProperty  long
            event               NoteProperty  string
            finalBalance        NoteProperty  double
            gameName            NoteProperty  string
            moduleId            NoteProperty  long
            sessionId           NoteProperty  long
            transactionNumber   NoteProperty  long
            transactionTime     NoteProperty  datetime

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://reviewdocs.gameassists.co.uk/internal/document/BackOffice/Help%20Desk%20Express%20API/1/Resources/FinancialAudits

    #>
    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = "Transaction")]
    [alias("qfaudit")]
    param (
    # The Operators API Token
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    # Hosting Site ID
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$HostingSiteID,

    # Player's UserID
    [Parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$UserID,

    # Player's CasinoID
    [Parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$CasinoID,

    # Start Date to search transactions
    [Parameter(Position = 4, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [datetime]$StartDate = $((Get-Date).AddDays(-40)),

    # End Date to search transactions
    [Parameter(Position = 5, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [datetime]$EndDate = $(Get-Date),

    # Transaction audit mode switch (default)
    [Parameter(ParameterSetName = 'Transaction')]
    [switch]$TransactionAudit,

    # Financial audit mode switch
    [Parameter(ParameterSetName = 'Financial')]
    [switch]$FinancialAudit,

    # Sort Ascending switch, default is descending - sets the order of the TransactionTime field
    [Parameter()]
    [switch]$SortAscending,

    # ModuleID to filter transaction audits by (sets ModuleFilter object in the request body)
    [Parameter(ParameterSetName = 'Transaction')]
    [int]$ModuleID,

    # Don't convert currency values from 2 decimals to cents (whole numbers)
    [Parameter()]
    [switch]$NoCurrencyConversion,

    # The address of the API endpoint host. You should only need to change this if the host name changes.
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$APIHost = "api.valueactive.eu"
    )

    # The address of the API host. Strip http/s if provided and any additional path after the host name
    $APIHost = $APIHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
    Write-Verbose ("[$(Get-Date)] API Endpoint address: $APIHost")

    Write-Verbose ("[$(Get-Date)] Requesting a $($PSCMDLet.ParameterSetName) audit")

    # Build the request body
    $Body = @{
        UserId = $UserID
        ProductId = $CasinoID
        StartDate = $((Get-Date $StartDate).ToString("yyyy-MM-ddTHH:mm:ssK"))
        EndDate = $((Get-Date $EndDate).ToString("yyyy-MM-ddTHH:mm:ssK"))
        SortDirection = $(if ($SortAscending) {"1"} else {"2"})
    }

    # Optional ModuleFilter request parameter - Transaction Audit Only
    If ($ModuleID -gt 0) {
        $ModuleFilter = @{
            moduleIdMin = $ModuleID
            moduleIdMax = $ModuleID
        }
        $Body.Add("moduleFilter",$ModuleFilter)
    }

    # Convert Body hash table to JSON format
    $Body = $Body | ConvertTo-JSON

    Write-Verbose ("[$(Get-Date)] API Request Body: $Body")

    # Switch between financial or transaction audits
    If ($PSCMDLet.ParameterSetName -eq 'Financial') {
        $RequestPath = "/BackOffice/HelpDeskExpress/v1/financialAudits/financialDetails"
    } elseif ($PSCMDLet.ParameterSetName -eq 'Transaction') {
        $RequestPath = "/BackOffice/HelpDeskExpress/v1/financialAudits/transactionDetails"
    } else {
        Throw "You must specify either a Transaction audit or a Financial audit."
    }

    # Make the API request, try up to 3 times 
    $i = 1
    Do {
        Try {
            Write-Verbose ("[$(Get-Date)] Invoking API request, attempt $i")
            $AuditData = Invoke-RESTMethod -Uri ("https://" + "api" + $HostingSiteID + "." + $APIHost + $RequestPath) -Body `
            $Body -ContentType 'application/json' -Headers  @{ Authorization = "Bearer $Token" } -Method Post -ErrorAction Stop
            $i = 4
        } Catch {
            # The API can be a bit flaky, sometimes it will give an error code and work after retrying, check common error codes make the request again
            $StatusCode = $_.Exception.Response.StatusCode.value__
            If (($StatusCode -eq 404 -or $StatusCode -eq 500 -or $StatusCode -eq 502 -or $StatusCode -eq 503 -or $StatusCode -eq 504) -and $i -lt 3) {
                $i++
                Write-Verbose ("[$(Get-Date)] Received HTTP Status Code $StatusCode")
            } else {
                # something else went wrong, or we tried 3 times already - exit with an error
                Write-Verbose ("[$(Get-Date)] Failed to retrieve audit data from the API after $i attempts.")
                $errorDetails = $_.errordetails.message|Convertfrom-Json -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Details
                If ($null -ne $errorDetails -and $errorDetails -ne "") {
                    Write-Error $errorDetails
                }
                Throw $_.Exception.Message
            }
        }
    } until ($i -gt 3)
    
    # API returns monetary values in 2 decimal places, convert it to cents (whole numbers)
    If ($PSCMDLet.ParameterSetName -eq 'Transaction') {
        If (!($NoConvertCurrency.IsPresent)) {
            $AuditData.transactionAuditDetails | ForEach-Object {
                If ($_.amount -lt 0) {
                    $_.Amount = $_.Amount * -1
                    }
                $_.Amount = $_.Amount * 100
            }
        }
        # Output audit data to pipeline
        $AuditData.transactionAuditDetails
    } elseif ($PSCMDLet.ParameterSetName -eq 'Financial') {
        If (!($NoConvertCurrency.IsPresent)) {
            $AuditData.financialAuditDetails | ForEach-Object {
                $_.ChangeAmount = $_.ChangeAmount * 100
                $_.FinalBalance = $_.FinalBalance * 100
            }
        }
        # Output audit data to pipeline
        $AuditData.financialAuditDetails
    }
}


function Get-QFUser {
    <#
    .SYNOPSIS
        Checks that the specified Player Login exists on the specified CasinoID, and returns the matching UserID.

    .DESCRIPTION
        Checks that the specified Player Login exists on the specified CasinoID, and returns the matching UserID.

        You must provide an operator API Key, a player Login and a CasinoID (aka ProductID/ServerID).
        You must also provide a HostingSiteID, this can be retrieved via Invoke-QFPortalRequest with a CasinoID parameter.

        You must provide a player Login matching the exact value from the Casino database, otherwise the player will not be found.
        You may optionally provide the Casino Login Prefix (2 characters followed by an underscore) but this is not required.
        Wildcards are not supported due to a limitation of the Account API.

        By default this cmdlet will use the API endpoint api.valueactive.eu - this can be adjusted by changing the APIHost parameter.

        API documentation is available at:
        https://reviewdocs.gameassists.co.uk/internal/document/Account/Account%20API/1/Resources/Accounts/01C3E50E29

    .PARAMETER APIHost
        The hostname of the API endpoint host server.
        You should not need to adjust this value unless the server name changes.
        The default value is "api.valueactive.eu"

    .PARAMETER CasinoID
        The CasinoID of you wish to search in for the specified player.

    .PARAMETER HostingSiteID
        The ID Number of the Hosting Site for the specified CasinoID.
        This can be retrieved via Invoke-QFPortalRequest with a CasinoID parameter.

        Quickfire Site ID's are:
            2   Malta (MAL)
            3   Canada (MIT)
            9   Gibralta (GIC)
            25   Croatia (CIL)
            29   IOA Staging Environment

        Note that there is no distinction between different systems at each site. i.e. MAL1, MAL2, and MAL3 systems all have the same HostingSiteID of 2.

    .PARAMETER Login 
        The login of the player you wish to search for.
        Wildcards are not supported.
        You may optionally provide the Login Prefix from the CasinoDB however this is not required.


    .PARAMETER Token
        The Operator Bearer Token that will be used to authenticate to the API endpoint.
        You can use Get-QFOperatorToken to retrieve this token using an API key.
        The token will be in the member 'AccessToken' of the object returned from Get-QFOperatorToken
        The token must be valid for the Operator that owns the specified CasinoID.

    .EXAMPLE
        Get-QFUser -Token $OpToken.accesstoken -HostingSiteID 3 -Login GuyIncognito -CasinoID 98765

            Checks for the existence of a player with Login matching 'GuyIncognito' on MIT, for CasinoID 98765.
            Will return the UserID of a matching player if one is found.

    .INPUTS
        This cmdlet accepts pipeline input for the Operator API Token.

    .OUTPUTS
        A PSCustomObject with the UserID and CasinoID of any matching user.
        If no matching user was found, there will be no pipeline output.

        System.Management.Automation.PSCustomObject
            Name                MemberType    Definition
            ----                ----------    ----------
            casinoId            NoteProperty  int
            userId              NoteProperty  int
    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://reviewdocs.gameassists.co.uk/internal/document/Account/Account%20API/1/Resources/Accounts/01C3E50E29

    #>
    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = "Casino")]
    [alias("qfuser")]
    param (
    # The Operators API Token
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Token,

    # Hosting Site ID
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$HostingSiteID,

    # Player's Login
    [Parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Login,

    # Player's CasinoID
    [Parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$CasinoID,

    # The address of the API endpoint host. You should only need to change this if the host name changes.
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$APIHost = "api.valueactive.eu"
    )

    # The address of the API host. Strip http/s if provided and any additional path after the host name
    $APIHost = $APIHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
    Write-Verbose ("[$(Get-Date)] API Endpoint address: $APIHost")

    # Build the request body
    $Body = @{
        productId = $CasinoID
        username = $Login.trim()
    }

    # Convert Body hash table to JSON format
    $Body = $Body | ConvertTo-JSON

    Write-Verbose ("[$(Get-Date)] API Request Body: $Body")

    # Make the API request
    $Response = @()
    Try {
        $Response += Invoke-RESTMethod -Uri ("https://" + "api" + $HostingSiteID + "." + $APIHost + "/Account/v1/accounts/checkUserExists") -Body `
        $Body -ContentType 'application/json' -Headers  @{ Authorization = "Bearer $Token" } -Method Post -ErrorAction Stop
    } Catch {
        # Error handling
        $StatusCode = $_.Exception.Response.StatusCode.value__
        If ($StatusCode -eq 500) {
            Write-Error "Account API request failed - please check that you have specified the correct HostingSiteId; or there may be a server issue, please try again later. "
        } elseif ($StatusCode -eq 403) {
            Write-Error "Account API request failed - please check that you have a valid Operator Token and that it has permission for the specified CasinoId."
        } else {
            # something else went wrong,  exit with an error
            Write-Error "Account API request failed."
        }
        Throw $_
    }

    # If we got any results, write output to pipeline
    $Output = New-Object -TypeName pscustomobject
    $Output | Add-Member -MemberType NoteProperty -Name "CasinoId" -value $Response.RegisteredProductId
    $Output | Add-Member -MemberType NoteProperty -Name "UserId" -value $Response.UserId
    If ($null -ne $Output.userId) {$Output}
}


function Search-QFUser {
    <#
    .SYNOPSIS
        Searches multiple Casinos for a player with the specified Login.
        If a player was found, return the matching UserID plus details of the Casino where the player was located.

    .DESCRIPTION
        Searches multiple Casinos for a player with the specified Login.
        You must specify a player Login and another parameter to search with - either an OperatorID, CasinoID or Casino Name.

        You must provide a player Login matching the exact value from the Casino database, otherwise the player will not be found.
        You may optionally provide the Casino Login Prefix (2 characters followed by an underscore) but this is not required.
        Wildcards are not supported due to a limitation of the Account API.

        This cmdlet can take a long time to complete if there are a large number of casinos to search through.

        This cmdlet will automatically retrieve an Operator API Token for any OperatorID's found that match the specified search option. 
        An Operator API Key for 'All Products' must exist on the Operator Security site.

        By default this cmdlet will use the API endpoint api.valueactive.eu - this can be adjusted by changing the APIHost parameter.

        This cmdlet will exclude all UAT casinos from the search. UAT search may be added in the future if such a requirement arises.

    .PARAMETER APIHost
        The hostname of the API endpoint host server.
        You should not need to adjust this value unless the server name changes.
        The default value is "api.valueactive.eu"

    .PARAMETER CasinoID
        The CasinoID which you wish to search for the specified player Login.
        You can specify multiple CasinoID's seperated by commas.
        This parameter Cannot be used with OperatorID or CasinoName parameters.

    .PARAMETER CasinoName
        The name of a casino that you wish to search for the specified player Login.
        First a search will be performed for all casinos with a name matching the specified CasinoName, then each of these casinos will be searched for a player with a matching Login.
        This parameter Cannot be used with OperatorID or CasinoID parameters.

        Wildcards are not supported for this parameter; but the search will find any casino that contains the specified CasinoName anywhere in its name.

    .PARAMETER Login 
        The login of the player you wish to search for.
        Wildcards are not supported, the Login must match exactly
        You may optionally provide the Login Prefix from the CasinoDB however this is not required.

    .PARAMETER OperatorID
        The OperatorID which you wish to search in all linked Casinos for the specified player Login.
        You can specify multiple OperatorID's seperated by commas.
        This parameter Cannot be used with CasinoID or CasinoName parameters.
        
    .EXAMPLE
        Search-QFUser -Login GuyIncognito -CasinoName BigCasino

            Checks for the existence of a player with Login matching 'GuyIncognito' any Casinos with a name that contains 'BigCasino'
            Will return the UserID of a matching player if one is found on any of these Casinos.
            Will also return details about this casino if a player is located.

    .EXAMPLE
        Search-QFUser -Login GuyIncognito -CasinoID 12345

            Checks for the existence of a player with Login matching 'GuyIncognito' on CasinoID 12345.
            Will return the UserID of a matching player if one is found on the specified CasinoID.
            Will also return details about this casino if a player is located.
    
    .EXAMPLE
        Search-QFUser -Login GuyIncognito -CasinoID 12345,23456,34567

            Checks for the existence of a player with Login matching 'GuyIncognito' on CasinoID's 12345, 23456 and 34567
            Will return the UserID of a matching player if one is found on any of these CasinoIDs.
            Will also return details about the casino where the player was located.

    .EXAMPLE
        Search-QFUser -Login GuyIncognito -OperatorID 98765

            Checks for the existence of a player with Login matching 'GuyIncognito' for OperatorID 98765.
            Will return the UserID of a matching player if one is found on any of the Casinos that are linked to this OperatorID.
            Will also return details about the casino where the player was located.

    .EXAMPLE
        Search-QFUser -Login GuyIncognito -OperatorID 98765,87654,76543

            Checks for the existence of a player with Login matching 'GuyIncognito' on MIT, for OperatorIDs 98765, 87654 and 76543.
            Will return the UserID of a matching player if one is found on any of the Casinos that are linked to these OperatorIDs.
            Will also return details about the casino where the player was located.

    .INPUTS
        This cmdlet accepts pipeline input for the player Login and casino/operator search criteria.

    .OUTPUTS
        A PSCustomObject with the UserID and Casino Details of any matching user.
        If no matching user was found on any Casinos, there will be no pipeline output.

        System.Management.Automation.PSCustomObject
            Name                MemberType      Definition
            ----                ----------      ----------
            CasinoId            NoteProperty    int
            CasinoName          NoteProperty    string
            GamingSystem        NoteProperty    string
            GamingServerID      NoteProperty    int
            HostingSiteID       NoteProperty    int
            Market              NoteProperty    string
            UserId              NoteProperty    int
            OperatorId          NoteProperty    int
            LoginPrefix         NoteProperty    string

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        Account API - Check User Exists
        https://reviewdocs.gameassists.co.uk/internal/document/Account/Account%20API/1/Resources/Accounts/01C3E50E29
        
        Casino Portal - Casino Search, Operator Security Passwords
        https://casinoportal.gameassists.co.uk/api/swagger/index.html

        Operator Security API - Get Operator Tokens
        https://reviewdocs.gameassists.co.uk/internal/document/System/Operator%20Security%20API/1/Resources/OperatorTokens/3EFA1721EA

    #>
    
    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = "CasinoName")]
    [alias("qffind")]
    param (

    # Player's Login
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Login,

    # OperatorID
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true, ParameterSetName = "OperatorID")]
    [ValidateNotNullOrEmpty()]
    [int[]]$OperatorID,

    # Casino Name
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true, ParameterSetName = "CasinoName")]
    [ValidateNotNullOrEmpty()]
    [string]$CasinoName,

    # Casino ID
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true, ParameterSetName = "CasinoID")]
    [ValidateNotNullOrEmpty()]
    [int[]]$CasinoID,

    # The address of the API endpoint host. You should only need to change this if the host name changes.
    [Parameter(ValueFromPipelineByPropertyName = $true, Position = 2)]
    [ValidateNotNullOrEmpty()]
    [string]$APIHost = "api.valueactive.eu"
    )

    # The address of the API host. Strip http/s if provided and any additional path after the host name
    $APIHost = $APIHost.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
    Write-Verbose ("[$(Get-Date)] API Endpoint address: $APIHost")

    # Array for Casino Data retrieved from Casino Portal API
    $CasinoData = @()

    switch ($PSCMDLet.ParameterSetName) {
        "OperatorID" {
            # Look up all Casino's for these Operator Ids
            try {
                $CasinoData += Invoke-QFPortalRequest -CasinosForOperatorID $OperatorID
            } catch {
                Write-Error "Failed to retrieve list of Casinos for the specified Operator IDs."
                Write-Error $_
                return
            }
        }
        "CasinoID" {
            # Look up all Casino's for these Casino Ids
            try {
                $CasinoData += Invoke-QFPortalRequest -CasinoID $CasinoID
            } catch {
                Write-Error "Failed to retrieve list of Casinos for the specified Casino IDs."
                Write-Error $_
                return
            }
        }
        "CasinoName" {
                # Look up all Casino's matching specified Name
                try {
                    $CasinoData += Invoke-QFPortalRequest -CasinoName $CasinoName.trim()
                } catch {
                    Write-Error "Failed to retrieve list of Casinos matching the name '$CasinoName'"
                    Write-Error $_
                    return
                }
            }
        }

    # Pick out the relevant data into an array of hash tables, exclude any UAT casinos (HostingSiteID 29) or OpID < 1
    $Casinos = @()
    $CasinoData | ForEach-Object {
        If ($_.HostingSiteID -ne 29 -and $_.OperatorID -gt 1) {
            $CasinoTemp = @{
                CasinoId = $_.ProductId
                CasinoName = $_.ProductName
                HostingSiteID = $_.HostingSiteID
                GamingSystem = $_.GamingSystemName
                GamingServerID = $_.GamingServerID
                Market = $_.MarketType
                OperatorID = $_.OperatorID
                LoginPrefix = $(
                    [string]($_.productSettings | 
                    Where-Object {$_.name.trim() -like "Register - SGI JIT Account Creation Prefix"} |
                    Select-Object -expandProperty StringValue).trim()
                )
            }
            $Casinos += $CasinoTemp
        } elseif ($_.HostingSiteID -eq 29) {
            Write-Verbose ("[$(Get-Date)] CasinoId $($_.ProductId) is a UAT casino - skipping...")
        } elseif ($_.OperatorID -lt 1) {
            Write-Verbose ("[$(Get-Date)] CasinoId $($_.ProductId) is linked to OperatorID $($_.OperatorID) - skipping...")
        }
    }

    # Check if we got any Casino data
    If (@($Casinos).count -eq 0) {
        Write-Host "No Casinos found matching the specified search options."
        Return
    } 
    Write-Verbose ("[$(Get-Date)] Found $(@($Casinos).count) Casinos matching the specified search options.")
    
    # If we searched by OperatorID when cmdlet was run, then just request keys and tokens for those OpIDs, otherwise extract them from the retrieved casino data
    If ($null -eq $OperatorID) {
        $OperatorID = $Casinos.OperatorID | Sort-Object -Unique
        Write-Verbose ("[$(Get-Date)] Found $(@($OperatorID).count) OperatorIDs for these casinos, now attempting to retrieve API Keys")
    } else {
        $OperatorID = $OperatorID | Sort-Object -Unique
    }

    # Hashtable to store API Tokens
    $APITokens = @{}

    # Make the request for API Keys then generate tokens
    Foreach ($OpID in @($OperatorID)) {
        Try {
            $APIKey = $null
            # API key request
            $APIKey = (Get-QFOperatorAPIKeys -OperatorID $OpID | Where-Object {$_.AllProducts -eq $true} -ErrorAction Stop| Select-Object -First 1).APIKey
            # Check we actually got an API key
            If ($null -eq $APIKey) {
                Throw "No API key found - check that an API key has been generated for All Products in the Operator Security site."
            }
        } catch {
            Write-Error "Failed to retrieve an operator API Key for OperatorID $OpID - Check credentials for Operator Security Site are valid."
            Write-Error $_
            Continue
        }
        Write-Verbose ("[$(Get-Date)] OperatorID: $OpID  - API Key: $APIKey")
        try {
            # API token request using the key we just got
            $APIToken = (Get-QFOperatorToken -APIKey $APIKey -ErrorAction Stop).AccessToken
            # add the token to the hashtable for this OperatorID
            $APITokens.Add($OpId.tostring(),$APIToken)
        } catch {
        Write-Error "Failed to generate operator API Token for OperatorID $OpID"
        Write-Error $_
        }
    }

    # Now loop through all our casinos and run a player search
    $Output = @()
    Foreach ($Casino in $Casinos) {
        # check we have a token for this OperatorID, otherwise skip this casino
        If ($null -eq $APITokens["$($Casino.OperatorID)"]) {
            Write-Warning "Couldn't generate API token for CasinoID $($Casino.CasinoID) - $($Casino.CasinoName) - OperatorID - $($Casino.OperatorID) - skipping this casino... "
            Continue
        }
        Write-Verbose ("[$(Get-Date)] Searching for player $Login on $($Casino.CasinoID) - $($Casino.CasinoName) - $($Casino.GamingSystem)")
        Try {
            $PlayerData = Get-QFUser -Token $APITokens["$($Casino.OperatorID)"] -APIHost $APIHost -HostingSiteID $Casino.HostingSiteID -Login $Login -CasinoID $Casino.CasinoID
            If ($Null -ne $PlayerData) {
                # add results into Output array
                $Output += [PsCustomObject]@{
                    CasinoId = $PlayerData.CasinoId
                    UserId = $PlayerData.UserId
                    CasinoName = $Casino.CasinoName
                    GamingSystem = $Casino.GamingSystem
                    GamingServerId = $Casino.GamingServerId
                    Market = $Casino.Market
                    HostingSiteID = $Casino.HostingSiteID
                    OperatorID = $Casino.OperatorID
                    LoginPrefix = $Casino.LoginPrefix
                }


                Write-Verbose ("[$(Get-Date)] Found matching player with UserID $($PlayerData.UserId) on $($Casino.CasinoID)")
            } else {
                Write-Verbose ("[$(Get-Date)] No matching player found on $($Casino.CasinoID)")
            }
        } catch {
            $Exception = $_
            Write-Warning "An error occured searching for players on $($Casino.CasinoID) - $($Casino.CasinoName)"
            
            Try {
                Write-Warning $($Exception|ConvertFrom-JSON)
            } catch {
                Write-Warning $Exception
            }
        }
    }
    # Finally output results to pipeline
    Write-Verbose ("[$(Get-Date)] Found $($Output.count) matching players.")
    $Output
}


function Invoke-QFReconAPIRequest {
    <#
    .SYNOPSIS
        Invokes Reconciliation API functions such as managing Commit and Rollback queues for QuickFire operators.

    .DESCRIPTION
        Invokes Reconciliation API functions such as managing Commit and Rollback queues for QuickFire operators.
        An Operator API Token is required for authentication, and must be provided using the 'Token' parameter.

        Documentation for the Reconciliation API is available at https://reviewdocs.gameassists.co.uk/internal/document/ExternalOperators/Reconciliation%20API/1
        This cmdlet does not implement all functions of the API.

        Data returned from the API will be output to pipeline. If no data is returned from the API, e.g. a non-existent UserID was specified, there will be no pipeline output.

    .PARAMETER CasinoID
        A CasinoID (aka ServerID/ProductID) that the specified UserID belongs to.

    .PARAMETER HostingSiteID
        The ID Number of the Hosting Site for the specified CasinoID.
        This can be retrieved via Invoke-QFPortalRequest with a CasinoID parameter.

        Quickfire Site ID's are:
            2   Malta (MAL)
            3   Canada (MIT)
            9   Gibralta (GIC)
           25   Croatia (CIL)
           29   IOA Staging Environment

        Note that there is no distinction between different systems at each site. i.e. MAL1, MAL2, and MAL3 systems all have the same HostingSiteID of 2.

    .PARAMETER Reference
        The External Reference for Transaction Unlock requests.
        This will be recorded against the transaction in the VanguardState database when a transaction is successfully unlocked.
        A suggested Reference 

    .PARAMETER RoundInfo
        Specifies that a Game Round Info request should be made to Reconcilation API.
        Only one type of request parameter can be specified at a time.

        Various details about a game round will be output to pipeline. This includes the round status (Success, Rollback, Commit, etc), value of winnings, timestamps etc.
        This request does not modify or alter any game rounds on our system.

    .PARAMETER Token
        The Operator Bearer Token that will be used to authenticate to the API endpoint.
        You can use Get-QFOperatorToken to retrieve this token using an API key.
        The token will be in the member 'AccessToken' of the object returned from Get-QFOperatorToken
        The token must be valid for the Operator that owns the specified CasinoID.

    .PARAMETER TransactionIDs
        Specify a list of Game Round Transaction ID's. This is required for Game Round Info and Transaction Unlock requests.
        You can specify a single TransactionID, or multiple seperated by commas (without any spaces.)
        You can also specify a range of TransactionIDs using the Range operator syntax: (x..y)

    .PARAMETER QueueInfo
        Retrieves Queue details for the specified CasinoID and/or UserID.
        This includes a count of total transactions in each Queue, plus the details of each queued transaction.
        The Commit and Rollback queues are included, plus the PendingAdminEventQueue and the IncompleteGameQueue (pending endgames).
        Note that IncompleteGameQueue is NOT the same as game rounds open in the Casino Database.

        If a UserID is not specified, all queued transactions for all players belonging to the specified CasinoID will be returned.

        The API response will be output to pipeline as a nested PSCustomObject.

    .PARAMETER Unlock
        Specifies that a Transaction Unlock request should be made to Reconcilation API.
        Only one type of request parameter can be specified at a time.
        This request will commit/rollback transactions from the Commit or Rollback queues in the VanguardState database.

        Note that the operator must be informed when transactions are unlocked, as this may affect a player's balance.
        The operator may need to manually credit or refund the player.
        It is also recommended that a Reference parameter is specified, so unlocked transactions can be identified in the Back Office or VanguardState database.

        The results of the API request will be output to pipeline as a PSCustomObject, with one member for each TransactionID.
        If a transaction does not need to be unlocked - that is, it is not in the Commit or Rollback queue, this request will have no effect.
        The pipeline output will still show that the transaction was successfully unlocked, as the API responds like this if a transaction was unlocked or not.

    .PARAMETER UserId
        The UserID number of the player you wish to perform an operation on via Reconcilation API.

    .EXAMPLE
        Invoke-QFReconAPIRequest -Token $token -HostingSiteID 2 -CasinoID 12345 -UserID 654321 -TransactionIDs (100,101,102) -GameRoundDetails
            Retrieves Game Round Details for the specified UserID, CasinoID and TransactionIDs.

    .EXAMPLE
        Invoke-QFReconAPIRequest -Token $token -HostingSiteID 2 -CasinoID 12345 -UserID 654321 -TransactionIDs (100,101,102) -Unlock
            Unlocks the specified TransactionIDs. 
            This will commit or rollback any transactions from the Commit and Rollback queues in the VanguardState database.
            If the transaction is not in these queues, this will have no effect.

    .EXAMPLE 
        Invoke-QFReconAPIRequest -Token $token -HostingSiteID 2 -CasinoID 12345 -UserID 654321 -QueueInfo
            Retrieves queue information for the specified UserID, such as the details of queued transactions and the total count of transactions in each Queue.

    .INPUTS
        This cmdlet accepts pipeline input for the various parameters such as CasinoID or UserID,
        and a string value for the Token parameter.

    .OUTPUTS
        A PSCustomObject consisting of the output from the Reconciliation API.
        This output will vary depending on the parameters provided to this cmdlet.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://reviewdocs.gameassists.co.uk/internal/document/ExternalOperators/Reconciliation%20API/1

    #>

    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [alias("qfr")]
    param (
        # The Operators API Token
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Token,
    
        # Hosting Site ID - eg 2 for MAL
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$HostingSiteID,
    
        # Player's UserID
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true, Mandatory = $true, ParameterSetName = 'UnlockTransaction')]
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true, Mandatory = $true, ParameterSetName = 'RoundInfo')]
        [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true, Mandatory = $false, ParameterSetName = 'QueueInfo')]
        [ValidateNotNullOrEmpty()]
        [int]$UserID,
    
        # Player's CasinoID
        [Parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$CasinoID,

        # TransactionIDs for game round details/unlocks etc
        [Parameter(Position = 4, ValueFromPipelineByPropertyName = $true, Mandatory = $true, ParameterSetName = 'UnlockTransaction')]
        [Parameter(Position = 4, ValueFromPipelineByPropertyName = $true, Mandatory = $true, ParameterSetName = 'RoundInfo')]
        [ValidateNotNullOrEmpty()]
        [int[]]$TransactionIDs,

        # Specific game round details request
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'RoundInfo')]
        [ValidateNotNullOrEmpty()]
        [switch]$RoundInfo,

        # Unlock round request
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'UnlockTransaction')]
        [ValidateNotNullOrEmpty()]
        [switch]$Unlock,

        # Queue info request
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'QueueInfo')]
        [ValidateNotNullOrEmpty()]
        [switch]$QueueInfo,

        # External Reference for round unlock requests
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'UnlockTransaction')]
        [ValidateNotNullOrEmpty()]
        [string]$Reference,

        # The base address of the API host. The hosting site ID will be prepended to this, e.g. api2.$APIHost for MAL. You should only need to change this if the host name changes.
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $APIHost = "api.valueactive.eu"
    )


    # Get the URI of the request
    $RequestURI = "https://api$HostingSiteID.$APIHost/ExternalOperators/Reconciliation/v1/api".trim()

    # Switch statement controls the different API requests based on parameters provided to the cmdlet.
    switch ($psCmdlet.ParameterSetName) {
        RoundInfo {
            # Specific Game Round details
            If ($null -eq  $TransactionIDs) {Throw "No TransactionIDs specified to retrieve game round info."}
            # remove duplicates and process each remaining Transacton ID
            $TransactionIDs = $TransactionIDs | Sort-Object -Unique
            $TransactionDetails = @()
            foreach ($TransactionID in $TransactionIDs) {
                Write-Verbose ("[$(Get-Date)] Recon API request URI: $RequestURI/pendingtransactions/transaction/$TransactionID/user/$userId/product/$CasinoId")
                Try {
                    # Make the URI request and save response into an array
                    $TransactionDetails += Invoke-RestMethod -Uri "$RequestURI/pendingtransactions/transaction/$TransactionID/user/$userId/product/$CasinoId" -ErrorAction Stop `
                    -Method GET -ContentType "application/json" -Headers @{ Authorization = "Bearer $Token" }
                } Catch {
                    Write-Error $("An error occured while requesting Round Details for Transaction $TransactionID from Reconciliation API - " + $_.Exception.Message)
                    If (Test-Json $_.ErrorDetails.Message -ErrorAction SilentlyContinue) {
                        Write-Error ($_.ErrorDetails.Message | ConvertFrom-JSON).Details
                    }
                    Continue
                }
            }
            $TransactionDetails  
        }

        QueueInfo {
            # Gets details for queued transactions
            Write-Verbose ("[$(Get-Date)] Recon API request URI: $RequestURI/pendingtransactions/all")
            $Body = @{
                CommitQueue = @{
                    include = 'true'
                }
                RollbackQueue = @{
                    include = 'true'
                }
                IncompleteGameQueue = @{
                    include = 'true'
                }
                PendingAdminEventQueue = @{
                    include = 'true'
                }
                productIds = @($CasinoID) 
            } 
            If ($null -ne $UserID -and $UserID -gt 0) {
                $Body.Add('UserId',$UserID)
            }
            $Body = $Body | ConvertTo-JSON
            Try {
                # Make the URI request and save response into an array
                $QueueDetails += Invoke-RestMethod -Uri "$RequestURI/pendingtransactions/all" -ErrorAction Stop `
                -Method POST -ContentType "application/json" -Headers @{ Authorization = "Bearer $Token" } -Body $Body
            } Catch {
                Write-Error $("An error occured while requesting Queue Info from Reconciliation API - " + $_.Exception.Message)
                If (Test-Json $_.ErrorDetails.Message -ErrorAction SilentlyContinue) {
                    Write-Error ($_.ErrorDetails.Message | ConvertFrom-JSON).Details
                }
                Continue
            }
            $QueueDetails  
        }


        UnlockTransaction {
            # Unlock specific transactions
            If ($null -eq  $TransactionIDs) {Throw "No TransactionIDs specified to unlock."}
            # remove duplicates and process each remaining Transacton ID
            $TransactionIDs = $TransactionIDs | Sort-Object -Unique
            $Response = @()
            foreach ($TransactionID in $TransactionIDs) {
                $UnlockRequest = $null
                $Description = $null
                Write-Verbose ("[$(Get-Date)] Recon API request URI: $RequestURI/pendingtransactions/transaction/$TransactionID/user/$userId/product/$CasinoId/unlock")
                Try {
                    # Make the URI request
                    $UnlockRequest = Invoke-WebRequest -Uri "$RequestURI/pendingtransactions/transaction/$TransactionID/user/$userId/product/$CasinoId/unlock" -ErrorAction Stop `
                    -Method POST -ContentType "application/json" -Headers @{Authorization = "Bearer $Token"} -Body $(@{externalReference = "$Reference"}|ConvertTo-JSON)
                    # successful unlock will give HTTP 200 with description "OK"
                    If ($UnlockRequest.StatusCode -eq 200) {
                        $StatusCode = 200
                        $Description = "Transaction $TransactionId was successfully unlocked."
                    } else {
                        $Description = $UnlockRequest.StatusDescription
                    }
                } Catch {
                    # If an error occurs get the HTTP response code and the detailed error if one was supplied from the API
                    Write-Error $("An error occured while attempting to unlock Transaction $TransactionID via Reconciliation API - " + $_.Exception.Message)
                    $StatusCode = $_.Exception.Response.StatusCode.value__
                    If (Test-Json $_.ErrorDetails.Message -ErrorAction SilentlyContinue) {
                        $Description = ($_.ErrorDetails.Message | ConvertFrom-JSON).Details
                    } else {
                        $Description = $_.Exception.Message
                    }
                }
                $Response += [PSCustomObject]@{
                    TransactionID = $TransactionID
                    StatusCode = $StatusCode
                    Description = $Description
                }
            }
            $Response
        }

        Default {
            Throw "No Reconciliation API request parameter has been specified."
        }
    }
}

function Get-QFETIProviderInfo {
    <#
        .SYNOPSIS
            Displays support information for Quickfire ETI Providers.

        .DESCRIPTION
            Looks up support contact information for Quickfire ETI providers, from the file 'ETIProviders.csv'
            This CSV file must be present in the same folder as this PowerShell Module file.

            The information returned for each ETI provider includes a support email address, and a support portal URI and login credentials.
            If the provider does not have any of this information available in the CSV file, an empty value will be returned.

            The ETI provider information was taken from: https://confluence.derivco.co.za/pages/viewpage.action?pageId=360307215
            The information in the CSV file is not automatically updated. It may become out of date as ETI provider arrangements are changed.            

        .PARAMETER Id
            Specifies the ETI Provider ID number to search for. This can found in the Master Games List excel spreadsheet, downloaded from the Games Global Client Zone.

        .PARAMETER Name
            Specifies the Name of the ETI Provider ID to search for. The search is case-insensitive. 
            This will match if the specified Name parameter appears anywhere in an ETI provider's name.
            e.g. specifying 'gaming' for this parameter will match '1X2 Gaming', 'Gaming Corps', etc.

        .EXAMPLE
            Get-QFETIProviderInfo -Name Oryx
                Retrieves information regarding the ETI providers with a name matching 'Oryx'.

        .EXAMPLE
            Get-QFETIProviderInfo -Id 57
                Retrieves information regarding the ETI providers with an ID number of 57.

        .EXAMPLE
            Get-QFETIProviderInfo -Name *
                Outputs all information for all available ETI providers.

        .INPUTS
            This cmdlet accepts pipeline input for the Name or ID parameters.

        .NOTES
            This cmdlet reads its information from a CSV File which must be manually updated as ETI provider information changes.
            Eventually, we may seek to obtain an API key for Confluence and have it automatically look up the information directly from the Confluence page.
            It would then need to be formatted and converted into a PowerShell object.
            However, this is a long term, low priority goal.

        .OUTPUTS
            A PSCustomObject will be output to pipeline with the following members:
                System.Management.Automation.PSCustomObject
                    Name                MemberType      Definition
                    ----                ----------      ----------
                    Email               NoteProperty    string
                    ETIProvider         NoteProperty    string
                    ETIProviderId       NoteProperty    string
                    PortalPassword      NoteProperty    string
                    PortalURI           NoteProperty    string
                    PortalUsername      NoteProperty    string

        .LINK
            https://confluence.derivco.co.za/pages/viewpage.action?pageId=360307215
            
    #>
    
    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = 'Id')]
    [alias("qfeti")]
    param (
        [Parameter(Position=1,ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Id')]
        [ValidateNotNullOrEmpty()]
        [int]$Id,

        [Parameter(Position=1,ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Name')]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # Check that we have our CSV file with provider details and try to load it
    If (Test-Path "$PSScriptRoot\ETIProviders.csv" -PathType Leaf) {
        Try {
            Write-Verbose ("[$(Get-Date)] Reading ETI Providers list from $PSScriptRoot\ETIProviders.csv")
            $ETIProviders = Get-Content -Encoding UTF8 "$PSScriptRoot\ETIProviders.csv" -ErrorAction Stop | ConvertFrom-CSV -ErrorAction Stop
            # Check if we got any records from the CSV file
            If (@($ETIProviders).Count -lt 1) {Throw "The CSV file did not contain any records."}
            # Check if there is a property named SQLHost on the SQLDBList object
            If (!($ETIProviders|Get-Member -Name ETIProvider -MemberType Properties)) {Throw "The CSV file did not contain an ETIProviders column."}
            Write-Verbose ("[$(Get-Date)] Records retrieved from CSV file: $(@($ETIProviders).Count)")
        }
        Catch {
            $Exception = $_.Exception
            Throw "Failed to import ETI Providers list from the CSV file $PSScriptRoot\ETIProviders.csv - $Exception.Message"
        }
    } else {
        Throw "$PSScriptRoot\ETIProviders.csv file not found - cannot proceed."
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Id' {
            # Look up ETI provider ID number
            $ETIProviders | Where-Object {$_.ETIProviderId -eq $Id}
        }
        'Name' {
            # search the ETI provider name
            $ETIProviders | Where-Object {$_.ETIProvider -ilike "*" + $Name.trim() + "*"}
        }
        Default {
            # Shouldn't get here, function requires one of the above parameter sets to be active
            Throw "Unspecified action, please specify valid cmdlet parameters."
        }
    }
}


function Get-QFGameBlocking {
    <#
        .SYNOPSIS
            Displays game blocking information for the specified Casino, Country and Game.

        .DESCRIPTION
            Requests game blocking information from gw2.mgsops.net for the specified Casino, Country and Game.
            
            You may specify multiple CountryID, ModuleID and ClientID parameters.
            Game Blocking will be checked for each combination of provided values.

            A list of Countries is presented to the user via Out-GridView if the CountryId parameter is not specified.
            
            When Game Blocking details are retrieved, a brief summary of each game's blocking status is displayed to the user.
            The full details of game blocking are then output to pipeline.

            This cmdlet currently only supports checking game blocking for English language. 
            Please reach out to the author if you have a need for adding additional languages.

        .PARAMETER CasinoId
            Specifies the CasinoID/ProductID/ServerID to check for game blocking.

        .PARAMETER CID
            Specifies the ClientID of the game to check for game blocking. 
            You may specify multiple values for this parameter, and game blocking will be checked for each one.

        .PARAMETER CountryId
            Specifies the CountryId of the Country to check for game blocking.
            This is a 'ISO 3166-1' numeric country code, consiting of 1-3 digits.
            You may specify multiple values for this parameter, and game blocking will be checked for each one.

            If you do not specify this parameter, the cmdlet will look for a CSV file named 'Countries.CSV' which contains the CountryID and Country Name.
            It will then present a Grid View window for you to select the required Country.

            This parameter should be specified if you are calling this cmdlet from another function, as the Grid View will halt execution until the user makes a selection.

        .PARAMETER GamingSystemID
            The Gaming System ID of the Casino you wish to check for blocking.
            This value can be found using Invoke-QFPortalRequest and specifying the CasinoID of the required Casino. 
            The member 'gamingSystemID' contains this value.

            If this parameter is not specified, it will be looked up automatically.

        .PARAMETER MID
            Specifies the ModuleID of the game to check for game blocking.
            You may specify multiple values for this parameter, and game blocking will be checked for each one.

        .PARAMETER OktaToken
            An OKTA Bearer Token. You can retrieve this using the cmdlet 'Get-QFOktaToken'.
            This parameter is optional; if not supplied a new token will be requested automatically.

        .EXAMPLE
            Get-QFGameBlocking -CasinoId 39443 -MID 19964 -CID 50301
                Retrieves Game Blocking information for the specified Game MID/CID and Casino.
                A Grid-View is presented to the user to allow them to select the Countries they wish to check for blocking.

        .EXAMPLE
            Get-QFGameBlocking -CasinoId 39443 -MID 19964 -CID 50300 -CountryID 833
                Retrieves Game Blocking information for the specified Game MID/CID and Casino.
                The Country ID 833 is specified, corresponding to 'Isle of Man'.
                Specifying a CountryID parameter prevents the Grid-View from appearing.

                This is useful if you are calling this cmdlet from another function, as the Grid-View will halt execution until the user makes a selection.

        .EXAMPLE
            Get-QFGameBlocking -CasinoId 39443 -MID 19964,10976 -CID 40300,50300 -CountryID 833,380
                Retrieves Game Blocking information for the specified Game MID/CID's and Casino.
                
                The Country IDs 833 and 380 are specified, corresponding to 'Isle of Man' and 'Italy' respectively.
                Two Values for MID and CID are specified, so game blocking information will be retrieved for each combination of values.
                
                In this example, that is:
                ModuleID: 19964 ClientID: 40300 Country: Isle of Man
                ModuleID: 19964 ClientID: 50300 Country: Isle of Man
                ModuleID: 10976 ClientID: 40300 Country: Isle of Man
                ModuleID: 10976 ClientID: 40300 Country: Isle of Man
                ModuleID: 19964 ClientID: 40300 Country: Italy
                ModuleID: 19964 ClientID: 50300 Country: Italy
                ModuleID: 10976 ClientID: 40300 Country: Italy
                ModuleID: 10976 ClientID: 40300 Country: Italy

                The details of all eight game/country combinations will be retrieved and output to pipeline.

        .INPUTS
            This cmdlet accepts named pipeline input for all parameters.

        .NOTES
            Author:     Chris Byrne
            Email:      christopher.byrne@derivco.com.au

        .OUTPUTS
            A PSCustomObject will be output to pipeline with the following members:
                System.Management.Automation.PSCustomObject
                    Name            MemberType      Definition
                    ----            ----------      ----------
                    Blocks          NoteProperty    Object[]
                    CasinoName      NoteProperty    string
                    ClientID        NoteProperty    long
                    ClientName      NoteProperty    string
                    Country         NoteProperty    string
                    CountryID       NoteProperty    long
                    IsBlocked       NoteProperty    bool
                    LanguageID      NoteProperty    long
                    Market          NoteProperty    string
                    MarketTypeID    NoteProperty    long
                    ModuleID        NoteProperty    long
                    ServerID        NoteProperty    long
                    Territory       NoteProperty    string
                    TerritoryID     NoteProperty    long


        .LINK
            https://gw2.mgsops.net/GameBlocking/Detailed
            
    #>
    
    # Set up parameters for this function
    [CmdletBinding()]
    [alias("qfblock")]
    param (
        [Parameter(Position=0,ValueFromPipelineByPropertyName = $true,Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$CasinoId,

        [Parameter(Position=2,ValueFromPipelineByPropertyName = $true,Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [int[]]$CID,

        [Parameter(Position=3,ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int[]]$CountryID,

        [Parameter(Position=4,ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$GamingSystemId,

        [Parameter(Position=1,ValueFromPipelineByPropertyName = $true,Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [int[]]$MID,

        [Parameter(ValueFromPipelineByPropertyName = $true, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [string]$OktaToken
    )

    Begin {

        # If an Okta Token was not passed as a parameter, we'll try to request one. 
        if ($PSBoundParameters.ContainsKey('OktaToken')) {
            Write-Verbose ("[$(Get-Date)] OKTA token parameter supplied, will not request another.")
        } else {
            Write-Verbose ("[$(Get-Date)] No OKTA token parameter supplied, requesting/refreshing a token...")
            $OktaToken = (Get-QFOktaToken).Token               
        }    

        # If no CountryID parameter, Check that we have our CSV file with country details and try to load it
        If (!($PSBoundParameters.ContainsKey('CountryId')) -and (Test-Path "$PSScriptRoot\Countries.csv" -PathType Leaf)) {
            Try {
                Write-Verbose ("[$(Get-Date)] Reading Country list from $PSScriptRoot\Countries.csv")
                $CountryList = Get-Content -Encoding UTF8 "$PSScriptRoot\Countries.csv" -ErrorAction Stop | ConvertFrom-CSV -ErrorAction Stop
                # Check if we got any records from the CSV file
                If (@($CountryList).Count -lt 1) {Throw "The CSV file did not contain any records."}
                # Check if there is a property named CountryID on the CountryList object
                If (!($CountryList|Get-Member -Name CountryID -MemberType Properties)) {Throw "The CSV file did not contain a CountryID column."}
                Write-Verbose ("[$(Get-Date)] Records retrieved from CSV file: $(@($CountryList).Count)")
            }
            Catch {
                $Exception = $_.Exception
                Throw ("Failed to import Country list from the CSV file $PSScriptRoot\Countries.csv - " + $Exception.Message)
            }
            # Present Out-GridView for user to select Countries
            $CountryID = $CountryList | Out-GridView -OutputMode Multiple -Title "Please select Countries to check for game blocking" | Select-Object -ExpandProperty CountryId
            # If no countries were selected, exit
            if (@($CountryID).count -eq 0) {
                Throw "No countries selected, and no CountryID parameter specified. Unable to continue with game blocking check."
            }
            Write-Verbose ("[$(Get-Date)] CountryIDs selected from GridView: $CountryId ")
        } else {
            Write-Verbose ("[$(Get-Date)] Country ID parameter specified, will not load country list from CSV file. CountryID: $CountryId ")
        }

        # If no GamingSystemID parameter set, look it up
        If ($PSBoundParameters.ContainsKey('GamingSystemId')) {
            Write-Verbose ("[$(Get-Date)] Gaming System ID parameter specified, will not attempt to look it up. GamingSystemId: $GamingSystemId")
        } else {
            $GamingSystem = Invoke-RestMethod -URI "https://gw2.mgsops.net/GameBlocking/GetProductGamingSystems?term=$CasinoId" -Headers @{ Authorization = "Bearer $OktaToken" }
            # API will respond with ID column containing gamingsystemID and CasinoID in the 'Id' member, seperated by a semicolon.
            # 4 digit CasinoIds will return multiple results so match up the correct CasinoId tp the GS ID
            ForEach ($GSObject in $GamingSystem) {
                if (($GSObject.id -split ";")[1] -eq $CasinoID) {
                    [int]$GamingSystemId = ($GSObject.id -split ";")[0]
                    Break
                }
            }
            If ($GamingSystemId -le 0) {Throw "Unable to retrieve the Gaming System ID for the specified Casino. Cannot continue."}
            Write-Verbose ("[$(Get-Date)] GamingSystemId retrieved from gw2.mgsops.net : $GamingSystemId")
        }
    }

    Process {
        # Make a request to GW2 site looping through each country and game
        $OutputObject = @()
        Foreach ($Country in $CountryID) {
            foreach ($ModuleID in $MID) {
                Foreach ($ClientID in $CID)
                {
                    $RequestBody = @{ 
                        gamingSystemID = $GamingSystemId
                        productID = $CasinoId
                        countryID = $Country
                        moduleId = $ModuleID
                        clientId = $ClientID
                        languageId = 1 # English - ID's for languages seem to be arbitrary, just hard coding to English for now, may add a language option later
                    }

                    Write-Verbose ("[$(Get-Date)] Invoking REST method request to gw2.mgsops.net for game blocking info - Request Body:")
                    foreach($k in $RequestBody.Keys) { Write-Verbose "$k $($RequestBody[$k])" }
                    $GameBlock = Invoke-RestMethod -URI "https://gw2.mgsops.net/GameBlocking/GetLiveCasinoGameAvailability" -Headers @{Authorization = "Bearer $OktaToken"} -Body $RequestBody
                    If ($GameBlock.IsBlocked) {
                        Write-Host "$([char]27)[31mBLOCKED - Casino:$([char]27)[0m $($GameBlock.CasinoName) $([char]27)[31mGame:$([char]27)[0m $($GameBlock.ClientName) " -NoNewline
                        Write-Host "$([char]27)[31mCountry:$([char]27)[0m $($GameBlock.Country) $([char]27)[31mMID:$([char]27)[0m $($GameBlock.ModuleID) $([char]27)[31mCID:$([char]27)[0m $($GameBlock.ClientID)"
                        Write-Host "Found $(@($GameBlock.Blocks).Count) game blocking record(s):"
                        Foreach ($GameBlockDetail in $GameBlock.Blocks) {
                            $GameBlockDetail | Format-List -Property OBSNumber,BatchDescription,BlockSource | Out-Host
                        }
                    } else {
                        Write-Host "$([char]27)[32mAVAILABLE - Casino:$([char]27)[0m $($GameBlock.CasinoName) $([char]27)[32mGame:$([char]27)[0m $($GameBlock.ClientName) " -NoNewline
                        Write-Host "$([char]27)[32mCountry:$([char]27)[0m $($GameBlock.Country) $([char]27)[32mMID:$([char]27)[0m $($GameBlock.ModuleID) $([char]27)[32mCID:$([char]27)[0m $($GameBlock.ClientID)"
                    }
                    Write-Host ""
                    $OutputObject += $GameBlock
                }
            }
        }
    }
    End {
        $OutputObject
    }
}



function Start-QFGame {
    <#
        .SYNOPSIS
            Launches a Quickfire Game in the default web browser.

        .DESCRIPTION
            Launches a Quickfire Game in the default web browser via ServerID 21699 and OperatorId 47600 (Quickfire FakeAPI)
            
            You must specify either MID and CID parameters, or a UGL Launch Code, for the desired Game.

            A launch token is generated by making a request to gameshub.gameassists.co.uk - game launch will fail if this site is unreachable.

        .PARAMETER Balance
            Sets the player's account balance. 
            By default, this will be set to 10,000.00 in 1:1 currencies such as USD, GBP, EUR etc, or equivalent for other currencies.
            e.g. if you are launching a game with a 5x currency, the player's balance will be set to 50,000.00
            A seperate account balance is maintained for each different currency (as these are all differnt UserID's in the database).
            The balance will be reset on every game launch.

            This parameter allows you to specify a value for the player's account balance. Note that this will not be adjusted for currency multipliers.
            e.g. if you set this parameter to 5000 then launch a game the account balance will be 50.00 regardless of currency.
            
        .PARAMETER CID
            Specifies the ClientID of the game to launch.
            If this is specified, you must also specify a ModuleID.
            The UGL launch code for the desired Game will be retrieved via Invoke-QFPortalRequest.

            If this parameter is not specified, it will default to 50300.

        .PARAMETER Currency
            Specifies the Currency you wish the game to launch with. This is a 3-digit ISO currency code.
            If you do not specify this parameter, it will default to Euro (Currency code 'EUR').

            Not all currencies are supported by Games Hub. 
            An unsupported currency will give a 404 error while trying to update the player balance.

        .PARAMETER Language
            The Language code for the game launch. 
            This is optional; if not specified, will default to 'EN' for English.

            This is not validated, if an invalid language code is specified, most games will just launch in English.

        .PARAMETER LaunchCode
            The UGL Launch Code for the desired Game.
        
        .PARAMETER MID
            Specifies the ModuleID of the game to launch.
            If this is specified you must also specify a ClientID.
            The UGL launch code for the desired Game will be retrieved via Invoke-QFPortalRequest.

        .PARAMETER NoConfirm
            Launches the game without confirming.
            By default, the user will be prompted to press ENTER to launch a game or press C to copy the Launch URI to clipboard.
            Any other input will exit the cmdlet without further action.

            If this parameter is specified, the game will be launched in the default browser without confirmation.

        .PARAMETER OktaToken
            An OKTA Bearer Token. You can retrieve this using the cmdlet 'Get-QFOktaToken'.
            This parameter is optional; if not supplied, a new token will be requested automatically.

        .PARAMETER QFGames
            Uses the old game launch method, via qfgames.gameassists.co.uk
            The games will be launched using CasinoID 18226 and OperatorId 41662 (Quickfire Showcase UAT).
            The Games Hub site, used by default, is the newer replacement for the QFGames site.
            
            This parameter will stop working at some point when the QFGames site is shut down.

            Note that the Balance, CasinoID and Showcase parameters have no effect if this parameter is specified.

        .PARAMETER ServerID
            Specifies the ServerID (aka CasinoID or ProductID) used to launch the game.
            Games Hub supports several different ServerID's across different sites and markets.

            Please refer to the Games Hub website to see the supported ServerID's.
            Adjusting the Site and Market drop-down menus will change the displayed ServerID.

            Specifying a ServerID that is not supported by Games Hub (including a ServerID for any live Casino) will fail.

        .PARAMETER Showcase
            Games will be launched via Games Hub using ServerID 18226 and OperatorId 41662 (Quickfire Showcase UAT).
            The default behaviour, when this parameter is not specified, is to launch games from ServerID 21699 and OperatorId 47600 (Quickfire FakeAPI).
            This allows you to launch games using an alternative Casino if the default FakeAPI Casino is not working.

            This is effectively the same as setting the 'ServerID' parameter to 18226 and is just provided for convenience. 

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop
                Launches the game '9 pots of Gold - Desktop' in the default browser.
                
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
                The currency used in the game will be Euros. The player's account balance will be set to 10,000.00 EUR

        .EXAMPLE
            Start-QFGame -MID 19964
                Launches the game with ModuleID 19964 and ClientID 50300 (9 Masks of Fire - Desktop) in the default browser.
                Since the CID parameter was not specified, it defaults to 50300.
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
               The currency used in the game will be Euros. The player's account balance will be set to 10,000.00 EUR

        .EXAMPLE
            Start-QFGame -MID 19964 -CID 50301
                Launches the game with ModuleID 19964 and ClientID 50301 (9 Pots of Gold - Desktop) in the default browser.
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
                The currency used in the game will be Euros. The player's account balance will be set to 10,000.00 EUR

        .EXAMPLE
            Start-QFGame -MID 19964 -CID 50301 -Currency GBP
                Launches the game with ModuleID 19964 and ClientID 50301 (9 Pots of Gold - Desktop) in the default browser.
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
                The currency used in the game will be British Pounds (GBP). The player's account balance will be set to 10,000.00 GBP

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop -Currency GBP
                Launches the game '9 Pots of Gold - Desktop' in the default browser.
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
                The currency used in the game will be British Pounds (GBP). The player's account balance will be set to 10,000.00 GBP
                

        .EXAMPLE
            Start-QFGame -MID 19964 -CID 50301 -Currency GBP -NoConfirm
                Launches the game with ModuleID 19964 and ClientID 50301 (9 Pots of Gold - Desktop) in the default browser.
                The currency used in the game will be British Pounds (GBP). The player's account balance will be set to 10,000.00 GBP
                The game will be launched without asking the user for confirmation.

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop -Currency GBP -NoConfirm
                Launches the game '9 Pots of Gold - Desktop' in the default browser.
                The currency used in the game will be British Pounds (GBP). The player's account balance will be set to 10,000.00 GBP
                The game will be launched without asking the user for confirmation.

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop -Currency EUR -Language ES
                Launches the game '9 Pots of Gold - Desktop' in the default browser.
                The currency used in the game will be Euros. The player's account balance will be set to 10,000.00 EUR
                The game will be in Spanish language (ES).
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop -Balance 1000 -Currency SEK
                Launches the game '9 Pots of Gold - Desktop' in the default browser. 
                The currency used in the game will be Swedish Krona (SEK). The player's account balance will be set to 10.00 SEK.
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.

        .EXAMPLE
            Start-QFGame -LaunchCode 9potsOfGoldDesktop -ServerID 2000
                Launches the game '9 Pots of Gold - Desktop' in the default browser. 
                The currency used in the game will be Euros. The player's account balance will be set to 10,000.00 EUR
                The user will be prompted to hit ENTER to launch the game, or C to copy the Game Launch URI to clipboard.
                The game will be launched on ServerID/CasinoID 2000 (GIC Showcase) - OperatorID 41796

        .INPUTS
            This cmdlet accepts named pipeline input for all parameters.

        .NOTES
            Author:     Chris Byrne
            Email:      christopher.byrne@derivco.com.au

        .OUTPUTS
            This cmdlet produces no pipeline output.


        .LINK
            https://qfgames.gameassists.co.uk/
            
    #>
    
    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName='UGL')]
    [alias("qflaunch")]
    param (

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$Balance,

        [Parameter(Position=1,ValueFromPipelineByPropertyName = $true,ParameterSetName = 'MIDCID')]
        [ValidateNotNullOrEmpty()]
        [int]$CID=50300,

        [Parameter(Position=2,ValueFromPipelineByPropertyName = $true,ParameterSetName = 'MIDCID')]
        [Parameter(Position=1,ValueFromPipelineByPropertyName = $true,ParameterSetName = 'UGL')]
        [ValidateNotNullOrEmpty()]
        [string]$Currency='EUR',

        [Parameter(Position=4,ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Language='EN',

        [Parameter(Position=0,ValueFromPipelineByPropertyName = $true,ParameterSetName = 'UGL')]
        [ValidateNotNullOrEmpty()]
        [string]$LaunchCode,

        [Parameter(Position=0,ValueFromPipelineByPropertyName = $true,Mandatory = $true,ParameterSetName = 'MIDCID')]
        [ValidateNotNullOrEmpty()]
        [int]$MID,

        [Parameter(Position=3,ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [switch]$NoConfirm,

        [Parameter(ValueFromPipelineByPropertyName = $true, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [string]$OktaToken,

        [Parameter()]
        [switch]$QFGames,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$ServerID=21699,

        [Parameter()]
        [switch]$Showcase
    )

    If ($Showcase -eq $True) {
        # Force showcase launch instead of provided ServerID parameter
        $ServerId = 18226
    }

    # Find the UGL Launch Code if MID/CID specified
    If ($PSCmdlet.ParameterSetName -eq 'MIDCID') {
        Write-Verbose ("[$(Get-Date)] Requesting UGL Launch code for MID $MID and CID $CID")
        $GameInfo = Invoke-QFPortalRequest -MID $MID -CID $CID
        If ($null -eq $GameInfo.uglGameId -or $GameInfo.uglGameId -eq "") {
            Throw "Unable to retrieve a UGL Launch Code from Casino Portal for the specified game MID/CID."
        }
        $LaunchCode = $GameInfo.uglGameId
    }


    if ($GameInfo.gameName -ne "" -and $null -ne $GameInfo.gameName) {
        Write-Host ("$([char]27)[36mLaunching: $([char]27)[0m" + $GameInfo.gameName + " ") -NoNewline
    }
    Write-Host ("$([char]27)[36mUGL Launch code: $([char]27)[0m" + $LaunchCode)

    # Try to get lookup the casino name from the ServerID
    $CasinoInfo = Invoke-QFPortalRequest -CasinoID $ServerId -ErrorAction SilentlyContinue
    If ($null -ne $CasinoInfo.productId -and $CasinoInfo.productId -ne 0) {
        Write-Host ("$([char]27)[36mServerID: $([char]27)[0m" + $ServerId + " - " + $CasinoInfo.productName `
        + " $([char]27)[36mOperatorID: $([char]27)[0m" + $CasinoInfo.operatorId + " $([char]27)[36mSite: $([char]27)[0m" + $CasinoInfo.GamingSystemName)
    } else {
        Write-Warning "Unable to validate the specified ServerID: $ServerId"
    }
 
    # Check for valid currency and display info
    $Currency = $Currency.ToUpper().Trim()
    $CurrencyList = Invoke-QFPortalRequest -Currency -ErrorAction Continue
    $CurrencyInfo = $CurrencyList | Where-Object {$_.ISOCode -eq $Currency}
    If ($null -eq $CurrencyInfo) {
        Write-Warning "Unable to validate the specified currency."
    } else {
        [PSCustomObject]@{
            Currency = $CurrencyInfo.isoName
            CurrencyID = $CurrencyInfo.currencyID
            ISOCode = $CurrencyInfo.isoCode
            Multiplier = $CurrencyInfo.multiplierMaxBet
        } | Format-List
    }

    If ($QFGames) {
        # Old QFGames launch method
        $TokenURI = "https://qfgames.gameassists.co.uk/QuickfireGamesAPI/Showcase/GetToken?loginname=" + $env:UserName + $Currency + "&currency=" + $Currency
        Write-Verbose ("[$(Get-Date)] Requesting launch token from QFGames portal - URI:  $TokenURI")
        $LaunchToken = Invoke-RestMethod $TokenURI
        $LaunchURI = "https://gamelauncheruat.gameassists.co.uk/launcher/Generic?casinoid=18226&gameName=" + $LaunchCode.trim() + "&authToken=" + $LaunchToken + "&languageCode=" + $Language.trim()
    } else {
        # New Gameshub launch method
        # If an Okta Token was not passed as a parameter, we'll try to request one. 
        if ($PSBoundParameters.ContainsKey('OktaToken')) {
            Write-Verbose ("[$(Get-Date)] OKTA token parameter supplied, will not request another.")
        } else {
            Write-Verbose ("[$(Get-Date)] No OKTA token parameter supplied, requesting/refreshing a token...")
            $OktaToken = (Get-QFOktaToken).Token
        } 

        # Get the UserID using OKTA auth
        $HubUserID = Invoke-RESTMethod -Uri 'https://gameshub.gameassists.co.uk/api/User/login' -Method POST -ContentType 'text/plain' -Headers @{ Authorization = "Bearer $OktaToken" } `
        | Select-Object -ExpandProperty userId -ErrorAction Stop
        Write-Verbose ("[$(Get-Date)] Games Hub UserID: $HubUserID")

        
        If ($Showcase -eq $True) {
            # Force showcase launch instead of provided ServerID parameter
            $ServerId = 18226
        }

        # Update the balance for this player; different playerID for each currency
        # First check the wallet for this playerID/currency exists, this API call will create it if not
        try {
            $Wallet = Invoke-RestMethod -Uri ("https://gameshub.gameassists.co.uk/api/Player/ServerId/" + $ServerId + "/currencyCode/" + $Currency.trim()) -ContentType 'application/json' `
            -Headers @{ Authorization = "Bearer $OktaToken" }

            # Find the playerID for this currency
            $PlayerID = $Wallet | Where-Object {$_.CurrencyCode -eq $Currency.trim() -and $_.serverId -eq $ServerId} | Select-Object -ExpandProperty playerId
            Write-Verbose ("[$(Get-Date)] Player/Wallet ID: $PlayerID")
            
            # set the balance to 10,000.00 in 1:1 or equivalent in other currencies
            if ($Null -eq $Balance -or $Balance -le 0) {
                $Balance = ($CurrencyInfo.multiplierMaxBet * 1000000)
            }
            Write-Verbose ("[$(Get-Date)] Setting balance to $Balance")
            Invoke-RestMethod -Uri ("https://gameshub.gameassists.co.uk/api/Player/$PlayerID/CashBalance/" + [string]$Balance + "/BonusBalance/" + `
            [string]$Balance) -Method POST -ContentType 'application/json' -Headers @{ Authorization = "Bearer $OktaToken" } -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "An error occured attempting to update the player balance. Games Hub UserID: $HubUserID PlayerID: $PlayerID `
            Error details: $_"
        }

        # Get the game launch URI including the launch token
        $LaunchParams = @{
            serverID            = $ServerId;
            userId              = $HubUserID.trim();
            currency            = $Currency.trim();
            language            = $Language.trim();
            gameName            = $LaunchCode.trim();
            showPlaycheck       = $true;
            showHelp            = $true;
            useIngameInterface  = $true
        }

        Write-Verbose "[$(Get-Date)] Games Hub Launchurl Request Body:"
        foreach($k in $LaunchParams.Keys) { Write-Verbose "$k $($LaunchParams[$k])" }

        try {
            $LaunchURI = Invoke-RESTMethod -Uri "https://gameshub.gameassists.co.uk/api/Game/launchUrl" -Method POST -ContentType 'application/json' -Headers @{ Authorization = "Bearer $OktaToken" } `
            -Body $($LaunchParams | ConvertTo-Json)
        } catch {
            Write-Error "An error occured retrieving the Game Launch URI from Games Hub. `
            ServerID: $ServerId userId: $HubUserID Currency: $Currency Language: $Language `
            UGL launch code: $LaunchCode"
            Throw $_
        }
    }

    Write-Host -Foregroundcolor DarkCyan "Launch URI: $LaunchURI"
    
    if (!($NoConfirm.IsPresent)) {
        Write-Host "Press ENTER to launch the game, C to copy the launch URI to clipboard, anything else to continue without any further action..."
        $Waitkey = [System.Console]::ReadKey()
        Write-Host ""

        if ($Waitkey.Key -eq 'C') {
            Set-Clipboard $LaunchURI
            Write-Host -Foregroundcolor DarkCyan "Launch URI copied to clipboard."
        }
        if ($Waitkey.Key -ne 'Enter') {
            Return
        }
    }
    Write-Host -Foregroundcolor DarkCyan "Opening default browser to launch the selected game, please wait..."
    Start-Process $LaunchURI
    Write-Host ""
}



function Get-QFAAMSStatus {
    <#
        .SYNOPSIS
            Retrieves AAMS Participation status from the ADM Italy site.

        .DESCRIPTION
            Retrieves AAMS Participation status from the ADM Italy site www.adm.gov.it
            You must specify an AAMS Participation code (A 16 digit code beginning with the letter N).

            Details of the round returned include the Date, Bet Amount, Remote Session ID and Status.
            The Status can be either 'Riscosso' (Round is completed and closed) or 'Registrato' (Round is open and needs to be processed).

        .PARAMETER AAMSCode
            Specifies the Participation Code of the game session you wish to check.
            You may specify multiple Participation Codes, seperated by commas.

        .EXAMPLE
            Get-QFAAMSStatus N123456789012345,N098765432109876
                Retrieves AAMS status for the two specified Participation Codes.

        .INPUTS
            This cmdlet accepts pipeline input for the AAMS Participation Code.

        .OUTPUTS
            A PSCustomObject for each Game Session with the following members:
                System.Management.Automation.PSCustomObject
                    Name            MemberType      Definition
                    ----            ----------      ----------
                    Agent              NoteProperty string
                    Bet Amount         NoteProperty string
                    Date               NoteProperty datetime
                    Participation Code NoteProperty string
                    Remote Session ID  NoteProperty string
                    Status             NoteProperty string

        .NOTES
            Author:     Chris Byrne
            Email:      christopher.byrne@derivco.com.au

        .LINK
            https://confluence.derivco.co.za/pages/viewpage.action?pageId=450497969

    #>
    [cmdletBinding()]
    [alias('aams')]
    param(
    # AAMS participation codes
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$AAMSCode
    )

    # Loop through each code, remove any duplicates
    $AAMSCode =  $AAMSCode | Sort-Object -Unique
    Foreach ($Code in $AAMSCode) {

        # Link to the AAMS site to get the participation details
        $AAMSLink = "https://www.adm.gov.it/portale/web/guest/monopoli/giochi/giochi_abilita/giochi_ab_verifica?p_p_id=" +
            "it_sogei_wda_web_portlet_WebDisplayAamsPortlet&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view&p_p_cacheability=cacheLevelPage&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_matr=" + 
            $Code + "&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_op=inizia&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_sub=" +
            "Cerca&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_invia=Invia&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_CACHE=NONE&_it_sogei_wda_web_portlet_WebDisplayAamsPortlet_coolcap="
        
        # Contact AAMS site and retrieve the participation details
        $AAMSTemp = Invoke-RESTMethod $AAMSLink

        # normal -match syntax doesn't work so use the Matches method of a regex object, to process the returned AAMS data
        $RegexMatches = ([regex]"<strong>(.*):?</strong>.*<span>(.*)</span>").Matches($AAMSTemp)

        # $RegexMatches will now be an array, with a member for each matching line. 
        # Create a hash table for this participation data, before we process it and add it to $AAMSData
        $AAMSTemp = @{}

        # Process each member of $RegexMatches and check the Groups member of each one, this will have the text from our regex capture groups.
        Foreach ($Match in $RegexMatches) {

            # switch statement used to translate each different line in the results
            switch (($Match.groups[1].Value).trim()) {
                "Diritto di partecipazione:" {
                    # The same 'right to participate' heading is used to show the participation code and the status (riscosso/registrato), only difference seems to be the colon at the end for the code.
                    $AAMSTemp.add("Participation Code", ($Match.groups[2].Value).trim())
                }

                "Sessione:" {
                    $AAMSTemp.add("Remote Session ID", $(($Match.groups[2].Value).trim() -replace " del .*$"))
                }

                "Concessionario:" {
                    $AAMSTemp.add("Agent", ($Match.groups[2].Value).trim())
                }

                "Importo giocata:" {
                    $AAMSTemp.add("Bet Amount", ($Match.groups[2].Value).trim())
                }

                "Giocato in data:" {
                    [datetime]$Date = Get-Date $(($Match.groups[2].Value).trim() -replace " alle ore:") -Format "yyyy-MM-dd HH:mm:ss"
                    $AAMSTemp.add("Date", $Date)
                }

                "Diritto di partecipazione" {
                    # The same 'right to participate' heading is used to show the participation code and the status (riscosso/registrato), only difference seems to be the colon at the end for the code.
                    $AAMSTemp.add("Status", ($Match.groups[2].Value).trim())
                }

                Default {
                    # None of the above matched, just add it to the hash table as is
                    $AAMSTemp.add(($Match.groups[1].Value).trim(), ($Match.groups[2].Value).trim())
                }
            }
        }
        # Output to pipeline
        [pscustomobject]$AAMSTemp
    }
}
