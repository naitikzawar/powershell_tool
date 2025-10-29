###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                         Server Description Functions                        #
#                                    v1.6.4                                   #
#                                                                             #
###############################################################################

#Author: Chris Byrne - christopher.byrne@derivco.com.au

function Search-QFServerDetails {
    <#
    .SYNOPSIS
        Searches the Server Details web site for the specified Server Name and opens in the default browser.

    .DESCRIPTION
        This cmdlet is used to search the Server Details web site for the specified Server Name and opens the results in the default browser.
        This function is useful if you need to request the SA password for a particular SQL server or look at other server details.
        if you enter server name containing a backslash (\) the cmdlet assumes you are looking for a SQL Instance name and will remove everything before the slash.

        This command accepts multiple server names seperated by commas, or simply run the command with no parameters and you will be prompted to enter multiple server names on seperate lines.
        Press Enter on a blank line to begin the search.

    .EXAMPLE
        Search-QFServerDetails IOAUATCAS1\IOAUATCAS1
        Opens the Server Details site in the default browser and displays results for IOAUATCAS1

    .PARAMETER ServerName
        The name of the server you wish to search for.

    .INPUTS
        System.String
            You can pipe a string containing a single server name, or an array of multiple strings containing server names to this cmdlet.

    .OUTPUTS
        This cmdlet does not provide any pipeline output.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding()]
    [alias("sd")]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ServerName
    )
    begin {
        # The address of the Server Details site
        $SDSite = "sd.mgsops.net"

        # progressPreference set to silently continue, so test-netconnection doesn't show the progress bar. We'll temp set it to hide then restore the previous setting
        $oldProgressPreference = $global:progressPreference
        $global:progressPreference = 'silentlyContinue'
        try {
            # Confirm we have connectivity to the Server Details site
            Test-NetConnection $SDSite -port 443 -WarningAction Stop | Out-Null
        }
        catch {
            Throw "Unable to connect to the Server Details site: $SDSite - Please ensure you have network connectivity."
        }
        finally {
            $global:progressPreference = $oldProgressPreference # Restore the previous setting
        }
    }
    process {
        foreach ($Server in $ServerName) {
            # if a slash was passed in the server name it is probably the server host name followed by the SQL instance. We only want the instance name so strip the host name portion
            $Server = ($Server.ToString().trim() -replace "^.*\\","")

            # Buid the URI for the SD site request
            $SDURI = "https://" + $SDSite + "/Search/ResultsPage?q=" + $Server + "&chk_field_option=FQDN&chk_field_option=Description&chk_type=Contains&chk_option=Server&chk_option=SQL&chk_option=Web&chk_option=Cluster&chk_option=BaseNode&chk_option=Storage&chk_option=NetworkDevice&chk_option=StandAlone&chk_status_option=-1&chk_status_option=2&chk_status_option=3&chk_status_option=4&chk_status_option=5&chk_status_option=7'"

            Write-Verbose ("[$(Get-Date)] Server Name: $Server")
            Write-Verbose ("[$(Get-Date)] Server Details URI: $SDURI")

            # Open the URI in the default browser
            Start-Process $SDURI
        }
    }
}



function Get-QFSQLServerSAPassword {
    <#
    .SYNOPSIS
        Retrieves the SA password for the specified SQL Server.

    .DESCRIPTION
        This cmdlet will retrieve the SA password for the specified SQL Server.

        By default, a menu will be displayed allowing you to scroll through the results for each SQL server,
        copy the hostname or password to clipboard, or export all results to Excel.
        If the 'NoMenu' parameter is specified, an object will be output to pipeline containing
        the SQL server host name, description and its SA password, or any error returned from the SD API request.

        You may pass a SQL Server host name using the -ServerName parameter; multiple server names can be provided
        either on the pipeline in an array, or by seperating each server name with a comma.
        
        If you do not provide a server name, the cmdlet will look for a CSV file named 'SQLServers.csv' in the same
        folder where the module was imported from. This CSV file must contain a column named 'SQLHost' with the
        fully qualified domain name (FQDN) of the SQL server. The user will then be presented with a GridView to
        select the required SQL servers. Any additional coloumns will also be displayed in the GridView for the
        user's information, but will otherwise not be used by this cmdlet.

        The CSV file may optionally contain a 'Username' column that specifies the SQL user account to retrieve a password for.
        This field may specify any account supported by the SD API.
        e.g. DBReadOnly_PI, DBReadWrite_PI etc
        If the CSV file does not contain any value in this column, or the column is missing,
        the 'sa' password will be retrieved by default.
        
        Specifying the 'Username' parameter will allow you to manually specify the desired SQL user account.
        This parameter will always take precedence over any value in the CSV file.

        This command requires that all SQL servers must be in fully qualified domain name (FQDN) format.
        e.g. 'UATQF1CAS5.ioa.mgsops.com'

        You must also provide a reason for requesting the SA password - for example, an REQ ticket number.
        You may specify this via the 'Reason' parameter, or the cmdlet will prompt you to enter it.

        A request is made to SD API for each specified SQL server, including the specified Reason.
        If requesting passwords for multiple servers, the same Reason will be supplied for each one.

        This command can also be invoked with the alias 'sa' for interactive mode. 
        The alias 'saa' will invoke the command as if the 'NoMenu' parameter was specified.

    .EXAMPLE
        Get-QFSQLServerSAPassword
        Prompts the user to enter a Reason for request, then presents a GridView allowing the user to select one
        or multiple SQL servers; then attempts to retrieve the SA password for each SQL server.

    .EXAMPLE
        Get-QFSQLServerSAPassword -Username DBReadOnly_PI
        Prompts the user to enter a Reason for request, then presents a GridView allowing the user to select one
        or multiple SQL servers; then attempts to retrieve the password for the 'DBReadOnly_PI' account on each SQL server.

    .EXAMPLE
        Get-QFSQLServerSAPassword UATQF1CAS5.ioa.mgsops.com
        Retrieves the SA password for the UAT casino server. You will be prompted to enter a reason.
        A menu will be displayed allow you to copy the password or server hostname to clipboard, or export these details to excel.

    .EXAMPLE
        Get-QFSQLServerSAPassword -ServerName UATQF1CAS5.ioa.mgsops.com -Reason REQ123456
        Retrieves the SA password for the UAT casino server and provides an REQ ticket number for the Reason parameter.
        A menu will be displayed allow you to copy the password or server hostname to clipboard, or export these details to excel.

    .EXAMPLE
        Get-QFSQLServerSAPassword -ServerName UATQF1CAS5.ioa.mgsops.com -Reason REQ123456
        Retrieves the SA password for the UAT casino server and provides an REQ ticket number for the Reason parameter.
        A menu will be displayed allow you to copy the password or server hostname to clipboard, or export these details to excel.

    .EXAMPLE
        Get-QFSQLServerSAPassword -ServerName UATQF1CAS5.ioa.mgsops.com,UATQF1ARC5.ioa.mgsops.com -Reason REQ123456
        Retrieves the SA password for the UAT casino server, and the UAT Casino Archive server.
        Provides an REQ ticket number for the Reason parameter for both requests.
        A menu will be displayed allow you to copy the password or server hostname to clipboard, or export these details to excel.

    .EXAMPLE
        Get-QFSQLServerSAPassword -ServerName UATQF1CAS5.ioa.mgsops.com,UATQF1ARC5.ioa.mgsops.com -Reason REQ123456 -NoMenu
        Retrieves the SA password for the UAT casino server, and the UAT Casino Archive server.
        Provides an REQ ticket number for the Reason parameter for both requests.
        No menu will be displayed - all results will be output directly to pipeline without any further action.

    .EXAMPLE
        Get-QFSQLServerSAPassword -ServerName UATQF1CAS5.ioa.mgsops.com,UATQF1ARC5.ioa.mgsops.com -Reason REQ123456 -NoMenu -Username DBReadOnly_PI
        Retrieves the 'DBReadOnly_PI' account password for the UAT casino server, and the UAT Casino Archive server.
        Provides an REQ ticket number for the Reason parameter for both requests.
        No menu will be displayed - all results will be output directly to pipeline without any further action.

    .PARAMETER NoMenu
        Disables the interactive menu for scrolling through the list of returned SQL server passwords.
        All returned passwords are output to pipeline without any further input.
        This parameter is useful if you want to call this cmdlet from another cmdlet or function.
    
    .PARAMETER ServerName
        The name of the server you wish to retrieve the SA password for.

        This must be in fully qualified domain name (FQDN) format.
        e.g. 'UATQF1CAS5.ioa.mgsops.com'
        This is required by the SD API, so if an FQDN is not supplied, the API will return the error:
        'Username does not exist for this FQDN.'

    .PARAMETER Reason
        A string describing your reason for requesting an SA password.
        This is recorded along with your user name in the API audit logs and may be subject to review by Security.

    .PARAMETER SDURI
        The URI of the SD API endpoint. This defaults to 'https://sdapi.mgsops.net/ServerDetailsAPI.svc' if not specified.

    .PARAMETER Username
        Specify the Username of the account that you wish to retrieve the password for.
        By default, the 'sa' account password will be retrieved, but you can specify any account supported by the SD API.
        e.g. DBReadOnly_PI, DBReadWrite_PI etc

        If this parameter is not specified, it defaults to 'sa'; or if a CSV file is used that contains a 'Username' column, 
        the value in this column will be used instead of 'sa'.

        If this parameter is specified, it takes precedence and will be used instead of any Username in the CSV file.
        If a username is provided that is not supported by SD API, the API will return the error:
        'Username does not exist for this FQDN.'

    .INPUTS
        System.String
            You can pipe a string containing a single server name, or an array of multiple strings containing server names to this cmdlet.
            You may also pipe a string for the Reason and Username parameters.

    .OUTPUTS
        TypeName: System.Management.Automation.PSCustomObject
            Name        MemberType      Definition
            --------------------------------------
            Name        NoteProperty    string
            Hostname    NoteProperty    string
            Username    NoteProperty    string
            Password    NoteProperty    string

    .NOTES
        Author:     Alejandro Oliva - original SD API request (New-WebServiceProxy)
                    Chris Byrne - conversion to Invoke-RESTMethod, menu and excel features
        Email:      alejandro.oliva@derivco.es / christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding()]
    [alias("sa","saa")]
    param(

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]$NoMenu,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [string[]]$ServerName,

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Reason,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$SDURI="https://sdapi.mgsops.net/ServerDetailsAPI.svc",

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,position=2)]
        [ValidateNotNullOrEmpty()]
        [string]$Username="sa"
    )
    Begin {

        function Get-SQLSAPassword {
            # Local function to invoke the SOAP request to the SD API
            param(
                [string]$SdFqdn,
                [string]$SdReason,
                [string]$SdUsername
            )
            # Local variables for the SOAP API request, including the request header and XML body
            
            [int]$SdPasswordType = 0
            $Headers = @{ 'Content-Type' = 'text/xml'; "SOAPAction" = "http://tempuri.org/IServerDetailsApi/GetServerPassword" }
            $Body = @"
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
            <soapenv:Header/>
            <soapenv:Body>
                <tem:GetServerPassword>
                    <!--Optional:-->
                    <tem:fqdn>$SdFqdn</tem:fqdn>
                    <!--Optional:-->
                    <tem:username>$SdUsername</tem:username>
                    <!--Optional:-->
                    <tem:reason>$SdReason</tem:reason>
                    <!--Optional:-->
                    <tem:passwordType>$SdPasswordType</tem:passwordType>
                </tem:GetServerPassword>
            </soapenv:Body>
            </soapenv:Envelope>
"@
            # Unfortunately the here-string for the body doesn't parse properly if there is additional white space on the last line, so we have to remove the indentation.

            <# OLD SOAP API request method - WebServiceProxy - not supported in PowerShell Core or 7
            $sd = New-WebServiceProxy -Uri $SDURI -UseDefaultCredential
            $DBPassword = $sd.GetServerPassword($SdFqdn, $SdUsername, $SdReason, $SdPasswordType, $SdPasswordTypeSpecified);
            Now changed to Invoke-RestMethod.
            #>

            # This hash table will be parameters for splatting to Invoke-RestMethod
            $RequestArgs =
            @{
                Uri = "$SDURI/ServerDetailsHttpsAPI.svc"
                Method = "POST"
                Headers = $Headers
                Body = $Body
                UseDefaultCredentials = $true
                UseBasicParsing = $true
            }

            # If running PS 7, also add -SkipCertificateCheck parameter. This is not available in PS 5.
            # this fixes issue some users report with certificate errors when invoking the SD API
            # Seems to be a local environment issue - works fine without this parameter in Australia *shrug*
            If ($PSVersionTable.PSVersion.Major -gt 6 ) {
                $RequestArgs.Add('SkipCertificateCheck',$true)
            }

            Write-Verbose ("[$(Get-Date)] Invoke-RestMethod Parameters:")
            foreach($k in $RequestArgs.Keys) { Write-Verbose "$k $($RequestArgs[$k])" }

            Try {
                Write-Verbose ("[$(Get-Date)] Invoking SOAP API request for $SdFqdn")
                $DBPassword = Invoke-RestMethod @RequestArgs 
            }
            Catch {
                $Exception = $_.Exception
                Write-Warning "Failed to retrieve password for $SqFqdn"
                Write-Warning $Exception.Message
                Return
            }
            $DBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult
        }

        

        function Invoke-SQLServersMenu {
            # Local function to display multiple returned SA passwords and allow scrolling or other options
            
            param(
            # Counter of items to loop through
            [Parameter(Mandatory = $true)]
            [int]$Counter,

            # Max number of items
            [int]$ItemCount = 1
            )

            # Display menu prompting user for action
            # Hide next/back/all options depending on how many items are available
            If ($Counter + 1 -lt $ItemCount) {
                Write-Host "$([char]27)[36m[SPACE/`u{2192}]$([char]27)[0m Next Server`t" -NoNewline
            }
            If ($Counter -gt 0) {
                Write-Host "$([char]27)[36m[B/`u{2190}]$([char]27)[0m Back/Previous Server`t" -NoNewline
            }
            If ($Counter + 1 -lt $ItemCount) {
                Write-Host "$([char]27)[36m[A]$([char]27)[0m Display ALL remaining Servers`t" -NoNewline
            }
            Write-Host ""
            Write-Host "$([char]27)[36m[C]$([char]27)[0m Copy SA Password`t$([char]27)[36m[S]$([char]27)[0m Copy Server Hostname`t$([char]27)[36m[U]$([char]27)[0m Copy Username`t$([char]27)[36m[X]$([char]27)[0m Export to Excel`t$([char]27)[36m[Anything else]$([char]27)[0m Quit"
            $Waitkey = [System.Console]::ReadKey()
            Write-Host ""
            Write-Verbose ("[$(Get-Date)] Option selected:" + $Waitkey.key)
    
                switch ($Waitkey.key) {
                    'A' { 
                        # All remaining items - return -1 so calling function should know to display everything remaining
                        If ($Counter + 1 -ge $ItemCount) {
                            Break
                        }
                        Return -1
                    }
                    'Spacebar' {
                        # Increment counter so calling function knows to display the next object
                        Return $Counter + 1
                    }
                    'RightArrow' {
                        Return $Counter + 1
                    }
                    'B' {
                        # Decrement counter so calling function knows to display the previous object
                        if ($Counter -le 0) {
                            Break
                        } else {
                            Return $Counter - 1
                        }
                    }
                    'LeftArrow' {
                        if ($Counter -le 0) {
                            Break 
                        } else {
                            Return $Counter - 1
                        }
                    }
                    'X' {
                        # export to excel
                        Return -999
                    }
                    'C' {
                        # copy password to clipboard
                        Return -997
                    }
                    'S' {
                        # copy hostname to clipboard
                        Return -998
                    }
                    'U' {
                        # copy hostname to clipboard
                        Return -996
                    }
                    Default {
                        # Exit
                        Break
                    }
                }
            }

        # End local functions

        # Check if a SQL server FQDN hostname was passed as a parameter
        If ($Null -eq $ServerName -or $ServerName -eq '') {
            # If no ServerName parameter passed, see if there is a CSV file we can use
            If (Test-Path "$PSScriptRoot\SQLServers.csv" -PathType Leaf) {
                Try {
                    Write-Verbose ("[$(Get-Date)] Reading SQL Servers list from $PSScriptRoot\SQLServers.csv")
                    $SQLDBList = Get-Content -Encoding UTF8 "$PSScriptRoot\SQLServers.csv" -ErrorAction Stop | ConvertFrom-CSV -ErrorAction Stop
                    # Check if we got any records from the CSV file
                    If (@($SQLDBList).Count -lt 1) {Throw "The CSV file did not contain any records."}
                    # Check if there is a property named SQLHost on the SQLDBList object
                    If (!($SQLDBList|Get-Member -Name SQLHost -MemberType Properties)) {Throw "The CSV file did not contain a SQLHost column."}
                    Write-Verbose ("[$(Get-Date)] Records retrieved from CSV file: $(@($SQLDBList).Count)")
                }
                Catch {
                    $Exception = $_.Exception
                    Throw "No ServerName parameter was specified, and failed to import SQL Servers list from the CSV file $PSScriptRoot\SQLServers.csv - $Exception.Message"
                }

                # If we got some servers from the CSV file, bring up Out-GridView to let the user pick the servers.
                If ($null -ne $SQLDBList) {
                    # create PSStandardMembers object that tells PowerShell which CSV columns should be visible in Out-GridView
                    # show "Name", "Group" and "SQLHost" columns. Add column names to the list here if you want any others to be visible
                    [string[]]$CsvVisible = 'Name', 'Group', 'SQLHost'
                    [System.Management.Automation.PSMemberInfo[]]$info = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet',$CsvVisible)
                    # Add-Member will set DefaultDisplayPropertySet on the object, to only show columns specified in $CsvVisible
                    $SQLDBList | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $info
                    # Now display list of SQL servers in Out-GridView
                    [PSCustomObject]$ServerName = $SQLDBList | Out-GridView -OutputMode Multiple -Title "Select the required SQL Servers"
                    Write-Verbose ("[$(Get-Date)] Selected SQL Servers: " + $ServerName.SQLHost)
                }
            }
            # After selecting with gridview, check again to make sure we actually do have something in ServerName
            If ($Null -eq $ServerName.SQLHost -or $ServerName.SQLHost -eq '') {
                Throw "No SQL Server name specified. You must specify a fully qualified domain name of a SQL Server host in the ServerName parameter, or provide a CSV file with a SQLHost column containing this data."
            }
        }
 




    }
    Process {

        # array object of returned SQL server passwords, for pipeline output or displaying in menu
        $OutputObject = @()

        foreach ($SQLHost in $ServerName){
            $TempObject = New-Object -TypeName psobject

            # Check if a username parameter was set. if so, use that value for the request, otherwise use the value in the CSV. Otherwise use the default value for Username parameter.
            if (($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Username')) -or ($null -eq $SQLHost.Username -or $SQLHost.Username.trim() -eq '')) {
                $SDUsername = [string]$Username.trim()
            } else  {
                $SDUsername = [string]$SQLHost.Username.trim()
            }

            If ($PSBoundParameters.ContainsKey("ServerName")) {
                # Servers provided by command line parameter
                $TempObject | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $SQLHost.trim()
                $SdFqdn = $SQLHost
            } else {
                # Servers selected from Out-Gridview
                $TempObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $SQLHost.Name.trim()
                $TempObject | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $SQLHost.SQLHost.trim()
                $SdFqdn = $SQLHost.SQLHost
            }

            # Call the function to invoke the SOAP request and build up our output object
            $DatabasePassword = Get-SQLSAPassword -SdFqdn $SdFqdn -SdReason $Reason -SDUsername $SDUsername

            # Add the current SQL Server password and Username to the output array object
            $TempObject | Add-Member -MemberType NoteProperty -Name "Username" -Value $SDUsername
            $TempObject | Add-Member -MemberType NoteProperty -Name "Password" -Value $DatabasePassword.trim()
            $OutputObject += $TempObject
        }

        if ($NoMenu.IsPresent -or $psCmdlet.myinvocation.line -match "^saa") {
            #Output all retrieved passwords and other details to pipeline without any further action
            $OutputObject
        } else {
            # Interactive menu - Loop through all returned SQL Servers, call Invoke-SQLServersMenu for each one
            $i = 0
            do {
                # Arrays are 0 indexed so add 1 to our counter for display
                Write-Host -ForegroundColor Yellow "Displaying SQL Server $($i + 1) of $(@($OutputObject).Count)"
                $OutputObject[$i] | Format-List
                
                If ($i -ge (@($OutputObject).Count) -1) {
                    # Disable DisplayAll and show item menu on last item
                    $DisplayAll = $false
                }
                If ($DisplayAll -ne $true) {
                    $Counter = Invoke-SQLServersMenu -Counter $i -ItemCount @($OutputObject).Count
                }
                Write-Verbose ("[$(Get-Date)] Invoke-SQLServersMenu return value: $Counter - i object value: $i")
                if ($Counter -eq -1) {
                    # Display all the remaining items
                    $DisplayAll = $true
                    $i += 1
                } elseif ($Counter -eq -999) {
                    # Export to Excel
                    try {
                        [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'SQL Server Passwords - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                        If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "SQL Server Passwords - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                        If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                        Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData $OutputObject -ExcelDestWorksheetName SQLServers
                        if (Test-Path $ExcelFilename -PathType Leaf) {
                            Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                        }
                    } catch {
                        Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                    }
                } elseif ($Counter -eq -998) {
                    Set-Clipboard $OutputObject[$i].Hostname
                    Write-Host -ForegroundColor Yellow "SQL Server Hostname copied to clipboard!"
                    Write-Host ""
                } elseif ($Counter -eq -997) {
                    Set-Clipboard $OutputObject[$i].Password
                    Write-Host -ForegroundColor Yellow "SA password copied to clipboard!"
                    Write-Host ""
                } elseif ($Counter -eq -996) {
                    Set-Clipboard $OutputObject[$i].Username
                    Write-Host -ForegroundColor Yellow "SQL Username copied to clipboard!"
                    Write-Host ""
                } else {
                    $i = $Counter
                }
            } until ($i -lt 0 -or $i -ge @($OutputObject).Count)
        }
    }
}