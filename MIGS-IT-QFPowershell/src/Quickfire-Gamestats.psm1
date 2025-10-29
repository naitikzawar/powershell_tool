###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                       Game Statistics Report Functions                      #
#                                   v1.6.4                                   #
#                                                                             #
###############################################################################

#Author: Chris Byrne - christopher.byrne@derivco.com.au


function Get-QFGameStats {
    <#
    .SYNOPSIS
        Generates a Game Monitor / Game Statistics report link for the specified player.

    .DESCRIPTION
        This cmdlet is used to generate a Game Monitor / Game Statistics report link for the specified player.
        The default behaviour is to output a list of game statistics reports available for the player.
        The list of reports can be filtered by Game Name (using the -GameName parameter); Game ModuleID (using the -MID parameter) and Game ClientID (Using the -CID parameter).

        The player Login (aka UserName) must include the prefix from the Casino database. The prefix consists of two characters followed by an underscore.
        Alternatively you can specify a numeric UserID, CasinoID and GamingSystemID.
        You can retrieve these details using Invoke-QFPortalRequest and looking up the details for the CasinoID.
        The GamingSystemID can be found in the output property GamingSystemID.
        The Prefix can be found under the productSettings property, under "Register - SGI JIT Account Creation Prefix".

        You may optionally pass the "-OpenBrowser" parameter to open each Game Statistics report in the default system web browser.
        You may optionally pass the '-SavePDF' switch parameter to automatically save the generated report to a PDF file using Edge's 'Save as PDF' printer, then open in the default PDF Viewer program.
        Specifying the '-NoViewPDF' switch parameter in combination with the '-SavePDF' switch parameter will not open the PDF file, and just silently save the generated PDF without further user interaction.

        Running this cmdlet using the alias 'gsw' is equivalent to 'Get-QFGameStats -OpenBrowser'
        Running this cmdlet using the alias 'gsp' is equivalent to 'Get-QFGameStats -SavePDF'
        Running this cmdlet using the alias 'gss' is equivalent to 'Get-QFGameStats -SavePDF -NoViewPDF'

    .EXAMPLE
        Get-QFGameStats -Login XY_abcdef
        Generates a list of Game Statistics reports for the specified Player Login Name. The Login Name must include the prefix from the Casino database (two characters followed by an underscore).

    .EXAMPLE
        Get-QFGameStats -UserID 12345 -GamingSystemID 321 -CasinoID 98765
        Generates a list of Game Statistics reports for the specified Player UserID. The GamingSystemID and CasinoID parameters are required if a player UserID is specified.

    .EXAMPLE
        Get-QFGameStats -Login XY_abcdef -CasinoID 54321
        Generates a list of Game Statistics reports for the specified Player Login Name and CasinoID. This can be used to specify a CasinoID if the same Login Name exists on different sites.

    .EXAMPLE
        Get-QFGameStats -Login XY_abcdef -MID 10000 -CID 50300
        Generates a list of Game Statistics reports for the specified Player Login Name and filters the list to games with the ModuleID 10000 and the ClientID 50300.

    .EXAMPLE
        Get-QFGameStats -Login XY_abcdef -GameName "Thunderstruck"
        Generates a list of Game Statistics reports for the specified Player Login Name and filters the list to games with the word "Thunderstruck" in the name.

    .EXAMPLE
        Get-QFGameStats -UserID 12345 -GamingSystemID 321 -CasinoID 98765 -GameName "Thunderstruck" -OpenBrowser
        Looks up all Game Statistics reports for the specified player UserID and for games with the word "Thunderstruck" in the name, outputs a list of reports to the pipeline, then opens each one in the default web browser.

    .EXAMPLE
        Get-QFGameStats -Login XY_abcdef -MID 10000 -SavePDF
        Looks up all Game Statistics reports for the specified Player Login Name and for games with ModuleID 10000, outputs a list of reports to the pipeline, saves each report as a PDF in the current working folder, and then opens them in the default PDF viewer.

    .EXAMPLE
        Get-QFGameStats -UserID 12345 -GamingSystemID 321 -CasinoID 98765 -MID 10000 -SavePDF -NoViewPDF
        Looks up all Game Statistics reports for the specified player UserID and for games with ModuleID 10000, outputs a list of reports to the pipeline, and silently saves each report as a PDF in the current working folder. Does not open the saved PDF files.

    .PARAMETER CasinoID
        The CasinoID/ServerID that the specified player belongs to.

    .PARAMETER CID
        Filter list of Game Statistics reports by Game ClientID. Requires exact match.

    .PARAMETER GameName
        Filter list of Game Statistics reports by Game Name. Supports regular expressions as per PowerShell's Match operator.

    .PARAMETER GamingSystemID
        The Gaming System ID number for the specified CasinoID.
        This is required if the UserID parameter is specified.

        You can retrieve this using Invoke-QFPortalRequest and looking up the details for the CasinoID.
        The property GamingSystemID in the output will contain the correct value for this parameter.

    .PARAMETER HostName
        Specifies the Host Name of the Game Monitor Reports web app. Defaults to https://gamemonitor.mgsops.net - this parameter doesn't need to be set unless the Game Monitor Reports web app address changes.

    .PARAMETER Login
        The Login Name of the player you wish to generate a Game Statistics report for.
        You cannot specify both a Login Name and a UserID.

    .PARAMETER MID
        Filter list of Game Statistics reports by Game ModuleID. Requires exact match.

    .PARAMETER NoViewPDF
        if this switch is present, Play Check PDF's will not be opened in the default PDF viewer program automatically after they are created.
        This is useful if you are automating this function as it will not provide any other graphical output.
        This switch has no effect if the SavePDF switch is not also present.
        Running the command using the alias 'gss' is equivalent to setting this parameter.

    .PARAMETER OpenBrowser
        if this switch is present, the function will open each Game Statistics report in the default system web browser.
        Running the command using the alias 'gsw' is equivalent to setting this parameter.

    .PARAMETER SavePDF
        if this switch is present, the function will use Edge's Print To PDF feature, to save the play check as a PDF in the current directory.
        The PDF file will then open in the default PDF viewer program.
        Running the command using the alias 'gsp' is equivalent to setting this parameter.

    .PARAMETER UserID
        The UserID of the player you wish to generate a Game Statistics report for.
        You cannot specify both a Login and a UserID.
        If you specify a UserID you must also specify a GamingSystemID.

    .INPUTS
        System.String
            You can pipe a string that contains a Casino Player Login Name, or a UserID, or a Game Name to this cmdlet.

        System.Int32
            You can pipe an integer that contains a Game MID, game CID, or a CasinoID to this cmdlet.

    .OUTPUTS
        A list of Game Statistics reports available for the specified player and other parameters.
        The Game Name, MID, and CID will be output to host by default.
        Additional properties such as URI, CasinoID and GamingSystemID will also be output to pipeline.

            System.Management.Automation.PSCustomObject

                Name        MemberType   Definition
                ----        ----------   ----------
                GameName        NoteProperty string
                CID             NoteProperty int
                MID             NoteProperty int
                URI             NoteProperty string
                CasinoID        NoteProperty int
                GamingSystemID  NoteProperty int

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding(DefaultParameterSetName="Login")]
    [alias("gs","gsp","gss","gsw")]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0,ParameterSetName="Login")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript ({($_.Length -gt 3)})]
        [string]$Login,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=4,ParameterSetName="Login")]
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=4,ParameterSetName="UserID")]
        [int]$CasinoID,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
        [string]$GameName,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=2)]
        [int]$MID,

        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=3)]
        [int]$CID,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,Position=5,ParameterSetName="Login")]
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=5,ParameterSetName="UserID")]
        [int]$GamingSystemID,

        [Parameter(Mandatory=$false)]
        [switch]$OpenBrowser,

        [Parameter(Mandatory=$false)]
        [switch]$SavePDF,

        [Parameter(Mandatory=$false)]
        [switch]$NoViewPDF,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Hostname = "gamemonitor.mgsops.net",

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0,ParameterSetName="UserID")]
        [ValidateNotNullOrEmpty()]
        [int]$UserID

    )

    begin {
        # Required for HTML Encode/Decode
        Add-Type -AssemblyName System.Web

        # The address of the Game Monitor site; strip http/s and any path after the hostname
        $Hostname = $Hostname.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
        Write-Verbose ("[$(Get-Date)] Game Monitor site hostname: $Hostname")

        # Check if MS Edge browser is installed
        if (!(Test-Path -Path 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -PathType Leaf)) {
            Throw "MS Edge browser must be installed to generate Game Statistics reports. Please ensure it is installed properly."
        }

        # progressPreference set to silently continue, so test-netconnection, invoke-webrequest etc doesn't show the progress bar. Big speedup in some cases
        $global:progressPreference = 'silentlyContinue'

        # Confirm we have connectivity to the game monitor site
        try {
            Test-NetConnection $Hostname -port 443 -WarningAction Stop | Out-Null
        }
        catch {
            Throw "Unable to connect to the Game Monitor site: $Hostname - Please ensure you have internet connectivity."
        }

    }
    process {
        # Several REST methods are used to load the list of games for a player and ultimately provide the Game Stats report URI

        # If we were given a Login parameter, look up the gaming systems for the specified Login
        # If we were given a UserID parameter we can skip this as GamingSystemID is mandatory parameter and must be provided already
        If ($psCmdlet.ParameterSetName -eq "Login") {
            $Login = $Login.Trim()
                if ($null -eq $GamingSystemID -or $GamingSystemID -eq 0) {
                # Get the Gaming System ID's for the specified player login.
                $URI = "https://" + $Hostname + "/Players/GetGamingSystemsForLoginName?loginname=" + [System.Web.HttpUtility]::UrlEncode($Login)
                try {
                    Write-Verbose ("[$(Get-Date)] Invoking REST method: $URI")
                    $GamingSystemID = (Invoke-Restmethod -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop -Method GET -Uri $URI).Value | Select-Object -Unique
                }
                catch {
                    Write-Error "An error occured while trying to request Gaming Systems for the specified LoginID. Try manually specifying a GamingSystemID."
                    Throw $_.exception.Message
                }
            }
        }

        # Check we retrieved some values from this request
        if ($null -eq $GamingSystemID -or $GamingSystemID -eq 0) {
            Throw "No Gaming Systems found for the specified LoginID $Login or invalid GamingSystemID parameter specified."
        }
        Write-Verbose ("[$(Get-Date)] GamingSystemID: $GamingSystemID")

        # If Login is too short it will get results for multiple players on different Gaming Systems
        if ($GamingSystemID.Count -gt 1) {
            Throw "Multiple values returned for GamingSystemID. Please ensure you have entered the player's Login correctly including prefix."
        }

        # Get the Casino ID's for the specified player login and Gaming System
        # if a CasinoID paramter was specified we can skip this
        if ($CasinoID -gt 0) {
            $GamingSystemCasinos = $CasinoID
        }
        else {
            $URI = "https://" + $Hostname + "/Players/GetCasinosForGamingSystemAndLoginName?gamingSystemID=" + $GamingSystemID + "&loginName=" + [System.Web.HttpUtility]::UrlEncode($Login)
            try {
                Write-Verbose ("[$(Get-Date)] Invoking REST method: $URI")
                $GamingSystemCasinos = (Invoke-Restmethod -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop -Method GET -Uri $URI).Value | Select-Object -Unique
            }
            catch {
                Write-Error "An error occured while trying to request CasinoID's for the specified player."
                throw $_.exception.Message
            }
        }

        # Check we retrieved some values from this request or GamingSystemCasinos parameter was specified
        if ($null -eq $GamingSystemCasinos) {
            Throw "No CasinoID's found for the specified LoginID: $Login"
        }
        Write-Verbose ("[$(Get-Date)] GamingSystemCasinos: $GamingSystemCasinos")

        # Not sure what happens if we get multiple values returned from this request... so far I haven't been able to make this happen, its either 1 value or none.
        if ($GamingSystemCasinos.Count -gt 1) {
            Write-Warning "Multiple values returned for GamingSystemCasinos. This has never happened before!"
            Write-Warning "Please contact Christopher Byrne and provide the player details so I can investigate further!"
            return
        }

        # Get the list of Game Statistics Reports for the specified player login/userID, gaming system and CasinoID's
        If ($psCmdlet.ParameterSetName -eq "Login") {
            $URI = "https://" + $Hostname + "/Players/PlayerReportPartial?loginName=" + [System.Web.HttpUtility]::UrlEncode($Login) + "&gamingSystemID=" + $GamingSystemID + "&casinoID=" + $GamingSystemCasinos
        } elseif ($psCmdlet.ParameterSetName -eq "UserID") {
            $URI = "https://" + $Hostname + "/Players/PlayerReportPartial?userId=" + $UserID + "&gamingSystemID=" + $GamingSystemID + "&casinoID=" + $GamingSystemCasinos
        } else {
            Throw "No UserID or Login specified... we need those to get a Game Stats Report!"
        }

        try {
            Write-Verbose ("[$(Get-Date)] Invoking Web request: $URI")
            $PlayerGameReportsData = (Invoke-WebRequest -UseDefaultCredentials -ErrorAction Stop -Method GET -Uri $URI -UseBasicParsing).Links
        }
        catch {
            Write-Error "An error occured while trying to retrieve the list of Game Statistics reports for the specified player!"
            throw $_.exception.Message
        }

        # Check we retrieved some values from this request
        if ($PlayerGameReportsData.Count -eq 0) {
            Write-Verbose ("[$(Get-Date)] No Game Statistics Reports found for the specified Player Details:")
            Write-Verbose ("[$(Get-Date)] $(if ($Login -ne ''){"Login: $Login"})$(if ($UserID -gt 0){"UserID: $UserID"})$(if ($CasinoID -gt 0){" and CasinoID: $CasinoID"})")
            return
        }
        Write-Verbose ("[$(Get-Date)] Number of Game Stats reports returned for this player: $($PlayerGameReportsData.Count)")

        # $PlayerGameReports will be an array of objects for each game monitor page available for this player's UserID. The link to each page is in the href property.
        # Create a Custom Object to hold the full URI for each page, MID, CID and game name, and filter the list based on command parameters
        $PlayerGameReports = @()
        foreach ($PlayerGameReport in $PlayerGameReportsData) {
            # Get the Game Name from the link text on the game reports page. HTML Decode any escaped characters eg apostrophes
            $PlayerGameReportTemp = [PSCustomObject]@{ GameName = [System.Web.HttpUtility]::htmlDecode($PlayerGameReport.'data-sort') }

            # Check if the GameName matches the provided command line parameter - if not, exit the foreach loop and don't add this report to the list
            if ($null -ne $GameName -and $PlayerGameReportTemp.GameName -notmatch $GameName.trim()) {
                Write-Verbose ("[$(Get-Date)] Game Name: $($PlayerGameReportTemp.GameName) does not match GameName parameter: $($GameName.trim()) - won't add to the list")
                continue
            }
            Write-Verbose ("[$(Get-Date)] Game Name: $($PlayerGameReportTemp.GameName)")

            $PlayerGameReport.href -match "moduleID=([0-9]+)" | Out-Null # Populate $Matches with the MID
            # Check if the MID matches the provided command line parameter - if not, exit the foreach loop and don't add this report to the list
            if ($MID -gt 0 -and $MID -ne $Matches[1]) {
                Write-Verbose ("[$(Get-Date)] Game MID: $($Matches[1]) does not match MID parameter: $MID - won't add to the list")
                continue
            }
            $PlayerGameReportTemp | Add-Member -Name "MID" -Value $Matches[1] -MemberType NoteProperty
            Write-Verbose ("[$(Get-Date)] Game MID: $($PlayerGameReportTemp.MID)")

            $PlayerGameReport.href -match "clientID=([0-9]+)" | Out-Null # Populate $Matches with the CID
            # Check if the CID matches the provided command line parameter - if not, exit the foreach loop and don't add this report to the list
            if ($CID -gt 0 -and $CID -ne $Matches[1]) {
                Write-Verbose ("[$(Get-Date)] Game CID: $($Matches[1]) does not match CID parameter: $CID - won't add to the list")
                continue
            }
            $PlayerGameReportTemp | Add-Member -Name "CID" -Value $Matches[1] -MemberType NoteProperty
            Write-Verbose ("[$(Get-Date)] Game CID: $($PlayerGameReportTemp.CID)")
            # Add the hostname to the URI, and decode escaped '&' characters in the href property
            $PlayerGameReportTemp | Add-Member -Name "URI" -Value $("https://" + $Hostname + ($PlayerGameReport.href -replace "&amp;","&")) -MemberType NoteProperty
            $PlayerGameReportTemp | Add-Member -Name "CasinoID" -Value $GamingSystemCasinos -MemberType NoteProperty # CasinoID for this player
            $PlayerGameReportTemp | Add-Member -Name "GamingSystemID" -Value $GamingSystemID -MemberType NoteProperty # GamingSystem ID for this player
            $PlayerGameReports += $PlayerGameReportTemp # finally, add this custom object to the $PlayerGameReports array
            Write-Verbose ("[$(Get-Date)] Game Report URI: $($PlayerGameReportTemp.URI)")
        }

        # if no other parameters were set, just return $PlayerGameReports without any further action - skip the whole section below
        if ($OpenBrowser.IsPresent -or $SavePDF.IsPresent -or $NoViewPDF.IsPresent -or ($($psCmdlet.myinvocation.line) -match "gs[wps] ")) {
            foreach ($PlayerGameReport in $PlayerGameReports) {
                # if SavePDF parameter was passed, or command was run with alias 'gsp' or 'gss', print to PDF automatically without displaying in browser
                # Unfortunately using Edge with --save-as-pdf gives a blank page so here's a kinda kludgey workaround.
                # Save the page as HTML in a temp file, run a regex to replace all the relative paths to absolute for scripts, images, css etc; THEN use Edge to save as PDF
                if (($SavePDF.IsPresent) -or ($($psCmdlet.myinvocation.line) -match "gs[ps] ")) {
                    $Outfile = (Get-Location).path + "\Game Statistics - " + $($PlayerGameReport.GameName.trim() -replace '[\\/:"*?<>|]+','') +  " - " + $PlayerGameReport.MID + "-" + $PlayerGameReport.CID  + ".pdf"
                    $GameReportURI = $PlayerGameReport | Select-Object -ExpandProperty URI
                    $TempFile = "$Env:Temp\GameStatsTemp.html"
                    Write-Verbose ("[$(Get-Date)] Current Game Report URI: $GameReportURI")

                    # Save the report page into a temp file, use Replace operator to run a regex to adjust paths
                    try {
                        # Clear any existing temp file
                        if (Test-Path $TempFile -PathType Leaf) {
                            Remove-Item $TempFile -Force
                        }
                        Write-Verbose ("[$(Get-Date)] Saving Game Report into temp file: $TempFile")
                        $GameReportContent = (Invoke-WebRequest -uri $GameReportURI -ErrorAction Stop -UseDefaultCredentials -UseBasicParsing).Content
                        # Regex to insert Hostname of the game stats site into start of any relative links, so images/scripts aren't broken when the file is loaded in Edge again when converting to PDF
                        $GameReportContent -replace '(href|src)="(/.*)"', $('$1="https://' + $Hostname + '$2"') | Set-Content $TempFile -Encoding UTF8
                    }
                    catch {
                        Write-Warning "Unable to generate this Game Statistics Report: $($PlayerGameReport.GameName.trim()) - $($PlayerGameReport.MID)-$($PlayerGameReport.CID) and save into temporary file: $TempFile"
                        Write-Host $_
                        continue
                    }

                    # Print out the result of the game monitor report (green/yellow/red box on the report page indicating player's payout is within volatility or not)
                    if ($GameReportContent -match '<div class=\"alert-box (.*)\">(.*)</div>') {
                        # Matches[1] will be the class of alert box (success/warning/error). Matches[2] will be the message text
                        switch -wildcard ($Matches[1]) {
                            "*success*" { 
                                $TextColour = "DarkGreen"
                                $Result = "OK"
                            }
                            "*error*"  { 
                                $TextColour = "DarkRed"
                                $Result = "ERROR"
                            }
                            "*warning*" { 
                                $TextColour = "DarkYellow"
                                $Result = "WARNING"
                            }
                            Default { $TextColour = "White" }
                        }                         
                        $ResultText = $($Matches[2] -replace " his "," their ") # prefer gender neutral language for copy/pasting into ticket :)
                        Write-Host -ForegroundColor $TextColour "$($PlayerGameReport.GameName.trim()) - $($PlayerGameReport.MID)-$($PlayerGameReport.CID)"
                        Write-Host -ForegroundColor $TextColour $ResultText 
                        Write-Host ""

                        # Add the results to our reports array, which will be output to pipeline later

                        Add-Member -InputObject $PlayerGameReport -MemberType NoteProperty -Name "Result" -Value $Result
                        Add-Member -InputObject $PlayerGameReport -MemberType NoteProperty -Name "ResultText" -Value $ResultText
                    }

                    # Check that the PDF was generated correctly... sometimes its a blank page, check file size and retry until its big enough to have some content. try 5 times, waiting a bit longer each time
                    $i = 1
                    do {
                        Write-Verbose ("[$(Get-Date)] Saving $Outfile as a PDF, attempt $i")
                        # Launch Edge and open the Game Statistics Report temp file, and save it as a PDF
                        Start-Process 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -Wait -argumentlist "--headless=old --no-sandbox --run-all-compositor-stages-before-draw --virtual-time-budget=30000 --proxy-bypass-list=* --proxy-server= --safe-mode --print-to-pdf=""$Outfile"" --no-pdf-header-footer ""$TempFile""" -RedirectStandardOutput NUL
                        Start-Sleep $i
                        $i++
                    }
                    until
                    (
                        # PDF file is greater than 1KB in size; give up if we've already checked 5 times.
                        ((Test-Path $OutFile -PathType Leaf) -and ((Get-ChildItem $Outfile -ErrorAction SilentlyContinue).Length -gt 1024 -or $i -gt 5))
                    )
                    if ($i -gt 5) {
                        Write-Warning "There was a problem converting a Game Statistics Report to PDF. Please check the output file: $Outfile"
                    }
                    elseif ((!($NoViewPDF.IsPresent)) -and ($($psCmdlet.myinvocation.line) -notmatch "gss ")) {
                        # Open the generated PDF file in the default PDF viewer, if NoViewPDF parameter isnt set or we didn't run the command via the gss alias
                        Write-Verbose ("[$(Get-Date)] Attempting to open $Outfile in the default PDF Viewer...")
                        Start-Process $Outfile
                    }
                }
                if (($OpenBrowser.IsPresent) -or ($($psCmdlet.myinvocation.line) -match "gsw ")) {
                    # Opens in browser where you can save/print it
                    Write-Verbose ("[$(Get-Date)] Opening $($PlayerGameReport.URI) in the default browser...")
                    Start-Process $($PlayerGameReport.URI)
                }
            }
        }
    }
    End {
        if ($PlayerGameReportsData.Count -eq 0) {return}
        # create a PSPropertySet with the default property names.
        # This controls which properties are displayed on the console. so we can hide properties that we want for other objects in our pipeline, but not written out on screen
        [string[]]$visible = 'GameName','MID','CID','Result','ResultText'
        [Management.Automation.PSMemberInfo[]]$visibleProperties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet',$visible)

        # add the PSPropertySet to the PlayerGameReports onject, and send the list of reports to pipeline
        $PlayerGameReports | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $visibleProperties -PassThru
    }
}
