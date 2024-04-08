###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                             Playcheck Functions                             #
#                                     v1.6.3                                  #
#                                                                             #
###############################################################################

#Author: Chris Byrne - christopher.byrne@derivco.com.au



function Get-QFPlayCheck {
    <#
    .SYNOPSIS
        Generates a Play Check for the specified Player Login, CasinoID and TransactionID.

    .DESCRIPTION
        This cmdlet is used to generate a Play Check for the specified Player Login, CasinoID and TransactionID.
        You may optionally pass multiple TransactionIDs in a list and a Play Check will be generated for each one.
        The default web browser will open to display each Play Check report. This cmdlet generates no other output.

        You may optionally pass the '-SavePDF' switch parameter to automatically save the generated Play Check to a PDF file using Edge's 'Save as PDF' printer, then open in the default PDF Viewer program.
        Specifying the '-NoViewPDF' switch parameter will not open the PDF file, and just silently save the generated PDF without further user interaction.

        Running this cmdlet using the alias 'pcp' is equivalent to 'Get-QFPlaycheck -SavePDF'
        Running this cmdlet using the alias 'pcs' is equivalent to 'Get-QFPlaycheck -SavePDF -NoViewPDF'

    .EXAMPLE
        Get-QFPlayCheck -Login 12345 -CasinoID 54321 -TransID 100
        Generates a play check for a single transaction, and opens it in the default web browser.

    .EXAMPLE
        Get-QFPlayCheck -Login 12345 -CasinoID 54321 -TransID 100 -SavePDF
        Generates a play check for a single transaction and saves it as a PDF file int the current folder, without opening in a web browser.
        A list of game names, CIDs and MIDs for all games that a play check was generated for will be output to pipeline.

    .EXAMPLE
        Get-QFPlayCheck -Login 12345 -CasinoID 54321 -TransID 100,200,300
        Generates play checks for multiple transactions and opens in the default web browser.
        Transaction ID's are seperated by commas.

    .EXAMPLE
        Get-QFPlayCheck -Login 12345 -CasinoID 54321 -TransID (100..120)
        Play checks for multiple transactions in sequence (e.g. 20 TransID's from 100 to 120) and opens in the default web browser.
        Users PowerShell's Range operator to generate list of sequential numbers in the specified range.

    .EXAMPLE
        Get-QFPlayCheck -Login 12345 -CasinoID 54321 -TransID 100,200,300 -SavePDF -FileName "CasinoGame"
        Generates play checks for multiple transactions and saves as a PDF in the current folder.
        Transaction ID's are seperated by commas.
        
        The file name of each generated PDF will be in the format "CasinoGame_###.pdf"
        e.g. CasinoGame_100.pdf CasinoGame_200.pdf CasinoGame_300.pdf

    .PARAMETER CasinoID
        The CasinoID/ServerID that the specified Login belongs to.

    .PARAMETER FileName
        Specifies the File Name used when saving Play Checks to PDF.
        An underscore and the current TransactionID number will be appended to this filename.

        e.g. Setting this parameter to "CasinoPlayer" will results in filenames like:
        CasinoPlayer_1.pdf
        CasinoPlayer_2.pdf
        CasinoPlayer_3.pdf

        If not specified, the default file name is "PlayCheck" with a space and the current TransactionID number appended.

    .PARAMETER Hostname
        The Host Name for the Play Check site. Defaults to 'redirector3.valueactive.eu'

    .PARAMETER Login
        The Login Name of the player you wish to generate a play check for.
        The Play Check currently system doesn't currently support UserID's, you must specify the Login Name.

    .PARAMETER NoViewPDF
        if this switch is present, Play Check PDF's will not be opened in the default PDF viewer program automatically after they are created.
        This is useful if you are automating this function as it will not provide any other graphical output.
        This switch has no effect if the SavePDF switch is not also present.

    .PARAMETER OpenExplorer
        Opens a File Explorer window to the location where the PDF PlayCheck files are saved.

    .PARAMETER SavePDF
        if this switch is present, the function will use Edge's Print To PDF feature, to save the play check as a PDF in the current directory.
        The PDF file will then open in the default PDF viewer program.
        if this switch is not present, the play check will open in the default web browser for you to view and manually save.

    .PARAMETER TransID
        The player's Transaction ID's to generate Play Checks for. Passing multiple Transaction ID's in a list will generate a play check for each one.

    .INPUTS
        System.String
            You can pipe a string that contains a Casino Player LoginID to this cmdlet.

        System.Int32
            You can pipe an integer that contains a CasinoID/ServerID to this cmdlet.
            You can pipe an array of integers that contain Transaction ID numbers to this cmdlet. A Play Check Report will be generated for each Transaction ID number.
            You may alternatively pipe a single System.Int32 object containing a Transaction ID Number to play check a single round.

    .OUTPUTS
        If script runs in default mode (open playcheck in browser) no pipeline output is produced.

        If script runs in SavePDF mode, a list of any games for which a play check has been generated will be output to pipeline:

            System.Management.Automation.PSCustomObject

                Name        MemberType   Definition
                ----        ----------   ----------
                GameName    NoteProperty string
                CID         NoteProperty int
                MID         NoteProperty int
                ETI         NoteProperty bool

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function
    [CmdletBinding()]
    [alias("pc","pcp","pcs")]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Login,

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
        [int]$CasinoID,

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=2)]
        [string[]]$TransID,

        [Parameter(Mandatory=$false)]
        [switch]$SavePDF,

        [Parameter(Mandatory=$false)]
        [switch]$NoViewPDF,

        [Parameter(Mandatory=$false)]
        [string]$Hostname = "redirector3.valueactive.eu",

        [Parameter(Mandatory=$false)]
        [switch]$OpenExplorer,

        [Parameter(Mandatory=$false)]
        [string]$FileName

    )
    begin {

        # Required for URL Encode function
        Add-Type -AssemblyName System.Web

        # The address of the play check site. Strip http/s if provided and any additional path after the host name
        $Hostname = $Hostname.ToLower().trim() -replace "^.*://","" -replace "/.*$",""
        Write-Verbose ("[$(Get-Date)] Playcheck Site Hostname: $Hostname")

        # Check if MS Edge browser is installed
        if (!(Test-Path -Path 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -PathType Leaf)) {
            Throw "MS Edge browser must be installed to generate Play Checks. Please ensure it is installed properly."
        }

        # progressPreference set to silently continue, so test-netconnection and invoke-webrequest doesn't show the progress bar. Big speedup in some cases
        $global:progressPreference = 'silentlyContinue'

        # Confirm we have connectivity to the play check site
        try {
            Test-NetConnection $Hostname -port 443 -WarningAction Stop | Out-Null
        }
        catch {
            Write-Error "Unable to connect to the Play Check site: $Hostname - Please ensure you have internet connectivity."
            Return
        }
        # Array used to store game names, mids and cids before duplicates are removed and output to pipeline
        $PlayCheckGameArray = @()
    }
    process {
        Write-Host "Generating Play Checks, please wait...."
        # Remove any duplicate transaction IDs
        $TransID = $TransID | Sort-Object -Unique
        foreach ($Transaction in $TransID) {
            # Confirm Transaction IDs are integers
            # While we could have forced the TransID parameter in the function to only accept integers, allowing string lets us check each member if an array is passed.
            # If some of the members are integers we can still playcheck those. If we only accepted integers, the function would terminate as soon as it's invoked.
            # This is useful if you copy and paste a whole heap of transaction ID numbers and acceidently paste in some letters as well.
            $Transaction = $Transaction.ToString()
            if ($Transaction.trim() -notmatch "^[0-9]+$") {
                Write-Error "Transaction ID $Transaction must be a whole number."
                continue
            }
            # Buid the URI for the play check request
            $PlaycheckURI = "https://" + $Hostname + "/casino/default.aspx?applicationID=1001&ServerID=" + $CasinoID + "&username=" + [System.Web.HttpUtility]::UrlEncode($Login.trim()) + "&ssousername=&password=PTS_ADMIN&lang=en&SessionID=&TransactionID=" + $Transaction.trim() + "&ssopassword=&usertype=0&ssologintype=0&requestedlogin=0&accounttype=0&Clienttype=5&Directx=0&PCMGUID="
            Write-Verbose ("[$(Get-Date)] Transaction ID: $Transaction")
            Write-Verbose ("[$(Get-Date)] Playcheck URI: $PlayCheckURI")

            # if SavePDF parameter was passed, or command was run with alias 'pcp' or 'pcs', print to PDF automatically without displaying in browser
            if (($SavePDF.IsPresent) -or ($($psCmdlet.myinvocation.line) -match "^pc[ps]")) {
                # Open the playcheck silently in Edge, and dump the DOMcontaining the HTML source into a temp file.
                # This is required to reformat the play check so all content appears on one page.
                # Basically the same process as the Format-QFPlaycheck function but have integrated it into this function as its only a few regexes to find and replace strings in the HTML source of the page.
                # Can't use Invoke-WebRequest for this as the scripts on the page won't run.
                $TempFile = "$Env:Temp\PlayCheckTemp$Transaction.html"

                # It can sometimes take a couple of goes to generate the play check, try 4 times before giving up
                $i = 1
                do {
                    # Clear any existing temp file
                    if (Test-Path $TempFile -PathType Leaf) {
                        Remove-Item $TempFile -Force
                    }
                    Write-Verbose ("[$(Get-Date)] Calling Edge to dump PlayCheck to temporary file: $TempFile - Attempt $i")
                    try {
                        Start-Process -WindowStyle Hidden -FilePath 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -Wait -argumentlist "--headless --dump-dom --no-sandbox --run-all-compositor-stages-before-draw --virtual-time-budget=30000 --proxy-bypass-list=* --proxy-server= --user-agent=""Mozilla/5.0 (Windows NT 10.0; Win64; x64)"" --safe-mode ""$PlaycheckURI""" -RedirectStandardOutput "$TempFile" -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Error calling Edge to retrieve PlayCheck (Transaction ID: $Transaction) and saving to a temporary file: $TempFile"
                        continue
                    }
                    Write-Verbose ("[$(Get-Date)] Size of the TEMP file: $((Get-ChildItem $TempFile -ErrorAction SilentlyContinue).Length) ")
                    $i++
                }
                Until
                (
                    # Check a temp file was created with some data in it, at least 1kb, or we've made 4 attempts
                    # Also check we didn't just get a session expired error page, sometimes it needs to re-auth for some reason
                    (((Test-Path $TempFile -PathType Leaf) -and `
                    ((!(Get-Content -Encoding UTF8 $TempFile -ErrorAction SilentlyContinue|select-string -pattern "Session Expired")) -and `
                    ((Get-ChildItem $TempFile -ErrorAction SilentlyContinue).Length -gt 1024))) -or `
                    ($i -gt 4))
                )

                if ($i -gt 4) {
                    Write-Warning "Error retrieving PlayCheck (Transaction ID: $Transaction) and saving to a temporary file after $i attempts - Skipping this playcheck."
                    # Finally, Clear up any existing temp file
                    if (Test-Path $TempFile -PathType Leaf) {
                        Remove-Item $TempFile -Force
                    }
                    continue
                }

                # Now get the temp file into a variable so we can process it
                try {
                    $PlaycheckTemp = Get-Content $TempFile -Encoding UTF8 -ErrorAction Stop -Raw
                }
                catch {
                    Write-Warning "Error reading PlayCheck data (Transaction ID: $Transaction) from the temporary file: $TempFile"
                    continue
                }

                # Basic check to see if its really playcheck data in the temp file
                if (!($PlaycheckTemp -like "*/playcheck/*" -and $PlaycheckTemp -like "*<html*")) {
                    Write-Warning "Doesn't look like a valid playcheck file for Transaction ID: $Transaction - This playcheck will be skipped."
                    Write-Warning "Playcheck URI: $PlaycheckURI"
                    continue
                }

                # Regex to get the current host server in Matches[2], this will vary based on player site; also checks if we were redirected to an error page in Matches[1]
                # Also gets the MID and CID of the game in $Matches[3] and $Matches[4] respectively
                $PlaycheckTemp | Select-String -Pattern "trackPageView" | ForEach-Object { $_ -match 'trackPageView\(.*(https?://([^/?#]*).+?),.*\{.*ModuleId.:(\d+),.*?ClientId.:(\d+),.*\}'} | Out-Null
                if ($null -ne $Matches[4]) { [int]$PlaycheckGameCID = $Matches[4] ; Write-Verbose ("[$(Get-Date)] Found Game CID: $PlaycheckGameCID")}
                if ($null -ne $Matches[3]) { [int]$PlaycheckGameMID = $Matches[3] ; Write-Verbose ("[$(Get-Date)] Found Game MID: $PlaycheckGameMID")}
                $PlayCheckHost = $Matches[2]
                $PlayCheckRedir = $Matches[1]

                if ($PlayCheckRedir -match "(Failed|Error)") {
                    Write-Warning "Received an error when generating PlayCheck Transaction ID: $Transaction  - Please confirm player LoginID, CasinoID and Transaction Numbers are valid."
                }

                # try to read the game name from the page header into $Matches[1]
                $PlaycheckTemp | Select-String -Pattern "<h1>.*</h1>" | ForEach-Object { $_ -match '<h1>(.*)</h1>'} | Out-Null
                if ($null -ne $Matches[1] -and $Matches[1] -ne "") {
                    $PlaycheckGameName = $Matches[1].Trim()
                    Write-Verbose ("[$(Get-Date)] Found Game Name: $PlaycheckGameName")
                }

                # Run the replace regex to edit the temp file HTML source - this reformats the play check so all elements are visible on one page and overwrites the temp file
                $PlaycheckTemp = $PlaycheckTemp -replace '<div id(=|=3D)"event_(.*)" style(=|=3D)"display: ?none;?" ?>','<div id$1"event_$2" style$3"display: inline;">' # Unhide all extra page sections, so all content on one page
                $PlaycheckTemp = $PlaycheckTemp -replace '<div id(=|=3D)"event_(.*)" style(=|=3D)"display: ?none;?" ?>','<div id$1"event_$2" style$3"display: inline;">' # Run this regex twice in case it misses some lines... if you have a better solution let me know!
                $PlaycheckTemp = $PlaycheckTemp -replace '(href|src)="(/.*)"', $('$1="https://' + $PlayCheckHost + '$2"') # Insert the Playcheck Host at the start of any relative links, otherwise images etc will break when Edge loads the file
                $PlaycheckTemp = $PlaycheckTemp -replace '(?mi)<img src="\.\./images/Cards/(.+?)">', $('<img src="https://' + $PlayCheckHost + '/playcheck/Home/GameDetail/images/Cards/$1">') # Insert Playcheck host and full path in relative links for Card images
                $PlaycheckTemp = $PlaycheckTemp -replace '(<script.*GameDetails.*\/script>)','<!-- $1 -->' # Comment out the GameDetails script as this will hide all the pages again when the playcheck loads in Edge
                $PlaycheckTemp = $PlaycheckTemp -replace '<img src="\.\./images/.*\.gif">','' # Remove the + Expand image for extra reel positions/bonus rounds; we will expand these sections so this image isnt needed
                $PlaycheckTemp = $PlaycheckTemp -replace '<div class(=|=3D)"controls bottom" ?>','<div class$1"controls bottom" style="display:none">' # Hide the page controls at the bottom of the page
                $PlaycheckTemp = $PlaycheckTemp -replace '<div class(=|=3D)"(ToggleContent|collapse)" style(=|=3D)"display: ?none;? ?" ?>','<div class$1"$2" style$3"display: inline;">' # Expands sections for extra reel positions/bonus rounds e.g. Reel Position 2 in 'Thunderstruck Stormchaser'
                
                # ETI checks - Supported ETI games have an additional ETI script not present in regular games
                if ($PlaycheckTemp -like "*/Playcheck/Scripts/eti?v=*") {
                    $ETIGame = $true
                    # regex to get the URI for the ETI provider's playcheck
                    $PlaycheckTemp -imatch '(?si)<div class="etiGameDetails">.*<iframe.*src="(.+?)".*</iframe' | Out-Null
                    # URI will be in $Matches[1], Convert $amp; to & - required for Hacksaw URI to work
                    $ETIPlaycheckURI = $Matches[1] -replace "&amp;","&"
                    Write-Verbose ("[$(Get-Date)] ETI game detected! Playcheck URI: $ETIPlaycheckURI")

                    # Hacksaw or MGA ETI Games - insert link to replay video into the HTML
                    If ($ETIPlaycheckURI -like '*hacksawgaming*' -or $ETIPlaycheckURI -like '*mgagamesmicrogaming*') {
                        Write-Verbose ("[$(Get-Date)] ETI game with video replay, will insert the video replay link into the playcheck")
                        $PlaycheckTemp = $PlaycheckTemp -replace '(?si)(<div class="etiGameDetails">.*<h2>Details</h2>)',$('$1<div><a href="' +
                        $ETIPlaycheckURI + '">A visual replay of this round is available. Click here to view.</a>')
                    } elseif ($ETIPlaycheckURI -like '*api-rgs.oryxgaming*') {
                        # Oryx gaming - fix for incorrect URI and resize iframe
                        if ($PlaycheckTemp -match '(?si)<div class="etiGameDetails">.*<iframe.*src="(.+?)".*style="display: none">.*</iframe') {
                            # Oryx has multiple domains - api-rgs Malta/Germany; api-r3hr for NL and others; api-rghr for Curacao (not sure if we have any operators there)
                            # See if the alternative domain works and replace the iframe URI if so
                            Try {
                                $ETIPlaycheckURI = $ETIPlaycheckURI -replace 'api-rgs','api-r3hr'
                                Write-Verbose ("[$(Get-Date)] ETI Oryx Gaming - trying alternative playcheck URI: $ETIPlaycheckURI")
                                Invoke-WebRequest $ETIPlaycheckURI | Out-Null
                                $PlaycheckTemp = $PlaycheckTemp -replace '(?si)<iframe(.+?)src=".+?"(.+?)style="display: none">',('<iframe$1src="' + $ETIPlaycheckURI +'"$2style="display: block; height: 1800px;">')
                            } catch {
                                Write-Verbose ("[$(Get-Date)] Oryx alternative URI failed, could not get a valid Play Check for this round.")
                            }
                        } else {
                            # If the iframe didn't have style="display: none" then the original URI worked, but Oryx playchecks still don't seem to resize properly. Force the iframe to a specific height.
                            $PlaycheckTemp = $PlaycheckTemp -replace '(?si)<iframe(.+?)style="display: block;">',('<iframe$1style="display: block; height: 1800px;">')
                        }
                        # Insert a link to the Oryx playcheck page in case it gets cut off
                        $PlaycheckTemp = $PlaycheckTemp -replace '(?si)(<div class="etiGameDetails">.*<h2>Details</h2>)',$('$1<div><a href="' +
                        $ETIPlaycheckURI + '">Click here to view detailed game round results.</a>')
                    }
                } else {
                    $ETIGame = $false
                }

                # Some ETI games aren't supported for playcheck at all, or if game round incomplete/playcheck is over 40 days old, skip PDF creation for these
                if ($PlaycheckTemp -like "*Sorry, this game is not supported*") {
                    $ETIGame = $true
                    Write-Warning "TransactionID: $Transaction - $PlaycheckGameName (MID: $PlaycheckGameMID CID: $PlayCheckGameCID) is not supported for Playcheck."
                } elseif ($PlaycheckTemp -like "*The detailed result cannot be displayed due to incomplete games*") {
                    Write-Warning "TransactionID $Transaction is still open pending free spins or a similar feature - player needs to return to the game and complete this round before playcheck will be available."
                } elseif ($PlaycheckTemp -like "*This transaction is older than 40 days and has been archived*") {
                    Write-Warning "TransactionID $Transaction is older than 40 days and has been archived. Playcheck is no longer available for this round."
                } elseif ($PlaycheckTemp -match '(?si)<div class="etiGameDetails">.*<iframe.*style="display: none".*</iframe') {
                    # The ETI playcheck iframe didn't load correctly
                    Write-Warning "TransactionID: $Transaction - $PlaycheckGameName (MID: $PlaycheckGameMID CID: $PlayCheckGameCID) - Failed to get a valid playcheck for this ETI game, please refer to the ETI provider's support team."
                } else {
                    # Write the modified HTML back to the temp file - encoding UTF8 otherwise you get weird symbols
                    $PlaycheckTemp | Set-Content $TempFile -Encoding UTF8
                     
                    # Our temp file should now be ready to convert to PDF.
                    # Check that the PDF is generated correctly... sometimes its a blank page, check file size and retry until its big enough to have some content. try 5 times, waiting a bit longer each time
                    $i = 1
                    do {
                        if ($FileName -ne "" -and $FileName -ne $null) {
                            $Outfile = (Get-Location).path + "\$FileName" + "_" + $Transaction + ".pdf"
                        }else{
                            $Outfile = (Get-Location).path + "\PlayCheck " + $Transaction.trim() + ".pdf"
                        }
                        
                        Write-Verbose ("[$(Get-Date)] Saving $Outfile as a PDF, attempt $i")
                        # Launch Edge and open the Playcheck temp file, save it as PDF
                        Start-Process 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -Wait -argumentlist "--headless --no-sandbox --run-all-compositor-stages-before-draw --virtual-time-budget=30000 --proxy-bypass-list=* --proxy-server= --safe-mode --print-to-pdf=""$Outfile"" --no-pdf-header-footer --user-agent=""Mozilla/5.0 (Windows NT 10.0; Win64; x64)"" ""$TempFile""" -RedirectStandardOutput NUL
                        Start-Sleep $i
                        $i++
                    }
                    until (
                        ((Test-Path $OutFile -PathType Leaf) -and ((Get-ChildItem $Outfile -ErrorAction SilentlyContinue).Length -gt 1024 -or $i -gt 5))
                    )
                    if ($i -gt 5) {
                        Write-Warning "There was a problem converting playcheck $Transaction to PDF. Please check the output file: $Outfile"
                    }
                    elseif ((!($NoViewPDF.IsPresent)) -and ($($psCmdlet.myinvocation.line) -notmatch "^pcs")) {
                        # Open the generated PDF file in the default PDF viewer, if NoViewPDF parameter isnt set or we didn't run the command via the pcs alias
                        Write-Verbose ("[$(Get-Date)] Attempting to open $Outfile in the default PDF Viewer...")
                        Start-Process $Outfile
                        
                        
                    }
                }

                # Add the game name, MID and CID to array for output to pipeline at end of function.
                if ($null -ne $PlaycheckGameCID -and $null -ne $PlaycheckGameMID -and $null -ne $PlaycheckGameName) {
                $PlayCheckGameArray += [PsCustomObject]@{Id = $($PlayCheckGameArray.Count + 1); GameName = $PlaycheckGameName; MID = $PlaycheckGameMID; CID = $PlaycheckGameCID; ETI = $ETIGame}
                }
                
            } else {
                # Opens the playcheck in the default browser where you can save/print it
                Write-Verbose ("[$(Get-Date)] Opening Play Check $Transaction in default web browser...")
                Start-Process $PlaycheckURI
                Continue
            }
            # Finally, Clear up any existing temp file
            if (Test-Path $TempFile -PathType Leaf) {
                Remove-Item $TempFile -Force
            }
        }
        if($OpenExplorer.IsPresent){
            explorer.exe (Get-Location).path
        }
    }
    end {
        # if we added any game info to our array, check and remove any duplicates then output to pipeline
        if ($PlayCheckGameArray.count -gt 0) {
            Write-Verbose ("[$(Get-Date)] Current Temp Game Array record count: $($PlayCheckGameArray.count)")
            $PlayCheckGameArray | Select-Object GameName,MID,CID,ETI -unique
        }
    }
}


function Format-QFPlayCheck {
    <#
    .SYNOPSIS
        Converts a Play Check with multiple pages into a single page.

    .DESCRIPTION
        Play Checks for some games are created with multiple pages, that can be selected from a control element at the bottom of the page.
        This makes it cumbersome to save to PDF as you must select and save each individual page.
        This function will modify a play check saved as MHTML to remove the page selector, expand any collapsed/minimised sections and display all content on one page.
        It will also convert the file to PDF, deleting the source file.

        To use it:
        -Generate the play check normally in your web browser
        -Save the play check as 'Web Page - Single File (*.mhtml)'
        -Run this function and specify the file name

    .EXAMPLE
        Format-QFPlaycheck PlayCheck.mhtml

    .EXAMPLE
        Get-ChildItem Playcheck*.mhtml | Format-QFPlaycheck

    .PARAMETER NoViewPDF
        Does not open the saved PDF file in the PDF viewer. This is useful if this function is called by an automated process as it requires no user interaction.

    .PARAMETER Path
        The file name of the Play Check (in HTML format) to modify. Supports wildcards.

    .INPUTS
        System.String
            You can pipe a string to this cmdlet that contains a path to a saved Play Check MHTML file.

    .OUTPUTS
        This cmdlet does not provide any pipeline output.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    #>

    # Set up parameters that can be passed to the function.
    # Parameter Set Path = literal filenames passed at command line. default parameter set.
    # Parameter Set LiteralPath = filenames passed through pipeline (any object with a property of PSPath) eg Get-ChildItem
    # Only one parameter set can be used at a time
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [alias("fpc")]
    param(
        [Parameter(
            Mandatory,
            ParameterSetName  = 'Path',
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]$Path,
        [Parameter(
            Mandatory,
            ParameterSetName = 'LiteralPath',
            Position = 0,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [string[]]$LiteralPath,

        [Parameter(Mandatory=$false)]
        [switch]$NoViewPDF
    )
    begin {
           # Check if MS Edge browser is installed - required for converting into PDF
        if (!(Test-Path -Path 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -PathType Leaf)) {
            Throw "MS Edge browser must be installed to generate Play Checks. Please ensure it is installed properly."
            }
    }
    process {
        # Resolve path(s) based on the parameter set
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $ResolvedPaths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            $ResolvedPaths = (Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path).replace("Microsoft.PowerShell.Core\FileSystem::","")
        }
        # Process each item in resolved paths, check it exists
        foreach ($SourceFile in $ResolvedPaths) {
            if (!(Test-Path -Path $SourceFile -PathType Leaf)) {
                Write-Host "Cannot find playcheck file: $SourceFile"
                continue
            }

            Write-Verbose ("[$(Get-Date)] MHTML File: $SourceFile")
            # Check the source file isn't locked by trying to get a read/write handle on it
            try {
                $FileStream = [System.IO.File]::Open($SourceFile,'Open','Write')
                $FileStream.Close()
                $FileStream.Dispose()
            }
            catch {
                Write-Host "Unable to open playcheck file $SourceFile ! Ensure it is not open in another program."
                continue
            }

            # Get the source file content
            $FileContent = (Get-Content -raw -Encoding UTF8 -Path $SourceFile)

            # Basic check to see if its really a MHTML playcheck file
            if (!($FileContent -like "*/playcheck/*" -and $FileContent -like "*<html*")) {
                Write-Host "$SourceFile doesn't look like a valid playcheck file. This file won't be modified."
                continue
            }

            # Run the replace regex to edit the MHTML file - this reformats the play check so all elements are visible on one page and overwrites the source file
            $FileContent = $FileContent -replace '<div id(=|=3D)"event_(.*)" style(=|=3D)"display: ?none;?" ?>','<div id$1"event_$2" style$3"display: inline;">'  # Makes every page visible
            $FileContent = $FileContent -replace '<img src="\.\./images/.*\.gif">','' # Removes the + Expand image for additional rounds or bonuses
            $FileContent = $FileContent -replace '<div class(=|=3D)"controls bottom" ?>','<div class$1"controls bottom" style="display:none">' # Hides the page controls on the bottom of the page
            $FileContent = $FileContent -replace '<div class(=|=3D)"(ToggleContent|collapse)" style(=|=3D)"display: ?none;? ?" ?>','<div class$1"$2" style$3"display: inline;">' # Expands the additional rounds or bonuses e.g. Reel Position 2 in 'Thunderstruck Stormchaser'
            $FileContent | Set-Content -Encoding UTF8 -Path $SourceFile

            # Call Edge to convert to PDF
            $Outfile = $SourceFile.trim() -replace  ("\..*$",".pdf")
            Write-Verbose ("[$(Get-Date)] Saving as a PDF - Filename: $OutFile")
            Start-Process 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' -wait -argumentlist "--headless --disable-gpu --no-sandbox --run-all-compositor-stages-before-draw --virtual-time-budget=30000 --proxy-bypass-list=* --proxy-server= --safe-mode --print-to-pdf=""$Outfile"" --no-pdf-header-footer ""$SourceFile"""

            # try to delete the source MHTML file, attempt 10 times with 1 sec delay between each attempt
            $i = 1
            do {
                try {
                    # Check the source file isn't locked which will prevent it getting deleted
                    Start-Sleep 1
                    $FileStream = [System.IO.File]::Open($OutFile,'Open','Write')
                    $FileStream.Close()
                    $FileStream.Dispose()
                    if (Test-Path -Path $OutFile -PathType Leaf) {
                        # Delete the old MHTML file
                        Write-Verbose ("[$(Get-Date)] Deleting the source file: $SourceFile Attempt $i")
                        Remove-Item $SourceFile -Force
                        # Set i to 999 so it exits the do loop and won't print the error message
                        $i = 999
                    }
                }
                catch {
                    $i++
                }
            }
            until
            (
                ($i -gt 10)
            )
            if ($i -gt 10 -and $i -lt 999) {
                Write-Host "Unable to delete the source playcheck file: $SourceFile"
            }

            # Open the saved PDF file in the default viewer
            if (!($NoViewPDF.IsPresent)) {
                Write-Verbose ("[$(Get-Date)] Opening $OutFile in the default PDF viewer...")
                Start-Process $Outfile
            }
        }
    }
}
