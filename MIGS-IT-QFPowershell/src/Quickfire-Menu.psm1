###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                            Invoke-QFMenu Function                           #
#                                    v1.6.4                                   #
#                                                                             #
###############################################################################

# Author: Chris Byrne - christopher.byrne@derivco.com.au

# Note - if you are looking for the New-QFTicket or 'zz' menu, that is located in the file Quickfire.psm1


function Invoke-QFMenu {
    <#
    .SYNOPSIS
        Presents an interactive menu for requesting information from Casino Portal or Reconciliation API.

    .DESCRIPTION
        This cmdlet presents an interactive menu for requesting information from the Casino Portal API or Reconciliation API.
        For example, you can search for a Quickfire Casino via CasinoID or Casino Name, or list all casinos belonging to an OperatorID.
        You can search for Quickfire Games by ModuleID/ClientID or Game Name.
        You can also perform Reconciliation API functions such as checking for queued transactions and unlocking them.

        A number of optional parameters are provided that allow you to set default values for various options. 
        This includes CasinoID, UserID, Player Login names, etc. 
        The user will be prompted to enter a value or simply press ENTER to use the default.
        These parameters are useful if this menu is invoked from another function; you can offer default values instead of making the user enter them manually.

    .PARAMETER CasinoID
        Optional parameter that specifies a default value for the Casino ID. 
        The user will be asked to enter a CasinoID or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.

    .PARAMETER CasinoID
        Optional parameter that specifies a default value for the Casino ID. 
        When performing any operations relating to a Casino, the user will be asked to enter a CasinoID or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.

    .PARAMETER HostingSiteID
        The ID Number of the Hosting Site for the specified CasinoID. This is required for Reconciliation API functions.

        Quickfire Hosting Site ID's are:
            2   Malta (MAL)
            3   Canada (MIT)
            9   Gibralta (GIC)
           25   Croatia (CIL)
           29   IOA Staging Environment

        Note that there is no distinction between different systems at each site. i.e. MAL1, MAL2, and MAL3 systems all have the same HostingSiteID of 2.

    .PARAMETER Login
        Optional parameter that specifies a default value for the player Login. 
        When searching for a player, the user will be asked to enter a Login or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.

    .PARAMETER OperatorID
        Optional parameter that specifies a default value for the Operator ID. 
        When performing any operations relating to a Casino, the user will be asked to enter an Operator ID or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.

    .PARAMETER REQNumber
        Optional parameter that specifies a default value for the ticket REQ Number. This is currently only used by Reconciliation API functions.
        When unlocking a transaction, the user will be asked to enter a Transaction Reference or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.
    
    .PARAMETER UserId
        Optional parameter that specifies a default value for the player UserID. This is mainly used by Reconciliation API functions.
        When performing operations relating to a player, the user will be asked to enter a UserId or simply press ENTER to use the default value.
        If this parameter is not set, no default value will be provided to the user.

    .EXAMPLE
        Invoke-QFMenu
            Displays an interactive menu of API functions. 

    .EXAMPLE
        Invoke-QFMenu -CasinoID 12345
            Displays an interactive menu of API functions, with a default Casino ID value of '12345'.
            This default value is displayed to the user to accept or overwrite with a different value.

    .INPUTS
        This cmdlet accepts pipeline input for the various parameters such as CasinoID or OperatorID.

    .OUTPUTS
        A PSCustomObject consisting of the output from the Casino Portal API.
        This output will vary depending on the parameters provided to this cmdlet.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

        When the user selects option Q to exit, the global object $global:QFMenuExitFlag will be set to True.
        Parent functions calling this function can check this object and if true, can stop execution completely.
        Otherwise the function will simply return to the calling function as per normal behaviour of the Return keyword.

    .LINK
        https://casinoportal.gameassists.co.uk/api/swagger/index.html
        https://reviewdocs.gameassists.co.uk/internal/document/ExternalOperators/Reconciliation%20API/1

    #>

    # Set up parameters for this function
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [alias("qfm","qfmenu")]
    param (

    # The default CasinoID value to offer to the user
    [Parameter(Position = 3, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$CasinoID,

    # The HostingSiteID - required for ReconAPI
    [Parameter(Position = 5, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$HostingSiteID,

    # The default Player Login value to offer to the user
    [Parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Login,

    # The default OperatorID value to offer to the user
    [Parameter(Position = 4, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$OperatorID,

    # The REQNumber of the current ticket. Will be used as the reference for Transaction Unlock operations.
    [Parameter(Position = 0, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$REQNumber,

    # The default UserID value to offer to the user
    [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [int]$UserID
    )

    begin {

        function Invoke-MainMenu {
            # Local function to display the main menu
            :MainMenu do {
                $MainMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$(If ($REQNumber){"$REQNumber - "})QuickFire API functions$([char]27)[0m", "Select an operation from the menu. Press Q to quit.", $MainMenu, 0)
                Write-Host ""
                Write-Verbose ("[$(Get-Date)] Main Menu option selected: $MainMenuChoice")
                switch ($MainMenuChoice) {
                    1 { # Casino info menu
                        Invoke-CasinoMenu
                    }
                    2 { # Game search menu
                        Invoke-GameSearchMenu
                    }
                    3 { # User search menu
                        Invoke-UserSearchMenu
                    }
                    4 { # Recon API
                        Invoke-ReconciliationMenu
                    }
                    5 {
                        Invoke-AAMSMenu
                    }
                    6 { # Quit - should only appear if run from inside another function e.g. New-QFTicket
                        if ($PSCmdlet.MyInvocation.CommandOrigin -eq "Internal") {
                            # Set this flag so any parent function knows to exit
                            $global:QFMenuExitFlag = $True
                        }
                    }
                }
            }
            until ($MainMenuChoice -eq 0 -or $global:QFMenuExitFlag -eq $True)
        }
        

        function Invoke-CasinoMenu {
            # Local function to display the Casino search menu and prompt for input
            
            # Sets up the PromptForChoice method for the Casino Info menu
            $CasinoMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&CasinoID',"Display details for the specified CasinoID"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Casino &Name', "Search for casinos matching the specified Name"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&OperatorID', "Display details for the specified OperatorID"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'OpSec &Passwords', "Display Operator Security credentials for the specified OperatorID"))
            $CasinoMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Website Passwords', "Display Casino Website credentials (e.g. test player accounts) for the specified CasinoID"))

            :CasinoMenu do {
                $CasinoMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$(If ($REQNumber){"$REQNumber - "})QuickFire Casino Info lookup$([char]27)[0m", "Select an operation from the menu. Press Q to quit.", $CasinoMenu, 0)
                Write-Host ""
                Write-Verbose ("[$(Get-Date)] Casino Menu option selected: $CasinoMenuChoice")
                switch ($CasinoMenuChoice) {
                    1 { # Quit completely
                        # Set this flag so any parent function knows to exit
                        $global:QFMenuExitFlag = $True
                        Return
                    }
                    2 { # Casino ID
                        Invoke-CasinoSearch -ByCasinoID
                    }
                    3 { # Casino Name
                        Invoke-CasinoSearch -ByCasinoName
                    }
                    4 { # OperatorId
                        Invoke-CasinoSearch -ByOperatorID
                    }
                    5 { # Operator Security Passwords
                        Get-OpsecCreds
                    }
                    6 { # Operator Security Passwords
                        Get-WebsiteCreds
                    }
                }
            } until ($CasinoMenuChoice -eq 0 -or $global:QFMenuExitFlag -eq $True)
        }


        function Invoke-GameSearchMenu {
            # Local function to display Game Search menu and prompt for input

            # Sets up the PromptForChoice method for the Game Search menu
            $GameMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $GameMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
            $GameMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
            $GameMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&MID/CID',"Display games by ModuleID and optionally ClientID"))
            $GameMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Game &Name', "Search for casinos matching the specified Name"))

            :GameSearchMenu do {
                $GameMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$(If ($REQNumber){"$REQNumber - "})QuickFire Game Search$([char]27)[0m", "Select an operation from the menu. Press Q to quit.", $GameMenu, 0)
                Write-Host ""
                Write-Verbose ("[$(Get-Date)] Menu option selected: $GameMenuChoice")
                switch ($GameMenuChoice) {
                    1 { # Quit completely
                        # Set this flag so parent function knows to exit
                        $global:QFMenuExitFlag = $True
                        Return
                    }
                    2 { # Search By MID and CID
                        Invoke-GameSearch -ByMIDCID
                    }
                    3 { # Search By Name
                        Invoke-GameSearch -ByName
                    }
                }
            } until ($GameMenuChoice -eq 0 -or $global:QFMenuExitFlag -eq $True)
        }


        function Invoke-UserSearchMenu {
            # Local function to display the user search menu and prompt for input.

            # Sets up the PromptForChoice method for the User Search menu
            $UserSearchMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $UserSearchMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
            $UserSearchMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
            $UserSearchMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&CasinoID',"Display details for the specified CasinoID"))
            $UserSearchMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Casino &Name', "Search for casinos matching the specified Name"))
            $UserSearchMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&OperatorID', "Display details for the specified OperatorID"))

            :UserSearchMenu do {
                $UserSearchMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$(If ($REQNumber){"$REQNumber - "})QuickFire User Search$([char]27)[0m", "Select an operation from the menu. Press Q to quit.", $UserSearchMenu, 0)
                Write-Host ""
                Write-Verbose ("[$(Get-Date)] User Search Menu option selected: $UserSearchMenuChoice")
                switch ($UserSearchMenuChoice) {
                    1 { # Quit completely
                        # Set this flag so parent function knows to exit
                        $global:QFMenuExitFlag = $True
                        Return
                    }
                    2 { # Search By Casino ID
                        
                        Try {
                            # Request the CasinoID
                            $CasinoIDInput = Get-CasinoID $script:CasinoID
                            # Request the user Login to search for
                            $UserLoginInput = Get-Login $script:Login
                            $Params = @{
                                CasinoID  = $CasinoIDInput
                                Login = $UserLoginInput
                            }
                            $UserSearch = Search-QFUser @Params
                            If ($Null -eq $UserSearch -or $UserSearch.Count -eq 0) {Throw "No Users found matching the specified criteria."}
                            $UserSearch | Format-List
                            # Set OperatorID and Login to the input value so we can offer it as a default for the next request
                            $script:Login = $UserLoginInput
                            $script:CasinoID = $CasinoIDInput
                        } Catch {
                            Write-Warning $_.Exception.Message
                        }
                    }
                    3 { # Casino Name
                        Try {
                            [string]$CasinoName = Read-Host -Prompt "Please enter the Casino Name to search for"
                            If ($Null -eq $CasinoName -or $CasinoName.trim().Length -lt 3) {
                                Throw "Please enter a valid Casino Name to search for of at least 3 characters in length."
                            }
                            # Request the user Login to search for
                            $UserLoginInput = Get-Login $script:Login
                            $Params = @{
                                CasinoName  = $CasinoName
                                Login = $UserLoginInput
                            }
                            $UserSearch = Search-QFUser @Params
                            If ($Null -eq $UserSearch -or $UserSearch.Count -eq 0) {Throw "No Users found matching the specified criteria."}
                            $UserSearch | Format-List
                        } Catch {
                            Write-Warning $_.Exception.Message
                        }
                        # Set Login to the input value so we can offer it as a default for the next request
                        $script:Login = $UserLoginInput
                    }
                    4 { # OperatorId
                        Try {
                            # offer OperatorID parameter as default option if present
                            $OperatorIDInput = Get-OperatorID $script:OperatorID
                            # Request the user Login to search for
                            $UserLoginInput = Get-Login $script:Login
                            $Params = @{
                                OperatorID = $OperatorIDInput
                                Login = $UserLoginInput
                            }
                            $UserSearch = Search-QFUser @Params
                            If ($Null -eq $UserSearch -or $UserSearch.Count -eq 0) {Throw "No Users found matching the specified criteria."}
                            $UserSearch | Format-List
                            # Set OperatorID and Login to the input value so we can offer it as a default for the next request
                            $script:OperatorID = $OperatorIDInput
                            $script:Login = $UserLoginInput
                        } Catch {
                            Write-Warning $_.Exception.Message
                        }
                    }
                }
            } until ($UserSearchMenuChoice -eq 0 -or $global:QFMenuExitFlag -eq $True)

        }


        function Invoke-ReconciliationMenu {
            # Local function to display the Reconciliation menu and prompt for input
            
            # Sets up the PromptForChoice method for the Casino Info menu
            $ReconMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription 'Round &Status',"Display detailed status for the specified TransactionID"))
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Vanguard Queues', "Display Commit and Rollback queues for a Casino or Player"))
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Unlock', "Commit or Rollback a queued Transaction"))
            $ReconMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Range Mode', "Range Mode calls Reconciliation API for all Transaction ID's between two provided numbers. The default mode lets you specify non-sequential Transaction ID's."))

            :ReconMenu do {
                $ReconMenuChoice = $Host.UI.PromptForChoice("$([char]27)[36m$(If ($REQNumber){"$REQNumber - "})Reconciliation API Functions$(if ($script:ReconRangeMode) {"$([char]27)[32m - Range Mode Enabled"})$([char]27)[0m", "Select an operation from the menu. Press Q to quit.", $ReconMenu, 0)
                Write-Host ""
                Write-Verbose ("[$(Get-Date)] Recon Menu option selected: $ReconMenuChoice")
                switch ($ReconMenuChoice) {
                    1 { # Quit completely
                        # Set this flag so any parent function knows to exit
                        $global:QFMenuExitFlag = $True
                        Return
                    }
                    2 { # Round Status
                        Invoke-RoundStatus
                    }
                    3 { # Vanguard Queues
                        Invoke-VanguardQueues
                    }
                    4 { # Unlock Transactions
                        Invoke-TransactionUnlock
                    }
                    5 {
                        # Range Mode Toggle
                        if ($script:ReconRangeMode) {
                            $script:ReconRangeMode = $false
                        } else {
                            $script:ReconRangeMode = $true
                        }
                    }
                }
            } until ($ReconMenuChoice -eq 0 -or $global:QFMenuExitFlag -eq $True)
        }


        function Invoke-ItemMenu {
            # Local function to assist displaying multiple returned items with lots of data, such as casinos and games.
            param(
                # Counter of items to loop through
                [Parameter(Mandatory = $true)]
                [int]$Counter,

                # Text label to display in the menu
                [Parameter(Mandatory = $true)]
                [ValidateSet("Game","Casino","OpSec Credential","Website Credential","Transaction","AAMS Session",ErrorMessage="Invalid item specified.")]
                [string]$Item,

                # Max number of items
                [int]$ItemCount = 1

            )
            # Display menu prompting user for action
            # Hide next/back/all options depending on how many items are available
            If ($Counter + 1 -lt $ItemCount) {
                Write-Host "$([char]27)[36m[SPACE/`u{2192}]$([char]27)[0m Next $Item`t" -NoNewline
            }
            If ($Counter -gt 0) {
                Write-Host "$([char]27)[36m[B/`u{2190}]$([char]27)[0m Back/Previous $($Item)`t" -NoNewline
            }
            If ($Counter + 1 -lt $ItemCount) {
                Write-Host "$([char]27)[36m[A]$([char]27)[0m Display ALL remaining $($Item)s`t" -NoNewline
            }
            Write-Host ""
            Write-Host "$([char]27)[36m[X]$([char]27)[0m Export to Excel`t" -NoNewline
            # Other options to display depending on the type of item
            If ($item -eq "OpSec Credential" -or $item -eq "Website Credential") {
                Write-Host "$([char]27)[36m[U]$([char]27)[0m Copy Username`t$([char]27)[36m[P]$([char]27)[0m Copy Password`t" -NonewLine 
                If ($item -eq "OpSec Credential") {Write-Host "$([char]27)[36m[W]$([char]27)[0m Copy OpSec URI`t$([char]27)[36m[T]$([char]27)[0m Toggle UAT/Prod" -NonewLine}
                If ($item -eq "Website Credential") {Write-Host "$([char]27)[36m[W]$([char]27)[0m Copy Website URL" -NonewLine }
            } elseif ($item -ne "AAMS Session") {
                Write-Host "$([char]27)[36m[D]$([char]27)[0m Toggle Full Details`t" -NoNewLine
            }
            if ($Item -eq "Casino") {Write-Host "$([char]27)[36m[O]$([char]27)[0m OpSec Passwords`t$([char]27)[36m[P]$([char]27)[0m Website Passwords" -NoNewLine}
            if ($Item -eq "Transaction") {Write-Host "`t$([char]27)[36m[U]$([char]27)[0m Unlock Transaction" -NoNewLine}
            if ($Item -eq "Game") {Write-Host "$([char]27)[36m[C]$([char]27)[0m Open in ClientZone`t$([char]27)[36m[L]$([char]27)[0m Launch Game`t$([char]27)[36m[G]$([char]27)[0m Game Blocking" -NoNewLine}
            Write-Host ""
            # Exit and Return options
            Write-Host "$([char]27)[36m[Q]$([char]27)[0m Quit`t`t$([char]27)[36m[Anything else]$([char]27)[0m Back to previous menu"
            # Prompt the user for input
            $Waitkey = [System.Console]::ReadKey()
            Write-Host ""
            Write-Verbose ("[$(Get-Date)] Option selected:" + $Waitkey.key)

            switch ($Waitkey.key) {
                'A' { # All remaining items - return -1 so calling function should know to display everything remaining
                    If ($Counter + 1 -ge $ItemCount) {
                        Break
                    }
                    Return -1
                }
                'C' { # open game in ClientZone
                    If ($Item -ne "Game") {Break}
                    Return -991
                }
                'D' { # Toggle details
                    If ($Item -ne "Game" -and $Item -ne "Casino" -and $Item -ne "Transaction") {Break}
                    If ($script:QFMenuDetailDisplay -eq $true) {
                        $script:QFMenuDetailDisplay = $false
                    } else {
                        $script:QFMenuDetailDisplay = $true
                    }
                    Write-Verbose ("[$(Get-Date)] Details display toggle mode: $script:QFMenuDetailDisplay")
                    Return $Counter
                }
                'G' { # Game Blocking check
                    # would have used B but thats already set for Back... oh well
                    If ($Item -ne "Game") {Break}
                    Return -989
                }
                'Spacebar' { # Increment counter so calling function knows to display the next object
                    Return $Counter + 1
                }
                'RightArrow' {
                    Return $Counter + 1
                }
                'B' { # Decrement counter so calling function knows to display the previous object
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
                'X' { # export to excel 
                    Return -999
                }
                'U' { # Copy Username
                    if ($Item -eq "Opsec Credential") {
                        Set-Clipboard $OpSecCreds.username
                        Write-Host -ForegroundColor Yellow "Operator Security Username copied to clipboard!"
                        Return $Counter
                    } elseif ($Item -eq "Website Credential") {
                        Return -994
                    } elseif ($Item -eq "Transaction") {
                        Return -992
                    } else {
                        Break
                    }
                }
                'L' { #Launch game via qfgames.gameassists.co.uk
                    If ($Item -ne "Game") {Break}
                    Return -990
                }
                'O' { # Opsec Passwords
                    If ($Item -eq "Casino") {
                        Return -997
                    } else {Break}
                }
                'P' { # Copy Password
                    if ($Item -eq "Casino") {
                        # Get Website Passwords
                        Return -996
                    } elseif ($Item -eq "Opsec Credential") {
                        Set-Clipboard $OpSecCreds.password
                        Write-Host -ForegroundColor Yellow "Operator Security Password copied to clipboard!"
                        Return $Counter
                    } elseif ($Item -eq "Website Credential") {
                        Return -995
                    } else {
                        Break
                    }
                }
                'T' { # Toggle UAT/Production.
                    If ($Item -ne "Opsec Credential") {Break}
                    Return -998
                }
                'W' { # Copy site address
                    If ($Item -eq "Opsec Credential") {
                        If ($Params.UAT -eq $true) {
                            $OpsecURI = "https://operatorsecurityuat.valueactive.eu/system/operatorsecurityweb/v1/#/login"
                        } else {
                            $OpsecURI = "https://operatorsecurity.valueactive.eu/system/operatorsecurityweb/v1/#/login"
                        }
                        Set-Clipboard $OpsecURI
                        Write-Host -ForegroundColor Yellow "Operator Security Site Address copied to clipboard: $OpsecURI"
                        Return 0
                    } ElseIf ($Item -eq "Website Credential") {
                        Return -993
                    } else {Break}
                }
                'Q' { # Quit completely
                    $global:QFMenuExitFlag = $true
                    Break
                }
                Default { # Return to menu, exit do loop from calling function
                    Break
                }
            }
        }


        function Invoke-AAMSMenu {
            # Local function to request AAMS status
            try {
                Write-Host ""
                Write-Host "Please enter the AAMS Participation Code (Begining with N) - press ENTER on an empty line to begin..."

                $AAMSData = Get-QFAAMSStatus
                if (@($AAMSData).Count -le 0) {
                    Throw "Did not retrieve any AAMS participation data. Please check the Participation Codes are correct and the adm.gov.it site is online."
                }
                # Loop through all returned Transactions, call Invoke-ItemMenu for each one
                $i = 0
                do {
                    # Arrays are 0 indexed so add 1 to our counter for display
                    Write-Host -ForegroundColor Yellow "Displaying AAMS Session $($i + 1) of $(@($AAMSData).Count)"
                    $AAMSData[$i] | Format-List
                    If ($i -ge (@($AAMSData).Count) -1) {
                        # Disable DisplayAll and show item menu on last item
                        $DisplayAll = $false
                    }
                    If ($DisplayAll -ne $true) {
                        $Counter = Invoke-ItemMenu -Counter $i -Item 'AAMS Session' -ItemCount @($AAMSData).Count
                    }
                    Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                    if ($Counter -eq -1) {
                        # Display all the remaining items
                        $DisplayAll = $true
                        $i += 1
                    } elseif ($Counter -eq -999) {
                        # Export to Excel.
                        try {
                            [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'AAMS Sessions - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                            If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "AAMS Sessions - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                            If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                            Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                $AAMSData -ExcelDestWorksheetName AAMS
                            if (Test-Path $ExcelFilename -PathType Leaf) {
                                Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                            }
                        } catch {
                            Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                        }
                    } else {
                        $i = $Counter
                    }
                } until ($i -lt 0 -or $i -ge @($AAMSData).Count -or $global:QFMenuExitFlag -eq $true)

            } Catch {
                Write-Warning $_.Exception.Message
            }
        }


        function Invoke-RoundStatus {
            # Local function to request round status details from reconciliation API
            try {
                # offer CasinoID parameter as default option if present
                $CasinoIDInput = Get-CasinoID $script:CasinoID
                # get the UserID
                $UserIDInput = Get-UserID $script:UserID
                # Get the Casino Info
                $CasinoData = Invoke-QFPortalRequest -CasinoID $CasinoIDInput
                If ($Null -eq $CasinoData -or $CasinoData.Count -ne 1) {
                    Throw "Could not retrieve any data for the specified CasinoID, unable to continue with Round Status request."
                }
                If ($CasinoData.operatorID -le 0) {
                    Throw "The specified CasinoID is not linked to a valid OperatorID, or may be decommisioned. Unable to continue with Round Status request."
                }

                # UAT Casino check - QF UAT casinos seem to have incorrect siteID of 29 instead of 32
                If ($CasinoData.hostingSiteID -eq 29) {
                    $HostingSiteID = 32
                    $APIToken = Get-APIToken -OperatorID $CasinoData.operatorID -UAT
                } else {
                    $HostingSiteID = $CasinoData.hostingSiteID
                    $APIToken = Get-APIToken -OperatorID $CasinoData.operatorID
                }
                Write-Verbose ("[$(Get-Date)] Operator ID for specified Casino: $($CasinoData.operatorID) HostingSiteID: $HostingSiteID")

                # Set up array for splatting to the Round Status Request. 
                $RoundStatusParams = @{
                    HostingSiteID = $HostingSiteID
                    Token = $APIToken
                    CasinoID = $CasinoIDInput
                    UserID = $UserIDInput
                    RoundInfo = $true
                }
                if ($script:ReconRangeMode) {
                    Write-Host -ForegroundColor DarkGreen "Range mode - All transaction ID's between the two specified numbers will be checked. Hit ENTER on an empty line to cancel."
                    [int]$TransIDStart = Read-Host "Enter FIRST Transaction ID of the range to request round status"
                    [int]$TransIDEnd = Read-Host "Enter LAST Transaction ID of the range to request round status"
                    if ($TransIDEnd -eq 0 -or $TransIDStart -eq 0) {Return}
                    $RoundStatusParams.Add("TransactionIDs",($TransIDStart..$TransIDEnd))
                } else {
                    # if not Range Mode, Mandatory TransactionID parameter of this function will prompt user for the Transaction ID's
                    Write-Host "Enter TransactionID's for the Round Status request. Press ENTER on a blank line to continue."
                }
                # Make the Recon API Request
                $RoundStatus = Invoke-QFReconAPIRequest @RoundStatusParams
                If ($RoundStatus.Count -lt 1) {
                    Write-Host ""
                    Write-Host -ForegroundColor DarkYellow "Did not receive any data for the specified Transactions."
                    Write-Host ""
                    Return
                }
                # Add properties to the TransactionInfo nested object. These are members of the parent object but not on the TransactionInfo object.
                # numberOfEvents is how many records are in the transactionInfo object, eg free spins can have many wagers and winnings in one Transaction.
                $RoundStatus | ForEach-Object {
                    If ($null -ne $_.transactionInfo) {
                        $_.transactionInfo | Add-Member -Name 'currency' -value $_.currency -MemberType NoteProperty 
                        $_.transactionInfo | Add-Member -Name 'productId' -value $_.productID -MemberType NoteProperty
                        $_.transactionInfo | Add-Member -Name 'userId' -value $_.userId -MemberType NoteProperty
                        $_.transactionInfo | Add-Member -Name 'transactionNumber' -value $_.transactionNumber -MemberType NoteProperty
                        $_ | Add-Member -Name 'numberOfEvents' -value $_.transactionInfo.count -MemberType NoteProperty
                    } else {
                        # If nothing in Transaction Info (e.g. Unknown status for the round) add NumberOfEvents member with value 0
                        $_ | Add-Member -Name 'numberOfEvents' -value 0 -MemberType NoteProperty
                    }
                }
                
                # Loop through all returned Transactions, call Invoke-ItemMenu for each one
                $i = 0
                do {
                    # Arrays are 0 indexed so add 1 to our counter for display
                    Write-Host -ForegroundColor Yellow "Displaying Transaction $($i + 1) of $(@($RoundStatus).Count)"
                    if ($script:QFMenuDetailDisplay -and $null -ne $RoundStatus[$i].transactionInfo -and $RoundStatus[$i].numberOfEvents -gt 0) {
                        # If nothing in Transaction Info (e.g. Unknown status for the round) just display the parent object
                        $RoundStatus[$i].TransactionInfo | Format-List
                    } else {
                        $RoundStatus[$i] | Format-List -Property TransactionNumber,RoundStatusName,WinAmount,Currency,numberOfEvents
                    }
                    If ($i -ge (@($RoundStatus).Count) -1) {
                        # Disable DisplayAll and show item menu on last item
                        $DisplayAll = $false
                    }
                    If ($DisplayAll -ne $true) {
                        $Counter = Invoke-ItemMenu -Counter $i -Item Transaction -ItemCount @($RoundStatus).Count
                    }
                    Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                    if ($Counter -eq -1) {
                        # Display all the remaining items
                        $DisplayAll = $true
                        $i += 1
                    } elseif ($Counter -eq -992) {
                        # Transaction Unlock
                        Invoke-TransactionUnlock -CasinoID $CasinoIDInput -UserID $UserIDInput -TransactionIDs ($RoundStatus[$i].TransactionNumber) -APIToken $APIToken -HostingSiteID $HostingSiteID -OperatorID $CasinoData.operatorID
                        Write-Host ""
                        Write-Host -ForegroundColor DarkYellow "NOTE: Any change in status will not be shown in the below details. You will need to run another 'Round Status' request to view the updated status."
                        Write-Host ""
                    } elseif ($Counter -eq -999) {
                        # Export to Excel. Make two worksheets, one for Summary and one for Detail
                        try {
                            [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Round Status - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                            If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Round Status - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                            If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                            Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                $($RoundStatus | Select-Object -ExcludeProperty transactionInfo) -ExcelDestWorksheetName RoundStatusSummary
                            if ($null -ne $RoundStatus.TransactionInfo) {
                                # Special condition to handle some rounds with no TransactionInfo
                                Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                $($RoundStatus.TransactionInfo | Where-Object { $null -ne $_ }) -ExcelDestWorksheetName RoundStatusDetail
                            }
                            if (Test-Path $ExcelFilename -PathType Leaf) {
                                Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                            }
                        } catch {
                            Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                        }
                    } else {
                        $i = $Counter
                    }
                } until ($i -lt 0 -or $i -ge @($RoundStatus).Count -or $global:QFMenuExitFlag -eq $true)
                # Set Casino and User ID to provided options to offer them as defaults for next Recon API functions. 
                $script:CasinoID = $CasinoIDInput
                $script:UserID = $UserIDInput
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }

        function Get-APIToken {
            # Local function to request an operator API Key and generate a Token.
            # If UAT param set, will get keys/tokens for the UAT environment, otherwise Prod.
            param(

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [int]$OperatorID,

            [ValidateNotNullOrEmpty()]
            [switch]$UAT
            )

            # Get the API Key for the Casino's matching OperatorID - set up parameters for splatting to the function, check for UAT casino
            $APIKeyParams = @{
                OperatorID = $OperatorID
            }
            If ($UAT) {$APIKeyParams.Add('UAT',$true)}
            # Call function to get the API key
            $APIKey = (Get-QFOperatorAPIKeys @APIKeyParams).APIKey | Select-Object -First 1
            # Check we actually got an API key
            If ($null -eq $APIKey) {
                Throw "Couldn't get an API Key for OperatorID $OperatorID - check credentials for Operator Security Site are valid, and the operator has generated a key."
            }
            Write-Verbose ("[$(Get-Date)] API Key: $APIKey")
            # Now get an Operator Token - set up parameters for splatting to the function, check for UAT casino
            $APITokenParams = @{
                APIKey = $APIKey
            }
            If ($UAT) {$APITokenParams.Add('APIHost','operatorsecurityuat.valueactive.eu')}
            # Call function to get the token
            $APIToken = Get-QFOperatorToken @APITokenParams
            # Check we have a Token
            If ($null -eq $APIToken.AccessToken) {
                Throw "Failed to generate an API Token for CasinoID $CasinoIDInput "
            }
            Write-Verbose ("[$(Get-Date)] API Token expiry: $($APIToken.Expiry)")
            # output the API token to pipeline and return to calling function
            $APIToken.AccessToken
        }


        function Invoke-GameSearch {
            # Local function to search games by MID and CID
            Param(
                [switch]$ByName,

                [switch]$ByMIDCID
            )

            Try {
                If ($ByMIDCID) {
                    [int]$MIDInput = Read-Host -Prompt "Please enter the Game ModuleID to search for"
                    if ($MIDInput -lt  1) {
                        Continue
                    }
                    [int]$CIDInput = Read-Host -Prompt "Please enter the Game ClientID to search for (leave blank to list all games with ModuleID $MIDInput)"
                    # Parameters to splat to Invoke-QFPortalRequest
                    $GameSearchParams = @{
                        MID         = $MIDInput
                    }
                    If ($CIDInput -gt 0) {
                        $GameSearchParams.Add("CID",$CIDInput)
                    }
                } elseif ($ByName) {
                    [string]$GameNameInput = Read-Host -Prompt "Please enter the Game Name to search for"
                    if ($GameNameInput.trim().length -lt  3) {
                        Throw "Please enter a game name of at least 3 characters in length to search for."
                    }
                    $GameSearchParams = @{
                        GameName = $GameNameInput.trim()          
                    }
                } else {
                    Throw "Unable to determine search mode! Specify ByMIDCID or ByName parameters."
                }

                # Make the game search request
                $GameSearch = Invoke-QFPortalRequest @GameSearchParams
                If ($Null -eq $GameSearch -or $GameSearch.Count -eq 0) {Throw "No Games found matching the specified criteria."}
                # Loop through all returned Games, call Invoke-ItemMenu for each one
                $i = 0
                do {
                    # Arrays are 0 indexed so add 1 to our counter for display
                    Write-Host -ForegroundColor Yellow "Displaying Game $($i + 1) of $(@($GameSearch).Count)"
                    if ($script:QFMenuDetailDisplay -and $i -ge 0) {
                        $GameSearch[$i] | Format-List -Property *
                    } else {
                        $GameSearch[$i]
                    }
                    if ($GameSearch[$i].etiProductId -gt 0) {
                        Write-Host -ForegroundColor Yellow "ETI Support details:"
                        Get-QFETIProviderInfo -Id $GameSearch[$i].etiProductId | Format-List -Property Email,PortalURI,PortalUsername,PortalPassword
                    } elseif ($GameSearch[$i].provider -like "*On Air Entertainment*") {
                        Write-Host -ForegroundColor Yellow "ETI Support details:"
                        Get-QFETIProviderInfo -Name "On Air Entertainment" | Format-List -Property Email,PortalURI,PortalUsername,PortalPassword
                    }
                    If ($i -ge (@($GameSearch).Count) -1) {
                        # Disable DisplayAll and show item menu on last item
                        $DisplayAll = $false
                    }
                    If ($DisplayAll -ne $true) {
                        $Counter = Invoke-ItemMenu -Counter $i -Item Game -ItemCount @($GameSearch).Count
                    }
                    Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                    if ($Counter -eq -1) {
                        # Display all the remaining items
                        $DisplayAll = $true
                        $i += 1
                    } elseif ($Counter -eq -990) {
                        # Launch game
                        # Uses qfgames.gameassists.co.uk 
                        Write-Host -Foregroundcolor DarkCyan "Launching $($GameSearch[$i].gameName) - CasinoID 21699 - OperatorID 47600"
                        Write-Host ""
                        [string]$Currency = Read-Host "Please enter the 3 character Currency Code (Or just press ENTER for EUR)"
                        If ($Currency.trim() -eq "") {
                            $Currency = "EUR"
                        } else {
                            $Currency = $Currency.ToUpper().Trim()
                        }
                        Start-QFGame -LaunchCode $GameSearch[$i].uglGameId -Currency $Currency
                    } elseif ($Counter -eq -991) {
                        # Open game page in Client Zone
                        # Use a regex to find the game name in the CZ URL, this seems to work in 99% of cases
                        # remove any non-alphanumeric characters and  variant number off the end
                        $CZLink = $GameSearch[$i].gameName -replace "[^a-zA-Z0-9]+","" -replace "([vV][0-9][0-9])$",""
                        # Some special cases... probably more to follow
                        If ($CZLink.trim() -like "HiLo") { 
                            $CZLink = "HacksawHiLo"
                        } elseIf ($CZLink.trim() -like "FruitFiestaETI") { 
                            $CZLink = "FruitFiesta"
                        } elseIf ($CZLink.trim() -like "ToshiVideoClub") { 
                            $CZLink = "ToshVideoClub"
                        } elseIf ($CZLink.trim() -like "QueensofRaPOWERCOMBO") { 
                            $CZLink = "QueensofRaPowerCombo"
                        } elseif ($CZLink.trim() -match "GameofThrones(243Ways|15Lines)") {
                            $CZLink = "GameOfThrones" + $Matches[1]
                        }
                        # Open browser to ClientZone URI
                        Start-Process ("https://clientzone.gamesglobal.com/Games/" + $CZLink.trim())
                    } elseif ($Counter -eq -989) {
                        Write-Host -ForegroundColor DarkCyan "Game Blocking Check"
                        # Get the CasinoID
                        $CasinoIDInput = Get-CasinoID $script:CasinoID
                        Write-Host ""
                        # Params for splatting to Get-QFGameBlocking 
                        $GameBlockParams = @{
                            CasinoId = $CasinoIDInput
                            MID = $GameSearch[$i].ModuleId
                            CID = $GameSearch[$i].ClientId
                        }
                        # Call the function for game block data
                        $GameBlock = Get-QFGameBlocking @GameBlockParams
                        # Get an array of CRQ Numbers
                        $CRQNumbers = @()
                        $GameBlock.Blocks.OBSNumber | Foreach-Object {
                            If ($_ -match "CRQ[0-9]+") {
                                $CRQNumbers += $Matches.Values
                            }
                        }
                        if (@($CRQNumbers.Count) -gt 0) {
                            Write-Host "Found these CRQ numbers for these game blocking records:"
                            Write-Host ($CRQNumbers | Sort-Object -Unique)
                            Write-Host "Press SPACEBAR or ENTER to copy them to clipboard, or anything else to return to previous menu:"
                            $Waitkey = [System.Console]::ReadKey()
                            Write-Host ""
                            If ($Waitkey.Key -eq "Spacebar" -or $Waitkey.Key -eq "Enter") {
                                Set-Clipboard ($CRQNumbers | Sort-Object -Unique)
                            }
                        }
                        $script:CasinoID = $CasinoIDInput
                    } elseif ($Counter -eq -999) {
                        # Export to Excel
                        try {
                            [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Game Search - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                            If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Game Search - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                            If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                            Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData $GameSearch -ExcelDestWorksheetName Games
                            if (Test-Path $ExcelFilename -PathType Leaf) {
                                Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                            }
                        } catch {
                            Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                        }
                    } else {
                        $i = $Counter
                    }
                } until ($i -lt 0 -or $i -ge @($GameSearch).Count)
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }

        function Invoke-CasinoSearch {
            # Local function to search Casinos by CasinoID, Name, or OperatorID
            Param(
                [switch]$ByCasinoID,

                [switch]$ByOperatorID,

                [switch]$ByCasinoName
            )

            Try {
                If ($ByCasinoID) {
                    $CasinoIDInput = Get-CasinoID $script:CasinoID
                    $CasinoSearchParams = @{
                        CasinoID = $CasinoIDInput
                    }
                } elseif ($ByCasinoName) {
                    [string]$CasinoName = Read-Host -Prompt "Please enter the Casino Name to search for"
                    If ($Null -eq $CasinoName -or $CasinoName.trim().Length -lt 3) {
                        Throw "Please enter a valid Casino Name to search for of at least 3 characters in length."
                    }
                    $CasinoSearchParams = @{
                        CasinoName = $CasinoName
                    }
                } elseif ($ByOperatorID) {
                    $OperatorIDInput = Get-OperatorID $script:OperatorID
                    $CasinoSearchParams = @{
                        OperatorID = $OperatorIDInput
                    }
                } else {
                    Throw "Unable to determine search mode!"
                }

                $CasinoInfo = Invoke-QFPortalRequest @CasinoSearchParams
                If ($CasinoInfo -ne "") {
                    if ($ByOperatorID) {
                        $CasinoInfo | Out-Host
                        # Set OperatorID to the input value so we can offer it as a default for the next request
                        $script:OperatorID = $OperatorIDInput
                        # Make a new request for CasinosByOperatorID now that we've confirmed the OperatorID value is valid
                        $CasinoSearchParams.Remove('OperatorId')
                        $CasinoSearchParams.Add('CasinosForOperatorID',$OperatorIdInput)
                        $CasinoInfo = Invoke-QFPortalRequest @CasinoSearchParams
                        Write-Host "$(@($CasinoInfo).Count) Casino(s) found linked to this Operator. Hit SPACEBAR or ENTER to list these casinos:"
                        $Waitkey = [System.Console]::ReadKey()
                        Write-Host ""
                        If ($Waitkey.Key -ne "Enter" -and $Waitkey.Key -ne "Spacebar") {Return}
                    }
                    # Set CasinoID to the input value so we can offer it as a default for the next request
                    If ($ByCasinoID) {$script:CasinoID = $CasinoIDInput}

                    # Loop through all returned Casinos, call Invoke-ItemMenu for each one
                    $i = 0
                    do {
                        # Arrays are 0 indexed so add 1 to our counter for display
                        Write-Host -ForegroundColor Yellow "Displaying Casino $($i + 1) of $(@($CasinoInfo).Count)"
                        if ($script:QFMenuDetailDisplay -and $i -ge 0) {
                            $CasinoInfo[$i] | Select-Object -ExcludeProperty LobbySettings,productSettings,databases,vanguardApiSettings,vanguardOperatorSettings -Property *
                            $CasinoInfo[$i].productSettings
                            $CasinoInfo[$i].vanguardApiSettings
                        } else {
                            $CasinoInfo[$i]
                        }
                        If ($i -ge (@($CasinoInfo).Count) -1) {
                            # Disable DisplayAll and show item menu on last item
                            $DisplayAll = $false
                        }
                        If ($DisplayAll -ne $true) {
                            $Counter = Invoke-ItemMenu -Counter $i -Item Casino -ItemCount @($CasinoInfo).Count
                        }
                        Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                        if ($Counter -eq -1) {
                            # Display all the remaining items
                            $DisplayAll = $true
                            $i += 1
                        } elseif ($Counter -eq -997) {
                            # Operator Security Credentials
                            if ($CasinoInfo[$i].OperatorId -le 0) {
                                Write-Host -ForegroundColor DarkYellow 'Not a valid OperatorID, cannot retrieve Operator Security credentials.'
                                Continue
                            } 
                            $OpsecParams = @{
                                OperatorIDInput = $CasinoInfo[$i].OperatorId
                            }
                            If ($CasinoInfo[$i].Environment -eq "UAT") {$OpsecParams.Add('UAT',$true)}
                            Get-OpsecCreds @OpsecParams
                        } elseif ($Counter -eq -996) {
                            # Website Credentials
                            Get-WebsiteCreds -CasinoIDInput $CasinoInfo[$i].productId
                        } elseif ($Counter -eq -999) {
                            # Export to Excel
                            try {
                                [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Casinos - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                                If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Casinos - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                                If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                                Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                $($CasinoInfo | Select-Object -ExcludeProperty LobbySettings,productSettings,databases,vanguardApiSettings,vanguardOperatorSettings,Lobbies) -ExcelDestWorksheetName Casino
                                if (Test-Path $ExcelFilename -PathType Leaf) {
                                    Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                                }
                            } catch {
                                Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                            }
                        } else {
                            $i = $Counter
                        }
                    } until ($i -lt 0 -or $i -ge @($CasinoInfo).Count -or $global:QFMenuExitFlag -eq $true)
                } else {
                    Throw "No matching casinos found."
                }
            } Catch {
                Write-Warning $_.Exception.Message
            }
            
        }        


        function Get-CasinoID {
            # Local function to prompt user for CasinoID if a default value was not provided by a parameter when cmdlet was invoked
            param (
                [int]$CasinoID
            )

            If ($Null -ne $CasinoID -and $CasinoID -gt 0) {
                [int]$CasinoIDInput = Read-Host -Prompt "Please enter the CasinoID/ServerID (ENTER for $CasinoID)"
                if ($CasinoIDInput -eq 0) {
                    $CasinoIDInput = $CasinoID
                }
            } else {
                [int]$CasinoIDInput = Read-Host -Prompt "Please enter the CasinoID/ServerID"
                If ($CasinoIDInput -lt  1) {
                    Throw "Please enter a valid numeric CasinoID."
                }
            }
            $CasinoIDInput
        }


        function Get-OperatorID {
            # Local function to prompt user for Operator ID if a default value was not provided by a parameter when cmdlet was invoked
            param (
                [int]$OperatorID
            )

            If ($Null -ne $OperatorID -and $OperatorID -gt 0) {
                [int]$OperatorIDInput = Read-Host -Prompt "Please enter the OperatorID (ENTER for $OperatorID)"
                if ($OperatorIDInput -eq 0) {
                    $OperatorIDInput = $OperatorID
                }
            } else {
                [int]$OperatorIDInput = Read-Host -Prompt "Please enter the OperatorID"
                If  ($OperatorIDInput -lt  1) {
                    Throw "Please enter a valid numeric OperatorID."
                }
            }
            $OperatorIDInput
        }


        function Get-UserID {
            # Local function to prompt user for UserID if a default value was not provided by a parameter when cmdlet was invoked
            param (
                [int]$UserID
            )

            If ($Null -ne $UserID -and $UserID -gt 0) {
                [int]$UserIDInput = Read-Host -Prompt "Please enter the player's UserID (ENTER for $UserID)"
                if ($UserIDInput -eq 0) {
                    $UserIDInput = $UserID
                }
            } else {
                [int]$UserIDInput = Read-Host -Prompt "Please enter the player's UserID"
                If  ($UserIDInput -lt  1) {
                    Throw "Please enter a valid numeric UserID."
                }
            }
            $UserIDInput
        }


        function Get-Login {
            # Local function to prompt user for player Login if a default value was not provided by a parameter when cmdlet was invoked
            param (
                [string]$Login
            )
            If ($Null -ne $Login -and [string]$Login.length -ge 3) {
                [string]$UserLoginInput = Read-Host -Prompt "Please enter the player Login to search for (ENTER for $Login)"
                if ($UserLoginInput -eq "") {
                    $UserLoginInput = $Login
                }
            } else {
                [string]$UserLoginInput = Read-Host -Prompt "Please enter the player Login to search for (Casino login prefix is optional)"
            }
            If ($Null -eq $UserLoginInput -or $UserLoginInput.trim().Length -lt 3) {
                Throw "Please enter a valid user Login to search for of at least 3 characters in length."
            }
            $UserLoginInput
        }


        function Invoke-TransactionUnlock {
            # Local function to call Reconciliation API and unlock transactions
            # These parameters can be passed by a calling function, otherwise will prompt the user to enter them manually
            param (
                [ValidateNotNullOrEmpty()]
                [int]$CasinoID,

                [ValidateNotNullOrEmpty()]
                [int]$UserID,

                [ValidateNotNullOrEmpty()]
                [int]$OperatorID,

                [ValidateNotNullOrEmpty()]
                [int]$HostingSiteID,

                [ValidateNotNullOrEmpty()]
                [int[]]$TransactionIDs,

                [ValidateNotNullOrEmpty()]
                [string]$APIToken
            )

            try {
                Write-Host -ForegroundColor Yellow "Unlocking Transactions from Vanguard Queues may affect a player's balance."
                Write-Host -ForegroundColor Yellow "You MUST inform the operator whenever you unlock any transactions, as they may need to manually credit or refund the player."
                Write-Host -ForegroundColor Yellow "Operators also have the ability to unlock transactions from their Back Office, and any transactions that you unlock will also be visible there."
                Write-Host -ForegroundColor Yellow "Enter Y to acknowledge and proceed, or anything else to cancel: " -NoNewline
                $Waitkey = [System.Console]::ReadKey()
                If ($Waitkey.Key -ne "Y") {Return}
                Write-Host ""
                # offer CasinoID parameter as default option if present
                if ($PSBoundParameters.ContainsKey('CasinoID')) {
                    $CasinoIDInput = $CasinoID
                } else {
                    $CasinoIDInput = Get-CasinoID $script:CasinoID
                }
                Write-Verbose ("[$(Get-Date)] CasinoID: $CasinoIDInput")

                # get the UserID
                if ($PSBoundParameters.ContainsKey('UserId')) {
                    $UserIDInput = $UserID
                } else {
                    $UserIDInput = Get-UserID $script:UserID
                }
                Write-Verbose ("[$(Get-Date)] UserID: $UserIDInput")

                # Transaction Unlock Reference.
                Write-Host -ForegroundColor Yellow "Please enter the Transaction Unlock Reference. This will be visible to Operators in the Back Office and in Transaction Audits."
                If ($null -ne $REQNumber -and [string]$REQNumber.trim() -ne "") {
                    [string]$Reference = Read-Host -Prompt "Transaction Unlock Reference (ENTER for $REQNumber)"
                    if ($Reference.trim() -eq "") {
                        $Reference = $REQNumber
                    }
                } else {
                    [string]$Reference = Read-Host -Prompt "Transaction Unlock Reference"
                }

                # Get the Casino Info
                if (!($PSBoundParameters.ContainsKey('OperatorID') -and $PSBoundParameters.ContainsKey('HostingSiteID'))) {
                    $CasinoData = Invoke-QFPortalRequest -CasinoID $CasinoIDInput
                    If ($Null -eq $CasinoData -or $CasinoData.Count -ne 1) {
                        Throw "Could not retrieve any data for the specified CasinoID, unable to continue with Round Status request."
                    }
                    If ($CasinoData.operatorID -le 0) {
                        Throw "The specified CasinoID is not linked to a valid OperatorID, or may be decommisioned. Unable to continue with Round Status request."
                    }
                    $OperatorID = $CasinoData.operatorID
                    # UAT Casino check - QF UAT casinos seem to have incorrect siteID of 29 instead of 32
                    If ($CasinoData.hostingSiteID -eq 29) {
                        $HostingSiteID = 32
                    } else {
                        $HostingSiteID = $CasinoData.hostingSiteID
                    }
                }
                Write-Verbose ("[$(Get-Date)] OperatorID: $operatorID HostingSiteID: $HostingSiteID ")
                
                If (!($PSBoundParameters.ContainsKey('APIToken'))) {
                    # Get the API Key for the Casino's matching OperatorID
                    If ($HostingSiteID -eq 32) {
                        Write-Verbose ("[$(Get-Date)] UAT Casino, requesting UAT API tokens")
                        $APIToken = Get-APIToken -OperatorID $OperatorID -UAT
                    } else {
                        $APIToken = Get-APIToken -OperatorID $OperatorID
                    }
                }

                # Set up array for splatting to the Transaction Unlock Request. 
                $UnlockParams = @{
                    HostingSiteID = $hostingSiteID
                    Token = $APIToken
                    CasinoID = $CasinoIDInput
                    UserID = $UserIDInput
                    Unlock = $true
                }
                If ($null -ne $Reference -and $Reference.trim() -ne "") {
                    $UnlockParams.Add("Reference",$Reference)
                }
                if ($PSBoundParameters.ContainsKey('TransactionIDs')) {
                    $UnlockParams.Add("TransactionIDs",$TransactionIDs)
                } elseif ($script:ReconRangeMode) {
                    Write-Host -ForegroundColor DarkGreen "Range mode - All transaction ID's between the two specified numbers will be unlocked. Hit ENTER on an empty line to cancel."
                    [int]$TransIDStart = Read-Host "Enter FIRST Transaction ID of the range to unlock"
                    [int]$TransIDEnd = Read-Host "Enter LAST Transaction ID of the range to unlock"
                    if ($TransIDEnd -eq 0 -or $TransIDStart -eq 0) {Return}
                    $UnlockParams.Add("TransactionIDs",($TransIDStart..$TransIDEnd))
                } else {
                    # if not Range Mode, Mandatory TransactionID parameter of this function will prompt user for the Transaction ID's
                    Write-Host "Enter TransactionID's to Unlock. Press ENTER on a blank line to begin unlocking transactions. If you make a mistake, hit CTRL + C to exit."
                }
                Write-Verbose ("[$(Get-Date)] Unlock request parameters:")
                foreach($k in $UnlockParams.Keys) { Write-Verbose "$k $($UnlockParams[$k])" }
                # Make the Recon API Request
                $UnlockStatus = Invoke-QFReconAPIRequest @UnlockParams
                If ($UnlockStatus.Count -lt 1) {
                    Write-Host ""
                    Write-Host -ForegroundColor DarkYellow "Did not receive any data for the specified Transaction(s)."
                    Write-Host ""
                    Return
                }
                $UnlockStatus | Format-Table
                # Set Casino and User ID to provided options to offer them as defaults for next Recon API functions. 
                $script:CasinoID = $CasinoIDInput
                $script:UserID = $UserIDInput
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }


        function Invoke-VanguardQueues {
            try {
                # offer CasinoID parameter as default option if present
                $CasinoIDInput = Get-CasinoID $script:CasinoID
                # get the UserID
                $UserIDInput = Get-UserID $script:UserID
                # Get the Casino Info
                $CasinoData = Invoke-QFPortalRequest -CasinoID $CasinoIDInput
                If ($Null -eq $CasinoData -or $CasinoData.Count -ne 1) {
                    Throw "Could not retrieve any data for the specified CasinoID, unable to continue with Round Status request."
                }
                If ($CasinoData.operatorID -le 0) {
                    Throw "The specified CasinoID is not linked to a valid OperatorID, or may be decommisioned. Unable to continue with Round Status request."
                }

                Write-Verbose ("[$(Get-Date)] Operator ID for specified Casino: $($CasinoData.operatorID)")

                # UAT Casino check - QF UAT casinos seem to have incorrect siteID of 29 instead of 32
                # Get the API Key for the Casino's matching OperatorID
                If ($CasinoData.hostingSiteID -eq 29) {
                    Write-Verbose ("[$(Get-Date)] UAT Casino, requesting UAT API tokens")
                    $HostingSiteID = 32
                    $APIToken = Get-APIToken -OperatorID $CasinoData.operatorID -UAT
                } else {
                    $HostingSiteID = $CasinoData.hostingSiteID
                    $APIToken = Get-APIToken -OperatorID $CasinoData.operatorID
                }
                

            
                # Set up array for splatting to the Vanguard Queues Request. 
                $QueueInfoParams = @{
                    HostingSiteID = $HostingSiteID
                    Token = $APIToken
                    CasinoID = $CasinoIDInput
                    UserID = $UserIDInput
                    QueueInfo = $true
                }

                # Make the Recon API Request
                $QueueInfo = Invoke-QFReconAPIRequest @QueueInfoParams

                # Save the userID and casino ID for next request
                $script:CasinoID = $CasinoIDInput
                $script:UserID = $UserIDInput

                If ($QueueInfo.Count -lt 1) {
                    Write-Host ""
                    Write-Host -ForegroundColor DarkYellow "Did not receive any Transaction Queue data for the specified player."
                    Write-Host ""
                    Return
                }
                # Output count of queued transactions. Do loop to return to transaction count menu after looking at individual queues.
                Do {
                    Write-Host ""
                    [PSCustomObject]@{
                        'Commit Queue Transactions' = $QueueInfo.CommitCount
                        'Rollback Queue Transactions' = $QueueInfo.RollbackCount
                        'Incomplete Games' = $QueueInfo.IncompleteCount
                    } | Format-List

                    # If no queued transactions nothing more to do, just return
                    If ($QueueInfo.CommitCount -le 0 -and $QueueInfo.RollbackCount -le 0 -and $QueueInfo.IncompleteCount -le 0) {Return}

                    # Show a simple menu for each Queue with transactions
                    Write-Host "Please select a queue to view detailed information:"
                    If ($QueueInfo.CommitCount -gt 0) {Write-Host "$([char]27)[36m[C]$([char]27)[0m Commit Queue`t`t" -NoNewline}
                    If ($QueueInfo.RollbackCount -gt 0) {Write-Host "$([char]27)[36m[R]$([char]27)[0m Rollback Queue`t`t" -NoNewline}
                    If ($QueueInfo.IncompleteCount -gt 0) {Write-Host "$([char]27)[36m[I]$([char]27)[0m Incomplete Games" -NoNewline}
                    Write-Host ""
                    Write-Host "$([char]27)[36m[Q]$([char]27)[0m Quit`t$([char]27)[36m[X]$([char]27)[0m Export to Excel`t$([char]27)[36m[Anything else]$([char]27)[0m Return to main menu"
                    $Waitkey = [System.Console]::ReadKey()
                    Write-Host ""
                    Write-Verbose ("[$(Get-Date)] Option selected:" + $Waitkey.key)
            
                    switch ($Waitkey.key) {
                        C {
                            If (@($QueueInfo.CommitCount).count -le 0) {Return}
                            # Loop through all Commit Queue Transactions, call Invoke-ItemMenu for each one
                            $i = 0
                            do {
                                # Arrays are 0 indexed so add 1 to our counter for display
                                Write-Host -ForegroundColor Yellow "Displaying Transaction $($i + 1) of $(@($QueueInfo.CommitQueue).Count)"
                                if ($script:QFMenuDetailDisplay -and $null -ne $QueueInfo.CommitQueue[$i]) {
                                    # If nothing in Transaction Info (e.g. Unknown status for the round) just display the parent object
                                    $QueueInfo.CommitQueue[$i] | Format-List
                                } else {
                                    $QueueInfo.CommitQueue[$i] | Format-List -Property TransactionNumber,winAmount,Currency,gameName,dateCreated
                                }
                                If ($i -ge (@($QueueInfo.CommitQueue[$i]).Count) -1) {
                                    # Disable DisplayAll and show item menu on last item
                                    $DisplayAll = $false
                                }
                                If ($DisplayAll -ne $true) {
                                    $Counter = Invoke-ItemMenu -Counter $i -Item Transaction -ItemCount @($QueueInfo.CommitQueue).Count
                                }
                                Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                                if ($Counter -eq -1) {
                                    # Display all the remaining items
                                    $DisplayAll = $true
                                    $i += 1
                                } elseif ($Counter -eq -992) {
                                    # Transaction Unlock
                                    Invoke-TransactionUnlock -CasinoID $CasinoIDInput -UserID $UserIDInput -TransactionIDs ($QueueInfo.CommitQueue[$i].TransactionNumber) -APIToken $APIToken -HostingSiteID $HostingSiteID -OperatorID $CasinoData.operatorID
                                    Write-Host ""
                                    Write-Host -ForegroundColor DarkYellow "NOTE: Any change in status will not be shown in the below details. You will need to run another 'Round Status' or 'Vanguard Queues' request to view the updated status."
                                    Write-Host ""
                                } elseif ($Counter -eq -999) {
                                    # Export to Excel.
                                    try {
                                        [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Commit Queue - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                                        If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Commit Queue - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                                        If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                                        Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                        $QueueInfo.CommitQueue -ExcelDestWorksheetName CommitQueue
                                        if (Test-Path $ExcelFilename -PathType Leaf) {
                                            Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                                        }
                                    } catch {
                                        Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                                    }
                                } else {
                                    $i = $Counter
                                }
                            } until ($i -lt 0 -or $i -ge @($QueueInfo.CommitQueue).Count -or $global:QFMenuExitFlag -eq $true)
                        }
                        I {
                            If (@($QueueInfo.IncompleteGameQueue).Count -le 0) {Return}
                            # Loop through all Incomplete Games, call Invoke-ItemMenu for each one
                            $i = 0
                            do {
                                # Arrays are 0 indexed so add 1 to our counter for display
                                Write-Host -ForegroundColor Yellow "Displaying Transaction $($i + 1) of $(@($QueueInfo.IncompleteGameQueue).Count)"
                                if ($script:QFMenuDetailDisplay -and $null -ne $QueueInfo.IncompleteGameQueue[$i]) {
                                    # If nothing in Transaction Info (e.g. Unknown status for the round) just display the parent object
                                    $QueueInfo.IncompleteGameQueue[$i] | Format-List
                                } else {
                                    $QueueInfo.IncompleteGameQueue[$i] | Format-List -Property TransactionNumber,gameName,dateCreated
                                }
                                If ($i -ge (@($QueueInfo.IncompleteGameQueue[$i]).Count) -1) {
                                    # Disable DisplayAll and show item menu on last item
                                    $DisplayAll = $false
                                }
                                If ($DisplayAll -ne $true) {
                                    $Counter = Invoke-ItemMenu -Counter $i -Item Transaction -ItemCount @($QueueInfo.IncompleteGameQueue).Count
                                }
                                Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                                if ($Counter -eq -1) {
                                    # Display all the remaining items
                                    $DisplayAll = $true
                                    $i += 1
                                } elseif ($Counter -eq -992) {
                                    # Transaction Unlock
                                    Invoke-TransactionUnlock -CasinoID $CasinoIDInput -UserID $UserIDInput -TransactionIDs ($QueueInfo.IncompleteGameQueue[$i].TransactionNumber) -APIToken $APIToken -HostingSiteID $HostingSiteID -OperatorID $CasinoData.operatorID
                                    Write-Host ""
                                    Write-Host -ForegroundColor DarkYellow "NOTE: Any change in status will not be shown in the below details. You will need to run another 'Round Status' or 'Vanguard Queues' request to view the updated status."
                                    Write-Host ""
                                } elseif ($Counter -eq -999) {
                                    # Export to Excel
                                    try {
                                        [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Incomplete Games - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                                        If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Incomplete Games - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                                        If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                                        Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                        $QueueInfo.IncompleteGameQueue -ExcelDestWorksheetName IncompleteGameQueue
                                        if (Test-Path $ExcelFilename -PathType Leaf) {
                                            Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                                        }
                                    } catch {
                                        Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                                    }
                                } else {
                                    $i = $Counter
                                }
                            } until ($i -lt 0 -or $i -ge @($QueueInfo.IncompleteGameQueue).Count -or $global:QFMenuExitFlag -eq $true)
                        }
                        R {
                            If (@($QueueInfo.RollbackQueue).count -le 0) {Return}
                            # Loop through all Rollback queue Transactions, call Invoke-ItemMenu for each one
                            $i = 0
                            do {
                                # Arrays are 0 indexed so add 1 to our counter for display
                                Write-Host -ForegroundColor Yellow "Displaying Transaction $($i + 1) of $(@($QueueInfo.RollbackQueue).Count)"
                                if ($script:QFMenuDetailDisplay -and $null -ne $QueueInfo.RollbackQueue[$i]) {
                                    # If nothing in Transaction Info (e.g. Unknown status for the round) just display the parent object
                                    $QueueInfo.RollbackQueue[$i] | Format-List
                                } else {
                                    $QueueInfo.RollbackQueue[$i] | Format-List -Property TransactionNumber,refundAmount,Currency,gameName,dateCreated
                                }
                                If ($i -ge (@($QueueInfo.RollbackQueue[$i]).Count) -1) {
                                    # Disable DisplayAll and show item menu on last item
                                    $DisplayAll = $false
                                }
                                If ($DisplayAll -ne $true) {
                                    $Counter = Invoke-ItemMenu -Counter $i -Item Transaction -ItemCount @($QueueInfo.RollbackQueue).Count
                                }
                                Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                                if ($Counter -eq -1) {
                                    # Display all the remaining items
                                    $DisplayAll = $true
                                    $i += 1
                                } elseif ($Counter -eq -992) {
                                    # Transaction Unlock
                                    Invoke-TransactionUnlock -CasinoID $CasinoIDInput -UserID $UserIDInput -TransactionIDs ($QueueInfo.RollbackQueue[$i].TransactionNumber) -APIToken $APIToken -HostingSiteID $HostingSiteID -OperatorID $CasinoData.operatorID
                                    Write-Host ""
                                    Write-Host -ForegroundColor DarkYellow "NOTE: Any change in status will not be shown in the below details. You will need to run another 'Round Status' or 'Vanguard Queues' request to view the updated status."
                                    Write-Host ""
                                } elseif ($Counter -eq -999) {
                                    # Export to Excel
                                    try {
                                        [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Rollback Queue - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                                        If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Rollback Queue - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                                        If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                                        Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                        $QueueInfo.RollbackQueue -ExcelDestWorksheetName RollbackQueue
                                        if (Test-Path $ExcelFilename -PathType Leaf) {
                                            Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                                        }
                                    } catch {
                                        Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                                    }
                                } else {
                                    $i = $Counter
                                }
                            } until ($i -lt 0 -or $i -ge @($QueueInfo.RollbackQueue).Count -or $global:QFMenuExitFlag -eq $true)
                        }
                        X {
                            # Export to Excel. Make a worksheet for each Queue
                            try {
                                [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Vanguard Queues - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                                If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Vanguard Queues - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                                If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                                If (@($QueueInfo.CommitQueue).Count -gt 0) {
                                    Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                    $QueueInfo.CommitQueue -ExcelDestWorksheetName CommitQueue
                                }
                                If (@($QueueInfo.RollbackQueue).Count -gt 0) {
                                    Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                    $QueueInfo.RollbackQueue -ExcelDestWorksheetName RollbackQueue
                                }
                                If (@($QueueInfo.IncompleteGameQueue).Count -gt 0) {
                                    Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData `
                                    $QueueInfo.IncompleteGameQueue -ExcelDestWorksheetName IncompleteGames
                                }
                                if (Test-Path $ExcelFilename -PathType Leaf) {
                                    Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                                }
                            } catch {
                                Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                            }
                        }
                        Q { # quit completely
                            $global:QFMenuExitFlag = $true
                            Return
                        }
                        Default { # back to main menu
                            Return
                        }
                    }
                } while ($true)
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }


        function Get-OpsecCreds {
            # Local function to retrieve Operator Security Credentials
            param (
                # If this is set to an operatorID, won't prompt user to enter one - just make the request straight away
                [int]$OperatorIDInput,
                [switch]$UAT
            )

            # If $OperatorIDInput not set, offer OperatorID parameter as default option if present
            Try {
                If ($OperatorIDInput -le 0) {
                    $OperatorIDInput = Get-OperatorID $script:OperatorID
                }
                $Params = @{
                    OpSecID = $OperatorIDInput
                }
                If ($UAT) {$Params.add('UAT',$true)} else {$Params.add('UAT',$false)}
                # Make portal API request
                $OpSecCreds = Invoke-QFPortalRequest @Params
                If ($OpSecCreds.Count -lt 1) {
                    Write-Host ""
                    Write-Host -ForegroundColor DarkYellow "No Operator Security Credentials found for the specified Operator ID."
                    Write-Host ""
                    Return
                }
                $script:OperatorID = $OperatorIDInput
                do {
                    $OpSecCreds | Format-List -Property operatorId,username,password,env
                    # Display Options to user
                    $Counter = Invoke-ItemMenu -Counter 0 -Item "OpSec Credential" -ItemCount 1
                    Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter")
                    if ($Counter -eq -998) {
                        # Toggle UAT/Prod and make another API request
                        If ($Params.UAT -eq $true) {$Params.UAT = $false} else {$Params.UAT = $true}
                        $OpSecCreds = Invoke-QFPortalRequest @Params
                        $i = 0
                    } elseif ($Counter -eq -999) {
                        # Export to Excel
                        try {
                            [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Opsec Credentials $OperatorIDInput - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                            If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Opsec Credentials $OperatorIDInput - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                            If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                            Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData $OpSecCreds -ExcelDestWorksheetName OpSecCreds
                            if (Test-Path $ExcelFilename -PathType Leaf) {
                                Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                            }
                            $i = 0
                        } catch {
                            Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                        }
                    } else {$i = $Counter}
                } until ($i -ne 0 -or $global:QFMenuExitFlag -eq $true)
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }


        function Get-WebsiteCreds {
            # Local function to retrieve Casino Website Credentials
            param (
                # If this is set to a CasinoID, won't prompt user to enter one - just make the request straight away
                [int]$CasinoIDInput
            )

            # If $CasinoIDInput not set, offer CasinoID parameter as default option if present
            Try {
                If ($CasinoIDInput -le 0) {
                    $CasinoIDInput = Get-CasinoID $script:CasinoID
                }
                $Params = @{
                    WebsiteCasinoID = $CasinoIDInput
                }
                # Make portal API request
                $WebsiteCreds = Invoke-QFPortalRequest @Params
                If ($WebsiteCreds.Count -lt 1) {
                    Write-Host ""
                    Write-Host -ForegroundColor DarkYellow "No Website Credentials found for the specified Casino."
                    Write-Host ""
                    Return
                }
                $script:CasinoID = $CasinoIDInput
                # Loop through all returned Credentials, call Invoke-ItemMenu for each one
                $i = 0
                do {
                    # Arrays are 0 indexed so add 1 to our counter for display
                    Write-Host -ForegroundColor Yellow "Displaying Casino Website Credentials $($i + 1) of $(@($WebsiteCreds).Count)"
                    $WebsiteCreds[$i] | Format-List -Property serverid,url,username,password,notes,loggedby
                    If ($i -ge (@($WebsiteCreds).Count) -1) {
                        # Disable DisplayAll and show item menu on last item
                        $DisplayAll = $false
                    }
                    If ($DisplayAll -ne $true) {
                        $Counter = Invoke-ItemMenu -Counter $i -Item "Website Credential" -ItemCount @($WebsiteCreds).count
                    }
                    Write-Verbose ("[$(Get-Date)] Invoke-ItemMenu return value: $Counter i object value: $i")
                    if ($Counter -eq -1) {
                        # Display all the remaining items
                        $DisplayAll = $true
                        $i += 1
                    } elseif ($Counter -eq -993) {
                        # Copy website URI
                        Set-Clipboard $WebsiteCreds[$i].url
                        Write-Host -ForegroundColor Yellow "Website Address copied to clipboard!"
                    } elseif ($Counter -eq -994) {
                        # Copy website Username
                        Set-Clipboard $WebsiteCreds[$i].username
                        Write-Host -ForegroundColor Yellow "Username copied to clipboard!"
                    } elseif ($Counter -eq -995) {
                        # Copy password
                        Set-Clipboard $WebsiteCreds[$i].password
                        Write-Host -ForegroundColor Yellow "Password copied to clipboard!"
                    } elseif ($Counter -eq -999) {
                        # Export to Excel
                        try {
                            [string]$ExcelFilename = Read-Host "Enter filename (ENTER for 'Website Credentials $CasinoIDInput - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx')"
                            If ($ExcelFilename.trim() -eq "") {$ExcelFilename = "Website Credentials $CasinoIDInput - $(Get-Date -Format 'yyyy-MM-dd hh-mm').xlsx"}
                            If ($ExcelFilename.trim() -notlike "*.xlsx") {$ExcelFilename = $ExcelFilename.trim() + ".xlsx"}
                            Export-QFExcel -StartRow 1 -ExcelFileName "$ExcelFilename" -ExcelTemplate $null -ExcelData $WebsiteCreds -ExcelDestWorksheetName WebsiteCreds
                            if (Test-Path $ExcelFilename -PathType Leaf) {
                                Start-Process $ExcelFilename -ErrorAction SilentlyContinue
                            }
                            $i = 0
                        } catch {
                            Write-Warning ("Failed to export data to Excel: " + $_.Exception.Message)
                        }
                    } else {
                        $i = $Counter
                    }
                } until ($i -lt 0 -or $i -ge @($WebsiteCreds).Count  -or $global:QFMenuExitFlag -eq $true)
            } Catch {
                Write-Warning $_.Exception.Message
            }
        }


        # End of local functions


        # set this so any parent function that called this cmdlet knows to exit completely
        # using pipeline returns is a challenge as we want to output objects to screen, or allow user to capture pipeline output
        # Ignore VSCode complaining that it is assigned but never used.
        $global:QFMenuExitFlag = $false

        # Script scoped variables. Set these initially based on any parameters passed to the function. 
        # We will also update them whenever the user enters a different value.
        # Script scoped as these could be updated from within a nested function but we want them to be changed across all functions within this cmdlet.
        $script:CasinoID = $CasinoID
        $script:UserID = $UserID
        $script:OperatorID = $OperatorID
        $script:Login = $Login

        # Sets up the PromptForChoice method for the main menu
        $MainMenu = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        # If called from another function, offer Back option as default, otherwise Quit if run directly from PowerShell command line
        if ($PSCmdlet.MyInvocation.CommandOrigin -eq "Internal ") {
            $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Back',"Go back to the previous menu"))
        } else {
            $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
        }
        $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Casino Info',"Request information regarding a Quickfire Casino or Operator"))
        $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Game Info', "Request information regarding a Quickfire Game"))
        $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&User Search',"Search for a Quickfire user"))
        $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Reconciliation',"Reconciliation API functions e.g. managing Commit and Rollback queues"))
        $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&AAMS Status',"Check status of game session with a Participation Code from the AAMS site, for Italian casinos."))
        # If called from another function, Quit is the last option.
        if ($PSCmdlet.MyInvocation.CommandOrigin -eq "Internal") {
            $MainMenu.Add((New-Object Management.Automation.Host.ChoiceDescription '&Quit',"Exit without any further action"))
        }
            
    }

    Process {
        # Display the main menu and prompt for input
        Invoke-MainMenu
    }
    
}