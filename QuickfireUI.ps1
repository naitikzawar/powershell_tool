Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#\/\/\/\/download location here\/\/\/\/
$PSFolderLocation = Join-Path $PSScriptRoot "MIGS-IT-QFPowershell"
#/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

$PSFolderLocation + "\"
Get-ChildItem -Path $PSFolderLocation -Recurse -Filter "*.ps*1" | Unblock-File
$Quickfirepsd1 = Join-Path -Path $PSFolderLocation -ChildPath "Quickfire.psd1"


try { Get-ChildItem -Path $PSFolderLocation -Recurse -ErrorAction Stop | Out-Null }
catch {
    function Grant-FullControl {
        param (
            [Parameter(Mandatory = $true)]
            [string] $Path,
            [Parameter(Mandatory = $true)]
            [string] $User
        )
    
        & icacls $Path /grant ($User + ':(OI)(CI)F') /T
    }
    Grant-FullControl -Path $PSFolderLocation -User $currentUser
}
Import-Module $Quickfirepsd1
#$GetGamesList = Invoke-RestMethod -Method 'Get' -uri "https://casinoportal.gameassists.co.uk/api/Games/List"
$mainform = New-Object System.Windows.Forms.Form
$Icon = Join-Path $PSScriptRoot "Derivco_logo.ico"
$mainform.Icon = New-Object System.Drawing.Icon($Icon)
$mainform.Text = "Derivco"
$mainform.ClientSize = '700,500'
$mainform.StartPosition = 'CenterScreen'
$mainform.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

#################
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'

################## Create SD password tab
$SDPasswordsTab = New-Object System.Windows.Forms.TabPage
$SDPasswordsTab.Text = "SD Passwords"
$tabControl.TabPages.Add($SDPasswordsTab)

if ($PSVersionTable.PSVersion.Major -lt 7 -and $PSVersionTable.PSVersion.Minor -lt 4) {
    $PSVersionLbl = New-Object System.Windows.Forms.Label -Property `
    @{Location = New-Object System.Drawing.Point(10, 10); Text = "Need at least Powershell version 7.4"; AutoSize = $true }
    
    $GetPS7 = New-Object System.Windows.Forms.TextBox
    $GetPS7.Location = New-Object System.Drawing.Point(10, 30)
    $GetPS7.Text = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.1/PowerShell-7.4.1-win-x64.msi"
    $GetPS7.Width = 500

    $SDPasswordsTab.Text = "Powershell Version Check"
    $SDPasswordsTab.Controls.Add($PSVersionLbl)
    $SDPasswordsTab.Controls.Add($GetPS7)
    $mainform.Controls.Add($tabControl)
}
else {

    $CasinoList = 'UATQF1CAS5.ioa.mgsops.com,
GICQF1CAS5.gic.mgsops.com,
GICQF2CAS5.gic.mgsops.com,
MALITQF1CAS5.mal.mgsops.com,
MALQF1CAS5.mal.mgsops.com,
MALQF2CAS5.mal.mgsops.com,
MITQF1CAS5.mit.mgsops.com,
MITQF2CAS5.mit.mgsops.com,
MALQF3CAS5.mal.mgsops.com,
CILQF1CAS1.cil.mgsops.com'

    $SDLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 3; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $SDLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 90); ColumnCount = 2; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $SDPasswordButton = New-Object System.Windows.Forms.Button -Property @{ Autosize = $true; Text = 'Generate'; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    $SDPasswordResetCasinos = New-Object System.Windows.Forms.Button -Property @{AutoSize = $true; Text = 'Reset Tab' }
    $SDPasswordsReasonTxt = New-Object System.Windows.Forms.TextBox -Property @{ Size = New-Object System.Drawing.Size(160, 23); Text = 'Reason' }
    
    $SDPasswordResetCasinos.Add_Click({
            $DSBoxes.Text = $CasinoList
            $SDPasswordsReasonTxt.Text = 'Reason'
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $font = New-Object System.Drawing.Font("Arial", 12)

    $DSBoxes = New-Object System.Windows.Forms.TextBox -Property `
    @{ Multiline = $true; Size = New-Object System.Drawing.Size(300, 300); Font = $font; Text = $CasinoList }
    
    $SDPasswords = New-Object System.Windows.Forms.TextBox -Property `
    @{ Multiline = $true; Size = New-Object System.Drawing.Size(300, 300); Font = $font }
    
    $SDPasswordButton.Add_Click({
            $SDBoxesText = $DSBoxes.Text
            $SDBoxesText = $SDBoxesText.Split(',') | ForEach-Object { $_.Trim() }

            # Function or script block to execute when the button is clicked
            foreach ($box in $SDBoxesText) {
                [string]$SdUsername = 'sa'
                [int]$SdPasswordType = 0
                [string]$SdReason = $SDPasswordsReasonTxt.Text
                [string]$SDURI = "https://sdapi.mgsops.net/ServerDetailsAPI.svc"

                $Headers = @{ 'Content-Type' = 'text/xml'; "SOAPAction" = "http://tempuri.org/IServerDetailsApi/GetServerPassword" }

                $Body = @"
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
            <soapenv:Header/>
            <soapenv:Body>
                <tem:GetServerPassword>
                    <tem:fqdn>$box</tem:fqdn>
                    <tem:username>$SdUsername</tem:username>
                    <tem:reason>$SdReason</tem:reason>
                    <tem:passwordType>$SdPasswordType</tem:passwordType>
                </tem:GetServerPassword>
            </soapenv:Body>
            </soapenv:Envelope>
"@

                $DBPassword = Invoke-RestMethod -Uri "$SDURI/ServerDetailsHttpsAPI.svc" -Method 'POST' -Headers $Headers -Body $Body -UseDefaultCredentials -UseBasicParsing
                $DBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult
                $SDPasswords.Text += $DBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult += "`r`n"

                $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
            }
            $SDPasswords.Text += "`r`n" + $SDPasswordsReasonTxt.Text
        })
    $SDLayoutPanel1.Controls.Add($SDPasswordsReasonTxt, 0, 0)
    $SDLayoutPanel1.Controls.Add($SDPasswordButton, 1, 0)
    $SDLayoutPanel1.Controls.Add($SDPasswordResetCasinos, 2, 0)

    $SDLayoutPanel2.Controls.Add($DSBoxes, 0, 0)
    $SDLayoutPanel2.Controls.Add($SDPasswords, 1, 0)
        
    $SDPasswordsTab.Controls.Add($SDLayoutPanel1)
    $SDPasswordsTab.Controls.Add($SDLayoutPanel2)
    
    ################## End of SD password tab

    ################## Create Traditional SD password tab
    $TradSDPasswordsTab = New-Object System.Windows.Forms.TabPage
    $TradSDPasswordsTab.Text = "Trad SD Passwords"
    $tabControl.TabPages.Add($TradSDPasswordsTab)

    $TradSDLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 3; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TradSDLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 60); ColumnCount = 2; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TradSDPasswordButton = New-Object System.Windows.Forms.Button -Property @{ Autosize = $true; Text = 'Generate'; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    
    $TradSDPasswordsReasonTxt = New-Object System.Windows.Forms.TextBox -Property @{ Size = New-Object System.Drawing.Size(160, 23); Text = 'Reason' }
    
    $TradSDPasswordResetButton = New-Object System.Windows.Forms.Button -Property  @{ AutoSize = $true; Text = 'Reset Tab'; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    $TradBoxList = Join-Path $PSScriptRoot "TradServers.txt"
    $TradBoxListContent = Get-Content $TradBoxList | Out-String
    $TradBoxListContent = $TradBoxListContent -replace ",", ",`n"
    $TradBox = $TradBoxListContent.ToString()

    $font = New-Object System.Drawing.Font("Arial", 12)

    $TradDSBoxes = New-Object System.Windows.Forms.TextBox -Property `
    @{ Multiline = $true; Size = New-Object System.Drawing.Size(500, 300); Font = $font; ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; Text = $TradBox }
    
    $TradSDPasswords = New-Object System.Windows.Forms.TextBox -Property `
    @{Multiline = $true; Size = New-Object System.Drawing.Size(150, 300); Font = $font }

    $TradSDPasswordResetButton.Add_Click({
            $TradDSBoxes.Text = $TradBox
            $TradSDPasswordsReasonTxt.Text = 'Reason'
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $TradSDPasswordButton.Add_Click({
            $TradSDPasswords.Text = ""
            $SDBoxesText = $TradDSBoxes.Text
            $SDBoxesText = $SDBoxesText.Split(',') | ForEach-Object { $_.Trim() }

            foreach ($box in $SDBoxesText) {
                $box = $box.Split('--') | ForEach-Object { $_.Trim() }
                $box = $box[1]
                [string]$SdUsername = 'DBReadOnly_PI'
                [int]$SdPasswordType = 0
                [string]$SdReason = $TradSDPasswordsReasonTxt.Text
                [string]$SDURI = "https://sdapi.mgsops.net/ServerDetailsAPI.svc"

                $Headers = @{ 'Content-Type' = 'text/xml'; "SOAPAction" = "http://tempuri.org/IServerDetailsApi/GetServerPassword" }

                $Body = @"
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
            <soapenv:Header/>
            <soapenv:Body>
                <tem:GetServerPassword>
                    <!--Optional:-->
                    <tem:fqdn>$box</tem:fqdn>
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

                $TradDBPassword = Invoke-RestMethod -Uri "$SDURI/ServerDetailsHttpsAPI.svc" -Method 'POST' -Headers $Headers -Body $Body -UseDefaultCredentials -UseBasicParsing
                $TradDBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult
                $TradSDPasswords.Text += $TradDBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult += "`r`n"

                $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
            }
            $TradSDPasswords.Text += "`r`n" + $TradSDPasswordsReasonTxt.Text
        })
    
    $TradSDLayoutPanel1.Controls.Add($TradSDPasswordButton, 0, 0)
    $TradSDLayoutPanel1.Controls.Add($TradSDPasswordsReasonTxt, 1, 0)
    $TradSDLayoutPanel1.Controls.Add($TradSDPasswordResetButton, 2, 0)

    $TradSDLayoutPanel2.Controls.Add($TradDSBoxes, 0, 0)
    $TradSDLayoutPanel2.Controls.Add($TradSDPasswords, 1, 0)

    $TradSDPasswordsTab.Controls.Add($TradSDLayoutPanel1)
    $TradSDPasswordsTab.Controls.Add($TradSDLayoutPanel2)

    ################## End of SD password tab

    ################## Start of vanguard transaction audit tab
    $VanguardState = New-Object System.Windows.Forms.TabPage
    $VanguardState.Text = "Player transaction audit"

    $TransactionAuditLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 7; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TransactionAuditLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 90); ColumnCount = 4; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TransactionAuditLayoutPanel3 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 150); ColumnCount = 4; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $UserID = New-Object System.Windows.Forms.TextBox -Property @{ Text = "UserID" }
    $CasinoID = New-Object System.Windows.Forms.TextBox -Property @{ Text = "CasinoID" }
    
    $StartTimeDate = New-Object System.Windows.Forms.DateTimePicker
    $StartTimeDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short

    $StartTimeTime = New-Object System.Windows.Forms.DateTimePicker
    $StartTimeTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
    $StartTimeTime.ShowUpDown = $true

    $StartTimeLbl = New-Object System.Windows.Forms.Label -Property @{Text = "From (older date)"; AutoSize = $true }
    $EndTimeLbl = New-Object System.Windows.Forms.Label -Property @{Text = "To (newer date)"; AutoSize = $true }

    $EndTimeDate = New-Object System.Windows.Forms.DateTimePicker
    $EndTimeDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short

    $EndTimeTime = New-Object System.Windows.Forms.DateTimePicker
    $EndTimeTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
    $EndTimeTime.ShowUpDown = $true

    $GetAudit = New-Object System.Windows.Forms.Button -Property @{Text = "Fetch Audit" }
    $GetAuditLbl = New-Object System.Windows.Forms.Label -Property @{ AutoSize = $true; Text = ""; Name = "GetAuditLbl" }
    $ClearAudit = New-Object System.Windows.Forms.Button -Property @{ Width = 90; Text = "Clear Results" }
    $ResetTab = New-Object System.Windows.Forms.Button -Property @{ Width = 90; Text = "Reset Tab" }
    $ExportExcel = New-Object System.Windows.Forms.Button -Property @{ AutoSize = $true; Text = "Export to Excel"; Name = "ExportExcel" }
    $ExportExcelLbl = New-Object System.Windows.Forms.Label -Property @{ AutoSize = $true; Text = ""; Name = "ExportExcelLbl" }
    $NameExcelChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Name Document?"; AutoSize = $true; Name = "NameExcelChkBx" }
    $OpenExcelChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Open export location?"; AutoSize = $true; Name = "OpenExcelChkBx" }
    $NameExcelChkBxTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Document Name"; Name = "NameExcelChkBxTxt" }
    $ModuleIDTxtBx = New-Object System.Windows.Forms.TextBox -Property @{Text = "Module ID"; Name = "ModuleIDTxt" }
    $ModuleIDChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Search by Module ID?"; Name = "ModuleIDChk"; AutoSize = $true }

    $tabControl.TabPages.Add($VanguardState)

    $controls = @(@{Control = $UserID; Column = 0; Row = 0 }, @{Control = $CasinoID; Column = 1; Row = 0 }, @{Control = $ModuleIDChkBx; Column = 2; Row = 0 }, @{Control = $ModuleIDTxtBx; Column = 3; Row = 0 })
    foreach ($item in $controls) { $TransactionAuditLayoutPanel1.Controls.Add($item.Control, $item.Column, $item.Row) }
    $VanguardState.Controls.Add($TransactionAuditLayoutPanel1)

    $controls2 = @(@{Control = $EndTimeLbl; Column = 0; Row = 1 }, @{Control = $EndTimeDate; Column = 1; Row = 1 }, @{Control = $EndTimeTime; Column = 2; Row = 1 }, @{Control = $StartTimeDate; Column = 1; Row = 0 }, @{Control = $StartTimeTime; Column = 2; Row = 0 }, @{Control = $StartTimeLbl; Column = 0; Row = 0 })

    foreach ($item in $controls2) { $TransactionAuditLayoutPanel2.Controls.Add($item.Control, $item.Column, $item.Row) }
    $VanguardState.Controls.Add($TransactionAuditLayoutPanel2)

    $controls3 = @(@{Control = $GetAudit; Column = 0; Row = 0 }, @{Control = $ClearAudit; Column = 1; Row = 0 }, @{Control = $ResetTab; Column = 2; Row = 0 }, @{Control = $GetAuditLbl; Column = 3; Row = 0 })
    foreach ($item in $controls3) { $TransactionAuditLayoutPanel3.Controls.Add($item.Control, $item.Column, $item.Row) }
    $VanguardState.Controls.Add($TransactionAuditLayoutPanel3)
    
    $ExportExcel.Add_Click({

            $StartTimeValue = $StartTimeDate.value
            $StartTimePickerValue = $StartTimeTime.value
            $combinedStartTime = New-Object DateTime -ArgumentList @( $StartTimeValue.Year, $StartTimeValue.Month, $StartTimeValue.Day, $StartTimePickerValue.Hour, $StartTimePickerValue.Minute, 0)
            $formattedStartTime = $combinedStartTime.ToString("yyyy-MM-ddTHH:mm:ssK")

            $EndTimeValue = $EndTimeDate.value
            $EndTimePickerValue = $EndTimeTime.value
            $combinedEndTime = New-Object DateTime -ArgumentList @( $EndTimeValue.Year, $EndTimeValue.Month, $EndTimeValue.Day, $EndTimePickerValue.Hour, $EndTimePickerValue.Minute, 0)
            $formattedEndTime = $combinedEndTime.ToString("yyyy-MM-ddTHH:mm:ssK")

            $ExportExcelLbl.Text = "Processing"
            $ExportExcelLbl.Refresh()

            if ($formattedEndTime -lt $formattedStartTime) {
                $GetAuditLbl.Text = "Older date must be older"
                $ExportExcelLbl.Text = ""
                return
            }
            else {
                $GetOperatorFromCasino = Invoke-QFPortalRequest -CasinoID $CasinoID.Text
                $OP_Key = Get-QFOperatorAPIKeys -OperatorID $GetOperatorFromCasino.operatorID
                $OP_Token = Get-QFOperatorToken -APIKey $OP_key[0].APIKey
                $startDate = $formattedStartTime
                $EndtDate = $formattedEndTime
                if ($NameExcelChkBx.Checked -eq $true) {
                    
                    $fileName = $NameExcelChkBxTxt.Text.Trim() + ".xlsx"
                    if ($ModuleIDChkBx.Checked -eq $true) {
                        $ExportExcelLbl.Text = "Fetching audit data for 1st sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -ModuleID $ModuleIDTxtBx.Text.Trim()
                        Export-QFExcel -ExcelData $AuditData -ExcelFileName $fileName

                        $ExportExcelLbl.Text = "Fetching audit data for 2nd sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -FinancialAudit 
                        $ExportExcelLbl.Text = "Exporting to Excel"
                        Export-QFExcel -ExcelData $AuditData -ExcelFileName $fileName -ExcelSourceWorksheetName "Financial Audit"
                    }
                    else {
                        $ExportExcelLbl.Text = "Fetching audit data for 1st sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate
                        Export-QFExcel -ExcelData $AuditData -ExcelFileName $fileName
    
                        $ExportExcelLbl.Text = "Fetching audit data for 2nd sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -FinancialAudit 
                        $ExportExcelLbl.Text = "Exporting to Excel"
                        Export-QFExcel -ExcelData $AuditData -ExcelFileName $fileName -ExcelSourceWorksheetName "Financial Audit"
                    }
                    
                }
                else {
                    if ($ModuleIDChkBx.Checked -eq $true) {
                        $ExportExcelLbl.Text = "Fetching audit data for 1st sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -ModuleID $ModuleIDTxtBx.Text.Trim()
                        Export-QFExcel -ExcelData $AuditData 

                        $ExportExcelLbl.Text = "Fetching audit data for 2nd sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -FinancialAudit 
                        $ExportExcelLbl.Text = "Exporting to Excel"
                        Export-QFExcel -ExcelData $AuditData -ExcelSourceWorksheetName "Financial Audit"
                    }
                    else {
                        $ExportExcelLbl.Text = "Fetching audit data for 1st sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate
                        Export-QFExcel -ExcelData $AuditData 
    
                        $ExportExcelLbl.Text = "Fetching audit data for 2nd sheet"
                        $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -FinancialAudit 
                        $ExportExcelLbl.Text = "Exporting to Excel"
                        Export-QFExcel -ExcelData $AuditData -ExcelSourceWorksheetName "Financial Audit"
                    }
                }
                
                if ($OpenExcelChkBx.Checked -eq $true) { 
                    explorer.exe (Get-Location).path 
                }
                $ExportExcelLbl.Refresh()
                $ExportExcelLbl.Text = "Complete!"
            }
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $ResetTab.Add_Click({
            $UserID.Text = "UserID"
            $CasinoID.Text = "CasinoID"
            $GetAuditLbl.Text = ""
            $ModuleIDTxtBx.Text = "ModuleID"
            $OpenExcelChkBx.Checked = $false
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $ClearAudit.Add_Click({
            $controlsToRemove = New-Object System.Collections.Generic.List[System.Windows.Forms.Control]

            $controlNames = @("TransactionAuditLayoutPanel4", "TransactionAuditLayoutPanel5")

            foreach ($control in $VanguardState.Controls) {
                if ($control.Name -in $controlNames) { $controlsToRemove.Add($control) }
                $GetAuditLbl.Text = ""
            }
            foreach ($control in $controlsToRemove) { $VanguardState.Controls.Remove($control) }
        })

    $GetAudit.Add_Click({
            $GetOperatorFromCasino = Invoke-QFPortalRequest -CasinoID $CasinoID.Text

            $StartTimeValue = $StartTimeDate.value
            $StartTimePickerValue = $StartTimeTime.value
            $combinedStartTime = New-Object DateTime -ArgumentList @( $StartTimeValue.Year, $StartTimeValue.Month, $StartTimeValue.Day, $StartTimePickerValue.Hour, $StartTimePickerValue.Minute, 0)
            $formattedStartTime = $combinedStartTime.ToString("yyyy-MM-ddTHH:mm:ssK")

            $EndTimeValue = $EndTimeDate.value
            $EndTimePickerValue = $EndTimeTime.value
            $combinedEndTime = New-Object DateTime -ArgumentList @( $EndTimeValue.Year, $EndTimeValue.Month, $EndTimeValue.Day, $EndTimePickerValue.Hour, $EndTimePickerValue.Minute, 0)
            $formattedEndTime = $combinedEndTime.ToString("yyyy-MM-ddTHH:mm:ssK")

            $OP_Key = Get-QFOperatorAPIKeys -OperatorID $GetOperatorFromCasino.operatorID
            $OP_Token = Get-QFOperatorToken -APIKey $OP_key[0].APIKey

            if ($formattedEndTime -lt $formattedStartTime) {
                $GetAuditLbl.Text = "Older date must be older"
                return
            }

            $GetAuditLbl.Text = "Processing"
            $GetAuditLbl.Refresh()

            $startDate = $formattedStartTime
            $EndtDate = $formattedEndTime

            if ($ModuleIDChkBx.Checked -eq $true) {
                $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate -ModuleID $ModuleIDTxtBx.Text.Trim()
            }
            else {
                $AuditData = Get-QFAudit -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -UserID $UserID.Text -CasinoID $CasinoID.Text -StartDate $startDate -EndDate $EndtDate
            }

            $dataTable = New-Object System.Data.DataTable
            $dataGridView = New-Object System.Windows.Forms.DataGridView
            $dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
            $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

            $columns = @('actionTime', 'amount', 'CID', 'currencyCode', 'externalActionId', 'ExternalGameName', 'externalReference', 'MID', 'NoOfAttempts', 'productId', 'source', 'statusDescription', 'statusId', 'TransNo', 'TransTime', 'TransType', 'UID', 'userName')

            foreach ($column in $columns) { $dataTable.Columns.Add($column) }
        
            $A = @()

            foreach ($tran in $AuditData) {
                $tranMember = $tran | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                foreach ($item in $tranMember) {
                    $itemValue = $tran.$item
                    $A += $itemValue
                }
                $dataTable.Rows.Add($A[0], $A[1], $A[2], $A[3], $A[4], $A[5], $A[6], $A[7], $A[8], $A[9], $A[10], $A[11], $A[12], $A[13], $A[14], $A[15], $A[16], $A[17]);
                $A = @()
            }

            $dataGridView.Width = 600
            $dataGridView.Height = 200
            $dataGridView.DataSource = $dataTable

            $TransactionAuditLayoutPanel4 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
            @{Location = New-Object System.Drawing.Point(30, 190); ColumnCount = 1; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink; `
                    Name = "TransactionAuditLayoutPanel4"
            }

            $TransactionAuditLayoutPanel5 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
            @{Location = New-Object System.Drawing.Point(30, 410); ColumnCount = 5; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink; `
                    Name = "TransactionAuditLayoutPanel5"
            }

            $TransactionAuditLayoutPanel4.Controls.Add($dataGridView, 0, 0)
            $VanguardState.Controls.Add($TransactionAuditLayoutPanel4)

            $TransactionAuditLayoutPanel5.Controls.Add($ExportExcel, 3, 0)
            $TransactionAuditLayoutPanel5.Controls.Add($ExportExcelLbl, 4, 0)
            $TransactionAuditLayoutPanel5.Controls.Add($NameExcelChkBx, 0, 0)
            $TransactionAuditLayoutPanel5.Controls.Add($NameExcelChkBxTxt, 1, 0)
            $TransactionAuditLayoutPanel5.Controls.Add($OpenExcelChkBx, 2, 0)
            $VanguardState.Controls.Add($TransactionAuditLayoutPanel5)
            

            $GetAuditLbl.Text = "Complete!"
            $GetAuditLbl.Refresh()
            $VanguardState.Refresh()
        })
    ################## End of vanguard transaction audit tab

    ################## Start of lookup player tab
    $QFPlayerLookup = New-Object System.Windows.Forms.TabPage
    $QFPlayerLookup.Text = "Locate player"
    $tabControl.TabPages.Add($QFPlayerLookup)

    $QFPlayerLookupLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 2; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $QFPlayerLookupLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 60); ColumnCount = 4; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $PlayerSearchTerm = New-Object System.Windows.Forms.TextBox -Property @{Text = "Login Name"; Width = 100 }
    $PlayerSearchOperatorID = New-Object System.Windows.Forms.TextBox -Property @{Text = "Operator ID"; Width = 100 }
    $PlayerSearchButton = New-Object System.Windows.Forms.Button -Property @{Text = "Search player"; AutoSize = $true }
    $PlayerSearchingLbl = New-Object System.WIndows.Forms.Label -Property @{AutoSize = $true }
    $ClearPlayer = New-Object System.Windows.Forms.Button -Property @{Text = "Clear Results"; AutoSize = $true }
    $PlayerSearchResetTab = New-Object System.Windows.Forms.Button -Property @{Text = "Reset Tab"; AutoSize = $true }

    $QFPlayerLookupLayoutPanel1.Controls.Add($PlayerSearchTerm, 0, 0)
    $QFPlayerLookupLayoutPanel1.Controls.Add($PlayerSearchOperatorID, 1, 0)
    $QFPlayerLookup.Controls.Add($QFPlayerLookupLayoutPanel1)

    $QFPlayerLookupLayoutPanel2.Controls.Add($PlayerSearchButton, 0, 0)
    $QFPlayerLookupLayoutPanel2.Controls.Add($ClearPlayer, 1, 0)
    $QFPlayerLookupLayoutPanel2.Controls.Add($PlayerSearchResetTab, 2, 0)
    $QFPlayerLookupLayoutPanel2.Controls.Add($PlayerSearchingLbl, 3, 0)
    $QFPlayerLookup.Controls.Add($QFPlayerLookupLayoutPanel2)
    

    $PlayerSearchResetTab.Add_Click({
            $PlayerSearchTerm.Text = "Login Name"
            $PlayerSearchOperatorID.Text = "Operator ID"
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $ClearPlayer.Add_Click({
            $controlsToRemove = New-Object System.Collections.Generic.List[System.Windows.Forms.Control]

            foreach ($control in $QFPlayerLookup.Controls) {
                if ($control -is [System.Windows.Forms.DataGridView]) { $controlsToRemove.Add($control) }
            }

            foreach ($control in $controlsToRemove) { $QFPlayerLookup.Controls.Remove($control) }
        })

    $PlayerSearchButton.Add_Click({

            if ($PlayerSearchTerm.Text -eq "Login Name" -or $PlayerSearchTerm.Text -eq "" -or $PlayerSearchOperatorID.Text -eq "Operator ID" -or $PlayerSearchOperatorID.Text -eq "") {
                $PlayerSearchingLbl.Text = "Player login name and Operator ID cannot be blank"
            }
            else {
                $PlayerSearchingLbl.Text = "Searching... this may take a while"
                function SwitchUser {
                    [CmdletBinding()]
                    param(
                        [Parameter(Mandatory = $true)]
                        [string]$Token,
    
                        [Parameter(Mandatory = $true)]
                        [string]$OperatorID
                    )
    
                    $SwitchOperatorBody = @{ "operatorId" = $OperatorID.Trim() }
                    $UserSearchAppHeader = @{ "Authorization" = "UserSession " + $Token }
                    $NewLoginToken = Invoke-RestMethod -Method Post -uri "https://quickfireapp.gameassists.co.uk/Framework/api/Security/SwitchUser" -Body $SwitchOperatorBody -Headers $UserSearchAppHeader
                    return $NewLoginToken.sessionToken
                }
    
    
                $UserSearchLoginBody = @{ 
                    "loginName" = "alex.bowker@derivco.co.im"
                    "password"  = "pas$\/\/o7d6" 
                }
                $PlayerSearchingLbl.Text = "Logging in"
                $LoginResult = Invoke-RestMethod -Method 'Post' -uri "https://quickfireapp.gameassists.co.uk/Framework/api/Security/Login" -Body $UserSearchLoginBody
                $InitialSesstionToken = $LoginResult.sessionToken

                $PlayerSearchingLbl.Text = "Swapping operator"
                SwitchUser -Token $InitialSesstionToken -OperatorID $PlayerSearchOperatorID.Text.Trim()

                $PlayerSearchingLbl.Text = "Logging in"
                $LoginResult = Invoke-RestMethod -Method 'Post' -uri "https://quickfireapp.gameassists.co.uk/Framework/api/Security/Login" -Body $UserSearchLoginBody
                $NewSesstionToken = $LoginResult.sessionToken
    
                $NewSessionLogin = @{ "Authorization" = "UserSession " + $NewSesstionToken }
    
                $UserSearchBody = @{ "exactMatch" = "false"
                    "pageNumber"                  = "1"
                    "rowsPerPage"                 = "25"
                    "searchPhrase"                = $PlayerSearchTerm.Text.Trim() 
                }
                $PlayerSearchingLbl.Text = "Searching player"
                $PlayerSearch = Invoke-RestMethod -Method Post -uri "https://quickfireapp.gameassists.co.uk/helpdeskexpress/players/search" -Headers $NewSessionLogin -Body $UserSearchBody
    
                $playerGridView = New-Object System.Windows.Forms.DataGridView -Property `
                @{Location = New-Object System.Drawing.Point(30, 120); ScrollBars = [System.Windows.Forms.ScrollBars]::Both; Width = 500; Height = 100; AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells }
                
                $playerDataTable = New-Object System.Data.DataTable
                $playerDataTableColumns = @('productId', 'userId', 'productName', 'player login')

                $productIDs = @()
    
                foreach ($column in $playerDataTableColumns) { $playerDataTable.Columns.Add($column) }
                
                foreach ($player in $PlayerSearch.searchRecords) { 
                    if ($productIDs -notcontains $player.productId ) { $productIDs += $player.productId } 
                }
                    
                if ($productIDs.Length -eq 0) {
                    $PlayerSearchingLbl.Text = "No player located or operator not mapped"
                }
                else {
                    $prefixes = Invoke-QFPortalRequest -CasinoID $productIDs
                    $casinoPrefix = $prefixes.productSettings.StringValue -join ""
                    foreach ($player in $PlayerSearch.searchRecords) { 
                        $playerUsername = $casinoPrefix + $player.username
                        $playerDataTable.Rows.Add($player.productId, $player.userId, $player.productName, $playerUsername); 
                    }
    
                    $playerGridView.DataSource = $playerDataTable
                    $QFPlayerLookup.Controls.Add($playerGridView)
                    $PlayerSearchingLbl.Text = ""
                }
                
            }
        })
    ################## End of lookup player tab

    ################## Start of playcheck tab
    $playcheck = New-Object System.Windows.Forms.TabPage
    $playcheck.Text = "Playcheck"
    $tabControl.TabPages.Add($playcheck)

    $PSServerTxt = New-Object System.Windows.Forms.TextBox -Property @{ Text = "Server ID"; Location = New-Object System.Drawing.Point(30, 30); Width = 100; }
    $PSUserLoginTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Player login"; Location = New-Object System.Drawing.Point(140, 30); Width = 300; }
    $PSUserTransTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Trans IDs (123,234,345)"; Location = New-Object System.Drawing.Point(30, 60); Width = 300; }
    $PSOpenExplorer = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Open in explorer"; Location = New-Object System.Drawing.Point(30, 90); AutoSize = $true }
    $PSPDFName = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Name PDF?"; Location = New-Object System.Drawing.Point(150, 90); Width = 90 }
    $PSPDFNameTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "PDF name"; Location = New-Object System.Drawing.Point(250, 90); Width = 250 }
    $PSGetPlaycheckButton = New-Object System.Windows.Forms.Button -Property @{Text = "Fetch playcheck(s)"; Location = New-Object System.Drawing.Point(30, 120); AutoSize = $true }
    $PSGetPlaycheckResetTab = New-Object System.Windows.Forms.Button -Property @{Text = "Reset Tab"; Location = New-Object System.Drawing.Point(150, 120); AutoSize = $true }
    $PSOutputTxtArea = New-Object System.Windows.Forms.TextBox -Property @{Multiline = $true; Location = New-Object System.Drawing.Point(30, 160); Width = 630; Height = 300; Name = "PSOutputTxtArea" }
    

    $PSGetPlaycheckResetTab.Add_Click({
            $playcheck.Refresh()
            $PSServerTxt.Text = "Server ID"
            $PSUserLoginTxt.Text = "Player login"
            $PSUserTransTxt.Text = "Trans IDs (123,234,345)"
            $PSPDFNameTxt.Text = "PDF name"
            $PSPDFName.Checked = $false
            $PSOpenExplorer.Checked = $false
            $controlsToRemove = @()
        })

    $PSGetPlaycheckButton.Add_Click({
            $playcheck.Controls.Add($PSOutputTxtArea)
            $PSOutputTxtArea.Text = ""
            $PSOutputTxtArea.Text += "Please wait `r`n"
            $PSUserTransTxt = $PSUserTransTxt.Text.Trim()
            $PSUserTransArr = $PSUserTransTxt -split ',' | ForEach-Object { [int]$_ }
        
            function CheckPlaycheckSaved {
                param( [string]$TestPath, [string]$TransactionNo)
                write-host $TestPath
                if ($TestPath -like "*\") {
                    $PSOutputTxtArea.Text += "Unable to generate PDF for transaction $TransactionNo `r`n"
                }
                else {
                    if (Test-Path -Path $TestPath) {
                        $PSOutputTxtArea.Text += "Complete!, Saved here: $TestPath `r`n"
                    }
                    else {
                        $PSOutputTxtArea.Text += "Unable to generate PDF for transaction $TransactionNo `r`n"
                    }
                }
            }

            foreach ($Transaction in $PSUserTransArr) {
                $userDefinedPlaycheckName = "$($PSPDFNameTxt.Text)_$Transaction.pdf"
                $userDefinedPlaycheckNameFilePath = Join-Path -Path $PSScriptRoot -ChildPath $userDefinedPlaycheckName
                #Name PDF
                if ($PSPDFName.Checked -eq $true -and $PSOpenExplorer.Checked -eq $false) {
                    $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                    Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF -FileName $PSPDFNameTxt.Text.Trim()
                    CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                }
                #Open in explorer
                if ($PSOpenExplorer.Checked -eq $true -and $PSPDFName.Checked -eq $false) {
                    if ($Transaction -eq $PSUserTransArr[$PSUserTransArr.Count - 1]) {
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF -OpenExplorer
                        CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                    else {
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF
                        CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                }
                #Open in explorer & name the PDF
                if ($PSOpenExplorer.Checked -eq $true -and $PSPDFName.Checked -eq $true) {
                    if ($Transaction -eq $PSUserTransArr[$PSUserTransArr.Count - 1]) {
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF -OpenExplorer -FileName $PSPDFNameTxt.Text.Trim()
                        CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                    else {
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF -FileName $PSPDFNameTxt.Text.Trim()
                        CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                }
                #neither open in explorer or name PDF
                if ($PSOpenExplorer.Checked -eq $false -and $PSPDFName.Checked -eq $false) {
                    $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                    $autoPlaycheckName = "PlayCheck $Transaction.pdf"

                    Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF
                    
                    $autoPlaycheckNameFilePath = Join-Path -Path $PSScriptRoot -ChildPath $autoPlaycheckName
                    if (Test-Path -Path $autoPlaycheckNameFilePath) {
                        $PSOutputTxtArea.Text += "Complete! `r`n"
                    }
                    else {
                        $PSOutputTxtArea.Text += "Error retrieving playcheck for transaction $Transaction) `r`n"
                    }
                
                }
            }
            $PSOutputTxtArea.Text += "All transactions processed"
        })
    $playcheckControls = $PSServerTxt, $PSUserLoginTxt, $PSUserTransTxt, $PSOpenExplorer, $PSPDFName, $PSPDFNameTxt, $PSGetPlaycheckButton, $PSGetPlaycheckResetTab

    foreach ($control in $playcheckControls) { $playcheck.Controls.Add($control) }

    ################## End of playcheck tab
    $mainform.Controls.Add($tabControl)
}
$mainform.ShowDialog()