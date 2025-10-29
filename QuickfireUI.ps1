#https://confluence.derivco.co.za/display/DE/ES+-+Quality+Gate+Reference
#https://confluence.derivco.co.za/pages/viewpage.action?pageId=34571320

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
$mainform = New-Object System.Windows.Forms.Form
$Icon = Join-Path $PSScriptRoot "Derivco_logo.ico"
$mainform.Icon = New-Object System.Drawing.Icon($Icon)
$mainform.Text = "Derivco"
$mainform.ClientSize = '900,500'
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
MALQF3CAS5.mal.mgsops.com,
MITQF1CAS5.mit.mgsops.com,
MITQF2CAS5.mit.mgsops.com,
MITQF3CAS5.mit.mgsops.com,
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
    ################## Start of Vanguard passwords tab
    $VanguardSDPasswordsTab = New-Object System.Windows.Forms.TabPage
    $VanguardSDPasswordsTab.Text = "Vanguard Passwords"
    $tabControl.TabPages.Add($VanguardSDPasswordsTab)

    $VnaguardSDLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 3; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $VnaguardSDLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 60); ColumnCount = 2; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $VanguardSDPasswordButton = New-Object System.Windows.Forms.Button -Property @{ Autosize = $true; Text = 'Generate'; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    
    $VanguardSDPasswordsReasonTxt = New-Object System.Windows.Forms.TextBox -Property @{ Size = New-Object System.Drawing.Size(160, 23); Text = 'Reason' }
    
    $VanguardSDPasswordResetButton = New-Object System.Windows.Forms.Button -Property  @{ AutoSize = $true; Text = 'Reset Tab'; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    
    $VanguardList = 
    'CILCNTCFG1.cil.mgsops.com,
GICSH1VGDB5.gic.mgsops.com,
MALIT1VGDB5.mal.mgsops.com,
MALQF1VGDB5.mal.mgsops.com,
MITSH1VGDB5.mit.mgsops.com,
UATSH1VGDB5.ioa.mgsops.com'

    $font = New-Object System.Drawing.Font("Arial", 12)

    $VanguardDSBoxes = New-Object System.Windows.Forms.TextBox -Property `
    @{ Multiline = $true; Size = New-Object System.Drawing.Size(300, 200); Font = $font; ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; Text = $VanguardList }
    
    $VanguardSDPasswords = New-Object System.Windows.Forms.TextBox -Property `
    @{Multiline = $true; Size = New-Object System.Drawing.Size(200, 200); Font = $font }

    $VanguardSDPasswordResetButton.Add_Click({
            $VanguardDSBoxes.Text = $VanguardList
            $VanguardSDPasswordsReasonTxt.Text = 'Reason'
            $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
        })

    $VanguardSDPasswordButton.Add_Click({
            $VanguardSDPasswords.Text = ""
            $VanguardDSBoxesText = $VanguardDSBoxes.Text
            $VanguardDSBoxesText = $VanguardDSBoxesText.Split(',') | ForEach-Object { $_.Trim() }

            foreach ($box in $VanguardDSBoxesText) {
                [string]$SdUsername = 'sa'
                [int]$SdPasswordType = 0
                [string]$SdReason = $VanguardSDPasswordsReasonTxt.Text
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

                $VanguardDBPassword = Invoke-RestMethod -Uri "$SDURI/ServerDetailsHttpsAPI.svc" -Method 'POST' -Headers $Headers -Body $Body -UseDefaultCredentials -UseBasicParsing
                $VanguardDBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult
                $VanguardSDPasswords.Text += $VanguardDBPassword.Envelope.body.GetServerPasswordResponse.GetServerPasswordResult += "`r`n"

                $mainform.DialogResult = [System.Windows.Forms.DialogResult]::None
            }
            $VanguardSDPasswords.Text += "`r`n" + $VanguardSDPasswordsReasonTxt.Text
        })

    $VnaguardSDLayoutPanel1.Controls.Add($VanguardSDPasswordButton)
    $VnaguardSDLayoutPanel1.Controls.Add($VanguardSDPasswordsReasonTxt)
    $VnaguardSDLayoutPanel1.Controls.Add($VanguardSDPasswordResetButton)

    $VnaguardSDLayoutPanel2.Controls.Add($VanguardDSBoxes)
    $VnaguardSDLayoutPanel2.Controls.Add($VanguardSDPasswords)

    $VanguardSDPasswordsTab.Controls.Add($VnaguardSDLayoutPanel1)
    $VanguardSDPasswordsTab.Controls.Add($VnaguardSDLayoutPanel2)
    ################## End of Vanguard passwords tab
    ################## Start of vanguard transaction audit tab
    $VanguardState = New-Object System.Windows.Forms.TabPage
    $VanguardState.Text = "Transaction Audit"

    $TransactionAuditLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 7; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TransactionAuditLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 90); ColumnCount = 4; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $TransactionAuditLayoutPanel3 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 150); ColumnCount = 5; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

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
    $AuditNameExcelChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Name audit?"; AutoSize = $true; Name = "NameExcelChkBx" }
    $OpenExcelChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Open export location?"; AutoSize = $true; Name = "OpenExcelChkBx" }
    $NameExcelChkBxTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Document name"; Name = "NameExcelChkBxTxt" }
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
                if ($AuditNameExcelChkBx.Checked -eq $true) {
                    
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

            try {
                $OP_Key = Get-QFOperatorAPIKeys -OperatorID $GetOperatorFromCasino.operatorID
                write-host "OP_Key "$OP_Key
            }
            catch {
                $GetAuditLbl.Text = "Failed to fetch operator bearer token, please try again"
                return
            }
            $OP_Token = Get-QFOperatorToken -APIKey $OP_key[0].APIKey
            

            if ($formattedEndTime -lt $formattedStartTime) {
                $GetAuditLbl.Text = "Older date must be older"
                return
            }

            if ($formattedEndTime -eq $formattedStartTime) {
                $GetAuditLbl.Text = "Entered dates cannot be the same"
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
        
            if ($null -eq $AuditData) {
                $GetAuditLbl.Text = "No gameplay found for date range"
            }
            else {

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

                $TransactionAuditLayoutPanel5.Controls.Add($AuditNameExcelChkBx, 0, 0)
                $TransactionAuditLayoutPanel5.Controls.Add($NameExcelChkBxTxt, 1, 0)
                $TransactionAuditLayoutPanel5.Controls.Add($OpenExcelChkBx, 2, 0)
                $TransactionAuditLayoutPanel5.Controls.Add($ExportExcel, 3, 0)
                $TransactionAuditLayoutPanel5.Controls.Add($ExportExcelLbl, 4, 0)
                $VanguardState.Controls.Add($TransactionAuditLayoutPanel5)

                $TransactionAuditLayoutPanel5.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
                    $_.Add_MouseClick({ $this.Text = "" })
                }
            
                $GetAuditLbl.Text = "Complete! Please wait for results to appear"
                $GetAuditLbl.Refresh()
                $VanguardState.Refresh()
            }
        })
        

    $TransactionAuditLayoutPanel1.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
        $_.Add_MouseClick({ $this.Text = "" })
    }
        
    ################## End of vanguard transaction audit tab

    ################## Start of lookup player tab
    $QFPlayerLookup = New-Object System.Windows.Forms.TabPage
    $QFPlayerLookup.Text = "Locate player"
    $tabControl.TabPages.Add($QFPlayerLookup)

    $QFPlayerLookupLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 2; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $QFPlayerLookupLayoutPanel2 = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 60); ColumnCount = 4; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $PlayerSearchTerm = New-Object System.Windows.Forms.TextBox -Property @{Text = "Login Name"; Width = 300 }
    $PlayerSearchOperatorID = New-Object System.Windows.Forms.TextBox -Property @{Text = "Operator ID"; Width = 100 }
    $PlayerSearchButton = New-Object System.Windows.Forms.Button -Property @{Text = "Search player"; AutoSize = $true }
    $PlayerSearchingLbl = New-Object System.WIndows.Forms.Label -Property @{AutoSize = $true }
    $ClearPlayer = New-Object System.Windows.Forms.Button -Property @{Text = "Clear Results"; AutoSize = $true }
    $PlayerSearchResetTab = New-Object System.Windows.Forms.Button -Property @{Text = "Reset Tab"; AutoSize = $true }

    $controlsLayoutPanel1 = @( @{ Control = $PlayerSearchTerm; Column = 0; Row = 0 }, @{ Control = $PlayerSearchOperatorID; Column = 1; Row = 0 } )
    foreach ($item in $controlsLayoutPanel1) { $QFPlayerLookupLayoutPanel1.Controls.Add($item.Control, $item.Column, $item.Row) }
    $QFPlayerLookup.Controls.Add($QFPlayerLookupLayoutPanel1)

    ####
    $controlsLayoutPanel2 = @( 
        @{ Control = $PlayerSearchButton; Column = 0; Row = 0 }, @{ Control = $ClearPlayer; Column = 1; Row = 0 }, @{ Control = $PlayerSearchResetTab; Column = 2; Row = 0 }, @{ Control = $PlayerSearchingLbl; Column = 3; Row = 0 }
    )
    
    foreach ($item in $controlsLayoutPanel2) { $QFPlayerLookupLayoutPanel2.Controls.Add($item.Control, $item.Column, $item.Row) }
    
    $QFPlayerLookup.Controls.Add($QFPlayerLookupLayoutPanel2)
    ######
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
                    "password"  = "P@55w0rD$" 
                }
                $PlayerSearchingLbl.Text = "Logging in"
                try {
                    $LoginResult = Invoke-RestMethod -Method 'Post' -uri "https://quickfireapp.gameassists.co.uk/Framework/api/Security/Login" -Body $UserSearchLoginBody
                }
                catch {
                    write-host $_
                    $PlayerSearchingLbl.Text = "Operator not mapped"
                    return
                }
                
                $InitialSesstionToken = $LoginResult.sessionToken

                $PlayerSearchingLbl.Text = "Swapping operator"
                SwitchUser -Token $InitialSesstionToken -OperatorID $PlayerSearchOperatorID.Text.Trim()

                $PlayerSearchingLbl.Text = "Logging in"
                try {
                    $LoginResult = Invoke-RestMethod -Method 'Post' -uri "https://quickfireapp.gameassists.co.uk/Framework/api/Security/Login" -Body $UserSearchLoginBody
                }
                catch {
                    $PlayerSearchingLbl.Text = "Operator not mapped"
                }
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
                    $PlayerSearchingLbl.Text = "No player located"
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

    $QFPlayerLookupLayoutPanel1.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
        $_.Add_MouseClick({ $this.Text = "" })
    }

    $QFPlayerLookupLayoutPanel2.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
        $_.Add_MouseClick({ $this.Text = "" })
    }
    ################## End of lookup player tab

    ################## Start of playcheck tab
    $playcheck = New-Object System.Windows.Forms.TabPage
    $playcheck.Text = "Playcheck"
    $tabControl.TabPages.Add($playcheck)

    $PSServerTxt = New-Object System.Windows.Forms.TextBox -Property @{ Text = "Server ID"; Location = New-Object System.Drawing.Point(30, 30); Width = 100; }
    $PSUserLoginTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Player login"; Location = New-Object System.Drawing.Point(140, 30); Width = 300; }
    $PSUserTransTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Trans IDs (123,234,345)"; Location = New-Object System.Drawing.Point(30, 60); Width = 300; }
    $PSUserTransBatchingOptionTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "Transactions to process, default is all"; Location = New-Object System.Drawing.Point(350, 60); Width = 200; }
    $PSUserTransBatchingLbl = New-Object System.WIndows.Forms.Label -Property @{Text = "- Left-most transaction is number 0"; Location = New-Object System.Drawing.Point(560, 60); AutoSize = $true }
    $PSOpenExplorer = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Open in explorer"; Location = New-Object System.Drawing.Point(30, 90); AutoSize = $true }
    $PSPDFName = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Name PDF?"; Location = New-Object System.Drawing.Point(150, 90); Width = 90 }
    $PSPDFNameTxt = New-Object System.Windows.Forms.TextBox -Property @{Text = "PDF name"; Location = New-Object System.Drawing.Point(250, 90); Width = 250 }
    $PSGetPlaycheckButton = New-Object System.Windows.Forms.Button -Property @{Text = "Fetch playcheck(s)"; Location = New-Object System.Drawing.Point(30, 120); AutoSize = $true }
    $PSGetPlaycheckResetTab = New-Object System.Windows.Forms.Button -Property @{Text = "Reset Tab"; Location = New-Object System.Drawing.Point(150, 120); AutoSize = $true }
    $PSOutputTxtArea = New-Object System.Windows.Forms.TextBox -Property @{Multiline = $true; Location = New-Object System.Drawing.Point(30, 160); Width = 630; Height = 300; Name = "PSOutputTxtArea" }
    
    $PSServerTxt.Add_MouseClick({ $this.text = "" })
    $PSUserLoginTxt.Add_MouseClick({ $this.text = "" })
    $PSUserTransTxt.Add_MouseClick({ $this.text = "" })
    $PSPDFNameTxt.Add_MouseClick({ $this.text = "" })

    $PSGetPlaycheckResetTab.Add_Click({
            $playcheck.Refresh()
            $PSServerTxt.Text = "Server ID"
            $PSUserLoginTxt.Text = "Player login"
            $PSUserTransTxt.Text = "Trans IDs (123,234,345)"
            $PSPDFNameTxt.Text = "PDF name"
            $PSPDFName.Checked = $false
            $PSOpenExplorer.Checked = $false
            $controlsToRemove = @()
            $PSUserTransBatchingOptionTxt.Text = "Transactions to process, default is all"
        })

    $PSGetPlaycheckButton.Add_Click({
            $playcheck.Controls.Add($PSOutputTxtArea)
            $PSOutputTxtArea.Text = ""
            $PSOutputTxtArea.Text += "Please wait `r`n"
            $PSUserTransTxt = $PSUserTransTxt.Text.Trim()
            $PSUserTransArr = $PSUserTransTxt -split ',' | ForEach-Object { [int]$_ }
        
            function CheckPlaycheckSaved {
                param( [string]$TestPath, [string]$TransactionNo)
                if ($TestPath -like "*\") {
                    $PSOutputTxtArea.Text += "Unable to generate PDF for transaction $TransactionNo `r`n"
                }
                else {
                    if (Test-Path -Path $TestPath) {
                        $PSOutputTxtArea.Text += "Complete!, saved here: $TestPath `r`n"
                    }
                    else {
                        $PSOutputTxtArea.Text += "Unable to generate PDF for transaction $TransactionNo `r`n"
                    }
                }
            }
            $SpecificTransactionArrayIndexes = @()
            $SpecificTransactionArray = @()

            if ($PSUserTransBatchingOptionTxt.Text.Substring(0, 1) -ne "T") {
                try {
                    write-host
                    $SpecificTransactionArrayIndexes = $PSUserTransBatchingOptionTxt.Text -split ',' | ForEach-Object { 
                        if (-not [int]::TryParse($_, [ref]$null)) {
                            throw "Invalid selection"
                        }
                        [int]$_ 
                    }
                }
                catch {
                    $PSUserTransBatchingOptionTxt.Text = "Invalid selection"
                    return
                }

                $invalidIndexes = @()
                foreach ($index in $SpecificTransactionArrayIndexes) {
                    if ($index -ge 0 -and $index -lt $PSUserTransArr.Count) {
                        $SpecificTransactionArray += $PSUserTransArr[$index]
                    }
                    else {
                        $invalidIndexes += $index
                    }
                }

                if ($invalidIndexes.Count -gt 0) {
                    $PSOutputTxtArea.Text += "Index $invalidIndexes is invalid `r`n"
                }

                $PSUserTransArr = $SpecificTransactionArray
            }
            
            foreach ($Transaction in $PSUserTransArr) {
                $autoPlaycheckName = "PlayCheck $Transaction.pdf"
                $userDefinedPlaycheckName = "$($PSPDFNameTxt.Text)_$Transaction.pdf"
                $userDefinedPlaycheckNameFilePath = Join-Path -Path $PSScriptRoot -ChildPath $userDefinedPlaycheckName
                $autoPlaycheckNameFilePath = Join-Path -Path $PSScriptRoot -ChildPath $autoPlaycheckName

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
                        CheckPlaycheckSaved -TestPath $autoPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                    else {
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF
                        CheckPlaycheckSaved -TestPath $autoPlaycheckNameFilePath -TransactionNo $Transaction
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
                        $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                        Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF -FileName $PSPDFNameTxt.Text.Trim()
                        CheckPlaycheckSaved -TestPath $userDefinedPlaycheckNameFilePath -TransactionNo $Transaction
                    }
                }
                #neither open in explorer or name PDF
                if ($PSOpenExplorer.Checked -eq $false -and $PSPDFName.Checked -eq $false) {
                    $PSOutputTxtArea.Text += "Now processing transaction $Transaction `r`n"
                    Get-QFPlayCheck -Login $PSUserLoginTxt.Text.Trim() -CasinoID $PSServerTxt.Text.Trim() -TransID $Transaction -SavePDF
                    $autoPlaycheckNameFilePath = Join-Path -Path $PSScriptRoot -ChildPath $autoPlaycheckName
                    CheckPlaycheckSaved -TestPath $autoPlaycheckNameFilePath -TransactionNo $Transaction
                
                }
            }
            $PSOutputTxtArea.Text += "All transactions processed"
        })
    $playcheckControls = $PSServerTxt, $PSUserLoginTxt, $PSUserTransTxt, $PSOpenExplorer, $PSPDFName, $PSPDFNameTxt, $PSGetPlaycheckButton, $PSGetPlaycheckResetTab, $PSUserTransBatchingOptionTxt, $PSUserTransBatchingLbl

    foreach ($control in $playcheckControls) { $playcheck.Controls.Add($control) }

    ################## End of playcheck tab
    ################## Start of player incomplete bet tab
    $incompleteBetTab = New-Object System.Windows.Forms.TabPage
    $incompleteBetTab.Text = "Incomplete bets"

    $PlayerInfoLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 30); ColumnCount = 5; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink }

    $ExportResultLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 70); ColumnCount = 5; RowCount = 1; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink; Name = "ExportResultLayoutPanel" }

    $CommitQueueLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 100); ColumnCount = 3; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink; Name = "CommitQueueLayoutPanel" }

    $RollbackQueueLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel -Property `
    @{Location = New-Object System.Drawing.Point(30, 230); ColumnCount = 3; RowCount = 2; Autosize = $true; AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink; Name = "RollbackQueueLayoutPanel" }

    $ServerTxtBx = New-Object System.Windows.Forms.TextBox -Property @{ Text = "Casino ID"; }
    $UserIDTxtBx = New-Object System.Windows.Forms.TextBox -Property @{Text = "User ID"; }
    $GetQueueButton = New-Object System.Windows.Forms.Button -Property @{Text = "Fetch queues"; AutoSize = $true }
    $ResetTabButton = New-Object System.Windows.Forms.Button -Property @{Text = "Reset tab"; AutoSize = $true }
    $CommitQueueLbl = New-Object System.Windows.Forms.Label -Property @{Text = "Commit queue"; AutoSize = $true }
    $RollbackQueueLbl = New-Object System.Windows.Forms.Label -Property @{Text = "Rollback queue"; AutoSize = $true }
    $QueueStatusLbl = New-Object System.Windows.Forms.Label -Property @{Text = ""; AutoSize = $true }
    $ClearQueueResultsBtn = New-Object System.Windows.Forms.Button -Property @{Text = "Clear results"; Name = "ClearQueueResultsBtn" }
    $IncompleteBetNameExcelChkBx = New-Object System.Windows.Forms.CheckBox -Property @{Text = "Name export?"; Name = "IncompleteBetNameExcelChkBx" }
    $NameExcelTxtBx = New-Object System.Windows.Forms.TextBox -Property @{ Text = "document name"; Width = 150; Name = "NameExcelTxtBx" }
    $ExportQueueResultsBtn = New-Object SYstem.Windows.Forms.Button -Property @{Text = "Export results"; Width = 100; Name = "ExportQueueResultsBtn" }
    $ExportQueueResultsLbl = New-Object System.Windows.Forms.Label -Property @{Text = "" }

    $commitQueueGridView = New-Object System.Windows.Forms.DataGridView -Property `
    @{ScrollBars = [System.Windows.Forms.ScrollBars]::Both; Width = 630; Height = 200; AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells }
    $rollbackQueueGridView = New-Object System.Windows.Forms.DataGridView -Property `
    @{ScrollBars = [System.Windows.Forms.ScrollBars]::Both; Width = 630; Height = 200; AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells }
            
    $orderedKeysRollback = @("rowId", "externalReference", "loginName", "transactionNumber", "betReference", "refundAmount", "gameName", "currency", "dateCreated", "freeGameOffer", "productId", "productName", "userId")
    $orderedKeysCommit = @("rowId", "externalReference", "loginName", "transactionNumber", "winReference", "winAmount", "gameName", "currency", "dateCreated", "progressiveWin", "progressiveDescription", "productId", "productName", "userId", "freeGameOffer")
            

    $PlayerInfoLayoutPanel.Controls.Add($ServerTxtBx, 0, 0)
    $PlayerInfoLayoutPanel.Controls.Add($UserIDTxtBx, 1, 0)
    $PlayerInfoLayoutPanel.Controls.Add($GetQueueButton, 2, 0)
    $PlayerInfoLayoutPanel.Controls.Add($ResetTabButton, 3, 0)
    $PlayerInfoLayoutPanel.Controls.Add($QueueStatusLbl, 4, 0)
    $incompleteBetTab.Controls.Add($PlayerInfoLayoutPanel)

    $PlayerInfoLayoutPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
        $_.Add_MouseClick({ $this.Text = "" })
    }

    $global:commitQueue = @()
    $global:rollbackQueue = @()
    $global:commitQueueExcel = ""
    $global:rollbackQueueExcel = ""

    $ClearQueueResultsBtn.Add_Click({
            $incompleteBetTab.Refresh()
            $QueueStatusLbl.Text = ""
            $controlNames = @("CommitQueueLayoutPanel", "RollbackQueueLayoutPanel", "ExportResultLayoutPanel")

            for ($i = $incompleteBetTab.Controls.Count - 1; $i -ge 0; $i--) {
                $control = $incompleteBetTab.Controls[$i]
                if ($control.Name -in $controlNames) {
                    $incompleteBetTab.Controls.Remove($control)
                }
            }
        })

    $ResetTabButton.Add_Click({
            $ServerTxtBx.Text = "Casino ID"
            $UserIDTxtBx.Text = "User ID"
        })

    $ExportQueueResultsBtn.Add_Click({
            $ExportQueueResultsLbl.Text = "Processing"
            $Commit = 0
            $ExcelName = $NameExcelTxtBx.Text.Trim()
            $UserIDTxtBxTEXT = $UserIDTxtBx.Text
            if ($IncompleteBetNameExcelChkBx.Checked -eq $true) {
                if ($commitQueueExcel.Count -gt 0) {
                    $Commit = 1
                    Export-QFExcel -ExcelData $commitQueueExcel -ExcelTemplate $null -ExcelFileName "$ExcelName.xlsx" -ExcelDestWorksheetName "Commit Queue" -StartRow 1
                }
                if ($rollbackQueueExcel[0]) {
                    if ($Commit -eq 1) {
                        Export-QFExcel -ExcelData $rollbackQueueExcel -ExcelTemplate $null -ExcelSourceWorksheetName "$ExcelName.xlsx" -ExcelDestWorksheetName "Rollback Queue" -StartRow 1
                    }
                    else {
                        Export-QFExcel -ExcelData $rollbackQueueExcel -ExcelTemplate $null -ExcelFileName "$ExcelName.xlsx" -ExcelDestWorksheetName "Rollback Queue" -StartRow 1
                    }
                }
            }
            else {
                if ($commitQueueExcel.Count -gt 0) {
                    $Commit = 1
                    Export-QFExcel -ExcelData $commitQueueExcel -ExcelTemplate $null -ExcelFileName "$UserIDTxtBxTEXT Incomplete Transactions.xlsx" -ExcelDestWorksheetName "Commit Queue" -StartRow 1
                }
                if ($rollbackQueueExcel.Count -gt 0) {
                    if ($Commit -eq 1) {
                        Export-QFExcel -ExcelData $rollbackQueueExcel -ExcelTemplate $null -ExcelSourceWorksheetName "$UserIDTxtBxTEXT Incomplete Transactions.xlsx" -ExcelDestWorksheetName "Rollback Queue" -StartRow 1
                    }
                    else {
                        Export-QFExcel -ExcelData $rollbackQueueExcel -ExcelTemplate $null -ExcelFileName "$UserIDTxtBxTEXT Incomplete Transactions.xlsx" -ExcelDestWorksheetName "Rollback Queue" -StartRow 1
                    }
                }
            }
            $ExportQueueResultsLbl.Text = "Complete!"
            explorer.exe (Get-Location).path 
        })

    $GetQueueButton.Add_Click({
            $commitQueueDataTable = New-Object System.Data.DataTable
            $rollbackQueueDataTable = New-Object System.Data.DataTable
            $QueueStatusLbl.Text = "Processing"
            $global:commitQueue = @()
            $global:rollbackQueue = @()
            $casinoID = $ServerTxtBx.Text.Trim()
            $userID = $UserIDTxtBx.Text.Trim()

            $commitTableHasHeaders = 1
            if ($commitQueueDataTable.Columns.Count -eq 0) {
                $commitTableHasHeaders = 0
            }

            $rollbackTableHasHeaders = 1
            if ($rollbackQueueDataTable.Columns.Count -eq 0) {
                $rollbackTableHasHeaders = 0
            }

            $GetOperatorFromCasino = Invoke-QFPortalRequest -CasinoID $casinoID
            $OP_Key = Get-QFOperatorAPIKeys -OperatorID $GetOperatorFromCasino.operatorID
            $OP_Token = Get-QFOperatorToken -APIKey $OP_key[0].APIKey

            #get player queue data
            $QueueStatusLbl.Text = "Fetching incomplete bets"
            $Queue = Invoke-QFReconAPIRequest -Token $OP_Token.AccessToken -HostingSiteID $GetOperatorFromCasino.hostingSiteID -CasinoID $casinoID -UserID $userID  -QueueInfo
            $global:commitQueueExcel = $Queue.commitQueue

            #create commit queue table headers if they dont already exist
            if ($commitTableHasHeaders -eq 1) { Write-Host "The commit DataTable has headers set." }
            else {
                $commitQueueItem = $Queue.commitQueue[0]
                foreach ($key in $commitQueueItem.PSObject.Properties.Name) {
                    $commitQueueDataTable.Columns.Add($key)
                } 
            }
            
            #if records in commit queue, populate array to store them
            if ($Queue.commitCount -gt 0) {
                foreach ($Item in $Queue.commitQueue) {
                    $global:commitQueue += $Item
                }
                
                #loop through commit queue array, add each record to commit queue data table
                foreach ($Record in $commitQueue) {
                    $CR = @()
                    foreach ($Item in $orderedKeysCommit) {
                        if ($Record.PSObject.Properties.Name -contains $Item) {
                            $itemValue = $Record.$Item
                            $CR += $itemValue
                        }
                    }
                    $commitQueueDataTable.Rows.Add(@($CR))
                    $CR = @()
                }
                
            }
            
            #create rollback queue table headers
            if ($rollbackTableHasHeaders -eq 1) { Write-Host "The rollback DataTable has headers set." }
            else {
                $rollbackQueueItem = $Queue.rollbackQueue[0]
                foreach ($key in $rollbackQueueItem.PSObject.Properties.Name) {
                    $rollbackQueueDataTable.Columns.Add($key)
                }
            }
            

            #if records exist in rollback queue, populate array to store them
            if ($Queue.rollbackCount -gt 0) {
                $global:rollbackQueueExcel = $Queue.rollbackQueue
                $global:rollbackQueue += $Queue.rollbackQueue
            
                foreach ($Record in $rollbackQueue) {
                    $RR = $orderedKeysRollback | ForEach-Object {
                        if ($Record.PSObject.Properties.Name -contains $_) {
                            $Record.$_
                        }
                    }
                    $rollbackQueueDataTable.Rows.Add(@($RR))
                }
            }
            
            $commitQueueGridView.DataSource = $commitQueueDataTable

            if ($Queue.commitCount -gt 0) {
                $CommitQueueLayoutPanel.Controls.Add($commitQueueGridView, 0, 1)
                $CommitQueueLayoutPanel.Controls.Add($CommitQueueLbl, 0, 0)
                $incompleteBetTab.Controls.Add($CommitQueueLayoutPanel)
            }
            else {
                $CommitQueueLbl.Text = "No bets in commit queue"
                $CommitQueueLayoutPanel.Controls.Add($CommitQueueLbl, 0, 0)
                $incompleteBetTab.Controls.Add($CommitQueueLayoutPanel)
            }
            
            $rollbackQueueGridView.DataSource = $rollbackQueueDataTable

            if ($Queue.rollbackCount -gt 0) {
                $RollbackQueueLayoutPanel.Controls.Add($rollbackQueueGridView, 0, 1)
                $RollbackQueueLayoutPanel.Controls.Add($RollbackQueueLbl, 0, 0)
                $incompleteBetTab.Controls.Add($RollbackQueueLayoutPanel)
            }
            else {
                $RollbackQueueLbl.Text = "No bets in rollback queue"
                $RollbackQueueLayoutPanel.Controls.Add($RollbackQueueLbl, 0, 0)
                $incompleteBetTab.Controls.Add($RollbackQueueLayoutPanel)
            }
            
            $QueueStatusLbl.Text = "Complete!"
            if (($Queue.commitCount -gt 0) -or ($Queue.rollbackCount -gt 0)) {
                $ExportResultLayoutPanel.Controls.Add($ClearQueueResultsBtn, 0, 0)
                $ExportResultLayoutPanel.Controls.Add($IncompleteBetNameExcelChkBx, 1, 0)
                $ExportResultLayoutPanel.Controls.Add($NameExcelTxtBx, 2, 0)
                $ExportResultLayoutPanel.Controls.Add($ExportQueueResultsBtn, 3, 0)
                $ExportResultLayoutPanel.Controls.Add($ExportQueueResultsLbl, 4, 0)
                $incompleteBetTab.Controls.Add($ExportResultLayoutPanel)
            }
            $PlayerInfoLayoutPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object {
                $_.Add_MouseClick({ $this.Text = "" })
            }
        })
        
    $tabControl.TabPages.Add($incompleteBetTab)
    ################## End of player incomplete bet tab
    $mainform.Controls.Add($tabControl)
}
$mainform.ShowDialog()