###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                             Bet Setting Proc Builder                        #
#                                   v1.6.3                                    #
#                                                                             #
###############################################################################

#Author: Harley Osgood - harley.osgood@derivco.com.au



Function Get-QFBetSettingProc {
    <#
    .SYNOPSIS
        Create bet setting procs for 1 to 1 and currency multiplier variations.

    .DESCRIPTION
        This cmdlet is used to create bet settings procs that CasinoPortal will not generate. Casino Portal and Sherlock
        do not contain the logic for, or allow settings to be changed for table games, and video bingo games.
        This cmdlet does not contain the pre-existing logic that sherlock has for functional bet value options.
        It will therefore not find the closest viable bet settings options to the users input. Instead this needs to be tested calculated before this cmdlet is used.
        The cmdlet will take the players input OperatorID, MID-CID combos, Currencies, SettingsIDs and values and pre-fill all required SQL procs to be run on the DB

    .EXAMPLE
        Entering the information as prompted for OperatorID 41662. For Currencies USD, GBP, MXN, ARS. For the game variants 10876,50300 and 10876,4033. Entering Setting 105 for max bet. Adding Value of 800 ($8.00) will generate the following procs;

            EXEC pr_UpdateDefaultModuleSetting 10876,50300, 105, 800, Null, 41662, Null
            EXEC pr_UpdateDefaultModuleSetting 10876,40300, 105, 800, Null, 41662, Null
            EXEC pr_UpdateCurrencyModuleSetting 45, 10876,50300, 105, 16000, Null, 41662, Null
            EXEC pr_UpdateCurrencyModuleSetting 45, 10876,40300, 105, 16000, Null, 41662, Null
            EXEC pr_UpdateCurrencyModuleSetting 51, 10876,50300, 105, 40000, Null, 41662, Null
            EXEC pr_UpdateCurrencyModuleSetting 51, 10876,40300, 105, 40000, Null, 41662, Null

        There is only one UpdateDefaultModuleSetting proc for each MID/CID combo as both USD and GBP are 1:1 Currencies.
        There is an UpdateCurrencyModuleSettings for each MID-CID combo for each CurrencyID provided on non 1:1 Currencies.
        Entering an unknown Currency such as 'abc' will trigger an error for that currency commented out in SQL formatting as follows;
            --Incorrect currency provided abc

    .PARAMETER OpID
        The OperatorID relevant to the casino/operator changes are being made to.

    .PARAMETER Currency
        The Currency ISO in a 3 character format. This is not case sensitive as the cmdlet will adjust as needed.
        Uknown currency ISO's will trigger an error for that entry while continuing to output the rest.

    .PARAMETER MIDCID
        The MIDCID is the ModuleID of the game and ClientID in the format of *****,#####.
        This must be entered for every variant of the game being changed. 
        The format follows mid comma cid to ensure both parameters are filled into the proc correctly.

    .PARAMETER SettingID
        The bet setting ID relevant to the setting being changed on the DB.
        

    .PARAMETER Value
        Value of the setting being changed. As per the database this is in cents value.
        For all settings this will be multiplied by the currency multiplier associated with the currency ISO.
        This excludes settings 215 and 287 which do not increase with currency multipliers.

    

    .NOTES
        Author:     Harley Osgood
        Email:      harley.osgood@derivco.com.au

    #>
    [alias ("BSP")]
    [CmdletBinding()] param(
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [int]$OpID,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Currency,

        [Parameter(Mandatory = $true, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string[]]$MIDCID,

        [Parameter(Mandatory = $true, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [string]$SettingID,
 
        [Parameter(Mandatory = $true, Position = 5)]
        [ValidateNotNullOrEmpty()]
        [int]$Value


    )

  
    $processedMIDCIDs = @()  # To keep track of processed MIDCIDs for -1 CurrencyID
    $CurrencyList = Invoke-QFPortalRequest -Currency -ErrorAction Continue
    Foreach ($CurrencyEntry IN $currency) {
        $currencyid = $null
    
     # Check for valid currency and display info
    $CurrencyID = $Currencyentry.ToUpper().Trim()
    $CurrencyInfo = $CurrencyList | Where-Object {$_.ISOCode -eq $CurrencyID}
    If ($null -eq $CurrencyInfo) {
        Write-Warning "Unable to validate the specified currency. $CurrencyID"
        continue
    }
  
        write-debug "currencyid $currencyid"
        write-debug "currencyentry $currencyentry"

       
        if ($currencyinfo.multipliermaxbet -eq 1) {
            
            Foreach ($MIDCIDentry IN $MIDCID) {
              
                if ($processedMIDCIDs -notcontains $MIDCIDentry) {
                    "--MID,CID $MIDCIDentry, Currency $($CurrencyInfo.isocode)"
                    "EXEC pr_UpdateDefaultModuleSetting $MIDCIDentry, $SettingID, $Value, Null, $OpID, Null"
                
                    $processedMIDCIDs += $MIDCIDentry
                }
            }
        }
        else {
            Foreach ($MIDCIDentry IN $MIDCID) {
                "--MID,CID $MIDCIDentry, Currency $($CurrencyInfo.isocode)"
                if ($SettingID -eq 215 -or $SettingID -eq 287) {
             
                   "EXEC pr_UpdateCurrencyModuleSetting $($Currencyinfo.currencyid), $MIDCIDentry, $SettingID, $Value, Null, $OpID, Null"
                   }

                else {
                "EXEC pr_UpdateCurrencyModuleSetting $($Currencyinfo.currencyid), $MIDCIDentry, $SettingID, $($Value * [int]$currencyinfo.multipliermaxbet), Null, $OpID, Null"
                }
               
            }
        }
    }
}