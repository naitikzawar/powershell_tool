###############################################################################
#                                                                             #
#                         Quickfire Powershell Module                         #
#                             Excel Functions                                 #
#                                   v1.6.4                                    #
#                                                                             #
###############################################################################

# Author: Chris Byrne - christopher.byrne@derivco.com.au


# Import system.drawing assembly, required for PowerShell 5 - load before function to prevent error
if ($PSVersionTable.PSVersion.Major -lt 7) {
        [reflection.assembly]::LoadWithPartialName("System.Drawing") | Out-Null
}

function Export-QFExcel {
    <#
    .SYNOPSIS
        Exports data into an Excel spreadsheet.

    .DESCRIPTION
        This cmdlet exports data into an Excel spreadsheet.
        It is tailored for working with QuickFire / Games Global transaction audits by default, but can be used for any Excel spreadsheet.

        This cmdlet requires the Import-Excel third party module. It will attempt to download and install this module automatically.
        The module can be downloaded manually via the command:
        Install-Module -Scope CurrentUser ImportExcel

        See the module website for details: https://www.powershellgallery.com/packages/ImportExcel/

        By default, this cmdlet will copy a worksheet from a source file, into a new or existing Excel file.
        The cmdlet will then populate the new worksheet with data provided in the $ExcelData parameter.
        An 'Audit.xlsx' file is included in the QFPowerShell repository under the 'template' folder.
        This Excel source file is configured with the standard Quickfire / Games Global transaction audit header.
        If the source file cannot be found, a new empty Excel file will be created (or a blank worksheet in an existing Excel target file).

        This cmdlet can also set date or number formatting on specified cell ranges, or change text and background fill colours.

        A number of default parameter values are configured, which can be overwritten by specifying parameters for this function. See the full help text for details.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData
            Exports the object $InputData into an Excel file named "Transaction_Audit.xlsx" in the current working folder.
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx"
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder.
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx" -ExcelTemplate $null
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder.
            Instead of attempting to copy a worksheet named 'Transaction Audit' out of the ExcelTemplate file, a new blank worksheet will be created instead.
            The worksheet in Output.xlsx will be bamed 'Transaction Audit' as per the default value for this parameter.
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx" -ExcelTemplate $null -ExcelDestWorksheetName "Data"
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder.
            Instead of attempting to copy a worksheet named 'Transaction Audit' out of the ExcelTemplate file, a new blank worksheet will be created instead.
            The worksheet in Output.xlsx will be named 'Data'.
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx" -ExcelSourceWorksheetName "Financial Audit"
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder, and a worksheet name "Financial Audit".
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx" -DateFormatRange "B:B" -NumberFormatRange "C:C" -DecimalFormatRange "D:D"
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder.
            Data in Column B will be formatted as Date.
            Data in Column C will be formatted as whole numbers.
            Data in Column D will be formatted as numbers to 2 decimal places.
            Other default parameter values will be used.

    .EXAMPLE
        Export-QFExcel -ExcelData $InputData -ExcelFileName "Output.xlsx" -ColourRange "A1:Z1" -ColourText "Blue" -ColourFill "Green"
            Exports the object $InputData into an Excel file named "Output.xlsx" in the current working folder.
            Cells in the range A1:Z1 will have a colour style applied.
            Text colour will be set to Blue, and Green solid fill will be applied to these cells.
            Other default parameter values will be used.

    .PARAMETER ExcelData
        The object containing data to be exported into an Excel spreadsheet.

    .PARAMETER ExcelTemplate
        The name and path to an existing Excel file. This cmdlet will look for a work sheet with a name matching the 'ExcelSourceWorksheetName' parameter.
        This worksheet will then be copied into the target Excel file matching the 'ExcelFileName' parameter.

        By default, this parameter will be set to look in the QFPowerShell repository folder for a folder name 'Template', and then look for a file name 'Audit.xlsx'.
        This 'Audit.xlsx' file is included in the QFPowerShell repository and is configured with the standard Quickfire / Games Global transaction audit header.
        It includes two worksheets, one for Transaction Audits and one for Financial Audits (only the text box in the heading is different between the two).

        If the template file does not exist, or this parameter is set to $Null, a new empty Excel file will be created instead.

    .PARAMETER ExcelFileName
        The name and path to the output Excel file. By default this is a file named "Transaction_Audit.xlsx" in the current working folder.
        If the file does not exist it will be created. If the file already exists, a new worksheet will be added and populated with the data in the 'ExcelData' parameter.
        Any worksheets that already exist with a name matching the 'ExcelDestWorksheetName' parameter will be overwritten.

        An existing Excel file must not be open in Excel while this cmdlet runs or an error will occur.

    .PARAMETER ExcelSourceWorksheetName
        The name of the worksheet to copy out of the 'ExcelTemplate' file and into the target 'ExcelFileName' file. By default this is set to 'Transaction Audit'.
        If the ExcelTemplate file does not exist, or the ExcelTemplate parameter is set to $Null, this parameter has no effect.

    .PARAMETER ExcelDestWorksheetName
        The name of the worksheet to create in the target 'ExcelFileName' file. By default this is set to match the 'ExcelSourceWorksheetName' parameter.
        Any worksheets that already exist in the target file with a name matching this parameter will be overwritten.

    .PARAMETER StartRow
        The row number to start populating data in the target worksheet.
        By default this is set to 7, this leaves enough room for the standard Quickfire / Games Global transaction audit header.

    .PARAMETER DateFormatRange
        A range of cells to format as a Date inside the target worksheet. The range should be enclosed in quotation marks (" ")
        For example, setting this parameter to "A:A" would set all cells in column A to a Date format.
        This parameter accepts multiple objects as input, so you can specify multiple ranges seperated by commas.
        For example, "A:A","B1:B20","F:G"

        The specified cells will be set to a format of 'yyyy/MM/dd hh:mm:ss.000'

    .PARAMETER NumberFormatRange
        A range of cells to format as Whole Numbers inside the target worksheet. The range should be enclosed in quotation marks (" ")
        For example, setting this parameter to "A:A" would set all cells in column A to a Date format.
        This parameter accepts multiple objects as input, so you can specify multiple ranges seperated by commas.
        For example, "A:A","B1:B20","F:G"

        The specified cells will be set to a format of '#' - only whole numbers will be shown with no decimal places and no thousands seperators.
        e.g. a value of 1,000.111 will be displayed as 1000

    .PARAMETER DecimalFormatRange
        A range of cells to format as numbers to 2 decimal places inside the target worksheet. The range should be enclosed in quotation marks (" ")
        For example, setting this parameter to "A:A" would set all cells in column A to a Date format.
        This parameter accepts multiple objects as input, so you can specify multiple ranges seperated by commas.
        For example, "A:A","B1:B20","F:G"

        The specified cells will be set to a format of '0.00' - numbers will be shown with 2 decimal places and no thousands seperators.
        e.g. a value of 1,000.111 will be displayed as 1000.11

    .PARAMETER ColourRange
        A range of cells to apply a text colour and solid background fill to.
        For example, setting this parameter to "A:A" would apply the colour to all cells in column A.
        This parameter accepts multiple objects as input, so you can specify multiple ranges seperated by commas.
        For example, "A:A","B1:B20","F:G"

        By default, this will set white text on a dark blue background, to match the standard Quickfire / Games Global transaction audit header.
        These colours can be adjusted with 'ColourText' and 'ColourFill' parameters.

        This parameter has no default value, as the number of columns in the table will vary depending on the data exported into excel.
        e.g. If you have 7 columns you should set a range of "A7:G7" for this parameter (assuming the default StartRow value).
        You could apply the colour style to an entire row, e.g. "7:7" but this makes the horizontal scroll bar very small.

    .PARAMETER ColourText
        The colour of the text that will be applied to the cell range specified by the 'ColourRange' parameter.
        This parameter accepts any object that is of the type '[System.Drawing.Color]'
        You can specify the name of a basic colour, or an RGB value using '[System.Drawing.Color]::FromArgb(R,G,B)'
        By default this is set to 'White'.

        If 'ColourRange' parameter is not specified, this parameter has no effect.

    .PARAMETER ColourFill
        The colour of the solid background fill that will be applied to the cell range specified by the 'ColourRange' parameter.
        This parameter accepts any object that is of the type '[System.Drawing.Color]'
        You can specify the name of a basic colour, or an RGB value using '[System.Drawing.Color]::FromArgb(R,G,B)'
        By default this is set to a dark blue colour to match the standard Quickfire / Games Global transaction audit header.
        The default colour value is '[System.Drawing.Color]::FromArgb(10,31,64)'

        If 'ColourRange' parameter is not specified, this parameter has no effect.

    .INPUTS
        This cmdlet does not accept pipeline input. Please specify the data to export to Excel using the -ExcelData parameter.

    .OUTPUTS
        This cmdlet does not produce any pipeline output.

    .NOTES
        Author:     Chris Byrne
        Email:      christopher.byrne@derivco.com.au

    .LINK
        https://www.powershellgallery.com/packages/ImportExcel/

    #>

    # Set up parameters for this function
    [CmdletBinding()]
    param (
        # The object holding the data we want to export into excel
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        $ExcelData,

        # The Excel template file name where we look for the source worksheet. Default looks for Audit.xlsx in the Template folder included with the QFPowershell repo.
        [Parameter()]
        [string]$ExcelTemplate = $($PSScriptRoot -replace "\\src$","\template\Audit.xlsx"),

        # The Excel output file name, will be created if it doesn't exist
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$ExcelFileName = "Transaction_Audit.xlsx",

        # The Excel worksheet name in $ExcelTemplate that gets copied into $ExcelFilename, renamed to $ExcelDestWorksheetName and populated with data
        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$ExcelSourceWorksheetName = "Transaction Audit",

        # The Excel worksheet name to create and export data into, copied from file $ExcelTemplate and worksheet $ExcelSourceWorksheetName. Default is same name as $ExcelSourceWorksheetName
        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string]$ExcelDestWorksheetName = $ExcelSourceWorksheetName,

        # The row to start populating data in the worksheet. Default is set to 7, to allow for the standard GGL transaction audit header.
        [Parameter(Mandatory = $false)]
        [int]$StartRow = 7,

        # Excel row/column range to format as Date (yyyy/MM/dd hh:mm:ss.000), multiple ranges can be done by passing an array of ranges.
        [Parameter(Mandatory = $false)]
        [string[]]$DateFormatRange,

        # Excel row/column range to format as Whole Numbers, multiple ranges can be done by passing an array of ranges.
        [Parameter(Mandatory = $false)]
        [string[]]$NumberFormatRange,

        # Excel row/column range to format as Decimal Numbers to two decimal places, multiple ranges can be done by passing an array of ranges.
        [Parameter(Mandatory = $false)]
        [string[]]$DecimalFormatRange,

        # Excel row/column range to apply colour settings. Multiple ranges can be done by passing an array of ranges. Default color is white text with dark blue fill
        [Parameter(Mandatory = $false)]
        [string[]]$ColourRange,

        # Font colour applied to $ColourRange. This parameter has no effect if $ColourRange is not set.
        # Default setting shows how you can specify a colour name
        [Parameter(Mandatory = $false)]
        [System.Drawing.Color]$ColourText = "White",

        # Background fill colour applied to $ColourRange. This parameter has no effect if $ColourRange is not set.
        # Default setting shows how you can specify an RGB colour value (10,31,64 is a dark blue to match our standard GGL audit template)
        [Parameter(Mandatory = $false)]
        [System.Drawing.Color]$ColourFill = $([System.Drawing.Color]::FromArgb(10,31,64))

    )
    # Check if Import-Excel module is installed, attempt to download and install if not
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    If ((Get-Module ImportExcel).Count -lt 1) {
        try {
            Install-Module -Scope CurrentUser ImportExcel -Force -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to import or install the ImportExcel module. Please install this module before running this cmdlet again. Visit https://www.powershellgallery.com/packages/ImportExcel/ for more info."
            Return
        }
    }

    # If ExcelTemplate is null or the Excel Template file can't be found, don't try to copy the template worksheet into the destination file, we'll just create a new empty file instead
    If ([string]$ExcelTemplate -ne "") {
        If (Test-Path -Path $ExcelTemplate -PathType Leaf) { # seperate if statements to avoid an error from test-path if $ExcelTemplate is null
            # copy spreadsheets from the template into our target excel file. Will create a new file if it doesn't exist
            Copy-ExcelWorksheet $ExcelTemplate -SourceWorksheet $ExcelSourceWorksheetName -DestinationWorkbook $ExcelFileName -DestinationWorksheet $ExcelDestWorksheetName -ErrorAction Stop
        }
    }

    # Now export the data into our new Excel worksheet and save it
    Try {
        If (Test-Path -Path $ExcelFileName -PathType Leaf -ErrorAction SilentlyContinue) {
            # Check if the target file already exists e.g. an existing file, or we just created it from the template
            # Open-ExcelPackage creates an Excel Package object for use with the various functions of the ImportExcel module
            $ExcelObject = Open-ExcelPackage -path $ExcelFileName
            $ExcelData | Export-Excel -ExcelPackage $ExcelObject -WorksheetName $ExcelDestWorksheetName -StartRow $StartRow -Tablename $(($ExcelDestWorksheetName.trim() -replace " ","") + "Table") -AutoSize -AutoFilter -MaxAutoSizeRows 0 -BoldTopRow -PassThru | Out-Null
        } Else {
            # Target file doesn't exist so create a new empty excel file named $ExcelFileName and an Excel Object package pointing to it
            $ExcelObject = $ExcelData | Export-Excel -Path $ExcelFileName -WorksheetName $ExcelDestWorksheetName -StartRow $StartRow -Tablename $(($ExcelDestWorksheetName.trim() -replace " ","") + "Table") -AutoSize -AutoFilter -MaxAutoSizeRows 0 -BoldTopRow -PassThru
        }
        # finally save the newly created excel package object to disk
        #$ExcelObject.Save()
    } Catch {
        Write-Error "An error occured attempting to save data into the Excel file $ExcelFileName and worksheet $ExcelDestWorksheetName"
        Throw $_.Exception.Message
    }

    # Apply the date formatting for each specified range
    If ($null -ne $DateFormatRange) {
        Try {
            foreach ($Range in $DateFormatRange) {
                Set-ExcelRange -Worksheet $ExcelObject.Workbook.Worksheets[$ExcelDestWorksheetName] -Range $Range -NumberFormat 'yyyy/MM/dd hh:mm:ss.000' -AutoFit
            }
        } Catch {
            Write-Error "An error occured attempting to format cell range $DateFormatRange as a Date type in the Excel file $ExcelFileName and worksheet $ExcelDestWorksheetName"
            Throw $_.Exception.Message
        }
    }

    # Apply the whole number formatting for each specified range
    If ($null -ne $NumberFormatRange) {
        Try {
            foreach ($Range in $NumberFormatRange) {
                Set-ExcelRange -Worksheet $ExcelObject.Workbook.Worksheets[$ExcelDestWorksheetName] -Range $Range -NumberFormat '0' -AutoFit
            }
        } Catch {
            Write-Error "An error occured attempting to format cell range $NumberFormatRange as a Whole Number type in the Excel file $ExcelFileName and worksheet $ExcelDestWorksheetName"
            Throw $_.Exception.Message
        }
    }

    # Apply the decimal number formatting for each specified range
    If ($null -ne $DecimalFormatRange) {
        Try {
            foreach ($Range in $DecimalFormatRange) {
                Set-ExcelRange -Worksheet $ExcelObject.Workbook.Worksheets[$ExcelDestWorksheetName] -Range $Range -NumberFormat '0.00' -AutoFit
            }
        } Catch {
            Write-Error "An error occured attempting to format cell range $DecimalFormatRange as a Decimal Number type in the Excel file $ExcelFileName and worksheet $ExcelDestWorksheetName"
            Throw $_.Exception.Message
        }
    }

    # Apply the header colour formatting for each specified range
    If ($null -ne $ColourRange) {
        Try {
            foreach ($Range in $ColourRange) {
                Set-ExcelRange -WorkSheet $ExcelObject.Workbook.Worksheets[$ExcelDestWorksheetName] -Range $Range -FontColor $ColourText -BackgroundColor $ColourFill -BackgroundPattern Solid -BorderColor $ColourFill -BorderAround Thin
            }
        } Catch {
            Write-Error "An error occured attempting to apply styles to cell range $ColourRange in the Excel file $ExcelFileName and worksheet $ExcelDestWorksheetName"
            Throw $_.Exception.Message
        }
    }
    # all done, so close off the Excel Package. Try 3 times
    $i = 1
    do {
        Write-Verbose ("[$(Get-Date)] Saving Excel file $ExcelFileName, attempt $i...")
        try {
            Close-ExcelPackage $ExcelObject
            $i = 4
        } catch {
            Start-Sleep $i
            $i += 1
            if ($i -gt 3) {Throw $_.Exception.Message}
        }
    } until ($i -gt 3)

}