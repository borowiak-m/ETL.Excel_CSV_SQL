#  LOGIC EXPLAINED
# ---------------
# This script aims to export data in an Excel file to csv files, one per worksheet. We achieve this following this logic:
#
# - Read process params from a settings text file, like folder locations and file names.
#
# - Check when was the file last written to, if there are new changes. Then compare last modified date and time to a stored
#   value of date and time of the last data extract. Should the modification time be more recent, this triggers a new export.
#
# - Load the excel file and export all data from a given sheet name. This will be exported into a csv file, but first we need 
#   to check if the previously exported csv files are still in the export folder. These files is to be picked up by another process.
#   If the previous export is still sat in the folder, we move it to an Error folder and replace it with a new export instead and log the error.
#
# To be continued:
# ----------------
# - CSV inmport into SQL tables script to pick up the exports
# - Potential housekeeping of old files in Error folder
# - potential clashes with excel sheet being edited at the same time or if open

# Import the ImportExcel module
Import-Module -Name ImportExcel

function Write-Error($errorFolderPath, $errorMsg, $errorLvl) {

    # Log error in cmd
    Write-Host $errorMsg                            

    # Generate a timestamp for error file name
    $timestamp = Get-Date -format "yyyy.MM.dd hh.mm.ss"

    # Generate a date for error file name
    $errorDate = Get-Date -format "yyyyMMdd"

    # Generate a unique error file path
    $errorLogFilePath = Join-Path -Path $errorFolderPath -ChildPath("$errorDate FileImportError.txt")

    # Check if file already exists, if so append the error message to existing file, if not create a new error file
    If (Test-Path $errorLogFilePath) {
        Add-Content $errorLogFilePath "$timestamp $errorMsg"
    } else {
        Set-Content $errorLogFilePath "$timestamp $errorMsg"
    }

    If ($errorLvl -eq "Fatal") {
        Write-Host "Debug: Fatal error, exiting program."
        Exit
    } else {
        Write-Host "Debug: Error of level $errorLvl. Recommencing program."
    }

}

# Initialize settings from file, where all business folder paths are stored
$settingFileName        = "ExcelImport_settings.txt"
$settingsFolderPath     = "D:\Scripts\"
$settingsFilePath       = Join-Path -Path $settingsFolderPath -ChildPath ($settingFileName)

# Initialize default folder locations
$errorFolderPath        = "D:\Scripts\ExcelImport_Error\"
If (-Not (Test-Path $errorFolderPath)) {New-Item -Path $errorFolderPath -ItemType Directory}

# Check for existence of the settings file
If (-Not(Test-Path $settingsFilePath)) {Throw "No settings document found"}

$settings = @{}

# Get variablename=filepath strings per each line, split by "=" and store in $settings dict
Get-Content $settingsFilePath | ForEach-Object {
    $paramValuePair = $_ -split "="
    $settings[$paramValuePair[0]] = $paramValuePair[1]
}

# Read settings into variables

$excelFilePath           = $settings['excelFilePath']
$lastTimeFilePath        = $settings['lastTimeFilePath']
$csvExportFolderPath     = $settings['csvExportFolderPath']
$sheetsToExport          = $settings['sheetsToExport'] -split "," | ForEach-Object trim($it)

$paramsToCheck = @($excelFilePath, $lastTimeFilePath, $csvExportFolderPath, $sheetsToExport)

# Check for empty params 
ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { Write-Error $errorFolderPath "Params missing. Review settings file under $settingsFilePath" Fatal} }

# Check for the existence of excel document
If (-Not(Test-Path $excelFilePath)) {Write-Error $errorFolderPath "No document found under $excelFilePath" Fatal}

# Check for existence of csv export folder
If (-Not (Test-Path $csvExportFolderPath)) {New-Item -Path $csvExportFolderPath -ItemType Directory}

# Fetch last edit time of the excel document
$lastModifiedTime = (Get-Item $excelFilePath).LastWriteTime

If (Test-Path $lastTimeFilePath) {

    # If file exists get the last modified datetime
    $lastKnownTime = New-Object DateTime (Get-Content $lastTimeFilePath)
    Write-Host "Debug: lastKnownTime = $lastKnownTime"

} else {

    # If the file file does not exist for some reason assume the beginning of time
    # as the last known modification datetime
    $lastKnownTime = [DateTime]::MinValue
    Write-Host "Debug: Last modification time not found in file. lastKnownTime set to $lastKnownTime."
}

if ($lastModifiedTime -gt $lastKnownTime) {

    # File has been modified since last check
    Write-Host "Debug: File modified since $lastKnownTime."

    # Get all workshet names from the Excel document to be extracted
    $sheets = Get-ExcelSheetInfo $excelFilePath | Select-Object -ExpandProperty Name

    $matchCounter = 0

    ForEach ($sheet in $sheets) {

        ForEach ($sheetToExport in $sheetsToExport) {

            Write-Host ("Debug: Processing worksheet: " + $sheet)
            Write-Host ("Debug: Matching with sheet name: " + $sheetToExport)
            
            # If worksheet name matches a sheet name to be extracted
            If ($sheet -eq $sheetToExport) {
                
                Write-Host "Debug:  - - - > $sheet and $sheetToExport [MATCH]"
                $matchCounter = $matchCounter + 1

                # Build file path for csv file export with sheet name
                $csvExportFilePath = Join-Path -Path $csvExportFolderPath -ChildPath ("$sheet.csv")

                # Read from excel file
                $allData = Import-Excel -Path $excelFilePath -WorksheetName $sheet

                # Check if an export csv file already exists (if so move it to Error folder and replace it)
                If (Test-Path $csvExportFilePath) {

                    # Log unprocessed file
                    Write-Error $errorFolderPath "Unprocessed file $csvExportFilePath. File renamed and moved to $errorFolderPath." NotFatal

                    # Generate a timestamp for error file name
                    $timestamp = Get-Date -format "yyyy.MM.dd hh.mm.ss"

                    # Move existing file to error folder and rename it with a timestamp
                    $csvErrorFilePath = Join-Path -Path $errorFolderPath -ChildPath ("Unprocessed $sheet $timestamp.csv")
                    Move-Item -Path $csvExportFilePath -Destination $csvErrorFilePath

                } 

                #Export to csv
                $allData | Export-Csv -Path $csvExportFilePath -NoTypeInformation -Encoding UTF8


            } else {

                Write-Host "Debug: $sheet and $sheetToExport [NO MATCH]"

            }
        }

    }

    Write-Host "Debug: Total matches : $matchCounter"

    # After all tabs are processed, update the last modification date in the text file for next run
    Set-Content $lastTimeFilePath $lastModifiedTime.Ticks

} else {

    Write-Host "Debug: No new changes were written to the file"

}


