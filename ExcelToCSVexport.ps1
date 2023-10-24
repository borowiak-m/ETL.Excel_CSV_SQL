#  LOGIC EXPLAINED
# ---------------
# This script aims to export data in an Excel file to csv files, one per worksheet. We achieve this following this logic:
#
# - Read process params from a settings text file, like folder locations and file names.
#
# - Check when was the file last written to, if there are new changes. Then compare last modified date and time to a stored
#   value of date and time of the last data export. Should the modification time be more recent, this triggers a new export.
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
    $timestamp = Get-Date -format "yyyy.MM.dd HH.mm"

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

# Export settings files are all post-fixed with "_export_settings". They are divided into two groups:
#   - one is prefixed with "main": this should be a single file cantaining settings for the export process 
#     that apply to all exported data (example filename main_export_settings.txt)
#   - another one is prefixed with "filename" which is going to correspond with an Exel filename for exportion process: there will be as many files of this type
#     as many files there are to process, each containing specific params for the file, like it's path, and sheets to export (file name: Stockboard_export_settings.txt)

# Initialize default error folder locations and file names
$processingSettingsFolderPath   = "D:\Scripts\Stock Blackboard\Settings\"
$processingSettingsFiles        = Get-ChildItem -Path $processingSettingsFolderPath -Filter *_export_settings.txt
$errorFolderPath                = "D:\Scripts\Stock Blackboard\Error\"
$lastModLogFileName             = "last_time_modified.txt"
$mainExportSettingsFileName     = "export_settings.txt"
$mainExportSettingsFilePath     = Join-Path -Path $processingSettingsFolderPath -ChildPath ($mainexportSettingsFileName)

# Check for existence of the main export settings file
If (-Not(Test-Path $mainexportSettingsFilePath)) {Throw "No main export settings document found."}

# Check for existence of error folder
If (-Not (Test-Path $errorFolderPath)) {New-Item -Path $errorFolderPath -ItemType Directory}

# Get params from main export settings file
$exportSettings = @{}

# Get variablename=filepath strings per each line, split by "=" and store in $settings dict
Get-Content $mainexportSettingsFilePath | ForEach-Object {
    $key, $val      = $_ -split "="
    $exportSettings[$key] = $val
}

# Initialize settings from file, where all business folder paths are stored
$lastModLogFolderPath    = $exportSettings['lastModLogFolderPath']
$csvExportFolderPath     = $exportSettings['csvExportFolderPath']

$paramsToCheck = @($lastModLogFolderPath, $csvExportFolderPath)

# Check for empty params 
ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { Write-Error $errorFolderPath "Params missing. Review settings file under $mainExportSettingsFilePath" Fatal} }

# Check for existence of time modified log folder and csv export folder
If (-Not (Test-Path $lastModLogFolderPath)) {New-Item -Path $lastModLogFolderPath -ItemType Directory}
If (-Not (Test-Path $csvExportFolderPath)) {New-Item -Path $csvExportFolderPath -ItemType Directory}


# Process export files from export file settings
ForEach ($settingsFile in $processingSettingsFiles) {

    # Read the settings file and fetch params
    $exportSettings = @{}

    Get-Content $settingsFile.FullName | ForEach-Object {
        $key, $val      = $_ -split "="
        $exportSettings[$key] = $val
    }

    # Initialize settings from file, where all business folder paths are stored
    $exportFileExtention     = $exportSettings['exportFileExtention']
    $exportSourceFolderPath  = $exportSettings['exportSourceFolderPath']
    $sheetsToExport          = $exportSettings['sheetsToExport'] -split "," | ForEach-Object trim($it)

    $exportFileBaseName      = ($settingsFile.BaseName -replace "_export_settings", "")
    $exportFileName          = ($settingsFile.BaseName -replace "_export_settings", "") + $exportFileExtention
    $exportFilePath          = Join-Path -Path $exportSourceFolderPath -ChildPath ($exportFileName)

    Write-Host "Debug: processing file : $settingsFile"
    Write-Host "Debug: processing param exportFileExtention : $exportFileExtention" 
    Write-Host "Debug: processing param exportSourceFolderPath : $exportSourceFolderPath"
    Write-Host "Debug: processing param sheetsToExport : $sheetsToExport"
    Write-Host "Debug: processing param exportFileBaseName : $exportFileBaseName"
    Write-Host "Debug: processing param exportFileName  : $exportFileName"
    Write-Host "Debug: processing param exportFilePath : $exportFilePath "

    # Check for empty params 
    $hasEmptyParams = $false
    $paramsToCheck           = @($exportFileExtention, $exportSourceFolderPath, $sheetsToExport)
    ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { $hasEmptyParams = $true } }

    If ($hasEmptyParams){ 
        Write-Error $errorFolderPath "Params missing. File $exportFileBaseName is skipped from export process. Review settings file under $settingsFile" NotFatal
        Continue
    }

    # Check for the existence of export document
    If (-Not(Test-Path $exportFilePath)) {  
        Write-Error $errorFolderPath "File missing. File $exportFilePath is skipped from export process. Review settings file under $settingsFile" NotFatal
        Continue
    }

    # Check for existence of csv export folder
    If (-Not (Test-Path $csvExportFolderPath)) {New-Item -Path $csvExportFolderPath -ItemType Directory}

    # Build last modified log file path
    $lastModLogFilePath = Join-Path -Path $lastModLogFolderPath -ChildPath ($exportFileBaseName + "_" + $lastModLogFileName)

    # Fetch last edit time of the excel document
    $lastModifiedTime = (Get-Item $exportFilePath).LastWriteTime

    If (Test-Path $lastModLogFilePath) {

        # If file exists get the last modified datetime
        $lastKnownTime = New-Object DateTime (Get-Content $lastModLogFilePath)
        Write-Host "Debug: lastKnownTime = $lastKnownTime"

    } else {

        # If the file file does not exist for some reason assume the beginning of time
        # as the last known modification datetime
        $lastKnownTime = [DateTime]::MinValue
        Write-Host "Debug: Last modification time not found in file. lastKnownTime set to $lastKnownTime."
    }

    If ($lastModifiedTime -gt $lastKnownTime) {

        # File has been modified since last check
        Write-Host "Debug: File modified since $lastKnownTime."

        # Get all workshet names from the Excel document to be exported
        $sheets = Get-ExcelSheetInfo $exportFilePath | Select-Object -ExpandProperty Name

        $matchCounter = 0

        ForEach ($sheet in $sheets) {

            ForEach ($sheetToExport in $sheetsToExport) {

                Write-Host ("Debug: Processing worksheet: " + $sheet)
                Write-Host ("Debug: Matching with sheet name: " + $sheetToExport)
                
                # If worksheet name matches a sheet name to be exported
                If ($sheet -eq $sheetToExport) {
                    
                    Write-Host "Debug:  - - - > $sheet and $sheetToExport [MATCH]"
                    $matchCounter = $matchCounter + 1

                    # Build file path for csv file export with sheet name
                    $csvExportFilePath = Join-Path -Path $csvExportFolderPath -ChildPath ("$sheet.csv")

                    # Read from excel file
                    $allData = Import-Excel -Path $exportFilePath -WorksheetName $sheet

                    # Check if an export csv file already exists (if so move it to Error folder and replace it)
                    If (Test-Path $csvExportFilePath) {

                        # Log unprocessed file
                        Write-Error $errorFolderPath "Unprocessed file $csvExportFilePath. File renamed and moved to $errorFolderPath." NotFatal

                        # Generate a timestamp for error file name
                        $timestamp = Get-Date -format "yyyy.MM.dd HH.mm"

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
        If ($matchCounter -gt 0) { Set-Content $lastModLogFilePath $lastModifiedTime.Ticks }

    } else {

        Write-Host "Debug: No new changes were written to the file"

    }
    

}