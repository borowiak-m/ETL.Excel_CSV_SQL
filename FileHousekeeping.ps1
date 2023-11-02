#  LOGIC EXPLAINED
# ---------------
# This script aims to clear any old files from the monitored folder paths. We achieve this following this logic:
#
# - Read process params (folder paths) from a settings text file
# - Loop through the monitored folders and check file age
# - If file is older than X days, delete the file
# - Folders to monitor for old files: 
#     - $errorFolderPath - path where erros logs are kept and files not processed
#     - $importErrorFolderPath - error path specified in the import settings file, could be different to default error folder
#     - $exportErrorFolderPath - error path specified in the export settings file, could be different to default error folder
#     - $csvExportFolderPath - path where csv sheets are exported from business docs, and imported from to db 
#     - $importProcessedFolderPath - path where processed import csv files are moved

# Initialize default logging and settings folders

$processingSettingsFolderPath   = "D:\Scripts\Stock Blackboard\Settings\"
$processingSettingsFiles        = Get-ChildItem -Path $processingSettingsFolderPath -Filter *settings.txt
$errorFolderPath                = "D:\Scripts\Stock Blackboard\Error\"
$importErrorFolderPath = ""
$exportErrorFolderPath = ""
$csvExportFolderPath = ""
$importProcessedFolderPath = ""



# Check if settings files exist
If (-Not (Test-Path $processingSettingsFolderPath)) { Throw "No settings folder found" } else {"File was found"}

# Check that we found any settings files in the folder
If (-Not ($processingSettingsFiles.Count -gt 0))    { Throw "No settings files were found" } else {"File was found"}

# $errorFolderPath - path where erros logs are kept and files not processed
# $csvExportFolderPath - path where csv sheets are exported from business docs, and imported from to db 
# $importProcessedFolderPath - path where processed import csv files are moved
# $importErrorFolderPath - error path specified in the import settings file, could be different to default error folder
# $exportErrorFolderPath - error path specified in the export settings file, could be different to default error folder

ForEach ($sFile in $processingSettingsFiles) {

    # Get import settings params
    If ($sFile.BaseName -eq 'import_settings') {
        "- - - Found import_settings: $($sFile.FullName)"
        # Get params from main extract settings file
        $importSettings = @{}

        # Get variablename=filepath strings per each line, split by "=" and store in $settings dict
        Get-Content $sFile.FullName | ForEach-Object {
            $key, $val      = $_ -split "="
            $importSettings[$key] = $val
        }

        # Initialize settings from file, where all business folder paths are stored
        $importProcessedFolderPath          = $importSettings['importProcessedFolderPath']
        $importErrorFolderPath              = $importSettings['errorFolderPath']
    }

    # Get export settings file params
    If ($sFile.BaseName -eq 'export_settings') {
        "- - - Found export_settings: $($sFile.FullName)"

        # Get params from main export settings file
        $exportSettings = @{}

        # Get variablename=filepath strings per each line, split by "=" and store in $settings dict
        Get-Content $sFile.FullName | ForEach-Object {
            $key, $val      = $_ -split "="
            $exportSettings[$key] = $val
        }

        # Initialize settings from file, where all business folder paths are stored
        $csvExportFolderPath     = $exportSettings['csvExportFolderPath']
        $exportErrorFolderPath   = $exportSettings['errorFolderPath']

    }

}

$foldersToMonitor =@($errorFolderPath, $importErrorFolderPath, $exportErrorFolderPath, $csvExportFolderPath, $importProcessedFolderPath)

Write-Host "Default error folder:     $errorFolderPath"
Write-Host "Import error folder:      $importErrorFolderPath"
Write-Host "Export error folder:      $exportErrorFolderPath"
Write-Host "CSV export folder:        $csvExportFolderPath"
Write-Host "Processed imports folder: $importProcessedFolderPath"
