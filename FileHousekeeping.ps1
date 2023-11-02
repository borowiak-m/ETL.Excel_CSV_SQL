#  LOGIC EXPLAINED
# ---------------
# This script aims to clear any old files from the monitored folder paths. We achieve this following this logic:
#
# - Read process params (folder paths) from a settings text file
# - Loop through the monitored folders and check file age
# - If file is older than X days, delete the file
# - Folders to monitor for old files: $errorFolderPath, $csvExportFolderPath, $importProcessedFolderPath:
#     - $errorFolderPath - path where erros logs are kept and files not processed
#     - $csvExportFolderPath - path where csv sheets are exported from business docs, and imported from to db 
#     - $importProcessedFolderPath - path where processed import csv files are moved

# Initialize default logging and settings folders

$processingSettingsFolderPath   = "D:\Scripts\Stock Blackboard\Settings\"
$processingSettingsFiles        = Get-ChildItem -Path $processingSettingsFolderPath -Filter *settings.txt
$errorFolderPath                = "D:\Scripts\Stock Blackboard\Error\"

# Check if settings files exist
If (-Not (Test-Path $processingSettingsFolderPath)) { Throw "No settings folder found" }

# Check that we found any settings files in the folder
If (-Not ($processingSettingsFiles.Count -gt 0)) { Throw "No settings files were found" }


