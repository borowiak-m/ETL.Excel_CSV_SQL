#  SCRIPT MISSION
# ---------------
# This script aims to export data in an Excel file. We achieve this following this logic:
#
# - Check when was the file last written to, if there are new changes. Then compare last modified date and time to a stored
#   value of date and time of the last data extract. Should the modification time be more recent, this triggers a new export.
#
# - Load the excel file and export all data with a non-empty column B. This will be exported into a csv file, but first we need 
#   to check if the previously exported csv file is still in the export folder. This file is to be picked up by another process.
#   If the previous export is still sat in the folder, we move it to an Error folder and replace it with a new export instead.
#
# To be continued:
# ----------------
# - How to notify users that there was a csv file left behind from previous cycle
# - Potential housekeeping of old files in Error folder
# - potential clashes with excel sheet being edited at the same time or if open


# Initialize settings from file, where all business folder paths are stored

$settingsFilePath = "D:\Scripts\ExcelImport_settings.txt"

if (Test-Path $settingsFilePath) {
    $settings = @{}

    # Get variablename=filepath strings per each line, split by "=" and store in $settings dict

    Get-Content $settingsFilePath | ForEach-Object {
    $splitValue = $_ -split "="
    Write-Host $splitValue[0] #Print in command line
    Write-Host $splitValue[1] #Print in command line
    $settings[$splitValue[0]] = $splitValue[1]
}
} else {
    Throw "Missing settings file in $settingsFilePath"
}


# Read settings into variables

$excelFilePath = $settings['excelFilePath']
$lastTimeFilePath = $settings['lastTimeFilePath']

$lastModifiedTime = (Get-Item $excelFilePath).LastWriteTime

Write-Host "Debug: excelFilePath = $excelFilePath" #Print in command line
Write-Host "Debug: lastModifiedTime = $lastModifiedTime" #Print in command line

if (Test-Path $lastTimeFilePath) {

    # If file exists get the last modified datetime
    $lastKnownTime = New-Object DateTime (Get-Content $lastTimeFilePath)
} else {
    $lastKnownTime = [DateTime]::MinValue
}

if ($lastModifiedTime -gt $lastKnownTime) {

    # File has been modified since last check
    Set-Content $lastTimeFilePath $lastModifiedTime.Ticks
}
