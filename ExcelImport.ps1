#  LOGIC EXPLAINED
# ---------------
# This script aims to export data in an Excel file. We achieve this following this logic:
#
# - Check when was the file last written to, if there are new changes. Then compare last modified date and time to a stored
#   value of date and time of the last data extract. Should the modification time be more recent, this triggers a new export.
#
# - Load the excel file and export all data from a given sheet name. This will be exported into a csv file, but first we need 
#   to check if the previously exported csv file is still in the export folder. This file is to be picked up by another process.
#   If the previous export is still sat in the folder, we move it to an Error folder and replace it with a new export instead.
#
# To be continued:
# ----------------
# - How to notify users that there was a csv file left behind from previous cycle
# - Potential housekeeping of old files in Error folder
# - potential clashes with excel sheet being edited at the same time or if open

# Import the ImportExcel module
Import-Module -Name ImportExcel

function Write-Error($errorFolderPath, $errorMsg, $riskLvl) {

    Write-Host $errorMsg
    Write-Host $errorFolderPath

    # Generate a timestamp for error file name
    $timestamp = Get-Date -format "yyyy.MM.dd hh.mm.ss"

    # Generate a unique error file path
    $errorLogFilePath = Join-Path -Path $errorFolderPath -ChildPath("FileImportError $timestamp.txt")

    # Write error message to a file
    Set-Content $errorLogFilePath $errorMsg

    If ($riskLvl -eq "Fatal") {Exit}

}

# Initialize settings from file, where all business folder paths are stored
$settingFileName        = "ExcelImport_settings.txt"
$settingsFolderPath     = "D:\Scripts\"
$settingsFilePath       = Join-Path -Path $settingsFolderPath -ChildPath ($settingFileName)

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

ForEach ($param in $paramsToCheck) {
    
    $paramLength = $param | Measure-Object -Character | Select-Object -expand Characters

    Write-Host $param "Length : " $paramLength

    If ($paramLength -eq 0) {Throw "Params missing. Review settings file under $settingsFilePath"} 
}

# Check for the existence of excel document
If (-Not(Test-Path $excelFilePath)) {Write-Error $errorFolderPath "No document found under $excelFilePath" "Fatal"}

# Check if necessary folders exist, if not create them
$foldersToCheck = @($csvExportFolderPath, $csvErrorFolderPath)

ForEach ($folder in $foldersToCheck) {
    If (-Not (Test-Path $folder)){
        Write-Error $errorFolderPath "Folder $folder not found. Created a new directory." "Continue"
        New-Item -Path $folder -ItemType Directory
    }
}

# Fetch last edit time of the excel document
$lastModifiedTime = (Get-Item $excelFilePath).LastWriteTime

Write-Host "Debug: excelFilePath                = $excelFilePath"
Write-Host "Debug: excelFilePath.Length         = " ($excelFilePath | Measure-Object -Character | Select-Object -expand Characters)
Write-Host "Debug: lastModifiedTime             = $lastModifiedTime"
Write-Host "Debug: lastModifiedTime.Length      = " ($lastModifiedTime | Measure-Object -Character | Select-Object -expand Characters)
Write-Host "Debug: csvExportFolderPath          = $csvExportFolderPath"
Write-Host "Debug: csvExportFolderPath.Length   = " ($csvExportFolderPath | Measure-Object -Character | Select-Object -expand Characters)
Write-Host "Debug: errorFolderPath              = $errorFolderPath"
Write-Host "Debug: errorFolderPath.Length       = " ($errorFolderPath | Measure-Object -Character | Select-Object -expand Characters)
Write-Host "Debug: sheetsToExport               = $sheetsToExport"
Write-Host "Debug: sheetsToExport.Length        = " ($sheetsToExport | Measure-Object -Character | Select-Object -expand Characters)

if (Test-Path $lastTimeFilePath) {

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

    $sheets = Get-ExcelSheetInfo $excelFilePath | Select-Object -ExpandProperty Name

    $matchCounter = 0

    ForEach ($sheet in $sheets) {

        ForEach ($sheetToExport in $sheetsToExport) {

            Write-Host ("Debug: Processing worksheet: " + $sheet)
            Write-Host ("Debug: Matching with sheet name: " + $sheetToExport)
            
            If ($sheet -eq $sheetToExport) {
                Write-Host "Debug:  - - - > $sheet and $sheetToExport [MATCH]"
                $matchCounter = $matchCounter + 1

                # Build file path for csv file export with sheet name
                $csvExportFilePath = Join-Path -Path $csvExportFolderPath -ChildPath ("$sheet.csv")

                # Read from excel file
                $allData = Import-Excel -Path $excelFilePath 

                # Check if an export csv file already exists (if so move it to Error folder and replace it)
                If (Test-Path $csvExportFilePath) {

                    # Log unprocessed file
                    Write-Error $errorFolderPath "Unprocessed file $csvExportFilePath. File renamed and moved to $errorFolder." "Continue"

                    # Generate a timestamp for error file name
                    $timestamp = Get-Date -format "yyyy.MM.dd hh.mm.ss"

                    # Move existing file to error folder and rename it with timestamp
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


