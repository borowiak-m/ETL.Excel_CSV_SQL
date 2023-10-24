#  LOGIC EXPLAINED
# ---------------
# This script aims to import data from csv files into SQL tables. How this is done is as follows:
#
# - Script monitors for [fileName]_settings.txt files in a folder. Each *_settings file is a collection of instructions
#   on how to import and what to import. The fileName prefix in the file name correspond to a fileName.csv file to be
#   processed. Its settings file will contain all fields that we are importing, and also the same fild names we will
#   be updating in the database table specified in this settings file.
#
# - We iterate over this collection of files, gather a list of files to be imported and process each csv one by one, 
#   updating their corresponding database tables.
#
# - Any new files will need a table created with same field names as in the csv file and a settings file created and dropped in 
#   the monitored folder. 
#
# - There will be two types of data imports. And append and an overwrite, where some tables collect ongoing business data,
#   where others will be config updates which will overwrite an entire table. These will be small in size. The indicating
#   flag on how to update data will also be held in a settings file, particular to its import.
#


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

# Initialize default error folder locations and file names
$processingSettingsFolderPath   = "D:\Scripts\Stock Blackboard\Settings\"
$processingSettingsFiles        = Get-ChildItem -Path $processingSettingsFolderPath -Filter *_import_settings.txt
$errorFolderPath                = "D:\Scripts\Stock Blackboard\Error\"
$lastImpLogFileName             = "last_time_imported.txt"
$mainImportSettingsFileName     = "import_settings.txt"
$mainImportSettingsFilePath     = Join-Path -Path $processingSettingsFolderPath -ChildPath ($mainImportSettingsFileName)

# Check for existence of the main extract settings file
If (-Not(Test-Path $mainImportSettingsFilePath)) {Throw "No main import settings document found."}

# Check for existence of error folder
If (-Not (Test-Path $errorFolderPath)) {New-Item -Path $errorFolderPath -ItemType Directory}

# Get params from main extract settings file
$importSettings = @{}

# Get variablename=filepath strings per each line, split by "=" and store in $settings dict
Get-Content $mainImportSettingsFilePath | ForEach-Object {
    $key, $val      = $_ -split "="
    $importSettings[$key] = $val
}

# Initialize settings from file, where all business folder paths are stored
$lastImpLogFolderPath               = $importSettings['lastImpLogFolderPath']
$importFilesFolderPath              = $importSettings['importFilesFolderPath']
$importProcessedFolderPath          = $importSettings['importProcessedFolderPath']
$overwriteMode                      = $importSettings['overwriteMode']
$appendMode                         = $importSettings['appendMode']

$paramsToCheck = @($lastImpLogFolderPath, $importFilesFolderPath, $importProcessedFolderPath, $overwriteMode, $appendMode)

# Check for empty params 
ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { Write-Error $errorFolderPath "Params missing. Review settings file under $mainImportSettingsFilePath" Fatal} }

# Check for existence of last time imported log folder and the import folder
If (-Not (Test-Path $importFilesFolderPath)) {Write-Error $errorFolderPath "Import folder missing. Review settings file under $mainImportSettingsFilePath" Fatal}
If (-Not (Test-Path $lastImpLogFolderPath)) {New-Item -Path $lastImpLogFolderPath -ItemType Directory}
If (-Not (Test-Path $importProcessedFolderPath)) {New-Item $importProcessedFolderPath -ItemType Directory}

# Loop through each settings file
ForEach ($settingsFile in $processingSettingsFiles) {

    # Build current settings file path
    $settingsFilePath   = Join-Path -Path $processingSettingsFolderPath -ChildPath ($settingsFile.Name)

    # Read the settings file and fetch params
    $settings = @{}

    Get-Content $settingsFilePath | ForEach-Object {
        $key, $val      = $_ -split "="
        $settings[$key] = $val
    }

    $importTable                    = $settings['importTable']
    $importTablePK                  = $settings['importTablePK']
    $importFieldNames               = $settings['importFieldNames']
    $importMode                     = $settings['importMode']
    $importServerName               = $settings['importServerName']
    $importDatabaseName             = $settings['importDatabaseName']

    # Processing file path
    $importFileName                 = ($settingsFile.BaseName -replace "_import_settings", "") + ".csv"
    $importFilePath                 = Join-Path -Path $importFilesFolderPath -ChildPath ($importFileName)

    # Check if there is a file to pick up
    If (-Not (Test-Path $importFilePath)) { 
        Write-Host "No file $importFilePath to process. Moving to the next file"
        Continue 
    }

    # Check for empty params 
    $hasEmptyParams                 = $false
    $paramsToCheck                  = @($importTable, $importTablePK, $importFieldNames, $importMode, $importServerName, $importDatabaseName)
    ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { $hasEmptyParams = $true } }

    If ($hasEmptyParams){ 
        Write-Error $errorFolderPath "Params missing. File $importFileName is skipped from extract process. Review settings file under $settingsFilePath." NotFatal
        # Leaving file in folder for next pickup cycle
        # Need to communicate to users / service team that the file or settings file needs attention as no updates are happening while unresolved
        Continue
    }

    # SQL Server connection
    $connectionString               = "Server=$importServerName;Database=$importDatabaseName"
    $connection                     = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString    = $connectionString

    $connection.Open()

    # Check if connection was successfully open 
    If ($connection.State -ne 'Open') {
        Write-Error $errorFolderPath "Could not establish connection to $importServerName for CSV import process for file $importFilePath." NotFatal
        # Leaving file in folder for next pickup cycle
        # Need to communicate to users / service team that the file or settings file needs attention as no updates are happening while unresolved
    }

    # Check if destination table exists
    try {
        $tableCheckCommand             = $connection.CreateCommand()
        $tableCheckCommand.CommandText = "SELECT TOP 1 * FROM $importTable"
        $tableCheckCommand.ExecuteNonQuery()
    }
    catch {
        Write-Error $errorFolderPath "Could not connect to $importTable. File $importFileName can't be processed." NotFatal
        # Leaving file in folder for next pickup cycle
        # Need to communicate to users / service team that the file or settings file needs attention as no updates are happening while unresolved
        Continue
    }

    # If import mode is overwrite
    # Clear import table before import
    If ($importMode -eq $overwriteMode) {
        $truncateCommand             = $connection.CreateCommand()
        $truncateCommand.CommandText = "TRUNCATE TABLE $importTable"
        $truncateCommand.ExecuteNonQuery()
    }

    # Import csv file
    $importFileData                 = Import-Csv -Path $importFilePath

    # Process rows from csv data
    ForEach ($row in $importFileData) {

        $values     = @($importFieldNames.Split(',').ForEach({ $row.$_ }) -join "','")
        $updates    = ($importFieldNames -split "," | ForEach-Object {"$_ = '$($row.$_)'"}) -join ","

        # If update flag is append
        If ($importMode -eq $appendMode) {
            # upsert query
            $sqlQuery = "IF EXISTS (SELECT $importTablePK FROM $importTable WHERE $importTablePK = '$($row.$importTablePK)')
                            UPDATE $importTable
                            SET $updates
                            WHERE $importTablePK = '$($row.$importTablePK)'
                        ELSE
                            INSERT INTO $importTable ($importFieldNames) VALUES ('$values')"
        } else {
            # simple insert query
            $sqlQuery = "INSERT INTO $importTable ($importFieldNames) VALUES ('$values')"
        }

        $command                        = $connection.CreateCommand()
        $command.CommandText            = $sqlQuery 
        $command.ExecuteNonQuery()
        #Write-Host $sqlQuery
    }

    # Close connection
    $connection.Close()

    # Generate a timestamp for processed file name
    $timestamp                          = Get-Date -format "yyyy.MM.dd HH.mm"

    # Processed file path with a timestamp
    $importProcessedFilePath            = Join-Path -Path $importProcessedFolderPath -ChildPath ("$timestamp $importFileName")

    # Move processed file to Completed folder
    Move-Item -Path $importFilePath -Destination $importProcessedFilePath

    # Log update
    $lastImpLogFilePath = Join-Path -Path $lastImpLogFolderPath -ChildPath ($importFileName + "_" + $lastImpLogFileName)
    Set-Content $lastImpLogFilePath (Get-Date)

}

