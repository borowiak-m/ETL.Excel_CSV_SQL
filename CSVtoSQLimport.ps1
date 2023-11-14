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



function Write-Error($errorFolderPath, $errorMsg, $errorLvl) {

    # Log error in cmds
    Write-Host $errorMsg                            

    # Generate a timestamp for error file name
    $timestamp = Get-Date -format "yyyy.MM.dd HH.mm"

    # Generate a date for error file name
    $errorDate = Get-Date -format "yyyyMMdd"

    # Generate a unique error file path
    $errorLogFilePath = Join-Path -Path $errorFolderPath -ChildPath("$errorDate FileImportError.txt")

    # Check if file already exists, if so append the error message to existing file, if not create a new error file
    If (Test-Path $errorLogFilePath) {
        Add-Content $errorLogFilePath "$timestamp File name: $importFileName. Error: $errorMsg"
    } else {
        Set-Content $errorLogFilePath "$timestamp File name: $importFileName. Error: $errorMsg"
    }

    If ($errorLvl -eq "Fatal") {
        Write-Host "Debug: Fatal error, exiting program."
        Exit
    } else {
        Write-Host "Debug: Error of level $errorLvl. Recommencing program."
    }

}

function EncloseWithBrackets($name) {
    # Check for spaces in SQL object names and enslose in [] if found
    If ($name -like "* *") {
        return "[$name]"
    } else {
        return $name
    }

}

function SanitizeString($inputString) {
    # Define forbidden characters
    $forbiddenChars     = @("'",";","--")

    # Define a list of common SQL syntaxt words to look out for
    $sqlSyntaxWords     = @("SELECT", "DROP", "INSERT", "DELETE", "UPDATE", "EXEC", "EXECUTE", "ALTER", "CREATE", "GRANT", "REVOKE", "TRUNCATE", "TABLE", "TABLES"
                            "select", "drop", "insert", "delete", "update", "exec" , "execute", "alter", "create", "grant", "revoke", "truncate", "table", "tables")

    # Remove forbidden characters
    ForEach ($char in $forbiddenChars) {
        $inputString    = $inputString.Replace($char,'')
    }

    # Alter any SQL syntax words
    ForEach ($word in $sqlSyntaxWords) {
        $inputString = $inputString.Replace($word, "[[$word]]")
    }

    return $inputString
}

function ConvertExcelDateToSQL($excelDate) {
    Write-Host "Passed excel date value: $excelDate"
    $origin = [datetime]"1900-01-01"
    try {
        $date   = [datetime]$origin.AddDays([double]$excelDate)
    }
    catch {
        Write-Error $errorFolderPath "Incorrect date format in value $excelDate" NotFatal
        return $null
    }
    
    return  $date.ToString("yyyy-MM-dd HH:mm:ss")
}

# Initialize default error folder locations and file names
$processingSettingsFolderPath   = "D:\Scripts\Stock Blackboard\Settings\"
$processingSettingsFiles        = Get-ChildItem -Path (Join-Path $processingSettingsFolderPath -ChildPath "import files") -Filter *_import_settings.txt
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

If (-Not([String]::IsNullOrEmpty($importSettings['errorFolderPath']))) {$errorFolderPath = $importSettings['errorFolderPath']}

$paramsToCheck = @($lastImpLogFolderPath, $importFilesFolderPath, $importProcessedFolderPath, $overwriteMode, $appendMode)

# Check for empty params 
ForEach ($param in $paramsToCheck) { If ([string]::IsNullOrEmpty($param)) { Write-Error $errorFolderPath "Params missing. Review settings file under $mainImportSettingsFilePath" Fatal} }

# Check for existence of last time imported log folder and the import folder
If (-Not (Test-Path $importFilesFolderPath)) {Write-Error $errorFolderPath "Import folder missing. Review settings file under $mainImportSettingsFilePath" Fatal}
If (-Not (Test-Path $lastImpLogFolderPath)) {New-Item -Path $lastImpLogFolderPath -ItemType Directory}
If (-Not (Test-Path $importProcessedFolderPath)) {New-Item $importProcessedFolderPath -ItemType Directory}

# Loop through each settings file
ForEach ($settingsFile in $processingSettingsFiles) {

    # Read the settings file and fetch params
    $settings = @{}

    Get-Content $settingsFile.FullName | ForEach-Object {
        $key, $val      = $_ -split "=="
        $settings[$key] = $val
    }

    $importTable                    = $settings['importTable']
    $importTablePK                  = $settings['importTablePK']
    $importFieldNames               = $settings['importFieldNames']
    $importServerName               = $settings['importServerName']
    $importDatabaseName             = $settings['importDatabaseName']
    $importDatetimeFields           = $settings['importDatetimeFields']

    # Check if any datetime conversion is required
    $isDatetimeConversionRequired = $false
    ForEach ($field in $importDatetimeFields) { If (-Not([string]::IsNullOrEmpty($field))) { $isDatetimeConversionRequired = $true } }
    Write-Host "Datetime conversion is required? : $isDatetimeConversionRequired"

    # Processing file path
    $importFileName                 = ($settingsFile.BaseName -replace "_import_settings", "") + ".csv"
    $importFilePath                 = Join-Path -Path $importFilesFolderPath -ChildPath ($importFileName)

    # Enclose in [] names with spaces if required
    $enclosedImportTable            = EncloseWithBrackets $importTable

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

    # Clear table before import
    try {
        $truncateCommand             = $connection.CreateCommand()
        $truncateCommand.CommandText = "TRUNCATE TABLE $enclosedImportTable"
        $truncateCommand.ExecuteNonQuery()
    }
    catch {
        # Log error if sql command didn't execute as excpected  
        Write-Error $errorFolderPath $_.Exception.Message NotFatal
        # Skip to next file in the loop
        Continue
    }

    # Import csv file
    $importFileData              = Import-Csv -Path $importFilePath

    # Process rows from csv data
    ForEach ($row in $importFileData) {

        # Validate CSV properties (.Name = header name and .Value = actual row value)
        $row.PSObject.Properties | ForEach-Object { 
            # Sanitize all string values for SQL input
            $_.Value = SanitizeString $_.Value 

            # If in row there is a defined datetime field for conversion, convert it to SQL format
            If ($isDatetimeConversionRequired){
                If ($importDatetimeFields -contains $_.Name) { 
                    $_.Value = ConvertExcelDateToSQL $_.Value
                }
            }
        }

        # If primary key field is empty, skip the row
        If ([String]::IsNullOrEmpty($row.$importTablePK)) {Continue}

        # If all fields are inserted as they are in the CSV file
        If ($importFieldNames -eq "All") {
            $values                 = ($row.PSObject.Properties.Value) -join "','"
        } else {
            # If inserts are only for specified fields, fetch their values only
            $values                 = $importFieldNames.Split(',').ForEach({ $row.$_ }) -join "','"
        }

        $sqlQuery               = "INSERT INTO $enclosedImportTable VALUES ('$values')"
        Write-Host "SQL query: $sqlQuery"
        
        try {
            $command                        = $connection.CreateCommand()
            $command.CommandText            = $sqlQuery 
            $command.ExecuteNonQuery()
        } 
        catch {
            Write-Error $errorFolderPath $_.Exception.Message NotFatal
            # Skip to next row in the loop
            Continue
        }
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

