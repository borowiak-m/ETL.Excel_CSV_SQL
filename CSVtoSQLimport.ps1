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
#

# Define the path of the processing settings files to monitor
$processingSettingsFolderPath   = "D:\Scripts\ImportSettings\"
$processingSettingsFiles        = Get-ChildItem -Path $processingSettingsFolderPath -Filter *_settings.text

# Define import files folder
$importFilesFolderPath          = "D:\Scripts\Import\"

$overwriteMode                  = "overwrite"
$appendMode                     = "append"

# Loop through each settings file
ForEach ($settingsFile in $processingSettingsFiles) {
    $settingsFilePath   = Join-Path -Path $processingSettingsFolderPath -ChildPath ($settingsFile.Name)

    # Read the settings file and fetch params
    $settings = @{}

    Get-Content $settingsFilePath | ForEach-Object {
        $key, $val      = $_ -split "="
        $settings[$key] = $val
    }

    $importTable        = $settings['importTable']
    $importTablePK      = $settings['importTablePK']
    $importFileName     = ($settingsFile.BaseName -replace "_settings.txt", "") + ".csv"
    $importFilePath     = Join-Path -Path $importFilesFolderPath -ChildPath ($importFileName)
    $importFieldNames   = $settings['fieldNames']
    $importMode         = $settings['importMode']
    $importServerName   = $settings['importServerName']
    $importDatabaseName = $settings['importDatabaseName']

    # SQL Server connection
    $connectionString               = "Server=$importServerName;Database=$importDatabaseName"
    $connection                     = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString    = $connectionString

    $connection.Open()

    # Check if connection was successfully open 
    # TO BE ERROR HANDLED 
    If ($connection.State -ne 'Open') {Throw "Could not establish connection to $serverName for CSV import process"}

    # If import mode is overwrite
    #   Clear import table before import
    If ($importMode -eq $overwriteMode) {
        $truncateCommand             = $connection.CreateCommand()
        $truncateCommand.CommandText = "TRUNCATE TABLE $importTable"
        $truncateCommand.ExecuteNonQuery()
    }

    # Imprt csv file
    $importFileData                 = Import-Csv -Path $importFilePath

    # Process rows from csv data
    ForEach ($row in $importFileData) {

        $values     = @($importFieldNames.Split(',').ForEach({ $row.$_ }) -join "','")

        # If update flag is append
        If ($importMode -eq $appendMode) {
            # upsert query
            $sqlQuery = "IF EXISTS (SELECT $mportTablePK FROM $importTable WHERE $importTablePK = '$($row.$importTablePK)')
                            UPDATE $importTable
                            SET $fieldNames = $values
                        ELSE
                            INSERT INTO $importTable ($fieldNames) VALUES ('$values')"
        } else {
            # simple insert query
            $sqlQuery = "INSERT INTO $importTable ($fieldNames) VALUES ('$values')"
        }

        $command                        = $connection.CreateCommand()
        $command.CommandText            = $sqlQuery 
        $command.ExecuteNonQuery()

    }

    #Close connection
    $connection.Close()

}

