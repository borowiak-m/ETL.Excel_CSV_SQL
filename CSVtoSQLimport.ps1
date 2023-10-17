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


# Define csv file to import
$csvFolderPath                  = "D:\Scripts\Import\"
$csvFileName                    = "CSVimport.csv"
$csvFilePath                    = Join-Path -Path $csvFolderPath -ChildPath ($csvFileName)

# Define header variables if required to import only selected field from CSV
$fieldsToImport                = @("ProductDepot", "Commentary")

# Imprt csv file
$csvData                        = Import-Csv -Path $csvFilePath

# SQL Server connection
$serverName                     = ""
$databaseName                   = ""
$connectionString               = "Server=$serverName;Database=$databaseName"
$connection                     = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString    = $connectionString

$connection.Open()

# Check if connection was successfully open 
# TO BE ERROR HANDLED 
If ($connection.State -ne 'Open') {Throw "Could not establish connection to $serverName for CSV import process"}

# Process rows from csv data
ForEach ($row in $csvData) {
    ForEach ($field in $fieldsToImport) {
        # -----
        #assign $row.$field to a viariable for each field
        # -----
        $upsertQuery = @"
IF EXISTS (SELECT 1 FROM table_name WHERE ProductDepot = $ProductDepot)
    UPDATE table_name 
    SET
    Commentary = $Commentary
    WHERE ProductDepot = $ProductDepot
ELSE
    INSERT INTO table_name (ProductDepot, Commentary)
    VALUES ( $ProductDepot, $Commentary)
"@
        $command = New-Object System.Data.SqlClient.SqlCommand($upsertQuery, $connection)
        $command.ExecuteNonQuery()
    }
}

#Close connection
$connection.Close()

