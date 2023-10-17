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

# Process rows from csv data
ForEach ($row in $csvData) {
    ForEach ($field in $fieldsToImport) {
        # -----
        #assign $row.$field to a viariable for each field
        # -----
        $upsertQuery = @"
IF EXISTS (SELECT 1 FROM table_name WHERE PK_field = $PK_field)
    UPDATE table_name 
    SET
    field = $field
    WHERE PK_field = $PK_field
ELSE
    INSERT INTO table_name (field1, field2)
    VALUES ($field1, $field2)
"@
        $command = New-Object System.Data.SqlClient.SqlCommand($upsertQuery, $connection)
        $command.ExecuteNonQuery()
    }
}

#Close connection
$connection.Close()

