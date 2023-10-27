
Import-Module SqlServer

# Init vars
$settingsFilePath           = "D:\Scripts\WAGA_PODS\Settings\settings.txt"

# Get settings params from a settings txt file
$settings = @{}

# Get variablename=filepath strings per each line, split by "=" and store in $settings dict
Get-Content $settingsFilePath | ForEach-Object {
    $key, $val      = $_ -split "=="
    $settings[$key] = $val
}

# Initialize settings from file
$serverName                 = $settings['serverName']
$dbName                     = $settings['dbName']
$invoiceNumsCsvFilePath     = $settings['invoiceNumsCsvFilePath']
$parentFolder               = $settings['parentFolder']
$destinationFolderPath      = $settings['destinationFolderPath']
$query                      = $settings['query']
$foundFiles                 = 0

Write-Host "Settings from file:
            Server name: $serverName
            db name: $dbName
            CSV file path: $invoiceNumsCsvFilePath
            Parent folder: $parentFolder  
            Destination folder: $destinationFolderPath 
            SQL query: $query 
            "

# If destination folder doesn't exist, create it
If (-Not(Test-Path -Path $destinationFolderPath)){New-Item -Path $destinationFolderPath -ItemType Directory}

# Get CSV file data
$csvData = Import-Csv -Path $invoiceNumsCsvFilePath 
# Filter out only invoices that are not found
$invoiceNumbers = $csvData | Where-Object {$_.Found -ne 'Yes' } | Select-Object -ExpandProperty 'InvoiceNumber'

# Check for no returned values
If ($invoiceNumbers.Count -eq 0) {
    Write-Host "No invoices returned from the CSV file $invoiceNumsCsvFilePath"
    Exit
} Else {
    Write-Host "$($invoiceNumbers.Count) Invoice numbers were returned from the CSV file."
}

# Insert invoice numbers from the CSV file into the query insteaf of the INVOICE_NUMBERS placeholder
$query = $query -replace 'INVOICE_NUMBERS', ($invoiceNumbers -join "','")

# Connect and execute query
$connectionString = "Server=$serverName;Database=$dbName;Integrated Security=true;"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

try {
    $connection.Open()

    If ($connection.State -eq 'Open'){

        $command                = $connection.CreateCommand()
        $command.CommandText    = $query
        $reader                 = $command.ExecuteReader()

        Write-Host "Connection established"
exit
        While ($reader.Read()) {

            $invNum         = $reader['invnum']
            $subFolder      = $reader['SubFolder']
            $fileName       = $reader['FileName']
            $customerRef    = $reader['CustomerRef']

            # Check for "\" or "/" in customer ref since it is a manually typed field 
            $customerRef    = $customerRef -replace '\\', ' '
            $customerRef    = $customerRef -replace '\/', ' '

            $sourcePath     = Join-Path -Path $parentFolder -ChildPath (Join-Path -Path $subFolder -ChildPath ($fileName))

            #Write-Host "Processing $sourcePath"

            # If file exists under this path
            If([System.IO.File]::Exists($sourcePath)) {
                # Count found files
                $foundFiles++

                $newFileName            = "$invNum - $customerRef - $fileName"
                $destinationFilePath    = Join-Path -Path $destinationFolderPath -ChildPath ($newFileName)
                Write-Host "Copying source: $sourcePath"
                Write-Host "Copying destination: $destinationFilePath"

                # Copy found file if it doesn't already exist
                If (-Not([System.IO.File]::Exists($destinationFilePath)) ) {
                    try {
                        #CURRENTLY CMDLET NOT WORKING Copy-Item -Path $sourcePath -Destination $destinationFilePath -Force
                        [System.IO.File]::Copy($sourcePath,$destinationFilePath,$true)
                    } catch {
                        Write-Host $_.ExceptionMessage
                    }
                }

                # Report that invnum was found
                ($csvData | Where-Object {$_.InvoiceNumber -eq $invNum} ).Found = 'Yes'
                
            } else {
                Write-Host "File not found."
            }

        }

    } else {
        Throw "Could not establish connection"
    }

} catch {
    Write-Host $_.Exception.Message
} finally {
    $connection.Close()
}

# Compare number of found files with invoice numbers from csv file to see if all files were found
If ($foundFiles -eq $invoiceNumbers.Count) {
    Write-Host "All files were found and copied"
} else {
    If ($foundFiles -eq 0) { Write-Host "No files were found" } Else { Write-Host "Some files where not found. Copied files: $foundFiles for $($invoiceNumbers.Count) invnums. Still missing $($($invoiceNumbers.Count)- $foundFiles) invoices (if list doesn't contain duplications)." }
}

# Write updates back to the CSV file
try {
    $csvData | Export-Csv -Path $invoiceNumsCsvFilePath -NoTypeInformation -Force
} catch {
    # If blocked by another user, save to a new file
    $today = Get-Date -Format "yyyyMMdd"
    $backupCSVexportFilePath = $invoiceNumsCsvFilePath -replace '\.csv$', "_$today.csv"
    $csvData | Export-Csv -Path $backupCSVexportFilePath -NoTypeInformation -Force 
    Write-Host "Original CSv file ws blocked by a user, output to file $backupCSVexportFilePath"
}


