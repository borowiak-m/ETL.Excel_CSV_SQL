# Initialize settings from file
$settingsFilePath = "D:\Scripts\ExcelImport_settings.txt"

if (Test-Path $settingsFilePath) {
    $settings = @{}
    Get-Content $settingsFilePath | ForEach-Object {
    $splitValue = $_ -split "="
    Write-Host $splitValue[0]
    Write-Host $splitValue[1]
    $settings[$splitValue[0]] = $splitValue[1]
}
} else {
    Throw "Missing settings file in $settingsFilePath"
}


# Read settings into variables

$excelFilePath = $settings['excelFilePath']
$lastTimeFilePath = $settings['lastTimeFilePath']

$lastModifiedTime = (Get-Item $excelFilePath).LastWriteTime

Write-Host "Debug: excelFilePath = $excelFilePath"
Write-Host "Debug: lastModifiedTime = $lastModifiedTime"

if (Test-Path $lastTimeFilePath) {
    $lastKnownTime = New-Object DateTime (Get-Content $lastTimeFilePath)
} else {
    $lastKnownTime = [DateTime]::MinValue
}

if ($lastModifiedTime -gt $lastKnownTime) {
    # File has been modified since last check
    Set-Content $lastTimeFilePath $lastModifiedTime.Ticks
}
