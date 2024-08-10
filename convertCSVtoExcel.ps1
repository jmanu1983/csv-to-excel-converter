<#
.SYNOPSIS
    Batch CSV to Excel converter.

.DESCRIPTION
    Converts all CSV files in the input directory to Excel (.xlsx) format,
    then renames the processed CSV files with a timestamp suffix.

.NOTES
    Requires the ImportExcel PowerShell module.
    Install with: Install-Module -Name ImportExcel -Force
#>

# Ensure the ImportExcel module is available
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Resolve directories relative to the script location
$scriptRoot = $PSScriptRoot
$csvDirectory   = Join-Path $scriptRoot "input"
$excelDirectory = Join-Path $scriptRoot "output"
$logDirectory   = Join-Path $scriptRoot "logs"

# Check if log directory exists, if not create it
if (!(Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory
}

# Get current date and time for log file
$currentDateTime = Get-Date -Format "ddMMyyyy_HHmm"

# Log file name
$logFileName = "Log_" + $currentDateTime + ".txt"

# Full log file path
$logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName

# Enumerate all CSV files
Get-ChildItem -Path $csvDirectory -Filter *.csv | ForEach-Object {
    # Log CSV file opening
    Add-Content -Path $logFilePath -Value ("Ouvert le fichier CSV : " + $_.FullName)

    # Import CSV file content
    $csvContent = Import-Csv $_.FullName -Encoding Default -Delimiter ";"

    # Excel file path
    $excelFilePath = Join-Path -Path $excelDirectory -ChildPath ($_.BaseName + ".xlsx")

    try {
        # Convert CSV to Excel
        $csvContent | Export-Excel -Path $excelFilePath -WorksheetName Sheet1 -AutoSize

        # Log Excel file saving
        Add-Content -Path $logFilePath -Value ("Sauvegardé en tant que fichier Excel : " + $excelFilePath)
    }
    catch {
        # Log error
        Add-Content -Path $logFilePath -Value ("Erreur lors de la sauvegarde en tant que fichier Excel : " + $excelFilePath)
    }

    # Get current date and time for renaming CSV file
    $currentDateTime = Get-Date -Format "ddMMyyyy_HHmm"

    # Renamed CSV file path
    $renamedCsvFilePath = Join-Path -Path $csvDirectory -ChildPath ($_.BaseName + "_" + $currentDateTime + ".csv")

    # Rename CSV file
    Rename-Item -Path $_.FullName -NewName $renamedCsvFilePath

    # Log CSV file renaming
    Add-Content -Path $logFilePath -Value ("Renommé le fichier CSV : " + $renamedCsvFilePath)
}
