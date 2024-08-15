# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out",
    [string]$startTime = "",
    [string]$endTime = ""
)

# Dot sourcing
. .\Edit-TimeRange.ps1
. .\Convert-TimeFormat.ps1
. .\Get-ValidName.ps1
. .\Get-Stats.ps1
. .\Save-GraphByExcel.ps1

# Make a foloder to save output files
New-Item -Path $outDirPath -ItemType Directory -Force
$outDirPath = Convert-Path $outDirPath

# Get a file path for the CSV file to store all counters
$allCSVPath = "$outDirPath\all.csv"

# Convert the specified blg file to a CSV file
Start-Process -NoNewWindow -Wait -FilePath "relog.exe" -ArgumentList "$blgPath", "-f", "CSV", "-o", "$allCSVPath"

# Load all.csv
$csv = Import-Csv -Path $allCSVPath

# Get counter names from the CSV file
$counterNames = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# Extract between the start time and the end time
$csv = Edit-TimeRange -startTime $startTime -endTime $endTime -timeColName $counterNames[0] -csv $csv

# Convert the time format MM/dd/yyyy HH:mm:ss.fff into yyyy/MM/dd HH:mm:ss.fff
$csv = Convert-TimeFormat -timeColName $counterNames[0] -csv $csv

# Overwrite all.csv without unnecessary quotes
$csv | Export-Csv -Path $allCSVPath -NoTypeInformation -UseQuotes AsNeeded

# Export a CSV file for each counter
for ($i = 1; $i -lt $counterNames.Count; $i++) {
    # Replace invalid characters in the file name to underscores
    $outFileName = Get-ValidName -fileName $counterNames[$i]

    # Remove duplicate backslashes
    $outFileName = $outFileName -replace '(\\{2,})', '\'

    # Create the parent folders for the exported CSV file
    $parentPath = Split-Path -Path $outDirPath$outFileName -Parent
    New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    
    # Export the CSV file for the counter without unnecessary quotes
    $outPath = "$outDirPath$outFileName.csv"
    $csv | Select-Object -Property $counterNames[0], $counterNames[$i] | Export-Csv -Path $outPath -NoTypeInformation -UseQuotes AsNeeded
    
    # Get an array of stats for the i-th column
    $outPath = "$outDirPath$outFileName" + "_stats.csv"
    Get-Stats -outPath $outPath -$colName $counterNames[$i] -csv $csv

    # Draw a line graph for the counter and save as an Excel workbook
    $outPath = "$outDirPath$outFileName.xlsx"
    Save-GraphByExcel -outPath $outPath -xAxisName $counterNames[0] -yAxisName $counterNames[$i] -csv $csv
}
