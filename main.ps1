# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out",
    [string]$startTime = "",
    [string]$endTime = ""
)

# Dot sourcing
. .\edit_time_range.ps1
. .\convert_time_format.ps1
. .\save_graph_by_excel.ps1

# Make a foloder to save output files
New-Item -Path $outDirPath -ItemType Directory -Force
$outDirPath = Convert-Path $outDirPath

# Get a file path for the CSV file to store all counters
$allCSVPath = "$outDirPath\all.csv"

# Search for invalid characters in file paths, excluding backslashes
$invalidChars = [IO.Path]::GetInvalidFileNameChars() | Where-Object { $_ -ne "\\" }

# Convert the specified blg file to a CSV file
Start-Process -NoNewWindow -Wait -FilePath "relog.exe" -ArgumentList "$blgPath", "-f", "CSV", "-o", "$allCSVPath"

# Load all.csv
$csv = Import-Csv -Path $allCSVPath

# Get counter names from the CSV file
$counterNames = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# Extract between the start time and the end time
$csv = Edit-Time-Range -startTime $startTime -endTime $endTime -timeColName $counterNames[0] -csv $csv

# Convert the time format MM/dd/yyyy HH:mm:ss.fff into yyyy/MM/dd HH:mm:ss.fff
$csv = Convert-Time-Format -timeColName $counterNames[0] -csv $csv

# Overwrite all.csv without unnecessary quotes
$csv | Export-Csv -Path $allCSVPath -NoTypeInformation -UseQuotes AsNeeded

# Export a CSV file for each counter
for ($i = 1; $i -lt $counterNames.Count; $i++) {
    # Get a name of a CSV file for storing a counter
    $outFileName = $counterNames[$i]

    # Replace invalid characters in the file name to underscores 
    foreach($char in $invalidChars) {
        $escapedChar = [Regex]::Escape($char)
        if($counterName -match $escapedChar) {
            $outFileName = $outFileName -replace $escapedChar,"_"
        }
    }

    # Remove duplicate backslashes
    $outFileName = $outFileName -replace '(\\{2,})', '\'

    # Create the parent folders for the exported CSV file
    $parentPath = Split-Path -Path $outDirPath$outFileName -Parent
    New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    
    # Export the CSV file for the counter without unnecessary quotes
    $outPath = "$outDirPath$outFileName.csv"
    $csv | Select-Object -Property $counterNames[0], $counterNames[$i] | Export-Csv -Path $outPath -NoTypeInformation -UseQuotes AsNeeded
    
    # Get values in the i-th column
    $columnValues = $csv | ForEach-Object { $_.($counterNames[$i]) }

    # Get an array of stats for the i-th column
    $stats = @{
        Count = $columnValues.Count
        Maximum = ($columnValues | Measure-Object -Maximum).Maximum
        Minimum = ($columnValues | Measure-Object -Minimum).Minimum
        Average = ($columnValues | Measure-Object -Average).Average
    }
    $outPath = "$outDirPath$outFileName" + "_stats.csv"

    # Export the stats array
    $stats | Select-Object Count, Maximum, Minimum, Average | Export-Csv -Path $outPath -NoTypeInformation -UseQuotes AsNeeded
    
    # Draw a line graph for the counter and save as an Excel workbook
    $outPath = "$outDirPath$outFileName.xlsx"
    Save-Graph-By-Excel -outPath $outPath -xAxisName $counterNames[0] -yAxisName $counterNames[$i] -csv $csv
}
