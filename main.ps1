# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out"
)

# Make a foloder to save output files
New-Item -Path $outDirPath -ItemType Directory -Force

# Get a file path for the CSV file to store all counters
$allCSVPath = "$outDirPath\all.csv"

# Search for invalid characters in file paths, excluding backslashes
$invalidChars = [IO.Path]::GetInvalidFileNameChars() | Where-Object { $_ -ne "\\" }

# Convert the specified blg file to a CSV file
Start-Process -NoNewWindow -Wait -FilePath "relog.exe" -ArgumentList "$blgPath", "-f", "CSV", "-o", "$allCSVPath"

# Load the CSV file 
$csv = Import-Csv -Path $allCSVPath

# Get counter names from the CSV file
$counterNames = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

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

    # Create the parent folders for the exported CSV file
    $parentPath = Split-Path -Path $outDirPath\$outFileName -Parent
    New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    
    # Export the CSV file for the counter
    $outPath = "$outDirPath\$outFileName.csv"
    $csv | Select-Object -Property $counterNames[0], $counterNames[$i] | Export-Csv -Path $outPath -NoTypeInformation
}
