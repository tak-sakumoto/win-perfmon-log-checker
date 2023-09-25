# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out"
)

New-Item -Path $outDirPath -ItemType Directory -Force
$allCSVPath = "$outDirPath\all.csv"

$invalidChars = [IO.Path]::GetInvalidFileNameChars()

# Convert the specified blg file to a CSV file
Start-Process -NoNewWindow -Wait -FilePath "relog.exe" -ArgumentList "$blgPath", "-f", "CSV", "-o", "$allCSVPath"

# Load the CSV file 
$csv = Import-Csv -Path $allCSVPath

# Get counter names from the CSV file
$counterNames = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# Export a CSV file for each counter
for ($i = 1; $i -lt $counterNames.Count; $i++) {
    $counterName = $counterNames[$i]
    foreach($char in $invalidChars) {
        $escapedChar = [Regex]::Escape($char)
        if($counterName -match $escapedChar) {
            $outFileName = $counterName -replace $escapedChar,"_"
        }
    }
    $outFileName = $outFileName -replace "\\", " "

    $outPath = "$outDirPath\$outFileName.csv"
    $csv | Select-Object -Property $counterNames[0], $counterNames[$i] | Export-Csv -Path $outPath -NoTypeInformation
}
