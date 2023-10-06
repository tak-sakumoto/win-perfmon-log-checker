# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out"
)

# Constants
$chartHeight = 300
$chartWidth = 400
$chartType = 4 # 4: line graph

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
# Overwrite all.csv without unnecessary quotes
$csv | Export-Csv -Path $allCSVPath -NoTypeInformation -UseQuotes AsNeeded

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

    # Remove duplicate backslashes
    $outFileName = $outFileName -replace '(\\{2,})', '\'

    # Create the parent folders for the exported CSV file
    $parentPath = Split-Path -Path $outDirPath$outFileName -Parent
    New-Item -Path $parentPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    
    # Export the CSV file for the counter without unnecessary quotes
    $outPath = "$outDirPath$outFileName.csv"
    $csv | Select-Object -Property $counterNames[0], $counterNames[$i] | Export-Csv -Path $outPath -NoTypeInformation -UseQuotes AsNeeded

    # Prepare to handle Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Prepare a workbook with a sheet
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)

    # Set column names to the 1st row of the workbook
    $worksheet.Cells.Item(1, 1) = $counterNames[0]
    $worksheet.Cells.Item(1, 2) = $counterNames[$i]

    # Set data after the 2nd row of the workbook
    $j = 2
    foreach ($line in $csv) {
        $worksheet.Cells.Item($j, 1) = $line.($counterNames[0])
        $worksheet.Cells.Item($j, 2) = $line.($counterNames[$i])
        $j++
    }

    # Add a chart in the workbook
    $chart = $worksheet.Shapes.AddChart2(-1, $chartType, 0, 0, $chartWidth, $chartHeight, $true)

    # Specify a range of data in the workbook
    $dataRange = $worksheet.Range($worksheet.Cells.Item(1, 2), $worksheet.Cells.Item($j, 2))
    $chart.Chart.SetSourceData($dataRange)

    # Save the Excel book
    $outPath = "$outDirPath$outFileName.xlsx"
    $workbook.SaveAs($outPath)

    $workbook.Close()
    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}
