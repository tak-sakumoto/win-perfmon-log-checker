# Arguments
param (
    [string]$blgPath,
    [string]$outDirPath = ".\out",
    [string]$startTime = "",
    [string]$endTime = ""
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

# Get counter names from the CSV file
$counterNames = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# Extract between the start time and the end time
if ($startTime.Trim().Length -ne 0) {
    $startTime = [DateTime]::ParseExact($startTime, "yyyy/MM/dd HH:mm:ss", $null)
    $csv = $csv | Where-Object { $_.($counterNames[0]) -ge $startTime }
}
if ($endTime.Trim().Length -ne 0) {
    $endTime = [DateTime]::ParseExact($endTime, "yyyy/MM/dd HH:mm:ss", $null)
    $csv = $csv | Where-Object { $_.($counterNames[0]) -le $endTime }
}

# Convert the time format
$csv = $csv | ForEach-Object {
    $time = [DateTime]::ParseExact($_.($counterNames[0]), "MM/dd/yyyy HH:mm:ss.fff", $null)
    $_.($counterNames[0]) = $time.ToString("yyyy/MM/dd HH:mm:ss.fff")
    $_
}

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
    $seriesCollection = $chart.Chart.SeriesCollection()
    $series = $seriesCollection.NewSeries()
    $series.XValues = $worksheet.Range($worksheet.Cells.Item(2, 1), $worksheet.Cells.Item($j, 1))
    $series.Values = $worksheet.Range($worksheet.Cells.Item(2, 2), $worksheet.Cells.Item($j, 2))

    # Get the axes
    $xAxis = $chart.Chart.Axes(1, 1)
    $yAxis = $chart.Chart.Axes(2, 1)

    # Set the axis titles
    $xAxis.HasTitle = $true
    $xAxis.AxisTitle.Text = $counterNames[0]
    $yAxis.HasTitle = $true
    $yAxis.AxisTitle.Text = $counterNames[$i]

    # Save the Excel book
    $outPath = "$outDirPath$outFileName.xlsx"
    $workbook.SaveAs($outPath)

    $workbook.Close()
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($yAxis) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xAxis) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($series) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($seriesCollection) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($chart) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Perform garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}
