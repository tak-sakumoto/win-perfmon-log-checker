# Draw a graph for a counter and save as an Excel workbook

# Constants
$chartHeight = 300
$chartWidth = 400
$chartType = 4 # 4: line graph

# Function
function Save-Graph-By-Excel {
    param (
        [string]$outPath,
        [string]$xAxisName,
        [string]$yAxisName,
        $csv
    )

    # Prepare to handle Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Prepare a workbook with a sheet
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)

    # Set column names to the 1st row of the workbook
    $worksheet.Cells.Item(1, 1) = $xAxisName
    $worksheet.Cells.Item(1, 2) = $yAxisName

    # Set data after the 2nd row of the workbook
    $i = 2
    foreach ($line in $csv) {
        $worksheet.Cells.Item($i, 1) = $line.($xAxisName)
        $worksheet.Cells.Item($i, 2) = $line.($yAxisName)
        $i++
    }

    # Add a chart in the workbook
    $chart = $worksheet.Shapes.AddChart2(-1, $chartType, 0, 0, $chartWidth, $chartHeight, $true)

    # Specify a range of data in the workbook
    $seriesCollection = $chart.Chart.SeriesCollection()
    $series = $seriesCollection.NewSeries()
    $series.XValues = $worksheet.Range($worksheet.Cells.Item(2, 1), $worksheet.Cells.Item($i, 1))
    $series.Values = $worksheet.Range($worksheet.Cells.Item(2, 2), $worksheet.Cells.Item($i, 2))

    # Get the axes
    $xAxis = $chart.Chart.Axes(1, 1)
    $yAxis = $chart.Chart.Axes(2, 1)

    # Set the axis titles
    $xAxis.HasTitle = $true
    $xAxis.AxisTitle.Text = $worksheet.Cells.Item(1, 1)
    $yAxis.HasTitle = $true
    $yAxis.AxisTitle.Text = $worksheet.Cells.Item(1, 2)

    # Save the Excel book
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