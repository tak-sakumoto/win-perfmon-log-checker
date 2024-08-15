# Get data statistics and write them to a file

# Function
function Get-Stats {
    param (
        [string]$outPath,
        [string]$colName,
        $csv
    )

    # Get values in the i-th column
    $columnValues = $csv | ForEach-Object { $_.($colName) }

    # Get an array of stats for the i-th column
    $stats = @{
        Count = $columnValues.Count
        Maximum = ($columnValues | Measure-Object -Maximum).Maximum
        Minimum = ($columnValues | Measure-Object -Minimum).Minimum
        Average = ($columnValues | Measure-Object -Average).Average
    }

    # Export the stats array
    $stats | Select-Object Count, Maximum, Minimum, Average | Export-Csv -Path $outPath -NoTypeInformation -UseQuotes AsNeeded
}
