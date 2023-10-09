# Convert the time format MM/dd/yyyy HH:mm:ss.fff into yyyy/MM/dd HH:mm:ss.fff

# Constants
$originalTimeFormat = "MM/dd/yyyy HH:mm:ss.fff"
$targetTimeFormat = "yyyy/MM/dd HH:mm:ss.fff"

# Function
function Convert-Time-Format {
    param (
        [string]$timeColName,
        $csv
    )
    
    # Extract data after the start time
    $csv = $csv | ForEach-Object {
        $time = [DateTime]::ParseExact($_.($timeColName), $originalTimeFormat, $null)
        $_.($timeColName) = $time.ToString($targetTimeFormat)
        $_
    }
    
    return $csv
}
