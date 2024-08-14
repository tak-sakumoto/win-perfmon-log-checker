# Extract data of the CSV file between the start time and the end time

# Constants
$argTimeFormat = "yyyy/MM/dd HH:mm:ss"

# Function
function Edit-Time-Range {
    param (
        [string]$startTime,
        [string]$endTime,
        [string]$timeColName,
        $csv
    )
    
    # Extract data after the start time
    if ($startTime.Trim().Length -ne 0) {
        $startTime = [DateTime]::ParseExact($startTime, $argTimeFormat, $null)
        $csv = $csv | Where-Object { $_.($timeColName) -ge $startTime }
    }

    # Extract data before the end time
    if ($endTime.Trim().Length -ne 0) {
        $endTime = [DateTime]::ParseExact($endTime, $argTimeFormat, $null)
        $csv = $csv | Where-Object { $_.($timeColName) -le $endTime }
    }
    
    return $csv
}