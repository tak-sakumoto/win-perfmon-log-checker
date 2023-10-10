# If the filename contains invalid characters, replace them with valid ones

# Constants
# Search for invalid characters in file paths, excluding backslashes
$invalidChars = [IO.Path]::GetInvalidFileNameChars() | Where-Object { $_ -ne "\\" }
$targetChar = "_"

# Function
function Get-Valid-Name {
    param (
        [string]$fileName
    )

    foreach($char in $invalidChars) {
        $escapedChar = [Regex]::Escape($char)
        if($counterName -match $escapedChar) {
            $fileName = $outFileName -replace $escapedChar,$targetChar
        }
    }
    return $fileName
}
