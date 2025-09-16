# OSCDIMG tool management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load ADK path utilities
. (Join-Path $PSScriptRoot "AdkPaths.ps1")

function Find-OscdimgPath {
    $adkPaths = Get-AdkExecutablePaths
    
    foreach ($path in $adkPaths) {
        if (Test-Path $path) {
            Write-ColorOutput "Found oscdimg.exe at: $path" -Color "Green" -Indent 1
            return $path
        }
    }
    
    throw "oscdimg.exe not found. Please ensure Windows ADK is properly installed."
}
