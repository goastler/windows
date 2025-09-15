# OSCDIMG tool management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

function Find-OscdimgPath {
    $adkPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
    )
    
    foreach ($path in $adkPaths) {
        if (Test-Path $path) {
            Write-ColorOutput "Found oscdimg.exe at: $path" -Color "Green" -Indent 1
            return $path
        }
    }
    
    throw "oscdimg.exe not found. Please ensure Windows ADK is properly installed."
}
