# Tools orchestrator for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load all tool dependencies
. (Join-Path $PSScriptRoot "Chocolatey.ps1")
. (Join-Path $PSScriptRoot "WindowsADK.ps1")
. (Join-Path $PSScriptRoot "OSCDIMG.ps1")
. (Join-Path $PSScriptRoot "DISM.ps1")

function Test-RequiredTools {
    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"
    
    # Always install Windows ADK via Chocolatey (it won't reinstall if already present)
    Install-WindowsADK
    
    # Find oscdimg.exe path
    $script:oscdimgPath = Find-OscdimgPath
    
    # Ensure DISM is available
    Test-DismAvailability
}
