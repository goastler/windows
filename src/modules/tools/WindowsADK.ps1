# Windows ADK management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load Chocolatey module dependency
. (Join-Path $PSScriptRoot "Chocolatey.ps1")

# Load ADK path utilities
. (Join-Path $PSScriptRoot "AdkPaths.ps1")

function Install-WindowsADK {
    Write-Host ""
    Write-ColorOutput "=== Windows ADK Installation ===" -Color "Cyan"

    Install-Chocolatey

    Write-ColorOutput "Installing Windows ADK via Chocolatey..." -Color "Yellow"

    & choco install windows-adk -y
    if ($LASTEXITCODE -eq 0) {
        Write-ColorOutput "Windows ADK installed via Chocolatey" -Color "Green"
        
        # Add ADK paths to current session PATH
        $adkPaths = Get-AdkPaths
        
        foreach ($path in $adkPaths) {
            if (Test-Path $path) {
                $env:Path += ";$path"
                Write-ColorOutput "Added to PATH: $path" -Color "Cyan" -Indent 1
            }
        }
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        return
    }

    throw "Windows ADK installation failed"
}
