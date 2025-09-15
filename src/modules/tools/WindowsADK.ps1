# Windows ADK management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load Chocolatey module dependency
. (Join-Path $PSScriptRoot "Chocolatey.ps1")

function Install-WindowsADK {
    Write-ColorOutput "=== Windows ADK Installation ===" "Cyan"

    Install-Chocolatey

    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"

    $result = Start-Process -FilePath "choco" -ArgumentList @("install","windows-adk","-y") -Wait -PassThru -NoNewWindow
    if ($result.ExitCode -eq 0) {
        Write-ColorOutput "[OK] Windows ADK installed successfully via Chocolatey!" "Green"
        
        # Add ADK paths to current session PATH
        $adkPaths = @(
            "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
            "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
            "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg",
            "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg"
        )
        
        foreach ($path in $adkPaths) {
            if (Test-Path $path) {
                $env:Path += ";$path"
                Write-ColorOutput "Added to PATH: $path" "Cyan" -Indent 1
            }
        }
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        return
    }

    throw "Windows ADK installation failed"
}
