# Chocolatey package manager for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

function Test-Chocolatey {
    return (Get-Command "choco" -ErrorAction SilentlyContinue) -ne $null
}

function Install-Chocolatey {
    if (Test-Chocolatey) {
        Write-ColorOutput "[OK] Chocolatey already installed!" "Green"
        return
    }

    Write-ColorOutput "Installing Chocolatey package manager..." "Yellow"
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    
    Start-Sleep -Seconds 1

    if (Test-Chocolatey) {
        Write-ColorOutput "[OK] Chocolatey installed successfully!" "Green"
        return
    } else {
        throw "Chocolatey installation failed to become available on PATH."            
    }
}
