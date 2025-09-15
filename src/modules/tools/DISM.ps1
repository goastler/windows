# DISM tool management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load Chocolatey module dependency
. (Join-Path $PSScriptRoot "Chocolatey.ps1")

function Test-DismAvailability {
    Write-ColorOutput "Checking DISM availability..." "Yellow"
    
    # Check if DISM is available in PATH
    $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
    if ($dismCommand) {
        Write-ColorOutput "DISM found at: $($dismCommand.Source)" "Green" -Indent 1
        return
    }
    
    # DISM should be available on Windows 7+ by default, but let's check common locations
    $dismPaths = @(
        "${env:SystemRoot}\System32\dism.exe",
        "${env:SystemRoot}\SysWOW64\dism.exe"
    )
    
    foreach ($path in $dismPaths) {
        if (Test-Path $path) {
            Write-ColorOutput "DISM found at: $path" "Green" -Indent 1
            # Add to PATH for current session if not already there
            $dismDir = Split-Path $path -Parent
            if ($env:Path -notlike "*$dismDir*") {
                $env:Path += ";$dismDir"
                Write-ColorOutput "Added DISM directory to PATH: $dismDir" "Cyan" -Indent 2
            }
            return
        }
    }
    
    # If DISM is not found, try to install it via Windows Features
    Write-ColorOutput "DISM not found in standard locations. Attempting to enable via Windows Features..." "Yellow"
    try {
        # Try to enable DISM via DISM itself (ironic but sometimes works)
        $result = Start-Process -FilePath "dism.exe" -ArgumentList "/?" -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "DISM is now available" "Green" -Indent 1
            return
        }
    } catch {
        # DISM is not available, try alternative approaches
    }
    
    # Try to enable via PowerShell
    try {
        Write-ColorOutput "Attempting to enable DISM via PowerShell..." "Yellow"
        Enable-WindowsOptionalFeature -Online -FeatureName "Deployment-Tools-Foundation" -NoRestart -ErrorAction SilentlyContinue | Out-Null
        Enable-WindowsOptionalFeature -Online -FeatureName "Deployment-Tools-Foundation-FoD" -NoRestart -ErrorAction SilentlyContinue | Out-Null
        
        # Check again
        $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
        if ($dismCommand) {
            Write-ColorOutput "DISM enabled successfully at: $($dismCommand.Source)" "Green" -Indent 1
            return
        }
    } catch {
        Write-ColorOutput "Failed to enable DISM via PowerShell: $($_.Exception.Message)" "Yellow"
    }
    
    # Last resort: try to install via Chocolatey
    try {
        Write-ColorOutput "Attempting to install DISM via Chocolatey..." "Yellow"
        Install-Chocolatey
        
        $result = Start-Process -FilePath "choco" -ArgumentList @("install", "windows-adk-deployment-tools", "-y") -Wait -PassThru -NoNewWindow
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Windows ADK Deployment Tools installed via Chocolatey" "Green"
            
            # Refresh PATH
            $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                        [System.Environment]::GetEnvironmentVariable('Path','User')
            
            # Check again
            $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
            if ($dismCommand) {
                Write-ColorOutput "DISM now available at: $($dismCommand.Source)" "Green" -Indent 1
                return
            }
        }
    } catch {
        Write-ColorOutput "Failed to install DISM via Chocolatey: $($_.Exception.Message)" "Yellow"
    }
    
    # If we get here, DISM is not available
    Write-ColorOutput "ERROR: DISM is not available and could not be installed automatically." "Red"
    Write-ColorOutput "DISM is required for VirtIO driver integration." "Red"
    Write-ColorOutput "Please ensure you are running on Windows 7 or later, or install Windows ADK manually." "Red"
    throw "DISM is not available. Required for VirtIO driver integration."
}

function Get-DismPath {
    # Try to get DISM from PATH first
    $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
    if ($dismCommand) {
        return $dismCommand.Source
    }
    
    # Check common DISM locations
    $dismPaths = @(
        "${env:SystemRoot}\System32\dism.exe",
        "${env:SystemRoot}\SysWOW64\dism.exe"
    )
    
    foreach ($path in $dismPaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    throw "DISM not found. Please ensure DISM is available on the system."
}
