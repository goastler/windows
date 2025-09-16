# DISM tool management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
. $commonPath

# Load Chocolatey module dependency
. (Join-Path $PSScriptRoot "Chocolatey.ps1")

function Test-DismAvailability {
    Write-ColorOutput "Checking DISM availability..." -Color "Yellow"
    
    # Check if DISM is available in PATH
    $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
    if ($dismCommand) {
        $dismSource = $dismCommand.Source
        $dismSource = Assert-ValidPath -VariableName "dismSource" -Path $dismSource -ErrorMessage "DISM command source path is invalid: $dismSource"
        Write-ColorOutput "DISM found at: $dismSource" -Color "Green" -Indent 1
        return
    }
    
    # DISM should be available on Windows 7+ by default, but let's check common locations
    $dismPaths = @(
        "${env:SystemRoot}\System32\dism.exe",
        "${env:SystemRoot}\SysWOW64\dism.exe"
    )
    $dismPaths = Assert-ArrayNotEmpty -VariableName "dismPaths" -Value $dismPaths -ErrorMessage "DISM paths array is empty"
    
    foreach ($path in $dismPaths) {
        $path = Assert-ValidPath -VariableName "path" -Path $path -ErrorMessage "DISM path is invalid: $path"
        
        if (Test-Path $path) {
            Write-ColorOutput "DISM found at: $path" -Color "Green" -Indent 1
            # Add to PATH for current session if not already there
            $dismDir = Split-Path $path -Parent
            $dismDir = Assert-ValidPath -VariableName "dismDir" -Path $dismDir -ErrorMessage "DISM directory path is invalid: $dismDir"
            if ($env:Path -notlike "*$dismDir*") {
                $env:Path += ";$dismDir"
                Write-ColorOutput "Added DISM directory to PATH: $dismDir" -Color "Cyan" -Indent 2
            }
            return
        }
    }
    
    # If DISM is not found, try to install it via Windows Features
    Write-ColorOutput "DISM not found in standard locations. Attempting to enable via Windows Features..." -Color "Yellow"
    try {
        # Try to enable DISM via DISM itself (ironic but sometimes works)
        & dism.exe /? > $null 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "DISM is now available" -Color "Green" -Indent 1
            return
        }
    } catch {
        # DISM is not available, try alternative approaches
    }
    
    # Try to enable via PowerShell
    try {
        Write-ColorOutput "Attempting to enable DISM via PowerShell..." -Color "Yellow"
        Enable-WindowsOptionalFeature -Online -FeatureName "Deployment-Tools-Foundation" -NoRestart -ErrorAction SilentlyContinue | Out-Null
        Enable-WindowsOptionalFeature -Online -FeatureName "Deployment-Tools-Foundation-FoD" -NoRestart -ErrorAction SilentlyContinue | Out-Null
        
        # Check again
        $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
        if ($dismCommand) {
            $dismSource = $dismCommand.Source
            $dismSource = Assert-ValidPath -VariableName "dismSource" -Path $dismSource -ErrorMessage "DISM command source path is invalid: $dismSource"
            Write-ColorOutput "DISM enabled successfully at: $dismSource" -Color "Green" -Indent 1
            return
        }
    } catch {
        Write-ColorOutput "Failed to enable DISM via PowerShell: $($_.Exception.Message)" -Color "Yellow"
    }
    
    # Last resort: try to install via Chocolatey
    try {
        Write-ColorOutput "Attempting to install DISM via Chocolatey..." -Color "Yellow"
        Install-Chocolatey
        
        & choco install windows-adk-deployment-tools -y
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "Windows ADK Deployment Tools installed via Chocolatey" -Color "Green"
            
            # Refresh PATH
            $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                        [System.Environment]::GetEnvironmentVariable('Path','User')
            
            # Check again
            $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
            if ($dismCommand) {
                $dismSource = $dismCommand.Source
                $dismSource = Assert-ValidPath -VariableName "dismSource" -Path $dismSource -ErrorMessage "DISM command source path is invalid: $dismSource"
                Write-ColorOutput "DISM now available at: $dismSource" -Color "Green" -Indent 1
                return
            }
        }
    } catch {
        Write-ColorOutput "Failed to install DISM via Chocolatey: $($_.Exception.Message)" -Color "Yellow"
    }
    
    # If we get here, DISM is not available
    throw "DISM is not available and could not be installed automatically. DISM is required for VirtIO driver integration. Please ensure you are running on Windows 7 or later, or install Windows ADK manually."
}

function Get-DismPath {
    # Try to get DISM from PATH first
    $dismCommand = Get-Command "dism.exe" -ErrorAction SilentlyContinue
    if ($dismCommand) {
        $dismSource = $dismCommand.Source
        $dismSource = Assert-ValidPath -VariableName "dismSource" -Path $dismSource -ErrorMessage "DISM command source path is invalid: $dismSource"
        return $dismSource
    }
    
    # Check common DISM locations
    $dismPaths = @(
        "${env:SystemRoot}\System32\dism.exe",
        "${env:SystemRoot}\SysWOW64\dism.exe"
    )
    $dismPaths = Assert-ArrayNotEmpty -VariableName "dismPaths" -Value $dismPaths -ErrorMessage "DISM paths array is empty"
    
    foreach ($path in $dismPaths) {
        $path = Assert-ValidPath -VariableName "path" -Path $path -ErrorMessage "DISM path is invalid: $path"
        
        if (Test-Path $path) {
            return $path
        }
    }
    
    throw "DISM not found. Please ensure DISM is available on the system."
}
