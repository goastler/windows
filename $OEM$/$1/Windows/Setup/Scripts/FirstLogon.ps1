# FirstLogon.ps1 - Windows Setup Script
# This script downloads setup.ps1 from GitHub, creates a scheduled task, and runs it

# Set execution policy to allow script execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

# Define paths
$scriptPath = "C:\Windows\Setup\Scripts\setup.ps1"
$scriptDir = Split-Path -Parent $scriptPath

# Create directory if it doesn't exist
if (!(Test-Path $scriptDir)) {
    New-Item -ItemType Directory -Path $scriptDir -Force
}

# Download setup.ps1 from GitHub repository
Write-Host "Downloading setup.ps1 from GitHub repository..."
try {
    $url = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/setup.ps1"
    Invoke-WebRequest -Uri $url -OutFile $scriptPath -UseBasicParsing
    Write-Host "Successfully downloaded setup.ps1 to $scriptPath"
} catch {
    Write-Error "Failed to download setup.ps1: $($_.Exception.Message)"
}

# Run the downloaded setup.ps1 script in a new window
Write-Host "Running setup.ps1 script in a new window..."
try {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Normal
} catch {
    Write-Error "Failed to run setup.ps1: $($_.Exception.Message)"
}

Pause