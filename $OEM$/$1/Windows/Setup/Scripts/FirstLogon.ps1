# FirstLogon.ps1 - Windows Setup Script
# This script downloads onLogin.ps1 from GitHub, creates a scheduled task, and runs it

# Set execution policy to allow script execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

# Define paths
$scriptPath = "C:\Windows\Setup\Scripts\onLogin.ps1"
$scriptDir = Split-Path -Parent $scriptPath

# Create directory if it doesn't exist
if (!(Test-Path $scriptDir)) {
    New-Item -ItemType Directory -Path $scriptDir -Force
}

# Download onLogin.ps1 from GitHub repository
Write-Host "Downloading onLogin.ps1 from GitHub repository..."
try {
    $url = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/onLogin.ps1"
    Invoke-WebRequest -Uri $url -OutFile $scriptPath -UseBasicParsing
    Write-Host "Successfully downloaded onLogin.ps1 to $scriptPath"
} catch {
    Write-Error "Failed to download onLogin.ps1: $($_.Exception.Message)"
}

# Create scheduled task to run onLogin.ps1 on user logon
Write-Host "Creating scheduled task for onLogin.ps1..."
try {
    # Remove existing task if it exists
    $taskName = "OnLoginScript"
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }

    # Create new scheduled task
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId "S-1-5-18" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Runs onLogin.ps1 script on user logon"
    Write-Host "Successfully created scheduled task '$taskName'"
} catch {
    Write-Error "Failed to create scheduled task: $($_.Exception.Message)"
}

# Trigger the scheduled task
Write-Host "Triggering scheduled task 'OnLoginScript'..."
try {
    Start-ScheduledTask -TaskName "OnLoginScript"
    Write-Host "Successfully triggered scheduled task 'OnLoginScript'"
} catch {
    Write-Error "Failed to trigger scheduled task: $($_.Exception.Message)"
}

Write-Host "FirstLogon.ps1 completed. Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
