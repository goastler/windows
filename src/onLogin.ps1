
# Create log file
$logFile = "$env:TEMP\onLogin.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Set error action preference for the entire script
$ErrorActionPreference = "Stop"

function Write-Log {
    param($Message)
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $logFile -Value $logMessage
}



try {
    Write-Log "Starting on login setup..."

    # =============================================================================
    # USER CANCELLATION DELAY
    # =============================================================================

    Write-Host "`n=== On Login Setup Starting ===" -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to cancel..." -ForegroundColor Yellow
    Write-Host ""
    Start-Sleep -Seconds 30
    Write-Host "Starting now..." -ForegroundColor Green

    # =============================================================================
    # CHOCOLATEY INSTALLATION AND PACKAGE SETUP
    # =============================================================================

    Write-Log "Installing Chocolatey and packages..."

    # Check if Chocolatey is already installed
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Log "Chocolatey is already installed. Updating..."
        $chocoUpgradeResult = choco upgrade chocolatey -y
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to upgrade Chocolatey. Exit code: $LASTEXITCODE"
        }
    } else {
        Write-Log "Installing Chocolatey..."
        # Install Chocolatey
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        
        Write-Log "Chocolatey installed successfully"
    }

    # Wait a moment for Chocolatey to be fully available
    Start-Sleep -Seconds 5

    # Clear any existing Chocolatey downloads to prevent corrupted package issues
    Write-Log "Clearing existing Chocolatey downloads..."
    $clearDownloadsResult = choco cache remove --expired -y
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to clear Chocolatey downloads. Exit code: $LASTEXITCODE"
    }
    Write-Log "Chocolatey downloads cleared successfully"

    # Define common packages to install
    $packages = @(
        # Web Browsers
        "googlechrome",
        "firefox",
        "microsoft-edge"
    )

    Write-Log "Installing common packages..."

    # Install packages
    foreach ($package in $packages) {
        Write-Log "Installing $package..."
        $installResult = choco install $package -y
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to install package '$package'. Exit code: $LASTEXITCODE"
        }
        Write-Log "$package installed successfully"
    }

    # Update all packages
    Write-Log "Updating all installed packages..."
    $upgradeResult = choco upgrade all -y
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to upgrade packages. Exit code: $LASTEXITCODE"
    }
    Write-Log "Package updates completed"

    # Clean up
    Write-Log "Cleaning up Chocolatey cache..."
    $cacheResult = choco cache remove --expired -y
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to clean up Chocolatey cache. Exit code: $LASTEXITCODE"
    }

    Write-Log "Chocolatey installation and package setup completed!"

    # =============================================================================
    # WINDOWS UPDATE CONFIGURATION
    # =============================================================================

    Write-Log "Configuring Windows Update settings..."

    # Configure Windows Update registry settings
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Set automatic update configuration
    Set-ItemProperty -Path $regPath -Name "AUOptions" -Value 4 -Type DWord  # Auto download and install
    Set-ItemProperty -Path $regPath -Name "NoAutoRebootWithLoggedOnUsers" -Value 0 -Type DWord  # Allow auto-reboot
    Set-ItemProperty -Path $regPath -Name "NoAutoUpdate" -Value 0 -Type DWord  # Enable auto updates
    Set-ItemProperty -Path $regPath -Name "ScheduledInstallDay" -Value 0 -Type DWord  # Every day
    Set-ItemProperty -Path $regPath -Name "ScheduledInstallTime" -Value 3 -Type DWord  # 3 AM

    # Configure Windows Update to include optional updates
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "IncludeRecommendedUpdates" -Value 1 -Type DWord

    # Configure Windows Update for Business (if applicable)
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "DeferUpgrade" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "DeferUpgradePeriod" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "ExcludeWUDriversInQualityUpdate" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "DisableWindowsUpdateAccess" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "DisableWindowsUpdateAccessAsUser" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "SetDisableUXWUAccess" -Value 0 -Type DWord
    Set-ItemProperty -Path $regPath -Name "SetDisableUXWUAccessAsUser" -Value 0 -Type DWord

    Write-Log "Windows Update registry settings configured"

    # =============================================================================
    # CHECK FOR WINDOWS UPDATES AND CLEANUP
    # =============================================================================

    Write-Log "Checking for available Windows updates..."

    # Create Windows Update session
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    # Search for all updates including optional
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' or IsInstalled=0 and Type='Driver'")

    if ($SearchResult.Updates.Count -eq 0) {
        Write-Log "No Windows updates available. System is up to date."
    } else {
        Write-Log "Found $($SearchResult.Updates.Count) Windows updates available. Installing updates..."
    }

    # List available updates
    foreach ($Update in $SearchResult.Updates) {
        Write-Log "Available update: $($Update.Title)"
    }
    
    # Shuffle the order of updates (only if updates are available)
    $UpdatesArray = @($SearchResult.Updates)
    if ($UpdatesArray.Count -gt 0) {
        $ShuffledUpdates = $UpdatesArray | Get-Random -Count $UpdatesArray.Count
    } else {
        $ShuffledUpdates = @()
    }
    
    # Install the updates one by one
    $rebootRequired = $false
    
    foreach ($Update in $ShuffledUpdates) {
        Write-Log "Processing update: $($Update.Title)"
        
        # Create update collection for single update
        $SingleUpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        $SingleUpdateCollection.Add($Update) | Out-Null
        
        # Download single update
        Write-Log "Downloading update: $($Update.Title)"
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $SingleUpdateCollection
        $DownloadResult = $Downloader.Download()
        
        if ($DownloadResult.ResultCode -eq 2) {
            Write-Log "Update downloaded successfully: $($Update.Title)"
            
            # Install single update
            Write-Log "Installing update: $($Update.Title)"
            $Installer = $UpdateSession.CreateUpdateInstaller()
            $Installer.Updates = $SingleUpdateCollection
            $InstallResult = $Installer.Install()
            
            if ($InstallResult.ResultCode -eq 2) {
                Write-Log "Update installed successfully: $($Update.Title)"
                if ($InstallResult.RebootRequired) {
                    $rebootRequired = $true
                    Write-Log "Reboot required after: $($Update.Title)"
                }
            } else {
                $errorMsg = "Failed to install update: $($Update.Title). Result code: $($InstallResult.ResultCode)"
                Write-Log $errorMsg
                Write-Log "Continuing with next update..."
            }
        } else {
            $errorMsg = "Failed to download update: $($Update.Title). Result code: $($DownloadResult.ResultCode)"
            Write-Log $errorMsg
            Write-Log "Continuing with next update..."
        }
    }
    
    Write-Log "All updates processed"
    if ($rebootRequired) {
        Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
        Write-Host "Windows updates require a system reboot." -ForegroundColor Yellow
        Write-Host "The computer will restart in 30 seconds..." -ForegroundColor Yellow
        Write-Host "Press Ctrl+C to cancel the reboot" -ForegroundColor Red
        Write-Host ""
        Start-Sleep -Seconds 30
        Write-Log "Reboot required. Restarting computer..."
        Restart-Computer -Force
    } else {
        Write-Log "No reboot required. All updates processed successfully."
    }

    # =============================================================================
    # OFFICE INSTALLATION
    # =============================================================================

    Write-Log "Checking if Office is already installed..."

    # Check if Office is already installed by looking for Office applications
    $officeInstalled = $false
    $officeApps = @("winword.exe", "excel.exe", "powerpnt.exe", "outlook.exe")
    
    foreach ($app in $officeApps) {
        $appPath = Get-Command $app -ErrorAction SilentlyContinue
        if ($appPath) {
            $officeInstalled = $true
            Write-Log "Office is already installed (found $app at $($appPath.Source))"
            break
        }
    }

    if ($officeInstalled) {
        Write-Log "Office is already installed. Skipping Office installation process."
    } else {
        Write-Log "Office not found. Starting Office installation process..."

        # Create temporary directory for Office installation
        $officeTempDir = "$env:TEMP\OfficeInstall"
        Write-Log "Creating temporary directory: $officeTempDir"
        if (Test-Path $officeTempDir) {
            Remove-Item $officeTempDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $officeTempDir -Force | Out-Null
        Set-Location $officeTempDir
        Write-Log "Changed to directory: $officeTempDir"

        # Download Office deployment tool
        $officeDownloadUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20136.exe"
        $officeInstaller = "officedeploymenttool.exe"
        
        Write-Log "Downloading Office deployment tool from: $officeDownloadUrl"
        Invoke-WebRequest -Uri $officeDownloadUrl -OutFile $officeInstaller -UseBasicParsing
        Write-Log "Office deployment tool downloaded successfully"

        # Run the Office deployment tool to extract files
        Write-Log "Extracting Office deployment tool files..."
        & ".\$officeInstaller" /quiet /extract:$officeTempDir
        Write-Log "Office deployment tool extracted successfully"

        # Download office.xml from GitHub repository
        $officeXmlUrl = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/office.xml"
        $officeXmlDest = "$officeTempDir\office.xml"
        
        Write-Log "Downloading office.xml from GitHub repository..."
        Invoke-WebRequest -Uri $officeXmlUrl -OutFile $officeXmlDest -UseBasicParsing
        Write-Log "office.xml downloaded successfully"

        # Run Office setup with the configuration file
        Write-Log "Starting Office installation with configuration file..."
        $setupProcess = Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure", "office.xml" -Wait -PassThru -NoNewWindow
        if ($setupProcess.ExitCode -ne 0) {
            throw "Office installation completed with exit code: $($setupProcess.ExitCode)"
        }
        Write-Log "Office installation completed successfully"

        # Clean up temporary directory
        Write-Log "Cleaning up temporary Office installation directory..."
        Set-Location $env:TEMP
        Remove-Item $officeTempDir -Recurse -Force
        Write-Log "Temporary directory cleaned up successfully"

        Write-Log "Office installation process completed!"
    }



    # =============================================================================
    # FINAL CLEANUP - REMOVE SCHEDULED TASK
    # =============================================================================

    Write-Log "Setup completed successfully. Removing setup scheduled task..."
    Unregister-ScheduledTask -TaskName "OnLoginSetup" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Setup scheduled task removed successfully"

    # =============================================================================
    # COMPLETION
    # =============================================================================

    Write-Log "On login setup completed. Log saved to: $logFile"

} catch {
    Write-Host "`n=== ERROR ===" -ForegroundColor Red
    Write-Host "An error occurred during setup: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception)" -ForegroundColor Red
    Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
    Write-Host "Due to the error, the system will reboot in 60 seconds..." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to cancel the reboot" -ForegroundColor Red
    Write-Host ""
    Start-Sleep -Seconds 30
    Write-Host "Rebooting now..." -ForegroundColor Red
    Restart-Computer -Force
}
