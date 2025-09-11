
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

function Invoke-CommandWithExitCode {
    param(
        [string]$Command,
        [string]$Description = "",
        [int]$ExpectedExitCode = 0
    )
    
    Write-Log "Executing: $Command"
    
    # Execute command and stream output in real-time
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "powershell.exe"
    $processInfo.Arguments = "-Command `"$Command`""
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    
    # Start the process
    $process.Start() | Out-Null
    
    # Stream output in real-time and collect separately
    $stdoutOutput = @()
    $stderrOutput = @()
    while ($true) {
        $hasOutput = $false
        
        # Read stdout
        if (!$process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($line) {
                Write-Host $line -ForegroundColor Green
                $stdoutOutput += $line
                $hasOutput = $true
            }
        }
        
        # Read stderr
        if (!$process.StandardError.EndOfStream) {
            $line = $process.StandardError.ReadLine()
            if ($line) {
                Write-Host $line -ForegroundColor Red
                $stderrOutput += $line
                $hasOutput = $true
            }
        }
        
        # If process has exited and no more output, break
        if ($process.HasExited -and !$hasOutput) {
            break
        }
        
        # If no output this iteration, sleep briefly
        if (!$hasOutput) {
            Start-Sleep -Milliseconds 50
        }
    }
    
    # Wait for process to exit and get exit code
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    
    if ($exitCode -ne $ExpectedExitCode) {
        $errorMsg = if ($Description) { 
            "Failed to $Description. Command: $Command. Exit code: $exitCode (expected: $ExpectedExitCode)" 
        } else { 
            "Command failed: $Command. Exit code: $exitCode (expected: $ExpectedExitCode)" 
        }
        throw $errorMsg
    }
    
    Write-Log "Command completed successfully: $Command"
    
    # Return stdout, stderr, and exit code
    return @{
        StdOut = $stdoutOutput
        StdErr = $stderrOutput
        ExitCode = $exitCode
    }
}

function Invoke-WebRequestWithCleanup {
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Description = "download file"
    )
    
    Write-Log "Downloading $Description from: $Uri"
    
    $webRequest = Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing
    Write-Log "$Description downloaded successfully"
    
    # Ensure proper cleanup of web request object and file handles
    $webRequest = $null
    [System.GC]::Collect()
    Start-Sleep -Seconds 2  # Brief pause to ensure file handles are released
}

function Wait-ForUserCancellation {
    param(
        [int]$Seconds = 30,
        [string]$Message = "Press any key to cancel..."
    )
    
    $timeout = $Seconds
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()

    Write-Host "$Message" -ForegroundColor Yellow

    while ($stopwatch.Elapsed.TotalSeconds -lt $timeout) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)  # $true prevents echo to screen
            Write-Host "`nYou pressed: $($key.Key)" -ForegroundColor Red
            Write-Host "User cancelled. Exiting script..." -ForegroundColor Red
            exit 1
        }
        Start-Sleep -Milliseconds 100
    }

    if (-not [Console]::KeyAvailable -and $stopwatch.Elapsed.TotalSeconds -ge $timeout) {
        Write-Host "No key was pressed within $timeout seconds." -ForegroundColor Green
    }
}

function Install-ChocolateyAndPackages {
    Write-Log "Installing Chocolatey and packages..."

    # Check if Chocolatey is already installed
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Log "Chocolatey is already installed. Updating..."
        $result = Invoke-CommandWithExitCode -Command "choco upgrade chocolatey -y" -Description "upgrade Chocolatey"
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
    $result = Invoke-CommandWithExitCode -Command "choco cache remove --expired -y" -Description "clear Chocolatey downloads"

    # Define common packages to install
    $packages = @(
        # Web Browsers
        "googlechrome",
        "firefox",
        "microsoft-edge"
    )

    Write-Log "Installing common packages..."

    # Install all packages in one command
    $packageList = $packages -join " "
    $result = Invoke-CommandWithExitCode -Command "choco install $packageList -y --ignore-checksums" -Description "install all packages: $packageList"

    # Update all packages
    $result = Invoke-CommandWithExitCode -Command "choco upgrade all -y --ignore-checksums" -Description "upgrade all packages"

    # Clean up
    $result = Invoke-CommandWithExitCode -Command "choco cache remove --expired -y" -Description "clean up Chocolatey cache"

    Write-Log "Chocolatey installation and package setup completed!"
}

function Configure-WindowsUpdates {
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
}

function Install-WindowsUpdates {
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
        Wait-ForUserCancellation -Seconds 30 -Message "Press any key to cancel the reboot"
        Write-Log "Reboot required. Restarting computer..."
        Restart-Computer -Force
    } else {
        Write-Log "No reboot required. All updates processed successfully."
    }
}

function Install-Office {
    Write-Log "Starting Office installation process..."

    # Create temporary directory for Office installation
    $officeTempDir = "$env:TEMP\OfficeInstall"
    $originalLocation = Get-Location
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
    
    Invoke-WebRequestWithCleanup -Uri $officeDownloadUrl -OutFile $officeInstaller -Description "Office deployment tool"

    # Run the Office deployment tool to extract files
    Write-Log "Extracting Office deployment tool files..."
    & ".\$officeInstaller" /quiet /extract:$officeTempDir
    Write-Log "Office deployment tool extracted successfully"

    # Download office.xml from GitHub repository
    $officeXmlUrl = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/office.xml"
    $officeXmlDest = "$officeTempDir\office.xml"
    
    Invoke-WebRequestWithCleanup -Uri $officeXmlUrl -OutFile $officeXmlDest -Description "office.xml configuration file"

    # Run Office setup with the configuration file
    Write-Log "Starting Office installation with configuration file..."
    $setupProcess = Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure", "office.xml" -Wait -PassThru -NoNewWindow
    if ($setupProcess.ExitCode -ne 0) {
        throw "Office installation completed with exit code: $($setupProcess.ExitCode)"
    }
    Write-Log "Office installation completed successfully"

    # Clean up temporary directory
    Write-Log "Cleaning up temporary Office installation directory..."
    Set-Location $originalLocation
    Remove-Item $officeTempDir -Recurse -Force
    Write-Log "Temporary directory cleaned up successfully"

    Write-Log "Office installation process completed!"
}

function Setup-BgInfo {
    Write-Log "Setting up BgInfo..."

    # Create BgInfo directory
    $bgInfoDir = "C:\Tools\BgInfo"
    if (!(Test-Path $bgInfoDir)) {
        New-Item -ItemType Directory -Path $bgInfoDir -Force | Out-Null
        Write-Log "Created BgInfo directory: $bgInfoDir"
    }

    # Download BgInfo
    $bgInfoUrl = "https://download.sysinternals.com/files/BGInfo.zip"
    $bgInfoZip = "$env:TEMP\BgInfo.zip"
    $bgInfoExe = "$bgInfoDir\BgInfo.exe"
    
    if (!(Test-Path $bgInfoExe)) {
        Write-Log "Downloading BgInfo..."
        Invoke-WebRequestWithCleanup -Uri $bgInfoUrl -OutFile $bgInfoZip -Description "BgInfo"
        
        # Extract BgInfo
        Write-Log "Extracting BgInfo..."
        Expand-Archive -Path $bgInfoZip -DestinationPath $bgInfoDir -Force
        Write-Log "BgInfo extracted successfully"
        
        # Clean up zip file
        Remove-Item $bgInfoZip -Force
        Write-Log "BgInfo zip file cleaned up"
    } else {
        Write-Log "BgInfo is already installed at: $bgInfoExe"
    }

    # Download BgInfo configuration file from GitHub repository
    $bgInfoConfigUrl = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/BgInfo.bgi"
    $bgInfoConfigFile = "$bgInfoDir\BgInfo.bgi"
    
    Write-Log "Downloading BgInfo configuration file from GitHub..."
    Invoke-WebRequestWithCleanup -Uri $bgInfoConfigUrl -OutFile $bgInfoConfigFile -Description "BgInfo configuration file"
    Write-Log "BgInfo configuration file downloaded: $bgInfoConfigFile"

    # Create BgInfo startup task
    $taskName = "BgInfo"
    $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    
    if (!$taskExists) {
        Write-Log "Creating BgInfo startup task..."
        
        # Create the action
        $action = New-ScheduledTaskAction -Execute $bgInfoExe -Argument "/nolicprompt /timer:0 /all"
        
        # Create the trigger (at startup)
        $trigger = New-ScheduledTaskTrigger -AtStartup
        
        # Create the principal (run as SYSTEM)
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        
        # Create the settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false
        
        # Register the task
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "BgInfo - Display system information on desktop background"
        
        Write-Log "BgInfo startup task created successfully"
    } else {
        Write-Log "BgInfo startup task already exists"
    }

    # Run BgInfo immediately to apply the configuration
    Write-Log "Running BgInfo to apply configuration..."
    $result = Invoke-CommandWithExitCode -Command "& '$bgInfoExe' /nolicprompt /timer:0 /silent /accepteula" -Description "run BgInfo to apply configuration"

    Write-Log "BgInfo setup completed!"
}

function Install-MicrosoftActivationScripts {
    Write-Log "Setting up Microsoft Activation Scripts..."

    # Create temporary directory for MAS
    $masTempDir = "$env:TEMP\MAS"
    $originalLocation = Get-Location
    Write-Log "Creating temporary directory: $masTempDir"
    if (Test-Path $masTempDir) {
        Remove-Item $masTempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $masTempDir -Force | Out-Null
    Set-Location $masTempDir
    Write-Log "Changed to directory: $masTempDir"

    # Download Microsoft Activation Scripts
    $masUrl = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd"
    $masFile = "MAS_AIO.cmd"
    
    Write-Log "Downloading Microsoft Activation Scripts..."
    Invoke-WebRequestWithCleanup -Uri $masUrl -OutFile $masFile -Description "Microsoft Activation Scripts"

    # Run MAS with /HWID parameter
    Write-Log "Running Microsoft Activation Scripts with /HWID parameter..."
    $result = Invoke-CommandWithExitCode -Command ".\$masFile /HWID" -Description "run MAS with /HWID parameter"

    # Run MAS with /Ohook parameter
    Write-Log "Running Microsoft Activation Scripts with /Ohook parameter..."
    $result = Invoke-CommandWithExitCode -Command ".\$masFile /Ohook" -Description "run MAS with /Ohook parameter"

    # Clean up temporary directory
    Write-Log "Cleaning up temporary MAS directory..."
    Set-Location $originalLocation
    Remove-Item $masTempDir -Recurse -Force
    Write-Log "Temporary directory cleaned up successfully"

    Write-Log "Microsoft Activation Scripts setup completed!"
}

try {
    Write-Log "Starting on login setup..."

    # =============================================================================
    # USER CANCELLATION DELAY
    # =============================================================================

    Write-Host "`n=== On Login Setup Starting ===" -ForegroundColor Yellow
    Wait-ForUserCancellation -Seconds 30 -Message "Press any key to cancel..."
    Write-Host "Starting now..." -ForegroundColor Green

    # =============================================================================
    # WINDOWS UPDATE CONFIGURATION AND INSTALLATION
    # =============================================================================

    Configure-WindowsUpdates
    Install-WindowsUpdates

    # =============================================================================
    # CHOCOLATEY INSTALLATION AND PACKAGE SETUP
    # =============================================================================

    Install-ChocolateyAndPackages

    # =============================================================================
    # BGINFO SETUP
    # =============================================================================

    Setup-BgInfo

    # =============================================================================
    # OFFICE INSTALLATION
    # =============================================================================

    Install-Office

    # =============================================================================
    # MICROSOFT ACTIVATION SCRIPTS
    # =============================================================================

    Install-MicrosoftActivationScripts

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
}

Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
Write-Host "The system will reboot in 30 seconds..." -ForegroundColor Yellow
Wait-ForUserCancellation -Seconds 30 -Message "Press any key to cancel the reboot"
Write-Host "Rebooting now..." -ForegroundColor Red
Restart-Computer -Force
