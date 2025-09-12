
# Create log file
$logFile = "$env:TEMP\onLogin.log"

# Scheduled task name
$scheduledTaskName = "OnLogin"

# Store original directory location for restoration on error
$originalScriptLocation = Get-Location

# Set error action preference for the entire script
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage -ForegroundColor $Color
    Add-Content -Path $logFile -Value $logMessage
}

function Write-Log-Highlight {
    param(
        [string]$Message,
        [string]$HighlightText,
        [string]$Color = "White",
        [string]$HighlightColor = "Yellow"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "[$timestamp] $Message"
    
    # Write to console with highlighting
    $parts = $Message -split $HighlightText
    if ($parts.Count -gt 1) {
        Write-Host "[$timestamp] " -NoNewline -ForegroundColor $Color
        for ($i = 0; $i -lt $parts.Count; $i++) {
            Write-Host $parts[$i] -NoNewline -ForegroundColor $Color
            if ($i -lt $parts.Count - 1) {
                Write-Host $HighlightText -NoNewline -ForegroundColor $HighlightColor
            }
        }
        Write-Host ""
    } else {
        Write-Host $logMessage -ForegroundColor $Color
    }
    
    # Write plain text to log file
    Add-Content -Path $logFile -Value $logMessage
}

function Invoke-CommandWithExitCode {
    param(
        [string]$Command,
        [string]$Description = "",
        [int]$ExpectedExitCode = 0
    )
    
    Write-Log "Executing: $Command"
    
    # Check if the command is a batch file (.cmd or .bat) or contains a .cmd/.bat file
    if ($Command -match '\.(cmd|bat)') {
        # For batch files, use the original method (wait for completion)
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "cmd.exe"
        $processInfo.Arguments = "/c `"$Command`""
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true
        $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $processInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        # Start the process
        $process.Start() | Out-Null
        
        # Stream output in real-time and collect separately
        $stdoutOutput = @()
        $stderrOutput = @()
        
        # Use asynchronous reading to avoid blocking
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        
        # Wait for process to complete
        $process.WaitForExit()
        
        # Get the output after process has completed
        $stdoutText = if ($stdoutTask -and $stdoutTask.Result) { $stdoutTask.Result } else { "" }
        $stderrText = if ($stderrTask -and $stderrTask.Result) { $stderrTask.Result } else { "" }
        
        # Split output into lines and display
        if ($stdoutText) {
            $stdoutLines = $stdoutText -split "`r?`n"
            foreach ($line in $stdoutLines) {
                if ($line.Trim()) {
                    Write-Host $line -ForegroundColor Green
                    $stdoutOutput += $line
                }
            }
        }
        
        if ($stderrText) {
            $stderrLines = $stderrText -split "`r?`n"
            foreach ($line in $stderrLines) {
                if ($line.Trim()) {
                    Write-Host $line -ForegroundColor Red
                    $stderrOutput += $line
                }
            }
        }
        
        $exitCode = $process.ExitCode
    } else {
        # For non-batch commands, stream output in real-time
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "powershell.exe"
        $processInfo.Arguments = "-Command `"$Command`""
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true
        $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $processInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        # Start the process
        $process.Start() | Out-Null
        
        # Stream output in real-time
        $stdoutOutput = @()
        $stderrOutput = @()
        
        # Read both stdout and stderr in a single loop
        while (-not $process.StandardOutput.EndOfStream -or -not $process.StandardError.EndOfStream) {
            # Read stdout if available
            if (-not $process.StandardOutput.EndOfStream) {
                $line = $process.StandardOutput.ReadLine()
                if ($line) {
                    Write-Host $line -ForegroundColor Green
                    $stdoutOutput += $line
                }
            }
            
            # Read stderr if available
            if (-not $process.StandardError.EndOfStream) {
                $line = $process.StandardError.ReadLine()
                if ($line) {
                    Write-Host $line -ForegroundColor Red
                    $stderrOutput += $line
                }
            }
        }
        
        # Wait for process to complete
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    }
    
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
        [string]$Message = "Press Ctrl+C to cancel..."
    )
    
    Write-Host "$Message" -ForegroundColor Yellow
    Write-Host "Waiting $Seconds seconds..." -ForegroundColor Yellow
    Start-Sleep -Seconds $Seconds
    Write-Host "Continuing..." -ForegroundColor Green
}

function Install-ChocolateyAndPackages {
    Write-Log "Installing Chocolatey and packages..."

    # Check if Chocolatey is already installed
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Log "Chocolatey is already installed. Updating..."
        $chocoUpgradeResult = Invoke-CommandWithExitCode -Command "choco upgrade chocolatey -y" -Description "upgrade Chocolatey"
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
    $chocoCacheClearResult = Invoke-CommandWithExitCode -Command "choco cache remove --expired -y" -Description "clear Chocolatey downloads"

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
    $chocoInstallResult = Invoke-CommandWithExitCode -Command "choco install $packageList -y --ignore-checksums" -Description "install all packages: $packageList"

    # Update all packages
    $chocoUpgradeAllResult = Invoke-CommandWithExitCode -Command "choco upgrade all -y --ignore-checksums" -Description "upgrade all packages"

    # Clean up
    $chocoCacheCleanupResult = Invoke-CommandWithExitCode -Command "choco cache remove --expired -y" -Description "clean up Chocolatey cache"

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
    Write-Log "Preparing Windows Update check..."
    
    # Clear Windows Update cache to resolve stuck states
    # Write-Log "Clearing Windows Update cache..."
    # Stop-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue
    # Start-Sleep -Seconds 3
    # if (Test-Path "$env:SystemRoot\SoftwareDistribution\DataStore") {
    #     Remove-Item -Path "$env:SystemRoot\SoftwareDistribution\DataStore\*" -Recurse -Force -ErrorAction SilentlyContinue
    # }
    # Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
    # Start-Sleep -Seconds 5
    # Write-Log "Windows Update cache cleared and service restarted"

    Write-Log "Checking for available Windows updates..."

    # Create Windows Update session
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    # Search for updates with timeout to prevent hanging
    Write-Log "Searching for updates (timeout: 10 minutes)..."
    
    # Force online scan by setting Online property to true
    $UpdateSearcher.Online = $true
    
    # Begin asynchronous search
    $searchCriteria = "IsInstalled=0 and Type='Software' or IsInstalled=0 and Type='Driver'"
    $dummyCallback = [System.AsyncCallback]{}
    $dummyState1 = New-Object System.Object
    $searchJob = $UpdateSearcher.BeginSearch($searchCriteria, $dummyCallback, $dummyState1)
    
    # Monitor search completion with timeout
    $timeoutSeconds = 600  # 10 minutes
    $elapsedSeconds = 0
    
    Write-Host "Searching for updates..." -NoNewline -ForegroundColor Yellow
    while (-not $searchJob.IsCompleted -and $elapsedSeconds -lt $timeoutSeconds) {
        Start-Sleep -Seconds 2
        $elapsedSeconds+=2
        Write-Host "." -NoNewline -ForegroundColor Yellow
    }
    
    if (-not $searchJob.IsCompleted) {
        throw "Windows Update search timed out after $timeoutSeconds seconds"
    }

    Write-Host ""
        
    # End the search and get the result
    $searchResult = $UpdateSearcher.EndSearch($searchJob)
    
    if ($searchResult.Updates.Count -eq 0) {
        Write-Log "No Windows updates available. System is up to date."
        return
    } else {
        Write-Log "Found $($searchResult.Updates.Count) Windows updates..."
    }

    # List available updates
    foreach ($updateInfo in $searchResult.Updates) {
        Write-Log-Highlight "Available update: $($updateInfo.Title)" -HighlightText $updateInfo.Title -HighlightColor "Cyan"
    }
    
    # Shuffle the order of updates (only if updates are available)
    $UpdatesArray = @($searchResult.Updates)
    if ($UpdatesArray.Count -gt 0) {
        $ShuffledUpdates = $UpdatesArray | Get-Random -Count $UpdatesArray.Count
    }
    
    foreach ($Update in $ShuffledUpdates) {
        Write-Log-Highlight "Processing update: $($Update.Title)" -HighlightText $Update.Title -HighlightColor "Green"
        
        # Create update collection for single update
        $SingleUpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        $SingleUpdateCollection.Add($Update) | Out-Null
        
        # Download single update with progress tracking
        Write-Log-Highlight "Downloading update: $($Update.Title)" -HighlightText $Update.Title -HighlightColor "Yellow"
        
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $SingleUpdateCollection
        $Downloader.Priority = 3
        
        # Begin asynchronous download with callback
        $dummyCallback = [System.AsyncCallback]{}
        $dummyState1 = New-Object System.Object
        $dummyState2 = New-Object System.Object
        $DownloadJob = $Downloader.BeginDownload($dummyCallback, $dummyState1, $dummyState2)
        
        # Wait for download to complete with progress monitoring
        while (-not $DownloadJob.IsCompleted) {
            Start-Sleep -Seconds 1
            
            # Monitor download progress
            try {
                Write-Progress -Activity "Downloading Windows Update: $($Update.Title)" -Status "Downloading..." -PercentComplete $DownloadJob.GetProgress().PercentComplete
            } catch {
                Write-Log "Error monitoring download progress: $($_.Exception.Message)"
            }
        }
        
        # End the download and get the result
        $DownloadResult = $Downloader.EndDownload($DownloadJob)
        
        Write-Progress -Activity "Downloading Windows Update: $($Update.Title)" -Completed
        
        if ($Update.IsDownloaded) {
            Write-Log-Highlight "Update downloaded successfully: $($Update.Title)" -HighlightText $Update.Title -HighlightColor "Green"
            
            # Install single update with progress tracking
            Write-Log-Highlight "Installing update: $($Update.Title)" -HighlightText $Update.Title -HighlightColor "Magenta"
            $Installer = $UpdateSession.CreateUpdateInstaller()
            $Installer.Updates = $SingleUpdateCollection
            
            # Begin asynchronous installation with callback
            $dummyCallback = [System.AsyncCallback]{}
            $dummyState1 = New-Object System.Object
            $dummyState2 = New-Object System.Object
            $InstallJob = $Installer.BeginInstall($dummyCallback, $dummyState1, $dummyState2)
            
            # Wait for installation to complete with progress monitoring
            while (-not $InstallJob.IsCompleted) {
                Start-Sleep -Seconds 1
                
                # Monitor installation progress
                try {
                    Write-Progress -Activity "Installing Windows Update: $($Update.Title)" -Status "Installing..." -PercentComplete $InstallJob.GetProgress().PercentComplete
                } catch {
                    Write-Log "Error monitoring installation progress: $($_.Exception.Message)"
                }
            }
            
            # End the installation and get the result
            $InstallResult = $Installer.EndInstall($InstallJob)
            
            Write-Progress -Activity "Installing Windows Update: $($Update.Title)" -Completed

            if ($InstallResult.ResultCode -eq 2) {
                Write-Log-Highlight "Update installed successfully: $($Update.Title)" -HighlightText $Update.Title -HighlightColor "Green"
                if ($InstallResult.RebootRequired) {
                    Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
                    Write-Host "Windows updates require a system reboot." -ForegroundColor Yellow
                    Write-Host "The computer will restart in 30 seconds..." -ForegroundColor Yellow
                    Wait-ForUserCancellation -Seconds 30 -Message "Press any key to cancel the reboot"
                    Write-Log "Reboot required. Restarting computer..."
                    # Restart-Computer -Force
                    Pause
                }
            } else {
                throw "Failed to install update: $($Update.Title)"
            }
        } else {
            throw "Failed to download update: $($Update.Title)"
        }
    }
    
    Write-Log "All updates processed"
    Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
    Write-Host "Windows updates require a system reboot." -ForegroundColor Yellow
    Write-Host "The computer will restart in 30 seconds..." -ForegroundColor Yellow
    Wait-ForUserCancellation -Seconds 30 -Message "Press any key to cancel the reboot"
    Write-Log "Reboot required. Restarting computer..."
    # Restart-Computer -Force
    Pause
}

function Install-Office {
    Write-Log "Starting Office installation process..."

    # Create temporary directory for Office installation
    $officeTempDir = "$env:TEMP\OfficeInstall"
    Write-Log "Creating temporary directory: $officeTempDir"
    if (Test-Path $officeTempDir) {
        Remove-Item $officeTempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $officeTempDir -Force | Out-Null

    # Download Office deployment tool
    $officeDownloadUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20136.exe"
    $officeInstaller = "officedeploymenttool.exe"
    $officeInstallerFullPath = Join-Path $officeTempDir $officeInstaller
    
    Invoke-WebRequestWithCleanup -Uri $officeDownloadUrl -OutFile $officeInstallerFullPath -Description "Office deployment tool"

    # Run the Office deployment tool to extract files
    Write-Log "Extracting Office deployment tool files..."
    & $officeInstallerFullPath /quiet /extract:$officeTempDir
    Write-Log "Office deployment tool extracted successfully"

    # Download office.xml from GitHub repository
    $officeXmlUrl = "https://raw.githubusercontent.com/goastler/windows/refs/heads/main/src/office.xml"
    $officeXmlDest = "$officeTempDir\office.xml"
    
    Invoke-WebRequestWithCleanup -Uri $officeXmlUrl -OutFile $officeXmlDest -Description "office.xml configuration file"

    # Run Office setup with the configuration file
    Write-Log "Starting Office installation with configuration file..."
    $setupExePath = "$officeTempDir\setup.exe"
    $setupProcess = Start-Process -FilePath $setupExePath -ArgumentList "/configure", $officeXmlDest -Wait -PassThru -NoNewWindow
    if ($setupProcess.ExitCode -ne 0) {
        throw "Office installation completed with exit code: $($setupProcess.ExitCode)"
    }
    Write-Log "Office installation completed successfully"

    # Clean up temporary directory
    Write-Log "Cleaning up temporary Office installation directory..."
    Remove-Item $officeTempDir -Recurse -Force
    Write-Log "Temporary directory cleaned up successfully"

    Write-Log "Office installation process completed!"
}

function Setup-BgInfo {
    Write-Log "Setting up BgInfo..."

    # BgInfo command line arguments
    $bgInfoArgs = "/NOLICPROMPT /TIMER:0 /ALL /SILENT"

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
        $action = New-ScheduledTaskAction -Execute $bgInfoExe -Argument $bgInfoArgs
        
        # Create the trigger (at logon for current user)
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        
        # Create the principal (run as current user with highest privileges)
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
        
        # Create the settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false
        
        # Register the task
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "BgInfo - Display system information on desktop background"
        
        Write-Log "BgInfo startup task created successfully"
    } else {
        Write-Log "BgInfo startup task already exists"
    }

    # Trigger the BgInfo scheduled task to apply the configuration
    Write-Log "Triggering BgInfo scheduled task to apply configuration..."
    Start-ScheduledTask -TaskName $taskName
    Write-Log "BgInfo scheduled task triggered successfully"

    Write-Log "BgInfo setup completed!"
}

function Install-MicrosoftActivationScripts {
    Write-Log "Setting up Microsoft Activation Scripts..."

    # Use Downloads directory for MAS
    $downloadsDir = "$env:USERPROFILE\Downloads"
    Write-Log "Using Downloads directory: $downloadsDir"

    # Download Microsoft Activation Scripts
    $masUrl = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd"
    $masFile = "MAS_AIO.cmd"
    $masFullPath = Join-Path $downloadsDir $masFile
    
    Write-Log "Downloading Microsoft Activation Scripts..."
    Invoke-WebRequestWithCleanup -Uri $masUrl -OutFile $masFullPath -Description "Microsoft Activation Scripts"

    # Verify the file was downloaded successfully
    if (!(Test-Path $masFullPath)) {
        throw "Failed to download MAS_AIO.cmd. File not found at: $masFullPath"
    }
    
    Write-Log "MAS_AIO.cmd downloaded successfully to: $masFullPath"
    
    # Get file size to verify it's not empty
    $fileSize = (Get-Item $masFullPath).Length
    if ($fileSize -eq 0) {
        throw "Downloaded MAS_AIO.cmd file is empty (0 bytes)"
    }
    Write-Log "MAS_AIO.cmd file size: $fileSize bytes"

    # Run MAS with /HWID parameter
    Write-Log "Running Microsoft Activation Scripts with /HWID parameter..."
    $masHwidResult = Invoke-CommandWithExitCode -Command "`"$masFullPath`" /HWID" -Description "run MAS with /HWID parameter"

    # Run MAS with /Ohook parameter
    Write-Log "Running Microsoft Activation Scripts with /Ohook parameter..."
    $masOhookResult = Invoke-CommandWithExitCode -Command "`"$masFullPath`" /Ohook" -Description "run MAS with /Ohook parameter"

    Write-Log "Microsoft Activation Scripts setup completed!"
}

function Create-OnLoginScheduledTask {
    Write-Log "Creating scheduled task for onLogin.ps1 script..."
    
    # Define the script path (current script location)
    $scriptPath = $PSCommandPath
    
    # Check if task already exists
    $taskExists = Get-ScheduledTask -TaskName $scheduledTaskName -ErrorAction SilentlyContinue

    if ($taskExists) {
        Write-Log "Scheduled task '$scheduledTaskName' already exists. Removing it..."
        Unregister-ScheduledTask -TaskName $scheduledTaskName -Confirm:$false
        Write-Log "Scheduled task '$scheduledTaskName' removed."
    }
    
     Write-Log "Creating scheduled task: $scheduledTaskName"
     
     # Create the action to run this script with visible PowerShell window
     $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Normal -File `"$scriptPath`""
     
     # Create the trigger (at logon for current user)
     $trigger = New-ScheduledTaskTrigger -AtLogOn
     
     # Create the principal (run as current user with highest privileges)
     $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
     
     # Create the settings
     $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable:$false
     
     # Register the task
     Register-ScheduledTask -TaskName $scheduledTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "OnLogin Setup Script - Runs system setup and configuration on user logon"
     
     Write-Log "Scheduled task '$scheduledTaskName' created successfully"
     Write-Log "Task will run: $scriptPath"
}

try {
    Write-Log "Starting on login setup..."

    # =============================================================================
    # CREATE SCHEDULED TASK
    # =============================================================================

    Create-OnLoginScheduledTask

    # =============================================================================
    # USER CANCELLATION DELAY
    # =============================================================================

    Write-Host "`n=== On Login Setup Starting ===" -ForegroundColor Yellow
    Wait-ForUserCancellation -Seconds 10
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

    # Setup-BgInfo

    # =============================================================================
    # OFFICE INSTALLATION
    # =============================================================================

    Install-Office

    =============================================================================
    MICROSOFT ACTIVATION SCRIPTS
    =============================================================================

    Install-MicrosoftActivationScripts

    # =============================================================================
    # FINAL CLEANUP - REMOVE SCHEDULED TASK
    # =============================================================================

    Write-Log "Setup completed successfully. Removing setup scheduled task..."
    Unregister-ScheduledTask -TaskName $scheduledTaskName -Confirm:$false
    Write-Log "Setup scheduled task removed successfully"

    # =============================================================================
    # COMPLETION
    # =============================================================================

    Write-Log "On login setup completed. Log saved to: $logFile"

    Pause

    exit 0

} catch {
    Write-Host "`n=== ERROR ===" -ForegroundColor Red
    Write-Host "An error occurred during setup: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception)" -ForegroundColor Red
    
    # Restore original directory location on error
    try {
        Set-Location $originalScriptLocation
        Write-Log "Restored to original directory: $originalScriptLocation"
    }
    catch {
        Write-Log "Warning: Could not restore to original directory: $($_.Exception.Message)"
    }
}

Write-Host "`n=== REBOOT REQUIRED ===" -ForegroundColor Yellow
Write-Host "The system will reboot in 10 seconds..." -ForegroundColor Yellow
Wait-ForUserCancellation -Seconds 10
Write-Host "Rebooting now..." -ForegroundColor Red
# Restart-Computer -Force
Pause
