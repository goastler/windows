# #Requires -RunAsAdministrator

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Input ISO file does not exist: $_"
        }
        if ($_ -notmatch '\.iso$') {
            throw "Input file must have .iso extension: $_"
        }
        $true
    })]
    [string]$InputIso,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $parentDir = Split-Path $_ -Parent
        if (-not (Test-Path $parentDir -PathType Container)) {
            throw "Output directory does not exist: $parentDir"
        }
        if ($_ -notmatch '\.iso$') {
            throw "Output file must have .iso extension: $_"
        }
        $true
    })]
    [string]$OutputIso,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Autounattend XML file does not exist: $_"
        }
        $true
    })]
    [string]$AutounattendXml = (Join-Path $PSScriptRoot "autounattend.xml"),

    [Parameter(Mandatory = $false)]
    [string]$OemDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) '$OEM$'),

    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = "C:\WinIsoRepack_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [switch]$KeepWorkingDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeVirtioDrivers,

    [Parameter(Mandatory = $false)]
    [ValidateSet("stable", "latest")]
    [string]$VirtioVersion = "stable",

    [Parameter(Mandatory = $false)]
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache")

)

$ErrorActionPreference = "Stop"

function Invoke-WebRequestWithCleanup {
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Description = "download file",
        [int]$ProgressId = 3
    )
    
    Write-Log "Downloading $Description from: $Uri"
    
    $webRequest = $null
    try {
        # Create a web client for progress tracking
        $webClient = New-Object System.Net.WebClient
        
        # Set up progress tracking
        $totalBytes = 0
        $downloadedBytes = 0
        
        # Register for download progress
        Register-ObjectEvent -InputObject $webClient -EventName "DownloadProgressChanged" -Action {
            $global:downloadProgress = $Event.SourceEventArgs
        } | Out-Null
        
        # Get file size first
        try {
            $headRequest = [System.Net.WebRequest]::Create($Uri)
            $headRequest.Method = "HEAD"
            $response = $headRequest.GetResponse()
            $totalBytes = $response.ContentLength
            $response.Close()
        } catch {
            Write-Log "Could not determine file size, progress tracking may be limited"
        }
        
        # Start download with progress tracking
        $downloadTask = $webClient.DownloadFileTaskAsync($Uri, $OutFile)
        
        # Monitor progress
        while (-not $downloadTask.IsCompleted) {
            Start-Sleep -Milliseconds 100
            
            if ($global:downloadProgress) {
                $percentComplete = if ($totalBytes -gt 0) {
                    [math]::Round(($global:downloadProgress.BytesReceived / $totalBytes) * 100)
                } else {
                    [math]::Round(($global:downloadProgress.BytesReceived / ($global:downloadProgress.BytesReceived + 1)) * 100)
                }
                
                $downloadedMB = [math]::Round($global:downloadProgress.BytesReceived / 1MB, 2)
                $totalMB = if ($totalBytes -gt 0) { [math]::Round($totalBytes / 1MB, 2) } else { "Unknown" }
                
                Write-ProgressWithPercentage -Activity "Downloading $Description" -Status "Downloaded $downloadedMB MB of $totalMB MB" -PercentComplete $percentComplete -Id $ProgressId
            }
        }
        
        # Wait for completion and handle any exceptions
        $downloadTask.Wait()
        if ($downloadTask.Exception) {
            throw $downloadTask.Exception
        }
        
        Write-Progress -Activity "Downloading $Description" -Completed -Id $ProgressId
        Write-Host "" # Clear the progress line
        Write-Log "$Description downloaded successfully"
    }
    finally {
        # Clean up event subscription
        try {
            Get-EventSubscriber | Where-Object { $_.SourceObject -eq $webClient } | Unregister-Event
        } catch {
            # Ignore cleanup errors
        }
        
        # Properly dispose of web client
        if ($webClient) {
            try {
                $webClient.Dispose()
            }
            catch {
                # Ignore disposal errors
            }
        }
        
        # Properly dispose of web request object
        if ($webRequest) {
            try {
                $webRequest.Dispose()
            }
            catch {
                # Ignore disposal errors
            }
        }
        
        # Force garbage collection to ensure resources are released
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Brief pause to ensure file handles are released
        Start-Sleep -Seconds 1
    }
}

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White",
        [int]$Indent = 0
    )
    
    # Create indentation string (2 spaces per indent level)
    $indentString = "  " * $Indent
    
    # Combine indentation with message
    $indentedMessage = $indentString + $Message
    
    Write-Host $indentedMessage -ForegroundColor $Color
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Find-OscdimgPath {
    $adkPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\8.1\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
    )
    
    foreach ($path in $adkPaths) {
        if (Test-Path $path) {
            Write-ColorOutput "Found oscdimg.exe at: $path" "Green" -Indent 1
            return $path
        }
    }
    
    throw "oscdimg.exe not found. Please ensure Windows ADK is properly installed."
}

function Test-RequiredTools {
    Write-ColorOutput "Installing Windows ADK via Chocolatey..." "Yellow"
    
    # Always install Windows ADK via Chocolatey (it won't reinstall if already present)
    Install-WindowsADK
    
    # Find oscdimg.exe path
    $script:oscdimgPath = Find-OscdimgPath
    
    # Ensure DISM is available
    Test-DismAvailability
}

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
                Write-ColorOutput "Added to PATH: $path" "Cyan"
            }
        }
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        return
    }

    throw "Windows ADK installation failed"
}

function Extract-IsoContents {
    param(
        [string]$IsoPath,
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Mounting ISO: $IsoPath" "Yellow"
    $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter
    
    if (-not $driveLetter) {
        throw "Failed to mount ISO or get drive letter"
    }
    
    $mountedPath = "${driveLetter}:\"
    Write-ColorOutput "ISO mounted at: $mountedPath" "Green" -Indent 1
    
    try {
        Write-ColorOutput "Extracting ISO contents to: $ExtractPath" "Yellow" -Indent 1
        
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
        robocopy $mountedPath $ExtractPath /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
        if ($LASTEXITCODE -gt 7) {
            throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
        }

        Write-ColorOutput "ISO contents extracted successfully" "Green" -Indent 1
    } finally {
        Write-ColorOutput "Dismounting ISO..." "Yellow" -Indent 1
        Dismount-DiskImage -ImagePath $IsoPath
        Write-ColorOutput "ISO dismounted" "Green" -Indent 1
    }
}

function Add-AutounattendXml {
    param(
        [string]$ExtractPath,
        [string]$AutounattendXmlPath
    )
    Write-ColorOutput "Adding autounattend.xml to ISO contents..." "Yellow"
    $destinationPath = Join-Path $ExtractPath "autounattend.xml"
    Copy-Item $AutounattendXmlPath $destinationPath -Force
    Write-ColorOutput "autounattend.xml added to: $destinationPath" "Green" -Indent 1
}

function Add-OemDirectory {
    param(
        [string]$ExtractPath,
        [string]$OemSourcePath
    )
    Write-ColorOutput "Adding $OEM$ directory to ISO contents..." "Yellow"
    
    if (-not (Test-Path $OemSourcePath -PathType Container)) {
        Write-ColorOutput "Warning: $OEM$ directory not found at: $OemSourcePath" "Yellow" -Indent 1
        return
    }
    
    $destinationPath = Join-Path $ExtractPath '$OEM$'
    
    # Remove existing $OEM$ directory if it exists
    if (Test-Path $destinationPath) {
        Remove-Item $destinationPath -Recurse -Force
    }
    
    # Copy the entire $OEM$ directory structure
    Copy-Item $OemSourcePath $destinationPath -Recurse -Force
    Write-ColorOutput "$OEM$ directory added to: $destinationPath" "Green" -Indent 1
}

function New-IsoFromDirectory {
    param(
        [string]$SourcePath,
        [string]$OutputPath,
        [string]$OscdimgPath
    )
    Write-ColorOutput "Creating new ISO from directory: $SourcePath" "Yellow"

    # Resolve absolute paths
    $absSrc    = (Resolve-Path $SourcePath).ProviderPath
    $absOutDir = (Resolve-Path (Split-Path $OutputPath -Parent)).ProviderPath
    $absOutIso = Join-Path $absOutDir (Split-Path $OutputPath -Leaf)

    $etfsbootPath  = "$absSrc\boot\etfsboot.com"
    $efisysPath    = "$absSrc\efi\microsoft\boot\efisys.bin"

    Write-ColorOutput "Using source directly: $absSrc" "Cyan" -Indent 1

    $arguments = @(
        "-m"
        "-u2"
        "-udfver102"
        "-bootdata:2#p0,e,b`"$etfsbootPath`"#pEF,e,b`"$efisysPath`""
        "`"$absSrc`""
        "`"$absOutIso`""
    )

    Write-ColorOutput "Current working directory: $(Get-Location)" "Cyan" -Indent 1
    Write-ColorOutput "Running oscdimg with arguments: $($arguments -join ' ')" "Cyan" -Indent 1
    Write-ColorOutput "Full command: & `"$OscdimgPath`" $($arguments -join ' ')" "Cyan" -Indent 1
    
    & $OscdimgPath $arguments
    if ($LASTEXITCODE -ne 0) { throw "oscdimg failed with exit code: $LASTEXITCODE" }

    Write-ColorOutput "ISO created successfully: $absOutIso" "Green" -Indent 1
}

function Remove-WorkingDirectory {
    param([string]$Path)
    if (-not $KeepWorkingDirectory -and (Test-Path $Path)) {
        Write-ColorOutput "Cleaning up working directory: $Path" "Yellow"
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Working directory cleaned up" "Green"
    } elseif ($KeepWorkingDirectory) {
        Write-ColorOutput "Keeping working directory: $Path" "Cyan"
    }
}

function Get-VirtioDownloadUrl {
    param([string]$Version)
    
    $urls = @{
        "stable" = "https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso"
        "latest" = "https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/latest-virtio/virtio-win.iso"
    }
    
    return $urls[$Version]
}

function Get-VirtioDrivers {
    param(
        [string]$Version,
        [string]$CacheDirectory
    )
    
    Write-ColorOutput "=== VirtIO Drivers Download ===" "Cyan"
    
    # Create cache directory if it doesn't exist
    if (-not (Test-Path $CacheDirectory)) {
        New-Item -ItemType Directory -Path $CacheDirectory -Force | Out-Null
        Write-ColorOutput "Created cache directory: $CacheDirectory" "Green" -Indent 1
    }
    
    $downloadUrl = Get-VirtioDownloadUrl -Version $Version
    $fileName = "virtio-win-$Version.iso"
    $localPath = Join-Path $CacheDirectory $fileName
    
    # Check if we already have the file
    if (Test-Path $localPath) {
        Write-ColorOutput "VirtIO drivers already cached: $localPath" "Green" -Indent 1
        return $localPath
    }
    
    try {
        # Use Invoke-WebRequestWithCleanup for better progress tracking and resource cleanup
        Invoke-WebRequestWithCleanup -Uri $downloadUrl -OutFile $localPath -Description "VirtIO drivers ($Version)" -ProgressId 3
        Write-ColorOutput "VirtIO drivers downloaded successfully" "Green" -Indent 1
        return $localPath
    } catch {
        Write-ColorOutput "Failed to download VirtIO drivers: $($_.Exception.Message)" "Red"
        throw
    }
}

function Extract-VirtioDrivers {
    param(
        [string]$VirtioIsoPath,
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Extracting VirtIO drivers from: $VirtioIsoPath" "Yellow"
    
    # Create virtio directory in the ISO extract path
    $virtioDir = Join-Path $ExtractPath "virtio"
    if (Test-Path $virtioDir) {
        Remove-Item $virtioDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $virtioDir -Force | Out-Null
    
    try {
        # Mount the VirtIO ISO
        Write-ColorOutput "Mounting VirtIO ISO..." "Yellow"
        $mountResult = Mount-DiskImage -ImagePath $VirtioIsoPath -PassThru
        $driveLetter = ($mountResult | Get-Volume).DriveLetter
        
        if (-not $driveLetter) {
            throw "Failed to mount VirtIO ISO or get drive letter"
        }
        
        $mountedPath = "${driveLetter}:\"
        Write-ColorOutput "VirtIO ISO mounted at: $mountedPath" "Green"
        
        try {
            # Copy all contents from VirtIO ISO to the virtio directory
            Write-ColorOutput "Copying VirtIO drivers to: $virtioDir" "Yellow"
            robocopy $mountedPath $virtioDir /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
            
            if ($LASTEXITCODE -gt 7) {
                throw "Failed to copy VirtIO drivers. Robocopy exit code: $LASTEXITCODE"
            }
            
            Write-ColorOutput "VirtIO drivers extracted successfully" "Green"
        } finally {
            Write-ColorOutput "Dismounting VirtIO ISO..." "Yellow"
            Dismount-DiskImage -ImagePath $VirtioIsoPath
            Write-ColorOutput "VirtIO ISO dismounted" "Green"
        }
        
        return $virtioDir
    } catch {
        Write-ColorOutput "Failed to extract VirtIO drivers: $($_.Exception.Message)" "Red"
        throw
    }
}

function Add-VirtioDriversToWim {
    param(
        [hashtable]$WimInfo,
        [string]$VirtioDir,
        [string]$VirtioVersion
    )
    
    $arch = $WimInfo.Architecture
    $version = $WimInfo.Version
    $wimPath = $WimInfo.Path
    $wimType = $WimInfo.Type
    $imageIndex = $WimInfo.Index
    $imageName = $WimInfo.Name
    
    Write-ColorOutput "Processing $wimType image: $imageName" "Yellow" -Indent 1
    
    # Validate Windows 11 architecture compatibility
    if ($version -eq "w11" -and $arch -eq "x86") {
        Write-ColorOutput "Skipping Windows 11 x86 image (not supported)" "Yellow" -Indent 2
        return
    }
    
    # Validate VirtIO driver availability for ARM64
    if ($arch -eq "arm64") {
        Write-ColorOutput "Skipping ARM64 image (VirtIO drivers not available)" "Yellow" -Indent 2
        return
    }
    
    # Check if appropriate drivers exist
    $driverPath = Join-Path $VirtioDir $arch
    if (-not (Test-Path $driverPath)) {
        Write-ColorOutput "No VirtIO drivers found for architecture $arch" "Yellow" -Indent 2
        return
    }
    
    $windowsDriverPath = Join-Path $driverPath $version
    if (-not (Test-Path $windowsDriverPath)) {
        Write-ColorOutput "No VirtIO drivers found for Windows version $version" "Yellow" -Indent 2
        return
    }
    
    Write-ColorOutput "Adding VirtIO drivers (Arch: $arch, Version: $version)" "Green" -Indent 2
    
    try {
        if ($wimType -eq "boot") {
            Inject-VirtioDriversIntoBootWim -WimPath $wimPath -VirtioDir $VirtioDir -Arch $arch -Version $version -ImageIndex $imageIndex
        } else {
            Inject-VirtioDriversIntoInstallWim -WimPath $wimPath -VirtioDir $VirtioDir -Arch $arch -Version $version -ImageIndex $imageIndex
        }
    } catch {
        Write-ColorOutput "Failed to add VirtIO drivers to $wimType image: $($_.Exception.Message)" "Red" -Indent 2
        throw
    }
}

function Add-VirtioDrivers {
    param(
        [string]$ExtractPath,
        [string]$VirtioVersion,
        [string]$VirtioCacheDirectory,
        [array]$WimInfos
    )
    
    if (-not $IncludeVirtioDrivers) {
        Write-ColorOutput "VirtIO drivers not requested, skipping..." "Cyan"
        return
    }
    
    Write-ColorOutput "=== Adding VirtIO Drivers ===" "Cyan"
    
    try {
        # Download VirtIO drivers
        $virtioIsoPath = Get-VirtioDrivers -Version $VirtioVersion -CacheDirectory $VirtioCacheDirectory
        
        # Extract VirtIO drivers to a temporary directory
        $virtioDir = Extract-VirtioDrivers -VirtioIsoPath $virtioIsoPath -ExtractPath $ExtractPath
        
        Write-ColorOutput "VirtIO drivers extracted to: $virtioDir" "Green" -Indent 1
        
        # Log the driver structure for debugging
        $driverDirs = Get-ChildItem -Path $virtioDir -Directory | Select-Object -ExpandProperty Name
        Write-ColorOutput "Available driver directories: $($driverDirs -join ', ')" "Cyan" -Indent 1
        
        # Process each WIM image individually
        foreach ($wimInfo in $WimInfos) {
            Add-VirtioDriversToWim -WimInfo $wimInfo -VirtioDir $virtioDir -VirtioVersion $VirtioVersion
        }
        
    } catch {
        Write-ColorOutput "Failed to add VirtIO drivers: $($_.Exception.Message)" "Red"
        throw
    }
}

function Inject-VirtioDriversIntoBootWim {
    param(
        [string]$WimPath,
        [string]$VirtioDir,
        [string]$Arch,
        [string]$Version,
        [int]$ImageIndex
    )
    
    Write-ColorOutput "=== Injecting VirtIO Drivers into boot.wim (Index: $ImageIndex) ===" "Cyan" -Indent 1
    
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "boot_mount_$ImageIndex"
    
    if (-not (Test-Path $WimPath)) {
        throw "boot.wim not found at: $WimPath"
    }
    
    try {
        # Create mount directory
        if (Test-Path $mountDir) {
            Remove-Item $mountDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
        
        # Use the specified architecture and Windows version
        $driverPath = Join-Path $VirtioDir $Arch
        
        if (-not (Test-Path $driverPath)) {
            Write-ColorOutput "Error: No drivers found for architecture $Arch at: $driverPath" "Red" -Indent 2
            throw "VirtIO drivers not found for architecture: $Arch"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            Write-ColorOutput "Error: No drivers found for Windows version $Version at: $windowsDriverPath" "Red" -Indent 2
            throw "VirtIO drivers not found for Windows version: $Version"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" "Green" -Indent 2
        Write-ColorOutput "Architecture: $Arch" "Cyan" -Indent 2
        Write-ColorOutput "Version: $Version" "Cyan" -Indent 2
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Mount the specific boot.wim index
        Write-ColorOutput "Mounting boot.wim index $ImageIndex..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Mount-Wim",
            "/WimFile:`"$WimPath`"",
            "/Index:$ImageIndex",
            "/MountDir:`"$mountDir`""
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to mount boot.wim index $ImageIndex (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to mount boot.wim index $ImageIndex"
        }
        
        Write-ColorOutput "Successfully mounted boot.wim index $ImageIndex" "Green" -Indent 2
        
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to boot.wim..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Image:`"$mountDir`"",
            "/Add-Driver",
            "/Driver:`"$windowsDriverPath`"",
            "/Recurse"
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully added VirtIO drivers to boot.wim" "Green" -Indent 2
        } else {
            Write-ColorOutput "Warning: Failed to add drivers to boot.wim (exit code: $($result.ExitCode))" "Yellow" -Indent 2
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting boot.wim..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Unmount-Wim",
            "/MountDir:`"$mountDir`"",
            "/Commit"
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully unmounted and committed boot.wim changes" "Green" -Indent 2
        } else {
            Write-ColorOutput "Warning: Failed to unmount boot.wim (exit code: $($result.ExitCode))" "Yellow" -Indent 2
        }
        
    } catch {
        Write-ColorOutput "Error injecting drivers into boot.wim: $($_.Exception.Message)" "Red" -Indent 2
        throw
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" "Yellow" -Indent 2
            }
        }
    }
}

function Inject-VirtioDriversIntoInstallWim {
    param(
        [string]$WimPath,
        [string]$VirtioDir,
        [string]$Arch,
        [string]$Version,
        [int]$ImageIndex
    )
    
    Write-ColorOutput "=== Injecting VirtIO Drivers into install.wim (Index: $ImageIndex) ===" "Cyan" -Indent 1
    
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "install_mount_$ImageIndex"
    
    if (-not (Test-Path $WimPath)) {
        Write-ColorOutput "Warning: install.wim not found at: $WimPath" "Yellow" -Indent 2
        return
    }
    
    try {
        # Create mount directory
        if (Test-Path $mountDir) {
            Remove-Item $mountDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
        
        # Use the specified architecture and Windows version
        $driverPath = Join-Path $VirtioDir $Arch
        
        if (-not (Test-Path $driverPath)) {
            Write-ColorOutput "Error: No drivers found for architecture $Arch at: $driverPath" "Red" -Indent 2
            throw "VirtIO drivers not found for architecture: $Arch"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            Write-ColorOutput "Error: No drivers found for Windows version $Version at: $windowsDriverPath" "Red" -Indent 2
            throw "VirtIO drivers not found for Windows version: $Version"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" "Green" -Indent 2
        Write-ColorOutput "Architecture: $Arch" "Cyan" -Indent 2
        Write-ColorOutput "Version: $Version" "Cyan" -Indent 2
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Mount the specific install.wim index
        Write-ColorOutput "Mounting install.wim index $ImageIndex..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Mount-Wim",
            "/WimFile:`"$WimPath`"",
            "/Index:$ImageIndex",
            "/MountDir:`"$mountDir`""
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to mount install.wim index $ImageIndex (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to mount install.wim index $ImageIndex"
        }
        
        Write-ColorOutput "Successfully mounted install.wim index $ImageIndex" "Green" -Indent 2
        
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to install.wim..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Image:`"$mountDir`"",
            "/Add-Driver",
            "/Driver:`"$windowsDriverPath`"",
            "/Recurse"
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully added VirtIO drivers to install.wim" "Green" -Indent 2
        } else {
            Write-ColorOutput "Warning: Failed to add drivers to install.wim (exit code: $($result.ExitCode))" "Yellow" -Indent 2
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting install.wim..." "Yellow" -Indent 2
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Unmount-Wim",
            "/MountDir:`"$mountDir`"",
            "/Commit"
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully unmounted and committed install.wim changes" "Green" -Indent 2
        } else {
            Write-ColorOutput "Warning: Failed to unmount install.wim (exit code: $($result.ExitCode))" "Yellow" -Indent 2
        }
        
    } catch {
        Write-ColorOutput "Error injecting drivers into install.wim: $($_.Exception.Message)" "Red" -Indent 2
        throw
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" "Yellow" -Indent 2
            }
        }
    }
}

function Get-WimImageInfo {
    param(
        [string]$WimPath,
        [string]$DismPath
    )
    
    try {
        Write-ColorOutput "Getting WIM image information from: $WimPath" "Yellow"
        
        # Use DISM to get image information
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_wim_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM info (exit code: $($result.ExitCode))" "Yellow"
            return $null
        }
        
        # Parse the output to extract image information
        $wimInfo = Get-Content "temp_wim_info.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_wim_info.txt" -ErrorAction SilentlyContinue
        
        $images = @()
        $currentImage = $null
        
        foreach ($line in $wimInfo) {
            if ($line -match "Index\s*:\s*(\d+)") {
                if ($currentImage) {
                    $images += $currentImage
                }
                $currentImage = @{
                    Index = [int]$matches[1]
                    Name = ""
                    Description = ""
                }
            } elseif ($currentImage -and $line -match "Name\s*:\s*(.+)") {
                $currentImage.Name = $matches[1].Trim()
            } elseif ($currentImage -and $line -match "Description\s*:\s*(.+)") {
                $currentImage.Description = $matches[1].Trim()
            }
        }
        
        if ($currentImage) {
            $images += $currentImage
        }
        
        return $images
        
    } catch {
        Write-ColorOutput "Error getting WIM image info: $($_.Exception.Message)" "Yellow"
        return $null
    }
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

function Get-WimArchitecture {
    param(
        [string]$WimPath
    )
    
    Write-ColorOutput "Inferring architecture from WIM: $WimPath" "Yellow"
    
    try {
        $dismPath = Get-DismPath
        
        # Get WIM information using DISM
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_arch_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM architecture info (exit code: $($result.ExitCode))" "Yellow"
            return $null
        }
        
        # Parse the output to extract architecture information
        $wimInfo = Get-Content "temp_arch_info.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_arch_info.txt" -ErrorAction SilentlyContinue
        
        foreach ($line in $wimInfo) {
            if ($line -match "Architecture\s*:\s*(.+)") {
                $arch = $matches[1].Trim()
                Write-ColorOutput "Detected architecture: $arch" "Green" -Indent 1
                
                # Map DISM architecture names to our expected values
                switch ($arch.ToLower()) {
                    "x64" { return "amd64" }
                    "x86" { return "x86" }
                    "arm64" { return "arm64" }
                    default {
                        Write-ColorOutput "Unknown architecture: $arch" "Yellow"
                        return $null
                    }
                }
            }
        }
        
        Write-ColorOutput "Could not determine architecture from WIM" "Yellow"
        return $null
        
    } catch {
        Write-ColorOutput "Error inferring architecture from WIM: $($_.Exception.Message)" "Red"
        return $null
    }
}

function Get-WimVersion {
    param(
        [string]$WimPath
    )
    
    Write-ColorOutput "Inferring Windows version from WIM: $WimPath" "Yellow"
    
    try {
        $dismPath = Get-DismPath
        
        # Get WIM information using DISM
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_version_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM version info (exit code: $($result.ExitCode))" "Yellow"
            return $null
        }
        
        # Parse the output to extract version information
        $wimInfo = Get-Content "temp_version_info.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_version_info.txt" -ErrorAction SilentlyContinue
        
        foreach ($line in $wimInfo) {
            if ($line -match "Name\s*:\s*(.+)") {
                $name = $matches[1].Trim()
                Write-ColorOutput "Detected image name: $name" "Green" -Indent 1
                
                # Check for Windows 11 indicators
                if ($name -match "Windows 11" -or $name -match "Windows 1[1-9]") {
                    Write-ColorOutput "Detected Windows version: w11" "Green" -Indent 2
                    return "w11"
                }
                # Check for Windows 10 indicators
                elseif ($name -match "Windows 10") {
                    Write-ColorOutput "Detected Windows version: w10" "Green" -Indent 2
                    return "w10"
                }
            }
            elseif ($line -match "Description\s*:\s*(.+)") {
                $description = $matches[1].Trim()
                Write-ColorOutput "Detected image description: $description" "Green" -Indent 1
                
                # Check for Windows 11 indicators in description
                if ($description -match "Windows 11" -or $description -match "Windows 1[1-9]") {
                    Write-ColorOutput "Detected Windows version: w11" "Green" -Indent 2
                    return "w11"
                }
                # Check for Windows 10 indicators in description
                elseif ($description -match "Windows 10") {
                    Write-ColorOutput "Detected Windows version: w10" "Green" -Indent 2
                    return "w10"
                }
            }
        }
        
        Write-ColorOutput "Could not determine Windows version from WIM" "Yellow"
        return $null
        
    } catch {
        Write-ColorOutput "Error inferring Windows version from WIM: $($_.Exception.Message)" "Red"
        return $null
    }
}

function Get-AllWimInfo {
    param(
        [string]$ExtractPath
    )
    
    Write-ColorOutput "=== Analyzing All WIM Files ===" "Cyan"
    
    $wims = @()
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." "Yellow" -Indent 1
        $installWimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath (Get-DismPath)
        
        if ($installWimInfo) {
            foreach ($image in $installWimInfo) {
                $arch = Get-WimArchitecture -WimPath $installWimPath
                $version = Get-WimVersion -WimPath $installWimPath
                
                $wimInfo = @{
                    Path = $installWimPath
                    Type = "install"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $arch
                    Version = $version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $arch, Version: $version)" "Green" -Indent 2
            }
        }
    }
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." "Yellow" -Indent 1
        $bootWimInfo = Get-WimImageInfo -WimPath $bootWimPath -DismPath (Get-DismPath)
        
        if ($bootWimInfo) {
            foreach ($image in $bootWimInfo) {
                $arch = Get-WimArchitecture -WimPath $bootWimPath
                $version = Get-WimVersion -WimPath $bootWimPath
                
                $wimInfo = @{
                    Path = $bootWimPath
                    Type = "boot"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $arch
                    Version = $version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $arch, Version: $version)" "Green" -Indent 2
            }
        }
    }
    
    if ($wims.Count -eq 0) {
        throw "No WIM files found in the ISO"
    }
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" "Green"
    return $wims
}

function Get-WimInfo {
    param(
        [string]$ExtractPath
    )
    
    Write-ColorOutput "=== Inferring WIM Information ===" "Cyan"
    
    # Try to get info from install.wim first (more reliable)
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    $arch = $null
    $version = $null
    
    # Try install.wim first
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." "Yellow" -Indent 1
        $arch = Get-WimArchitecture -WimPath $installWimPath
        $version = Get-WimVersion -WimPath $installWimPath
    }
    
    # If we couldn't get info from install.wim, try boot.wim
    if (-not $arch -and (Test-Path $bootWimPath)) {
        Write-ColorOutput "Analyzing boot.wim..." "Yellow" -Indent 1
        $arch = Get-WimArchitecture -WimPath $bootWimPath
    }
    
    if (-not $version -and (Test-Path $bootWimPath)) {
        Write-ColorOutput "Analyzing boot.wim for version..." "Yellow" -Indent 1
        $version = Get-WimVersion -WimPath $bootWimPath
    }
    
    if (-not $arch) {
        throw "Could not determine architecture from WIM files"
    }
    
    if (-not $version) {
        throw "Could not determine Windows version from WIM files"
    }
    
    Write-ColorOutput "Inferred architecture: $arch" "Green"
    Write-ColorOutput "Inferred version: $version" "Green"
    
    return @{
        Architecture = $arch
        Version = $version
    }
}


try {
    Write-ColorOutput "=== Windows ISO Repack Script ===" "Cyan"
    Write-ColorOutput "Checking administrator privileges..." "Yellow"
    
    if (-not (Test-Administrator)) {
        Write-ColorOutput "ERROR: This script must be run as Administrator!" "Red"
        throw "Administrator privileges required."
    }
    
    Write-ColorOutput "Administrator privileges confirmed" "Green"
    Write-ColorOutput "Input ISO: $InputIso" "White"
    Write-ColorOutput "Output ISO: $OutputIso" "White"
    Write-ColorOutput "Autounattend XML: $AutounattendXml" "White"
    Write-ColorOutput "OEM Directory: $OemDirectory" "White"
    Write-ColorOutput "Working Directory: $WorkingDirectory" "White"
    Write-ColorOutput "Include VirtIO Drivers: $IncludeVirtioDrivers" "White"
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "VirtIO Version: $VirtioVersion" "White"
        Write-ColorOutput "VirtIO Cache Directory: $VirtioCacheDirectory" "White"
        Write-ColorOutput "Processing Mode: Per-WIM (architecture and version inferred from each WIM)" "White"
    }
    
    Test-RequiredTools
    
    Write-ColorOutput "Validating input files..." "Yellow"
    
    # Resolve and validate ISO paths
    try {
        $resolvedInputIso = Resolve-Path $InputIso -ErrorAction Stop
        Write-ColorOutput "Resolved input ISO path: $resolvedInputIso" "Cyan"
    } catch {
        throw "Cannot resolve input ISO file path: $InputIso. Error: $($_.Exception.Message)"
    }
    
    try {
        $resolvedOutputIso = Resolve-Path $OutputIso -ErrorAction SilentlyContinue
        if (-not $resolvedOutputIso) {
            # If path doesn't exist, resolve the parent directory and create the full path
            $outputDir = Split-Path $OutputIso -Parent
            $outputFile = Split-Path $OutputIso -Leaf
            $resolvedOutputDir = Resolve-Path $outputDir -ErrorAction Stop
            $resolvedOutputIso = Join-Path $resolvedOutputDir $outputFile
        }
        Write-ColorOutput "Resolved output ISO path: $resolvedOutputIso" "Cyan"
    } catch {
        throw "Cannot resolve output ISO file path: $OutputIso. Error: $($_.Exception.Message)"
    }
    
    if (-not (Test-Path $resolvedInputIso -PathType Leaf)) {
        throw "Input ISO file not found: $resolvedInputIso"
    }
    if (-not (Test-Path $AutounattendXml -PathType Leaf)) {
        throw "Autounattend XML file not found: $AutounattendXml"
    }
    Write-ColorOutput "Input files validated" "Green"
    
    if (Test-Path $resolvedOutputIso) {
        Write-ColorOutput "Output ISO already exists. Removing..." "Yellow"
        Remove-Item $resolvedOutputIso -Force
    }
    
    # Step 1: Extract ISO contents
    Write-ColorOutput "=== Step 1: Extracting ISO Contents ===" "Cyan"
    Extract-IsoContents -IsoPath $resolvedInputIso -ExtractPath $WorkingDirectory
    
    # Step 2: Add autounattend.xml and OEM directory
    Write-ColorOutput "=== Step 2: Adding Configuration Files ===" "Cyan"
    Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    Add-OemDirectory -ExtractPath $WorkingDirectory -OemSourcePath $OemDirectory
    
    # Step 3: Analyze all WIM files and add VirtIO drivers per-WIM
    if ($IncludeVirtioDrivers) {
        Write-ColorOutput "=== Step 3: Processing WIM Files and Adding VirtIO Drivers ===" "Cyan"
        $allWimInfos = Get-AllWimInfo -ExtractPath $WorkingDirectory
        Add-VirtioDrivers -ExtractPath $WorkingDirectory -VirtioVersion $VirtioVersion -VirtioCacheDirectory $VirtioCacheDirectory -WimInfos $allWimInfos
    } else {
        Write-ColorOutput "=== Step 3: Skipping VirtIO Drivers (not requested) ===" "Cyan"
    }
    
    # Step 4: Create new ISO
    Write-ColorOutput "=== Step 4: Creating New ISO ===" "Cyan"
    New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $resolvedOutputIso -OscdimgPath $script:oscdimgPath
    
    if (Test-Path $resolvedOutputIso) {
        $fileSize = (Get-Item $resolvedOutputIso).Length
        $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
        Write-ColorOutput "Output ISO created successfully!" "Green"
        Write-ColorOutput "File size: $fileSizeGB GB" "Green" -Indent 1
    } else {
        throw "Output ISO was not created successfully"
    }
} catch {
    Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
    exit 1
} finally {
    Remove-WorkingDirectory -Path $WorkingDirectory
}

Write-ColorOutput "=== Script completed successfully! ===" "Green"