# VirtIO drivers management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path $PSScriptRoot "Common.ps1"
. $commonPath

# Load DISM module dependency
$toolsPath = Join-Path $PSScriptRoot "tools"
. (Join-Path $toolsPath "DISM.ps1")

# Load WIM module dependency
$wimPath = Join-Path $PSScriptRoot "WIM.ps1"
. $wimPath

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
    
    Write-Host ""
    Write-ColorOutput "=== VirtIO Drivers Download ===" -Color "Cyan"
    
    # Create cache directory if it doesn't exist
    if (-not (Test-Path $CacheDirectory)) {
        New-Item -ItemType Directory -Path $CacheDirectory -Force | Out-Null
        Write-ColorOutput "Created cache directory: $CacheDirectory" -Color "Green" -Indent 1
    }
    
    $downloadUrl = Get-VirtioDownloadUrl -Version $Version
    $fileName = "virtio-win-$Version.iso"
    $localPath = Join-Path $CacheDirectory $fileName
    
    # Check if we already have the file
    if (Test-Path $localPath) {
        Write-ColorOutput "VirtIO drivers already cached: $localPath" -Color "Green" -Indent 1
        # Return absolute path for consistency
        return (Resolve-Path $localPath -ErrorAction Stop)
    }
    
    try {
        # Use Invoke-WebRequestWithCleanup for better progress tracking and resource cleanup
        Invoke-WebRequestWithCleanup -Uri $downloadUrl -OutFile $localPath -Description "VirtIO drivers ($Version)" -ProgressId 3
        
        # Verify the downloaded file exists and has content
        if (-not (Test-Path $localPath -PathType Leaf)) {
            throw "Downloaded VirtIO ISO file not found: $localPath"
        }
        
        $fileSize = (Get-Item $localPath).Length
        if ($fileSize -eq 0) {
            throw "Downloaded VirtIO ISO file is empty: $localPath"
        }
        
        Write-ColorOutput "VirtIO drivers downloaded" -Color "Green" -Indent 1
        Write-ColorOutput "File size: $([math]::Round($fileSize / 1MB, 2)) MB" -Color "Cyan" -Indent 2
        
        Write-Host ""
        # Return absolute path
        return (Resolve-Path $localPath -ErrorAction Stop)
    } catch {
        throw "Failed to download VirtIO drivers: $($_.Exception.Message)"
    }
}

function Extract-VirtioDrivers {
    param(
        [string]$VirtioIsoPath,
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Extracting VirtIO drivers from: $VirtioIsoPath" -Color "Yellow"
    
    Write-Host ""
    # Validate ISO file path
    if (-not (Test-Path $VirtioIsoPath -PathType Leaf)) {
        throw "VirtIO ISO file not found: $VirtioIsoPath"
    }
    
    # Get absolute path to avoid path issues
    $VirtioIsoPath = Resolve-Path $VirtioIsoPath -ErrorAction Stop
    
    # Create virtio directory in the ISO extract path
    $virtioDir = Join-Path $ExtractPath "virtio"
    if (Test-Path $virtioDir) {
        Remove-Item $virtioDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $virtioDir -Force | Out-Null
    
    $mounted = $false
    $mountResult = $null
    
    try {
        # Check if ISO is already mounted
        $existingMount = Get-DiskImage -ImagePath $VirtioIsoPath -ErrorAction SilentlyContinue
        if ($existingMount -and $existingMount.Attached) {
            Write-ColorOutput "VirtIO ISO already mounted, using existing mount" -Color "Yellow"
            $mountResult = $existingMount
            $mounted = $true
        } else {
            # Mount the VirtIO ISO
            Write-ColorOutput "Mounting VirtIO ISO..." -Color "Yellow"
            Mount-DiskImage -ImagePath $VirtioIsoPath -ErrorAction Stop | Out-Null
            $mountResult = Get-DiskImage -ImagePath $VirtioIsoPath -ErrorAction Stop
            $mounted = $true
        }
        
        # Get drive letter with better error handling
        $volume = $mountResult | Get-Volume -ErrorAction Stop
        $driveLetter = $volume.DriveLetter
        
        # Suppress any implicit output from volume operations
        $null = $volume
        
        if (-not $driveLetter) {
            throw "Failed to get drive letter for mounted VirtIO ISO"
        }
        
        $mountedPath = "${driveLetter}:\"
        Write-ColorOutput "VirtIO ISO mounted at: $mountedPath" -Color "Green"
        
        # Verify the mounted path is accessible
        if (-not (Test-Path $mountedPath)) {
            throw "Mounted VirtIO ISO path is not accessible: $mountedPath"
        }
        
        try {
            # Copy all contents from VirtIO ISO to the virtio directory
            Write-ColorOutput "Copying VirtIO drivers to: $virtioDir" -Color "Yellow"
            $robocopyResult = robocopy $mountedPath $virtioDir /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
            
            if ($LASTEXITCODE -gt 7) {
                throw "Failed to copy VirtIO drivers. Robocopy exit code: $LASTEXITCODE"
            }
            
            Write-ColorOutput "VirtIO drivers extracted" -Color "Green"
            
            Write-Host ""
        } finally {
            # Only dismount if we mounted it ourselves
            if ($mounted -and $mountResult) {
                try {
                    Write-ColorOutput "Dismounting VirtIO ISO..." -Color "Yellow"
                    Dismount-DiskImage -ImagePath $VirtioIsoPath -ErrorAction Stop | Out-Null
                    Write-ColorOutput "VirtIO ISO dismounted" -Color "Green"
                } catch {
                    Write-ColorOutput "Warning: Failed to dismount VirtIO ISO: $($_.Exception.Message)" -Color "Yellow"
                }
            }
        }
        
        return $virtioDir
    } catch {
        # Cleanup on error
        if ($mounted -and $mountResult) {
            try {
                Write-ColorOutput "Cleaning up failed mount..." -Color "Yellow"
                Dismount-DiskImage -ImagePath $VirtioIsoPath -ErrorAction SilentlyContinue | Out-Null
            } catch {
                # Ignore cleanup errors
            }
        }
        throw "Failed to extract VirtIO drivers: $($_.Exception.Message)"
    }
}

function Get-WindowsVersion {
    param(
        [hashtable]$WimInfo,
        [array]$AllWimInfo = @()
    )
    
    # For boot.wim files, determine version from the first install.wim image
    if ($WimInfo.Type -eq "boot") {
        # Find the first install.wim image
        $firstInstallImage = $AllWimInfo | Where-Object { $_.Type -eq "install" } | Select-Object -First 1
        
        if (-not $firstInstallImage) {
            throw "Cannot determine Windows version for boot.wim: No install.wim images found to reference"
        }
        
        # Get the Windows version from the first install image
        $installVersion = Get-WindowsVersion -WimInfo $firstInstallImage -AllWimInfo $AllWimInfo
        Write-ColorOutput "Boot.wim using Windows version from first install image: $installVersion" -Color "Cyan" -Indent 2
        return $installVersion
    }
    
    # For install.wim files, infer Windows version from image name
    $imageName = $WimInfo.Name
    if ($imageName -match "Windows 11" -or $imageName -match "Windows 1[1-9]") {
        return "11"
    } elseif ($imageName -match "Windows 10") {
        return "10"
    } elseif ($imageName -match "Windows 8\.1") {
        return "8.1"
    } elseif ($imageName -match "Windows 8") {
        return "8"
    } elseif ($imageName -match "Windows 7") {
        return "7"
    } elseif ($imageName -match "Windows Vista") {
        return "vista"
    } elseif ($imageName -match "Windows XP") {
        return "xp"
    } else {
        throw "Unable to detect Windows version from image name: '$imageName'. Expected Windows XP, Vista, 7, 8, 8.1, 10, or 11 (or higher)."
    }
}

function Get-VirtioDriverVersion {
    param(
        [string]$WindowsVersion
    )
    
    # Map Windows version to VirtIO driver directory format
    $versionMap = @{
        "11" = "w11"
        "10" = "w10"
        "8.1" = "w8.1"
        "8" = "w8"
        "7" = "w7"
    }
    
    if ($versionMap.ContainsKey($WindowsVersion)) {
        return $versionMap[$WindowsVersion]
    } else {
        throw "Unknown Windows version for VirtIO driver mapping: '$WindowsVersion'"
    }
}

function Add-VirtioDriversToWim {
    param(
        [hashtable]$WimInfo,
        [string]$VirtioDir,
        [string]$VirtioVersion,
        [array]$AllWimInfo = @()
    )
    
    Write-ColorOutput "WimInfo:`n$($WimInfo | Out-String)" -Color "Gray" -Indent 1
    
    Write-ColorOutput "AllWimInfo:`n$($AllWimInfo | Out-String)" -Color "Gray" -Indent 1
    
    $arch = $WimInfo.Architecture
    $windowsVersion = Get-WindowsVersion -WimInfo $WimInfo -AllWimInfo $AllWimInfo
    $version = Get-VirtioDriverVersion -WindowsVersion $windowsVersion
    $wimPath = $WimInfo.Path
    $wimType = $WimInfo.Type
    $imageIndex = $WimInfo.Index
    $imageName = $WimInfo.Name
    
    Write-ColorOutput "Processing $wimType image: $imageName" -Color "Yellow" -Indent 1     
    
    # Define the VirtIO driver components we want to inject
    $driverComponents = @("viostor", "vioscsi", "NetKVM")
    
    Write-ColorOutput "Adding VirtIO drivers (Arch: $arch, Windows: $windowsVersion -> VirtIO: $version, Components: $($driverComponents -join ', '))" -Color "Green" -Indent 2     
    
    # Install each driver component individually
    foreach ($component in $driverComponents) {
        $componentPath = Join-Path $VirtioDir $component
        if (-not (Test-Path $componentPath)) {
            throw "Component '$component' not found in VirtIO directory at: $componentPath"
        }
        
        # For each component, find the appropriate version directory
        $versionPath = Join-Path $componentPath $version
        if (-not (Test-Path $versionPath)) {
            throw "Version '$version' not found for component '$component' at: $versionPath"
        }
        
        # For each version, find the appropriate architecture directory
        $archPath = Join-Path $versionPath $arch
        if (-not (Test-Path $archPath)) {
            throw "Architecture '$arch' not found for component '$component' version '$version' at: $archPath"
        }
        
        Write-ColorOutput "Installing $component drivers from: $archPath" -Color "Cyan" -Indent 2
        
        try {
            if ($wimType -eq "boot") {
                Inject-VirtioDriversIntoBootWim -WimPath $wimPath -DriverPaths @($archPath) -ImageIndex $imageIndex
            } else {
                Inject-VirtioDriversIntoInstallWim -WimPath $wimPath -DriverPaths @($archPath) -ImageIndex $imageIndex
            }
            Write-ColorOutput "Successfully installed $component drivers" -Color "Green" -Indent 2
        } catch {
            throw "Failed to install $component drivers: $($_.Exception.Message)"
        }
    }
}

function Add-VirtioDrivers {
    param(
        [string]$ExtractPath,
        [string]$VirtioVersion,
        [string]$VirtioCacheDirectory
    )
    
    try {
        # Get WIM information for all WIM files in the ISO
        Write-Host ""
        Write-ColorOutput "Analyzing WIM files..." -Color "Yellow" -Indent 1
        $WimInfos = Get-AllWimInfo -ExtractPath $ExtractPath
        Write-ColorOutput "Processing all WIM images ($($WimInfos.Count) total)" -Color "Green" -Indent 1
        
        # Download VirtIO drivers
        $virtioIsoPath = Get-VirtioDrivers -Version $VirtioVersion -CacheDirectory $VirtioCacheDirectory
        
        # Extract VirtIO drivers to a temporary directory
        $virtioDir = Extract-VirtioDrivers -VirtioIsoPath $virtioIsoPath -ExtractPath $ExtractPath
        
        Write-ColorOutput "VirtIO drivers extracted to: $virtioDir" -Color "Green" -Indent 1         
        # Process each WIM image individually
        foreach ($wimInfo in $WimInfos) {
            Add-VirtioDriversToWim -WimInfo $wimInfo -VirtioDir $virtioDir -VirtioVersion $VirtioVersion -AllWimInfo $WimInfos
        }
        
    } catch {
        throw "Failed to add VirtIO drivers: $($_.Exception.Message)"
    }
}

function Inject-VirtioDriversIntoBootWim {
    param(
        [string]$WimPath,
        [array]$DriverPaths,
        [int]$ImageIndex
    )
    
    Write-Host ""
    Write-ColorOutput "=== Injecting VirtIO Drivers into boot.wim (Index: $ImageIndex) ===" -Color "Cyan" -Indent 1     
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
        
        # Use the provided driver paths
        Write-ColorOutput "Using drivers from $($DriverPaths.Count) component(s):" -Color "Green" -Indent 2
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "  - $driverPath" -Color "Cyan" -Indent 2
        }         
        # Check for administrator privileges
        Write-ColorOutput "Checking administrator privileges..." -Color "Cyan" -Indent 2
        Assert-Administrator -ErrorMessage "Administrator privileges are required to mount and modify WIM files. Please run PowerShell as Administrator."
        Write-ColorOutput "Administrator privileges confirmed" -Color "Green" -Indent 2
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Check WIM file permissions
        $wimFileInfo = Get-Item $WimPath -ErrorAction Stop
        Write-ColorOutput "WIM file: $($wimFileInfo.FullName)" -Color "Cyan" -Indent 2
        Write-ColorOutput "WIM file size: $([math]::Round($wimFileInfo.Length / 1GB, 2)) GB" -Color "Cyan" -Indent 2
        Write-ColorOutput "WIM file read-only: $($wimFileInfo.IsReadOnly)" -Color "Cyan" -Indent 2
        
        # Remove read-only attribute from WIM file if present
        if ($wimFileInfo.IsReadOnly) {
            Write-ColorOutput "Removing read-only attribute from WIM file..." -Color "Yellow" -Indent 2
            Set-ItemProperty -Path $WimPath -Name IsReadOnly -Value $false
            Write-ColorOutput "Read-only attribute removed" -Color "Green" -Indent 2
        }
        
        # Check mount directory permissions
        Write-ColorOutput "Mount directory: $mountDir" -Color "Cyan" -Indent 2
        if (Test-Path $mountDir) {
            $mountDirInfo = Get-Item $mountDir
            Write-ColorOutput "Mount directory exists: $($mountDirInfo.Exists)" -Color "Cyan" -Indent 2
        }
        
        # Mount the specific boot.wim index
        Write-ColorOutput "Mounting boot.wim index $ImageIndex..." -Color "Yellow" -Indent 2
        # Mount the WIM file using direct execution
        & $dismPath /Mount-Wim /WimFile:$WimPath /Index:$ImageIndex /MountDir:$mountDir
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount boot.wim index $ImageIndex (exit code: $LASTEXITCODE)"
        }
        
        Write-ColorOutput "Successfully mounted boot.wim index $ImageIndex" -Color "Green" -Indent 2         
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to boot.wim..." -Color "Yellow" -Indent 2
        
        # Add each driver component
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "Adding drivers from: $driverPath" -Color "Cyan" -Indent 2
            & $dismPath /Image:$mountDir /Add-Driver /Driver:$driverPath /Recurse
            
            if ($LASTEXITCODE -eq 0) {
                Write-ColorOutput "Successfully added drivers from: $driverPath" -Color "Green" -Indent 2
            } else {
                throw "Failed to add drivers from $driverPath (exit code: $LASTEXITCODE)"
            }
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting boot.wim..." -Color "Yellow" -Indent 2
        # Unmount and commit changes using direct execution
        & $dismPath /Unmount-Wim /MountDir:$mountDir /Commit
        
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "Successfully unmounted and committed boot.wim changes" -Color "Green" -Indent 2
        } else {
            throw "Failed to unmount boot.wim (exit code: $LASTEXITCODE). WIM may be left in inconsistent state."
        }
        
    } catch {
        throw "Error injecting drivers into boot.wim: $($_.Exception.Message)"
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" -Color "Yellow" -Indent 2
            }
        }
    }
}

function Inject-VirtioDriversIntoInstallWim {
    param(
        [string]$WimPath,
        [array]$DriverPaths,
        [int]$ImageIndex
    )
    
    Write-Host ""
    Write-ColorOutput "=== Injecting VirtIO Drivers into install.wim (Index: $ImageIndex) ===" -Color "Cyan" -Indent 1     
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "install_mount_$ImageIndex"
    
    if (-not (Test-Path $WimPath)) {
        Write-ColorOutput "Warning: install.wim not found at: $WimPath" -Color "Yellow" -Indent 2
        return
    }
    
    try {
        # Create mount directory
        if (Test-Path $mountDir) {
            Remove-Item $mountDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
        
        # Use the provided driver paths
        Write-ColorOutput "Using drivers from $($DriverPaths.Count) component(s):" -Color "Green" -Indent 2
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "  - $driverPath" -Color "Cyan" -Indent 2
        }         
        # Check for administrator privileges
        Write-ColorOutput "Checking administrator privileges..." -Color "Cyan" -Indent 2
        Assert-Administrator -ErrorMessage "Administrator privileges are required to mount and modify WIM files. Please run PowerShell as Administrator."
        Write-ColorOutput "Administrator privileges confirmed" -Color "Green" -Indent 2
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Check WIM file permissions
        $wimFileInfo = Get-Item $WimPath -ErrorAction Stop
        Write-ColorOutput "WIM file: $($wimFileInfo.FullName)" -Color "Cyan" -Indent 2
        Write-ColorOutput "WIM file size: $([math]::Round($wimFileInfo.Length / 1GB, 2)) GB" -Color "Cyan" -Indent 2
        Write-ColorOutput "WIM file read-only: $($wimFileInfo.IsReadOnly)" -Color "Cyan" -Indent 2
        
        # Remove read-only attribute if present
        if ($wimFileInfo.IsReadOnly) {
            Write-ColorOutput "Removing read-only attribute from WIM file..." -Color "Yellow" -Indent 2
            Set-ItemProperty -Path $WimPath -Name IsReadOnly -Value $false
            Write-ColorOutput "Read-only attribute removed" -Color "Green" -Indent 2
        }
        
        # Check mount directory permissions
        Write-ColorOutput "Mount directory: $mountDir" -Color "Cyan" -Indent 2
        if (Test-Path $mountDir) {
            $mountDirInfo = Get-Item $mountDir
            Write-ColorOutput "Mount directory exists: $($mountDirInfo.Exists)" -Color "Cyan" -Indent 2
        }
        
        # Mount the specific install.wim index
        Write-ColorOutput "Mounting install.wim index $ImageIndex..." -Color "Yellow" -Indent 2
        # Mount the WIM file using direct execution
        & $dismPath /Mount-Wim /WimFile:$WimPath /Index:$ImageIndex /MountDir:$mountDir
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to mount install.wim index $ImageIndex (exit code: $LASTEXITCODE)"
        }
        
        Write-ColorOutput "Successfully mounted install.wim index $ImageIndex" -Color "Green" -Indent 2         
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to install.wim..." -Color "Yellow" -Indent 2
        
        # Add each driver component
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "Adding drivers from: $driverPath" -Color "Cyan" -Indent 2
            & $dismPath /Image:$mountDir /Add-Driver /Driver:$driverPath /Recurse
            
            if ($LASTEXITCODE -eq 0) {
                Write-ColorOutput "Successfully added drivers from: $driverPath" -Color "Green" -Indent 2
            } else {
                throw "Failed to add drivers from $driverPath (exit code: $LASTEXITCODE)"
            }
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting install.wim..." -Color "Yellow" -Indent 2
        # Unmount and commit changes using direct execution
        & $dismPath /Unmount-Wim /MountDir:$mountDir /Commit
        
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "Successfully unmounted and committed install.wim changes" -Color "Green" -Indent 2
        } else {
            throw "Failed to unmount install.wim (exit code: $LASTEXITCODE). WIM may be left in inconsistent state."
        }
        
    } catch {
        throw "Error injecting drivers into install.wim: $($_.Exception.Message)"
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" -Color "Yellow" -Indent 2
            }
        }
    }
}
