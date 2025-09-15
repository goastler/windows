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
    
    Write-ColorOutput "Processing $wimType image: $imageName" -Color "Yellow" -Indent 1     
    # Validate Windows 11 architecture compatibility
    if ($version -eq "w11" -and $arch -eq "x86") {
        Write-ColorOutput "Skipping Windows 11 x86 image (not supported)" -Color "Yellow" -Indent 2
        return
    }
    
    # Validate VirtIO driver availability for ARM64
    if ($arch -eq "arm64") {
        Write-ColorOutput "Skipping ARM64 image (VirtIO drivers not available)" -Color "Yellow" -Indent 2
        return
    }
    
    # Check if appropriate drivers exist
    $driverPath = Join-Path $VirtioDir $arch
    if (-not (Test-Path $driverPath)) {
        Write-ColorOutput "No VirtIO drivers found for architecture $arch" -Color "Yellow" -Indent 2
        return
    }
    
    $windowsDriverPath = Join-Path $driverPath $version
    if (-not (Test-Path $windowsDriverPath)) {
        Write-ColorOutput "No VirtIO drivers found for Windows version $version" -Color "Yellow" -Indent 2
        return
    }
    
    Write-ColorOutput "Adding VirtIO drivers (Arch: $arch, Version: $version)" -Color "Green" -Indent 2     
    try {
        if ($wimType -eq "boot") {
            Inject-VirtioDriversIntoBootWim -WimPath $wimPath -VirtioDir $VirtioDir -Arch $arch -Version $version -ImageIndex $imageIndex
        } else {
            Inject-VirtioDriversIntoInstallWim -WimPath $wimPath -VirtioDir $VirtioDir -Arch $arch -Version $version -ImageIndex $imageIndex
        }
    } catch {
        throw "Failed to add VirtIO drivers to $wimType image: $($_.Exception.Message)"
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
        
        # Download VirtIO drivers
        $virtioIsoPath = Get-VirtioDrivers -Version $VirtioVersion -CacheDirectory $VirtioCacheDirectory
        
        # Extract VirtIO drivers to a temporary directory
        $virtioDir = Extract-VirtioDrivers -VirtioIsoPath $virtioIsoPath -ExtractPath $ExtractPath
        
        Write-ColorOutput "VirtIO drivers extracted to: $virtioDir" -Color "Green" -Indent 1         
        # Log the driver structure for debugging
        $driverDirs = Get-ChildItem -Path $virtioDir | Where-Object { $_.PSIsContainer } | Select-Object -ExpandProperty Name
        Write-ColorOutput "Available driver directories: $($driverDirs -join ', ')" -Color "Cyan" -Indent 1         
        # Process each WIM image individually
        foreach ($wimInfo in $WimInfos) {
            Add-VirtioDriversToWim -WimInfo $wimInfo -VirtioDir $virtioDir -VirtioVersion $VirtioVersion
        }
        
    } catch {
        throw "Failed to add VirtIO drivers: $($_.Exception.Message)"
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
        
        # Use the specified architecture and Windows version
        $driverPath = Join-Path $VirtioDir $Arch
        
        if (-not (Test-Path $driverPath)) {
            throw "VirtIO drivers not found for architecture: $Arch at: $driverPath"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            throw "VirtIO drivers not found for Windows version: $Version at: $windowsDriverPath"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" -Color "Green" -Indent 2
        Write-ColorOutput "Architecture: $Arch" -Color "Cyan" -Indent 2
        Write-ColorOutput "Version: $Version" -Color "Cyan" -Indent 2         
        # Check for administrator privileges
        Write-ColorOutput "Checking administrator privileges..." -Color "Cyan" -Indent 2
        Assert-Administrator -ErrorMessage "Administrator privileges are required to mount and modify WIM files. Please run PowerShell as Administrator."
        Write-ColorOutput "Administrator privileges confirmed" -Color "Green" -Indent 2
        
        # Get DISM path
        $dismPath = Get-DismPath
        
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
        # Add drivers using direct execution
        & $dismPath /Image:$mountDir /Add-Driver /Driver:$windowsDriverPath /Recurse
        
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "Successfully added VirtIO drivers to boot.wim" -Color "Green" -Indent 2
        } else {
            throw "Failed to add VirtIO drivers to boot.wim (exit code: $LASTEXITCODE)"
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
        [string]$VirtioDir,
        [string]$Arch,
        [string]$Version,
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
        
        # Use the specified architecture and Windows version
        $driverPath = Join-Path $VirtioDir $Arch
        
        if (-not (Test-Path $driverPath)) {
            throw "VirtIO drivers not found for architecture: $Arch at: $driverPath"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            throw "VirtIO drivers not found for Windows version: $Version at: $windowsDriverPath"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" -Color "Green" -Indent 2
        Write-ColorOutput "Architecture: $Arch" -Color "Cyan" -Indent 2
        Write-ColorOutput "Version: $Version" -Color "Cyan" -Indent 2         
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
        # Add drivers using direct execution
        & $dismPath /Image:$mountDir /Add-Driver /Driver:$windowsDriverPath /Recurse
        
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "Successfully added VirtIO drivers to install.wim" -Color "Green" -Indent 2
        } else {
            throw "Failed to add VirtIO drivers to install.wim (exit code: $LASTEXITCODE)"
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
