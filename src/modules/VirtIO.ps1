# VirtIO drivers management for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot) "Common.ps1"
. $commonPath

# Load DISM module dependency
$toolsPath = Join-Path (Split-Path $PSScriptRoot -Parent) "tools"
. (Join-Path $toolsPath "DISM.ps1")

# Load WIM module dependency
$wimPath = Join-Path (Split-Path $PSScriptRoot) "WIM.ps1"
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
        [string]$VirtioCacheDirectory
    )
    
    if (-not $IncludeVirtioDrivers) {
        Write-ColorOutput "VirtIO drivers not requested, skipping..." "Cyan"
        return
    }
    
    Write-ColorOutput "=== Adding VirtIO Drivers ===" "Cyan"
    
    try {
        # Get WIM information for all WIM files in the ISO
        Write-ColorOutput "Analyzing WIM files..." "Yellow" -Indent 1
        $WimInfos = Get-AllWimInfo -ExtractPath $ExtractPath
        
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
            Write-ColorOutput "Failed to add drivers to boot.wim (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to add VirtIO drivers to boot.wim. DISM exit code: $($result.ExitCode)"
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
            Write-ColorOutput "Failed to unmount boot.wim (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to unmount boot.wim. DISM exit code: $($result.ExitCode). WIM may be left in inconsistent state."
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
            Write-ColorOutput "Failed to add drivers to install.wim (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to add VirtIO drivers to install.wim. DISM exit code: $($result.ExitCode)"
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
            Write-ColorOutput "Failed to unmount install.wim (exit code: $($result.ExitCode))" "Red" -Indent 2
            throw "Failed to unmount install.wim. DISM exit code: $($result.ExitCode). WIM may be left in inconsistent state."
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
