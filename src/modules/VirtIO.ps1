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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("stable", "latest")]
        [string]$Version,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 40)]
        [int]$Indent = 0
    )
    
    $urls = @{
        "stable" = "https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso"
        "latest" = "https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/latest-virtio/virtio-win.iso"
    }
    
    $urls = Assert-Defined -VariableName "urls" -Value $urls -ErrorMessage "URL mapping hashtable is not defined"
    
    if ($urls.ContainsKey($Version)) {
        $downloadUrl = $urls[$Version]
        $downloadUrl = Assert-NotEmpty -VariableName "urls[$Version]" -Value $downloadUrl -ErrorMessage "Download URL for version '$Version' is empty"
        return $downloadUrl
    } else {
        throw "Unknown VirtIO version: '$Version'. Supported versions are: stable, latest"
    }
}


function Extract-VirtioDrivers {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("stable", "latest")]
        [string]$Version,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            # Validate that the path is a valid path format, but allow non-existent directories
            try {
                $null = [System.IO.Path]::GetFullPath($_)
                $true
            } catch {
                throw "Extract path is not a valid path format: $_"
            }
        })]
        [string]$ExtractPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 40)]
        [int]$Indent = 0
    )
    
    Write-Host ""
    Write-ColorOutput "=== VirtIO Drivers Download and Extract ===" -Color "Cyan" -Indent $Indent
    
    # Set up cache directory
    $cacheDir = Join-Path $env:TEMP "virtio-cache"
    if (-not (Test-Path $cacheDir)) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
    }
    
    # Use cached ISO if available
    $cachedIsoPath = Join-Path $cacheDir "$Version.iso"
    
    if (Test-Path $cachedIsoPath) {
        Write-ColorOutput "Using cached VirtIO drivers ($Version) from: $cachedIsoPath" -Color "Green" -Indent ($Indent + 1)
    } else {
        $downloadUrl = Get-VirtioDownloadUrl -Version $Version
        $downloadUrl = Assert-NotEmpty -VariableName "downloadUrl" -Value $downloadUrl -ErrorMessage "Failed to get download URL for version: $Version"
        
        Write-ColorOutput "Downloading VirtIO drivers ($Version)..." -Color "Yellow" -Indent ($Indent + 1)
        Invoke-WebRequestWithCleanup -Uri $downloadUrl -OutFile $cachedIsoPath -Description "VirtIO drivers ($Version)" -ProgressId 3
        
        # Verify the downloaded file exists and has content
        if (-not (Test-Path $cachedIsoPath -PathType Leaf)) {
            throw "Downloaded VirtIO ISO file not found: $cachedIsoPath"
        }
        
        $fileSize = (Get-Item $cachedIsoPath).Length
        if ($fileSize -eq 0) {
            throw "Downloaded VirtIO ISO file is empty: $cachedIsoPath"
        }
        
        Write-ColorOutput "VirtIO drivers downloaded and cached" -Color "Green" -Indent ($Indent + 1)
        Write-ColorOutput "File size: $([math]::Round($fileSize / 1MB, 2)) MB" -Color "Cyan" -Indent ($Indent + 2)
    }
    
    # Copy cached ISO to working directory for extraction
    $tempIsoPath = Join-Path $ExtractPath "virtio-$Version.iso"
    Copy-Item $cachedIsoPath $tempIsoPath -Force
    
    try {
        # Create virtio directory in the extract path
        $virtioDir = Join-Path $ExtractPath "virtio"
        if (Test-Path $virtioDir) {
            Remove-Item $virtioDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $virtioDir -Force | Out-Null
        
        $mounted = $false
        $mountResult = $null
        
        try {
            # Mount the VirtIO ISO
            Write-ColorOutput "Mounting VirtIO ISO..." -Color "Yellow" -Indent ($Indent + 1)
            Mount-DiskImage -ImagePath $tempIsoPath -ErrorAction Stop | Out-Null
            $mountResult = Get-DiskImage -ImagePath $tempIsoPath -ErrorAction Stop
            $mounted = $true
            
            # Get drive letter with better error handling
            $mountResult = Assert-Defined -VariableName "mountResult" -Value $mountResult -ErrorMessage "Failed to mount VirtIO ISO"
            
            $volume = $mountResult | Get-Volume -ErrorAction Stop
            $volume = Assert-Defined -VariableName "volume" -Value $volume -ErrorMessage "Failed to get volume information from mounted VirtIO ISO"
            
            $driveLetter = $volume.DriveLetter
            $driveLetter = Assert-NotEmpty -VariableName "driveLetter" -Value $driveLetter -ErrorMessage "Failed to get drive letter for mounted VirtIO ISO"
            
            # Suppress any implicit output from volume operations
            $null = $volume
            
            $mountedPath = "${driveLetter}:\"
            $mountedPath = Assert-ValidPath -VariableName "mountedPath" -Path $mountedPath -ErrorMessage "Generated mounted path is invalid: $mountedPath"
            Write-ColorOutput "VirtIO ISO mounted at: $mountedPath" -Color "Green" -Indent ($Indent + 1)
            
            # Copy VirtIO drivers from mounted ISO to extract directory
            try {
                Write-ColorOutput "Copying VirtIO drivers to: $virtioDir" -Color "Yellow" -Indent ($Indent + 1)
                # Use robocopy for better performance and error handling
                $robocopyArgs = @($mountedPath, $virtioDir, "/E", "/R:3", "/W:1", "/NFL", "/NDL", "/NJH", "/NJS", "/nc", "/ns", "/np")
                $robocopyResult = & robocopy @robocopyArgs
                $robocopyExitCode = $LASTEXITCODE
                
                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($robocopyExitCode -ge 8) {
                    throw "Robocopy failed with exit code: $robocopyExitCode"
                }
                
                Write-ColorOutput "VirtIO drivers extracted" -Color "Green" -Indent ($Indent + 1)
            } catch {
                throw "Failed to copy VirtIO drivers: $($_.Exception.Message)"
            } finally {
                # Always try to dismount the ISO
                if ($mounted) {
                    try {
                        Write-ColorOutput "Dismounting VirtIO ISO..." -Color "Yellow" -Indent ($Indent + 1)
                        Dismount-DiskImage -ImagePath $tempIsoPath -ErrorAction Stop | Out-Null
                        Write-ColorOutput "VirtIO ISO dismounted" -Color "Green" -Indent ($Indent + 1)
                    } catch {
                        Write-ColorOutput "Warning: Failed to dismount VirtIO ISO: $($_.Exception.Message)" -Color "Yellow" -Indent ($Indent + 1)
                    }
                }
            }
            
        } catch {
            # Cleanup on error
            if ($mounted) {
                try {
                    Write-ColorOutput "Cleaning up failed mount..." -Color "Yellow" -Indent ($Indent + 1)
                    Dismount-DiskImage -ImagePath $tempIsoPath -ErrorAction SilentlyContinue | Out-Null
                } catch {
                    # Ignore cleanup errors
                }
            }
            throw "Failed to extract VirtIO drivers: $($_.Exception.Message)"
        } finally {
            # Clean up temporary ISO file
            if (Test-Path $tempIsoPath) {
                try {
                    Remove-Item $tempIsoPath -Force -ErrorAction SilentlyContinue
                } catch {
                    # Ignore cleanup errors
                }
            }
        }
        
        return $virtioDir
        
    } catch {
        # Clean up temporary ISO file on download error
        if (Test-Path $tempIsoPath) {
            try {
                Remove-Item $tempIsoPath -Force -ErrorAction SilentlyContinue
            } catch {
                # Ignore cleanup errors
            }
        }
        throw "Failed to download and extract VirtIO drivers: $($_.Exception.Message)"
    }
}

function Get-WindowsVersion {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [hashtable]$WimInfo,
        
        [Parameter(Mandatory = $false)]
        [array]$AllWimInfo = @(),
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 40)]
        [int]$Indent = 0
    )
    
    # For boot.wim files, determine version from the first install.wim image
    if ($WimInfo.Type -eq "boot") {
        # Find the first install.wim image
        $firstInstallImage = $AllWimInfo | Where-Object { $_.Type -eq "install" } | Select-Object -First 1
        
        $firstInstallImage = Assert-Defined -VariableName "firstInstallImage" -Value $firstInstallImage -ErrorMessage "Cannot determine Windows version for boot.wim: No install.wim images found to reference"
        
        # Get the Windows version from the first install image
        $installVersion = Get-WindowsVersion -WimInfo $firstInstallImage -AllWimInfo $AllWimInfo -Indent $Indent
        Write-ColorOutput "Boot.wim using Windows version from first install image: $installVersion" -Color "Cyan" -Indent ($Indent + 1)
        return $installVersion
    }
    
    # For install.wim files, infer Windows version from image name
    $imageName = Assert-NotEmpty -VariableName "WimInfo.Name" -Value $WimInfo.Name -ErrorMessage "WIM image name is not defined"
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

function Get-VirtioArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowsArchitecture,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 40)]
        [int]$Indent = 0
    )
    
    # Map Windows architecture names to VirtIO architecture names
    $archMap = @{
        "x64" = "amd64"
        "AMD64" = "amd64"
        "x86" = "i386"
        "i386" = "i386"
        "ia64" = "ia64"
        "arm64" = "arm64"
        "aarch64" = "arm64"
    }
    
    $archMap = Assert-Defined -VariableName "archMap" -Value $archMap -ErrorMessage "VirtIO architecture mapping is not defined"
    
    if ($archMap.ContainsKey($WindowsArchitecture)) {
        $mappedArch = $archMap[$WindowsArchitecture]
        $mappedArch = Assert-NotEmpty -VariableName "archMap[$WindowsArchitecture]" -Value $mappedArch -ErrorMessage "Mapped architecture for '$WindowsArchitecture' is empty"
        return $mappedArch
    } else {
        throw "Unknown Windows architecture for VirtIO driver mapping: '$WindowsArchitecture'. Supported architectures: $($archMap.Keys -join ', ')"
    }
}

function Get-VirtioDriverVersion {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowsVersion,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 40)]
        [int]$Indent = 0
    )
    
    # Map Windows version to VirtIO driver directory format
    $versionMap = @{
        "11" = "w11"
        "10" = "w10"
        "8.1" = "w8.1"
        "8" = "w8"
        "7" = "w7"
    }
    
    $versionMap = Assert-Defined -VariableName "versionMap" -Value $versionMap -ErrorMessage "Version mapping hashtable is not defined"
    
    if ($versionMap.ContainsKey($WindowsVersion)) {
        $mappedVersion = $versionMap[$WindowsVersion]
        $mappedVersion = Assert-NotEmpty -VariableName "versionMap[$WindowsVersion]" -Value $mappedVersion -ErrorMessage "Mapped version for '$WindowsVersion' is empty"
        return $mappedVersion
    } else {
        throw "Unknown Windows version for VirtIO driver mapping: '$WindowsVersion'"
    }
}

function Add-VirtioDriversToWim {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [hashtable]$WimInfo,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "VirtIO directory does not exist: $_"
            }
            $true
        })]
        [string]$VirtioDir,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VirtioVersion,
        
        [Parameter(Mandatory = $false)]
        [array]$AllWimInfo = @(),
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$Indent = 0
    )
    
    # Validate and extract WIM info properties
    $windowsArch = Assert-NotEmpty -VariableName "WimInfo.Architecture" -Value $WimInfo.Architecture -ErrorMessage "WIM image architecture is not defined"
    $wimPath = Assert-NotEmpty -VariableName "WimInfo.Path" -Value $WimInfo.Path -ErrorMessage "WIM image path is not defined"
    $wimType = Assert-NotEmpty -VariableName "WimInfo.Type" -Value $WimInfo.Type -ErrorMessage "WIM image type is not defined"
    $imageIndex = Assert-PositiveNumber -VariableName "WimInfo.Index" -Value ([int]$WimInfo.Index) -ErrorMessage "WIM image index must be a positive number"
    $imageName = Assert-NotEmpty -VariableName "WimInfo.Name" -Value $WimInfo.Name -ErrorMessage "WIM image name is not defined"
    
    # Map Windows architecture to VirtIO architecture
    $arch = Get-VirtioArchitecture -WindowsArchitecture $windowsArch -Indent $Indent
    
    $windowsVersion = Get-WindowsVersion -WimInfo $WimInfo -AllWimInfo $AllWimInfo -Indent $Indent
    $version = Get-VirtioDriverVersion -WindowsVersion $windowsVersion -Indent $Indent
    
    Write-ColorOutput "Processing $wimType image: $imageName" -Color "Yellow" -Indent $Indent
    
    # Define the VirtIO driver components we want to inject
    $driverComponents = @("viostor", "vioscsi", "NetKVM")
    $driverComponents = Assert-ArrayNotEmpty -VariableName "driverComponents" -Value $driverComponents -ErrorMessage "Driver components array is empty"
    
    Write-ColorOutput "Adding VirtIO drivers (Arch: $windowsArch -> $arch, Windows: $windowsVersion -> VirtIO: $version, Components: $($driverComponents -join ', '))" -Color "Green" -Indent ($Indent + 2)     
    
    # Collect all driver paths
    $allDriverPaths = @()
    foreach ($component in $driverComponents) {
        $component = Assert-NotEmpty -VariableName "component" -Value $component -ErrorMessage "Component name is empty"
        
        $componentPath = Join-Path $VirtioDir $component
        $componentPath = Assert-ValidPath -VariableName "componentPath" -Path $componentPath -ErrorMessage "Generated component path is invalid: $componentPath"
        
        if (-not (Test-Path $componentPath)) {
            throw "Component '$component' not found in VirtIO directory at: $componentPath"
        }
        
        # For each component, find the appropriate version directory
        $versionPath = Join-Path $componentPath $version
        $versionPath = Assert-ValidPath -VariableName "versionPath" -Path $versionPath -ErrorMessage "Generated version path is invalid: $versionPath"
        
        if (-not (Test-Path $versionPath)) {
            throw "Version '$version' not found for component '$component' at: $versionPath"
        }
        
        # For each version, find the appropriate architecture directory
        $archPath = Join-Path $versionPath $arch
        $archPath = Assert-ValidPath -VariableName "archPath" -Path $archPath -ErrorMessage "Generated architecture path is invalid: $archPath"
        
        if (-not (Test-Path $archPath)) {
            throw "Architecture '$arch' not found for component '$component' version '$version' at: $archPath"
        }
        
        Write-ColorOutput "Found $component drivers at: $archPath" -Color "Cyan" -Indent ($Indent + 2)
        $allDriverPaths += $archPath
    }
    
    # Install all driver components at once
    try {
        Invoke-VirtioDriverInjection -WimPath $wimPath -DriverPaths $allDriverPaths -ImageIndex $imageIndex -WimType $wimType -Indent ($Indent + 2)
        Write-ColorOutput "Successfully installed all VirtIO drivers" -Color "Green" -Indent ($Indent + 2)
    } catch {
        throw "Failed to install VirtIO drivers: $($_.Exception.Message)"
    }
}

function Add-VirtioDrivers {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "Extract path does not exist: $_"
            }
            $true
        })]
        [string]$ExtractPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VirtioVersion,
        
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$Indent = 0
    )
    
    try {
        # Get WIM information for all WIM files in the ISO
        Write-Host ""
        Write-ColorOutput "Analyzing WIM files..." -Color "Yellow" -Indent $Indent
        $WimInfos = Get-AllWimInfo -ExtractPath $ExtractPath -Indent ($Indent + 1)
        Write-ColorOutput "Processing all WIM images ($($WimInfos.Count) total)" -Color "Green" -Indent $Indent
        
        # Download and extract VirtIO drivers
        $virtioDir = Extract-VirtioDrivers -Version $VirtioVersion -ExtractPath $ExtractPath -Indent ($Indent + 1)
        
        Write-ColorOutput "VirtIO drivers extracted to: $virtioDir" -Color "Green" -Indent ($Indent + 1)         
        # Process each WIM image individually
        foreach ($wimInfo in $WimInfos) {
            Add-VirtioDriversToWim -WimInfo $wimInfo -VirtioDir $virtioDir -VirtioVersion $VirtioVersion -AllWimInfo $WimInfos -Indent ($Indent + 1)
        }
        
    } catch {
        throw "Failed to add VirtIO drivers: $($_.Exception.Message)"
    }
}

function Invoke-VirtioDriverInjection {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "WIM path is invalid: $_"
            }
        })]
        [string]$WimPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [ValidateCount(1, [int]::MaxValue)]
        [array]$DriverPaths,
        
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ImageIndex,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("boot", "install")]
        [string]$WimType,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$Indent = 0
    )
    
    Write-Host ""
    Write-ColorOutput "=== Injecting VirtIO Drivers into $WimType.wim (Index: $ImageIndex) ===" -Color "Cyan" -Indent $Indent     
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "${WimType}_mount_$ImageIndex"
    
    # Validate WIM file exists
    $WimPath = Assert-FileExists -FilePath $WimPath -ErrorMessage "$WimType.wim not found at: $WimPath"
    
    try {
        # Create mount directory
        if (Test-Path $mountDir) {
            Remove-Item $mountDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
        
        # Use the provided driver paths
        Write-ColorOutput "Using drivers from $($DriverPaths.Count) component(s):" -Color "Green" -Indent ($Indent + 2)
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "  - $driverPath" -Color "Cyan" -Indent ($Indent + 2)
        }         
        # Check for administrator privileges
        Write-ColorOutput "Checking administrator privileges..." -Color "Cyan" -Indent ($Indent + 2)
        Assert-Administrator -ErrorMessage "Administrator privileges are required to mount and modify WIM files. Please run PowerShell as Administrator."
        Write-ColorOutput "Administrator privileges confirmed" -Color "Green" -Indent ($Indent + 2)
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Check WIM file permissions
        $wimFileInfo = Get-Item $WimPath -ErrorAction Stop
        Write-ColorOutput "WIM file: $($wimFileInfo.FullName)" -Color "Cyan" -Indent ($Indent + 2)
        Write-ColorOutput "WIM file size: $([math]::Round($wimFileInfo.Length / 1GB, 2)) GB" -Color "Cyan" -Indent ($Indent + 2)
        Write-ColorOutput "WIM file read-only: $($wimFileInfo.IsReadOnly)" -Color "Cyan" -Indent ($Indent + 2)
        
        # Remove read-only attribute from WIM file if present
        if ($wimFileInfo.IsReadOnly) {
            Write-ColorOutput "Removing read-only attribute from WIM file..." -Color "Yellow" -Indent ($Indent + 2)
            Set-ItemProperty -Path $WimPath -Name IsReadOnly -Value $false
            Write-ColorOutput "Read-only attribute removed" -Color "Green" -Indent ($Indent + 2)
        }
        
        # Check mount directory permissions
        Write-ColorOutput "Mount directory: $mountDir" -Color "Cyan" -Indent ($Indent + 2)
        if (Test-Path $mountDir) {
            $mountDirInfo = Get-Item $mountDir
            Write-ColorOutput "Mount directory exists: $($mountDirInfo.Exists)" -Color "Cyan" -Indent ($Indent + 2)
        }
        
        # Mount the specific WIM index
        Write-ColorOutput "Mounting $WimType.wim index $ImageIndex..." -Color "Yellow" -Indent ($Indent + 2)
        # Mount the WIM file using direct execution
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Mount-Wim", "/WimFile:$WimPath", "/Index:$ImageIndex", "/MountDir:$mountDir") -Description "Mount $WimType.wim index $ImageIndex" -SuppressOutput -Indent ($Indent + 1)
        
        Write-ColorOutput "Successfully mounted $WimType.wim index $ImageIndex" -Color "Green" -Indent ($Indent + 2)         
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to $WimType.wim..." -Color "Yellow" -Indent ($Indent + 2)
        
        # Add each driver component
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "Adding drivers from: $driverPath" -Color "Cyan" -Indent ($Indent + 2)
            Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Image:$mountDir", "/Add-Driver", "/Driver:$driverPath", "/Recurse") -Description "Add drivers from $driverPath" -SuppressOutput -Indent ($Indent + 1)
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting $WimType.wim..." -Color "Yellow" -Indent ($Indent + 2)
        # Unmount and commit changes using direct execution
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Unmount-Wim", "/MountDir:$mountDir", "/Commit") -Description "Unmount $WimType.wim" -SuppressOutput -Indent ($Indent + 1)
        
    } catch {
        throw "Error injecting drivers into $WimType.wim: $($_.Exception.Message)"
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" -Color "Yellow" -Indent ($Indent + 2)
            }
        }
    }
}

