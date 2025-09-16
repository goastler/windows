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
        [string]$Version
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

function Get-VirtioDrivers {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("stable", "latest")]
        [string]$Version,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "Cache directory path is invalid: $_"
            }
        })]
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
    $downloadUrl = Assert-NotEmpty -VariableName "downloadUrl" -Value $downloadUrl -ErrorMessage "Failed to get download URL for version: $Version"
    
    $fileName = "virtio-win-$Version.iso"
    $fileName = Assert-NotEmpty -VariableName "fileName" -Value $fileName -ErrorMessage "Generated filename is empty"
    
    $localPath = Join-Path $CacheDirectory $fileName
    $localPath = Assert-ValidPath -VariableName "localPath" -Path $localPath -ErrorMessage "Generated local path is invalid: $localPath"
    
    # Check if we already have the file
    if (Test-Path $localPath) {
        Write-ColorOutput "VirtIO drivers already cached: $localPath" -Color "Green" -Indent 1
        # Return absolute path for consistency
        $resolvedPath = Resolve-Path $localPath -ErrorAction Stop
        $resolvedPath = Assert-ValidPath -VariableName "resolvedPath" -Path $resolvedPath -ErrorMessage "Failed to resolve cached file path: $localPath"
        return $resolvedPath
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "VirtIO ISO path is invalid: $_"
            }
        })]
        [string]$VirtioIsoPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "Extract path is invalid: $_"
            }
        })]
        [string]$ExtractPath
    )
    
    Write-ColorOutput "Extracting VirtIO drivers from: $VirtioIsoPath" -Color "Yellow"
    
    Write-Host ""
    # Validate ISO file path
    $VirtioIsoPath = Assert-FileExists -FilePath $VirtioIsoPath -ErrorMessage "VirtIO ISO file not found: $VirtioIsoPath"
    
    # Get absolute path to avoid path issues
    $VirtioIsoPath = Resolve-Path $VirtioIsoPath -ErrorAction Stop
    $VirtioIsoPath = Assert-ValidPath -VariableName "VirtioIsoPath" -Path $VirtioIsoPath -ErrorMessage "Failed to resolve VirtIO ISO path: $VirtioIsoPath"
    
    # Create virtio directory in the ISO extract path
    $virtioDir = Join-Path $ExtractPath "virtio"
    $virtioDir = Assert-ValidPath -VariableName "virtioDir" -Path $virtioDir -ErrorMessage "Generated virtio directory path is invalid: $virtioDir"
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
        $mountResult = Assert-Defined -VariableName "mountResult" -Value $mountResult -ErrorMessage "Failed to mount VirtIO ISO"
        
        $volume = $mountResult | Get-Volume -ErrorAction Stop
        $volume = Assert-Defined -VariableName "volume" -Value $volume -ErrorMessage "Failed to get volume information from mounted VirtIO ISO"
        
        $driveLetter = $volume.DriveLetter
        $driveLetter = Assert-NotEmpty -VariableName "driveLetter" -Value $driveLetter -ErrorMessage "Failed to get drive letter for mounted VirtIO ISO"
        
        # Suppress any implicit output from volume operations
        $null = $volume
        
        $mountedPath = "${driveLetter}:\"
        $mountedPath = Assert-ValidPath -VariableName "mountedPath" -Path $mountedPath -ErrorMessage "Generated mounted path is invalid: $mountedPath"
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [hashtable]$WimInfo,
        
        [Parameter(Mandatory = $false)]
        [array]$AllWimInfo = @()
    )
    
    # For boot.wim files, determine version from the first install.wim image
    if ($WimInfo.Type -eq "boot") {
        # Find the first install.wim image
        $firstInstallImage = $AllWimInfo | Where-Object { $_.Type -eq "install" } | Select-Object -First 1
        
        $firstInstallImage = Assert-Defined -VariableName "firstInstallImage" -Value $firstInstallImage -ErrorMessage "Cannot determine Windows version for boot.wim: No install.wim images found to reference"
        
        # Get the Windows version from the first install image
        $installVersion = Get-WindowsVersion -WimInfo $firstInstallImage -AllWimInfo $AllWimInfo
        Write-ColorOutput "Boot.wim using Windows version from first install image: $installVersion" -Color "Cyan" -Indent 2
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

function Get-VirtioDriverVersion {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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
        [array]$AllWimInfo = @()
    )
    
    # Validate and extract WIM info properties
    $arch = Assert-NotEmpty -VariableName "WimInfo.Architecture" -Value $WimInfo.Architecture -ErrorMessage "WIM image architecture is not defined"
    $wimPath = Assert-NotEmpty -VariableName "WimInfo.Path" -Value $WimInfo.Path -ErrorMessage "WIM image path is not defined"
    $wimType = Assert-NotEmpty -VariableName "WimInfo.Type" -Value $WimInfo.Type -ErrorMessage "WIM image type is not defined"
    $imageIndex = Assert-PositiveNumber -VariableName "WimInfo.Index" -Value $WimInfo.Index -ErrorMessage "WIM image index must be a positive number"
    $imageName = Assert-NotEmpty -VariableName "WimInfo.Name" -Value $WimInfo.Name -ErrorMessage "WIM image name is not defined"
    
    $windowsVersion = Get-WindowsVersion -WimInfo $WimInfo -AllWimInfo $AllWimInfo
    $version = Get-VirtioDriverVersion -WindowsVersion $windowsVersion
    
    Write-ColorOutput "Processing $wimType image: $imageName" -Color "Yellow" -Indent 1     
    
    # Define the VirtIO driver components we want to inject
    $driverComponents = @("viostor", "vioscsi", "NetKVM")
    $driverComponents = Assert-ArrayNotEmpty -VariableName "driverComponents" -Value $driverComponents -ErrorMessage "Driver components array is empty"
    
    Write-ColorOutput "Adding VirtIO drivers (Arch: $arch, Windows: $windowsVersion -> VirtIO: $version, Components: $($driverComponents -join ', '))" -Color "Green" -Indent 2     
    
    # Install each driver component individually
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
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            try {
                $resolvedPath = Resolve-Path $_ -ErrorAction Stop
                $true
            } catch {
                throw "VirtIO cache directory path is invalid: $_"
            }
        })]
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
        [int]$ImageIndex
    )
    
    Write-Host ""
    Write-ColorOutput "=== Injecting VirtIO Drivers into boot.wim (Index: $ImageIndex) ===" -Color "Cyan" -Indent 1     
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "boot_mount_$ImageIndex"
    
    $WimPath = Assert-FileExists -FilePath $WimPath -ErrorMessage "boot.wim not found at: $WimPath"
    
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
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Mount-Wim", "/WimFile:$WimPath", "/Index:$ImageIndex", "/MountDir:$mountDir") -Description "Mount boot.wim index $ImageIndex" -SuppressOutput
        
        Write-ColorOutput "Successfully mounted boot.wim index $ImageIndex" -Color "Green" -Indent 2         
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to boot.wim..." -Color "Yellow" -Indent 2
        
        # Add each driver component
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "Adding drivers from: $driverPath" -Color "Cyan" -Indent 2
            Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Image:$mountDir", "/Add-Driver", "/Driver:$driverPath", "/Recurse") -Description "Add drivers from $driverPath" -SuppressOutput
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting boot.wim..." -Color "Yellow" -Indent 2
        # Unmount and commit changes using direct execution
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Unmount-Wim", "/MountDir:$mountDir", "/Commit") -Description "Unmount boot.wim" -SuppressOutput
        
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
        [int]$ImageIndex
    )
    
    Write-Host ""
    Write-ColorOutput "=== Injecting VirtIO Drivers into install.wim (Index: $ImageIndex) ===" -Color "Cyan" -Indent 1     
    $mountDir = Join-Path (Split-Path $WimPath -Parent) "install_mount_$ImageIndex"
    
    # Validate install.wim exists
    try {
        $WimPath = Assert-FileExists -FilePath $WimPath -ErrorMessage "install.wim not found at: $WimPath"
    } catch {
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
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Mount-Wim", "/WimFile:$WimPath", "/Index:$ImageIndex", "/MountDir:$mountDir") -Description "Mount install.wim index $ImageIndex" -SuppressOutput
        
        Write-ColorOutput "Successfully mounted install.wim index $ImageIndex" -Color "Green" -Indent 2         
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to install.wim..." -Color "Yellow" -Indent 2
        
        # Add each driver component
        foreach ($driverPath in $DriverPaths) {
            Write-ColorOutput "Adding drivers from: $driverPath" -Color "Cyan" -Indent 2
            Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Image:$mountDir", "/Add-Driver", "/Driver:$driverPath", "/Recurse") -Description "Add drivers from $driverPath" -SuppressOutput
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting install.wim..." -Color "Yellow" -Indent 2
        # Unmount and commit changes using direct execution
        Invoke-CommandWithExitCode -Command $dismPath -Arguments @("/Unmount-Wim", "/MountDir:$mountDir", "/Commit") -Description "Unmount install.wim" -SuppressOutput
        
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
