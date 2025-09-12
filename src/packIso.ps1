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
    [string]$VirtioCacheDirectory = (Join-Path $env:TEMP "virtio-cache"),

    [Parameter(Mandatory = $true)]
    [ValidateSet("amd64", "x86", "arm64")]
    [string]$Arch,

    [Parameter(Mandatory = $true)]
    [ValidateSet("w10", "w11")]
    [string]$Version

)

$ErrorActionPreference = "Stop"

# Validate Windows 11 architecture compatibility
if ($Version -eq "w11" -and $Arch -eq "x86") {
    throw "Windows 11 does not support x86 architecture. Windows 11 only supports amd64 and arm64 architectures. Please use Version 'w10' for x86 architecture or change Arch to 'amd64' or 'arm64' for Windows 11."
}

# Validate VirtIO driver parameters
if ($IncludeVirtioDrivers) {
    # Validate VirtIO driver availability for ARM64 (only when VirtIO drivers are requested)
    if ($Arch -eq "arm64") {
        throw "VirtIO drivers are not available for ARM64 architecture. VirtIO drivers are only available for x86 and amd64 architectures. Please change Arch to 'x86' or 'amd64'."
    }
}

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
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
            Write-ColorOutput "Found oscdimg.exe at: $path" "Green"
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
        Write-ColorOutput "DISM found at: $($dismCommand.Source)" "Green"
        return
    }
    
    # DISM should be available on Windows 7+ by default, but let's check common locations
    $dismPaths = @(
        "${env:SystemRoot}\System32\dism.exe",
        "${env:SystemRoot}\SysWOW64\dism.exe"
    )
    
    foreach ($path in $dismPaths) {
        if (Test-Path $path) {
            Write-ColorOutput "DISM found at: $path" "Green"
            # Add to PATH for current session if not already there
            $dismDir = Split-Path $path -Parent
            if ($env:Path -notlike "*$dismDir*") {
                $env:Path += ";$dismDir"
                Write-ColorOutput "Added DISM directory to PATH: $dismDir" "Cyan"
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
            Write-ColorOutput "DISM is now available" "Green"
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
            Write-ColorOutput "DISM enabled successfully at: $($dismCommand.Source)" "Green"
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
                Write-ColorOutput "DISM now available at: $($dismCommand.Source)" "Green"
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
    Write-ColorOutput "ISO mounted at: $mountedPath" "Green"
    
    try {
        Write-ColorOutput "Extracting ISO contents to: $ExtractPath" "Yellow"
        
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        
        robocopy $mountedPath $ExtractPath /E /COPY:DT /R:3 /W:10 /NFL /NDL /NJH /NJS /nc /ns /np
        
        if ($LASTEXITCODE -gt 7) {
            throw "Failed to extract ISO contents. Robocopy exit code: $LASTEXITCODE"
        }

        Write-ColorOutput "ISO contents extracted successfully" "Green"
    } finally {
        Write-ColorOutput "Dismounting ISO..." "Yellow"
        Dismount-DiskImage -ImagePath $IsoPath
        Write-ColorOutput "ISO dismounted" "Green"
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
    Write-ColorOutput "autounattend.xml added to: $destinationPath" "Green"
}

function Add-OemDirectory {
    param(
        [string]$ExtractPath,
        [string]$OemSourcePath
    )
    Write-ColorOutput "Adding $OEM$ directory to ISO contents..." "Yellow"
    
    if (-not (Test-Path $OemSourcePath -PathType Container)) {
        Write-ColorOutput "Warning: $OEM$ directory not found at: $OemSourcePath" "Yellow"
        return
    }
    
    $destinationPath = Join-Path $ExtractPath '$OEM$'
    
    # Remove existing $OEM$ directory if it exists
    if (Test-Path $destinationPath) {
        Remove-Item $destinationPath -Recurse -Force
    }
    
    # Copy the entire $OEM$ directory structure
    Copy-Item $OemSourcePath $destinationPath -Recurse -Force
    Write-ColorOutput "$OEM$ directory added to: $destinationPath" "Green"
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

    Write-ColorOutput "Using source directly: $absSrc" "Cyan"

    $arguments = @(
        "-m"
        "-u2"
        "-udfver102"
        "-bootdata:2#p0,e,b`"$etfsbootPath`"#pEF,e,b`"$efisysPath`""
        "`"$absSrc`""
        "`"$absOutIso`""
    )

    Write-ColorOutput "Current working directory: $(Get-Location)" "Cyan"
    Write-ColorOutput "Running oscdimg with arguments: $($arguments -join ' ')" "Cyan"
    Write-ColorOutput "Full command: & `"$OscdimgPath`" $($arguments -join ' ')" "Cyan"
    
    & $OscdimgPath $arguments
    if ($LASTEXITCODE -ne 0) { throw "oscdimg failed with exit code: $LASTEXITCODE" }

    Write-ColorOutput "ISO created successfully: $absOutIso" "Green"
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
        Write-ColorOutput "Created cache directory: $CacheDirectory" "Green"
    }
    
    $downloadUrl = Get-VirtioDownloadUrl -Version $Version
    $fileName = "virtio-win-$Version.iso"
    $localPath = Join-Path $CacheDirectory $fileName
    
    # Check if we already have the file
    if (Test-Path $localPath) {
        Write-ColorOutput "VirtIO drivers already cached: $localPath" "Green"
        return $localPath
    }
    
    Write-ColorOutput "Downloading VirtIO drivers from: $downloadUrl" "Yellow"
    Write-ColorOutput "Saving to: $localPath" "Yellow"
    
    try {
        # Use Invoke-WebRequest with progress
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($downloadUrl, $localPath)
        Write-ColorOutput "VirtIO drivers downloaded successfully" "Green"
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
        # Download VirtIO drivers
        $virtioIsoPath = Get-VirtioDrivers -Version $VirtioVersion -CacheDirectory $VirtioCacheDirectory
        
        # Extract VirtIO drivers to a temporary directory
        $virtioDir = Extract-VirtioDrivers -VirtioIsoPath $virtioIsoPath -ExtractPath $ExtractPath
        
        Write-ColorOutput "VirtIO drivers extracted to: $virtioDir" "Green"
        
        # Log the driver structure for debugging
        $driverDirs = Get-ChildItem -Path $virtioDir -Directory | Select-Object -ExpandProperty Name
        Write-ColorOutput "Available driver directories: $($driverDirs -join ', ')" "Cyan"
        
        # Inject drivers into boot.wim
        Inject-VirtioDriversIntoBootWim -ExtractPath $ExtractPath -VirtioDir $virtioDir -Arch $Arch -Version $Version
        
        # Inject drivers into install.wim
        Inject-VirtioDriversIntoInstallWim -ExtractPath $ExtractPath -VirtioDir $virtioDir -Arch $Arch -Version $Version
        
    } catch {
        Write-ColorOutput "Failed to add VirtIO drivers: $($_.Exception.Message)" "Red"
        throw
    }
}

function Inject-VirtioDriversIntoBootWim {
    param(
        [string]$ExtractPath,
        [string]$VirtioDir,
        [string]$Arch,
        [string]$Version
    )
    
    Write-ColorOutput "=== Injecting VirtIO Drivers into boot.wim ===" "Cyan"
    
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    $mountDir = Join-Path $ExtractPath "boot_mount"
    
    if (-not (Test-Path $bootWimPath)) {
        throw "boot.wim not found at: $bootWimPath"
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
            Write-ColorOutput "Error: No drivers found for architecture $Arch at: $driverPath" "Red"
            throw "VirtIO drivers not found for architecture: $Arch"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            Write-ColorOutput "Error: No drivers found for Windows version $Version at: $windowsDriverPath" "Red"
            throw "VirtIO drivers not found for Windows version: $Version"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" "Green"
        Write-ColorOutput "Architecture: $Arch" "Cyan"
        Write-ColorOutput "Version: $Version" "Cyan"
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Mount boot.wim (try both indexes)
        $mounted = $false
        foreach ($index in @(1, 2)) {
            try {
                Write-ColorOutput "Attempting to mount boot.wim index $index..." "Yellow"
                $result = Start-Process -FilePath $dismPath -ArgumentList @(
                    "/Mount-Wim",
                    "/WimFile:`"$bootWimPath`"",
                    "/Index:$index",
                    "/MountDir:`"$mountDir`""
                ) -Wait -PassThru -NoNewWindow
                
                if ($result.ExitCode -eq 0) {
                    Write-ColorOutput "Successfully mounted boot.wim index $index" "Green"
                    $mounted = $true
                    break
                } else {
                    Write-ColorOutput "Failed to mount boot.wim index $index (exit code: $($result.ExitCode))" "Yellow"
                }
            } catch {
                Write-ColorOutput "Error mounting boot.wim index $index`: $($_.Exception.Message)" "Yellow"
            }
        }
        
        if (-not $mounted) {
            throw "Failed to mount boot.wim with any index"
        }
        
        # Add drivers to the mounted image
        Write-ColorOutput "Adding VirtIO drivers to boot.wim..." "Yellow"
                $result = Start-Process -FilePath $dismPath -ArgumentList @(
                    "/Image:`"$mountDir`"",
                    "/Add-Driver",
                    "/Driver:`"$windowsDriverPath`"",
                    "/Recurse"
                ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully added VirtIO drivers to boot.wim" "Green"
        } else {
            Write-ColorOutput "Warning: Failed to add drivers to boot.wim (exit code: $($result.ExitCode))" "Yellow"
        }
        
        # Unmount and commit changes
        Write-ColorOutput "Unmounting boot.wim..." "Yellow"
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Unmount-Wim",
            "/MountDir:`"$mountDir`"",
            "/Commit"
        ) -Wait -PassThru -NoNewWindow
        
        if ($result.ExitCode -eq 0) {
            Write-ColorOutput "Successfully unmounted and committed boot.wim changes" "Green"
        } else {
            Write-ColorOutput "Warning: Failed to unmount boot.wim (exit code: $($result.ExitCode))" "Yellow"
        }
        
    } catch {
        Write-ColorOutput "Error injecting drivers into boot.wim: $($_.Exception.Message)" "Red"
        throw
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" "Yellow"
            }
        }
    }
}

function Inject-VirtioDriversIntoInstallWim {
    param(
        [string]$ExtractPath,
        [string]$VirtioDir,
        [string]$Arch,
        [string]$Version
    )
    
    Write-ColorOutput "=== Injecting VirtIO Drivers into install.wim ===" "Cyan"
    
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $mountDir = Join-Path $ExtractPath "install_mount"
    
    if (-not (Test-Path $installWimPath)) {
        Write-ColorOutput "Warning: install.wim not found at: $installWimPath" "Yellow"
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
            Write-ColorOutput "Error: No drivers found for architecture $Arch at: $driverPath" "Red"
            throw "VirtIO drivers not found for architecture: $Arch"
        }
        
        # Find the specified Windows version drivers
        $windowsDriverPath = Join-Path $driverPath $Version
        if (-not (Test-Path $windowsDriverPath)) {
            Write-ColorOutput "Error: No drivers found for Windows version $Version at: $windowsDriverPath" "Red"
            throw "VirtIO drivers not found for Windows version: $Version"
        }
        
        Write-ColorOutput "Using drivers from: $windowsDriverPath" "Green"
        Write-ColorOutput "Architecture: $Arch" "Cyan"
        Write-ColorOutput "Version: $Version" "Cyan"
        
        # Get DISM path
        $dismPath = Get-DismPath
        
        # Get install.wim image information
        Write-ColorOutput "Getting install.wim image information..." "Yellow"
        $imageInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath $dismPath
        
        if (-not $imageInfo) {
            Write-ColorOutput "Warning: Could not get install.wim image information" "Yellow"
            return
        }
        
        Write-ColorOutput "Found $($imageInfo.Count) image(s) in install.wim" "Green"
        
        # Process each image in install.wim
        foreach ($image in $imageInfo) {
            $imageIndex = $image.Index
            $imageName = $image.Name
            $imageDescription = $image.Description
            
            Write-ColorOutput "Processing image $imageIndex`: $imageName" "Yellow"
            if ($imageDescription) {
                Write-ColorOutput "  Description: $imageDescription" "Cyan"
            }
            
            try {
                # Mount the image
                Write-ColorOutput "  Mounting image $imageIndex..." "Yellow"
                $result = Start-Process -FilePath $dismPath -ArgumentList @(
                    "/Mount-Wim",
                    "/WimFile:`"$installWimPath`"",
                    "/Index:$imageIndex",
                    "/MountDir:`"$mountDir`""
                ) -Wait -PassThru -NoNewWindow
                
                if ($result.ExitCode -ne 0) {
                    Write-ColorOutput "  Warning: Failed to mount image $imageIndex (exit code: $($result.ExitCode))" "Yellow"
                    continue
                }
                
                Write-ColorOutput "  Successfully mounted image $imageIndex" "Green"
                
                # Add drivers to the mounted image
                Write-ColorOutput "  Adding VirtIO drivers to image $imageIndex..." "Yellow"
                $result = Start-Process -FilePath $dismPath -ArgumentList @(
                    "/Image:`"$mountDir`"",
                    "/Add-Driver",
                    "/Driver:`"$windowsDriverPath`"",
                    "/Recurse"
                ) -Wait -PassThru -NoNewWindow
                
                if ($result.ExitCode -eq 0) {
                    Write-ColorOutput "  Successfully added VirtIO drivers to image $imageIndex" "Green"
                } else {
                    Write-ColorOutput "  Warning: Failed to add drivers to image $imageIndex (exit code: $($result.ExitCode))" "Yellow"
                }
                
                # Unmount and commit changes
                Write-ColorOutput "  Unmounting image $imageIndex..." "Yellow"
                $result = Start-Process -FilePath $dismPath -ArgumentList @(
                    "/Unmount-Wim",
                    "/MountDir:`"$mountDir`"",
                    "/Commit"
                ) -Wait -PassThru -NoNewWindow
                
                if ($result.ExitCode -eq 0) {
                    Write-ColorOutput "  Successfully unmounted and committed image $imageIndex" "Green"
                } else {
                    Write-ColorOutput "  Warning: Failed to unmount image $imageIndex (exit code: $($result.ExitCode))" "Yellow"
                }
                
            } catch {
                Write-ColorOutput "  Error processing image $imageIndex`: $($_.Exception.Message)" "Red"
                # Try to unmount if there was an error
                try {
                    Start-Process -FilePath $dismPath -ArgumentList @(
                        "/Unmount-Wim",
                        "/MountDir:`"$mountDir`"",
                        "/Discard"
                    ) -Wait -PassThru -NoNewWindow | Out-Null
                } catch {
                    # Ignore unmount errors during cleanup
                }
            }
        }
        
        Write-ColorOutput "Completed processing all images in install.wim" "Green"
        
    } catch {
        Write-ColorOutput "Error injecting drivers into install.wim: $($_.Exception.Message)" "Red"
        throw
    } finally {
        # Cleanup mount directory
        if (Test-Path $mountDir) {
            try {
                Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-ColorOutput "Warning: Could not clean up mount directory: $mountDir" "Yellow"
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
        Write-ColorOutput "Architecture: $Arch" "White"
        Write-ColorOutput "Version: $Version" "White"
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
    
    Extract-IsoContents -IsoPath $resolvedInputIso -ExtractPath $WorkingDirectory
    Add-AutounattendXml -ExtractPath $WorkingDirectory -AutounattendXmlPath $AutounattendXml
    Add-OemDirectory -ExtractPath $WorkingDirectory -OemSourcePath $OemDirectory
    Add-VirtioDrivers -ExtractPath $WorkingDirectory -VirtioVersion $VirtioVersion -VirtioCacheDirectory $VirtioCacheDirectory
    New-IsoFromDirectory -SourcePath $WorkingDirectory -OutputPath $resolvedOutputIso -OscdimgPath $script:oscdimgPath
    
    if (Test-Path $resolvedOutputIso) {
        $fileSize = (Get-Item $resolvedOutputIso).Length
        $fileSizeGB = [math]::Round($fileSize / 1GB, 2)
        Write-ColorOutput "Output ISO created successfully!" "Green"
        Write-ColorOutput "File size: $fileSizeGB GB" "Green"
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
