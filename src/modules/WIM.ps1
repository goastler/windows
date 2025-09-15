# WIM analysis and information extraction for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path (Split-Path $PSScriptRoot) "Common.ps1"
. $commonPath

# Load DISM module dependency
$toolsPath = Join-Path (Split-Path $PSScriptRoot -Parent) "tools"
. (Join-Path $toolsPath "DISM.ps1")

function Get-WimImageInfo {
    param(
        [string]$WimPath,
        [string]$DismPath
    )
    
    try {
        Write-ColorOutput "Getting WIM image information from: $WimPath" -Color "Yellow"
        
        # Use DISM to get image information
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_wim_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM info (exit code: $($result.ExitCode))" -Color "Red"
            throw "Failed to get WIM image information. DISM exit code: $($result.ExitCode)"
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
        Write-ColorOutput "Error getting WIM image info: $($_.Exception.Message)" -Color "Red"
        throw "Failed to get WIM image information: $($_.Exception.Message)"
    }
}

function Get-WimArchitecture {
    param(
        [string]$WimPath
    )
    
    Write-ColorOutput "Inferring architecture from WIM: $WimPath" -Color "Yellow"
    
    try {
        $dismPath = Get-DismPath
        
        # Get WIM information using DISM
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_arch_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM architecture info (exit code: $($result.ExitCode))" -Color "Red"
            throw "Failed to get WIM architecture information. DISM exit code: $($result.ExitCode)"
        }
        
        # Parse the output to extract architecture information
        $wimInfo = Get-Content "temp_arch_info.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_arch_info.txt" -ErrorAction SilentlyContinue
        
        foreach ($line in $wimInfo) {
            if ($line -match "Architecture\s*:\s*(.+)") {
                $arch = $matches[1].Trim()
                Write-ColorOutput "Detected architecture: $arch" -Color "Green" -Indent 1
                
                # Map DISM architecture names to our expected values
                switch ($arch.ToLower()) {
                    "x64" { return "amd64" }
                    "x86" { return "x86" }
                    "arm64" { return "arm64" }
                    default {
                        Write-ColorOutput "Unknown architecture: $arch" -Color "Red"
                        throw "Unknown or unsupported architecture detected: $arch"
                    }
                }
            }
        }
        
        Write-ColorOutput "Could not determine architecture from WIM" -Color "Red"
        throw "Failed to determine architecture from WIM file. No architecture information found in WIM metadata."
        
    } catch {
        Write-ColorOutput "Error inferring architecture from WIM: $($_.Exception.Message)" -Color "Red"
        throw "Failed to infer architecture from WIM: $($_.Exception.Message)"
    }
}

function Get-WimVersion {
    param(
        [string]$WimPath
    )
    
    Write-ColorOutput "Inferring Windows version from WIM: $WimPath" -Color "Yellow"
    
    try {
        $dismPath = Get-DismPath
        
        # Get WIM information using DISM
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_version_info.txt"
        
        if ($result.ExitCode -ne 0) {
            Write-ColorOutput "Failed to get WIM version info (exit code: $($result.ExitCode))" -Color "Red"
            throw "Failed to get WIM version information. DISM exit code: $($result.ExitCode)"
        }
        
        # Parse the output to extract version information
        $wimInfo = Get-Content "temp_version_info.txt" -ErrorAction SilentlyContinue
        Remove-Item "temp_version_info.txt" -ErrorAction SilentlyContinue
        
        foreach ($line in $wimInfo) {
            if ($line -match "Name\s*:\s*(.+)") {
                $name = $matches[1].Trim()
                Write-ColorOutput "Detected image name: $name" -Color "Green" -Indent 1
                
                # Check for Windows 11 indicators
                if ($name -match "Windows 11" -or $name -match "Windows 1[1-9]") {
                    Write-ColorOutput "Detected Windows version: w11" -Color "Green" -Indent 2
                    return "w11"
                }
                # Check for Windows 10 indicators
                elseif ($name -match "Windows 10") {
                    Write-ColorOutput "Detected Windows version: w10" -Color "Green" -Indent 2
                    return "w10"
                }
            }
            elseif ($line -match "Description\s*:\s*(.+)") {
                $description = $matches[1].Trim()
                Write-ColorOutput "Detected image description: $description" -Color "Green" -Indent 1
                
                # Check for Windows 11 indicators in description
                if ($description -match "Windows 11" -or $description -match "Windows 1[1-9]") {
                    Write-ColorOutput "Detected Windows version: w11" -Color "Green" -Indent 2
                    return "w11"
                }
                # Check for Windows 10 indicators in description
                elseif ($description -match "Windows 10") {
                    Write-ColorOutput "Detected Windows version: w10" -Color "Green" -Indent 2
                    return "w10"
                }
            }
        }
        
        Write-ColorOutput "Could not determine Windows version from WIM" -Color "Red"
        throw "Failed to determine Windows version from WIM file. No version information found in WIM metadata."
        
    } catch {
        Write-ColorOutput "Error inferring Windows version from WIM: $($_.Exception.Message)" -Color "Red"
        throw "Failed to infer Windows version from WIM: $($_.Exception.Message)"
    }
}

function Get-AllWimInfo {
    param(
        [string]$ExtractPath
    )
    
    Write-ColorOutput "=== Analyzing All WIM Files ===" -Color "Cyan"
    
    $wims = @()
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1
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
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $arch, Version: $version)" -Color "Green" -Indent 2
            }
        }
    }
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1
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
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $arch, Version: $version)" -Color "Green" -Indent 2
            }
        }
    }
    
    if ($wims.Count -eq 0) {
        throw "No WIM files found in the ISO"
    }
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" -Color "Green"
    return $wims
}

function Get-WimInfo {
    param(
        [string]$ExtractPath
    )
    
    Write-ColorOutput "=== Inferring WIM Information ===" -Color "Cyan"
    
    # Try to get info from install.wim first (more reliable)
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    $arch = $null
    $version = $null
    
    # Try install.wim first
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1
        $arch = Get-WimArchitecture -WimPath $installWimPath
        $version = Get-WimVersion -WimPath $installWimPath
    }
    
    # If we couldn't get info from install.wim, try boot.wim
    if (-not $arch -and (Test-Path $bootWimPath)) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1
        $arch = Get-WimArchitecture -WimPath $bootWimPath
    }
    
    if (-not $version -and (Test-Path $bootWimPath)) {
        Write-ColorOutput "Analyzing boot.wim for version..." -Color "Yellow" -Indent 1
        $version = Get-WimVersion -WimPath $bootWimPath
    }
    
    if (-not $arch) {
        throw "Could not determine architecture from WIM files"
    }
    
    if (-not $version) {
        throw "Could not determine Windows version from WIM files"
    }
    
    Write-ColorOutput "Inferred architecture: $arch" -Color "Green"
    Write-ColorOutput "Inferred version: $version" -Color "Green"
    
    return @{
        Architecture = $arch
        Version = $version
    }
}
