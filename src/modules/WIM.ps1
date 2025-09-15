# WIM analysis and information extraction for Windows ISO repack script

# Load Common utilities
$commonPath = Join-Path $PSScriptRoot "Common.ps1"
. $commonPath

# Load DISM module dependency
$toolsPath = Join-Path $PSScriptRoot "tools"
. (Join-Path $toolsPath "DISM.ps1")

function Get-WimImageArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WimPath,
        [Parameter(Mandatory = $true)]
        [int]$ImageIndex,
        [Parameter(Mandatory = $true)]
        [string]$DismPath
    )
    
    try {
        # Try to get architecture using DISM /Get-WimInfo with specific index
        $result = Start-Process -FilePath $DismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`"",
            "/Index:$ImageIndex"
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_arch_info.txt"
        
        if ($result.ExitCode -eq 0) {
            $archInfo = Get-Content "temp_arch_info.txt" -ErrorAction SilentlyContinue
            
            # Debug: Show the architecture-specific DISM output
            Write-ColorOutput "DISM architecture output for index $($ImageIndex):" -Color "Cyan" -Indent 2
            foreach ($line in $archInfo) {
                Write-ColorOutput "  $line" -Color "Gray" -Indent 0 -InheritedIndent 2
            }
            Write-Host ""
            
            # Clean up the temporary file
            Remove-Item "temp_arch_info.txt" -ErrorAction SilentlyContinue
            
            foreach ($line in $archInfo) {
                if ($line -match "Architecture\s*:\s*(.+)") {
                    $arch = $matches[1].Trim().ToLower()
                    switch ($arch) {
                        "x64" { return "amd64" }
                        "x86" { return "x86" }
                        "arm64" { return "arm64" }
                        default { return "amd64" }
                    }
                }
            }
        }
    } catch {
        # Ignore errors and return default
    }
    
    return "amd64"  # Default fallback
}

function Get-WimImageInfo {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "WIM file does not exist: $_"
            }
            if ($_ -notmatch '\.wim$') {
                throw "File must have .wim extension: $_"
            }
            $true
        })]
        [string]$WimPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "DISM executable does not exist: $_"
            }
            $true
        })]
        [string]$DismPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    try {
        Write-ColorOutput "Getting WIM image information from: $WimPath" -Color "Yellow" -Indent 0 -InheritedIndent $InheritedIndent
        
        # Use DISM to get image information
        $result = Start-Process -FilePath $dismPath -ArgumentList @(
            "/Get-WimInfo",
            "/WimFile:`"$WimPath`""
        ) -Wait -PassThru -NoNewWindow -RedirectStandardOutput "temp_wim_info.txt"
        
        if ($result.ExitCode -ne 0) {
            throw "Failed to get WIM image information. DISM exit code: $($result.ExitCode)"
        }
        
        # Parse the output to extract image information
        $wimInfo = Get-Content "temp_wim_info.txt" -ErrorAction SilentlyContinue
        
        # Debug: Log the raw DISM output for troubleshooting
        Write-ColorOutput "Raw DISM output from temp_wim_info.txt:" -Color "Cyan" -Indent 1 -InheritedIndent $InheritedIndent
        Write-Host ""
        foreach ($line in $wimInfo) {
            Write-ColorOutput "  $line" -Color "Gray" -Indent 0 -InheritedIndent ($InheritedIndent + 1)
        }
        Write-Host ""
        
        # Clean up the temporary file
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
                    Architecture = ""
                    Version = ""
                }
            } elseif ($currentImage -and $line -match "Name\s*:\s*(.+)") {
                $currentImage.Name = $matches[1].Trim()
                
                # Detect version from name
                if ($currentImage.Name -match "Windows 11" -or $currentImage.Name -match "Windows 1[1-9]") {
                    $currentImage.Version = "w11"
                } elseif ($currentImage.Name -match "Windows 10") {
                    $currentImage.Version = "w10"
                }
            } elseif ($currentImage -and $line -match "Description\s*:\s*(.+)") {
                $currentImage.Description = $matches[1].Trim()
                
                # Detect version from description if not already found
                if (-not $currentImage.Version) {
                    if ($currentImage.Description -match "Windows 11" -or $currentImage.Description -match "Windows 1[1-9]") {
                        $currentImage.Version = "w11"
                    } elseif ($currentImage.Description -match "Windows 10") {
                        $currentImage.Version = "w10"
                    }
                }
            } elseif ($line -match "Architecture\s*:\s*(.+)") {
                $arch = $matches[1].Trim()
                
                # Map DISM architecture names to our expected values
                switch ($arch.ToLower()) {
                    "x64" { $arch = "amd64" }
                    "x86" { $arch = "x86" }
                    "arm64" { $arch = "arm64" }
                    default {
                        Write-ColorOutput "Warning: Unknown architecture detected: $arch, defaulting to amd64" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                        $arch = "amd64"
                    }
                }
                
                # Set architecture for current image if we have one
                if ($currentImage) {
                    $currentImage.Architecture = $arch
                }
            }
        }
        
        if ($currentImage) {
            $images += $currentImage
        }
        
        # Post-process images to add fallback architecture detection
        foreach ($image in $images) {
            if (-not $image.Architecture) {
                Write-ColorOutput "No architecture found for image $($image.Index), attempting fallback detection..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                
                # Try to get architecture using specific index method
                $detectedArch = Get-WimImageArchitecture -WimPath $WimPath -ImageIndex $image.Index -DismPath $DismPath
                
                if ($detectedArch -ne "amd64" -or $image.Architecture) {
                    Write-ColorOutput "Architecture detected via DISM: $detectedArch" -Color "Green" -Indent 1 -InheritedIndent $InheritedIndent
                } else {
                    # Try to detect architecture from image name or description
                    Write-ColorOutput "DISM method failed, trying name/description analysis..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                    
                    # Check name for architecture hints
                    if ($image.Name -match "x64|64-bit|amd64") {
                        $detectedArch = "amd64"
                    } elseif ($image.Name -match "x86|32-bit") {
                        $detectedArch = "x86"
                    } elseif ($image.Name -match "arm64|arm") {
                        $detectedArch = "arm64"
                    }
                    
                    # Check description for architecture hints
                    if (-not $detectedArch -and $image.Description) {
                        if ($image.Description -match "x64|64-bit|amd64") {
                            $detectedArch = "amd64"
                        } elseif ($image.Description -match "x86|32-bit") {
                            $detectedArch = "x86"
                        } elseif ($image.Description -match "arm64|arm") {
                            $detectedArch = "arm64"
                        }
                    }
                    
                    if ($detectedArch -ne "amd64") {
                        Write-ColorOutput "Architecture detected from name/description: $detectedArch" -Color "Green" -Indent 1 -InheritedIndent $InheritedIndent
                    } else {
                        Write-ColorOutput "Could not detect architecture, defaulting to amd64" -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
                    }
                }
                
                $image.Architecture = $detectedArch
            }
        }
        
        return $images
        
    } catch {
        throw "Failed to get WIM image information: $($_.Exception.Message)"
    }
}



function Get-AllWimInfo {
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
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 20)]
        [int]$InheritedIndent = 0
    )
    
    Write-ColorOutput "=== Analyzing All WIM Files ===" -Color "Cyan" -Indent 0 -InheritedIndent $InheritedIndent
    
    Write-Host ""
    $wims = @()
    $installWimPath = Join-Path $ExtractPath "sources\install.wim"
    $bootWimPath = Join-Path $ExtractPath "sources\boot.wim"
    
    # Analyze install.wim if it exists
    if (Test-Path $installWimPath) {
        Write-ColorOutput "Analyzing install.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $installWimInfo = Get-WimImageInfo -WimPath $installWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($installWimInfo) {
            foreach ($image in $installWimInfo) {
                $wimInfo = @{
                    Path = $installWimPath
                    Type = "install"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $image.Architecture
                    Version = $image.Version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1 -InheritedIndent $InheritedIndent
        $bootWimInfo = Get-WimImageInfo -WimPath $bootWimPath -DismPath (Get-DismPath) -InheritedIndent ($InheritedIndent + 1)
        
        if ($bootWimInfo) {
            foreach ($image in $bootWimInfo) {
                $wimInfo = @{
                    Path = $bootWimPath
                    Type = "boot"
                    Index = $image.Index
                    Name = $image.Name
                    Description = $image.Description
                    Architecture = $image.Architecture
                    Version = $image.Version
                }
                
                $wims += $wimInfo
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2 -InheritedIndent $InheritedIndent
            }
        }
    }
    
    Write-Host ""
    if ($wims.Count -eq 0) {
        throw "No WIM files found in the ISO"
    }
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" -Color "Green" -Indent 0 -InheritedIndent $InheritedIndent
    return $wims
}

