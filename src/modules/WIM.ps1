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
                        Write-ColorOutput "Unknown architecture: $arch" -Color "Red"
                        throw "Unknown or unsupported architecture detected: $arch"
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
        
        return $images
        
    } catch {
        Write-ColorOutput "Error getting WIM image info: $($_.Exception.Message)" -Color "Red"
        throw "Failed to get WIM image information: $($_.Exception.Message)"
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
                Write-ColorOutput "Found install image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2
            }
        }
    }
    
    # Analyze boot.wim if it exists
    if (Test-Path $bootWimPath) {
        Write-ColorOutput "Analyzing boot.wim..." -Color "Yellow" -Indent 1
        $bootWimInfo = Get-WimImageInfo -WimPath $bootWimPath -DismPath (Get-DismPath)
        
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
                Write-ColorOutput "Found boot image: $($image.Name) (Arch: $($image.Architecture), Version: $($image.Version))" -Color "Green" -Indent 2
            }
        }
    }
    
    if ($wims.Count -eq 0) {
        throw "No WIM files found in the ISO"
    }
    
    Write-ColorOutput "Found $($wims.Count) WIM image(s) total" -Color "Green"
    return $wims
}

