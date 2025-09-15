# Windows ISO Repack Script - Modular Architecture

This directory contains the modular components of the Windows ISO repack script, organized by functionality.

## Module Structure

### Common.ps1
Contains common utilities and helper functions:
- `Write-ColorOutput` - Enhanced logging function with indentation and color support
- `Test-Administrator` - Checks for administrator privileges
- `Invoke-WebRequestWithCleanup` - Robust web request function with progress tracking
- `Remove-WorkingDirectory` - Cleanup function for temporary directories

### Tools.ps1
Manages tools and prerequisites:
- `Find-OscdimgPath` - Locates Windows ADK oscdimg.exe
- `Test-DismAvailability` - Ensures DISM is available
- `Get-DismPath` - Gets DISM executable path
- `Test-Chocolatey` - Checks for Chocolatey package manager
- `Install-Chocolatey` - Installs Chocolatey if needed
- `Install-WindowsADK` - Installs Windows ADK via Chocolatey
- `Test-RequiredTools` - Orchestrates tool installation and verification

### ISO.ps1
Handles ISO file operations:
- `Extract-IsoContents` - Extracts ISO contents to working directory
- `Add-AutounattendXml` - Adds autounattend.xml to ISO contents
- `Add-OemDirectory` - Adds $OEM$ directory to ISO contents
- `New-IsoFromDirectory` - Creates new ISO from modified directory

### WIM.ps1
Manages WIM file analysis and information extraction:
- `Get-WimImageInfo` - Gets detailed information about WIM images (including architecture and version)
- `Get-AllWimInfo` - Analyzes all WIM files in the ISO

### VirtIO.ps1
Handles VirtIO driver management:
- `Get-VirtioDownloadUrl` - Gets VirtIO driver download URLs
- `Get-VirtioDrivers` - Downloads VirtIO drivers
- `Extract-VirtioDrivers` - Extracts VirtIO drivers from ISO
- `Add-VirtioDriversToWim` - Adds VirtIO drivers to individual WIM images
- `Add-VirtioDrivers` - Orchestrates VirtIO driver addition
- `Inject-VirtioDriversIntoBootWim` - Injects drivers into boot.wim
- `Inject-VirtioDriversIntoInstallWim` - Injects drivers into install.wim

## Usage

The main script (`packIso-Modular.ps1`) automatically loads all modules using dot sourcing:

```powershell
# Load modules
. (Join-Path $modulePath "Common.ps1")
. (Join-Path $modulePath "Tools.ps1")
. (Join-Path $modulePath "ISO.ps1")
. (Join-Path $modulePath "WIM.ps1")
. (Join-Path $modulePath "VirtIO.ps1")
```

## Benefits of Modular Architecture

1. **Maintainability**: Each module focuses on a specific area of functionality
2. **Reusability**: Modules can be used independently or in other scripts
3. **Testability**: Individual modules can be tested in isolation
4. **Readability**: Smaller, focused files are easier to understand and modify
5. **Collaboration**: Different team members can work on different modules
6. **Version Control**: Changes to specific functionality are isolated to relevant modules

## Dependencies

- **Common.ps1**: No dependencies (base utilities)
- **Tools.ps1**: Depends on Common.ps1
- **ISO.ps1**: Depends on Common.ps1
- **WIM.ps1**: Depends on Common.ps1 and Tools.ps1 (for Get-DismPath)
- **VirtIO.ps1**: Depends on Common.ps1, Tools.ps1, and WIM.ps1

## Loading Order

Modules should be loaded in dependency order:
1. Common.ps1 (base utilities)
2. Tools.ps1 (prerequisites)
3. ISO.ps1 (ISO operations)
4. WIM.ps1 (WIM analysis)
5. VirtIO.ps1 (VirtIO drivers)
