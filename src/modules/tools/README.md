# Tools Module - Individual Tool Scripts

This directory contains individual scripts for each tool used by the Windows ISO repack script, organized by functionality.

## Tool Scripts

### Chocolatey.ps1
Manages the Chocolatey package manager:
- `Test-Chocolatey` - Checks if Chocolatey is installed
- `Install-Chocolatey` - Installs Chocolatey if not present

### WindowsADK.ps1
Manages Windows Assessment and Deployment Kit:
- `Install-WindowsADK` - Installs Windows ADK via Chocolatey
- Dependencies: Chocolatey.ps1

### OSCDIMG.ps1
Manages the oscdimg.exe tool from Windows ADK:
- `Find-OscdimgPath` - Locates oscdimg.exe in standard ADK installation paths

### DISM.ps1
Manages the Deployment Image Servicing and Management tool:
- `Test-DismAvailability` - Ensures DISM is available and attempts installation if needed
- `Get-DismPath` - Gets the path to the DISM executable
- Dependencies: Chocolatey.ps1 (for installation fallback)

### ToolsOrchestrator.ps1
Orchestrates all tool installation and verification:
- `Test-RequiredTools` - Main function that ensures all required tools are available
- Dependencies: All other tool scripts

## Loading Order

The tools should be loaded in dependency order:

1. **Chocolatey.ps1** - Base package manager (no dependencies)
2. **WindowsADK.ps1** - Depends on Chocolatey
3. **OSCDIMG.ps1** - Depends on Windows ADK installation
4. **DISM.ps1** - Depends on Chocolatey (for fallback installation)
5. **ToolsOrchestrator.ps1** - Depends on all other tools

## Benefits of Individual Tool Scripts

1. **Focused Responsibility**: Each script manages a single tool
2. **Independent Testing**: Individual tools can be tested in isolation
3. **Selective Loading**: Only load the tools you need
4. **Easier Maintenance**: Changes to one tool don't affect others
5. **Clear Dependencies**: Dependencies are explicitly managed
6. **Reusability**: Individual tool scripts can be used in other projects

## Usage

The main script loads all tools in the correct order:

```powershell
$toolsPath = Join-Path $modulePath "tools"
. (Join-Path $toolsPath "Chocolatey.ps1")
. (Join-Path $toolsPath "WindowsADK.ps1")
. (Join-Path $toolsPath "OSCDIMG.ps1")
. (Join-Path $toolsPath "DISM.ps1")
. (Join-Path $toolsPath "ToolsOrchestrator.ps1")
```

Individual modules can also load specific tools as needed:

```powershell
# Load only DISM for WIM operations
$toolsPath = Join-Path (Split-Path $PSScriptRoot -Parent) "tools"
. (Join-Path $toolsPath "DISM.ps1")
```
