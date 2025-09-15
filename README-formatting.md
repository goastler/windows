# PowerShell Script Formatting and Analysis Tools

This directory contains several tools to format and analyze PowerShell scripts using PowerShell's built-in tools and community modules.

## Available Tools

### 1. `Format-Analyze.ps1` (Recommended)
A PowerShell script that provides comprehensive formatting and analysis capabilities.

**Usage:**
```powershell
# Format and analyze all scripts in current directory
.\Format-Analyze.ps1 -Format -Analyze

# Analyze all scripts recursively
.\Format-Analyze.ps1 -Analyze -Recurse

# Format and analyze scripts in specific directory
.\Format-Analyze.ps1 -Path ./src -Format -Analyze -Recurse

# Show help
.\Format-Analyze.ps1 -Help
```

### 2. `format-analyze.sh`
A bash script that wraps PowerShell commands for Unix-like systems.

**Usage:**
```bash
# Format all scripts
./format-analyze.sh --format

# Analyze all scripts
./format-analyze.sh --analyze

# Format and analyze all scripts
./format-analyze.sh --format-analyze

# Format and analyze specific script
./format-analyze.sh --script ./myscript.ps1

# Show help
./format-analyze.sh --help
```

### 3. `format-and-analyze.ps1`
A comprehensive bash script with interactive menu options.

**Usage:**
```bash
./format-and-analyze.ps1
```

## Direct PowerShell Commands

You can also run the PowerShell commands directly:

### Script Analysis
```powershell
# Analyze all scripts in current directory and subdirectories
Invoke-ScriptAnalyzer -Path . -Recurse

# Analyze specific script
Invoke-ScriptAnalyzer -Path ./myscript.ps1

# Analyze with specific rules
Invoke-ScriptAnalyzer -Path . -Recurse -IncludeRule @('PSAvoidUsingCmdletAliases', 'PSUseConsistentWhitespace')
```

### Script Formatting
```powershell
# Format a specific script
Import-Module PowerShell-Beautifier
$content = Get-Content ./myscript.ps1 -Raw
$formatted = Invoke-Beautifier -ScriptDefinition $content
Set-Content -Path ./myscript.ps1 -Value $formatted -Encoding UTF8
```

## Required Modules

The scripts will automatically install the required modules:

- **PSScriptAnalyzer**: For static code analysis
- **PowerShell-Beautifier**: For code formatting

## Features

### Formatting
- Consistent indentation and whitespace
- Proper line breaks and spacing
- Code structure improvements
- Automatic backup creation

### Analysis
- Static code analysis using PSScriptAnalyzer
- Detection of common PowerShell issues
- Best practice recommendations
- Detailed reporting with severity levels

### Reporting
- Console output with color coding
- Detailed analysis reports
- Summary statistics
- Issue categorization by severity

## Example Output

### Analysis Results
```
Found 5 issues across all scripts:

Script: packIso.ps1
  [Warning] Line 45: The cmdlet 'Write-Host' is used. Consider using 'Write-Output' instead.
  [Information] Line 123: Consider using 'Test-Path' instead of 'if (Get-Item)'.

Script: setup.ps1
  [Error] Line 67: The variable '$ErrorActionPreference' is assigned but never used.
```

### Formatting Results
```
Found 3 PowerShell script(s) to format
Formatting: ./src/packIso.ps1
[OK] Formatted: packIso.ps1
Formatting: ./src/setup.ps1
[OK] Formatted: setup.ps1
```

## Best Practices

1. **Always backup your scripts** before formatting
2. **Review analysis results** before making changes
3. **Test your scripts** after formatting
4. **Use version control** to track changes
5. **Run analysis regularly** during development

## Troubleshooting

### PowerShell Module Installation Issues
If module installation fails, try:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser -AllowClobber
```

### Permission Issues
If you encounter permission issues:
```powershell
# Run PowerShell as Administrator
# Or use -Scope CurrentUser for module installation
```

### Script Analysis Errors
If analysis fails:
1. Check that the script has valid PowerShell syntax
2. Ensure the script path is correct
3. Verify PSScriptAnalyzer is properly installed

## Integration with CI/CD

These tools can be integrated into CI/CD pipelines:

```yaml
# Example GitHub Actions step
- name: Analyze PowerShell Scripts
  run: |
    pwsh -Command "Invoke-ScriptAnalyzer -Path ./scripts -Recurse"
```

```bash
# Example in build script
./format-analyze.sh --analyze
if [ $? -ne 0 ]; then
    echo "Script analysis failed"
    exit 1
fi
```
