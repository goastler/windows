# PowerShell Script Formatter and Analyzer
# This script formats and analyzes all PowerShell scripts in the current directory

param(
    [Parameter(Mandatory = $false)]
    [string]$Path = ".",
    
    [Parameter(Mandatory = $false)]
    [switch]$Format,
    
    [Parameter(Mandatory = $false)]
    [switch]$Analyze,
    
    [Parameter(Mandatory = $false)]
    [switch]$Recurse,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckSpecialChars,
    
    [Parameter(Mandatory = $false)]
    [switch]$Help
)

function Show-Help {
    Write-Host "PowerShell Script Formatter and Analyzer" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage: .\Format-Analyze.ps1 [OPTIONS]" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "  -Path <path>           Path to analyze (default: current directory)" -ForegroundColor White
    Write-Host "  -Format                Format PowerShell scripts" -ForegroundColor White
    Write-Host "  -Analyze               Analyze PowerShell scripts" -ForegroundColor White
    Write-Host "  -CheckSpecialChars     Check for special characters in scripts" -ForegroundColor White
    Write-Host "  -Recurse               Include subdirectories" -ForegroundColor White
    Write-Host "  -Help                  Show this help message" -ForegroundColor White
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\Format-Analyze.ps1 -Format -Analyze" -ForegroundColor Green
    Write-Host "  .\Format-Analyze.ps1 -Path ./src -Recurse" -ForegroundColor Green
    Write-Host "  .\Format-Analyze.ps1 -Analyze -Recurse" -ForegroundColor Green
    Write-Host "  .\Format-Analyze.ps1 -CheckSpecialChars" -ForegroundColor Green
    Write-Host ""
    Write-Host "Direct PowerShell commands:" -ForegroundColor Yellow
    Write-Host "  Invoke-ScriptAnalyzer -Path . -Recurse" -ForegroundColor Green
    Write-Host "  Invoke-Beautifier -ScriptDefinition (Get-Content ./myscript.ps1 -Raw)" -ForegroundColor Green
}

function Install-RequiredModules {
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Yellow
    
    # Check for PSScriptAnalyzer
    if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
        Write-Host "Installing PSScriptAnalyzer..." -ForegroundColor Yellow
        Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser
    } else {
        Write-Host "[OK] PSScriptAnalyzer is available" -ForegroundColor Green
    }
    
    # Check for PowerShell-Beautifier
    if (-not (Get-Module -ListAvailable PowerShell-Beautifier)) {
        Write-Host "Installing PowerShell-Beautifier..." -ForegroundColor Yellow
        Install-Module -Name PowerShell-Beautifier -Force -Scope CurrentUser
    } else {
        Write-Host "[OK] PowerShell-Beautifier is available" -ForegroundColor Green
    }
    
    Write-Host ""
}

function Remove-SpecialCharacters {
    param(
        [string]$Content
    )
    
    # Only replace common problematic characters that are actually used in scripts
    $replacements = @{
        '✓' = '[OK]'
        '✗' = '[ERROR]'
        '✘' = '[ERROR]'
        '⚠' = '[WARNING]'
        'ℹ' = '[INFO]'
    }
    
    # Apply replacements
    $cleanContent = $Content
    foreach ($special in $replacements.Keys) {
        $cleanContent = $cleanContent -replace [regex]::Escape($special), $replacements[$special]
    }
    
    # Only remove null bytes and other truly problematic control characters
    # Preserve all whitespace including line breaks, tabs, and spaces
    $cleanContent = $cleanContent -replace '\x00', ''  # Remove null bytes only
    
    return $cleanContent
}

function Test-SpecialCharacters {
    param(
        [string]$Content
    )
    
    # Check for problematic characters: null bytes and common Unicode symbols that cause issues
    $problematicChars = @('✓', '✗', '✘', '⚠', 'ℹ', '\x00')
    $foundChars = @()
    
    foreach ($char in $problematicChars) {
        if ($Content.Contains($char)) {
            $foundChars += $char
        }
    }
    
    if ($foundChars.Count -gt 0) {
        return @{
            HasSpecialChars = $true
            SpecialChars = $foundChars
        }
    }
    
    return @{
        HasSpecialChars = $false
        SpecialChars = @()
    }
}

function Format-Scripts {
    param(
        [string]$ScriptPath,
        [bool]$Recurse
    )
    
    Write-Host "=== Formatting PowerShell Scripts ===" -ForegroundColor Cyan
    
    $scripts = Get-ChildItem -Path $ScriptPath -Filter "*.ps1" -Recurse:$Recurse
    
    if ($scripts.Count -eq 0) {
        Write-Host "No PowerShell scripts found in $ScriptPath" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($scripts.Count) PowerShell script(s) to format" -ForegroundColor Green
    
    foreach ($script in $scripts) {
        Write-Host "Processing: $($script.FullName)" -ForegroundColor Yellow
        
        try {
            # Create backup
            $backupPath = "$($script.FullName).backup"
            Copy-Item -Path $script.FullName -Destination $backupPath -Force
            
            # Read the script content
            $content = Get-Content $script.FullName -Raw
            
            # Check for special characters
            $specialCharTest = Test-SpecialCharacters -Content $content
            if ($specialCharTest.HasSpecialChars) {
                $charDetails = $specialCharTest.SpecialChars | ForEach-Object { 
                    if ($_ -match '[\x00-\x1F\x7F-\x9F]') {
                        "[Control: $(($_ -as [int]))]"
                    } else {
                        "'$_'"
                    }
                }
                Write-Host "  Found special characters: $($charDetails -join ', ')" -ForegroundColor Yellow
                Write-Host "  Removing special characters..." -ForegroundColor Yellow
                $content = Remove-SpecialCharacters -Content $content
            }
            
            # Format the script (if PowerShell-Beautifier is available)
            try {
                Import-Module PowerShell-Beautifier -Force -ErrorAction Stop
                $formatted = Invoke-Beautifier -ScriptDefinition $content
                Set-Content -Path $script.FullName -Value $formatted -Encoding UTF8
            } catch {
                # If PowerShell-Beautifier is not available, just save the cleaned content
                Write-Host "  PowerShell-Beautifier not available, saving cleaned content only" -ForegroundColor Yellow
                # Use ASCII encoding to ensure no special characters
                [System.IO.File]::WriteAllText($script.FullName, $content, [System.Text.Encoding]::ASCII)
            }
            
            Write-Host "[OK] Formatted: $($script.Name)" -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Failed to format: $($script.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host ""
}

function Analyze-Scripts {
    param(
        [string]$ScriptPath,
        [bool]$Recurse
    )
    
    Write-Host "=== Analyzing PowerShell Scripts ===" -ForegroundColor Cyan
    
    $scripts = Get-ChildItem -Path $ScriptPath -Filter "*.ps1" -Recurse:$Recurse
    
    if ($scripts.Count -eq 0) {
        Write-Host "No PowerShell scripts found in $ScriptPath" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($scripts.Count) PowerShell script(s) to analyze" -ForegroundColor Green
    
    try {
        Import-Module PSScriptAnalyzer -Force
        
        $results = Invoke-ScriptAnalyzer -Path $ScriptPath -Recurse:$Recurse
        
        if ($results.Count -eq 0) {
            Write-Host "[OK] No issues found in any PowerShell scripts!" -ForegroundColor Green
        } else {
            Write-Host "Found $($results.Count) issues across all scripts:" -ForegroundColor Yellow
            
            $groupedResults = $results | Group-Object ScriptName
            
            foreach ($group in $groupedResults) {
                Write-Host "`nScript: $($group.Name)" -ForegroundColor Cyan
                foreach ($result in $group.Group) {
                    $color = switch ($result.Severity) {
                        "Error" { "Red" }
                        "Warning" { "Yellow" }
                        "Information" { "White" }
                        default { "Gray" }
                    }
                    Write-Host "  [$($result.Severity)] Line $($result.Line): $($result.Message)" -ForegroundColor $color
                }
            }
        }
    } catch {
        Write-Host "[ERROR] Failed to analyze scripts: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
}

function Generate-Report {
    param(
        [string]$ScriptPath,
        [bool]$Recurse
    )
    
    Write-Host "=== Generating Analysis Report ===" -ForegroundColor Cyan
    
    $reportFile = "powershell-analysis-report.txt"
    $scripts = Get-ChildItem -Path $ScriptPath -Filter "*.ps1" -Recurse:$Recurse
    
    $report = @()
    $report += "PowerShell Script Analysis Report"
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "=" * 50
    $report += ""
    
    $report += "Found $($scripts.Count) PowerShell scripts:"
    foreach ($script in $scripts) {
        $report += "  - $($script.FullName)"
    }
    $report += ""
    
    try {
        Import-Module PSScriptAnalyzer -Force
        
        foreach ($script in $scripts) {
            $report += "Analyzing: $($script.Name)"
            $report += "-" * 30
            
            $results = Invoke-ScriptAnalyzer -Path $script.FullName
            
            if ($results.Count -eq 0) {
                $report += "[OK] No issues found"
            } else {
                $report += "Found $($results.Count) issues:"
                foreach ($result in $results) {
                    $report += "  [$($result.Severity)] Line $($result.Line): $($result.Message)"
                }
            }
            $report += ""
        }
        
        # Summary
        $allResults = Invoke-ScriptAnalyzer -Path $ScriptPath -Recurse:$Recurse
        $report += "SUMMARY"
        $report += "=" * 20
        $report += "Total scripts analyzed: $($scripts.Count)"
        $report += "Total issues found: $($allResults.Count)"
        
        if ($allResults.Count -gt 0) {
            $severityCounts = $allResults | Group-Object Severity
            $report += ""
            $report += "Issues by severity:"
            foreach ($severityGroup in $severityCounts) {
                $report += "  $($severityGroup.Name): $($severityGroup.Count)"
            }
        }
        
        # Write report to file
        $report | Out-File -FilePath $reportFile -Encoding UTF8
        Write-Host "Analysis report saved to: $reportFile" -ForegroundColor Green
        
    } catch {
        Write-Host "[ERROR] Failed to generate report: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
}

function Check-SpecialCharacters {
    param(
        [string]$ScriptPath,
        [bool]$Recurse
    )
    
    Write-Host "=== Checking for Special Characters ===" -ForegroundColor Cyan
    
    $scripts = Get-ChildItem -Path $ScriptPath -Filter "*.ps1" -Recurse:$Recurse
    
    if ($scripts.Count -eq 0) {
        Write-Host "No PowerShell scripts found in $ScriptPath" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($scripts.Count) PowerShell script(s) to check" -ForegroundColor Green
    
    $totalIssues = 0
    foreach ($script in $scripts) {
        Write-Host "Checking: $($script.FullName)" -ForegroundColor Yellow
        
        try {
            $content = Get-Content $script.FullName -Raw
            $specialCharTest = Test-SpecialCharacters -Content $content
            
            if ($specialCharTest.HasSpecialChars) {
                $charDetails = $specialCharTest.SpecialChars | ForEach-Object { 
                    if ($_ -match '[\x00-\x1F\x7F-\x9F]') {
                        "[Control: $(($_ -as [int]))]"
                    } else {
                        "'$_'"
                    }
                }
                Write-Host "  [WARNING] Found special characters: $($charDetails -join ', ')" -ForegroundColor Red
                $totalIssues++
            } else {
                Write-Host "  [OK] No special characters found" -ForegroundColor Green
            }
        } catch {
            Write-Host "  [ERROR] Failed to check: $($script.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    if ($totalIssues -eq 0) {
        Write-Host "[OK] No special characters found in any scripts!" -ForegroundColor Green
    } else {
        Write-Host "[WARNING] Found special characters in $totalIssues script(s)" -ForegroundColor Yellow
        Write-Host "Use -Format to automatically remove special characters" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Main execution
if ($Help) {
    Show-Help
    exit 0
}

# Install required modules
Install-RequiredModules

# If no specific actions are requested, show help
if (-not $Format -and -not $Analyze -and -not $CheckSpecialChars) {
    Write-Host "No actions specified. Use -Help for usage information." -ForegroundColor Yellow
    Show-Help
    exit 0
}

# Perform requested actions
if ($CheckSpecialChars) {
    Check-SpecialCharacters -ScriptPath $Path -Recurse $Recurse
}

if ($Format) {
    Format-Scripts -ScriptPath $Path -Recurse $Recurse
}

if ($Analyze) {
    Analyze-Scripts -ScriptPath $Path -Recurse $Recurse
}

Write-Host "=== Completed ===" -ForegroundColor Green
