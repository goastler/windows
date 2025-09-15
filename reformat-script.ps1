# Script to reformat corrupted PowerShell scripts
param(
    [string]$InputFile,
    [string]$OutputFile
)

# Read the corrupted content
$content = Get-Content $InputFile -Raw

# Define patterns for adding line breaks
$patterns = @(
    @{ Pattern = '\s*#'; Replace = "`n#" }
    @{ Pattern = '\s*param\s*\('; Replace = "`nparam(`n    " }
    @{ Pattern = '\)\s*\$ErrorActionPreference'; Replace = ")`n`n`$ErrorActionPreference" }
    @{ Pattern = '\s*function\s+'; Replace = "`n`nfunction " }
    @{ Pattern = '\s*try\s*\{'; Replace = "`n    try {`n        " }
    @{ Pattern = '\s*catch\s*\{'; Replace = "`n    } catch {`n        " }
    @{ Pattern = '\s*finally\s*\{'; Replace = "`n    } finally {`n        " }
    @{ Pattern = '\}\s*\}'; Replace = "}`n}" }
    @{ Pattern = '\s*if\s*\('; Replace = "`n        if (" }
    @{ Pattern = '\s*else\s*\{'; Replace = " else {`n            " }
    @{ Pattern = '\s*foreach\s*\('; Replace = "`n        foreach (" }
    @{ Pattern = '\s*while\s*\('; Replace = "`n        while (" }
    @{ Pattern = '\}\s*Write-Log'; Replace = "}`n        Write-Log" }
    @{ Pattern = '\}\s*return'; Replace = "}`n        return" }
    @{ Pattern = '\}\s*throw'; Replace = "}`n        throw" }
)

$formatted = $content

# Apply formatting patterns
foreach ($pattern in $patterns) {
    $formatted = $formatted -replace $pattern.Pattern, $pattern.Replace
}

# Clean up excessive whitespace
$formatted = $formatted -replace '\n\s*\n\s*\n', "`n`n"

# Write the formatted content
Set-Content -Path $OutputFile -Value $formatted -Encoding UTF8

Write-Host "Reformatted $InputFile -> $OutputFile"
