# Advanced PowerShell script reformatter
param(
    [string]$InputFile,
    [string]$OutputFile
)

# Read the corrupted content
$content = Get-Content $InputFile -Raw

# Split on function boundaries and reformat each section
$sections = $content -split '(?=function\s+[\w-]+\s*\{)'

$reformattedSections = @()

foreach ($section in $sections) {
    if ($section.Trim() -eq '') { continue }
    
    $reformatted = $section
    
    # Fix basic structure
    $reformatted = $reformatted -replace '(\w+)\s*\{', '$1 {' # Space before {
    $reformatted = $reformatted -replace '\}\s*([a-zA-Z])', "`}`n`n`$1" # Line break after }
    $reformatted = $reformatted -replace '\s+', ' ' # Normalize whitespace
    $reformatted = $reformatted -replace '\s*#', "`n#" # Comments on new lines
    $reformatted = $reformatted -replace 'param\s*\(', "param(`n    " # Param formatting
    $reformatted = $reformatted -replace '\[Parameter\(', "`n    [Parameter(" # Parameter attributes
    $reformatted = $reformatted -replace '\[string\]', "`n    [string]" # Type attributes
    $reformatted = $reformatted -replace '\[int\]', "`n    [int]" # Type attributes
    $reformatted = $reformatted -replace '\[switch\]', "`n    [switch]" # Type attributes
    $reformatted = $reformatted -replace '\[ValidateScript\(', "`n    [ValidateScript(" # Validation
    $reformatted = $reformatted -replace '\[ValidateSet\(', "`n    [ValidateSet(" # Validation
    $reformatted = $reformatted -replace '\)\s*\$ErrorActionPreference', ")`n`n`$ErrorActionPreference"
    $reformatted = $reformatted -replace 'function\s+', "`n`nfunction " # Functions
    $reformatted = $reformatted -replace '\s*try\s*\{', "`n    try {" # Try blocks
    $reformatted = $reformatted -replace '\s*catch\s*\{', "`n    } catch {" # Catch blocks
    $reformatted = $reformatted -replace '\s*finally\s*\{', "`n    } finally {" # Finally blocks
    $reformatted = $reformatted -replace '\s*if\s*\(', "`n        if (" # If statements
    $reformatted = $reformatted -replace '\s*else\s*\{', " else {" # Else statements
    $reformatted = $reformatted -replace '\s*foreach\s*\(', "`n        foreach (" # Foreach loops
    $reformatted = $reformatted -replace '\s*while\s*\(', "`n        while (" # While loops
    $reformatted = $reformatted -replace 'Write-Log\s+"', "`n        Write-Log `"" # Write-Log calls
    $reformatted = $reformatted -replace 'throw\s+"', "`n        throw `"" # Throw statements
    $reformatted = $reformatted -replace 'return\s+', "`n        return " # Return statements
    
    $reformattedSections += $reformatted
}

# Join sections and clean up
$final = $reformattedSections -join "`n"

# Clean up excessive whitespace and fix indentation issues
$lines = $final -split "`n"
$indentLevel = 0
$formattedLines = @()

foreach ($line in $lines) {
    $trimmedLine = $line.Trim()
    
    if ($trimmedLine -eq '') {
        $formattedLines += ''
        continue
    }
    
    # Adjust indent level based on braces
    if ($trimmedLine -match '^\}') {
        $indentLevel = [Math]::Max(0, $indentLevel - 1)
    }
    
    # Add proper indentation
    $indent = '    ' * $indentLevel
    $formattedLines += $indent + $trimmedLine
    
    # Increase indent for opening braces
    if ($trimmedLine -match '\{$') {
        $indentLevel++
    }
}

# Join and write
$finalFormatted = $formattedLines -join "`n"

# Write the formatted content
Set-Content -Path $OutputFile -Value $finalFormatted -Encoding UTF8

Write-Host "Advanced reformatted $InputFile -> $OutputFile"
