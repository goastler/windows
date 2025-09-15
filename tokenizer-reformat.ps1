# PowerShell tokenizer-based reformatter
param(
    [string]$InputFile,
    [string]$OutputFile
)

# Read the content
$content = Get-Content $InputFile -Raw

# Parse the script using PowerShell tokenizer
$tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$null)

$output = @()
$indentLevel = 0
$lastToken = $null

foreach ($token in $tokens) {
    $tokenText = $token.Content
    
    switch ($token.Type) {
        'Comment' {
            if ($lastToken -and $lastToken.Type -ne 'NewLine') {
                $output += "`n"
            }
            $output += ('    ' * $indentLevel) + $tokenText
        }
        'Keyword' {
            if ($tokenText -in @('function', 'if', 'else', 'elseif', 'foreach', 'while', 'try', 'catch', 'finally')) {
                if ($lastToken -and $lastToken.Type -ne 'NewLine') {
                    $output += "`n"
                }
                $output += ('    ' * $indentLevel) + $tokenText
            } else {
                $output += $tokenText
            }
        }
        'GroupStart' {
            $output += $tokenText
            if ($tokenText -eq '{') {
                $indentLevel++
                $output += "`n"
            }
        }
        'GroupEnd' {
            if ($tokenText -eq '}') {
                $indentLevel = [Math]::Max(0, $indentLevel - 1)
                if ($lastToken -and $lastToken.Type -ne 'NewLine') {
                    $output += "`n"
                }
                $output += ('    ' * $indentLevel) + $tokenText
            } else {
                $output += $tokenText
            }
        }
        'NewLine' {
            $output += "`n"
        }
        default {
            $output += $tokenText
        }
    }
    
    $lastToken = $token
}

# Join and clean up
$formatted = $output -join ''

# Clean up multiple newlines
$formatted = $formatted -replace '(\r?\n){3,}', "`n`n"

# Write the result
Set-Content -Path $OutputFile -Value $formatted -Encoding UTF8

Write-Host "Tokenizer reformatted $InputFile -> $OutputFile"
