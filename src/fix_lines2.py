#!/usr/bin/env python3

# Read the file
with open('packIso.ps1', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the split lines by replacing them with single lines
fixes = [
    # Fix line 165-166
    ('        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +\n [System.Environment]::GetEnvironmentVariable("Path", "User")', 
     '        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")'),
    
    # Fix line 220-221  
    ('            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +\n [System.Environment]::GetEnvironmentVariable("Path", "User")',
     '            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")'),
]

# Apply fixes
for old, new in fixes:
    content = content.replace(old, new)

# Write the fixed file
with open('packIso.ps1', 'w', encoding='utf-8') as f:
    f.write(content)

print("Fixed the split lines in packIso.ps1")
