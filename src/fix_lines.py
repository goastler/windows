#!/usr/bin/env python3

# Read the file
with open('packIso.ps1', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Fix the problematic lines
fixed_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    
    # Check if this is one of the problematic lines
    if 'GetEnvironmentVariable("Path", "Machine")' in line and 'GetEnvironmentVariable("Path", "User")' in line:
        # This is already a complete line, keep it
        fixed_lines.append(line)
    elif 'GetEnvironmentVariable("Path", "Machine")' in line and i + 1 < len(lines):
        # This is a split line, combine with the next line
        next_line = lines[i + 1]
        if 'GetEnvironmentVariable("Path", "User")' in next_line:
            # Combine the lines
            combined = line.rstrip() + ' ' + next_line.lstrip()
            fixed_lines.append(combined)
            i += 2  # Skip the next line
            continue
    
    fixed_lines.append(line)
    i += 1

# Write the fixed file
with open('packIso.ps1', 'w', encoding='utf-8') as f:
    f.writelines(fixed_lines)

print("Fixed the split lines in packIso.ps1")
