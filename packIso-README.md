# Windows ISO Repack Script (packIso.ps1)

This PowerShell script automates the process of unpacking a Windows ISO, adding an autounattend.xml file for unattended installation, and repacking it into a new ISO file.

## Features

- **Automated ISO Processing**: Mounts, extracts, modifies, and repacks Windows ISO files
- **Unattended Installation**: Adds autounattend.xml for automated Windows setup
- **Automatic Tool Installation**: Automatically installs Chocolatey and Windows ADK if needed
- **Chocolatey Integration**: Uses Chocolatey as the primary package manager for reliable installations
- **Error Handling**: Comprehensive error checking and cleanup
- **Flexible Configuration**: Customizable paths and options
- **Administrator Privileges**: Automatically checks for required permissions

## Prerequisites

### Required Software
1. **Windows Assessment and Deployment Kit (ADK)** - Contains `oscdimg.exe` tool
   - **Automatically installed** by the script via Chocolatey (requires internet connection)
   - Manual download from: https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install

2. **PowerShell 5.1 or later**
3. **Administrator privileges** (required for mounting/unmounting ISO files and installing tools)
4. **Internet connection** (for downloading Chocolatey and Windows ADK)

### Package Manager
- **Chocolatey** - Automatically installed by the script if not present
  - Used for installing Windows ADK
  - Reliable and well-maintained package manager for Windows

### System Requirements
- Windows 10/11 or Windows Server 2016/2019/2022
- Sufficient disk space (at least 2x the size of your ISO file)
- .NET Framework 4.5 or later

## Installation

1. Clone or download this repository
2. **Run PowerShell as Administrator** (required for the script to function)
3. Place your `autounattend.xml` file in the same directory as the script (or specify a custom path)
4. The script will automatically install Chocolatey and Windows ADK if needed (requires internet connection)

## Usage

### Basic Usage
**IMPORTANT: Run PowerShell as Administrator first!**

```powershell
# Right-click PowerShell and select "Run as Administrator", then:
.\packIso.ps1 -InputIso "C:\ISOs\Windows11.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso"
```

### Advanced Usage
```powershell
# Run PowerShell as Administrator first, then:
.\packIso.ps1 -InputIso "C:\ISOs\Windows11.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso" -AutounattendXml "C:\Custom\autounattend.xml" -KeepWorkingDirectory
```

### Parameters

| Parameter | Required | Description | Default |
|-----------|----------|-------------|---------|
| `InputIso` | Yes | Path to the input Windows ISO file | - |
| `OutputIso` | Yes | Path where the modified ISO will be created | - |
| `AutounattendXml` | No | Path to the autounattend.xml file | `autounattend.xml` in script directory |
| `WorkingDirectory` | No | Temporary directory for extraction | Auto-generated temp directory |
| `KeepWorkingDirectory` | No | Keep working directory after completion | False |
| `SkipAutoInstall` | No | Skip automatic Windows ADK installation | False |

## Examples

### Example 1: Basic repack with default autounattend.xml
```powershell
# Run PowerShell as Administrator first, then:
.\packIso.ps1 -InputIso "C:\Downloads\Windows11_22H2.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso"
```

### Example 2: Custom autounattend.xml and keep working files
```powershell
# Run PowerShell as Administrator first, then:
.\packIso.ps1 -InputIso "C:\Downloads\Windows11_22H2.iso" -OutputIso "C:\ISOs\Windows11_Custom.iso" -AutounattendXml "C:\Configs\my_autounattend.xml" -KeepWorkingDirectory
```

### Example 3: Specify custom working directory
```powershell
# Run PowerShell as Administrator first, then:
.\packIso.ps1 -InputIso "C:\Downloads\Windows11_22H2.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso" -WorkingDirectory "D:\Temp\IsoRepack"
```

### Example 4: Skip automatic ADK installation
```powershell
# Run PowerShell as Administrator first, then:
.\packIso.ps1 -InputIso "C:\Downloads\Windows11_22H2.iso" -OutputIso "C:\ISOs\Windows11_Unattended.iso" -SkipAutoInstall
```

## How It Works

1. **Validation**: Checks for administrator privileges, required tools, and input files
2. **Chocolatey Installation**: Automatically installs Chocolatey if not present
3. **Tool Installation**: Automatically installs Windows ADK via Chocolatey if oscdimg.exe is missing
4. **ISO Mounting**: Mounts the input ISO file to access its contents
5. **Extraction**: Copies all files from the mounted ISO to a temporary directory
6. **Modification**: Adds the autounattend.xml file to the root of the extracted contents
7. **Repacking**: Uses oscdimg.exe to create a new ISO file from the modified contents
8. **Cleanup**: Unmounts the original ISO and cleans up temporary files

## Troubleshooting

### Common Issues

**"oscdimg.exe not found"**
- The script will automatically install Chocolatey and then Windows ADK
- If automatic installation fails, follow the manual installation instructions provided
- Ensure the deployment tools are included in the installation
- The script will search common installation paths automatically

**"Access denied" or "Must run as Administrator"**
- Right-click PowerShell and select "Run as Administrator"
- The script requires admin privileges to mount/unmount ISO files and install Windows ADK

**"Input ISO file does not exist"**
- Verify the path to your input ISO file
- Ensure the file has a .iso extension
- Check that the file is not corrupted

**"Failed to mount ISO"**
- Ensure the ISO file is not already mounted
- Check that the ISO file is not corrupted
- Try running the script again after a few minutes

**"Robocopy failed"**
- Ensure sufficient disk space (at least 2x the ISO size)
- Check that the working directory is writable
- Verify no antivirus software is blocking the operation

**"Chocolatey installation failed"**
- Ensure you have an internet connection
- Check if your firewall or antivirus is blocking the download
- Use the `-SkipAutoInstall` parameter to skip automatic installation
- Install Chocolatey manually: `Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))`
- Then install Windows ADK manually: `choco install windows-adk -y`

### Performance Tips

- Use an SSD for better performance during extraction and repacking
- Ensure at least 8GB of free RAM for large ISO files
- Close unnecessary applications to free up system resources
- Use the `-KeepWorkingDirectory` parameter to debug issues

## File Structure

```
windows/
├── src/
│   ├── packIso.ps1          # Main PowerShell script
│   └── autounattend.xml     # Default unattended installation configuration
├── $OEM$/                   # Additional files for Windows setup
├── packIso-README.md       # This documentation
└── README.md               # General project documentation
```

## Security Notes

- The script requires administrator privileges to function
- Always verify the integrity of your input ISO files
- The autounattend.xml file may contain sensitive information (passwords, keys)
- Keep your working directories secure if using `-KeepWorkingDirectory`

## License

This project is provided as-is for educational and automation purposes. Please ensure you comply with Microsoft's licensing terms when using Windows ISO files.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this script.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify all prerequisites are met
3. Ensure you're running as Administrator
4. Check Windows Event Logs for additional error details
