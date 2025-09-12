# Windows ISO Customization Project

This repository contains tools and scripts for customizing Windows ISO files, primarily focused on creating unattended installation media.

## Project Overview

This project provides a comprehensive solution for:
- **Automated Windows ISO Processing**: Unpack, modify, and repack Windows installation media
- **Unattended Installation Setup**: Configure Windows to install without user interaction
- **Custom Configuration**: Add drivers, applications, and system modifications
- **Deployment Automation**: Streamline Windows deployment processes

## Quick Start

For the main ISO repacking script, see the dedicated documentation:

📖 **[packIso-README.md](packIso-README.md)** - Complete guide for the packIso.ps1 script

## Project Structure

```
windows/
├── src/
│   ├── packIso.ps1          # Main PowerShell script for ISO repacking
│   ├── autounattend.xml     # Default unattended installation configuration
│   ├── setup.ps1            # Post-installation configuration script
│   └── BgInfo.bgi           # Background information configuration
├── $OEM$/                   # Additional files for Windows setup
│   └── $1/Windows/Setup/Scripts/
│       └── FirstLogon.ps1   # First logon configuration script
├── packIso-README.md       # Detailed documentation for packIso.ps1
├── TODO.md                 # Development progress and future plans
└── README.md               # This general project overview
```

## Main Components

### 1. packIso.ps1 Script
The core script that handles:
- ISO mounting and extraction
- Adding autounattend.xml for unattended installation
- Repacking modified ISO files
- Automatic tool installation (Chocolatey, Windows ADK)

### 2. autounattend.xml
Comprehensive unattended installation configuration including:
- System bypasses (TPM, Secure Boot, RAM checks)
- User account creation and configuration
- Application removal and system optimization
- Network and security settings
- Custom scripts and configurations

### 3. Post-Installation Scripts
- **setup.ps1**: System configuration and optimization
- **FirstLogon.ps1**: Initial user setup and customization
- **BgInfo.bgi**: System information display configuration

## Features

- **Automated ISO Processing**: Mount, extract, modify, and repack Windows ISOs
- **Unattended Installation**: Fully automated Windows setup
- **System Optimization**: Remove bloatware, configure performance settings
- **Custom Configuration**: Add drivers, applications, and system modifications
- **Error Handling**: Comprehensive error checking and user guidance
- **Flexible Options**: Customizable paths, configurations, and installation options

## Prerequisites

- **Windows 10/11** or **Windows Server 2016/2019/2022**
- **PowerShell 5.1 or later**
- **Administrator privileges**
- **Internet connection** (for automatic tool installation)
- **Sufficient disk space** (at least 2x the size of your ISO file)

## Getting Started

1. **Clone or download** this repository
2. **Run PowerShell as Administrator**
3. **Review the packIso-README.md** for detailed usage instructions
4. **Place your Windows ISO** in an accessible location
5. **Run the packIso.ps1 script** with your desired parameters

## Documentation

- **[packIso-README.md](packIso-README.md)** - Complete guide for the main script
- **[TODO.md](TODO.md)** - Development progress and future enhancements

## Contributing

We welcome contributions! Please see the TODO.md file for current development priorities and ideas for future enhancements.

## License

This project is provided as-is for educational and automation purposes. Please ensure you comply with Microsoft's licensing terms when using Windows ISO files.

## Support

For issues or questions:
1. Check the packIso-README.md troubleshooting section
2. Review the TODO.md for known issues and planned fixes
3. Ensure you're running as Administrator
4. Verify all prerequisites are met