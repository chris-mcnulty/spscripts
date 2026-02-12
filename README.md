# spscripts
SharePoint and M365 Scripts

## Overview
This repository contains PowerShell scripts for managing and reporting on SharePoint Online and Microsoft 365 environments.

## Scripts

### Get-SPOSiteUsageReports.ps1
Enumerate usage reports from all SharePoint sites in a tenant and export to CSV.

**Quick Start:**
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
```

**Features:**
- Collects comprehensive usage statistics for all SharePoint sites
- Supports both SharePoint Online Management Shell and Microsoft Graph API
- Automatic module installation
- Exports to CSV with customizable output path
- Progress tracking and error handling

**[Full Documentation](README-Get-SPOSiteUsageReports.md)**

## Requirements
- PowerShell 5.1 or later
- Appropriate SharePoint/Microsoft 365 admin permissions
- Internet connection for module installation

## Getting Started

1. Clone this repository:
```powershell
git clone https://github.com/chris-mcnulty/spscripts.git
cd spscripts
```

2. Run the desired script with appropriate parameters

3. Required PowerShell modules will be automatically installed if missing

## Contributing
Contributions are welcome! Please feel free to submit pull requests or open issues.

## License
This project is provided as-is without warranty.
