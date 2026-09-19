# Autopilot OOBE Command Hub

Enterprise Provisioning & Endpoint Deployment Engine for Windows Setup (OOBE Shift+F10).

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%20%7C%207+-blue.svg)](https://microsoft.com)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-lightgrey.svg)](https://microsoft.com/windows)
[![Design](https://img.shields.io/badge/UI-WinUI%203%20Fluent%20Dark-0067C0.svg)](https://learn.microsoft.com/windows/apps/design/)

---

## Overview

The **Autopilot OOBE Command Hub** is an interactive, standalone provisioning console engineered for rapid technician deployment during Windows Out-of-Box Experience (`Shift` + `F10`). It unifies hardware diagnostics, OA3 hardware hash harvesting, direct Microsoft Intune enrollment via Microsoft Graph, batch application installation, Win32 packaging, and hardware refresh warranty auditing into a single Fluent dark WPF application.

### Key Capabilities
- **Autopilot & Cloud Registration**: Harvester with ASN.1 DER hardware hash validation, CSV export (with automatic USB drive detection), GroupTag and AssignedUser tagging, dynamic computer renaming, and direct Microsoft Intune registration.
- **Root-Level Graph Authentication**: Interactive OAuth 2.0 Device Code Flow (`https://login.microsoft.com/device`) using universal Azure PowerShell client (`1950a258-227b-4e31-a9cf-717495945fc2`) and Client Secret service principal authentication.
- **PowerShell 7 One-Click Installer**: Inbox header shortcut to upgrade from Windows PowerShell 5.1 to PowerShell 7 via winget or direct MSI download.
- **Application Deployment Hub**: Multi-select software installer for Browsers, Developer Tools, Productivity, and System Utilities.
- **Win32 App Packaging**: Compile `.intunewin` packages and publish directly to Intune.
- **Pre-Flight Diagnostics**: 7-stage network and hardware health ladder (Interface -> Gateway -> DNS -> TLS 1.3 -> Clock Sync -> TPM 2.0 -> Secure Boot).
- **Dell Warranty & Fleet Refresh Lifecycle**: Direct Dell Enterprise API (v5) query to assess asset warranty SLA and recommend deployment eligibility or refresh decommissioning.

---

## File Structure

```
C:\src\AutopilotCommandHub\
├── autopilot.ps1            # Standalone, zero-dependency production script
├── autopilot                # Extensionless bootstrap version for web deployment
├── build_autopilot.py       # Source compiler embedding all XAML, styling, and modules
├── Start-DeviceAuth.ps1     # Standalone CLI device code authenticator
├── Get-DellWarranty.ps1     # Standalone Dell asset warranty query tool
├── test_dell_warranty.py    # Automated verification test suite
├── render_ui.py             # UI rendering and preview tool
├── .env.example             # Configuration defaults template
├── .env                     # Local environment settings (gateway, charset, prefix)
└── media\                   # Screenshots and architecture assets
```

## Running the Hub

### From Local File:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\src\AutopilotCommandHub\autopilot.ps1"
```

### CLI / Headless Modes:
```powershell
# Harvest hardware hash only
powershell -File .\autopilot.ps1 -HarvestOnly

# Export Autopilot CSV to USB flash drive
powershell -File .\autopilot.ps1 -ExportCsv

# Check Dell warranty & refresh status
powershell -File .\autopilot.ps1 -DellWarranty -DellServiceTag "6BYQJW2"
```

### Compile from Source:
```powershell
python build_autopilot.py
```
