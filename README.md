# Autopilot OOBE Command Hub

> **Enterprise Provisioning & Endpoint Deployment Engine for Windows Setup (`Shift` + `F10`)**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%20%7C%207+-0078D4.svg?logo=powershell&logoColor=white)](https://microsoft.com)
[![Platform: Windows 10 / 11](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20%7C%20Server-00A4EF.svg?logo=windows&logoColor=white)](https://microsoft.com/windows)
[![UI: WinUI 3 Fluent Dark](https://img.shields.io/badge/UI-WinUI%203%20Fluent%20Dark-202020.svg?logo=windows11&logoColor=60CDFF)](https://learn.microsoft.com/windows/apps/design/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Zero Dependencies](https://img.shields.io/badge/Dependencies-Zero%20External%20Modules-brightgreen.svg)](https://github.com/thebubbsy/AutopilotCommandHub)

---

![Autopilot Command Hub UI](media/autopilotfast-oobe.png)

---

## ⚡ Instant Bootstrapping (OOBE Shift + F10)

During Windows Out-of-Box Experience (OOBE), press `Shift` + `F10` to open Command Prompt, then run:

```powershell
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/thebubbsy/AutopilotCommandHub/main/autopilot.ps1 | iex"
```

Or from an existing clone:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\src\AutopilotCommandHub\autopilot.ps1"
```

---

## 🌟 Executive Overview

**Autopilot OOBE Command Hub** is an all-in-one, standalone technician provisioning console designed for IT administrators and field engineers setting up endpoints. It packages hardware telemetry harvesting, Microsoft Graph cloud registration, batch application installation, Win32 packaging, 7-stage pre-flight diagnostics, and vendor hardware refresh lifecycle auditing into a single, zero-dependency Fluent Dark WPF application.

### Why Autopilot Command Hub?
* **Zero External Dependencies**: Does not require `Microsoft.Graph`, `AzureAD`, or external PowerShell gallery modules. Operates 100% natively in inbox Windows PowerShell 5.1 and modern PowerShell 7+.
* **Native Single-Threaded Apartment (STA) Engine**: Built on a non-blocking `DispatcherFrame` message loop that streams live color-coded logs without freezing the interface.
* **WinUI 3 Fluent Dark Architecture**: Crafted strictly with authentic Windows 11 dark theme design tokens (`#202020` canvas, `#2B2B2B` elevation surfaces, and `#0067C0` accent blue).
* **Cross-Subsystem Core Integration**: Combines core engines from [`AutopilotFast`](https://github.com/thebubbsy/AutopilotFast), [`IntuneShared`](https://github.com/thebubbsy/IntuneShared), and [`WingetIntune`](https://github.com/thebubbsy/WingetIntune).

---

## 🚀 Core Functional Hubs

### 1. Autopilot & Cloud Registration
* **High-Speed Hardware Hash Harvester**: Queries WMI `MDM_DevDetail_Ext01` and validates the genuine OA3 blob (magic `4F 41 33 00` = Base64 prefix `T0EzAA`, 2,048–16,384 bytes). The hub never fabricates a hash: if the provider cannot be read (not elevated, VM without OA3), it reports `HardwareHashStatus = Unavailable` with the reason and blocks Intune registration and CSV export.
* **Direct Intune Cloud Upload**: Imports hardware identities directly via `https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities`.
* **Smart CSV Exporter**: Automatically detects connected USB flash drives (`Win32_Volume` DriveType 2) and saves the standardized Microsoft Intune CSV format.
* **Dynamic Computer Renaming**: Resolves computer naming patterns (e.g. `WS-%SERIAL%` or `LT-%SERIAL%`) and applies them with a single click.

### 1b. Session Awareness: Privilege Badge, Device State & Restart Persistence
The Hub is built to run in two very different places - the OOBE `Shift+F10` prompt (SYSTEM) and a normal desktop (usually a standard user) - so it tells you where it is and what it can do, all the time:

* **Always-visible privilege badge** (`PRIV: SYSTEM | OOBE`, `PRIV: ADMINISTRATOR | DESKTOP`, `PRIV: STANDARD USER | DESKTOP`). Below the preferred level a red **Fix Privileges** button appears; clicking it (or the badge) opens a guide showing current vs. preferred level, why it matters, and the exact steps. On the desktop it offers **Relaunch as Administrator**: the running script is persisted to `%ProgramData%\AutopilotCommandHub` (works even when started via `irm | iex`), relaunched through UAC with the same `.env`, and the unprivileged window closes itself.
* **Device state on launch** - before you touch anything the Hub answers "is this PC already someone's?" from local evidence: the Autopilot profile the device fetched from the Deployment Service during OOBE (`AutopilotDDSZTDFile.json` / `Provisioning\Diagnostics\AutoPilot` - the thing that makes a registered PC boot into the branded OOBE), Intune MDM enrollment (`Enrollments\*` with provider `MS DM Server`) and `dsregcmd` Entra/domain join. Verdicts: **Autopilot registered** (no action needed, Register asks for confirmation), **Intune enrolled but no Autopilot profile**, **Entra joined**, **Domain joined**, or **Not registered** (harvest + register). Elevated sessions harvest the hash automatically at start.
* **Tenant-side lookup** - once signed in to Graph, the serial is checked against `windowsAutopilotDeviceIdentities` in *that* tenant and the banner is updated with group tag, enrollment state and last contact. (Only the device itself, during OOBE, can ask "which tenant owns me?" across all tenants - and its answer is exactly the cached profile above. Graph only sees the tenant you signed in to.)
* **Restart with persistence** - **Restart System** now asks *Yes = restart and re-open the Hub when OOBE / the desktop comes back*. It persists the script + `.env`, registers a one-shot logon-triggered scheduled task in the interactive session (`Administrators` group during OOBE, the current user on the desktop, `RunLevel Highest`) plus a `RunOnce` fallback, then restarts. The relaunched Hub (`-ResumeFromRestart`) removes both itself; a named mutex prevents a double launch.
* **App Deployment advisory** - on managed devices the App tab explains that app assignment belongs to Intune and this tab is for bench builds and one-offs.
* **Skip Autopilot in OOBE** - a first-class button (Autopilot & Cloud Registration tab) that finishes OOBE as a normal Windows with a local account: no Autopilot, no forced enrollment. Run it from Shift+F10 at the first OOBE screen, *before* connecting to a network. It removes any cached Autopilot profile, sets `Provisioning\Diagnostics\AutoPilot\IsAutopilotDisabled=1` and `OOBE\BypassNRO=1`, optionally disables the physical NICs for the rest of OOBE (with a **Re-enable Network Adapters** button), and launches the `ms-cxh:localonly` local-account page on 24H2+. The one rule that matters: the Autopilot Deployment Service must not be reachable while OOBE runs. The bypass method is intentionally simple to swap - if Microsoft changes the offline path, only `Invoke-AutopilotBypass` in `build_autopilot.py` needs updating.

### 2. Microsoft Graph & Entra ID Device Code Authentication
* **Interactive Device Code Flow**: Directly acquires tokens via RFC 8628 Device Authorization with universal pre-consented Azure PowerShell client ID (`1950a258-227b-4e31-a9cf-717495945fc2`).
* **One-Click Browser Launcher**: Dedicated **Open Browser** button alongside automated clipboard copy of the user code.
* **Silent `.env` Credential Fallback**: Supports automated unattended client credentials via `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET`.

### 3. Application Deployment Hub
* **Curated App Bundles**: One-click multi-select app queues across:
  * **Browsers**: Google Chrome, Mozilla Firefox, Brave, Microsoft Edge Dev
  * **Developer Tools**: Visual Studio Code, Git, Windows Terminal, Node.js, Python 3, GitHub CLI, Docker Desktop
  * **Productivity**: Microsoft 365 Office Suite, Teams, Slack, Zoom, Obsidian
  * **Utilities**: 7-Zip, Notepad++, VLC Media Player, Voidtools Everything, PowerToys
* **Custom Winget Package Queue**: Add any custom package ID (e.g., `WiresharkFoundation.Wireshark`) for batch unattended deployment.

### 4. Dell Enterprise Asset Warranty & Refresh Lifecycle (eAPI v5)
* **BIOS Service Tag Auto-Detection**: Interrogates `Win32_BIOS` to identify Dell service tags in <50ms.
* **Autonomous Refresh Verdict**:
  * 🟢 **ACTIVE WARRANTY (ELIGIBLE FOR DEPLOYMENT)**
  * 🟡 **REFRESH PLANNING (EXPIRING WITHIN 90 DAYS)**
  * 🔴 **REFRESH RECOMMENDED (DECOMMISSION / OUT OF WARRANTY)**
* **Detailed SLA Telemetry**: System model, product line, factory ship date, device age in years/days, and complete contract entitlements table.
* **Export Options**: Copy Markdown audit report to clipboard or export audit CSV to USB.

### 5. 7-Stage Pre-Flight Network & Hardware Ladder
1. **Network Interface**: Verifies active physical adapter is in `Up` status.
2. **Default Gateway**: Pings dynamic default gateway (or fallback router IP).
3. **DNS Resolution**: Validates public name resolution against Cloudflare/Google DNS.
4. **TLS 1.2 / 1.3 Handshake**: Tests secure HTTPS cryptographic tunnel to Microsoft endpoints.
5. **Clock Drift Sync**: Checks local machine clock against NIST / Windows Time servers.
6. **TPM 2.0 Security Module**: Verifies Trusted Platform Module is present, enabled, and ready for attestation.
7. **Secure Boot State**: Confirms UEFI Secure Boot policy is actively enforced.

---

## 💻 Non-GUI / CLI Automation Modes

Autopilot Command Hub can also be invoked headless in automated build pipelines, batch files, or USB setup scripts:

```powershell
# 1. Harvest OA3 hardware hash and print to console
powershell -File .\autopilot.ps1 -HarvestOnly

# 2. Export Intune CSV with automatic USB flash drive detection
powershell -File .\autopilot.ps1 -ExportCsv -GroupTag "Engineering" -AssignedUser "user@contoso.com"

# 3. Perform a Dell Asset Warranty assessment
powershell -File .\autopilot.ps1 -DellWarranty -DellServiceTag "6BYQJW2"

# 4. Check Dell Warranty and export result to CSV
powershell -File .\autopilot.ps1 -DellWarranty -ExportCsv -CsvPath "C:\temp\WarrantyReport.csv"

# 5. Rename computer using template
powershell -File .\autopilot.ps1 -RenameComputer -ComputerNamePrefix "WS" -ComputerNameTemplate "WS-%SERIAL%"

# 6. (internal) What the restart-resume scheduled task runs - cleans itself up on start
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "C:\ProgramData\AutopilotCommandHub\autopilot.ps1" -EnvFile "C:\ProgramData\AutopilotCommandHub\.env" -ResumeFromRestart
```

---

## 🛠️ Repository Architecture

```
AutopilotCommandHub/
├── autopilot.ps1            # Standalone, zero-dependency production script
├── autopilot                # Extensionless bootstrap version for web distribution
├── build_autopilot.py       # Source compiler (embeds XAML, styling, and modules)
├── Start-DeviceAuth.ps1     # Standalone CLI device code authenticator
├── Get-DellWarranty.ps1     # Standalone Dell asset warranty query tool
├── test_dell_warranty.py    # Automated verification test suite
├── render_ui.py             # UI rendering and preview tool
├── .env.example             # Configuration defaults template
├── .gitignore               # Secret and temp cache quarantine rules
└── media/
    └── autopilotfast-oobe.png # Hero interface preview
```

---

## ⚙️ Compiling from Source

To compile modifications to the embedded XAML, styling tokens, or PowerShell engines:

```powershell
python build_autopilot.py
```

This compiles `build_autopilot.py` into both `autopilot.ps1` and extensionless `autopilot` in `< 100ms`.

---

## 📄 License & Attribution

- **Author**: Matthew Bubb ([thebubbsy](https://github.com/thebubbsy))
- **Website**: [onyachamp.com](https://onyachamp.com)
- **License**: [MIT License](LICENSE)
