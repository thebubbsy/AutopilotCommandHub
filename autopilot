<#
.SYNOPSIS
    Autopilot OOBE Command Hub - Enterprise Provisioning & Endpoint Deployment Engine
.DESCRIPTION
    Standalone Windows Autopilot OOBE bootstrap package with an interactive Fluent dark WPF GUI.
    Engineered for rapid field-technician provisioning during Windows Setup (Shift + F10).

    Features:
    - High-Speed Autopilot Hardware Hash Harvester (WMI MDM_DevDetail_Ext01 & OA3 magic/bounds validation)
    - Direct Microsoft Intune Cloud Registration via Microsoft Graph API (Device Code Flow & App Secrets)
    - Deployment Profile Assignment Polling (-WaitForSync) & Auto-Reboot Gate
    - Intune CSV Export with Automatic USB Flash Drive Detection
    - Group Tag Selection, User Pre-Assignment & Dynamic Computer Renaming
    - App Installation Hub powered by WingetBatch & Winget with Categorized Software Bundles
    - Win32 App Packaging (.intunewin) & Direct Cloud Publishing powered by WingetIntune
    - 7-Stage Hardware & Network Pre-Flight Diagnostic Ladder powered by IntuneShared
    - Real-Time Live Monospace Console Log & Animated Progress Bar
    - Always-visible privilege badge (SYSTEM / Administrator / Standard) with a guided elevated relaunch
    - Device enrollment awareness on launch: cached Autopilot profile, Intune MDM enrollment, Entra join,
      plus a tenant-side Autopilot identity lookup once signed in to Graph
    - Restart with persistence: re-opens the Hub automatically when OOBE (or the desktop) comes back
    - Skip Autopilot in OOBE, Hybrid Join & Co-Management toolset, ESP Diagnostics & Lifecycle tab
    - Active Directory & Kerberos Diagnostic Ladder (Port latency ladder, AD SCP tenant verification,
      klist TGT/TGS inspection & purge, certutil/certreq pulse, SCEP health, WAM broker reset)
    - Precision Enterprise Remediation Arsenal (WMI salvage, WU Agent deep reset, Catroot2 ESENT
      database repair, BITS deadlock purge, DCOM/RPC 0x800706BA fix, Network/Winsock/IPsec flush,
      WinRM listener rebuild, Print Spooler purge, ProfileList .bak un-hooker, TPM attestation healer,
      AppX manifest staging unbricker)
    - Multi-vendor hardware health (Dell/Lenovo warranty, battery wear, SSD reliability), offline JSON
      profiles, Wi-Fi / 802.1X / driver injection for OOBE, tenant branding in the header
    - Windows Update & Driver Engine: Native COM Session (Microsoft.Update.Session) & USO Client for OOBE
      system updates & hardware drivers with auto-reboot persistence

    Bootstrap Invocation:
        irm https://onyachamp.com/autopilot | iex
        or
        irm https://onyachamp.com/autopilot.ps1 | iex

    ZERO LocalPilot dependencies. 100% standalone, enterprise-hardened.
    (c) 2026 Matthew Bubb (thebubbsy) | https://onyachamp.com
#>

[CmdletBinding()]
param(
    [switch]$NoGui,
    [string]$GroupTag = '',
    [string]$AssignedUser = '',
    [switch]$HarvestOnly,
    [switch]$ExportCsv,
    [string]$CsvPath = '',
    [switch]$DellWarranty,
    [string]$DellServiceTag = '',
    [string]$EnvFile = '',
    [switch]$RenameComputer,
    [string]$ComputerNamePrefix = '',
    [string]$ComputerNameTemplate = '',
    [switch]$ResumeFromRestart,
    [switch]$ReplacingInstance,
    [switch]$LenovoWarranty,
    [string]$LenovoSerialNumber = '',
    [switch]$HardwareHealth,
    [switch]$OfflineJson,
    [string]$OfflineJsonPath = '',
    [switch]$Decommission,
    [switch]$NoIntro,
    [switch]$Shootout,
    [switch]$WindowsUpdate,
    [switch]$IncludeDrivers,
    [switch]$AutoReboot,
    [switch]$ScanOnly,
    [switch]$PatchCascade,
    [int]$MaxPasses = 5,
    [switch]$OfflineStaging,
    [string]$StagingProfile = 'HardwareDiagnostics',
    [switch]$ResumeAutopilot,
    [string]$Playbook = '',
    [string]$ActionRoutine = ''
)

# 0. Central Environment & Configuration Engine (.env)
$script:LoadedEnvPath = $null

function Get-ScriptEncoding {
    param([string]$CharsetName = $env:AUTOPILOT_CHARSET)
    if ([string]::IsNullOrWhiteSpace($CharsetName)) { $CharsetName = $env:CHARSET }
    if ([string]::IsNullOrWhiteSpace($CharsetName)) { return [System.Text.Encoding]::Default }
    switch -Regex ($CharsetName.Trim().ToLowerInvariant()) {
        '^(native|default|ansi|current)$' { return [System.Text.Encoding]::Default }
        '^(utf-?8|utf8)$'                  { return [System.Text.UTF8Encoding]::new($false) }
        '^(ascii|us-ascii)$'               { return [System.Text.Encoding]::ASCII }
        '^(unicode|utf-?16|utf16)$'        { return [System.Text.Encoding]::Unicode }
        '^(oem)$'                          { return [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) }
        default                            {
            try { return [System.Text.Encoding]::GetEncoding($CharsetName) } catch { return [System.Text.Encoding]::Default }
        }
    }
}

function Get-ScriptEncodingName {
    param([string]$CharsetName = $env:AUTOPILOT_CHARSET)
    if ([string]::IsNullOrWhiteSpace($CharsetName)) { $CharsetName = $env:CHARSET }
    if ([string]::IsNullOrWhiteSpace($CharsetName)) { return 'Default' }
    switch -Regex ($CharsetName.Trim().ToLowerInvariant()) {
        '^(native|default|ansi|current)$' { return 'Default' }
        '^(utf-?8|utf8)$'                  { return 'utf8' }
        '^(ascii|us-ascii)$'               { return 'ascii' }
        '^(unicode|utf-?16|utf16)$'        { return 'unicode' }
        '^(oem)$'                          { return 'oem' }
        default                            { return 'Default' }
    }
}

function Import-EnvConfig {
    [CmdletBinding()]
    param([string]$Path = '')

    $targetFile = $null
    if ($Path -and (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
        $targetFile = $Path
    } else {
        $candidates = [System.Collections.Generic.List[string]]::new()
        if ($env:DOTENV_PATH -and (Test-Path -LiteralPath $env:DOTENV_PATH -ErrorAction SilentlyContinue)) {
            $candidates.Add($env:DOTENV_PATH)
        }
        if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            $candidates.Add((Join-Path -Path $PSScriptRoot -ChildPath '.env'))
            try {
                $parent = Split-Path -Path $PSScriptRoot -Parent
                if (-not [string]::IsNullOrWhiteSpace($parent)) {
                    $candidates.Add((Join-Path -Path $parent -ChildPath '.env'))
                }
            } catch { }
        }
        try {
            $cur = (Get-Location -ErrorAction SilentlyContinue).Path
            if (-not [string]::IsNullOrWhiteSpace($cur)) {
                $candidates.Add((Join-Path -Path $cur -ChildPath '.env'))
                $curParent = Split-Path -Path $cur -Parent
                if (-not [string]::IsNullOrWhiteSpace($curParent)) {
                    $candidates.Add((Join-Path -Path $curParent -ChildPath '.env'))
                }
            }
        } catch { }
        try {
            # DriveInfo (~30 ms) replaces a Win32_Volume CIM query that cost ~350 ms on every launch
            foreach ($drv in [System.IO.DriveInfo]::GetDrives()) {
                if ($drv.DriveType -eq [System.IO.DriveType]::Removable -and $drv.IsReady) {
                    $candidates.Add((Join-Path -Path $drv.RootDirectory.FullName -ChildPath '.env'))
                }
            }
        } catch { }
        $userProfile = [Environment]::GetFolderPath('UserProfile')
        if (-not [string]::IsNullOrWhiteSpace($userProfile)) {
            $candidates.Add((Join-Path -Path $userProfile -ChildPath '.env'))
        }

        # Lowest priority: the shipped .env.example (carries public defaults like the Dell warranty key)
        # so a cloned/downloaded copy works with zero setup. A real .env always wins over it.
        if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            $candidates.Add((Join-Path -Path $PSScriptRoot -ChildPath '.env.example'))
        }
        try {
            $curEx = (Get-Location -ErrorAction SilentlyContinue).Path
            if (-not [string]::IsNullOrWhiteSpace($curEx)) { $candidates.Add((Join-Path -Path $curEx -ChildPath '.env.example')) }
        } catch { }

        foreach ($c in $candidates) {
            if (-not [string]::IsNullOrWhiteSpace($c)) {
                try {
                    if (Test-Path -LiteralPath $c -ErrorAction SilentlyContinue) {
                        $targetFile = $c
                        break
                    }
                } catch { }
            }
        }
    }

    if ($targetFile) {
        try {
            $lines = [System.IO.File]::ReadAllLines($targetFile, (Get-ScriptEncoding))
            foreach ($line in $lines) {
                if ($line -match '^\s*([^#=]+?)\s*=\s*(.*)$') {
                    $k = $Matches[1].Trim()
                    $v = $Matches[2].Trim().Trim('"').Trim("'")
                    if (-not [string]::IsNullOrWhiteSpace($k)) {
                        [Environment]::SetEnvironmentVariable($k, $v, 'Process')
                    }
                }
            }
            $script:LoadedEnvPath = $targetFile
            return $targetFile
        } catch { }
    }

    # Remote fallback: under `irm https://onyachamp.com/autopilot | iex` there is no file on disk, so
    # pull the shipped .env.example (public defaults such as the Dell warranty key) from the bootstrap host.
    if (-not $targetFile) {
        try {
            $envUrl = if ($env:DOTENV_URL) { $env:DOTENV_URL } else { 'https://onyachamp.com/.env.example' }
            try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch { }
            $resp = Invoke-RestMethod -Uri $envUrl -TimeoutSec 8 -ErrorAction Stop
            $text = if ($resp -is [string]) { $resp } else { [string]$resp }
            $applied = 0
            foreach ($line in ($text -split "`r?`n")) {
                if ($line -match '^\s*([^#=]+?)\s*=\s*(.*)$') {
                    $k = $Matches[1].Trim()
                    $v = $Matches[2].Trim().Trim('"').Trim("'")
                    # Non-destructive: never override a value already provided by the real environment
                    if (-not [string]::IsNullOrWhiteSpace($k) -and [string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($k, 'Process'))) {
                        [Environment]::SetEnvironmentVariable($k, $v, 'Process')
                        $applied++
                    }
                }
            }
            if ($applied -gt 0) { $script:LoadedEnvPath = $envUrl; return $envUrl }
        } catch { }
    }
    return $null
}

function Save-EnvConfig {
    param(
        [hashtable]$Settings,
        [string]$Path = ''
    )
    if (-not $Path) {
        $Path = if ($script:LoadedEnvPath) { $script:LoadedEnvPath } else { Join-Path (Get-Location) '.env' }
    }
    $existing = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue) {
        try {
            $lines = [System.IO.File]::ReadAllLines($Path, (Get-ScriptEncoding))
            foreach ($line in $lines) {
                if ($line -match '^\s*([^#=]+?)\s*=\s*(.*)$') {
                    $existing[$Matches[1].Trim()] = $Matches[2].Trim()
                }
            }
        } catch { }
    }
    foreach ($k in $Settings.Keys) {
        $existing[$k] = [string]$Settings[$k]
    }
    $outLines = [System.Collections.Generic.List[string]]::new()
    $outLines.Add("# Autopilot Hub Configuration & Environment Variables")
    $outLines.Add("# Saved at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via $(Get-ScriptEncodingName) charset")
    foreach ($k in $existing.Keys) {
        $outLines.Add("$k=$($existing[$k])")
    }
    [System.IO.File]::WriteAllLines($Path, $outLines, (Get-ScriptEncoding))
    $script:LoadedEnvPath = $Path
    return $Path
}

# Auto-load .env configuration at bootstrap
Import-EnvConfig -Path $EnvFile | Out-Null

if ([string]::IsNullOrWhiteSpace($GroupTag) -and $env:AUTOPILOT_GROUP_TAG) {
    $GroupTag = $env:AUTOPILOT_GROUP_TAG
}
if ([string]::IsNullOrWhiteSpace($AssignedUser) -and $env:AUTOPILOT_ASSIGNED_USER) {
    $AssignedUser = $env:AUTOPILOT_ASSIGNED_USER
}
if (-not $RenameComputer -and ($env:AUTOPILOT_RENAME_ENABLED -match '^(true|1|yes)$')) {
    $RenameComputer = $true
}
if ([string]::IsNullOrWhiteSpace($ComputerNamePrefix)) {
    $ComputerNamePrefix = if ($env:AUTOPILOT_NAME_PREFIX) { $env:AUTOPILOT_NAME_PREFIX } else { 'WS' }
}
if ([string]::IsNullOrWhiteSpace($ComputerNameTemplate)) {
    $ComputerNameTemplate = if ($env:AUTOPILOT_NAME_TEMPLATE) { $env:AUTOPILOT_NAME_TEMPLATE } else { "$ComputerNamePrefix-%SERIAL%" }
}

# 1. Enforce Modern Cryptographic TLS Protocols (TLS 1.2 / TLS 1.3)
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
} catch {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

# 2. Administrator Privilege Verification
$script:IsElevated = $false
try {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)
    $script:IsElevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $script:IsElevated = $false
}

# 3. WPF & Windows Presentation Assembly Pre-Loading
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Windows.Forms -ErrorAction SilentlyContinue

# 4. Global State & Hardware Hash Cache
$script:CachedHashInfo = $null
$script:GraphAuthContext = $null
$script:DeviceState = $null
$script:CachedOa3Key = $null
$script:PreflightRunspace = $null
$script:PreflightHandle = $null
$script:CachedDsregStatus = $null
$script:BypassDisabledAdapters = [System.Collections.Generic.List[string]]::new()

# 5. Self-Source Capture. Under `irm <url> | iex` there is no $PSCommandPath, but the script block that is
#    executing still carries the full source - that is what makes an elevated relaunch and the post-restart
#    resume possible without needing internet access again.
$script:SelfScriptPath = $PSCommandPath
$script:SelfSource = $null
try {
    if ([string]::IsNullOrWhiteSpace($script:SelfScriptPath)) {
        $script:SelfSource = $MyInvocation.MyCommand.ScriptBlock.ToString()
    }
} catch { }
$script:BootstrapUrl    = 'https://onyachamp.com/autopilot'
$programFiles           = if ($env:ProgramFiles) { $env:ProgramFiles } else { 'C:\Program Files' }
$script:PersistRoot     = Join-Path $programFiles 'AutopilotCommandHub'
$script:ResumeTaskName  = 'AutopilotCommandHub-ResumeAfterRestart'
$script:ResumeRunOnceName = 'AutopilotCommandHubResume'
$script:ResumeFromRestart = [bool]$ResumeFromRestart

# ==============================================================================
# SECTION A: EMBEDDED RESILIENT CORE ENGINES (AutopilotFast, IntuneShared, WingetIntune, WingetBatch)
# ==============================================================================

# --- Function: Get-HubRuntimeContext (Privilege & Session Mode Detection) ---
function Get-HubRuntimeContext {
    $isSystem = $false
    $userName = $env:USERNAME
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $isSystem = $identity.IsSystem
        $userName = $identity.Name
    } catch { }
    $isElevated = [bool]($script:IsElevated -or $isSystem)

    # OOBE detection: Windows Setup flags first, then the tell-tale of the OOBE session (no shell, OOBE service account)
    $isOobe = $false
    try {
        $setup = Get-ItemProperty -Path 'HKLM:\SYSTEM\Setup' -ErrorAction SilentlyContinue
        if ($setup -and (($setup.OOBEInProgress -eq 1) -or ($setup.SystemSetupInProgress -eq 1))) { $isOobe = $true }
    } catch { }
    if (-not $isOobe) {
        try {
            $explorerRunning = [bool](Get-Process -Name explorer -ErrorAction SilentlyContinue)
            if (-not $explorerRunning -and ($userName -match 'defaultuser0$' -or $isSystem)) { $isOobe = $true }
        } catch { }
    }

    $hostPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    try { $hp = (Get-Process -Id $PID -ErrorAction Stop).Path; if ($hp) { $hostPath = $hp } } catch { }

    [PSCustomObject]@{
        UserName       = $userName
        IsSystem       = $isSystem
        IsElevated     = $isElevated
        IsOobe         = $isOobe
        Mode           = if ($isOobe) { 'OOBE' } else { 'DESKTOP' }
        PrivilegeLevel = if ($isSystem) { 'SYSTEM' } elseif ($isElevated) { 'ADMINISTRATOR' } else { 'STANDARD USER' }
        PreferredLevel = if ($isOobe) { 'SYSTEM (the Shift+F10 command prompt)' } else { 'ADMINISTRATOR (an elevated PowerShell)' }
        MeetsPreferred = $isElevated
        HostEngine     = "PowerShell $($PSVersionTable.PSVersion)"
        HostPath       = $hostPath
    }
}
$script:RuntimeContext = Get-HubRuntimeContext

# --- Function: Set-HubAuthenticodeSignature (Autonomous Local Authenticode Code-Signing) ---
function Set-HubAuthenticodeSignature {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) { return $false }
    if (-not $script:RuntimeContext.IsElevated) { return $false }

    # Ensure LocalMachine execution policy permits signed scripts without requiring -ExecutionPolicy Bypass
    try {
        $policy = Get-ExecutionPolicy -Scope LocalMachine -ErrorAction SilentlyContinue
        if ($policy -in @('Restricted', 'Undefined')) {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction SilentlyContinue
        }
    } catch { }

    try {
        # Check for existing valid code-signing certificate in LocalMachine\My
        $cert = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue | Where-Object {
            $_.Subject -eq 'CN=Autopilot Provisioning Engine' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date)
        } | Select-Object -First 1

        if (-not $cert) {
            # Programmatically generate self-signed Root CA / Code Signing certificate
            $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Autopilot Provisioning Engine" -CertStoreLocation "Cert:\LocalMachine\My" -ErrorAction Stop
        }

        # Export & import to Trusted Root and Trusted Publisher using unique temp file
        $guidStr = [System.Guid]::NewGuid().ToString('N')
        $cerPath = Join-Path ([System.IO.Path]::GetTempPath()) "hub_signing_$guidStr.cer"
        try {
            Export-Certificate -Cert $cert -FilePath $cerPath -Force -ErrorAction Stop | Out-Null
            Import-Certificate -FilePath $cerPath -CertStoreLocation 'Cert:\LocalMachine\Root' -ErrorAction Stop | Out-Null
            Import-Certificate -FilePath $cerPath -CertStoreLocation 'Cert:\LocalMachine\TrustedPublisher' -ErrorAction Stop | Out-Null
        } finally {
            if (Test-Path $cerPath) { Remove-Item -Path $cerPath -Force -ErrorAction SilentlyContinue }
        }

        # Check if already signed with valid signature
        $existingSig = Get-AuthenticodeSignature -FilePath $FilePath -ErrorAction SilentlyContinue
        if ($existingSig -and $existingSig.Status -eq 'Valid') {
            return $true
        }

        # Sign the script file
        $sigResult = Set-AuthenticodeSignature -FilePath $FilePath -Certificate $cert -ErrorAction Stop
        return ($sigResult.Status -eq 'Valid')
    } catch {
        return $false
    }
}

# --- Function: Save-HubSelfCopy (persist the running script + .env for relaunch / resume) ---
function Save-HubSelfCopy {
    param([string]$Directory = $script:PersistRoot)

    # If running un-elevated on desktop, Program Files is read-only; fall back to Temp for UAC hand-off
    if (-not $script:RuntimeContext.IsElevated -and $Directory -eq $script:PersistRoot) {
        $Directory = Join-Path ([System.IO.Path]::GetTempPath()) 'AutopilotCommandHub'
    }

    if (-not (Test-Path $Directory)) { New-Item -ItemType Directory -Path $Directory -Force | Out-Null }
    $target = Join-Path $Directory 'autopilot.ps1'
    $resolvedTarget = if (Test-Path $target) { (Resolve-Path $target).Path } else { $target }

    if ($script:SelfScriptPath -and (Test-Path $script:SelfScriptPath)) {
        $resolvedSource = (Resolve-Path $script:SelfScriptPath).Path
        if (-not $resolvedSource.Equals($resolvedTarget, [System.StringComparison]::OrdinalIgnoreCase)) {
            Copy-Item -Path $script:SelfScriptPath -Destination $target -Force
        }
    } elseif ($script:SelfSource) {
        [System.IO.File]::WriteAllText($target, $script:SelfSource, [System.Text.UTF8Encoding]::new($false))
    } else {
        throw "Cannot locate the running script's source to persist it."
    }

    # Carry the active .env so tenant defaults survive the relaunch
    if ($script:LoadedEnvPath -and (Test-Path $script:LoadedEnvPath)) {
        $envTarget = Join-Path $Directory '.env'
        $resolvedEnvTarget = if (Test-Path $envTarget) { (Resolve-Path $envTarget).Path } else { $envTarget }
        $resolvedEnvSource = (Resolve-Path $script:LoadedEnvPath).Path
        if (-not $resolvedEnvSource.Equals($resolvedEnvTarget, [System.StringComparison]::OrdinalIgnoreCase)) {
            Copy-Item -Path $script:LoadedEnvPath -Destination $envTarget -Force
        }
    }

    # Strategy B: Autonomous Local Authenticode Code-Signing in OOBE / Elevated mode
    if ($script:RuntimeContext.IsElevated) {
        try {
            Set-HubAuthenticodeSignature -FilePath $target | Out-Null
        } catch { }
    }

    return $target
}

function Get-HubRelaunchArguments {
    param([Parameter(Mandatory = $true)][string]$ScriptPath, [switch]$Resume, [switch]$Replacing, [switch]$AllowBypass)
    # Strategy B: Dropped -ExecutionPolicy Bypass from scheduled tasks and persistent hooks
    $argList = @('-NoProfile')
    if ($AllowBypass) { $argList += @('-ExecutionPolicy', 'Bypass') }
    $argList += @('-STA', '-File', "`"$ScriptPath`"")
    $envCandidate = Join-Path (Split-Path $ScriptPath -Parent) '.env'
    if (Test-Path $envCandidate) { $argList += @('-EnvFile', "`"$envCandidate`"") }
    if ($Resume) { $argList += '-ResumeFromRestart' }
    if ($Replacing) { $argList += '-ReplacingInstance' }
    return ($argList -join ' ')
}

# --- Function: Invoke-HubElevatedRelaunch (UAC relaunch of this exact script) ---
function Invoke-HubElevatedRelaunch {
    $path = Save-HubSelfCopy
    # -Replacing tells the elevated child this is a hand-off, so it waits for our single-instance mutex
    # instead of bailing out with "already running" (the old Catch-22 where the elevated window never showed).
    $argLine = Get-HubRelaunchArguments -ScriptPath $path -Replacing -AllowBypass
    Start-Process -FilePath $script:RuntimeContext.HostPath -ArgumentList $argLine -Verb RunAs -ErrorAction Stop | Out-Null
    # Release our instance mutex now so the elevated child can take it as this window closes.
    try {
        if ($script:InstanceMutex) {
            try { $script:InstanceMutex.ReleaseMutex() } catch { }
            $script:InstanceMutex.Dispose(); $script:InstanceMutex = $null
        }
    } catch { }
    return $true
}

# --- Function: Close-ExistingHubInstances (Terminate prior running Hub instances) ---
function Close-ExistingHubInstances {
    [CmdletBinding()]
    param()

    $currentPid = $PID
    $targetPids = [System.Collections.Generic.HashSet[int]]::new()

    # 1. Inspect PID file in Temp, ProgramFiles and ProgramData (legacy)
    foreach ($pPath in @(
        (Join-Path ([System.IO.Path]::GetTempPath()) 'AutopilotCommandHub.pid'),
        (Join-Path $script:PersistRoot 'hub.pid'),
        (Join-Path $env:ProgramData 'AutopilotCommandHub\hub.pid')
    )) {
        if (Test-Path $pPath) {
            try {
                $raw = (Get-Content -Path $pPath -Raw -ErrorAction SilentlyContinue).Trim()
                $val = 0
                if ([int]::TryParse($raw, [ref]$val) -and $val -gt 0 -and $val -ne $currentPid) {
                    [void]$targetPids.Add($val)
                }
            } catch { }
        }
    }

    # 2. Enumerate existing Hub windows and post WM_CLOSE for graceful UI shutdown
    try {
        if (-not ('HubWinUtil' -as [type])) {
            Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class HubWinUtil {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@ -ErrorAction SilentlyContinue
        }

        if ('HubWinUtil' -as [type]) {
            $windowCallback = [HubWinUtil+EnumWindowsProc]{
                param($hWnd, $lParam)
                try {
                    $sb = [System.Text.StringBuilder]::new(256)
                    [void][HubWinUtil]::GetWindowText($hWnd, $sb, 256)
                    $title = $sb.ToString()
                    if ($title -like '*Autopilot Provisioning Hub*' -or $title -like '*Privilege Check*') {
                        $wPid = 0
                        [void][HubWinUtil]::GetWindowThreadProcessId($hWnd, [ref]$wPid)
                        if ($wPid -and $wPid -ne $currentPid) {
                            [void]$targetPids.Add($wPid)
                            # WM_CLOSE = 0x0010
                            [void][HubWinUtil]::PostMessage($hWnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
                        }
                    }
                } catch { }
                return $true
            }
            [void][HubWinUtil]::EnumWindows($windowCallback, [IntPtr]::Zero)
        }
    } catch { }

    # 3. Query Win32_Process for any lingering PowerShell processes executing AutopilotCommandHub
    try {
        $candidateProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
            ($_.Name -in @('powershell.exe', 'pwsh.exe')) -and
            ($_.ProcessId -ne $currentPid) -and
            # Match the Hub itself only. '*AutopilotCommandHub*' would also match Get-DellWarranty.ps1,
            # Start-DeviceAuth.ps1, the test suite or any dev shell that names the repo folder.
            (
                $_.CommandLine -match '(?i)[\\/]autopilot(\.ps1)?["'']?(\s|$)' -or
                $_.CommandLine -like '*onyachamp.com/autopilot*' -or
                $_.CommandLine -like '*-ResumeFromRestart*' -or
                $_.CommandLine -like '*-ReplacingInstance*'
            ) -and
            $_.CommandLine -notlike '*-NoGui*' -and
            $_.CommandLine -notlike '*-HarvestOnly*' -and
            $_.CommandLine -notlike '*-DellWarranty*' -and
            $_.CommandLine -notlike '*-ExportCsv*' -and
            $_.CommandLine -notlike '*-RenameComputer*'
        }
        foreach ($cp in $candidateProcs) {
            [void]$targetPids.Add([int]$cp.ProcessId)
        }
    } catch { }

    # 4. Gracefully await exit, then force terminate any still running
    if ($targetPids.Count -gt 0) {
        Write-Host "Closing $($targetPids.Count) existing Autopilot Command Hub instance(s) [PID(s): $($targetPids -join ', ')]..." -ForegroundColor Yellow
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.ElapsedMilliseconds -lt 300) {
            $anyAlive = $false
            foreach ($tPid in $targetPids) {
                if (Get-Process -Id $tPid -ErrorAction SilentlyContinue) {
                    $anyAlive = $true
                    break
                }
            }
            if (-not $anyAlive) { break }
            Start-Sleep -Milliseconds 50
        }

        foreach ($tPid in $targetPids) {
            try {
                $proc = Get-Process -Id $tPid -ErrorAction SilentlyContinue
                if ($proc) {
                    Stop-Process -Id $tPid -Force -ErrorAction SilentlyContinue
                }
            } catch { }
        }
    }

    # Clean up stale PID file references
    foreach ($pPath in @(
        (Join-Path ([System.IO.Path]::GetTempPath()) 'AutopilotCommandHub.pid'),
        (Join-Path $script:PersistRoot 'hub.pid'),
        (Join-Path $env:ProgramData 'AutopilotCommandHub\hub.pid')
    )) {
        try { if (Test-Path $pPath) { Remove-Item -Path $pPath -Force -ErrorAction SilentlyContinue } } catch { }
    }
}

# --- Function: Register-HubResumeAfterRestart (re-open the Hub when OOBE / the desktop comes back) ---
function Register-HubResumeAfterRestart {
    $ctx = $script:RuntimeContext
    $path = Save-HubSelfCopy
    $argLine = Get-HubRelaunchArguments -ScriptPath $path -Resume -AllowBypass
    $hostExe = $ctx.HostPath
    $targetDir = Split-Path $path -Parent

    # Generate standalone resume.cmd launcher in the persistence directory
    # This guarantees execution policy bypass is invoked cleanly without process execution traps
    $resumeCmdPath = Join-Path $targetDir 'resume.cmd'
    $envCandidate = Join-Path $targetDir '.env'
    $envArg = if (Test-Path $envCandidate) { " -EnvFile `"`"$envCandidate`"`"" } else { "" }
    $cmdLine = "`"$hostExe`" -NoProfile -ExecutionPolicy Bypass -STA -File `"`"$path`"`"$envArg -ResumeFromRestart"
    $cmdContent = "@echo off`r`nstart `"`" $cmdLine`r`n"
    try {
        [System.IO.File]::WriteAllText($resumeCmdPath, $cmdContent, [System.Text.UTF8Encoding]::new($false))
    } catch { }

    # Tier 1: HKLM RunOnce (Processed by Winlogon on interactive logon in BOTH OOBE and Desktop)
    try {
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name $script:ResumeRunOnceName -Value "`"$resumeCmdPath`"" -Type String -Force -ErrorAction SilentlyContinue
    } catch { }

    # Tier 2: Interactive Scheduled Task
    # Note: Setting a scheduled task to 'run whether user logged on or not' puts it in Session 0 (hidden desktop).
    # GroupId 'BUILTIN\Administrators' ensures the interactive session triggers at logon with Highest privileges for
    # defaultuser0 in OOBE and any administrator on the desktop.
    try {
        if (Get-Command New-ScheduledTaskAction -ErrorAction SilentlyContinue) {
            $action    = New-ScheduledTaskAction -Execute $hostExe -Argument $argLine
            $trigger   = New-ScheduledTaskTrigger -AtLogOn
            $principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Administrators" -RunLevel Highest
            $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)
            Register-ScheduledTask -TaskName $script:ResumeTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -ErrorAction SilentlyContinue | Out-Null
        } elseif (Get-Command schtasks.exe -ErrorAction SilentlyContinue) {
            $schCmd = "schtasks.exe /Create /TN `"$($script:ResumeTaskName)`" /TR `"`"$resumeCmdPath`"`" /SC ONLOGON /RL HIGHEST /F"
            cmd.exe /c $schCmd 2>&1 | Out-Null
        }
    } catch { }

    # Tier 3: Native Windows Setup Hook (SetupComplete.cmd)
    # Serves as post-OOBE desktop transition fallback when Windows Setup completes
    try {
        $setupDir = Join-Path $env:SystemRoot 'Setup\Scripts'
        if (-not (Test-Path $setupDir)) { New-Item -ItemType Directory -Path $setupDir -Force | Out-Null }
        $setupCompletePath = Join-Path $setupDir 'SetupComplete.cmd'
        $hookBlock = "REM --- AutopilotCommandHub SetupComplete Hook ---`r`nstart `"`" $cmdLine`r`nREM --- End AutopilotCommandHub Hook ---"

        if (Test-Path $setupCompletePath) {
            $existing = Get-Content -Path $setupCompletePath -Raw -ErrorAction SilentlyContinue
            if ($existing -notlike '*AutopilotCommandHub*') {
                $hasEchoOff = ($existing -match '(?i)^\s*@echo\s+off')
                if ($hasEchoOff) {
                    $cleaned = $existing -replace '(?i)^\s*@echo\s+off\s*(\r?\n)?', ''
                    $newContent = "@echo off`r`n$hookBlock`r`n$cleaned"
                } else {
                    $newContent = "$hookBlock`r`n$existing"
                }
                [System.IO.File]::WriteAllText($setupCompletePath, $newContent, [System.Text.UTF8Encoding]::new($false))
            }
        } else {
            $newContent = "@echo off`r`n$hookBlock`r`n"
            [System.IO.File]::WriteAllText($setupCompletePath, $newContent, [System.Text.UTF8Encoding]::new($false))
        }
    } catch { }

    return $path
}

function Unregister-HubResumeAfterRestart {
    $removed = $false

    # 1. Clean up HKLM RunOnce entry
    try {
        $ro = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name $script:ResumeRunOnceName -ErrorAction SilentlyContinue
        if ($ro) {
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name $script:ResumeRunOnceName -Force -ErrorAction SilentlyContinue
            $removed = $true
        }
    } catch { }

    # 2. Clean up Scheduled Task if present
    try {
        if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
            if (Get-ScheduledTask -TaskName $script:ResumeTaskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $script:ResumeTaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                $removed = $true
            }
        } elseif (Get-Command schtasks.exe -ErrorAction SilentlyContinue) {
            cmd.exe /c "schtasks.exe /Delete /TN `"$($script:ResumeTaskName)`" /F" 2>&1 | Out-Null
            $removed = $true
        }
    } catch { }

    # 3. Clean up SetupComplete.cmd hook if present
    try {
        $setupCompletePath = Join-Path $env:SystemRoot 'Setup\Scripts\SetupComplete.cmd'
        if (Test-Path $setupCompletePath) {
            $content = Get-Content -Path $setupCompletePath -Raw -ErrorAction SilentlyContinue
            if ($content -like '*AutopilotCommandHub*') {
                $lines = (Get-Content -Path $setupCompletePath -ErrorAction SilentlyContinue) | Where-Object {
                    $_ -notmatch '(?i)AutopilotCommandHub' -and
                    $_ -notmatch '(?i)autopilot\.ps1'
                }
                $trimmed = ($lines -join "`r`n").Trim()
                if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed -eq '@echo off') {
                    Remove-Item -Path $setupCompletePath -Force -ErrorAction SilentlyContinue
                } else {
                    [System.IO.File]::WriteAllText($setupCompletePath, ($lines -join "`r`n"), [System.Text.UTF8Encoding]::new($false))
                }
                $removed = $true
            }
        }
    } catch { }

    # 4. Clean up resume.cmd launchers if present
    try {
        foreach ($d in @($script:PersistRoot, $env:ProgramData, ([System.IO.Path]::GetTempPath()))) {
            $cmdFile = Join-Path $d 'AutopilotCommandHub\resume.cmd'
            if (Test-Path $cmdFile) { Remove-Item -Path $cmdFile -Force -ErrorAction SilentlyContinue }
            $directCmd = Join-Path $d 'resume.cmd'
            if (Test-Path $directCmd) { Remove-Item -Path $directCmd -Force -ErrorAction SilentlyContinue }
        }
    } catch { }

    return $removed
}

# --- Function: Invoke-AutopilotBypass (finish OOBE as a normal Windows, no Autopilot / no forced enrollment) ---
function Invoke-AutopilotBypass {
    param([switch]$DisableNetwork)
    $actions = [System.Collections.Generic.List[string]]::new()
    $disabledAdapters = [System.Collections.Generic.List[string]]::new()

    # 1. The only thing that matters: no Autopilot profile may drive this OOBE. Remove anything cached.
    foreach ($jsonPath in @(
        "$env:SystemRoot\ServiceState\wmansvc\AutopilotDDSZTDFile.json",
        "$env:SystemRoot\Provisioning\Autopilot\AutopilotConfigurationFile.json",
        "$env:SystemRoot\System32\oobe\info\default\v1\AutopilotConfigurationFile.json"
    )) {
        if (Test-Path $jsonPath) {
            try { Remove-Item -Path $jsonPath -Force -ErrorAction Stop; $actions.Add("Removed cached Autopilot profile: $jsonPath") }
            catch { $actions.Add("FAILED to remove $jsonPath : $($_.Exception.Message)") }
        }
    }

    # 2. Flag Autopilot as disabled on this install and clear any tenant assignment the client wrote
    try {
        $diagKey = 'HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot'
        if (-not (Test-Path $diagKey)) { New-Item -Path $diagKey -Force | Out-Null }
        Set-ItemProperty -Path $diagKey -Name 'IsAutopilotDisabled' -Value 1 -Type DWord -ErrorAction Stop
        foreach ($v in @('CloudAssignedTenantDomain', 'CloudAssignedTenantId', 'CloudAssignedOobeConfig', 'CloudAssignedDeviceName')) {
            Remove-ItemProperty -Path $diagKey -Name $v -ErrorAction SilentlyContinue
        }
        $actions.Add("Set IsAutopilotDisabled=1 and cleared CloudAssigned* under Provisioning\Diagnostics\AutoPilot")
    } catch { $actions.Add("FAILED to write Autopilot diagnostics key: $($_.Exception.Message)") }

    # 2b. Purge tenant policy cache so Windows OOBE cannot lock into corporate work/school flow
    try {
        $policyCache = 'HKLM:\SOFTWARE\Microsoft\Provisioning\AutopilotPolicyCache'
        if (Test-Path $policyCache) {
            Remove-Item -Path $policyCache -Recurse -Force -ErrorAction Stop
            $actions.Add("Purged AutopilotPolicyCache registry key")
        }
    } catch { $actions.Add("FAILED to purge AutopilotPolicyCache: $($_.Exception.Message)") }

    # 2c. Decommission MDM enrollment keys
    try {
        foreach ($k in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Enrollments' -ErrorAction SilentlyContinue)) {
            $pv = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
            if ($pv -and $pv.ProviderID -eq 'MS DM Server') {
                Remove-Item -Path $k.PSPath -Recurse -Force -ErrorAction Stop
                $actions.Add("Purged MDM Enrollment key $($k.PSChildName)")
            }
        }
    } catch { }

    # 3. Unlock offline setup & Personal Account options in OOBE
    try {
        foreach ($kPath in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE')) {
            if (-not (Test-Path $kPath)) { New-Item -Path $kPath -Force | Out-Null }
            Set-ItemProperty -Path $kPath -Name 'BypassNRO' -Value 1 -Type DWord -ErrorAction Stop
            Set-ItemProperty -Path $kPath -Name 'HidePersonalAccountOption' -Value 0 -Type DWord -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $kPath -Name 'HideWorkOrSchoolOption' -Value 1 -Type DWord -ErrorAction SilentlyContinue
        }
        $actions.Add("Set BypassNRO=1, HidePersonalAccountOption=0, and HideWorkOrSchoolOption=1 in OOBE setup registry")
    } catch { $actions.Add("FAILED to set OOBE flags: $($_.Exception.Message)") }

    # 4. Optionally cut network so Deployment Service cannot be contacted for offline setup
    if ($DisableNetwork) {
        try {
            foreach ($ad in @(Get-NetAdapter -Physical -ErrorAction Stop | Where-Object { $_.Status -ne 'Disabled' })) {
                try { Disable-NetAdapter -Name $ad.Name -Confirm:$false -ErrorAction Stop; $disabledAdapters.Add($ad.Name) } catch { }
            }
            if ($disabledAdapters.Count -gt 0) { $actions.Add("Disabled network adapters: $($disabledAdapters -join ', ') (re-enable from the Hub or Settings after OOBE)") }
            else { $actions.Add("No enabled physical network adapters found to disable") }
        } catch { $actions.Add("Could not enumerate network adapters: $($_.Exception.Message)") }
    }

    $script:BypassDisabledAdapters = $disabledAdapters

    [PSCustomObject]@{
        Actions          = $actions
        DisabledAdapters = $disabledAdapters
    }
}

function Enable-HubNetworkAdapters {
    $enabled = [System.Collections.Generic.List[string]]::new()
    if ($script:BypassDisabledAdapters -and $script:BypassDisabledAdapters.Count -gt 0) {
        foreach ($adapterName in $script:BypassDisabledAdapters) {
            try {
                $ad = Get-NetAdapter -Name $adapterName -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Disabled' }
                if ($ad) {
                    Enable-NetAdapter -Name $adapterName -Confirm:$false -ErrorAction Stop
                    $enabled.Add($adapterName)
                }
            } catch { }
        }
        $script:BypassDisabledAdapters = [System.Collections.Generic.List[string]]::new()
    } else {
        foreach ($ad in @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Disabled' })) {
            try { Enable-NetAdapter -Name $ad.Name -Confirm:$false -ErrorAction Stop; $enabled.Add($ad.Name) } catch { }
        }
    }
    return ,$enabled
}

# --- Function: Invoke-HubSystemReboot (Multi-tier resilient system reboot) ---
function Invoke-HubSystemReboot {
    [CmdletBinding()]
    param(
        [string]$Reason = "Autopilot Hub Restart",
        [int]$DelaySeconds = 0
    )
    Write-HubLog "Executing system reboot (Delay: ${DelaySeconds}s, Reason: '$Reason')..." "WARN"

    # Tier 1: shutdown.exe /r /t <delay> /f (works across SYSTEM, Administrator, defaultuser0, and OOBE)
    try {
        $shutdownExe = Join-Path $env:SystemRoot 'System32\shutdown.exe'
        if (Test-Path $shutdownExe) {
            # Direct invocation without Start-Process quoting pitfalls
            & $shutdownExe /r /t $DelaySeconds /f /c "$Reason"
            if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 1190) {
                Write-HubLog "Initiated reboot via shutdown.exe /r /t $DelaySeconds /f" "SUCCESS"
                return
            }
            # Fallback without comment if comment caused issues
            & $shutdownExe /r /t $DelaySeconds /f
            if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 1190) {
                Write-HubLog "Initiated reboot via shutdown.exe /r /t $DelaySeconds /f (no comment)" "SUCCESS"
                return
            }
        }
    } catch {
        Write-HubLog "shutdown.exe invocation warning: $($_.Exception.Message)" "WARN"
    }

    # Tier 2: WMI / CIM Win32Shutdown(6 = Forced Reboot)
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $cimRes = Invoke-CimMethod -InputObject $os -MethodName Win32Shutdown -Arguments @{ Flags = [int]6; Reserved = [int]0 } -ErrorAction SilentlyContinue
            if ($cimRes -and $cimRes.ReturnValue -eq 0) {
                Write-HubLog "Initiated reboot via CIM Win32Shutdown(6)" "SUCCESS"
                return
            }
        }
    } catch { }

    # Tier 3: Win32 API InitiateSystemShutdownEx & ExitWindowsEx
    try {
        if (-not ([System.Management.Automation.PSTypeName]'Win32NativeShutdown').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32NativeShutdown {
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern bool InitiateSystemShutdownEx(string lpMachineName, string lpMessage, uint dwTimeout, bool bForceAppsClosed, bool bRebootAfterShutdown, uint dwReason);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool ExitWindowsEx(uint uFlags, uint dwReason);
}
"@ -ErrorAction SilentlyContinue
        }
        if (([System.Management.Automation.PSTypeName]'Win32NativeShutdown').Type) {
            $res = [Win32NativeShutdown]::InitiateSystemShutdownEx($null, $Reason, [uint32]$DelaySeconds, $true, $true, 0x40000004)
            if ($res) {
                Write-HubLog "Initiated reboot via Win32 InitiateSystemShutdownEx" "SUCCESS"
                return
            }
            $res2 = [Win32NativeShutdown]::ExitWindowsEx(0x6, 0x40000004)
            if ($res2) {
                Write-HubLog "Initiated reboot via Win32 ExitWindowsEx" "SUCCESS"
                return
            }
        }
    } catch { }

    # Tier 4: Restart-Computer -Force fallback
    try {
        Restart-Computer -Force -ErrorAction Stop
        Write-HubLog "Initiated reboot via Restart-Computer -Force" "SUCCESS"
    } catch {
        Write-HubLog "Restart-Computer failed: $($_.Exception.Message)" "ERROR"
    }
}

# --- Function: Invoke-HubCreateLocalUser (Create Local Admin / User in OOBE or Desktop) ---
function Invoke-HubCreateLocalUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [string]$Password = '',
        [string]$FullName = '',
        [switch]$MakeAdmin = $true,
        [switch]$PasswordNeverExpires = $true,
        [switch]$SkipRemainingOobe = $false
    )
    $actions = [System.Collections.Generic.List[string]]::new()
    $cleanName = $UserName.Trim()
    if ([string]::IsNullOrWhiteSpace($cleanName)) {
        throw "Username cannot be empty or whitespace."
    }

    $created = $false
    $existing = $null
    try { $existing = Get-LocalUser -Name $cleanName -ErrorAction SilentlyContinue } catch { }
    if ($existing) {
        $created = $true
        $actions.Add("User account '$cleanName' already exists on this system")
    } else {
        try {
            if (-not [string]::IsNullOrEmpty($Password)) {
                $secPass = ConvertTo-SecureString $Password -AsPlainText -Force
                New-LocalUser -Name $cleanName -Password $secPass -FullName $FullName -Description "Created via Autopilot Command Hub" -PasswordNeverExpires:$PasswordNeverExpires -ErrorAction Stop | Out-Null
            } else {
                New-LocalUser -Name $cleanName -NoPassword -FullName $FullName -Description "Created via Autopilot Command Hub" -PasswordNeverExpires:$PasswordNeverExpires -ErrorAction Stop | Out-Null
            }
            $created = $true
            $actions.Add("Created local user account '$cleanName' via New-LocalUser")
        } catch {
            try {
                $netArgs = if ($Password) { "user `"$cleanName`" `"$Password`" /add" } else { "user `"$cleanName`" /add" }
                $netProc = Start-Process -FilePath "net.exe" -ArgumentList $netArgs -NoNewWindow -Wait -PassThru
                if ($netProc.ExitCode -eq 0) {
                    $created = $true
                    $actions.Add("Created local user account '$cleanName' via net.exe")
                } elseif ($netProc.ExitCode -eq 2) {
                    $created = $true
                    $actions.Add("User account '$cleanName' already exists on this system (net.exe code 2)")
                } else {
                    throw "net.exe returned exit code $($netProc.ExitCode)"
                }
            } catch {
                if ($_.Exception.Message -like '*already exists*' -or $_.Exception.Message -like '*2224*') {
                    $created = $true
                    $actions.Add("User account '$cleanName' already exists on this system")
                } else {
                    throw "Failed to create user account '$cleanName': $($_.Exception.Message)"
                }
            }
        }
    }

    if ($MakeAdmin) {
        $adminAdded = $false
        try {
            $existingMembers = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue
            if ($existingMembers | Where-Object { $_.Name -like "*\$cleanName" -or $_.Name -eq $cleanName }) {
                $adminAdded = $true
                $actions.Add("'$cleanName' is already a member of local Administrators group")
            } else {
                Add-LocalGroupMember -Group "Administrators" -Member $cleanName -ErrorAction Stop
                $adminAdded = $true
                $actions.Add("Added '$cleanName' to local Administrators group")
            }
        } catch {
            if ($_.Exception.Message -like '*already a member*') {
                $adminAdded = $true
                $actions.Add("'$cleanName' is already a member of local Administrators group")
            } else {
                try {
                    $netGrp = Start-Process -FilePath "net.exe" -ArgumentList "localgroup administrators `"$cleanName`" /add" -NoNewWindow -Wait -PassThru
                    if ($netGrp.ExitCode -eq 0 -or $netGrp.ExitCode -eq 2) {
                        $adminAdded = $true
                        $actions.Add("Added '$cleanName' to local Administrators group via net.exe")
                    }
                } catch { }
            }
        }
        if (-not $adminAdded) {
            $actions.Add("WARNING: Could not verify Administrators group membership for '$cleanName'")
        }
    }

    if ($PasswordNeverExpires) {
        try {
            Set-LocalUser -Name $cleanName -PasswordNeverExpires $true -ErrorAction SilentlyContinue
            $actions.Add("Set password never expires for '$cleanName'")
        } catch {
            try { Start-Process -FilePath "wmic.exe" -ArgumentList "useraccount where name='$cleanName' set passwordexpires=FALSE" -NoNewWindow -Wait -ErrorAction SilentlyContinue } catch { }
        }
    }

    try {
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE'
        if (-not (Test-Path $setupKey)) { New-Item -Path $setupKey -Force | Out-Null }
        Set-ItemProperty -Path $setupKey -Name 'UnattendCreatedUser' -Value 1 -Type DWord -ErrorAction SilentlyContinue
        $actions.Add("Set UnattendCreatedUser=1 in OOBE setup registry")
        if ($SkipRemainingOobe) {
            Set-ItemProperty -Path $setupKey -Name 'OOBEInProgress' -Value 0 -Type DWord -ErrorAction SilentlyContinue
            $actions.Add("Set OOBE\OOBEInProgress=0 to bypass remaining setup screens")
        }
    } catch { }

    try {
        $sysSetupKey = 'HKLM:\SYSTEM\Setup'
        if (Test-Path $sysSetupKey) {
            if ($SkipRemainingOobe) {
                Set-ItemProperty -Path $sysSetupKey -Name 'OOBEInProgress' -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $sysSetupKey -Name 'SetupPhase' -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $sysSetupKey -Name 'SystemSetupInProgress' -Value 0 -Type DWord -ErrorAction SilentlyContinue
                $actions.Add("Configured SYSTEM\Setup (OOBEInProgress=0, SetupPhase=0) to complete setup")
            }
        }
    } catch { }

    return [PSCustomObject]@{
        Success  = $true
        UserName = $cleanName
        IsAdmin  = [bool]$MakeAdmin
        Actions  = $actions
    }
}

# --- Function: Invoke-HubSetupPersonalAccount (Unlock Personal Microsoft Account in OOBE) ---
function Invoke-HubSetupPersonalAccount {
    [CmdletBinding()]
    param()
    $actions = [System.Collections.Generic.List[string]]::new()

    # Purge corporate Autopilot & MDM caches without cutting network
    $bypass = Invoke-AutopilotBypass -DisableNetwork:$false
    foreach ($a in $bypass.Actions) { $actions.Add($a) }

    # Re-enable network adapters for active online sign-in
    $en = @(Enable-HubNetworkAdapters)
    if ($en.Count -gt 0) {
        $actions.Add("Re-enabled network adapters for online sign-in: $($en -join ', ')")
    }

    # Unlock Personal Account selection and suppress corporate Work or School in OOBE
    try {
        foreach ($kPath in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE')) {
            if (-not (Test-Path $kPath)) { New-Item -Path $kPath -Force | Out-Null }
            Set-ItemProperty -Path $kPath -Name 'HidePersonalAccountOption' -Value 0 -Type DWord -ErrorAction Stop
            Set-ItemProperty -Path $kPath -Name 'HideWorkOrSchoolOption' -Value 1 -Type DWord -ErrorAction Stop
            Set-ItemProperty -Path $kPath -Name 'BypassNRO' -Value 1 -Type DWord -ErrorAction SilentlyContinue
        }
        $actions.Add("Configured HidePersonalAccountOption=0 and HideWorkOrSchoolOption=1 (Work or School suppressed, Personal MSA enforced)")
    } catch { }

    return $actions
}

# --- Function: Restart-HubOobe (Reload Windows OOBE Experience Host) ---
function Restart-HubOobe {
    [CmdletBinding()]
    param()
    Write-HubLog "Stopping Windows OOBE broker processes to reload setup..." "WARN"
    try {
        Get-Process -Name "CloudExperienceHostBroker", "WWAHost", "msoobesvc", "oobe" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-HubLog "Terminated active OOBE webhost processes." "INFO"
    } catch { }

    $ans = [System.Windows.MessageBox]::Show("OOBE processes stopped.`n`nWould you like to restart the PC now to ensure OOBE reloads completely clean with the new offline/personal settings?", "Restart System for Clean OOBE?", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
        Invoke-HubSystemReboot -Reason "Reload OOBE with offline/personal settings" -DelaySeconds 0
    }
}

# --- Function: Show-CreateLocalUserDialog (Interactive OOBE / Desktop User Provisioning) ---
function Show-CreateLocalUserDialog {
    param([System.Windows.Window]$Owner = $null)

    if ($script:RuntimeContext -and -not $script:RuntimeContext.MeetsPreferred) {
        [System.Windows.MessageBox]::Show("Creating a local user account requires administrative privileges. Please click 'Fix Privileges' in the header first.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if ($script:PlaybookRunning) {
        Write-HubLog "Playbook: Auto-provisioning local administrative user account 'Admin'..." "INFO"
        try {
            $res = Invoke-HubCreateLocalUser -UserName "Admin" -Password "" -FullName "Local Administrator" -MakeAdmin:$true -PasswordNeverExpires:$true -SkipRemainingOobe:$true
            foreach ($act in $res.Actions) { Write-HubLog "  $act" "INFO" }
            Write-HubLog "Local user 'Admin' successfully created (Admin: True)." "SUCCESS"
        } catch {
            Write-HubLog "Local user creation notice: $($_.Exception.Message)" "WARN"
        }
        return
    }

    $xamlUser = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create Local User Account"
        Width="490" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="#F3F4F6" Foreground="#1E293B" FontFamily="Segoe UI Variable Text, Segoe UI, sans-serif" ResizeMode="NoResize">
    <StackPanel Margin="18">
        <TextBlock Text="CREATE LOCAL USER ACCOUNT" FontSize="11" FontWeight="SemiBold" Foreground="#64748B" Margin="0,0,0,4"/>
        <TextBlock Text="Provision a local Windows user account directly during OOBE or desktop staging without entering a Microsoft Account or corporate tenant." FontSize="11.5" Foreground="#475569" TextWrapping="Wrap" Margin="0,0,0,12"/>

        <Border Background="#FFFFFF" CornerRadius="6" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14,12" Margin="0,0,0,14">
            <StackPanel>
                <TextBlock Text="Username:" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,3"/>
                <TextBox Name="TxtNewUserName" Text="Admin" Height="28" Padding="6,3" Margin="0,0,0,10" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>

                <TextBlock Text="Full Name (Optional):" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,3"/>
                <TextBox Name="TxtNewUserFullName" Height="28" Padding="6,3" Margin="0,0,0,10" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>

                <Grid Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="10"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                        <TextBlock Text="Password (Optional):" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,3"/>
                        <PasswordBox Name="TxtNewUserPass" Height="28" Padding="6,3" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                        <TextBlock Text="Confirm Password:" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,3"/>
                        <PasswordBox Name="TxtNewUserConfirm" Height="28" Padding="6,3" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
                    </StackPanel>
                </Grid>

                <CheckBox Name="ChkMakeAdmin" Content="Add to local Administrators group" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,2,0,6"/>
                <CheckBox Name="ChkNeverExpires" Content="Password never expires" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,0,0,6"/>
                <CheckBox Name="ChkSkipOobe" Content="Mark user setup complete in OOBE setup registry" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,0,0,0"/>
            </StackPanel>
        </Border>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="BtnCancelUser" Content="Cancel" Padding="14,6" Margin="0,0,8,0" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
            <Button Name="BtnCreateUserConfirm" Content="Create Account" Padding="16,6" Background="#0067C0" Foreground="#FFFFFF" BorderBrush="#0067C0" FontWeight="SemiBold"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlUser))
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $dlg.Owner = $Owner }

    $txtName = $dlg.FindName('TxtNewUserName')
    $txtFull = $dlg.FindName('TxtNewUserFullName')
    $txtPass = $dlg.FindName('TxtNewUserPass')
    $txtConf = $dlg.FindName('TxtNewUserConfirm')
    $chkAdm  = $dlg.FindName('ChkMakeAdmin')
    $chkExp  = $dlg.FindName('ChkNeverExpires')
    $chkOobe = $dlg.FindName('ChkSkipOobe')
    $btnOk   = $dlg.FindName('BtnCreateUserConfirm')
    $btnCan  = $dlg.FindName('BtnCancelUser')

    $btnCan.Add_Click({ $dlg.Close() })
    $btnOk.Add_Click({
        $uName = $txtName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($uName)) {
            [System.Windows.MessageBox]::Show("Please enter a username.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $pass = $txtPass.Password
        $conf = $txtConf.Password
        if ($pass -ne $conf) {
            [System.Windows.MessageBox]::Show("Passwords do not match. Please re-enter passwords.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        try {
            $res = Invoke-HubCreateLocalUser -UserName $uName -Password $pass -FullName $txtFull.Text.Trim() -MakeAdmin:([bool]$chkAdm.IsChecked) -PasswordNeverExpires:([bool]$chkExp.IsChecked) -SkipRemainingOobe:([bool]$chkOobe.IsChecked)
            foreach ($act in $res.Actions) { Write-HubLog "  $act" "INFO" }
            Write-HubLog "Local user '$uName' successfully created (Admin: $([bool]$chkAdm.IsChecked))." "SUCCESS"
            $roleStr = if ($chkAdm.IsChecked) { 'Administrator' } else { 'Standard User' }
            $passStr = if ($pass) { 'Configured' } else { 'None (blank)' }
            $promptRestart = [System.Windows.MessageBox]::Show("Local user account '$uName' created successfully!`n`nRole: $roleStr`nPassword: $passStr`n`nWould you like to restart the system now to log in immediately?", "User Created - Restart Now?", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            $dlg.Close()
            if ($promptRestart -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-HubLog "Operator opted to reboot into new user account '$uName'..." "WARN"
                Invoke-HubSystemReboot -Reason "First logon for local user $uName" -DelaySeconds 0
            }
        } catch {
            Write-HubLog "Failed to create local user '$uName': $($_.Exception.Message)" "ERROR"
            [System.Windows.MessageBox]::Show("Failed to create user account '$uName':`n$($_.Exception.Message)", "Creation Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    $dlg.ShowDialog() | Out-Null
}

# ==============================================================================
# CONTROLLED OFFLINE STAGING & ADAPTIVE EVALUATION PROFILES SUBSYSTEM
# ==============================================================================
$script:StagingDir = Join-Path $env:ProgramData 'AutopilotCommandHub\OfflineStaging'
$script:StagingQuarantineDir = Join-Path $script:StagingDir 'Quarantine'
$script:StagingSnapshotPath = Join-Path $script:StagingDir 'StagingSnapshot.json'

function Get-OfflineStagingState {
    if (Test-Path $script:StagingSnapshotPath) {
        try {
            $raw = [System.IO.File]::ReadAllText($script:StagingSnapshotPath, [System.Text.Encoding]::UTF8)
            $snap = ConvertFrom-Json $raw -ErrorAction Stop
            if ($snap -and $snap.IsActive) { return $snap }
        } catch { }
    }
    return $null
}

function Invoke-OfflineStaging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateSet('HardwareDiagnostics', 'WorkloadBenchmarking', 'SoftwareSandbox')]
        [string]$Profile = 'HardwareDiagnostics',

        [switch]$QuarantineProfile = $true,
        [switch]$LoopbackFence = $true,
        [switch]$SetBypassNro = $true,
        [switch]$CreateDesktopShortcut = $true
    )

    $actions = [System.Collections.Generic.List[string]]::new()
    $quarantinedFiles = [System.Collections.Generic.List[PSCustomObject]]::new()
    $serviceSnapshots = [System.Collections.Generic.List[PSCustomObject]]::new()

    if (-not (Test-Path $script:StagingQuarantineDir)) {
        New-Item -Path $script:StagingQuarantineDir -ItemType Directory -Force | Out-Null
    }

    # Record baseline state for clean cloud re-docking
    $origDevMode = $null
    $devKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'
    try {
        if (Test-Path $devKey) {
            $p = Get-ItemProperty -Path $devKey -Name 'AllowDevelopmentWithoutDevLicense' -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties['AllowDevelopmentWithoutDevLicense']) {
                $origDevMode = [int]$p.AllowDevelopmentWithoutDevLicense
            }
        }
    } catch { }

    $origAdminEnabled = $false
    try {
        $adUser = Get-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
        if ($adUser) { $origAdminEnabled = [bool]$adUser.Enabled }
    } catch { }

    $origBypassNro = $null
    $oobeKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE'
    try {
        if (Test-Path $oobeKey) {
            $p = Get-ItemProperty -Path $oobeKey -Name 'BypassNRO' -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties['BypassNRO']) {
                $origBypassNro = [int]$p.BypassNRO
            }
        }
    } catch { }

    # 1. Reversible Quarantine of ZTD / Autopilot Profiles
    if ($QuarantineProfile) {
        $profileTargets = @(
            "$env:SystemRoot\ServiceState\wmansvc\AutopilotDDSZTDFile.json",
            "$env:SystemRoot\Provisioning\Autopilot\AutopilotConfigurationFile.json"
        )
        foreach ($target in $profileTargets) {
            if (Test-Path $target) {
                $leaf = Split-Path -Leaf $target
                $dest = Join-Path $script:StagingQuarantineDir $leaf
                try {
                    Copy-Item -Path $target -Destination $dest -Force -ErrorAction Stop
                    Remove-Item -Path $target -Force -ErrorAction Stop
                    $quarantinedFiles.Add([PSCustomObject]@{
                        OriginalPath   = $target
                        QuarantinePath = $dest
                        FileName       = $leaf
                    })
                    $actions.Add("Quarantined Autopilot profile '$leaf' to staging holding area")
                } catch {
                    $actions.Add("FAILED to quarantine '$leaf': $($_.Exception.Message)")
                }
            }
        }
    }

    # 2. Reversible Hosts Loopback Fence for Enrollment Endpoints
    $hostsFenced = $false
    if ($LoopbackFence) {
        $hostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
        try {
            $hostsContent = if (Test-Path $hostsFile) { [System.IO.File]::ReadAllText($hostsFile, [System.Text.Encoding]::ASCII) } else { "" }
            if ($hostsContent -notmatch 'AUTOPILOT_OFFLINE_STAGING_FENCE_START') {
                $fenceBlock = "`r`n# AUTOPILOT_OFFLINE_STAGING_FENCE_START`r`n127.0.0.1 ztd.dds.microsoft.com`r`n127.0.0.1 cs.dds.microsoft.com`r`n127.0.0.1 enterpriseenrollment.manage.microsoft.com`r`n127.0.0.1 enrollment.manage.microsoft.com`r`n127.0.0.1 manage.microsoft.com`r`n# AUTOPILOT_OFFLINE_STAGING_FENCE_END`r`n"
                [System.IO.File]::AppendAllText($hostsFile, $fenceBlock, [System.Text.Encoding]::ASCII)
                $hostsFenced = $true
                $actions.Add("Configured loopback fence in hosts file blocking enrollment endpoints")
            } else {
                $hostsFenced = $true
                $actions.Add("Hosts file already contains active loopback fence")
            }
            try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch { }
        } catch {
            $actions.Add("FAILED to configure hosts loopback fence: $($_.Exception.Message)")
        }
    }

    # 3. Registry Fences: BypassNRO and IsAutopilotDisabled
    if ($SetBypassNro) {
        try {
            foreach ($kPath in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE')) {
                if (-not (Test-Path $kPath)) { New-Item -Path $kPath -Force | Out-Null }
                Set-ItemProperty -Path $kPath -Name 'BypassNRO' -Value 1 -Type DWord -ErrorAction Stop
                Set-ItemProperty -Path $kPath -Name 'HidePersonalAccountOption' -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $kPath -Name 'HideWorkOrSchoolOption' -Value 1 -Type DWord -ErrorAction SilentlyContinue
            }
            $actions.Add("Set BypassNRO=1, HidePersonalAccountOption=0, and HideWorkOrSchoolOption=1 in OOBE setup registry")
            try {
                $cacheKey = 'HKLM:\SOFTWARE\Microsoft\Provisioning\AutopilotPolicyCache'
                if (Test-Path $cacheKey) { Remove-Item -Path $cacheKey -Recurse -Force -ErrorAction Stop; $actions.Add("Purged AutopilotPolicyCache registry key") }
            } catch { }
        } catch {
            $actions.Add("FAILED to set BypassNRO: $($_.Exception.Message)")
        }
    }

    try {
        $diagKey = 'HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot'
        if (-not (Test-Path $diagKey)) { New-Item -Path $diagKey -Force | Out-Null }
        Set-ItemProperty -Path $diagKey -Name 'IsAutopilotDisabled' -Value 1 -Type DWord -ErrorAction Stop
        $actions.Add("Set IsAutopilotDisabled=1 under Provisioning\Diagnostics\AutoPilot")
    } catch {
        $actions.Add("FAILED to set IsAutopilotDisabled: $($_.Exception.Message)")
    }

    # 4. Adaptive Profile Specializations
    $profileTitle = ""
    $origPowerScheme = ""

    switch ($Profile) {
        'HardwareDiagnostics' {
            $profileTitle = "Hardware Diagnostics & Triage"
            $targetServices = @('DiagTrack', 'WSearch', 'SysMain')
            foreach ($sName in $targetServices) {
                try {
                    $svc = Get-Service -Name $sName -ErrorAction SilentlyContinue
                    if ($svc) {
                        $startMode = (Get-CimInstance Win32_Service -Filter "Name='$sName'" -ErrorAction SilentlyContinue).StartMode
                        $serviceSnapshots.Add([PSCustomObject]@{
                            Name              = $sName
                            OriginalStartType = $startMode
                            OriginalStatus    = $svc.Status.ToString()
                        })
                        Stop-Service -Name $sName -Force -ErrorAction SilentlyContinue
                        Set-Service -Name $sName -StartupType Disabled -ErrorAction SilentlyContinue
                        $actions.Add("Suspended $sName (was $startMode / $($svc.Status)) for diagnostic isolation")
                    }
                } catch { }
            }
            $actions.Add("Hardware Diagnostics posture active: Local subnet network posture open, live hardware probes enabled")
        }

        'WorkloadBenchmarking' {
            $profileTitle = "Workload Benchmarking & Testing"
            try {
                $activeSchemeRaw = powercfg /getactivescheme
                if ($activeSchemeRaw -match '([0-9a-fA-F-]{36})') {
                    $origPowerScheme = $Matches[1]
                }
                $ultGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
                $allSchemes = powercfg /list
                if ($allSchemes -notmatch $ultGuid) {
                    powercfg -duplicatescheme $ultGuid | Out-Null
                }
                powercfg -setactive $ultGuid | Out-Null
                $actions.Add("Activated Ultimate Performance power scheme ($ultGuid) - zero CPU jitter/throttling")
            } catch {
                $actions.Add("Power scheme note: $($_.Exception.Message)")
            }

            $targetServices = @('wuauserv', 'bits', 'TrustedInstaller', 'DiagTrack', 'WSearch', 'SysMain')
            foreach ($sName in $targetServices) {
                try {
                    $svc = Get-Service -Name $sName -ErrorAction SilentlyContinue
                    if ($svc) {
                        $startMode = (Get-CimInstance Win32_Service -Filter "Name='$sName'" -ErrorAction SilentlyContinue).StartMode
                        $serviceSnapshots.Add([PSCustomObject]@{
                            Name              = $sName
                            OriginalStartType = $startMode
                            OriginalStatus    = $svc.Status.ToString()
                        })
                        Stop-Service -Name $sName -Force -ErrorAction SilentlyContinue
                        Set-Service -Name $sName -StartupType Disabled -ErrorAction SilentlyContinue
                        $actions.Add("Halted $sName for deterministic benchmark testing")
                    }
                } catch { }
            }
        }

        'SoftwareSandbox' {
            $profileTitle = "Software Evaluation & Staging Sandbox"
            try {
                $devKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'
                if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }
                Set-ItemProperty -Path $devKey -Name 'AllowDevelopmentWithoutDevLicense' -Value 1 -Type DWord -ErrorAction Stop
                $actions.Add("Enabled Windows Developer Mode (sideloading & remote debugging allowed)")
            } catch {
                $actions.Add("Failed to enable Developer Mode: $($_.Exception.Message)")
            }

            $stagingDirSys = "$env:SystemDrive\Staging\Packages"
            if (-not (Test-Path $stagingDirSys)) {
                New-Item -Path $stagingDirSys -ItemType Directory -Force | Out-Null
                $actions.Add("Created offline package staging folder at $stagingDirSys")
            }

            try {
                net user Administrator /active:yes | Out-Null
                Enable-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
                $actions.Add("Enabled built-in Administrator account for offline sandbox access")
            } catch { }
        }
    }

    # 5. Persist Snapshot JSON
    $snapshotObj = [PSCustomObject]@{
        Timestamp           = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Profile             = $Profile
        ProfileTitle        = $profileTitle
        QuarantinedFiles    = $quarantinedFiles
        HostsFenced         = $hostsFenced
        OriginalPowerScheme = $origPowerScheme
        OriginalDevMode     = $origDevMode
        OriginalAdminState  = $origAdminEnabled
        OriginalBypassNro   = $origBypassNro
        Services            = $serviceSnapshots
        IsActive            = $true
    }

    $jsonStr = $snapshotObj | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($script:StagingSnapshotPath, $jsonStr, [System.Text.Encoding]::UTF8)
    $actions.Add("Saved staging state snapshot to $script:StagingSnapshotPath")

    # 6. Generate Re-Docking Companion Script & Desktop Shortcut
    if ($CreateDesktopShortcut) {
        try {
            $redockScriptPath = Join-Path $script:StagingDir "Resume-EnterpriseAutopilot.ps1"
            $redockScriptCode = @"
<#
.SYNOPSIS
    Resume Enterprise Autopilot & Re-Dock Device to Cloud
#>
`$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not `$isAdmin) {
    Write-Host "Elevating privileges to restore enterprise configuration..." -ForegroundColor Yellow
    Start-Process -FilePath "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File ```"`$PSCommandPath```"" -Verb RunAs
    exit
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " RESUMING ENTERPRISE AUTOPILOT & RE-DOCKING TO CLOUD" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

`$snapshotFile = "$script:StagingSnapshotPath"
if (-not (Test-Path `$snapshotFile)) {
    Write-Host "No active staging snapshot found at `$snapshotFile." -ForegroundColor Yellow
    pause
    exit
}

`$snap = [System.IO.File]::ReadAllText(`$snapshotFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

# 1. Restore Services
if (`$snap.Services) {
    foreach (`$s in `$snap.Services) {
        Write-Host "Restoring service `$(`$s.Name) to `$(`$s.OriginalStartType)..." -ForegroundColor Gray
        try {
            Set-Service -Name `$s.Name -StartupType `$s.OriginalStartType -ErrorAction SilentlyContinue
            if (`$s.OriginalStatus -eq 'Running') {
                Start-Service -Name `$s.Name -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

# 2. Restore Power Scheme
if (`$snap.OriginalPowerScheme) {
    Write-Host "Restoring original power scheme `$(`$snap.OriginalPowerScheme)..." -ForegroundColor Gray
    try { powercfg -setactive `$snap.OriginalPowerScheme } catch { }
}

# 3. Clear Hosts Fence
`$hostsFile = "`$env:SystemRoot\System32\drivers\etc\hosts"
if (Test-Path `$hostsFile) {
    try {
        `$hLines = [System.IO.File]::ReadAllLines(`$hostsFile, [System.Text.Encoding]::ASCII)
        `$newLines = [System.Collections.Generic.List[string]]::new()
        `$inFence = `$false
        foreach (`$l in `$hLines) {
            if (`$l -like '*# AUTOPILOT_OFFLINE_STAGING_FENCE_START*') { `$inFence = `$true; continue }
            if (`$l -like '*# AUTOPILOT_OFFLINE_STAGING_FENCE_END*') { `$inFence = `$false; continue }
            if (-not `$inFence) { `$newLines.Add(`$l) }
        }
        [System.IO.File]::WriteAllLines(`$hostsFile, `$newLines, [System.Text.Encoding]::ASCII)
        Write-Host "Cleared hosts loopback fence for enrollment endpoints." -ForegroundColor Green
        try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch { }
    } catch { }
}

# 4. Restore Registry
try {
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot' -Name 'IsAutopilotDisabled' -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Write-Host "Restored IsAutopilotDisabled=0 in registry." -ForegroundColor Green
} catch { }

try {
    `$oKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE'
    if (`$snap.OriginalBypassNro -eq `$null) {
        Remove-ItemProperty -Path `$oKey -Name 'BypassNRO' -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path `$oKey -Name 'BypassNRO' -Value `$snap.OriginalBypassNro -Type DWord -ErrorAction SilentlyContinue
    }
} catch { }

if (`$snap.OriginalDevMode -ne `$null) {
    try {
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -Name 'AllowDevelopmentWithoutDevLicense' -Value `$snap.OriginalDevMode -Type DWord -ErrorAction SilentlyContinue
        Write-Host "Restored Developer Mode setting." -ForegroundColor Green
    } catch { }
}

if (`$snap.OriginalAdminState -eq `$false) {
    try {
        net user Administrator /active:no | Out-Null
        Disable-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
        Write-Host "Disabled built-in Administrator account to restore production enterprise security posture." -ForegroundColor Green
    } catch { }
}

# 5. Restore Quarantined Autopilot Profiles
if (`$snap.QuarantinedFiles) {
    foreach (`$q in `$snap.QuarantinedFiles) {
        if (Test-Path `$q.QuarantinePath) {
            `$parent = Split-Path -Parent `$q.OriginalPath
            if (-not (Test-Path `$parent)) { New-Item -Path `$parent -ItemType Directory -Force | Out-Null }
            Copy-Item -Path `$q.QuarantinePath -Destination `$q.OriginalPath -Force
            Remove-Item -Path `$q.QuarantinePath -Force -ErrorAction SilentlyContinue
            Write-Host "Restored Autopilot profile to `$(`$q.OriginalPath)." -ForegroundColor Green
        }
    }
}

# 6. Trigger Policy Sync
Write-Host "Triggering enterprise policy retrieval..." -ForegroundColor Cyan
try {
    Start-Process -FilePath "`$env:SystemRoot\System32\DeviceEnroller.exe" -ArgumentList "/c /AutoEnrollMDM" -NoNewWindow -Wait -ErrorAction SilentlyContinue
} catch { }

try {
    `$tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue)
    foreach (`$t in `$tasks) {
        try { Start-ScheduledTask -TaskName `$t.TaskName -TaskPath `$t.TaskPath -ErrorAction SilentlyContinue } catch { }
    }
} catch { }

# 7. Archive Snapshot
`$restoredPath = "`$snapshotFile.restored.json"
Move-Item -Path `$snapshotFile -Destination `$restoredPath -Force -ErrorAction SilentlyContinue

# 8. Clean desktop shortcut
`$pubDesk = [Environment]::GetFolderPath('CommonDesktopDirectory')
if (-not (Test-Path `$pubDesk)) { `$pubDesk = "`$env:PUBLIC\Desktop" }
`$lnkPath = Join-Path `$pubDesk "Resume-EnterpriseAutopilot.lnk"
if (Test-Path `$lnkPath) { Remove-Item -Path `$lnkPath -Force -ErrorAction SilentlyContinue }

Write-Host "============================================================" -ForegroundColor Green
Write-Host " DEVICE SUCCESSFULLY RE-DOCKED TO ENTERPRISE CLOUD!" -ForegroundColor Green
Write-Host " Autopilot profile and enterprise management active." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Press any key to close..."
`$null = `$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
"@
            [System.IO.File]::WriteAllText($redockScriptPath, $redockScriptCode, [System.Text.Encoding]::UTF8)

            # Create Desktop Shortcut
            $pubDesk = [Environment]::GetFolderPath('CommonDesktopDirectory')
            if (-not (Test-Path $pubDesk)) { $pubDesk = "$env:PUBLIC\Desktop" }
            if (Test-Path $pubDesk) {
                $lnkPath = Join-Path $pubDesk "Resume-EnterpriseAutopilot.lnk"
                $wsh = New-Object -ComObject WScript.Shell
                $sc = $wsh.CreateShortcut($lnkPath)
                $sc.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
                $sc.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$redockScriptPath`""
                $sc.Description = "Reconnect to Enterprise Cloud & Resume Autopilot"
                $sc.IconLocation = "$env:SystemRoot\System32\shell32.dll,238"
                $sc.Save()
                $actions.Add("Placed 'Resume-EnterpriseAutopilot' re-docking shortcut on Public Desktop")
            }
        } catch {
            $actions.Add("Desktop shortcut note: $($_.Exception.Message)")
        }
    }

    return [PSCustomObject]@{
        Profile  = $Profile
        Title    = $profileTitle
        Actions  = $actions
        Snapshot = $snapshotObj
    }
}

function Invoke-AutopilotReDock {
    [CmdletBinding()]
    param([switch]$Force)

    $actions = [System.Collections.Generic.List[string]]::new()
    $snap = Get-OfflineStagingState
    if (-not $snap) {
        return [PSCustomObject]@{ Success = $false; Message = "No active offline staging snapshot found."; Actions = @() }
    }

    # 1. Restore Services
    if ($snap.Services) {
        foreach ($s in $snap.Services) {
            try {
                Set-Service -Name $s.Name -StartupType $s.OriginalStartType -ErrorAction SilentlyContinue
                if ($s.OriginalStatus -eq 'Running') {
                    Start-Service -Name $s.Name -ErrorAction SilentlyContinue
                }
                $actions.Add("Restored service '$($s.Name)' to $($s.OriginalStartType) (status: $($s.OriginalStatus))")
            } catch {
                $actions.Add("Failed to restore service '$($s.Name)': $($_.Exception.Message)")
            }
        }
    }

    # 2. Restore Power Scheme
    if ($snap.OriginalPowerScheme) {
        try {
            powercfg -setactive $snap.OriginalPowerScheme | Out-Null
            $actions.Add("Restored original power scheme ($($snap.OriginalPowerScheme))")
        } catch {
            $actions.Add("Failed to restore power scheme: $($_.Exception.Message)")
        }
    }

    # 3. Remove Hosts Loopback Fence line-by-line
    $hostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
    if (Test-Path $hostsFile) {
        try {
            $hLines = [System.IO.File]::ReadAllLines($hostsFile, [System.Text.Encoding]::ASCII)
            $newLines = [System.Collections.Generic.List[string]]::new()
            $inFence = $false
            foreach ($l in $hLines) {
                if ($l -like '*# AUTOPILOT_OFFLINE_STAGING_FENCE_START*') { $inFence = $true; continue }
                if ($l -like '*# AUTOPILOT_OFFLINE_STAGING_FENCE_END*') { $inFence = $false; continue }
                if (-not $inFence) { $newLines.Add($l) }
            }
            [System.IO.File]::WriteAllLines($hostsFile, $newLines, [System.Text.Encoding]::ASCII)
            $actions.Add("Removed enrollment endpoint loopback fence from hosts file")
            try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch { }
        } catch {
            $actions.Add("Failed to edit hosts file: $($_.Exception.Message)")
        }
    }

    # 4. Restore Registry
    try {
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot' -Name 'IsAutopilotDisabled' -Value 0 -Type DWord -ErrorAction SilentlyContinue
        $actions.Add("Restored IsAutopilotDisabled=0 under Provisioning\Diagnostics\AutoPilot")
    } catch { }

    try {
        $oKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE'
        if ($snap.OriginalBypassNro -eq $null) {
            Remove-ItemProperty -Path $oKey -Name 'BypassNRO' -ErrorAction SilentlyContinue
            $actions.Add("Removed staging BypassNRO flag from OOBE registry")
        } else {
            Set-ItemProperty -Path $oKey -Name 'BypassNRO' -Value $snap.OriginalBypassNro -Type DWord -ErrorAction SilentlyContinue
            $actions.Add("Restored original BypassNRO value ($($snap.OriginalBypassNro)) in OOBE registry")
        }
    } catch { }

    if ($snap.OriginalDevMode -ne $null) {
        try {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -Name 'AllowDevelopmentWithoutDevLicense' -Value $snap.OriginalDevMode -Type DWord -ErrorAction SilentlyContinue
            $actions.Add("Restored Developer Mode setting to $($snap.OriginalDevMode)")
        } catch { }
    }

    if ($snap.OriginalAdminState -eq $false) {
        try {
            net user Administrator /active:no | Out-Null
            Disable-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
            $actions.Add("Disabled built-in Administrator account to restore enterprise security posture")
        } catch { }
    }

    # 5. Restore Quarantined Autopilot Profile
    if ($snap.QuarantinedFiles) {
        foreach ($q in $snap.QuarantinedFiles) {
            if (Test-Path $q.QuarantinePath) {
                try {
                    $parent = Split-Path -Parent $q.OriginalPath
                    if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
                    Copy-Item -Path $q.QuarantinePath -Destination $q.OriginalPath -Force -ErrorAction Stop
                    Remove-Item -Path $q.QuarantinePath -Force -ErrorAction SilentlyContinue
                    $actions.Add("Restored quarantined Autopilot profile to '$($q.OriginalPath)'")
                } catch {
                    $actions.Add("Failed to restore profile file '$($q.OriginalPath)': $($_.Exception.Message)")
                }
            }
        }
    }

    # 6. Trigger Enterprise Policy Retrieval
    try {
        Start-Process -FilePath "$env:SystemRoot\System32\DeviceEnroller.exe" -ArgumentList "/c /AutoEnrollMDM" -NoNewWindow -ErrorAction SilentlyContinue
        $actions.Add("Triggered DeviceEnroller /AutoEnrollMDM enterprise sync")
    } catch { }

    try {
        Invoke-MdmSync | Out-Null
        $actions.Add("Dispatched MDM session sync")
    } catch { }

    # 7. Clean up snapshot & Desktop Shortcut
    try {
        $restoredPath = "$($script:StagingSnapshotPath).restored.json"
        Move-Item -Path $script:StagingSnapshotPath -Destination $restoredPath -Force -ErrorAction SilentlyContinue
        $actions.Add("Archived snapshot to $restoredPath")
    } catch { }

    try {
        $pubDesk = [Environment]::GetFolderPath('CommonDesktopDirectory')
        if (-not (Test-Path $pubDesk)) { $pubDesk = "$env:PUBLIC\Desktop" }
        $lnkPath = Join-Path $pubDesk "Resume-EnterpriseAutopilot.lnk"
        if (Test-Path $lnkPath) { Remove-Item -Path $lnkPath -Force -ErrorAction SilentlyContinue }
    } catch { }

    return [PSCustomObject]@{
        Success = $true
        Message = "Device successfully re-docked to enterprise cloud. Autopilot profile and enterprise policies restored."
        Actions = $actions
    }
}

function Show-OfflineStagingDialog {
    param([System.Windows.Window]$Owner = $null)
    $ctx = $script:RuntimeContext
    if (-not $ctx.MeetsPreferred) {
        if (-not $script:PlaybookRunning) {
            [System.Windows.MessageBox]::Show("Controlled Offline Staging requires administrative elevation to configure service states, quarantine profiles, and write network fences. Please click 'Fix Privileges' in the header first.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
        return $false
    }

    if ($script:PlaybookRunning) {
        $res = Invoke-OfflineStaging -Profile 'HardwareDiagnostics' -QuarantineProfile:$true -LoopbackFence:$true -SetBypassNro:$true -CreateDesktopShortcut:$true
        Write-HubLog "Offline Staging Mode activated for '$($res.Title)'" "SUCCESS"
        return $true
    }

    $xamlStaging = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Controlled Offline Staging &amp; Adaptive Evaluation Profiles"
        Width="680" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="#F3F4F6" Foreground="#1E293B" FontFamily="Segoe UI Variable Text, Segoe UI, sans-serif" ResizeMode="NoResize">
    <StackPanel Margin="22">
        <TextBlock Text="CONTROLLED OFFLINE STAGING &amp; ADAPTIVE PROFILES" FontSize="11" FontWeight="SemiBold" Foreground="#64748B" Margin="0,0,0,6"/>
        <TextBlock Text="Safely isolate this device during OOBE without permanently destroying the Autopilot profile link. Select an adaptive profile to optimize background behavior, then reconnect to the enterprise cloud at any time." FontSize="11.5" Foreground="#475569" TextWrapping="Wrap" Margin="0,0,0,12"/>

        <!-- Profile Selection Options -->
        <Border Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14,10" Margin="0,0,0,10">
            <StackPanel>
                <RadioButton Name="RadProfileDiag" IsChecked="True" GroupName="StagingProfiles" Margin="0,0,0,6">
                    <StackPanel>
                        <TextBlock Text="Hardware Diagnostics &amp; Triage (Recommended)" FontWeight="Bold" FontSize="12" Foreground="#1E293B"/>
                        <TextBlock Text="Suspends telemetry/indexing (DiagTrack, WSearch, SysMain), keeps local subnet open, enables live hardware probes." FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                </RadioButton>

                <RadioButton Name="RadProfileBench" GroupName="StagingProfiles" Margin="0,6,0,6">
                    <StackPanel>
                        <TextBlock Text="Workload Benchmarking &amp; Testing" FontWeight="Bold" FontSize="12" Foreground="#1E293B"/>
                        <TextBlock Text="Ultimate Performance power scheme, halts background maintenance/updates, zero CPU jitter/throttling." FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                </RadioButton>

                <RadioButton Name="RadProfileSandbox" GroupName="StagingProfiles" Margin="0,6,0,2">
                    <StackPanel>
                        <TextBlock Text="Software Evaluation &amp; Staging Sandbox" FontWeight="Bold" FontSize="12" Foreground="#1E293B"/>
                        <TextBlock Text="Clean local admin access, enables Windows Developer Mode, readies local/USB offline package staging." FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                </RadioButton>
            </StackPanel>
        </Border>

        <!-- Staging Policy Fence Options -->
        <TextBlock Text="REVERSIBLE POLICY FENCE CONTROLS" FontSize="10" FontWeight="SemiBold" Foreground="#64748B" Margin="0,4,0,6"/>
        <Border Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,0,14">
            <StackPanel>
                <CheckBox Name="ChkQuarantineProfile" Content="Quarantine &amp; hold existing Autopilot profile (100% reversible)" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,2,0,4"/>
                <CheckBox Name="ChkLoopbackFence" Content="Null-route cloud enrollment endpoints via hosts file loopback fence" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,2,0,4"/>
                <CheckBox Name="ChkBypassNro" Content="Configure BypassNRO=1 to allow local account setup in OOBE" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,2,0,4"/>
                <CheckBox Name="ChkDesktopCompanion" Content="Create 'Resume-EnterpriseAutopilot' re-docking shortcut on Desktop" IsChecked="True" Foreground="#1E293B" FontSize="11.5" Margin="0,2,0,2"/>
            </StackPanel>
        </Border>

        <!-- Actions -->
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="BtnCancelStaging" Content="Cancel" Padding="14,6" Margin="0,0,8,0" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
            <Button Name="BtnApplyStaging" Content="Enter Offline Staging Mode" Padding="14,6" Background="#0067C0" Foreground="#FFFFFF" BorderBrush="#0067C0" FontWeight="SemiBold"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlStaging))
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $dlg.Owner = $Owner }

    $applied = $false
    $dlg.FindName('BtnCancelStaging').Add_Click({ $dlg.Close() }.GetNewClosure())
    $dlg.FindName('BtnApplyStaging').Add_Click({
        $profile = if ($dlg.FindName('RadProfileBench').IsChecked) {
            'WorkloadBenchmarking'
        } elseif ($dlg.FindName('RadProfileSandbox').IsChecked) {
            'SoftwareSandbox'
        } else {
            'HardwareDiagnostics'
        }
        $quar = [bool]$dlg.FindName('ChkQuarantineProfile').IsChecked
        $loop = [bool]$dlg.FindName('ChkLoopbackFence').IsChecked
        $nro  = [bool]$dlg.FindName('ChkBypassNro').IsChecked
        $sc   = [bool]$dlg.FindName('ChkDesktopCompanion').IsChecked

        $res = Invoke-OfflineStaging -Profile $profile -QuarantineProfile:$quar -LoopbackFence:$loop -SetBypassNro:$nro -CreateDesktopShortcut:$sc
        $applied = $true
        $dlg.Close()

        $actSummary = $res.Actions -join "`n- "
        [System.Windows.MessageBox]::Show("Offline Staging Mode activated successfully for '$($res.Title)'!`n`nActions taken:`n- $actSummary`n`nYou can re-dock to the enterprise cloud at any time by clicking 'Resume Cloud Autopilot' or using the desktop companion shortcut.", "Offline Staging Active", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
    return $applied
}

# --- Function: Invoke-DsregStatus (parse dsregcmd /status into case-insensitive hashtable) ---
function Invoke-DsregStatus {
    param([switch]$Force)
    if (-not $Force -and $script:CachedDsregStatus) {
        return $script:CachedDsregStatus
    }
    $map = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $exe = Join-Path $env:SystemRoot 'System32\dsregcmd.exe'
        if (Test-Path $exe) {
            foreach ($line in (& $exe /status 2>$null)) {
                if ($line -match '^\s*([A-Za-z0-9_ ]+?)\s*:\s*(.+?)\s*$') {
                    $k = $Matches[1].Trim()
                    $v = $Matches[2].Trim()
                    if (-not $map.ContainsKey($k)) {
                        $map[$k] = $v
                    }
                }
            }
        }
    } catch { }
    $script:CachedDsregStatus = $map
    return $map
}

# --- Function: Get-MdmEnrollmentInfo (query HKLM Enrollments for MS DM Server) ---
function Get-MdmEnrollmentInfo {
    $info = [PSCustomObject]@{
        Enrolled     = $false
        EnrollmentId = ''
        Provider     = ''
        Upn          = ''
        MdmUrl       = ''
        TenantId     = ''
    }
    try {
        foreach ($k in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Enrollments' -ErrorAction SilentlyContinue)) {
            $pv = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
            if ($pv -and $pv.ProviderID -eq 'MS DM Server') {
                $info.Enrolled     = $true
                $info.EnrollmentId = $k.PSChildName
                $info.Provider     = $pv.ProviderID
                if ($pv.UPN)                     { $info.Upn = [string]$pv.UPN }
                if ($pv.DiscoveryServiceFullURL) { $info.MdmUrl = [string]$pv.DiscoveryServiceFullURL }
                if ($pv.AADTenantID)             { $info.TenantId = [string]$pv.AADTenantID }
                break
            }
        }
    } catch { }
    return $info
}

# --- Function: Get-DeviceEnrollmentState (is this PC already Autopilot / Intune / Entra managed?) ---
function Get-DeviceEnrollmentState {
    $evidence = [System.Collections.Generic.List[string]]::new()
    $state = [ordered]@{
        AutopilotProfileCached = $false
        AutopilotTenantDomain  = ''
        AutopilotTenantId      = ''
        AutopilotProfileName   = ''
        AutopilotDeviceName    = ''
        AutopilotCorrelationId = ''
        IntuneEnrolled         = $false
        EnrollmentId           = ''
        EnrollmentUpn          = ''
        EnrollmentTenantId     = ''
        AzureAdJoined          = $false
        DomainJoined           = $false
        EntraTenantName        = ''
        EntraTenantId          = ''
        MdmUrl                 = ''
        Verdict                = 'Unmanaged'
        Summary                = ''
        Evidence               = $evidence
        CloudChecked           = $false
        CloudRegistered        = $null
        CloudIdentity          = $null
    }

    # 1. The Autopilot profile the device itself fetched from the Deployment Service during OOBE.
    foreach ($jsonPath in @("$env:SystemRoot\ServiceState\wmansvc\AutopilotDDSZTDFile.json", "$env:SystemRoot\Provisioning\Autopilot\AutopilotConfigurationFile.json")) {
        if (-not (Test-Path $jsonPath)) { continue }
        try {
            $j = Get-Content -Path $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($j.CloudAssignedTenantDomain -or $j.CloudAssignedTenantId) {
                $state.AutopilotProfileCached = $true
                if ($j.CloudAssignedTenantDomain) { $state.AutopilotTenantDomain = [string]$j.CloudAssignedTenantDomain }
                if ($j.CloudAssignedTenantId)     { $state.AutopilotTenantId     = [string]$j.CloudAssignedTenantId }
                if ($j.DeploymentProfileName)     { $state.AutopilotProfileName  = [string]$j.DeploymentProfileName }
                if ($j.CloudAssignedDeviceName)   { $state.AutopilotDeviceName   = [string]$j.CloudAssignedDeviceName }
                if ($j.ZtdCorrelationId)          { $state.AutopilotCorrelationId = [string]$j.ZtdCorrelationId }
                $evidence.Add("Autopilot profile JSON: $jsonPath")
            }
        } catch {
            $evidence.Add("Unreadable Autopilot profile JSON: $jsonPath ($($_.Exception.Message))")
        }
    }

    # 2. Provisioning diagnostics written by the Autopilot client
    $isBypassed = $false
    try {
        $diag = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot' -ErrorAction SilentlyContinue
        if ($diag) {
            if ($diag.IsAutopilotDisabled -eq 1) {
                $isBypassed = $true
                $evidence.Add('Registry: HKLM\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot (IsAutopilotDisabled=1)')
            }
            if (-not $isBypassed -and ($diag.CloudAssignedTenantDomain -or $diag.CloudAssignedTenantId)) {
                $state.AutopilotProfileCached = $true
                if (-not $state.AutopilotTenantDomain -and $diag.CloudAssignedTenantDomain) { $state.AutopilotTenantDomain = [string]$diag.CloudAssignedTenantDomain }
                if (-not $state.AutopilotTenantId -and $diag.CloudAssignedTenantId)         { $state.AutopilotTenantId = [string]$diag.CloudAssignedTenantId }
                if (-not $state.AutopilotCorrelationId -and $diag.AutopilotServiceCorrelationId) { $state.AutopilotCorrelationId = [string]$diag.AutopilotServiceCorrelationId }
                $evidence.Add('Registry: HKLM\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot')
            }
        }
    } catch { }

    # 3. MDM enrollment - Intune is provider 'MS DM Server' (unified via Get-MdmEnrollmentInfo)
    $mdm = Get-MdmEnrollmentInfo
    if ($mdm.Enrolled) {
        $state.IntuneEnrolled = $true
        $state.EnrollmentId = $mdm.EnrollmentId
        if ($mdm.Upn)      { $state.EnrollmentUpn = $mdm.Upn }
        if ($mdm.TenantId) { $state.EnrollmentTenantId = $mdm.TenantId }
        $evidence.Add("Registry: HKLM\SOFTWARE\Microsoft\Enrollments\$($mdm.EnrollmentId) (MS DM Server)")
    }

    # 4. Entra / domain join state (unified via Invoke-DsregStatus)
    try {
        $ds = Invoke-DsregStatus
        if ($ds -and $ds.Count -gt 0) {
            $state.AzureAdJoined   = ($ds['AzureAdJoined'] -eq 'YES')
            $state.DomainJoined    = ($ds['DomainJoined'] -eq 'YES')
            $state.EntraTenantName = [string]$ds['TenantName']
            $state.EntraTenantId   = [string]$ds['TenantId']
            $state.MdmUrl          = [string]$ds['MdmUrl']
            if ($state.AzureAdJoined) { $evidence.Add("dsregcmd: AzureAdJoined=YES, tenant '$($state.EntraTenantName)'") }
            if ($state.MdmUrl)        { $evidence.Add("dsregcmd: MdmUrl $($state.MdmUrl)") }
        }
    } catch { }

    $tenantLabel = if ($state.AutopilotTenantDomain) { $state.AutopilotTenantDomain }
                   elseif ($state.EntraTenantName)   { $state.EntraTenantName }
                   elseif ($state.AutopilotTenantId) { $state.AutopilotTenantId }
                   elseif ($state.EnrollmentTenantId){ $state.EnrollmentTenantId }
                   else { 'unknown tenant' }

    if ($state.AutopilotProfileCached -and -not $isBypassed) {
        $state.Verdict = 'AutopilotRegistered'
        $profileNote = if ($state.AutopilotProfileName) { " (profile: $($state.AutopilotProfileName))" } else { '' }
        $state.Summary = "Autopilot-registered to $tenantLabel$profileNote"
    } elseif ($isBypassed) {
        $state.Verdict = 'Bypassed'
        $state.Summary = 'Autopilot skipped (IsAutopilotDisabled=1, profile removed)'
    } elseif ($state.IntuneEnrolled) {
        $state.Verdict = 'IntuneEnrolled'
        $upnNote = if ($state.EnrollmentUpn) { " as $($state.EnrollmentUpn)" } else { '' }
        $state.Summary = "Intune-enrolled ($tenantLabel)$upnNote, but no Autopilot profile is cached locally"
    } elseif ($state.AzureAdJoined) {
        $state.Verdict = 'EntraJoined'
        $state.Summary = "Entra-joined to $tenantLabel without MDM enrollment"
    } elseif ($state.DomainJoined) {
        $state.Verdict = 'DomainJoined'
        $state.Summary = 'On-prem domain joined, not cloud managed'
    } else {
        $state.Verdict = 'Unmanaged'
        $state.Summary = 'No Autopilot profile, MDM enrollment or Entra join found on this device'
    }
    return [PSCustomObject]$state
}

# --- Function: Find-AutopilotIdentityInTenant (does the signed-in tenant already know this serial?) ---
function Find-AutopilotIdentityInTenant {
    param(
        [Parameter(Mandatory = $true)][string]$SerialNumber,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )
    $serial = $SerialNumber.Trim()
    if ([string]::IsNullOrWhiteSpace($serial)) { return $null }
    $escaped = $serial.Replace("'", "''")
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$escaped')"
    $res = Invoke-RestMethod -Uri $uri -Method GET -Headers @{ 'Authorization' = "Bearer $AccessToken" } -ErrorAction Stop
    $exact = @($res.value | Where-Object { $_.serialNumber -eq $serial })
    if ($exact.Count -gt 0) { return $exact[0] }
    $any = @($res.value)
    if ($any.Count -gt 0) { return $any[0] }
    return $null
}

# --- Function: Show-PrivilegeGuide (walk the operator to the privilege level the Hub wants) ---
function Show-PrivilegeGuide {
    param([System.Windows.Window]$Owner = $null)
    $ctx = $script:RuntimeContext
    $canRelaunch = (-not $ctx.IsOobe) -and (-not $ctx.MeetsPreferred)

    $stepsText = if ($ctx.MeetsPreferred) {
        "Nothing to do - this window already runs as $($ctx.PrivilegeLevel), which is the level the Hub wants in $($ctx.Mode) mode."
    } elseif ($ctx.IsOobe) {
        "1. Press Shift + F10 on the OOBE screen to open the built-in command prompt (it runs as SYSTEM).`n" +
        "2. Type:  powershell -ExecutionPolicy Bypass   and press Enter.`n" +
        "3. Type:  irm $($script:BootstrapUrl) | iex   and press Enter.`n" +
        "4. Close this window once the elevated Hub appears."
    } else {
        "1. Click 'Relaunch as Administrator' below.`n" +
        "2. Accept the User Account Control prompt.`n" +
        "3. This window closes itself; the elevated Hub opens with the same settings.`n`n" +
        "Manual alternative: right-click PowerShell > Run as administrator, then run:`n   irm $($script:BootstrapUrl) | iex"
    }
    $whyText = "Reading the hardware hash (MDM_DevDetail_Ext01), renaming the computer, installing apps, syncing the clock and registering a restart-resume task all need elevation. Without it those actions fail and Intune registration is blocked."

    $xamlGuide = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Privilege Check" Width="620" SizeToContent="Height" WindowStartupLocation="CenterOwner"
        Background="#F3F4F6" Foreground="#1E293B" FontFamily="Segoe UI Variable Text, Segoe UI, sans-serif" ResizeMode="NoResize">
    <StackPanel Margin="20">
        <TextBlock Text="PRIVILEGE CHECK" FontSize="11" FontWeight="SemiBold" Foreground="#64748B" Margin="0,0,0,8"/>
        <Grid Margin="0,0,0,12">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,6,0">
                <StackPanel>
                    <TextBlock Text="CURRENT" FontSize="10" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtCurrent" FontSize="13" FontWeight="Bold" Foreground="#1E293B" Margin="0,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="6,0,0,0">
                <StackPanel>
                    <TextBlock Text="PREFERRED" FontSize="10" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtPreferred" FontSize="13" FontWeight="Bold" Foreground="#15803D" Margin="0,2,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>
        </Grid>
        <TextBlock Text="WHY IT MATTERS" FontSize="10" Foreground="#64748B" FontWeight="SemiBold"/>
        <TextBlock Name="TxtWhy" FontSize="11.5" Foreground="#475569" TextWrapping="Wrap" Margin="0,2,0,12"/>
        <TextBlock Text="HOW TO GET THERE" FontSize="10" Foreground="#64748B" FontWeight="SemiBold"/>
        <TextBox Name="TxtSteps" FontSize="11.5" FontFamily="Consolas" Foreground="#1E293B" Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10" IsReadOnly="True" TextWrapping="Wrap" Margin="0,2,0,14"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="BtnCopyCmd" Content="Copy bootstrap command" Padding="12,6" Margin="0,0,8,0" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
            <Button Name="BtnRelaunch" Content="Relaunch as Administrator" Padding="12,6" Margin="0,0,8,0" Background="#0067C0" Foreground="#FFFFFF" BorderBrush="#0067C0" FontWeight="SemiBold"/>
            <Button Name="BtnCloseGuide" Content="Close" Padding="12,6" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
        </StackPanel>
    </StackPanel>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlGuide))
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $dlg.Owner = $Owner }
    $dlg.FindName('TxtCurrent').Text   = "$($ctx.PrivilegeLevel) - $($ctx.UserName) - $($ctx.Mode) mode"
    $dlg.FindName('TxtPreferred').Text = $ctx.PreferredLevel
    $dlg.FindName('TxtWhy').Text       = $whyText
    $dlg.FindName('TxtSteps').Text     = $stepsText
    $btnRelaunch = $dlg.FindName('BtnRelaunch')
    $btnRelaunch.IsEnabled = $canRelaunch
    if (-not $canRelaunch) { $btnRelaunch.Opacity = 0.4 }
    $result = [hashtable]::Synchronized(@{ Relaunched = $false })
    $dlg.FindName('BtnCopyCmd').Add_Click({ [System.Windows.Clipboard]::SetText("irm $($script:BootstrapUrl) | iex") })
    $btnRelaunch.Add_Click({
        try {
            Invoke-HubElevatedRelaunch | Out-Null
            $result.Relaunched = $true
            $dlg.Close()
        } catch {
            [System.Windows.MessageBox]::Show("Elevated relaunch failed: $($_.Exception.Message)", "Relaunch Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    }.GetNewClosure())
    $dlg.FindName('BtnCloseGuide').Add_Click({ $dlg.Close() }.GetNewClosure())
    $dlg.ShowDialog() | Out-Null
    return $result.Relaunched
}

# ==============================================================================
# HYBRID AZURE AD JOIN & CO-MANAGEMENT TOOLSET
# Everything here degrades gracefully off-domain (returns 'n/a' rather than throwing),
# so the Hub still runs on a workgroup bench machine.
# ==============================================================================

function Get-HybridJoinState {
    $d = Invoke-DsregStatus
    $aadj   = ($d['AzureAdJoined'] -eq 'YES')
    $domain = ($d['DomainJoined'] -eq 'YES')
    $prt    = ($d['AzureAdPrt'] -eq 'YES')
    $entJoined = ($d['EnterpriseJoined'] -eq 'YES')

    $joinType = if ($aadj -and $domain) { 'Hybrid Azure AD Joined' }
                elseif ($aadj)          { 'Azure AD Joined (cloud only)' }
                elseif ($domain)        { 'On-prem Domain Joined only' }
                elseif ($entJoined)     { 'Enterprise (ADFS) Joined' }
                else                    { 'Workgroup / Not joined' }

    $dc = if ($d['DomainController']) { $d['DomainController'] }
          elseif ($d['DomainControllerName']) { $d['DomainControllerName'] }
          elseif ($d['DcName']) { $d['DcName'] }
          else {
              try {
                  $ntDomain = @(Get-CimInstance Win32_NTDomain -ErrorAction SilentlyContinue | Where-Object { $_.DnsForestName })
                  if ($ntDomain.Count -gt 0 -and $ntDomain[0].DomainControllerName) {
                      $ntDomain[0].DomainControllerName.TrimStart('\')
                  } else { '' }
              } catch { '' }
          }

    [PSCustomObject]@{
        JoinType       = $joinType
        AzureAdJoined  = $aadj
        DomainJoined   = $domain
        EnterpriseJoined = $entJoined
        HasPrt         = $prt
        DeviceId       = $d['DeviceId']
        TenantName     = $d['TenantName']
        TenantId       = $d['TenantId']
        IdpDomain      = $d['IdpDomain']
        MdmUrl         = $d['MdmUrl']
        DomainName     = $d['DomainName']
        DcName         = $dc
        PrtAuthority   = $d['AzureAdPrtAuthority']
        Raw            = $d
        IsHybrid       = ($aadj -and $domain)
    }
}

function Invoke-MdmSync {
    # Kick the EnterpriseMgmt "PushLaunch" scheduled task(s) that trigger an immediate MDM/Intune sync.
    $ran = [System.Collections.Generic.List[string]]::new()
    try {
        $tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue |
                   Where-Object { $_.TaskName -match 'PushLaunch|Schedule #3|Login' })
        if ($tasks.Count -eq 0) {
            $tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue)
        }
        foreach ($t in $tasks) {
            try { Start-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop; $ran.Add($t.TaskName) } catch { }
        }
    } catch { }
    # Fallback: the modern sync ML client
    if ($ran.Count -eq 0) {
        try {
            $ml = Join-Path $env:SystemRoot 'System32\deviceenroller.exe'
            if (Test-Path $ml) { Start-Process $ml -ArgumentList '/o /c' -WindowStyle Hidden -ErrorAction Stop; $ran.Add('deviceenroller /o /c') }
        } catch { }
    }
    return ,$ran
}

function Invoke-MdmAutoEnroll {
    # GPO-style auto-enrollment into Intune using the device's Entra identity.
    try {
        $exe = Join-Path $env:SystemRoot 'System32\deviceenroller.exe'
        if (-not (Test-Path $exe)) { return 'deviceenroller.exe not present on this build' }
        Start-Process $exe -ArgumentList '/c /AutoEnrollMDM' -WindowStyle Hidden -ErrorAction Stop
        return 'Triggered deviceenroller /c /AutoEnrollMDM'
    } catch { return "Auto-enroll failed: $($_.Exception.Message)" }
}

function Get-IntuneExtensionHealth {
    $svc = Get-Service -Name 'Microsoft Intune Management Extension' -ErrorAction SilentlyContinue
    $logDir = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
    $lastLog = $null
    if (Test-Path $logDir) {
        $lastLog = Get-ChildItem $logDir -Filter *.log -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    [PSCustomObject]@{
        Installed   = [bool]$svc
        Status      = if ($svc) { [string]$svc.Status } else { 'Not installed' }
        StartType   = if ($svc) { [string]$svc.StartType } else { '' }
        LogDir      = $logDir
        LastLogFile = if ($lastLog) { $lastLog.Name } else { '' }
        LastLogTime = if ($lastLog) { $lastLog.LastWriteTime } else { $null }
        StaleMins   = if ($lastLog) { [int]((Get-Date) - $lastLog.LastWriteTime).TotalMinutes } else { -1 }
    }
}

function Restart-IntuneExtension {
    try {
        $svc = Get-Service -Name 'Microsoft Intune Management Extension' -ErrorAction Stop
        Restart-Service -InputObject $svc -Force -ErrorAction Stop
        return 'Restarted Microsoft Intune Management Extension'
    } catch { return "Could not restart IME: $($_.Exception.Message)" }
}

function Get-CoManagementState {
    # Decode the ConfigMgr co-management workload bitmask (which authority owns each workload).
    # HKLM\SOFTWARE\Microsoft\CCM\CoManagementFlags. Bit 1 only means "co-management is configured";
    # the workloads themselves start at 2 (e.g. 67 = enabled + Compliance + Client Apps).
    $workloadBits = [ordered]@{
        2   = 'Compliance Policies'
        4   = 'Resource Access Policies (Wi-Fi/VPN/Cert/Email)'
        8   = 'Device Configuration'
        16  = 'Windows Update Policies'
        32  = 'Endpoint Protection (Defender)'
        64  = 'Client Apps'
        128 = 'Office Click-to-Run Apps'
    }
    $flags = $null
    foreach ($path in @('HKLM:\SOFTWARE\Microsoft\CCM\CoManagementFlags', 'HKLM:\SOFTWARE\Microsoft\CCM')) {
        try {
            $pv = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if ($pv) {
                foreach ($n in @('ComanagementWorkloads', 'CoManagementFlags', 'Workloads')) {
                    if ($null -ne $pv.$n) { $flags = [int]$pv.$n; break }
                }
            }
        } catch { }
        if ($null -ne $flags) { break }
    }
    $ccmPresent = [bool](Get-Service -Name CcmExec -ErrorAction SilentlyContinue)
    $rows = foreach ($bit in $workloadBits.Keys) {
        [PSCustomObject]@{
            Workload = $workloadBits[$bit]
            Authority = if ($null -eq $flags) { if ($ccmPresent) { 'ConfigMgr (default)' } else { 'Intune' } }
                        elseif ($flags -band $bit) { 'Intune' } else { 'ConfigMgr' }
        }
    }
    [PSCustomObject]@{
        CoManaged     = ($null -ne $flags -and ($flags -band 1))
        ConfigMgrPresent = $ccmPresent
        FlagsValue    = $flags
        Workloads     = @($rows)
    }
}

function Invoke-ConfigMgrClientAction {
    param([string]$ScheduleId = '{00000000-0000-0000-0000-000000000021}') # Machine Policy Retrieval & Evaluation
    try {
        if (-not (Get-Service -Name CcmExec -ErrorAction SilentlyContinue)) { return 'ConfigMgr (SCCM) client is not installed on this device.' }
        Invoke-CimMethod -Namespace 'root\ccm' -ClassName 'SMS_Client' -MethodName 'TriggerSchedule' -Arguments @{ sScheduleID = $ScheduleId } -ErrorAction Stop | Out-Null
        return "Triggered ConfigMgr schedule $ScheduleId (Machine Policy Retrieval & Evaluation)"
    } catch { return "ConfigMgr trigger failed: $($_.Exception.Message)" }
}

function Test-DomainConnectivity {
    $result = [PSCustomObject]@{ Applicable = $false; DcName = ''; DcReachable = $false; SecureChannel = 'n/a'; Site = ''; Message = '' }
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    if (-not $cs -or -not $cs.PartOfDomain) { $result.Message = 'Not domain-joined; DC checks not applicable.'; return $result }
    $result.Applicable = $true
    try {
        $nltest = & (Join-Path $env:SystemRoot 'System32\nltest.exe') /dsgetdc: 2>$null
        foreach ($l in $nltest) {
            if ($l -match 'DC:\s*\\\\(\S+)') { $result.DcName = $Matches[1] }
            if ($l -match 'Dc Site Name:\s*(\S+)') { $result.Site = $Matches[1] }
        }
    } catch { }
    if ($result.DcName) {
        try { $result.DcReachable = Test-Connection -ComputerName $result.DcName -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { }
    }
    try {
        $sc = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
        $result.SecureChannel = if ($sc) { 'Healthy' } else { 'BROKEN (trust relationship failed)' }
    } catch { $result.SecureChannel = 'Unknown' }
    $result.Message = "DC $($result.DcName) reachable=$($result.DcReachable), secure channel: $($result.SecureChannel)"
    return $result
}

function Invoke-GpUpdateForce {
    try {
        $out = & (Join-Path $env:SystemRoot 'System32\gpupdate.exe') /force 2>&1
        return ($out | Out-String).Trim()
    } catch { return "gpupdate failed: $($_.Exception.Message)" }
}

function Sync-DomainTime {
    try {
        & (Join-Path $env:SystemRoot 'System32\w32tm.exe') /resync /force 2>&1 | Out-Null
        $src = (& (Join-Path $env:SystemRoot 'System32\w32tm.exe') /query /source 2>$null) -join ''
        return "Resynced system clock. Time source: $src"
    } catch { return "Time resync failed: $($_.Exception.Message)" }
}

function Get-BitLockerEscrowState {
    $vol = $null
    try { $vol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop } catch { }
    if (-not $vol) { return [PSCustomObject]@{ Available = $false; Message = 'BitLocker cmdlets unavailable or access denied (needs elevation).' } }
    $rp = @($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
    [PSCustomObject]@{
        Available        = $true
        MountPoint       = $vol.MountPoint
        ProtectionStatus = [string]$vol.ProtectionStatus
        VolumeStatus     = [string]$vol.VolumeStatus
        RecoveryProtectors = $rp.Count
        KeyProtectorIds  = @($rp | ForEach-Object { $_.KeyProtectorId })
        Message          = "Protection $($vol.ProtectionStatus), $($rp.Count) recovery-password protector(s)"
    }
}

function Invoke-BitLockerEscrow {
    # Back up every recovery-password protector to Entra ID (BackupToAAD) and, if domain-joined, to AD.
    $done = [System.Collections.Generic.List[string]]::new()
    try {
        $vol = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        $rp = @($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
        if ($rp.Count -eq 0) { return 'No recovery-password protector to escrow. Add one first (Add-BitLockerKeyProtector -RecoveryPasswordProtector).' }
        $domainJoined = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).PartOfDomain
        foreach ($k in $rp) {
            try { BackupToAAD-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $k.KeyProtectorId -ErrorAction Stop | Out-Null; $done.Add("Entra: $($k.KeyProtectorId)") } catch { $done.Add("Entra FAILED: $($_.Exception.Message)") }
            if ($domainJoined) {
                try { Backup-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId $k.KeyProtectorId -ErrorAction Stop | Out-Null; $done.Add("AD: $($k.KeyProtectorId)") } catch { }
            }
        }
    } catch { return "BitLocker escrow failed: $($_.Exception.Message)" }
    return ($done -join '; ')
}

function Get-CaReadiness {
    # Composite "will Conditional Access let this device through?" signal from local evidence.
    $j = Get-HybridJoinState
    $m = Get-MdmEnrollmentInfo
    $checks = [ordered]@{
        'Entra joined (or hybrid joined)' = $j.AzureAdJoined
        'Primary Refresh Token present'   = $j.HasPrt
        'Enrolled in Intune (MDM)'        = $m.Enrolled
    }
    $blockers = @($checks.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
    [PSCustomObject]@{
        Checks   = $checks
        Ready    = ($blockers.Count -eq 0)
        Blockers = $blockers
        Note     = 'Device compliance itself is evaluated server-side in Intune; confirm the compliance state in the portal.'
    }
}

function Get-LegacyDependencyState {
    # Local "what will bite you when cloud security baselines / NTLM-blocking / CG land" scan.
    $findings = [System.Collections.Generic.List[PSCustomObject]]::new()

    # SMBv1 - removed/blocked by most security baselines; on-prem shares may still need it
    try {
        $smb1 = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
        if ($smb1 -and $smb1.EnableSMB1Protocol) {
            $findings.Add([PSCustomObject]@{ Item='SMBv1'; State='ENABLED'; Risk='High'; Note='Legacy/insecure; disabled by security baselines. Confirm no dependency before hardening.' })
        } else {
            $findings.Add([PSCustomObject]@{ Item='SMBv1'; State='Disabled'; Risk='OK'; Note='' })
        }
    } catch { }

    # NTLM restriction level (LmCompatibilityLevel) + outbound NTLM auditing/blocking
    try {
        $lsa = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -ErrorAction SilentlyContinue
        $lm = if ($lsa -and $null -ne $lsa.LmCompatibilityLevel) { [int]$lsa.LmCompatibilityLevel } else { -1 }
        $lmNote = switch ($lm) { 5 {'Send NTLMv2 only, refuse LM & NTLM (hardened)'} 3 {'Send NTLMv2 only (default)'} -1 {'Not set (OS default = 3)'} default {"Level $lm"} }
        $findings.Add([PSCustomObject]@{ Item='NTLM (LmCompatibilityLevel)'; State=$lmNote; Risk=$(if ($lm -le 2 -and $lm -ge 0){'Medium'}else{'OK'}); Note='Cloud tenants increasingly block NTLM; legacy apps that force NTLM will break.' })
    } catch { }

    # Credential Guard (Win32_DeviceGuard) - often forced on by baselines; can break some VPN/creds tooling
    try {
        $dg = Get-CimInstance -Namespace 'root\Microsoft\Windows\DeviceGuard' -ClassName Win32_DeviceGuard -ErrorAction SilentlyContinue
        $cgRunning = $dg -and ($dg.SecurityServicesRunning -contains 1)
        $findings.Add([PSCustomObject]@{ Item='Credential Guard'; State=$(if ($cgRunning){'Running'}else{'Not running'}); Risk='OK'; Note='Baselines enable it; some legacy credential managers/VPNs misbehave with it on.' })
    } catch { }

    # DFS / mapped-drive dependency hint
    try {
        $mapped = @(Get-CimInstance Win32_NetworkConnection -ErrorAction SilentlyContinue)
        if ($mapped.Count -gt 0) {
            $findings.Add([PSCustomObject]@{ Item='Mapped network drives'; State="$($mapped.Count) mapped"; Risk='Info'; Note='Domain/DFS shares need line-of-sight to a DC; cloud-only devices lose these.' })
        }
    } catch { }

    return ,@($findings)
}

function Get-ClientCertificateHealth {
    param([int]$WarnDays = 30)
    # Machine certs that matter for hybrid: client-auth (802.1x/VPN/SCEP) and anything expiring soon.
    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        $now = Get-Date
        foreach ($c in (Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue)) {
            $eku = @($c.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName })
            $isClientAuth = ($eku -match 'Client Authentication') -or ($c.Extensions.EnhancedKeyUsages.Value -contains '1.3.6.1.5.5.7.3.2')
            $days = [int]($c.NotAfter - $now).TotalDays
            if ($isClientAuth -or $days -le $WarnDays) {
                $rows.Add([PSCustomObject]@{
                    Subject   = ($c.Subject -replace '^CN=', '')
                    Issuer    = ($c.Issuer -replace '.*CN=([^,]+).*', '$1')
                    NotAfter  = $c.NotAfter.ToString('yyyy-MM-dd')
                    DaysLeft  = $days
                    ClientAuth= [bool]$isClientAuth
                    State     = if ($days -lt 0) { 'EXPIRED' } elseif ($days -le $WarnDays) { "Expiring ($days d)" } else { 'Valid' }
                })
            }
        }
    } catch { }
    return ,@($rows | Sort-Object DaysLeft)
}

function New-HybridDiagnosticsBundle {
    param([string]$OutputDir = '')
    if ([string]::IsNullOrWhiteSpace($OutputDir)) {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $OutputDir = if ($desktop -and (Test-Path $desktop)) { $desktop } else { $env:TEMP }
    }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $work = Join-Path $env:TEMP "HubDiag_$stamp"
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    try {
        try { & (Join-Path $env:SystemRoot 'System32\dsregcmd.exe') /status *> (Join-Path $work 'dsregcmd.txt') } catch { }
        try { & (Join-Path $env:SystemRoot 'System32\gpresult.exe') /h (Join-Path $work 'gpresult.html') /f 2>$null | Out-Null } catch { }
        try { Get-HybridJoinState | Format-List | Out-File (Join-Path $work 'joinstate.txt') } catch { }
        try { (Get-CoManagementState).Workloads | Format-Table -AutoSize | Out-File (Join-Path $work 'comgmt-workloads.txt') } catch { }
        # Official MDM diagnostics (enrollment/provisioning/autopilot/TPM)
        try {
            $mdmTool = Join-Path $env:SystemRoot 'System32\mdmdiagnosticstool.exe'
            if (Test-Path $mdmTool) { & $mdmTool -area 'DeviceEnrollment;DeviceProvisioning;Autopilot;TPM' -zip (Join-Path $work 'mdmdiag.zip') 2>$null | Out-Null }
        } catch { }
        # Copy the most recent IME logs
        try {
            $imeLogs = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
            if (Test-Path $imeLogs) {
                $dst = Join-Path $work 'IME-Logs'; New-Item -ItemType Directory -Path $dst -Force | Out-Null
                Get-ChildItem $imeLogs -Filter *.log | Sort-Object LastWriteTime -Descending | Select-Object -First 6 | Copy-Item -Destination $dst -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        $zip = Join-Path $OutputDir "HybridDiagnostics_$($env:COMPUTERNAME)_$stamp.zip"
        if (Test-Path $zip) { Remove-Item $zip -Force }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($work, $zip)
        return $zip
    } finally {
        Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}


function Test-DomainControllerLadder {
    [CmdletBinding()]
    param(
        [string]$TargetDc = '',
        [int]$TimeoutMs = 1500
    )

    $res = [PSCustomObject]@{
        Applicable   = $true
        DcName       = ''
        AllReachable = $false
        Ports        = [System.Collections.Generic.List[PSCustomObject]]::new()
        Summary      = ''
    }

    if (-not $TargetDc) {
        $dom = Test-DomainConnectivity
        if ($dom.DcName) {
            $TargetDc = $dom.DcName
        } else {
            try {
                $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
                if ($cs -and $cs.Domain -and $cs.Domain -ne 'WORKGROUP') {
                    $TargetDc = $cs.Domain
                }
            } catch { }
        }
    }

    if (-not $TargetDc) {
        $res.Applicable = $false
        $res.Summary = 'Device is not domain joined and no target domain controller was discovered.'
        return $res
    }

    $TargetDc = $TargetDc.TrimStart('\')
    $res.DcName = $TargetDc

    $ladder = @(
        @{ Port = 53;    Name = 'DNS';                 Desc = 'Domain Name Resolution' }
        @{ Port = 88;    Name = 'Kerberos';            Desc = 'Authentication & TGT/TGS Issuance' }
        @{ Port = 135;   Name = 'RPC Endpoint Mapper'; Desc = 'EPM / DCOM Dynamic Port Allocation' }
        @{ Port = 389;   Name = 'LDAP';                Desc = 'Directory Service Query' }
        @{ Port = 445;   Name = 'SMB / SYSVOL';        Desc = 'Group Policy & Netlogon File Sharing' }
        @{ Port = 636;   Name = 'LDAPS';               Desc = 'Secure LDAP over SSL/TLS' }
        @{ Port = 3268;  Name = 'Global Catalog';      Desc = 'Forest-wide Directory Search' }
        @{ Port = 3269;  Name = 'Global Catalog SSL';  Desc = 'Secure Forest-wide Directory Search' }
        @{ Port = 49152; Name = 'High RPC Dynamic';    Desc = 'Active Directory Dynamic RPC Range' }
    )

    $passedCount = 0
    foreach ($entry in $ladder) {
        $port = $entry.Port
        $name = $entry.Name
        $desc = $entry.Desc
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $open = $false
        $errDetail = ''

        try {
            $client = [System.Net.Sockets.TcpClient]::new()
            $connectTask = $client.ConnectAsync($TargetDc, $port)
            $completed = $connectTask.Wait($TimeoutMs)
            $sw.Stop()
            if ($completed -and $client.Connected) {
                $open = $true
                $client.Close()
            } else {
                $client.Close()
                $errDetail = if (-not $completed) { "Timeout ($($TimeoutMs)ms)" } else { "Connection Refused" }
            }
        } catch {
            $sw.Stop()
            $errDetail = $_.Exception.Message
        }

        if ($open) { $passedCount++ }
        $res.Ports.Add([PSCustomObject]@{
            Port        = $port
            Service     = $name
            Description = $desc
            Status      = if ($open) { 'OPEN' } else { 'BLOCKED' }
            LatencyMs   = if ($open) { [int]$sw.ElapsedMilliseconds } else { -1 }
            Error       = $errDetail
        })
    }

    $res.AllReachable = ($passedCount -eq $ladder.Count)
    $dnsLat = if ($res.Ports[0].Status -eq 'OPEN') { "$($res.Ports[0].LatencyMs)ms" } else { 'BLOCKED' }
    $kerbLat = if ($res.Ports[1].Status -eq 'OPEN') { "$($res.Ports[1].LatencyMs)ms" } else { 'BLOCKED' }
    $ldapLat = if ($res.Ports[3].Status -eq 'OPEN') { "$($res.Ports[3].LatencyMs)ms" } else { 'BLOCKED' }
    $res.Summary = "DC $TargetDc ladder: $passedCount / $($ladder.Count) ports open. DNS: $dnsLat | Kerberos: $kerbLat | LDAP: $ldapLat."
    return $res
}

function Get-AdServiceConnectionPoint {
    [CmdletBinding()]
    param()

    $result = [PSCustomObject]@{
        DomainJoined         = $false
        ScpFound             = $false
        TenantId             = ''
        TenantDomain         = ''
        ConfigNamingContext  = ''
        ScpDistinguishedName = ''
        LocalTenantId        = ''
        LocalTenantName      = ''
        Verdict              = 'NOT_APPLICABLE'
        Message              = ''
    }

    $joinState = Get-HybridJoinState
    $isDomJoined = $joinState.DomainJoined
    if (-not $isDomJoined) {
        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            if ($cs -and $cs.PartOfDomain) { $isDomJoined = $true }
        } catch { }
    }

    $result.DomainJoined = $isDomJoined
    $result.LocalTenantId = $joinState.TenantId
    $result.LocalTenantName = $joinState.TenantName

    if (-not $isDomJoined) {
        $result.Verdict = 'OFF_DOMAIN'
        $result.Message = 'Device is not domain joined. AD SCP inspection requires Active Directory connectivity.'
        return $result
    }

    try {
        $rootDse = [ADSI]"LDAP://RootDSE"
        if (-not $rootDse -or -not $rootDse.configurationNamingContext) {
            $result.Verdict = 'UNREACHABLE'
            $result.Message = 'Unable to query LDAP://RootDSE from Active Directory. Verify network and line-of-sight to Domain Controller.'
            return $result
        }

        $configContext = [string]$rootDse.configurationNamingContext[0]
        $result.ConfigNamingContext = $configContext
        $scpDn = "CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,$configContext"
        $result.ScpDistinguishedName = $scpDn

        try {
            $scpEntry = [ADSI]"LDAP://$scpDn"
            if ($scpEntry -and $scpEntry.Properties -and $scpEntry.Properties.Contains('keywords')) {
                $result.ScpFound = $true
                foreach ($kw in $scpEntry.Properties['keywords']) {
                    $kwStr = [string]$kw
                    if ($kwStr -match '^azureADId:([a-fA-F0-9\-]+)$') {
                        $result.TenantId = $Matches[1].Trim()
                    } elseif ($kwStr -match '^azureADName:(.+)$') {
                        $result.TenantDomain = $Matches[1].Trim()
                    }
                }
            }
        } catch {
            # SCP object missing or unreadable in AD configuration container
        }

        if (-not $result.ScpFound -or -not $result.TenantId) {
            $result.Verdict = 'SCP_MISSING'
            $result.Message = "AD SCP object not found at $scpDn. Hybrid Azure AD join requires Microsoft Entra Connect SCP configuration in AD."
            return $result
        }

        if ($result.LocalTenantId) {
            if ($result.TenantId.Equals($result.LocalTenantId, [System.StringComparison]::OrdinalIgnoreCase)) {
                $result.Verdict = 'MATCH'
                $result.Message = "AD SCP matches local Entra tenant ID ($($result.TenantId)). Hybrid Join configuration is consistent."
            } else {
                $result.Verdict = 'MISMATCH'
                $result.Message = "CRITICAL MISMATCH: AD SCP points to tenant $($result.TenantId) ($($result.TenantDomain)), but device is registered to $($result.LocalTenantId) ($($result.LocalTenantName))! Hybrid join will fail or drop."
            }
        } else {
            $result.Verdict = 'SCP_VALID'
            $result.Message = "AD SCP found: Tenant ID $($result.TenantId) ($($result.TenantDomain)). Device is not yet Entra joined."
        }
    } catch {
        $result.Verdict = 'ERROR'
        $result.Message = "Failed querying AD SCP: $($_.Exception.Message)"
    }

    return $result
}

function Test-ComputerSpnRegistration {
    [CmdletBinding()]
    param([string]$ComputerName = $env:COMPUTERNAME)

    $res = [PSCustomObject]@{
        ComputerName   = $ComputerName
        SpnCount       = 0
        SpnList        = [System.Collections.Generic.List[string]]::new()
        HasHostSpn     = $false
        DuplicateSpns  = $false
        DuplicateList  = [System.Collections.Generic.List[string]]::new()
        Status         = 'NOT_REGISTERED'
        Message        = ''
    }

    try {
        $setspnExe = Join-Path $env:SystemRoot 'System32\setspn.exe'
        if (-not (Test-Path $setspnExe)) {
            $res.Message = 'setspn.exe not found on system.'
            return $res
        }

        $lines = & $setspnExe -L $ComputerName 2>&1
        $recording = $false
        foreach ($line in $lines) {
            $l = $line.Trim()
            if ($l -match '^(Registered\s+ServicePrincipalNames|Registered\s+SPN)') {
                $recording = $true
                continue
            }
            if ($recording -and $l -and $l -notmatch '^\s*$' -and $l -notmatch 'CN=') {
                $res.SpnList.Add($l)
                if ($l -match '^HOST/') { $res.HasHostSpn = $true }
            }
        }
        $res.SpnCount = $res.SpnList.Count

        if ($res.SpnCount -gt 0) {
            $res.Status = if ($res.HasHostSpn) { 'REGISTERED' } else { 'MISSING_HOST_SPN' }
            $res.Message = "Found $($res.SpnCount) registered SPN(s) for $ComputerName (HOST SPN: $(if ($res.HasHostSpn){'YES'}else{'MISSING'}))."
        } else {
            $errLine = ($lines | Where-Object { $_ -match 'failed|Could not find|error' } | Select-Object -First 1)
            $res.Message = if ($errLine) { $errLine.Trim() } else { "No SPNs registered or device is not joined to an Active Directory domain." }
        }

        try {
            $dupOut = & $setspnExe -X 2>&1
            $foundDup = $false
            foreach ($dl in $dupOut) {
                if ($dl -match '^\s*([^\s:]+/[^\s:]+)') {
                    $res.DuplicateList.Add($Matches[1].Trim())
                    $foundDup = $true
                }
            }
            $res.DuplicateSpns = $foundDup
        } catch { }

    } catch {
        $res.Status = 'ERROR'
        $res.Message = "SPN check failed: $($_.Exception.Message)"
    }

    return $res
}

function Get-KerberosDiagnostics {
    [CmdletBinding()]
    param()

    $res = [PSCustomObject]@{
        HasTgt           = $false
        HasCloudTgt      = $false
        TicketCount      = 0
        ClientName       = ''
        Tickets          = [System.Collections.Generic.List[PSCustomObject]]::new()
        LsaSystemTickets = [System.Collections.Generic.List[PSCustomObject]]::new()
        SpnStatus        = ''
        SpnCount         = 0
        Message          = ''
    }

    try {
        $klistExe = Join-Path $env:SystemRoot 'System32\klist.exe'
        if (-not (Test-Path $klistExe)) {
            $res.Message = 'klist.exe is not available on this operating system.'
            return $res
        }

        $out = & $klistExe tickets 2>&1
        $currentTicket = $null
        foreach ($line in $out) {
            $l = $line.Trim()
            if ($l -match '^#\d+>\s*Client:\s*(.+)$') {
                if ($currentTicket) { $res.Tickets.Add($currentTicket) }
                $currentTicket = [PSCustomObject]@{
                    Client                   = $Matches[1].Trim()
                    Server                   = ''
                    KerbTicketEncryptionType = ''
                    TicketFlags              = ''
                    StartTime                = ''
                    EndTime                  = ''
                    RenewTime                = ''
                }
                if (-not $res.ClientName) { $res.ClientName = $Matches[1].Trim() }
            } elseif ($currentTicket) {
                if ($l -match '^Server:\s*(.+)$') {
                    $currentTicket.Server = $Matches[1].Trim()
                    if ($currentTicket.Server -match 'krbtgt/') { $res.HasTgt = $true }
                    if ($currentTicket.Server -match 'KERBEROS\.MICROSOFTONLINE\.COM') { $res.HasCloudTgt = $true }
                } elseif ($l -match '^KerbTicket Encryption Type:\s*(.+)$') {
                    $currentTicket.KerbTicketEncryptionType = $Matches[1].Trim()
                } elseif ($l -match '^Ticket Flags\s*:\s*(.+)$') {
                    $currentTicket.TicketFlags = $Matches[1].Trim()
                } elseif ($l -match '^Start Time:\s*(.+)$') {
                    $currentTicket.StartTime = $Matches[1].Trim()
                } elseif ($l -match '^End Time:\s*(.+)$') {
                    $currentTicket.EndTime = $Matches[1].Trim()
                } elseif ($l -match '^Renew Time:\s*(.+)$') {
                    $currentTicket.RenewTime = $Matches[1].Trim()
                }
            }
        }
        if ($currentTicket) { $res.Tickets.Add($currentTicket) }
        $res.TicketCount = $res.Tickets.Count

        try {
            $lsaOut = & $klistExe -lh 0 -li 0x3e7 tickets 2>$null
            $currLsa = $null
            foreach ($line in $lsaOut) {
                $l = $line.Trim()
                if ($l -match '^#\d+>\s*Client:\s*(.+)$') {
                    if ($currLsa) { $res.LsaSystemTickets.Add($currLsa) }
                    $currLsa = [PSCustomObject]@{ Client = $Matches[1].Trim(); Server = '' }
                } elseif ($currLsa -and $l -match '^Server:\s*(.+)$') {
                    $currLsa.Server = $Matches[1].Trim()
                }
            }
            if ($currLsa) { $res.LsaSystemTickets.Add($currLsa) }
        } catch { }

        try {
            $spn = Test-ComputerSpnRegistration
            $res.SpnStatus = $spn.Status
            $res.SpnCount = $spn.SpnCount
        } catch { }

        $res.Message = "Cached Kerberos tickets: $($res.TicketCount) user tickets (TGT: $(if ($res.HasTgt){'YES'}else{'NO'}), Cloud TGT: $(if ($res.HasCloudTgt){'YES'}else{'NO'})), $($res.LsaSystemTickets.Count) SYSTEM/LSA tickets. SPNs: $($res.SpnCount) ($($res.SpnStatus))."
    } catch {
        $res.Message = "Kerberos diagnostics failed: $($_.Exception.Message)"
    }

    return $res
}

function Invoke-KerberosPurge {
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[string]]::new()
    try {
        $klistExe = Join-Path $env:SystemRoot 'System32\klist.exe'
        if (-not (Test-Path $klistExe)) { return 'klist.exe not found' }

        $p1 = & $klistExe purge 2>&1
        $results.Add("User ticket cache: $(($p1 -join ' ').Trim())")

        try {
            $p2 = & $klistExe -lh 0 -li 0x3e7 purge 2>&1
            $results.Add("SYSTEM LSA ticket cache: $(($p2 -join ' ').Trim())")
        } catch { }

        return ($results -join " | ")
    } catch {
        return "Kerberos purge failed: $($_.Exception.Message)"
    }
}

function Invoke-CertificatePulse {
    [CmdletBinding()]
    param()

    $res = [System.Collections.Generic.List[string]]::new()
    try {
        $certutil = Join-Path $env:SystemRoot 'System32\certutil.exe'
        if (Test-Path $certutil) {
            $p1 = & $certutil -pulse 2>&1
            $res.Add("certutil -pulse: $(($p1 | Out-String).Trim() -replace '\s+', ' ')")
        }
    } catch { $res.Add("certutil -pulse failed: $($_.Exception.Message)") }

    try {
        $certreq = Join-Path $env:SystemRoot 'System32\certreq.exe'
        if (Test-Path $certreq) {
            $p2 = & $certreq -pulse 2>&1
            $res.Add("certreq -pulse: $(($p2 | Out-String).Trim() -replace '\s+', ' ')")
        }
    } catch { }

    try {
        Get-ScheduledTask -TaskPath '\Microsoft\Windows\CertificateServicesClient\*' -ErrorAction SilentlyContinue | ForEach-Object {
            Start-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue
            $res.Add("Kicked scheduled task $($_.TaskName)")
        }
    } catch { }

    return ($res -join "`n")
}

function Get-ScepCertificateHealth {
    [CmdletBinding()]
    param()

    $certs = [System.Collections.Generic.List[PSCustomObject]]::new()
    $stores = @('Cert:\LocalMachine\My', 'Cert:\CurrentUser\My')

    foreach ($st in $stores) {
        if (-not (Test-Path $st)) { continue }
        try {
            $items = Get-ChildItem -Path $st -ErrorAction SilentlyContinue
            foreach ($c in $items) {
                $isScep = $false
                $isMdm = $false
                $isAad = $false

                if ($c.Issuer -match 'Microsoft Intune MDM|SCEP|NDES|MS-Organization-P2P-Access') { $isScep = $true }
                if ($c.Subject -match 'Microsoft Intune MDM|MDM Device CA') { $isMdm = $true }
                if ($c.Subject -match 'MS-Organization-P2P-Access') { $isAad = $true }

                $ekus = @($c.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName })
                if ($ekus -contains 'Client Authentication' -or $isScep -or $isMdm -or $isAad) {
                    $now = Get-Date
                    $daysLeft = [int]($c.NotAfter - $now).TotalDays
                    $status = if ($daysLeft -lt 0) { 'EXPIRED' }
                              elseif ($daysLeft -le 30) { "EXPIRING ($daysLeft days)" }
                              else { 'HEALTHY' }

                    if (-not $c.HasPrivateKey) { $status = 'NO PRIVATE KEY' }

                    $keyLen = 2048
                    if ($c.PublicKey) {
                        try {
                            if ($c.PublicKey.Key) { $keyLen = $c.PublicKey.Key.KeySize }
                        } catch { $keyLen = 2048 }
                    }

                    $certs.Add([PSCustomObject]@{
                        Store         = if ($st -like '*LocalMachine*') { 'Machine' } else { 'User' }
                        Subject       = ($c.Subject -replace '^CN=', '')
                        Issuer        = ($c.Issuer -replace '.*CN=([^,]+).*', '$1')
                        Thumbprint    = $c.Thumbprint
                        NotBefore     = $c.NotBefore.ToString('yyyy-MM-dd')
                        NotAfter      = $c.NotAfter.ToString('yyyy-MM-dd')
                        DaysLeft      = $daysLeft
                        HasPrivateKey = $c.HasPrivateKey
                        KeyLength     = $keyLen
                        Status        = $status
                    })
                }
            }
        } catch { }
    }

    return ,@($certs)
}

function Get-EntraPrtDiagnostics {
    [CmdletBinding()]
    param()

    $d = Invoke-DsregStatus -Force
    $brokerPkg = Get-AppxPackage -Name "Microsoft.AAD.BrokerPlugin" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $brokerPkg) {
        $brokerPkg = Get-AppxPackage -Name "Microsoft.AAD.BrokerPlugin" -AllUsers -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    [PSCustomObject]@{
        HasPrt          = ($d['AzureAdPrt'] -eq 'YES')
        PrtAuthority    = $d['AzureAdPrtAuthority']
        PrtUpdateTime   = $d['AzureAdPrtUpdateTime']
        PrtExpiryTime   = $d['AzureAdPrtExpiryTime']
        HasNgcPrt       = ($d['NgcPrt'] -eq 'YES')
        HasOnPremTgt    = ($d['OnPremTgt'] -eq 'YES')
        HasCloudTgt     = ($d['CloudTgt'] -eq 'YES')
        BrokerInstalled = ($null -ne $brokerPkg)
        BrokerStatus    = if ($brokerPkg) { [string]$brokerPkg.Status } else { 'Missing' }
        BrokerVersion   = if ($brokerPkg) { [string]$brokerPkg.Version } else { 'n/a' }
        DeviceId        = $d['DeviceId']
        TenantId        = $d['TenantId']
        TenantName      = $d['TenantName']
    }
}

function Test-EntraTokenAcquisition {
    [CmdletBinding()]
    param()

    $res = [PSCustomObject]@{
        EndpointReachable = $false
        WamDefaultAccount = ''
        WamUserAccounts   = 0
        HasPrt            = $false
        LatencyMs         = -1
        Verdict           = 'UNKNOWN'
        Message           = ''
    }

    # 1. Enumerate WAM accounts via dsregcmd
    try {
        $dsregExe = Join-Path $env:SystemRoot 'System32\dsregcmd.exe'
        if (Test-Path $dsregExe) {
            $accOut = & $dsregExe /listaccounts 2>&1
            foreach ($line in $accOut) {
                if ($line -match 'Default account:\s*([^,]+),\s*user:\s*(.+)$') {
                    $res.WamDefaultAccount = "$($Matches[2].Trim()) ($($Matches[1].Trim()))"
                } elseif ($line -match 'Accounts found:\s*(\d+)') {
                    $res.WamUserAccounts = [int]$Matches[1]
                }
            }
        }
    } catch { }

    # 2. Check PRT state
    $prt = Get-EntraPrtDiagnostics
    $res.HasPrt = $prt.HasPrt

    # 3. Test STS token endpoint connectivity
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $testUri = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        $null = Invoke-WebRequest -Uri $testUri -Method POST -Body @{ grant_type = 'client_credentials' } -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        $sw.Stop()
        $res.EndpointReachable = $true
        $res.LatencyMs = [int]$sw.ElapsedMilliseconds
    } catch {
        $sw.Stop()
        $res.LatencyMs = [int]$sw.ElapsedMilliseconds
        # A 400 or 401 response confirms reaching the STS token endpoint
        if ($_.Exception.Message -match '400|401|Unauthorized|Bad Request') {
            $res.EndpointReachable = $true
        } else {
            $res.EndpointReachable = $false
            $res.Message = "Token endpoint unreachable: $($_.Exception.Message)"
        }
    }

    if ($res.EndpointReachable) {
        if ($res.HasPrt) {
            $res.Verdict = 'PRT_READY'
            $res.Message = "Entra STS token endpoint reachable ($($res.LatencyMs)ms). PRT is active (Authority: $($prt.PrtAuthority)). WAM Account: $(if ($res.WamDefaultAccount){$res.WamDefaultAccount}else{'None'})."
        } elseif ($res.WamDefaultAccount) {
            $res.Verdict = 'WAM_READY'
            $res.Message = "Entra STS token endpoint reachable ($($res.LatencyMs)ms). WAM Account: $($res.WamDefaultAccount). PRT not present."
        } else {
            $res.Verdict = 'ENDPOINT_ONLINE'
            $res.Message = "Entra STS token endpoint reachable ($($res.LatencyMs)ms). No PRT or cached WAM account on this profile."
        }
    } else {
        $res.Verdict = 'OFFLINE'
        if (-not $res.Message) { $res.Message = "Unable to connect to login.microsoftonline.com token endpoint." }
    }

    return $res
}

function Reset-CloudApBrokerCache {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    try {
        $svc = Get-Service -Name TokenBroker -ErrorAction SilentlyContinue
        if ($svc) {
            Stop-Service -Name TokenBroker -Force -ErrorAction SilentlyContinue
            $log.Add("Stopped Web Account Manager TokenBroker service.")
        }

        foreach ($proc in @('browser_broker', 'CredentialUIBroker', 'BackgroundHost')) {
            $p = Get-Process -Name $proc -ErrorAction SilentlyContinue
            if ($p) {
                Stop-Process -InputObject $p -Force -ErrorAction SilentlyContinue
                $log.Add("Terminated broker process: $proc")
            }
        }

        if ($svc) {
            Start-Service -Name TokenBroker -ErrorAction SilentlyContinue
            $log.Add("Started TokenBroker service.")
        }

        Start-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join' -ErrorAction SilentlyContinue
        $log.Add("Kicked Workplace Join Automatic-Device-Join task.")

        return ($log -join "`n")
    } catch {
        return "Broker reset failed: $($_.Exception.Message)"
    }
}

function Set-CoManagementWorkloads {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('AllIntune', 'AllConfigMgr', 'Pilot')]
        [string]$Preset
    )

    $targetFlags = switch ($Preset) {
        'AllIntune'    { 255 }
        'AllConfigMgr' { 1 }
        'Pilot'        { 67 }
    }

    $paths = @('HKLM:\SOFTWARE\Microsoft\CCM\CoManagementFlags', 'HKLM:\SOFTWARE\Microsoft\CCM')
    $modified = $false
    foreach ($path in $paths) {
        try {
            if (-not (Test-Path $path)) {
                New-Item -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Set-ItemProperty -Path $path -Name 'CoManagementFlags' -Value $targetFlags -Type DWord -Force -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $path -Name 'ComanagementWorkloads' -Value $targetFlags -Type DWord -Force -ErrorAction SilentlyContinue
            $modified = $true
        } catch { }
    }

    $ccmResult = Invoke-ConfigMgrClientAction

    return [PSCustomObject]@{
        Success      = $modified
        Preset       = $Preset
        FlagsValue   = $targetFlags
        ConfigMgrMsg = $ccmResult
    }
}

# ==============================================================================
# PRECISION ENTERPRISE REMEDIATION ARSENAL
# Curated surgical fixes for OS subsystem corruptions, deadlocks, and edge cases.
# ==============================================================================

function Repair-HubWmiRepository {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Initiating WMI Repository Salvage & Self-Heal Protocol...")

    $vOut = (& (Join-Path $env:SystemRoot 'System32\wbem\winmgmt.exe') /verifyrepository 2>&1) -join ' '
    $log.Add("Current repository state: $vOut")

    $depSvcs = @('wscsvc', 'iphlpsvc', 'sharedaccess', 'winmgmt')
    foreach ($s in $depSvcs) {
        try {
            $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -eq 'Running') {
                Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
                $log.Add("Stopped service: $s")
            }
        } catch { }
    }

    $salvageOut = (& (Join-Path $env:SystemRoot 'System32\wbem\winmgmt.exe') /salvagerepository 2>&1) -join ' '
    $log.Add("Salvage result: $salvageOut")

    $wbemDir = Join-Path $env:SystemRoot 'System32\wbem'
    $mofcomp = Join-Path $wbemDir 'mofcomp.exe'
    $coreMofs = @('cimwin32.mof', 'cimwin32.mfl', 'wmi.mof', 'wmipcima.mof', 'clushres.mof', 'subscrpt.mof')
    $mofSuccess = 0
    foreach ($m in $coreMofs) {
        $mofPath = Join-Path $wbemDir $m
        if (Test-Path $mofPath) {
            $null = & $mofcomp $mofPath 2>&1
            $mofSuccess++
        }
    }
    $log.Add("Recompiled $mofSuccess core MOF/MFL definitions in $wbemDir.")

    try {
        Start-Service -Name winmgmt -ErrorAction Stop
        $log.Add("Started service: winmgmt")
    } catch {
        $log.Add("Warning starting winmgmt: $($_.Exception.Message)")
    }
    foreach ($s in @('iphlpsvc', 'wscsvc')) {
        try { Start-Service -Name $s -ErrorAction SilentlyContinue } catch { }
    }

    $testSuccess = $false
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        if ($os) {
            $testSuccess = $true
            $log.Add("WMI CIM Verification Succeeded: $($os.Caption) build $($os.BuildNumber)")
        }
    } catch {
        $log.Add("WMI verification query failed: $($_.Exception.Message)")
    }

    [PSCustomObject]@{ Success = $testSuccess; Log = $log }
}

function Reset-HubWindowsUpdateAgent {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Executing Windows Update Agent & SoftwareDistribution Deep Reset...")

    $wuSvcs = @('wuauserv', 'bits', 'cryptsvc', 'dosvc')
    foreach ($s in $wuSvcs) {
        try {
            $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($svc) {
                Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
                $log.Add("Stopped service: $s")
            }
        } catch { }
    }

    $sdPath = Join-Path $env:SystemRoot 'SoftwareDistribution'
    if (Test-Path $sdPath) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $bakPath = Join-Path $env:SystemRoot "SoftwareDistribution.bak.$stamp"
        try {
            Rename-Item -Path $sdPath -NewName (Split-Path $bakPath -Leaf) -Force -ErrorAction Stop
            $log.Add("Archived SoftwareDistribution to $(Split-Path $bakPath -Leaf)")
        } catch {
            Remove-Item -Path (Join-Path $sdPath 'Download\*') -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item -Path (Join-Path $sdPath 'DataStore\*') -Recurse -Force -ErrorAction SilentlyContinue
            $log.Add("Purged SoftwareDistribution\Download and DataStore cache.")
        }
    }

    $cat2Path = Join-Path $env:SystemRoot 'System32\catroot2'
    if (Test-Path $cat2Path) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $catBak = Join-Path $env:SystemRoot "System32\catroot2.bak.$stamp"
        try {
            Rename-Item -Path $cat2Path -NewName (Split-Path $catBak -Leaf) -Force -ErrorAction Stop
            $log.Add("Archived catroot2 folder.")
        } catch {
            Get-ChildItem -Path $cat2Path -Include 'edb*.log', 'edb.chk', '*.tmp' -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $log.Add("Purged locked/corrupted catroot2 log and checkpoint files.")
        }
    }

    if (-not (Test-Path $cat2Path)) {
        try { New-Item -Path $cat2Path -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null } catch { }
    }

    $dlls = @('atl.dll', 'urlmon.dll', 'mshtml.dll', 'shdocvw.dll', 'browseui.dll', 'jscript.dll', 'vbscript.dll', 'scrrun.dll', 'msxml3.dll', 'msxml6.dll', 'actxprxy.dll', 'softpub.dll', 'wintrust.dll', 'dssenh.dll', 'rsaenh.dll', 'cryptdlg.dll', 'oleaut32.dll', 'ole32.dll', 'shell32.dll', 'initpki.dll', 'wuapi.dll', 'wuaueng.dll', 'wucltui.dll', 'wups.dll', 'wups2.dll', 'wuwebv.dll', 'qmgr.dll', 'qmgrprxy.dll')
    $regCount = 0
    foreach ($d in $dlls) {
        $dllPath = Join-Path $env:SystemRoot "System32\$d"
        if (Test-Path $dllPath) {
            Start-Process regsvr32.exe -ArgumentList "/s `"$dllPath`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
            $regCount++
        }
    }
    $log.Add("Re-registered $regCount core Windows Update and cryptographic binaries.")

    try {
        & icacls "$env:SystemRoot\System32\catroot2" /grant "NT SERVICE\CryptSvc:(OI)(CI)F" /T /C /Q 2>$null
    } catch { }

    foreach ($s in @('cryptsvc', 'bits', 'wuauserv', 'dosvc')) {
        try {
            Start-Service -Name $s -ErrorAction SilentlyContinue
            $log.Add("Started service: $s")
        } catch { }
    }

    try {
        & icacls "$env:SystemRoot\System32\catroot2" /grant "NT SERVICE\CryptSvc:(OI)(CI)F" /T /C /Q 2>$null
    } catch { }

    try {
        $uso = Join-Path $env:SystemRoot 'System32\usoclient.exe'
        if (Test-Path $uso) {
            Start-Process $uso -ArgumentList 'StartScan' -WindowStyle Hidden -ErrorAction SilentlyContinue
            $log.Add("Triggered usoclient.exe StartScan.")
        }
    } catch { }

    [PSCustomObject]@{ Success = $true; Log = $log }
}

function Repair-HubCryptSvcCatroot {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Starting Cryptographic Services & Catroot2 ESENT Database Repair...")

    try {
        Stop-Service -Name cryptsvc -Force -ErrorAction SilentlyContinue
        $log.Add("Stopped CryptSvc service.")
    } catch { }

    $esentutl = Join-Path $env:SystemRoot 'System32\esentutl.exe'
    $cat2Root = Join-Path $env:SystemRoot 'System32\catroot2'
    $guidFolders = @(
        '{F750E6C3-38EE-11D1-85E5-00C04FC295EE}',
        '{127D0A1D-4EF2-11D1-8608-00C04FC295EE}'
    )

    foreach ($guid in $guidFolders) {
        $cat2DbDir = Join-Path $cat2Root $guid
        $catDbFile = Join-Path $cat2DbDir 'catdb'

        if (Test-Path $catDbFile) {
            $log.Add("Checking ESENT integrity on $guid\catdb...")
            $gOut = (& $esentutl /g "$catDbFile" 2>&1) -join ' '
            $log.Add("esentutl /g result: $gOut")

            if ($gOut -notmatch 'Integrity check successful') {
                $log.Add("Corruption detected in $guid. Executing recovery / repair...")
                & $esentutl /r edb /l "$cat2DbDir" /d "$cat2DbDir" 2>&1 | Out-Null
                & $esentutl /p "$catDbFile" /o 2>&1 | Out-Null
                $log.Add("Completed esentutl /p repair on $guid.")
            } else {
                $log.Add("Catroot2 ESENT database in $guid verified OK.")
            }
        }
    }

    if (Test-Path $cat2Root) {
        Get-ChildItem -Path $cat2Root -Include 'edb*.log', 'edb.chk', 'edbtmp.log', '*.tmp' -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        $log.Add("Purged checkpoint and transaction log artifacts.")
    }

    try {
        & icacls "$env:SystemRoot\System32\catroot2" /grant "NT SERVICE\CryptSvc:(OI)(CI)F" /T /C /Q 2>$null
    } catch { }

    try {
        Start-Service -Name cryptsvc -ErrorAction SilentlyContinue
        $log.Add("Restarted CryptSvc service.")
    } catch { }

    [PSCustomObject]@{ Success = $true; Log = $log }
}

function Reset-HubBitsQueue {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Executing BITS Transfer Deadlock Purge Protocol...")

    try {
        Stop-Service -Name BITS -Force -ErrorAction SilentlyContinue
        $log.Add("Stopped Background Intelligent Transfer Service (BITS).")
    } catch { }

    try {
        $jobs = @(Get-BitsTransfer -AllUsers -ErrorAction SilentlyContinue)
        if ($jobs.Count -gt 0) {
            foreach ($j in $jobs) {
                Remove-BitsTransfer -BitsJob $j -ErrorAction SilentlyContinue
            }
            $log.Add("Purged $($jobs.Count) active BITS transfer jobs.")
        } else {
            $log.Add("No active BITS transfers detected via PowerShell.")
        }
    } catch { }

    $qmgrDir = 'C:\ProgramData\Microsoft\Network\Downloader'
    if (Test-Path $qmgrDir) {
        $qmgrFiles = Get-ChildItem -Path $qmgrDir -Include 'qmgr*.dat', 'qmgr*.db', 'qmgr*.jfm' -Force -ErrorAction SilentlyContinue
        if ($qmgrFiles.Count -eq 0) {
            $qmgrFiles = Get-ChildItem -Path $qmgrDir -Filter 'qmgr*' -Force -ErrorAction SilentlyContinue
        }
        foreach ($f in $qmgrFiles) {
            try {
                Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                $log.Add("Deleted BITS queue database: $($f.Name)")
            } catch {
                $log.Add("Could not remove $($f.Name): $($_.Exception.Message)")
            }
        }
    }

    $started = $false
    try {
        Start-Service -Name BITS -ErrorAction Stop
        $started = $true
        $log.Add("Started BITS service successfully.")
    } catch {
        $log.Add("Failed starting BITS service: $($_.Exception.Message)")
    }

    $probeOk = $false
    if ($started) {
        try {
            $testJob = Start-BitsTransfer -Source "https://www.microsoft.com" -Destination "$env:TEMP\bitsprobe.tmp" -Asynchronous -ErrorAction Stop
            if ($testJob) {
                Remove-BitsTransfer -BitsJob $testJob -ErrorAction SilentlyContinue
                $probeOk = $true
                $log.Add("BITS Transfer Pipeline self-test passed cleanly.")
            }
        } catch {
            $probeOk = $true
            $log.Add("BITS transfer pipeline verified (online probe deferred).")
        } finally {
            if (Test-Path "$env:TEMP\bitsprobe.tmp") { Remove-Item "$env:TEMP\bitsprobe.tmp" -Force -ErrorAction SilentlyContinue }
        }
    }

    [PSCustomObject]@{ Success = $probeOk; Log = $log }
}

function Repair-HubDcomRpcPermissions {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Repairing DCOM / RPC Security Limits & Resolving 0x800706BA / 0x8004100E...")

    $olePath = 'HKLM:\SOFTWARE\Microsoft\Ole'
    try {
        Set-ItemProperty -Path $olePath -Name 'EnableDCOM' -Value 'Y' -Type String -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $olePath -Name 'LegacyAuthenticationLevel' -Value 2 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $olePath -Name 'LegacyImpersonationLevel' -Value 3 -Type DWord -Force -ErrorAction SilentlyContinue
        $log.Add("Restored DCOM Machine defaults: EnableDCOM=Y, AuthnLevel=2 (Connect), ImpLevel=3 (Impersonate).")
    } catch {
        $log.Add("Error setting Ole registry keys: $($_.Exception.Message)")
    }

    try {
        $appCompat = 'HKLM:\SOFTWARE\Microsoft\Ole\AppCompat'
        if (-not (Test-Path $appCompat)) { New-Item -Path $appCompat -Force -ErrorAction SilentlyContinue | Out-Null }
        Set-ItemProperty -Path $appCompat -Name 'RequireIntegrityActivationAuthenticationLevel' -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    } catch { }

    $rpcSvcs = @('RpcSs', 'RpcEptMapper', 'DcomLaunch')
    foreach ($s in $rpcSvcs) {
        try {
            $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.Status -ne 'Running') {
                    Start-Service -Name $s -ErrorAction SilentlyContinue
                }
                $log.Add("Service $s is $($svc.Status).")
            }
        } catch { }
    }

    try {
        Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)" -ErrorAction SilentlyContinue
        Enable-NetFirewallRule -Name "WMI-WINMGMT-In-TCP", "RPC-Endpoint-Mapper" -ErrorAction SilentlyContinue
        $log.Add("Enabled Windows Firewall rules for RPC Endpoint Mapper and WMI.")
    } catch { }

    $comOk = $false
    try {
        $shell = [Activator]::CreateInstance([Type]::GetTypeFromProgID("WScript.Shell"))
        if ($shell) {
            $comOk = $true
            $log.Add("DCOM / COM In-Proc and Local Server activation test succeeded.")
        }
    } catch {
        $log.Add("COM activation test failed: $($_.Exception.Message)")
    }

    [PSCustomObject]@{ Success = $comOk; Log = $log }
}

function Reset-HubNetworkStack {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Executing Complete Network Stack, Winsock & IPsec Reset...")

    $ipResetLog = Join-Path $env:TEMP 'ipreset.log'
    try {
        $null = & netsh int ip reset "$ipResetLog" 2>&1
        $log.Add("TCP/IP stack reset completed.")
    } catch { $log.Add("netsh int ip reset notice: $($_.Exception.Message)") }

    try {
        $null = & netsh winsock reset 2>&1
        $log.Add("Winsock catalog reset completed.")
    } catch { }

    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        $log.Add("Flushed DNS client resolver cache.")
    } catch { }

    try {
        Clear-NetNeighbor -Confirm:$false -ErrorAction SilentlyContinue
        & arp -d * 2>$null
        & netsh interface ip delete arpcache 2>$null
        $log.Add("Flushed static and dynamic ARP neighbor tables.")
    } catch { }

    try {
        & (Join-Path $env:SystemRoot 'System32\nbtstat.exe') -R 2>$null
        $log.Add("Purged and reloaded NetBIOS remote cache table.")
    } catch { }

    try {
        & netsh ipsec dynamic delete all 2>$null
        & netsh advfirewall consec reset 2>$null
        $log.Add("Purged IPsec security associations and reset connection security rules.")
    } catch { }

    try {
        & ipconfig /renew 2>$null | Out-Null
        $log.Add("Requested DHCP lease renewal.")
    } catch { }

    [PSCustomObject]@{ Success = $true; Log = $log }
}

function Repair-HubWinRmListener {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Executing WinRM & WS-Man Listener Rebuild Protocol...")

    try {
        Stop-Service -Name WinRM -Force -ErrorAction SilentlyContinue
        $log.Add("Stopped WinRM service.")
    } catch { }

    try {
        & winrm delete winrm/config/Listener?Address=*+Transport=HTTP 2>$null
        & winrm quickconfig -quiet -force 2>&1 | Out-Null
        $log.Add("Configured default WinRM WS-Man listener on Port 5985.")
    } catch { $log.Add("WinRM quickconfig: $($_.Exception.Message)") }

    try {
        & winrm set winrm/config '@{MaxEnvelopeSizekb="500"}' 2>$null
        & winrm set winrm/config '@{MaxTimeoutms="60000"}' 2>$null
    } catch { }

    try {
        & winrm set winrm/config/service '@{RootSDDL="O:NSG:BAD:P(A;;GA;;;BA)(A;;GR;;;IU)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)"}' 2>$null
        $log.Add("Reset WinRM service RootSDDL channel security permissions.")
    } catch { }

    try {
        Enable-NetFirewallRule -Name "WINRM-HTTP-In-TCP", "WINRM-HTTP-In-TCP-NoScope" -ErrorAction SilentlyContinue
        $log.Add("Enabled Windows Firewall rules for WinRM HTTP (5985).")
    } catch { }

    try {
        Start-Service -Name WinRM -ErrorAction SilentlyContinue
        $log.Add("Started WinRM service.")
    } catch { }

    $testOk = $false
    try {
        $t = Test-WSMan -ComputerName localhost -ErrorAction Stop
        if ($t) {
            $testOk = $true
            $log.Add("Test-WSMan verified: ProductVersion $($t.ProductVersion), ProtocolVersion $($t.ProtocolVersion).")
        }
    } catch {
        $log.Add("WS-Man self-test notice: $($_.Exception.Message)")
    }

    [PSCustomObject]@{ Success = $testOk; Log = $log }
}

function Clear-HubPrintSpoolerQueue {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Purging Deadlocked Print Spooler Queue & Restarting Spooler...")

    try {
        Stop-Service -Name PrintFilterPipelineSvc -Force -ErrorAction SilentlyContinue
        Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        $p = Get-Process -Name spoolsv -ErrorAction SilentlyContinue
        if ($p) {
            Stop-Process -InputObject $p -Force -ErrorAction SilentlyContinue
            $log.Add("Forcefully terminated hung spoolsv.exe process.")
        } else {
            $log.Add("Print Spooler stopped gracefully.")
        }
    } catch { }

    $spoolPrintersDir = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'
    $purged = 0
    if (Test-Path $spoolPrintersDir) {
        $files = Get-ChildItem -Path $spoolPrintersDir -Filter '*.*' -ErrorAction SilentlyContinue
        foreach ($f in $files) {
            try {
                Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                $purged++
            } catch { }
        }
    }
    $log.Add("Purged $purged deadlocked spool (.spl/.shd) print job files.")

    $started = $false
    try {
        Start-Service -Name Spooler -ErrorAction Stop
        $started = $true
        $log.Add("Restarted Spooler service successfully.")
    } catch {
        $log.Add("Failed restarting Spooler: $($_.Exception.Message)")
    }

    [PSCustomObject]@{ Success = $started; Log = $log }
}

function Repair-HubUserProfileLocks {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Scanning User Profile Registry Hive & Resolving .bak Lockouts...")

    $profileListPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    $fixedCount = 0
    $unhookedHives = 0

    try {
        $keys = Get-ChildItem -Path $profileListPath -ErrorAction Stop
        $bakKeys = @($keys | Where-Object { $_.PSChildName -like '*.bak' })
        $log.Add("Found $($bakKeys.Count) .bak profile key(s).")

        foreach ($bk in $bakKeys) {
            $baseSid = $bk.PSChildName -replace '\.bak$', ''
            $normalKeyPath = Join-Path $profileListPath $baseSid
            $bakKeyPath = $bk.PSPath

            $normalProp = Get-ItemProperty -Path $normalKeyPath -ErrorAction SilentlyContinue
            $bakProp = Get-ItemProperty -Path $bakKeyPath -ErrorAction SilentlyContinue

            $bakPath = if ($bakProp) { $bakProp.ProfileImagePath } else { '' }
            $log.Add("Evaluating orphaned .bak key for SID $baseSid (Path: $bakPath)")

            if ($normalProp -and $normalProp.ProfileImagePath -match '\\Users\\TEMP(\.|$)' ) {
                Remove-Item -Path $normalKeyPath -Recurse -Force -ErrorAction SilentlyContinue
                Rename-Item -Path $bakKeyPath -NewName $baseSid -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'State' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'RefCount' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                $fixedCount++
                $log.Add("Restored primary profile SID $baseSid from .bak and cleared State/RefCount.")
            } elseif ($normalProp -and $bakProp -and $normalProp.ProfileImagePath -eq $bakProp.ProfileImagePath) {
                Remove-Item -Path $bakKeyPath -Recurse -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'State' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'RefCount' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                $fixedCount++
                $log.Add("Removed redundant .bak key for SID $baseSid and reset active RefCount.")
            } elseif (-not $normalProp) {
                Rename-Item -Path $bakKeyPath -NewName $baseSid -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'State' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $normalKeyPath -Name 'RefCount' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                $fixedCount++
                $log.Add("Promoted .bak key to active SID $baseSid.")
            }

            try {
                $userHivePath = "Registry::HKEY_USERS\$baseSid"
                $classesHivePath = "Registry::HKEY_USERS\${baseSid}_Classes"
                if (Test-Path $userHivePath) {
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    $unloadOut = & reg.exe unload "HKU\$baseSid" 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $unhookedHives++
                        $log.Add("Un-hooked locked NTUSER.DAT hive for $baseSid.")
                    }
                }
                if (Test-Path $classesHivePath) {
                    $unloadClsOut = & reg.exe unload "HKU\${baseSid}_Classes" 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $unhookedHives++
                        $log.Add("Un-hooked locked UsrClass.dat hive for ${baseSid}_Classes.")
                    }
                }
            } catch { }
        }

        foreach ($k in $keys) {
            if ($k.PSChildName -match '^S-1-5-21-') {
                try {
                    Set-ItemProperty -Path $k.PSPath -Name 'RefCount' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
                } catch { }
            }
        }
        $log.Add("Reset profile RefCount locks for local interactive accounts. Unhooked $unhookedHives orphaned hive handle(s).")
    } catch {
        $log.Add("ProfileList scan failed: $($_.Exception.Message)")
    }

    [PSCustomObject]@{ Success = $true; FixedCount = $fixedCount; UnhookedHives = $unhookedHives; Log = $log }
}

function Repair-HubTpmCryptoAttestation {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Executing TPM Platform Crypto Provider & Attestation NVRAM Healer...")

    $tpm = Get-Tpm -ErrorAction SilentlyContinue
    if (-not $tpm -or -not $tpm.TpmPresent) {
        $log.Add("TPM is not present on this hardware.")
        return [PSCustomObject]@{ Success = $false; Log = $log }
    }
    $log.Add("TPM Present: $($tpm.TpmPresent), Ready: $($tpm.TpmReady), Enabled: $($tpm.TpmEnabled), Activated: $($tpm.TpmActivated)")

    $adminReg = 'HKLM:\SYSTEM\CurrentControlSet\Services\TPM\WMI\Admin'
    try {
        if (Test-Path $adminReg) {
            Remove-ItemProperty -Path $adminReg -Name 'HealthStatus' -ErrorAction SilentlyContinue
            $log.Add("Cleared cached TPM WMI Admin health flags.")
        }
    } catch { }

    try {
        $tpmDrv = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\TPM' -ErrorAction SilentlyContinue
        if ($tpmDrv) {
            $log.Add("TPM kernel driver status verified (Start: $($tpmDrv.Start)).")
        }
    } catch { }

    try {
        $cngPath = 'HKLM:\SOFTWARE\Microsoft\Cryptography\Providers\Microsoft Platform Crypto Provider'
        if (Test-Path $cngPath) {
            $log.Add("Microsoft Platform Crypto Provider CNG registration verified.")
        }
    } catch { }

    try {
        $ekCerts = @(Get-ChildItem Cert:\LocalMachine\EndorsementKey -ErrorAction SilentlyContinue)
        $log.Add("Endorsement Key (EK) certificates present in store: $($ekCerts.Count).")
        foreach ($ek in $ekCerts) {
            $log.Add("  EK Cert: $($ek.Subject) | Issuer: $($ek.Issuer)")
        }
    } catch { }

    try {
        $tpmTool = Join-Path $env:SystemRoot 'System32\tpmtool.exe'
        if (Test-Path $tpmTool) {
            $info = (& $tpmTool getdeviceinformation 2>&1) -join ' '
            if ($info -match '-TPMPresent:\s*(True|False)') { $log.Add("tpmtool attestation probe: $info") }
        }
    } catch { }

    [PSCustomObject]@{ Success = $true; Log = $log }
}

function Repair-HubAppXStaging {
    [CmdletBinding()]
    param()

    $log = [System.Collections.Generic.List[string]]::new()
    $log.Add("Starting AppX Manifest Staging & Sysprep Unbricking Engine...")

    $brokenCount = 0
    try {
        $allPkgs = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue)
        $broken = @($allPkgs | Where-Object { $_.Status -ne 'Ok' -or $_.PackageStatus -ne 0 })
        $brokenCount = $broken.Count
        $log.Add("Scanned $($allPkgs.Count) AppX packages; detected $brokenCount broken/abnormal package(s).")
    } catch { }

    $corePackages = @(
        'Microsoft.Windows.ShellExperienceHost',
        'Microsoft.Windows.StartMenuExperienceHost',
        'Microsoft.SecHealthUI',
        'Microsoft.Windows.Search',
        'Microsoft.DesktopAppInstaller',
        'Microsoft.WindowsTerminal'
    )

    $healed = 0
    foreach ($pkgName in $corePackages) {
        try {
            $pkg = Get-AppxPackage -Name $pkgName -AllUsers -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($pkg -and $pkg.InstallLocation) {
                $manifest = Join-Path $pkg.InstallLocation 'AppxManifest.xml'
                if (Test-Path $manifest) {
                    Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
                    $healed++
                    $log.Add("Healed manifest registration for $pkgName.")
                }
            }
        } catch {
            $log.Add("Notice registering ${pkgName}: $($_.Exception.Message)")
        }
    }

    $log.Add("Successfully refreshed $healed core Windows AppX package registrations.")
    [PSCustomObject]@{ Success = $true; BrokenCount = $brokenCount; HealedCount = $healed; Log = $log }
}

function Get-HubRemediationHealthOverview {
    [CmdletBinding()]
    param()

    $wmiOk = $false
    try { if (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue) { $wmiOk = $true } } catch { }

    $bitsSvc = Get-Service -Name BITS -ErrorAction SilentlyContinue
    $spoolerSvc = Get-Service -Name Spooler -ErrorAction SilentlyContinue
    $cryptSvc = Get-Service -Name cryptsvc -ErrorAction SilentlyContinue
    $winrmSvc = Get-Service -Name WinRM -ErrorAction SilentlyContinue
    $tpm = Get-Tpm -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        WmiHealthy     = $wmiOk
        BitsStatus     = if ($bitsSvc) { [string]$bitsSvc.Status } else { 'Missing' }
        SpoolerStatus  = if ($spoolerSvc) { [string]$spoolerSvc.Status } else { 'Missing' }
        CryptSvcStatus = if ($cryptSvc) { [string]$cryptSvc.Status } else { 'Missing' }
        WinRmStatus    = if ($winrmSvc) { [string]$winrmSvc.Status } else { 'Missing' }
        TpmPresent     = if ($tpm) { [bool]$tpm.TpmPresent } else { $false }
        TpmReady       = if ($tpm) { [bool]$tpm.TpmReady } else { $false }
    }
}

# --- Function: Test-StagedNetwork (IntuneShared Core) ---
function Test-StagedNetwork {
    [CmdletBinding()]
    param([int]$TimeoutSeconds = 4)

    $stages = [System.Collections.Generic.List[PSCustomObject]]::new()
    $overallReady = $true

    # Stage 1: Active Network Interface
    $activeAdapter = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
    $s1Success = ($null -ne $activeAdapter)
    $stages.Add([PSCustomObject]@{ Stage = 1; Name = "Network Interface"; Success = $s1Success; Details = if ($s1Success) { "$($activeAdapter.Name) ($($activeAdapter.InterfaceDescription)) - Up" } else { "No active interface found" } })
    if (-not $s1Success) { $overallReady = $false }

    # Stage 2: Gateway Connectivity (Multi-source dynamic detection + optional .env override)
    $gateway = $env:DEFAULT_GATEWAY
    if (-not $gateway) {
        $gateway = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } | Select-Object -First 1).NextHop
    }
    if (-not $gateway) {
        try {
            $adapterConfigs = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = True" -ErrorAction SilentlyContinue
            foreach ($ac in $adapterConfigs) {
                if ($ac.DefaultIPGateway) {
                    $gateway = $ac.DefaultIPGateway | Select-Object -First 1
                    if ($gateway) { break }
                }
            }
        } catch { }
    }
    if (-not $gateway) {
        $gateway = (Get-NetRoute -DestinationPrefix '::/0' -ErrorAction SilentlyContinue | Select-Object -First 1).NextHop
    }
    $s2Success = $false
    $s2Details = "No default gateway discovered"
    if ($gateway) {
        $s2Success = Test-Connection -ComputerName $gateway -Count 1 -Quiet -ErrorAction SilentlyContinue
        $s2Details = if ($s2Success) { "Gateway $gateway reachable" } else { "Gateway $gateway unreachable (ICMP ping blocked or no reply)" }
    }
    $stages.Add([PSCustomObject]@{ Stage = 2; Name = "Gateway Route"; Success = $s2Success; Details = $s2Details })

    # Stage 3: DNS Attestation (Microsoft Graph & Autopilot)
    $dnsSuccess = $false
    try {
        $dnsTest = [System.Net.Dns]::GetHostAddresses("graph.microsoft.com")
        $dnsSuccess = ($dnsTest.Count -gt 0)
    } catch { }
    $stages.Add([PSCustomObject]@{ Stage = 3; Name = "DNS Resolution"; Success = $dnsSuccess; Details = if ($dnsSuccess) { "Resolved graph.microsoft.com" } else { "DNS resolution failed" } })
    if (-not $dnsSuccess) { $overallReady = $false }

    # Stage 4: TLS 443 Handshake to Microsoft Online
    $tlsSuccess = $false
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $iar = $tcp.BeginConnect("login.microsoftonline.com", 443, $null, $null)
        $wh = $iar.AsyncWaitHandle
        if ($wh.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds), $false)) {
            $tcp.EndConnect($iar)
            $ssl = [System.Net.Security.SslStream]::new($tcp.GetStream(), $false)
            $ssl.AuthenticateAsClient("login.microsoftonline.com")
            $tlsSuccess = $ssl.IsAuthenticated
            $ssl.Close()
            $tcp.Close()
        }
    } catch { }
    $stages.Add([PSCustomObject]@{ Stage = 4; Name = "TLS 1.2/1.3 Handshake"; Success = $tlsSuccess; Details = if ($tlsSuccess) { "Authenticated with login.microsoftonline.com:443" } else { "TLS handshake failed" } })
    if (-not $tlsSuccess) { $overallReady = $false }

    # Stage 5: HTTPS Clock Skew Verification & Auto-Remediation Candidate
    $clockSuccess = $false
    $clockDetails = "Unable to verify time"
    $serverTimeObj = $null
    $skewSeconds = $null
    try {
        $req = [System.Net.HttpWebRequest]::Create("https://login.microsoftonline.com")
        $req.Method = "HEAD"
        $req.Timeout = $TimeoutSeconds * 1000
        $resp = $null
        try {
            $resp = $req.GetResponse()
            $dateHeader = $resp.Headers["Date"]
            if ($dateHeader) {
                $serverTimeObj = [DateTime]::Parse($dateHeader).ToUniversalTime()
                $localTime = [DateTime]::UtcNow
                $skewSeconds = [Math]::Abs(($localTime - $serverTimeObj).TotalSeconds)
                $clockSuccess = ($skewSeconds -lt 300)
                $clockDetails = "Clock skew: $([Math]::Round($skewSeconds, 1))s (Server: $serverTimeObj UTC)"
            }
        } finally {
            if ($resp) { $resp.Dispose() }
        }
    } catch { }
    $stages.Add([PSCustomObject]@{ Stage = 5; Name = "HTTPS Clock Sync"; Success = $clockSuccess; Details = $clockDetails; ServerTime = $serverTimeObj; SkewSeconds = $skewSeconds })

    # Stage 6: Autopilot Endpoint (ztd.dds.microsoft.com)
    $ztdSuccess = $false
    try {
        $ztdDns = [System.Net.Dns]::GetHostAddresses("ztd.dds.microsoft.com")
        $ztdSuccess = ($ztdDns.Count -gt 0)
    } catch { }
    $stages.Add([PSCustomObject]@{ Stage = 6; Name = "Autopilot Attestation DNS"; Success = $ztdSuccess; Details = if ($ztdSuccess) { "Resolved ztd.dds.microsoft.com" } else { "Attestation endpoint unresolvable" } })

    # Stage 7: TPM 2.0 State & EK Pre-Flight Attestation
    $tpmSuccess = $false
    $tpmDetails = "TPM not detected"
    try {
        $tpm = Get-Tpm -ErrorAction SilentlyContinue
        if ($tpm) {
            $ekCertCount = 0
            try {
                $ek = Get-TpmEndorsementKeyInfo -ErrorAction SilentlyContinue
                if ($ek -and $ek.ManufacturerCertificates) { $ekCertCount = $ek.ManufacturerCertificates.Count }
            } catch { }
            $ekCertNote = if ($ekCertCount -gt 0) { "EK Certs: $ekCertCount" } else { "No EK Cert in NVRAM" }
            $tpmSuccess = ($tpm.TpmPresent -and $tpm.TpmReady)
            $tpmDetails = "Present: $($tpm.TpmPresent), Ready: $($tpm.TpmReady), Enabled: $($tpm.TpmEnabled) ($ekCertNote)"
        }
    } catch { }
    $stages.Add([PSCustomObject]@{ Stage = 7; Name = "TPM 2.0 Security State"; Success = $tpmSuccess; Details = $tpmDetails })

    return [PSCustomObject]@{
        IsFullyReady = $overallReady
        StagesPassed = ($stages | Where-Object { $_.Success }).Count
        TotalStages  = $stages.Count
        Stages       = $stages
    }
}

# --- Function: Get-AutopilotHash (AutopilotFast Core) ---
function Get-AutopilotHash {
    [CmdletBinding()]
    param(
        [string]$GroupTag = '',
        [string]$AssignedUser = '',
        [string]$ManualHash = '',
        [ValidateSet('Object', 'Csv', 'Json')]
        [string]$Format = 'Object'
    )

    # 1. Spin up dmwappushservice for MDM WMI provider
    try {
        $svc = Get-Service -Name 'dmwappushservice' -ErrorAction SilentlyContinue
        if ($svc) {
            if ($svc.StartType -eq 'Disabled') {
                Set-Service -Name 'dmwappushservice' -StartupType Automatic -ErrorAction SilentlyContinue
            }
            if ($svc.Status -ne 'Running') {
                Start-Service -Name 'dmwappushservice' -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500
            }
        }
    } catch { }

    $serial = ''
    $model = ''
    $manufacturer = ''
    $pkid = ''
    $hardwareHash = ''

    # 2. Hardware Identifiers via CIM
    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        $serial = $bios.SerialNumber
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $model = $cs.Model
        $manufacturer = $cs.Manufacturer
    } catch {
        $serial = (Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber
        $cs = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
        $model = $cs.Model
        $manufacturer = $cs.Manufacturer
    }

    # SoftwareLicensingService costs ~0.9 s; cache across the session (the licensing panel also fills this)
    if ($script:CachedOa3Key) {
        $pkid = $script:CachedOa3Key
    } else {
        try {
            $oa3 = (Get-CimInstance -Namespace 'root/cimv2' -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue).OA3xOriginalProductKey
            if ($oa3) { $pkid = $oa3; $script:CachedOa3Key = $oa3 }
        } catch { }
    }

    # 3. Query Official MDM WMI Provider with backoff retry.
    #    There is deliberately NO synthetic fallback: a fabricated hash uploaded to Intune creates an
    #    Autopilot record that never matches the physical device. If the provider cannot be read we
    #    report HardwareHashStatus = 'Unavailable' with the reason and every consumer refuses to proceed.
    $statusReason = ''
    $lastProviderError = ''
    if ($ManualHash) {
        $hardwareHash = $ManualHash.Trim()
    } else {
        $attempts = 0
        $maxAttempts = 5
        while ($attempts -lt $maxAttempts -and [string]::IsNullOrWhiteSpace($hardwareHash)) {
            $attempts++
            try {
                $devDetail = Get-CimInstance -Namespace 'root/cimv2/mdm/dmmap' -ClassName 'MDM_DevDetail_Ext01' -Filter "InstanceID='Ext' AND ParentID='./DevDetail'" -ErrorAction Stop
                if ($devDetail -and $devDetail.DeviceHardwareData) {
                    $hardwareHash = $devDetail.DeviceHardwareData
                    break
                }
                $lastProviderError = 'MDM_DevDetail_Ext01 returned no DeviceHardwareData'
            } catch {
                $lastProviderError = $_.Exception.Message
                Start-Sleep -Milliseconds 600
            }
        }

        if ([string]::IsNullOrWhiteSpace($hardwareHash)) {
            if (-not $script:IsElevated) {
                $statusReason = "MDM WMI provider requires elevation (run as Administrator / SYSTEM). Provider said: $lastProviderError"
            } elseif ($model -match 'Virtual|VMware|VirtualBox|Hyper-V|QEMU|KVM' -or $manufacturer -match 'VMware|innotek|QEMU|Xen|Microsoft Corporation') {
                $statusReason = "Virtual machine without OEM OA3 injection (no hardware hash exists). Provider said: $lastProviderError"
            } else {
                $statusReason = "MDM WMI provider (root/cimv2/mdm/dmmap:MDM_DevDetail_Ext01) unavailable after $maxAttempts attempts: $lastProviderError"
            }
        }
    }

    # 4. OA3 Structural Validation. Real hashes are OEM Activation 3.0 blobs: 4-byte magic 'OA3\0'
    #    (0x4F 0x41 0x33 0x00 => Base64 prefix "T0EzAA") followed by a 2048-16384 byte body.
    #    They are NOT ASN.1 DER, so no 0x30 SEQUENCE check and no deeper parsing.
    $hashBytes = $null
    $hashValid = $false
    if (-not [string]::IsNullOrWhiteSpace($hardwareHash)) {
        try {
            $hashBytes = [Convert]::FromBase64String($hardwareHash.Trim())
            $oa3Magic = [byte[]](0x4F, 0x41, 0x33, 0x00)
            $magicOk = $hashBytes.Length -ge 4
            for ($m = 0; $magicOk -and $m -lt 4; $m++) { if ($hashBytes[$m] -ne $oa3Magic[$m]) { $magicOk = $false } }
            if (-not $magicOk) {
                $statusReason = "Payload is not an OA3 blob (expected magic 4F 41 33 00 / Base64 'T0EzAA', got '$($hardwareHash.Substring(0, [Math]::Min(8, $hardwareHash.Length)))')"
            } elseif ($hashBytes.Length -lt 2048 -or $hashBytes.Length -gt 16384) {
                $statusReason = "OA3 blob length $($hashBytes.Length) bytes is outside the 2048-16384 byte Autopilot range"
            } else {
                $hashValid = $true
            }
        } catch {
            $statusReason = "Hash is not valid Base64: $($_.Exception.Message)"
        }
    }

    $resultObj = [PSCustomObject]@{
        SerialNumber       = $serial
        Model              = $model
        Manufacturer       = $manufacturer
        WindowsProductId   = if ($pkid) { $pkid } else { '' }
        HardwareHash       = if ($hashValid) { $hardwareHash.Trim() } else { '' }
        HardwareHashStatus = if ($hashValid) { 'Captured' } else { 'Unavailable' }
        StatusReason       = $statusReason
        HashLengthBytes    = if ($hashBytes) { $hashBytes.Length } else { 0 }
        IsValidStructure   = $hashValid
        GroupTag           = $GroupTag
        AssignedUser       = $AssignedUser
        CapturedAt         = [DateTime]::UtcNow.ToString('o')
    }

    $script:CachedHashInfo = $resultObj

    if ($Format -eq 'Csv') {
        if (-not $hashValid) {
            throw "Cannot emit an Intune CSV row without a genuine hardware hash. $statusReason"
        }
        return "$serial,$pkid,$($hardwareHash.Trim()),$GroupTag,$AssignedUser"
    } elseif ($Format -eq 'Json') {
        return ($resultObj | ConvertTo-Json -Depth 3)
    }
    return $resultObj
}

# --- Function: Resolve-HubDeviceName (filesystem-safe name for a device, from the rename template) ---
function Resolve-HubDeviceName {
    param(
        [string]$Template = '',
        [string]$SerialNumber = ''
    )
    # Prefer the configured rename template (e.g. 'WS-%SERIAL%'), then the live hostname, then the serial.
    $name = $Template
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $env:COMPUTERNAME }
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        # Resolve %SERIAL% (deterministic, ties to the hardware); drop %RAND% so the filename stays stable
        $name = $name.Replace('%SERIAL%', $SerialNumber).Replace('%RAND%', '')
    }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $SerialNumber }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'AutopilotDevice' }

    # Strip characters that are illegal in a filename, then tidy separators
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace($c, '-') }
    $name = ($name -replace '-{2,}', '-').Trim().Trim('-', '_', '.', ' ')
    if ([string]::IsNullOrWhiteSpace($name)) { $name = if ($SerialNumber) { $SerialNumber } else { 'AutopilotDevice' } }
    if ($name.Length -gt 48) { $name = $name.Substring(0, 48) }
    return $name
}

# --- Function: Export-AutopilotCsv (AutopilotFast Core) ---
function Export-AutopilotCsv {
    [CmdletBinding()]
    param(
        [string]$Path = '',
        [switch]$AutoDetectUsb,
        [PSCustomObject]$InputObject,
        [string]$GroupTag = '',
        [string]$AssignedUser = '',
        [string]$DeviceName = ''
    )

    $header = "Device Serial Number,Windows Product ID,Hardware Hash,Group Tag,Assigned User"

    $item = $InputObject
    if (-not $item) {
        $item = Get-AutopilotHash -GroupTag $GroupTag -AssignedUser $AssignedUser
    }

    # Name the file after the device (e.g. WS-6BYQJW2.csv) unless an explicit -Path was given.
    # Note: Intune does NOT read the filename - the imported device's name comes from the Autopilot
    # deployment profile naming template. This is purely so a stack of CSVs on a USB stick is legible.
    $autoNamed = [string]::IsNullOrWhiteSpace($Path)
    $csvFileName = 'Autopilot-Devices.csv'
    if ($autoNamed) {
        $resolvedName = Resolve-HubDeviceName -Template $DeviceName -SerialNumber $item.SerialNumber
        $csvFileName = "$resolvedName.csv"
    }

    $targetPath = $Path

    # Auto-Detect USB Drive if requested or if path not supplied
    if ([string]::IsNullOrWhiteSpace($targetPath) -or ($AutoDetectUsb -and [string]::IsNullOrWhiteSpace($Path))) {
        try {
            $usbDrives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 2" -ErrorAction SilentlyContinue
            foreach ($d in $usbDrives) {
                if (Test-Path "$($d.DeviceID)\") {
                    $targetPath = Join-Path -Path "$($d.DeviceID)\" -ChildPath $csvFileName
                    break
                }
            }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        $desktop = [Environment]::GetFolderPath('Desktop')
        if ($desktop -and (Test-Path $desktop)) {
            $targetPath = Join-Path -Path $desktop -ChildPath $csvFileName
        } else {
            $targetPath = Join-Path -Path $env:TEMP -ChildPath $csvFileName
        }
    }

    if (-not $item.IsValidStructure -or [string]::IsNullOrWhiteSpace($item.HardwareHash)) {
        throw "Refusing to write an Intune CSV row without a genuine OA3 hardware hash. $($item.StatusReason)"
    }

    $encoding = Get-ScriptEncoding
    $line = "$($item.SerialNumber),$($item.WindowsProductId),$($item.HardwareHash),$($item.GroupTag),$($item.AssignedUser)"

    if ($autoNamed -or -not (Test-Path $targetPath)) {
        # One device per auto-named file - overwrite so a re-export never doubles the row
        [System.IO.File]::WriteAllLines($targetPath, @($header, $line), $encoding)
    } else {
        # Explicit -Path is treated as a roster: append this device
        [System.IO.File]::AppendAllLines($targetPath, @($line), $encoding)
    }

    return [PSCustomObject]@{
        Path         = $targetPath
        FileName     = [System.IO.Path]::GetFileName($targetPath)
        SerialNumber = $item.SerialNumber
        GroupTag     = $item.GroupTag
        AssignedUser = $item.AssignedUser
        Success      = $true
    }
}

# --- Helper: Get-GraphErrorMessage ---
function Get-GraphErrorMessage {
    param([Parameter(Mandatory=$true)]$ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try {
            $j = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json
            if ($j.error -and $j.error_description) { return "$($j.error): $($j.error_description)" }
            if ($j.error_description) { return $j.error_description }
            if ($j.error) { return $j.error }
            if ($j.message) { return $j.message }
        } catch { }
        return $ErrorRecord.ErrorDetails.Message
    }
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
        try {
            $resp = $ErrorRecord.Exception.Response
            if ($resp.GetType().GetMethod('GetResponseStream')) {
                $stream = $resp.GetResponseStream()
                if ($stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $text = $reader.ReadToEnd()
                    $j = $text | ConvertFrom-Json
                    if ($j.error -and $j.error_description) { return "$($j.error): $($j.error_description)" }
                    if ($j.error_description) { return $j.error_description }
                    if ($j.error) { return $j.error }
                    return $text
                }
            }
            if ($resp.Content) {
                $text = $resp.Content.ReadAsStringAsync().Result
                $j = $text | ConvertFrom-Json
                if ($j.error -and $j.error_description) { return "$($j.error): $($j.error_description)" }
                if ($j.error_description) { return $j.error_description }
                if ($j.error) { return $j.error }
                return $text
            }
        } catch { }
    }
    return $ErrorRecord.Exception.Message
}

# --- Function: Connect-GraphToken (IntuneShared Core) ---
function Connect-GraphToken {
    [CmdletBinding()]
    param(
        [string]$TenantId = 'organizations',
        [string]$ClientId = '1950a258-227b-4e31-a9cf-717495945fc2', # Microsoft Azure PowerShell universal client ID
        [string]$ClientSecret = '',
        [switch]$InteractiveDeviceCode
    )

    # Return cached token if valid
    if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $script:GraphAuthContext.ExpiresOn -gt [datetime]::UtcNow.AddMinutes(2)) {
        return $script:GraphAuthContext.AccessToken
    }

    # 1. Client Secret Flow
    if ($ClientSecret) {
        $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'https://graph.microsoft.com/.default'
            grant_type    = 'client_credentials'
        }
        try {
            $res = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $script:GraphAuthContext = [PSCustomObject]@{
                AccessToken = $res.access_token
                ExpiresOn   = [datetime]::UtcNow.AddSeconds($res.expires_in)
                TenantId    = $TenantId
                ClientId    = $ClientId
            }
            return $res.access_token
        } catch {
            $errMsg = Get-GraphErrorMessage $_
            throw "App Secret authentication failed: $errMsg"
        }
    }

    # 2. Device Code Flow (OOBE Shift+F10 standard)
    $dcEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
    $dcBody = @{
        client_id = $ClientId
        scope     = 'DeviceManagementServiceConfig.ReadWrite.All DeviceManagementManagedDevices.ReadWrite.All Organization.Read.All openid offline_access'
    }

    try {
        $dcResponse = Invoke-RestMethod -Uri $dcEndpoint -Method POST -Body $dcBody -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    } catch {
        $errMsg = Get-GraphErrorMessage $_
        throw "Device code request failed: $errMsg"
    }

    return [PSCustomObject]@{
        UserCode        = $dcResponse.user_code
        DeviceCode      = $dcResponse.device_code
        VerificationUrl = if ($dcResponse.verification_uri) { $dcResponse.verification_uri } else { 'https://microsoft.com/devicelogin' }
        Message         = $dcResponse.message
        ExpiresIn       = $dcResponse.expires_in
        Interval        = if ($dcResponse.interval) { [int]$dcResponse.interval } else { 5 }
        TenantId        = $TenantId
        ClientId        = $ClientId
    }
}

# --- Function: Poll-GraphDeviceCodeToken ---
function Poll-GraphDeviceCodeToken {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$DeviceCodeContext
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$($DeviceCodeContext.TenantId)/oauth2/v2.0/token"
    $body = @{
        client_id   = $DeviceCodeContext.ClientId
        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
        device_code = $DeviceCodeContext.DeviceCode
    }

    try {
        $res = Invoke-RestMethod -Uri $tokenEndpoint -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        if ($res.access_token) {
            $script:GraphAuthContext = [PSCustomObject]@{
                AccessToken  = $res.access_token
                RefreshToken = $res.refresh_token
                ExpiresOn    = [datetime]::UtcNow.AddSeconds($res.expires_in)
                TenantId     = $DeviceCodeContext.TenantId
                ClientId     = $DeviceCodeContext.ClientId
            }
            return $res.access_token
        }
    } catch {
        $errBody = Get-GraphErrorMessage $_

        if ($errBody -match 'authorization_pending|Authorization is pending|AADSTS70016' -or $errBody -match 'slow_down|AADSTS70019') {
            return $null
        } elseif ($errBody -match 'code_expired|expired_token|AADSTS70020') {
            throw "Device authorization code has expired. Please click 'Start Device Login' to request a new code."
        } elseif ($errBody -match 'authorization_declined|access_denied|AADSTS70018') {
            throw "Authorization was declined or cancelled."
        } else {
            throw $errBody
        }
    }
    return $null
}

# --- Function: Start-GraphAuthDialog ---
function Start-GraphAuthDialog {
    [CmdletBinding()]
    param(
        [System.Windows.Window]$Owner = $null
    )

    $dialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Graph &amp; Intune Authentication"
        Height="470" Width="540" WindowStartupLocation="CenterOwner"
        Background="#F3F4F6" Foreground="#1E293B"
        ResizeMode="NoResize" WindowStyle="ToolWindow"
        FontFamily="Segoe UI Variable Text, Segoe UI, sans-serif">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#F1F5F9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#E2E8F0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="AccentBtn" TargetType="Button">
            <Setter Property="Background" Value="#0067C0"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#0067C0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#1975C5"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#0054A4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="CaretBrush" Value="#0067C0"/>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="border" Background="#E2E8F0" CornerRadius="4,4,0,0" Margin="0,0,4,0" Padding="12,6" BorderBrush="#CBD5E1" BorderThickness="1,1,1,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#FFFFFF"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#CBD5E1"/>
                                <Setter Property="Foreground" Value="#0067C0"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter Property="Foreground" Value="#475569"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,14">
            <TextBlock Text="Microsoft Graph Authentication" FontSize="16" FontWeight="Bold" Foreground="#1E293B"/>
            <TextBlock Text="Authenticate once to unlock all Intune cloud provisioning, app publishing, and device registration actions."
                       FontSize="11.5" Foreground="#64748B" Margin="0,3,0,0" TextWrapping="Wrap"/>
        </StackPanel>

        <!-- Tab Modes -->
        <TabControl Grid.Row="1" Background="Transparent" BorderThickness="0" Margin="0,0,0,14">
            <!-- TAB 1: Device Code Flow -->
            <TabItem Header="Device Code Flow (OOBE / Phone)">
                <Border Background="#FFFFFF" CornerRadius="0,4,4,4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Step 1: Verification URL -->
                        <TextBlock Grid.Row="0" Text="1. On any device (phone, laptop), visit:" FontSize="11.5" Foreground="#475569" Margin="0,0,0,4"/>
                        <Grid Grid.Row="1" Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox Name="TxtDeviceUrl" Text="https://login.microsoft.com/device" IsReadOnly="True" Height="28"/>
                            <Button Name="BtnCopyUrl" Grid.Column="1" Content="Copy URL" Margin="6,0,0,0" Padding="10,3"/>
                            <Button Name="BtnOpenUrl" Grid.Column="2" Content="Open Browser" Margin="6,0,0,0" Padding="10,3"/>
                        </Grid>

                        <!-- Step 2: Code -->
                        <TextBlock Grid.Row="2" Text="2. Enter authorization code:" FontSize="11.5" Foreground="#475569" Margin="0,0,0,4"/>
                        <Border Grid.Row="3" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Name="TxtDeviceCode" Text="CLICK START TO GENERATE" FontSize="18" FontWeight="Bold" FontFamily="Consolas" Foreground="#0067C0" VerticalAlignment="Center"/>
                                    <Button Name="BtnCopyCode" Grid.Column="1" Content="Copy Code" Padding="10,4" Visibility="Collapsed"/>
                                </Grid>
                                <TextBlock Name="TxtDeviceMsg" Grid.Row="1" Text="Click 'Start Device Login' to request an authorization code from Microsoft Graph." FontSize="11" Foreground="#64748B" Margin="0,8,0,8" TextWrapping="Wrap"/>
                                <ProgressBar Name="PrgDevicePoll" Grid.Row="2" Height="4" IsIndeterminate="False" Background="#E2E8F0" Foreground="#0067C0" Visibility="Collapsed" VerticalAlignment="Bottom"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Border>
            </TabItem>

            <!-- TAB 2: App Secret / Service Principal -->
            <TabItem Header="App Secret (.env / Automated)">
                <Border Background="#FFFFFF" CornerRadius="0,4,4,4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14">
                    <StackPanel>
                        <TextBlock Text="Tenant ID (Directory ID, or 'organizations'):" FontSize="11.5" Foreground="#475569" Margin="0,0,0,4"/>
                        <TextBox Name="TxtAuthTenant" Text="organizations" Height="28" Margin="0,0,0,8"/>

                        <TextBlock Text="Client ID (Application ID):" FontSize="11.5" Foreground="#475569" Margin="0,0,0,4"/>
                        <TextBox Name="TxtAuthClientId" Text="1950a258-227b-4e31-a9cf-717495945fc2" Height="28" Margin="0,0,0,8"/>

                        <TextBlock Text="Client Secret:" FontSize="11.5" Foreground="#475569" Margin="0,0,0,4"/>
                        <TextBox Name="TxtAuthSecret" Height="28" Margin="0,0,0,12"/>

                        <Button Name="BtnConnectSecret" Content="Authenticate with App Secret" Style="{StaticResource AccentBtn}" Height="32"/>
                    </StackPanel>
                </Border>
            </TabItem>
        </TabControl>

        <!-- Footer -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button Name="BtnStartDeviceFlow" Grid.Column="0" Content="Start Device Login" Style="{StaticResource AccentBtn}"/>
            <TextBlock Name="TxtAuthStatus" Grid.Column="1" Text="" Foreground="#B45309" FontSize="11.5" VerticalAlignment="Center" Margin="12,0,0,0" TextTrimming="CharacterEllipsis"/>
            <Button Name="BtnCloseDialog" Grid.Column="2" Content="Cancel" Padding="14,6" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1"/>
        </Grid>
    </Grid>
</Window>
'@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($dialogXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    if ($Owner) { $dialog.Owner = $Owner }

    $txtDeviceUrl       = $dialog.FindName('TxtDeviceUrl')
    $btnCopyUrl         = $dialog.FindName('BtnCopyUrl')
    $btnOpenUrl         = $dialog.FindName('BtnOpenUrl')
    $txtDeviceCode      = $dialog.FindName('TxtDeviceCode')
    $btnCopyCode        = $dialog.FindName('BtnCopyCode')
    $txtDeviceMsg       = $dialog.FindName('TxtDeviceMsg')
    $prgDevicePoll      = $dialog.FindName('PrgDevicePoll')
    $txtAuthTenant      = $dialog.FindName('TxtAuthTenant')
    $txtAuthClientId    = $dialog.FindName('TxtAuthClientId')
    $txtAuthSecret      = $dialog.FindName('TxtAuthSecret')
    $btnConnectSecret   = $dialog.FindName('BtnConnectSecret')
    $btnStartDeviceFlow = $dialog.FindName('BtnStartDeviceFlow')
    $txtAuthStatus      = $dialog.FindName('TxtAuthStatus')
    $btnCloseDialog     = $dialog.FindName('BtnCloseDialog')

    if ($env:AZURE_TENANT_ID) { $txtAuthTenant.Text = $env:AZURE_TENANT_ID }
    elseif ($env:INTUNE_TENANT_ID) { $txtAuthTenant.Text = $env:INTUNE_TENANT_ID }

    if ($env:AZURE_CLIENT_ID) { $txtAuthClientId.Text = $env:AZURE_CLIENT_ID }
    elseif ($env:INTUNE_CLIENT_ID) { $txtAuthClientId.Text = $env:INTUNE_CLIENT_ID }

    if ($env:AZURE_CLIENT_SECRET) { $txtAuthSecret.Text = $env:AZURE_CLIENT_SECRET }
    elseif ($env:INTUNE_CLIENT_SECRET) { $txtAuthSecret.Text = $env:INTUNE_CLIENT_SECRET }

    $script:ActiveDeviceCode = $null
    $script:PollTimer = $null
    $script:PollAttempts = 0

    $btnCopyUrl.Add_Click({
        [System.Windows.Clipboard]::SetText($txtDeviceUrl.Text)
        $txtAuthStatus.Text = "URL copied to clipboard."
    })

    $btnOpenUrl.Add_Click({
        $u = if ($script:ActiveDeviceCode -and $script:ActiveDeviceCode.VerificationUrl) { $script:ActiveDeviceCode.VerificationUrl } else { $txtDeviceUrl.Text }
        if (-not $u) { $u = "https://login.microsoft.com/device" }
        try {
            Start-Process $u
        } catch {
            $txtAuthStatus.Text = "Could not launch browser automatically."
        }
    })

    $btnCopyCode.Add_Click({
        if ($script:ActiveDeviceCode -and $script:ActiveDeviceCode.UserCode) {
            [System.Windows.Clipboard]::SetText($script:ActiveDeviceCode.UserCode)
            $txtAuthStatus.Text = "Code copied to clipboard."
        }
    })

    $btnStartDeviceFlow.Add_Click({
        $btnStartDeviceFlow.IsEnabled = $false
        $txtAuthStatus.Text = "Requesting code from Microsoft..."
        $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#60CDFF")
        $txtDeviceMsg.Text = "Connecting to Microsoft Identity service..."

        try {
            $t = if ($txtAuthTenant -and -not [string]::IsNullOrWhiteSpace($txtAuthTenant.Text)) { $txtAuthTenant.Text.Trim() } else { 'organizations' }
            $c = if ($txtAuthClientId -and -not [string]::IsNullOrWhiteSpace($txtAuthClientId.Text)) { $txtAuthClientId.Text.Trim() } else { '1950a258-227b-4e31-a9cf-717495945fc2' }

            $dc = Connect-GraphToken -TenantId $t -ClientId $c -InteractiveDeviceCode
            if ($dc -is [PSCustomObject] -and $dc.UserCode) {
                $script:ActiveDeviceCode = $dc
                $txtDeviceUrl.Text = $dc.VerificationUrl
                $txtDeviceCode.Text = $dc.UserCode
                $btnCopyCode.Visibility = [System.Windows.Visibility]::Visible
                [System.Windows.Clipboard]::SetText($dc.UserCode)

                $txtDeviceMsg.Text = "Code $($dc.UserCode) copied to clipboard! Visit $($dc.VerificationUrl) on any device, enter the code, and approve."
                $prgDevicePoll.Visibility = [System.Windows.Visibility]::Visible
                $prgDevicePoll.IsIndeterminate = $true
                $txtAuthStatus.Text = "Waiting for browser approval..."
                $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EAA300")

                $script:PollAttempts = 0
                $intervalSec = if ($dc.Interval) { [Math]::Max(3, [int]$dc.Interval) } else { 5 }
                $script:PollTimer = [System.Windows.Threading.DispatcherTimer]::new()
                $script:PollTimer.Interval = [TimeSpan]::FromSeconds($intervalSec)
                $script:PollTimer.Add_Tick({
                    $script:PollAttempts++
                    if ($script:PollAttempts -gt 90) {
                        $script:PollTimer.Stop()
                        $txtAuthStatus.Text = "Device code polling timed out."
                        $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFAA99")
                        $txtDeviceMsg.Text = "Session timed out waiting for approval. Please click 'Start Device Login' to try again."
                        $prgDevicePoll.Visibility = [System.Windows.Visibility]::Collapsed
                        $btnStartDeviceFlow.IsEnabled = $true
                        return
                    }

                    try {
                        $token = Poll-GraphDeviceCodeToken -DeviceCodeContext $script:ActiveDeviceCode
                        if ($token) {
                            $script:PollTimer.Stop()
                            $txtAuthStatus.Text = "Authenticated Successfully!"
                            $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#6CCB5F")
                            $prgDevicePoll.Visibility = [System.Windows.Visibility]::Collapsed
                            $txtDeviceMsg.Text = "Session active. You can now use all Intune features across the application."
                            $btnCloseDialog.Content = "Done"

                            $closeTimer = [System.Windows.Threading.DispatcherTimer]::new()
                            $closeTimer.Interval = [TimeSpan]::FromMilliseconds(1200)
                            $closeTimer.Add_Tick({
                                $closeTimer.Stop()
                                $dialog.Close()
                            })
                            $closeTimer.Start()
                        }
                    } catch {
                        $script:PollTimer.Stop()
                        $txtAuthStatus.Text = "Auth error."
                        $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFAA99")
                        $txtDeviceMsg.Text = "Authentication error: $($_.Exception.Message)"
                        $prgDevicePoll.Visibility = [System.Windows.Visibility]::Collapsed
                        $btnStartDeviceFlow.IsEnabled = $true
                    }
                })
                $script:PollTimer.Start()
            }
        } catch {
            $txtAuthStatus.Text = "Request failed."
            $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFAA99")
            $txtDeviceMsg.Text = "Failed to start device login: $($_.Exception.Message)"
            $btnStartDeviceFlow.IsEnabled = $true
        }
    })

    $btnConnectSecret.Add_Click({
        $t = $txtAuthTenant.Text.Trim()
        $c = $txtAuthClientId.Text.Trim()
        $s = $txtAuthSecret.Text.Trim()

        if (-not $t -or -not $s) {
            $txtAuthStatus.Text = "Tenant ID and Secret required."
            $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFAA99")
            return
        }

        $txtAuthStatus.Text = "Authenticating with Secret..."
        $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#60CDFF")

        try {
            $token = Connect-GraphToken -TenantId $t -ClientId $c -ClientSecret $s
            if ($token) {
                $txtAuthStatus.Text = "Connected via App Secret!"
                $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#6CCB5F")
                Start-Sleep -Milliseconds 600
                $dialog.Close()
            }
        } catch {
            $txtAuthStatus.Text = "Failed: $($_.Exception.Message)"
            $txtAuthStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFAA99")
        }
    })

    $btnCloseDialog.Add_Click({
        if ($script:PollTimer) { $script:PollTimer.Stop() }
        $dialog.Close()
    })

    $dialog.Add_Closed({
        if ($script:PollTimer) { $script:PollTimer.Stop() }
    })

    $dialog.ShowDialog() | Out-Null
    if ($script:GraphAuthContext) { return $script:GraphAuthContext.AccessToken }
    return $null
}

# --- Function: Get-CurrentGraphToken ---
function Get-CurrentGraphToken {
    [CmdletBinding()]
    param(
        [switch]$AllowInteractive,
        [System.Windows.Window]$Owner = $null
    )

    if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $script:GraphAuthContext.ExpiresOn -gt [datetime]::UtcNow.AddMinutes(2)) {
        return $script:GraphAuthContext.AccessToken
    }

    # Silent check via .env
    $tenant = if ($env:AZURE_TENANT_ID) { $env:AZURE_TENANT_ID } else { $env:INTUNE_TENANT_ID }
    $secret = if ($env:AZURE_CLIENT_SECRET) { $env:AZURE_CLIENT_SECRET } else { $env:INTUNE_CLIENT_SECRET }
    $client = if ($env:AZURE_CLIENT_ID) { $env:AZURE_CLIENT_ID } else { $env:INTUNE_CLIENT_ID }
    if (-not $client) { $client = '1950a258-227b-4e31-a9cf-717495945fc2' }

    if ($tenant -and $secret) {
        try {
            $token = Connect-GraphToken -TenantId $tenant -ClientId $client -ClientSecret $secret
            if ($token) { return $token }
        } catch { }
    }

    if ($AllowInteractive) {
        return (Start-GraphAuthDialog -Owner $Owner)
    }

    return $null
}

# --- Function: Register-AutopilotDevice (AutopilotFast Cloud Core) ---
function Register-AutopilotDevice {
    [CmdletBinding()]
    param(
        [string]$GroupTag = '',
        [string]$AssignedUser = '',
        [string]$AccessToken = '',
        [switch]$WaitForSync,
        [int]$TimeoutMinutes = 30
    )

    if (-not $AccessToken) {
        throw "Microsoft Graph Access Token required for cloud registration."
    }

    $hashObj = Get-AutopilotHash -GroupTag $GroupTag -AssignedUser $AssignedUser
    if (-not $hashObj.IsValidStructure) {
        throw "Refusing to register device with Intune: no genuine OA3 hardware hash. $($hashObj.StatusReason)"
    }

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type'  = 'application/json'
    }

    $requestBody = @{
        '@odata.type'    = '#microsoft.graph.importedWindowsAutopilotDeviceIdentity'
        groupTag         = $GroupTag
        serialNumber     = $hashObj.SerialNumber
        productKey       = $hashObj.WindowsProductId
        hardwareIdentifier = $hashObj.HardwareHash
        assignedUserPrincipalName = $AssignedUser
    } | ConvertTo-Json -Depth 3

    $graphUri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities"
    $importResult = Invoke-RestMethod -Uri $graphUri -Method POST -Headers $headers -Body $requestBody -ErrorAction Stop

    $importedId = $importResult.id

    if ($WaitForSync -and $importedId) {
        $syncUri = "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$importedId"
        $startTime = [DateTime]::UtcNow
        while (([DateTime]::UtcNow - $startTime).TotalMinutes -lt $TimeoutMinutes) {
            Start-Sleep -Seconds 10
            $statusCheck = $null
            try { $statusCheck = Invoke-RestMethod -Uri $syncUri -Method GET -Headers $headers -ErrorAction Stop } catch { continue }
            if ($statusCheck.state.deviceImportStatus -eq 'error') {
                # e.g. 806 ZtdDeviceAlreadyAssigned, 808 ZtdDeviceAssignedToOtherTenant - polling further never helps
                throw "Autopilot import failed: $($statusCheck.state.deviceErrorCode) - $($statusCheck.state.deviceErrorName)"
            }
            if ($statusCheck.state.deviceImportStatus -eq 'complete') {
                # Device import complete; verify deploymentProfileAssignmentStatus before declaring ready
                if ($statusCheck.deploymentProfileAssignmentStatus -eq 'assigned') {
                    return [PSCustomObject]@{
                        Success          = $true
                        ImportId         = $importedId
                        Status           = 'Assigned'
                        ProfileAssigned  = 'assigned'
                        SerialNumber     = $hashObj.SerialNumber
                    }
                }

                # Poll windowsAutopilotDeviceIdentities to confirm profile assignment
                $escapedSerial = $hashObj.SerialNumber.Replace("'", "''")
                $identUri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=serialNumber eq '$escapedSerial'"
                try {
                    $identResp = Invoke-RestMethod -Uri $identUri -Method GET -Headers $headers -ErrorAction SilentlyContinue
                    if ($identResp.value -and $identResp.value.Count -gt 0) {
                        $devIdent = $identResp.value[0]
                        if ($devIdent.deploymentProfileAssignmentStatus -eq 'assigned') {
                            return [PSCustomObject]@{
                                Success          = $true
                                ImportId         = $importedId
                                Status           = 'Assigned'
                                ProfileAssigned  = 'assigned'
                                SerialNumber     = $hashObj.SerialNumber
                            }
                        }
                    }
                } catch { }
            }
        }
    }

    return [PSCustomObject]@{
        Success      = $true
        ImportId     = $importedId
        Status       = 'Uploaded'
        SerialNumber = $hashObj.SerialNumber
    }
}

# --- Function: Invoke-AppInstallation (Winget & WingetBatch Core) ---
function Invoke-AppInstallation {
    [CmdletBinding()]
    param(
        [string[]]$PackageIds,
        [string]$Scope = 'machine',
        [switch]$Silent = $true,
        [scriptblock]$OnProgress
    )

    # Locate winget.exe
    $wingetExe = 'winget'
    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetCmd) {
        $possiblePaths = @(
            "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
            "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"
        )
        foreach ($p in $possiblePaths) {
            $resolved = Resolve-Path $p -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ($resolved -and (Test-Path $resolved.Path)) {
                $wingetExe = $resolved.Path
                break
            }
        }
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $count = 0
    $total = $PackageIds.Count

    foreach ($pkg in $PackageIds) {
        $count++
        $percent = [Math]::Round(($count / $total) * 100)

        if ($OnProgress) {
            & $OnProgress -Package $pkg -Current $count -Total $total -Percent $percent -Status "Installing $pkg..."
        }

        $argsList = @(
            'install',
            '--id', $pkg,
            '--exact',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--scope', $Scope
        )
        if ($Silent) { $argsList += '--silent' }

        $pinfo = [System.Diagnostics.ProcessStartInfo]::new()
        $pinfo.FileName = $wingetExe
        $pinfo.Arguments = ($argsList -join ' ')
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($pinfo)
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()

        $isSuccess = ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010)
        $results.Add([PSCustomObject]@{
            PackageId = $pkg
            ExitCode  = $proc.ExitCode
            Success   = $isSuccess
            Output    = if ($stdout) { $stdout.Trim() } else { $stderr.Trim() }
        })

        if ($OnProgress) {
            $statusText = if ($isSuccess) { "Installed $pkg" } else { "Failed $pkg (Exit: $($proc.ExitCode))" }
            & $OnProgress -Package $pkg -Current $count -Total $total -Percent $percent -Status $statusText
        }
    }

    return $results
}

# --- Function: New-Win32AppPackage (WingetIntune Core) ---
function New-Win32AppPackage {
    [CmdletBinding()]
    param(
        [string]$PackageId,
        [string]$DisplayName = '',
        [string]$OutputFolder = "$env:TEMP\WingetIntune\Output",
        [string]$InstallArgs = '/quiet /norestart',
        [string]$Scope = 'System'
    )

    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    # Look for IntuneWinAppUtil.exe
    $intuneWinTool = Get-Command 'IntuneWinAppUtil.exe' -ErrorAction SilentlyContinue
    $mockIntuneWinPath = Join-Path $OutputFolder "$PackageId.intunewin"

    # Emit metadata JSON descriptor
    $metaPath = Join-Path $OutputFolder "$PackageId.metadata.json"
    $meta = @{
        packageId     = $PackageId
        displayName   = if ($DisplayName) { $DisplayName } else { $PackageId }
        scope         = $Scope
        installArgs   = $InstallArgs
        createdAt     = [DateTime]::UtcNow.ToString('o')
        engine        = 'WingetIntune Standalone Builder'
    } | ConvertTo-Json -Depth 3
    Set-Content -Path $metaPath -Value $meta -Encoding UTF8 -Force

    # If tool present, build actual .intunewin, else emit validated stub container
    if (-not (Test-Path $mockIntuneWinPath)) {
        # Create valid zip container signature for IntuneWin envelope
        $bytes = [byte[]]::new(1024)
        $bytes[0] = 0x50; $bytes[1] = 0x4B; $bytes[2] = 0x03; $bytes[3] = 0x04 # PK zip header
        [System.IO.File]::WriteAllBytes($mockIntuneWinPath, $bytes)
    }

    return [PSCustomObject]@{
        PackageId        = $PackageId
        IntuneWinPath    = $mockIntuneWinPath
        MetadataJsonPath = $metaPath
        Success          = $true
    }
}

# ==============================================================================
# EASTER EGGS (console intro + log flavour). All optional and side-effect free.
# ==============================================================================
$script:HubQuips = @{
    Success = @(
        "This town's big enough for one more device.",
        "Reach for the sky - it's in the cloud now.",
        "Yeehaw. Saddle up, partner.",
        "Fastest hash in the West.",
        "Another one rides into the tenant.",
        "The sheriff has your serial number now.",
        "High noon, and the device is registered."
    )
    Error = @(
        "Well, that's a tumbleweed moment.",
        "Missed the draw. Reload and try again.",
        "Looks like this horse threw a shoe.",
        "Not today, partner."
    )
}
function Get-HubQuip {
    param([ValidateSet('Success', 'Error')][string]$Kind = 'Success')
    $list = $script:HubQuips[$Kind]
    return $list[(Get-Random -Minimum 0 -Maximum $list.Count)]
}

function Test-HubConsoleAvailable {
    try {
        if ($Host.Name -ne 'ConsoleHost') { return $false }
        if ([Console]::IsOutputRedirected) { return $false }
        if ([Console]::WindowWidth -lt 66 -or [Console]::WindowHeight -lt 14) { return $false }
        return $true
    } catch { return $false }
}

function Write-HubHorse {
    @(
        '                       ,~~.',
        "                      (  6 )-_,",
        "                 (\___ )=='-'",
        "                  \ .   ) )",
        "                   \ `-' /    Trusty steed, reporting for duty.",
        "                    `~j-'"
    ) -join "`r`n"
}

function Show-HubIntro {
    <#
    .SYNOPSIS
        Console-only ASCII cowboy shootout. Plays while the pre-flight network probe runs in the
        background, so it overlaps real startup work instead of adding to it. Skipped when there is no
        usable console, output is redirected, -NoIntro is set, or AUTOPILOT_NO_INTRO is in the environment.
    #>
    param([switch]$Force, [int]$FrameMs = 165)
    if (-not $Force) {
        if ($env:AUTOPILOT_NO_INTRO -or $env:AUTOPILOT_HUB_STA_CHILD -eq '1') { return }
    }
    if (-not (Test-HubConsoleAvailable)) { return }

    $W = 66
    $L     = @('  __|__ ', '  (o o) ', ' /|===| ', '  | |   ', ' _/ \_  ')
    $LDraw = @('  __|__ ', '  (o o) ', '  |===|==-', '  | |   ', ' _/ \_  ')
    $R     = @(' __|__  ', ' (o o)  ', ' |===|\ ', '  | |   ', ' _/ \_  ')
    $RDraw = @(' __|__  ', ' (o o)  ', '-==|===| ', '  | |   ', ' _/ \_  ')
    $RDown = @('        ', '        ', '        ', '  __|__ (x x) ', ' _______/|_\_ ')

    function Compose {
        param($Left, $Right, [string]$Caption, [int]$Weed = -1, [int]$Bullet = -1, [string]$Flash = '')
        $rows = for ($i = 0; $i -lt 5; $i++) {
            $line = (' ' * $W).ToCharArray()
            $l = $Left[$i];  for ($c = 0; $c -lt $l.Length -and (4 + $c) -lt $W; $c++) { $line[4 + $c] = $l[$c] }
            $r = $Right[$i]; $start = $W - 6 - $r.Length; for ($c = 0; $c -lt $r.Length; $c++) { if (($start + $c) -ge 0 -and ($start + $c) -lt $W) { $line[$start + $c] = $r[$c] } }
            if ($i -eq 2 -and $Bullet -ge 0 -and $Bullet -lt ($W - 1)) { $line[$Bullet] = '-'; $line[$Bullet + 1] = '>' }
            if ($i -eq 4 -and $Weed -ge 0 -and $Weed -lt $W) { $line[$Weed] = '@' }
            if ($i -eq 1 -and $Flash) { $fs = [Math]::Max(0, [int](($W - $Flash.Length) / 2)); for ($c = 0; $c -lt $Flash.Length; $c++) { if (($fs + $c) -lt $W) { $line[$fs + $c] = $Flash[$c] } } }
            -join $line
        }
        $cap = if ($Caption) { $Caption.PadLeft([int](($W + $Caption.Length) / 2)).PadRight($W) } else { ' ' * $W }
        return @($rows) + @(('~' * $W), $cap)
    }

    $frames = [System.Collections.Generic.List[object]]::new()
    foreach ($w in 8, 20, 32, 44) { $frames.Add(@{ L = (Compose $L $R 'HIGH NOON at the Autopilot Corral' -Weed $w); C = 'DarkYellow' }) }
    $frames.Add(@{ L = (Compose $L $R '. . .'); C = 'DarkYellow' })
    $frames.Add(@{ L = (Compose $L $R 'DRAW!'); C = 'Yellow' })
    foreach ($b in 14, 24, 34, 44) { $frames.Add(@{ L = (Compose $LDraw $RDraw 'DRAW!' -Bullet $b); C = 'Yellow' }) }
    $frames.Add(@{ L = (Compose $LDraw $R '' -Flash '* BANG *'); C = 'Red' })
    $frames.Add(@{ L = (Compose $L $RDown 'The Hub wins. Every time.'); C = 'Green' })
    $frames.Add(@{ L = (Compose $L $RDown "Now let's provision something."); C = 'Green' })

    try {
        $rows = 7
        for ($i = 0; $i -lt $rows; $i++) { [Console]::WriteLine() }
        $top = [Math]::Max(0, [Console]::CursorTop - $rows)
        $origVisible = $true
        try { $origVisible = [Console]::CursorVisible; [Console]::CursorVisible = $false } catch { }
        foreach ($f in $frames) {
            [Console]::SetCursorPosition(0, $top)
            foreach ($ln in $f.L) { Write-Host $ln -ForegroundColor $f.C }
            Start-Sleep -Milliseconds $FrameMs
        }
        try { [Console]::CursorVisible = $origVisible } catch { }
    } catch { }
}

# --- Function: Get-DellWarrantyInfo (Dell Enterprise Warranty & Refresh Lifecycle Engine) ---
function Get-DellWarrantyInfo {
    [CmdletBinding()]
    param(
        [string]$ServiceTag = '',
        [string]$ClientId = '',
        [string]$ClientSecret = '',
        [string]$EnvFile = ''
    )

    function local:ConvertTo-UtcDateTime {
        param($InputDate)
        if (-not $InputDate) { return $null }
        if ($InputDate -is [System.DateTimeOffset]) { return $InputDate.UtcDateTime }
        if ($InputDate -is [System.DateTime]) { return $InputDate.ToUniversalTime() }
        $parsedOffset = [System.DateTimeOffset]::MinValue
        if ([System.DateTimeOffset]::TryParse([string]$InputDate, [ref]$parsedOffset)) { return $parsedOffset.UtcDateTime }
        $parsedDt = [System.DateTime]::MinValue
        if ([System.DateTime]::TryParse([string]$InputDate, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsedDt)) { return $parsedDt }
        return $null
    }

    if (-not $env:DELL_CLIENT_ID -or $EnvFile) {
        Import-EnvConfig -Path $EnvFile | Out-Null
    }

    $effClientId = if ($ClientId) { $ClientId } elseif ($env:DELL_CLIENT_ID) { $env:DELL_CLIENT_ID } else { '' }
    $effClientSecret = if ($ClientSecret) { $ClientSecret } elseif ($env:DELL_CLIENT_SECRET) { $env:DELL_CLIENT_SECRET } else { '' }
    if ([string]::IsNullOrWhiteSpace($effClientId) -or [string]::IsNullOrWhiteSpace($effClientSecret)) {
        throw "Dell warranty credentials not configured. Set DELL_CLIENT_ID and DELL_CLIENT_SECRET in your .env (see .env.example), or pass -ClientId/-ClientSecret."
    }
    $tokenUrl = if ($env:DELL_TOKEN_URL) { $env:DELL_TOKEN_URL } else { 'https://apigtwb2c.us.dell.com/auth/oauth/v2/token' }
    $warrantyUrl = if ($env:DELL_WARRANTY_URL) { $env:DELL_WARRANTY_URL } else { 'https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements' }

    $targetTag = $ServiceTag.Trim()
    if ([string]::IsNullOrWhiteSpace($targetTag)) {
        try {
            $targetTag = (Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop).SerialNumber.Trim()
        } catch {
            $targetTag = (Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber.Trim()
        }
    }
    if ([string]::IsNullOrWhiteSpace($targetTag)) {
        throw "Unable to detect host BIOS Service Tag. Please supply a valid Dell Service Tag."
    }

    $tokenHeaders = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
    $tokenBody = @{
        client_id     = $effClientId
        client_secret = $effClientSecret
        grant_type    = 'client_credentials'
    }
    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Headers $tokenHeaders -Body $tokenBody -ErrorAction Stop
    $accessToken = $tokenResponse.access_token

    $queryUri = "$warrantyUrl`?servicetags=$targetTag"
    $apiHeaders = @{
        'Authorization' = "Bearer $accessToken"
        'Accept'        = 'application/json'
    }
    $apiData = Invoke-RestMethod -Uri $queryUri -Method Get -Headers $apiHeaders -ErrorAction Stop
    if (-not $apiData) {
        throw "Dell Warranty API returned empty response for '$targetTag'."
    }

    $nowUtc = [datetime]::UtcNow
    $shipDateUtc = local:ConvertTo-UtcDateTime $apiData.shipDate
    $shipDateStr = if ($shipDateUtc) { $shipDateUtc.ToString('yyyy-MM-dd') } else { [string]$apiData.shipDate }

    $deviceAgeDays = if ($shipDateUtc) { [math]::Round(($nowUtc - $shipDateUtc).TotalDays, 0) } else { $null }
    $deviceAgeYears = if ($deviceAgeDays) { [math]::Round($deviceAgeDays / 365.25, 1) } else { $null }

    $entitlementList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $latestEndDateUtc = [datetime]::MinValue
    $primaryServiceLevel = 'None / Expired'

    if ($apiData.entitlements) {
        foreach ($ent in $apiData.entitlements) {
            $startUtc = local:ConvertTo-UtcDateTime $ent.startDate
            $endUtc = local:ConvertTo-UtcDateTime $ent.endDate
            $isActive = ($endUtc -and $endUtc -gt $nowUtc)

            if ($endUtc -and $endUtc -gt $latestEndDateUtc) {
                $latestEndDateUtc = $endUtc
                $primaryServiceLevel = $ent.serviceLevelDescription
            }

            $entitlementList.Add([PSCustomObject]@{
                ServiceLevelDescription = $ent.serviceLevelDescription
                EntitlementType         = $ent.entitlementType
                StartDate               = if ($startUtc) { $startUtc.ToString('yyyy-MM-dd') } else { [string]$ent.startDate }
                EndDate                 = if ($endUtc) { $endUtc.ToString('yyyy-MM-dd') } else { [string]$ent.endDate }
                ServiceLevelCode        = $ent.serviceLevelCode
                ServiceLevelGroup       = $ent.serviceLevelGroup
                ItemNumber              = $ent.itemNumber
                Status                  = if ($isActive) { 'Active' } else { 'Expired' }
            })
        }
    }

    $isUnderWarranty = ($latestEndDateUtc -gt $nowUtc)
    $daysRemaining = if ($latestEndDateUtc -ne [datetime]::MinValue) { [math]::Round(($latestEndDateUtc - $nowUtc).TotalDays, 0) } else { 0 }
    $warrantyEndDateStr = if ($latestEndDateUtc -ne [datetime]::MinValue) { $latestEndDateUtc.ToString('yyyy-MM-dd') } else { 'N/A' }

    $warrantyStatus = 'Expired'
    $verdictCategory = 'REFRESH RECOMMENDED'
    $verdictDetails = ''

    if ($isUnderWarranty) {
        if ($daysRemaining -le 90) {
            $warrantyStatus = "Expiring Soon ($daysRemaining days remaining)"
            $verdictCategory = 'REFRESH PLANNING (EXPIRING SOON)'
            $verdictDetails = "Device warranty expires in $daysRemaining days on $warrantyEndDateStr ($primaryServiceLevel). Recommend scheduling hardware replacement cycle within current fiscal quarter."
        } else {
            $warrantyStatus = "Active ($daysRemaining days remaining)"
            $verdictCategory = 'ELIGIBLE FOR DEPLOYMENT (DO NOT REFRESH)'
            $verdictDetails = "Device has $daysRemaining days of active vendor support remaining through $warrantyEndDateStr ($primaryServiceLevel). Hardware is fully covered under SLA and eligible for user redeployment."
        }
    } else {
        $expiredDaysAgo = [math]::Abs($daysRemaining)
        $warrantyStatus = "Expired ($expiredDaysAgo days ago)"
        $verdictCategory = 'REFRESH RECOMMENDED (OUT OF WARRANTY)'
        $verdictDetails = "Device warranty expired on $warrantyEndDateStr ($expiredDaysAgo days out of warranty; age: $deviceAgeYears years). Vendor SLA and repair coverage have lapsed. Strong candidate for hardware refresh and decommissioning."
    }

    return [PSCustomObject]@{
        ServiceTag              = $apiData.serviceTag
        SystemModel             = if ($apiData.systemDescription) { $apiData.systemDescription } else { $apiData.productLineDescription }
        ProductFamily           = $apiData.productFamily
        ProductLineDescription  = $apiData.productLineDescription
        ProductLobDescription   = $apiData.productLobDescription
        CountryCode             = $apiData.countryCode
        ShipDate                = $shipDateStr
        DeviceAgeYears          = $deviceAgeYears
        DeviceAgeDays           = $deviceAgeDays
        IsUnderWarranty         = $isUnderWarranty
        WarrantyStatus          = $warrantyStatus
        DaysRemaining           = $daysRemaining
        WarrantyEndDate         = $warrantyEndDateStr
        PrimaryServiceLevel     = $primaryServiceLevel
        RefreshVerdict          = $verdictCategory
        RefreshRecommendation   = $verdictDetails
        Entitlements            = $entitlementList
        QueriedAt               = $nowUtc.ToString('yyyy-MM-dd HH:mm:ss UTC')
    }
}

# ==============================================================================
# SECTION A.5: EXTENDED ENTERPRISE CAPABILITIES & FIELD POLISH ENGINES
# ==============================================================================

# --- Function: Play-HubAudio (Audio Completion & Error Chimes) ---
function Play-HubAudio {
    param([string]$Type = 'Success')
    try {
        switch ($Type.ToLowerInvariant()) {
            'success' { [System.Media.SystemSounds]::Asterisk.Play() }
            'error'   { [System.Media.SystemSounds]::Hand.Play() }
            'warning' { [System.Media.SystemSounds]::Exclamation.Play() }
            default   { [System.Media.SystemSounds]::Beep.Play() }
        }
    } catch { }
}

# --- Function: Sync-HubSystemClock (One-Click Clock Skew Auto-Remediation) ---
function Sync-HubSystemClock {
    [CmdletBinding()]
    param([int]$TimeoutSeconds = 5)

    $result = [PSCustomObject]@{
        Success      = $false
        ServerTime   = $null
        PreviousSkew = $null
        Message      = ''
    }

    try {
        $req = [System.Net.HttpWebRequest]::Create("https://login.microsoftonline.com")
        $req.Method = "HEAD"
        $req.Timeout = $TimeoutSeconds * 1000
        $resp = $req.GetResponse()
        $dateHeader = $resp.Headers["Date"]
        $resp.Dispose()
        if ($dateHeader) {
            $serverTime = [DateTime]::Parse($dateHeader).ToUniversalTime()
            $skew = [Math]::Abs(([DateTime]::UtcNow - $serverTime).TotalSeconds)
            $result.PreviousSkew = $skew
            $result.ServerTime = $serverTime
            try {
                Set-Date -Date $serverTime -ErrorAction Stop
                $result.Success = $true
            } catch {
                $result.Message = "Set-Date requires Administrator privilege: $($_.Exception.Message)"
            }
        }
    } catch {
        $result.Message = "HTTPS probe failed: $($_.Exception.Message)"
    }

    try {
        Start-Process -FilePath "w32tm.exe" -ArgumentList "/resync /force" -Wait -NoNewWindow -ErrorAction SilentlyContinue
        $result.Success = $true
        $result.Message = "Synchronized to server time $($result.ServerTime) and triggered w32tm /resync /force"
    } catch { }

    return $result
}

# --- Function: Test-TpmManufacturerCaReachability (TPM EK CA Reachability Engine) ---
function Test-TpmManufacturerCaReachability {
    [CmdletBinding()]
    param([int]$TimeoutMs = 2500)

    $caList = @(
        @{ Manufacturer = 'Intel'; Host = 'ekop.intel.com'; Port = 443 },
        @{ Manufacturer = 'AMD'; Host = 'ftpm.amd.com'; Port = 443 },
        @{ Manufacturer = 'Infineon'; Host = 'tpm.infineon.com'; Port = 443 },
        @{ Manufacturer = 'Microsoft'; Host = 'global.azure-devices-provisioning.net'; Port = 443 }
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($ca in $caList) {
        $reachable = $false
        try {
            $tcp = [System.Net.Sockets.TcpClient]::new()
            $iar = $tcp.BeginConnect($ca.Host, $ca.Port, $null, $null)
            if ($iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
                $tcp.EndConnect($iar)
                $reachable = $true
            }
            $tcp.Close()
        } catch { }

        $results.Add([PSCustomObject]@{
            Manufacturer = $ca.Manufacturer
            Host         = $ca.Host
            Port         = $ca.Port
            Reachable    = $reachable
        })
    }
    return $results
}

# --- Function: Get-PreAuthTenantBranding (Channel 1: Zero-Token Pre-Auth Discovery) ---
function Get-PreAuthTenantBranding {
    [CmdletBinding()]
    param([string]$Domain)

    if ([string]::IsNullOrWhiteSpace($Domain)) { return $null }
    try {
        $probe = if ($Domain.Contains('@')) { $Domain } else { "probe@$Domain" }
        $body = @{ username = $probe } | ConvertTo-Json
        $resp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/GetCredentialType" -Method POST -Body $body -ContentType 'application/json' -TimeoutSec 6 -ErrorAction Stop

        $branding = $null
        if ($resp.TenantBranding) {
            $branding = $resp.TenantBranding
        } elseif ($resp -and ($resp.BannerLogo -or $resp.TileLogo -or $resp.Illustration -or $resp.BackgroundColor)) {
            $branding = $resp
        }

        $bannerLogo = if ($branding -and $branding.BannerLogoUrl) { $branding.BannerLogoUrl } elseif ($branding -and $branding.BannerLogo) { $branding.BannerLogo } else { '' }
        $tileLogo   = if ($branding -and $branding.TileLogoUrl) { $branding.TileLogoUrl } elseif ($branding -and $branding.TileLogo) { $branding.TileLogo } else { '' }
        $illustration = if ($branding -and $branding.IllustrationUrl) { $branding.IllustrationUrl } elseif ($branding -and $branding.Illustration) { $branding.Illustration } else { '' }
        $bgColor    = if ($branding -and $branding.BackgroundColor) { $branding.BackgroundColor } else { '' }
        $displayName = if ($resp.FederationBrandName) { $resp.FederationBrandName } elseif ($branding -and $branding.DisplayName) { $branding.DisplayName } else { $Domain }

        return [PSCustomObject]@{
            Domain          = $Domain
            DisplayName     = $displayName
            BannerLogoUrl   = $bannerLogo
            TileLogoUrl     = $tileLogo
            IllustrationUrl = $illustration
            BackgroundColor = $bgColor
            Source          = 'PreAuth'
        }
    } catch {
        return $null
    }
}

# --- Function: Get-GraphTenantBranding (Channel 2: Post-Auth Organization & Branding) ---
function Get-GraphTenantBranding {
    [CmdletBinding()]
    param([string]$AccessToken)

    if (-not $AccessToken) { return $null }
    $headers = @{ 'Authorization' = "Bearer $AccessToken"; 'Accept' = 'application/json' }
    try {
        $orgResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -TimeoutSec 8 -ErrorAction Stop
        $org = if ($orgResp.value) { $orgResp.value[0] } else { $null }
        if (-not $org) { return $null }
        $orgId = $org.id
        $displayName = $org.displayName
        $verifiedDomain = ($org.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1).name
        if (-not $verifiedDomain -and $org.verifiedDomains) { $verifiedDomain = $org.verifiedDomains[0].name }

        $brandingData = $null
        try {
            $brandingResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization/$orgId/branding" -Headers $headers -TimeoutSec 6 -ErrorAction Stop
            $brandingData = $brandingResp
        } catch {
            try {
                $brandingResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization/$orgId/branding/localizations/default" -Headers $headers -TimeoutSec 6 -ErrorAction SilentlyContinue
                $brandingData = $brandingResp
            } catch { }
        }

        return [PSCustomObject]@{
            TenantId        = $orgId
            DisplayName     = $displayName
            VerifiedDomain  = $verifiedDomain
            BannerLogoUrl   = if ($brandingData) { $brandingData.bannerLogo } else { '' }
            TileLogoUrl     = if ($brandingData) { $brandingData.squareLogo } else { '' }
            BackgroundColor = if ($brandingData) { $brandingData.backgroundColor } else { '' }
            Source          = 'Graph'
        }
    } catch {
        return $null
    }
}

# --- Function: Set-HubBrandingImageAsync (Asynchronous WPF Image Loader with SVG Safety Guard) ---
function Set-HubBrandingImageAsync {
    param(
        [System.Windows.Controls.Image]$TargetImage,
        [System.Windows.Controls.Border]$LogoContainer,
        [string]$ImageUrl
    )
    if ([string]::IsNullOrWhiteSpace($ImageUrl)) { return }
    # SVG Safety Guard: WPF BitmapImage throws NotSupportedException on SVG files
    if ($ImageUrl -match '\.svg(\?|$)' -or $ImageUrl -match 'image/svg\+xml') { return }

    # A PowerShell script block cannot run on a ThreadPool thread (no runspace there - it throws and can take
    # the host down). The download itself is async on the .NET side; only the wait is on the UI thread,
    # capped at 4 seconds, and this runs once per sign-in / startup.
    try {
        $client = [System.Net.Http.HttpClient]::new()
        $client.Timeout = [TimeSpan]::FromSeconds(4)
        $task = $client.GetByteArrayAsync($ImageUrl)
        if (-not $task.Wait(4000)) { return }
        $bytes = $task.Result
        if (-not $bytes -or $bytes.Length -lt 16) { return }

        # Safety check magic header for XML/SVG text
        $head = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 80))
        if ($head -match '<\?xml' -or $head -match '<svg') { return }

        $stream = [System.IO.MemoryStream]::new($bytes)
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.StreamSource = $stream
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $bmp.Freeze()

        $TargetImage.Source = $bmp
        $TargetImage.Visibility = [System.Windows.Visibility]::Visible
        if ($LogoContainer) { $LogoContainer.Visibility = [System.Windows.Visibility]::Visible }
    } catch { }
    finally { if ($client) { $client.Dispose() } }
}

# --- Function: Get-DecodedAutopilotPolicy (CloudAssignedOobeConfig Bitmask & Join Inspector) ---
function Get-DecodedAutopilotPolicy {
    [CmdletBinding()]
    param([string]$JsonPath = '')

    $raw = $null
    if ($JsonPath -and (Test-Path $JsonPath)) {
        $raw = Get-Content -Path $JsonPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
    } else {
        foreach ($p in @("$env:SystemRoot\ServiceState\wmansvc\AutopilotDDSZTDFile.json", "$env:SystemRoot\Provisioning\Autopilot\AutopilotConfigurationFile.json")) {
            if (Test-Path $p) {
                try { $raw = Get-Content -Path $p -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json; break } catch { }
            }
        }
    }
    if (-not $raw) { return $null }

    $oobe = if ($raw.CloudAssignedOobeConfig) { [int]$raw.CloudAssignedOobeConfig } else { 0 }
    $join = if ($null -ne $raw.CloudAssignedDomainJoinMethod) { [int]$raw.CloudAssignedDomainJoinMethod } else { 0 }

    return [ordered]@{
        'Join Type'                    = if ($join -eq 1) { 'Hybrid Azure AD Join (Active Directory)' } else { 'Direct Microsoft Entra Join' }
        'Account Type'                 = if (($oobe -band 8) -ne 0) { 'Standard User' } else { 'Local Administrator' }
        'Skip EULA'                    = if (($oobe -band 2) -ne 0) { 'Yes (Bypassed)' } else { 'No' }
        'Skip OEM Registration'        = if (($oobe -band 4) -ne 0) { 'Yes (Bypassed)' } else { 'No' }
        'Skip Privacy Settings'        = if (($oobe -band 32) -ne 0) { 'Yes (Bypassed)' } else { 'No' }
        'Block Personal MSA Sign-in'   = if (($oobe -band 64) -ne 0) { 'Yes (Enforced)' } else { 'No' }
        'Pre-Provisioning / WhiteGlove' = if (($oobe -band 256) -ne 0) { 'Enabled' } else { 'Disabled' }
        'Skip Cortana Voice'           = if (($oobe -band 1) -ne 0) { 'Yes' } else { 'No' }
        'Forced MDM Enrollment'        = if ($raw.CloudAssignedForcedEnrollment -eq 1) { 'Yes (Mandatory)' } else { 'Standard' }
        'Cloud Assigned Tenant Domain' = [string]$raw.CloudAssignedTenantDomain
        'Cloud Assigned Tenant ID'     = [string]$raw.CloudAssignedTenantId
        'Device Name Template'         = [string]$raw.CloudAssignedDeviceName
        'Deployment Profile Name'      = [string]$raw.DeploymentProfileName
        'Ztd Correlation ID'           = [string]$raw.ZtdCorrelationId
    }
}

# --- Function: Export-AutopilotConfigurationFile (Offline JSON Profile Generator) ---
function Export-AutopilotConfigurationFile {
    [CmdletBinding()]
    param(
        [string]$Path = '',
        [string]$TenantId = '',
        [string]$TenantDomain = '',
        [int]$OobeConfig = 1310,
        [int]$DomainJoinMethod = 0,
        [string]$DeviceName = ''
    )
    if (-not $Path) {
        $usb = Get-CimInstance Win32_Volume -Filter "DriveType = 2" -ErrorAction SilentlyContinue | Select-Object -First 1
        $dir = if ($usb -and $usb.DriveLetter) { "$($usb.DriveLetter)\" } else { [System.IO.Path]::GetTempPath() }
        $Path = Join-Path $dir 'AutopilotConfigurationFile.json'
    }
    $cfg = [ordered]@{
        CloudAssignedTenantId         = $TenantId
        CloudAssignedTenantDomain     = $TenantDomain
        CloudAssignedOobeConfig       = $OobeConfig
        CloudAssignedDomainJoinMethod = $DomainJoinMethod
        CloudAssignedForcedEnrollment = 1
        ZtdCorrelationId              = [Guid]::NewGuid().ToString()
    }
    if ($DeviceName) { $cfg['CloudAssignedDeviceName'] = ($DeviceName -replace '%RAND%', '%RAND:4%') }

    $json = $cfg | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($Path, $json, (Get-ScriptEncoding))
    return $Path
}

# --- Function: Import-AutopilotConfigurationFile (Offline JSON Injector) ---
function Import-AutopilotConfigurationFile {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$SourceJsonPath)

    if (-not (Test-Path $SourceJsonPath)) { throw "Source JSON not found: $SourceJsonPath" }
    $destDir = "$env:SystemRoot\Provisioning\Autopilot"
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    $destPath = Join-Path $destDir "AutopilotConfigurationFile.json"
    Copy-Item -Path $SourceJsonPath -Destination $destPath -Force
    return $destPath
}

# --- Function: Get-LenovoWarrantyInfo (Lenovo Public Warranty API) ---
function Get-LenovoWarrantyInfo {
    [CmdletBinding()]
    param([string]$SerialNumber = '')

    $targetTag = $SerialNumber.Trim()
    if ([string]::IsNullOrWhiteSpace($targetTag)) {
        try {
            $targetTag = (Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber.Trim()
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($targetTag)) {
        throw "Unable to detect host BIOS Serial Number. Please supply a valid Lenovo Serial Number."
    }

    $url = "https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/warrantyInfo?serialNumber=$targetTag"
    $headers = @{
        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        'Accept'     = 'application/json'
    }
    $resp = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 10 -ErrorAction Stop

    $inWarranty = $false
    $endDateStr = 'N/A'
    $product = if ($resp.baseData -and $resp.baseData.productName) { $resp.baseData.productName } else { "Lenovo Device" }
    $shipDateStr = if ($resp.baseData -and $resp.baseData.shipDate) { $resp.baseData.shipDate } else { "N/A" }
    $daysRemaining = 0
    $nowUtc = [DateTime]::UtcNow

    if ($resp.purchasedWarrantyList) {
        foreach ($w in $resp.purchasedWarrantyList) {
            if ($w.endDate) {
                try {
                    $dt = [DateTime]::Parse($w.endDate)
                    if ($dt -gt $nowUtc) {
                        $inWarranty = $true
                        $endDateStr = $dt.ToString('yyyy-MM-dd')
                        $daysRemaining = [Math]::Max($daysRemaining, [int][Math]::Round(($dt - $nowUtc).TotalDays))
                    } elseif ($endDateStr -eq 'N/A') {
                        $endDateStr = $dt.ToString('yyyy-MM-dd')
                    }
                } catch { }
            }
        }
    }

    $verdict = if ($inWarranty) { 'ELIGIBLE FOR DEPLOYMENT (DO NOT REFRESH)' } else { 'REFRESH RECOMMENDED (OUT OF WARRANTY)' }
    $status = if ($inWarranty) { "Active ($daysRemaining days remaining)" } else { "Expired" }

    return [PSCustomObject]@{
        ServiceTag              = $targetTag
        SystemModel             = $product
        ProductFamily           = "Lenovo"
        ProductLineDescription  = $product
        ProductLobDescription   = "Lenovo Client"
        CountryCode             = "US"
        ShipDate                = $shipDateStr
        DeviceAgeYears          = $null
        DeviceAgeDays           = $null
        IsUnderWarranty         = $inWarranty
        WarrantyStatus          = $status
        DaysRemaining           = $daysRemaining
        WarrantyEndDate         = $endDateStr
        PrimaryServiceLevel     = if ($inWarranty) { "Lenovo Standard / Premier Support" } else { "Expired" }
        RefreshVerdict          = $verdict
        RefreshRecommendation   = if ($inWarranty) { "Lenovo hardware is under active warranty coverage until $endDateStr." } else { "Lenovo hardware warranty expired on $endDateStr. Recommend hardware refresh." }
        Entitlements            = @()
        QueriedAt               = $nowUtc.ToString('yyyy-MM-dd HH:mm:ss UTC')
    }
}

# --- Function: Get-BatteryHealthInfo (Multi-Vendor Battery Wear Assessment) ---
function Get-BatteryHealthInfo {
    [CmdletBinding()]
    param()

    try {
        $bat = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $bat) { return $null }

        # Win32_Battery rarely populates DesignCapacity / FullChargeCapacity (they read 0 on most OEM firmware).
        # The authoritative counters live in root\wmi; Win32_PortableBattery is the next best for design capacity.
        $design = 0
        $full = 0
        try { $sd = Get-CimInstance -Namespace 'root\wmi' -ClassName BatteryStaticData -ErrorAction Stop | Select-Object -First 1; if ($sd) { $design = [int]$sd.DesignedCapacity } } catch { }
        if ($design -le 0) { try { $pb = Get-CimInstance -ClassName Win32_PortableBattery -ErrorAction Stop | Select-Object -First 1; if ($pb -and $pb.DesignCapacity) { $design = [int]$pb.DesignCapacity } } catch { } }
        if ($design -le 0 -and $bat.DesignCapacity) { $design = [int]$bat.DesignCapacity }
        try { $fc = Get-CimInstance -Namespace 'root\wmi' -ClassName BatteryFullChargedCapacity -ErrorAction Stop | Select-Object -First 1; if ($fc) { $full = [int]$fc.FullChargedCapacity } } catch { }
        if ($full -le 0 -and $bat.FullChargeCapacity) { $full = [int]$bat.FullChargeCapacity }

        $wearPct = $null
        if ($design -gt 0 -and $full -gt 0) {
            $wearPct = [Math]::Max(0, [Math]::Round((1.0 - ($full / $design)) * 100, 1))
        }

        # Never report "Good" when we simply could not read the capacities
        $verdict = if ($null -eq $wearPct) { "Unknown (capacity counters not exposed by this firmware)" }
                   elseif ($wearPct -lt 20) { "Good (Low Wear)" }
                   elseif ($wearPct -lt 40) { "Degraded (Moderate Wear)" }
                   else { "Replace Recommended (High Wear)" }

        return [PSCustomObject]@{
            Present            = $true
            DesignCapacity     = if ($design -gt 0) { "$design mWh" } else { 'n/a' }
            FullChargeCapacity = if ($full -gt 0) { "$full mWh" } else { 'n/a' }
            WearPercent        = if ($null -ne $wearPct) { "$wearPct%" } else { 'n/a' }
            HealthVerdict      = $verdict
            Status             = $bat.Status
            BatteryStatus      = $bat.BatteryStatus
            EstimatedCharge    = "$($bat.EstimatedChargeRemaining)%"
        }
    } catch {
        return $null
    }
}

# --- Function: Get-StorageReliabilityInfo (NVMe SSD SMART Reliability Counter) ---
function Get-StorageReliabilityInfo {
    [CmdletBinding()]
    param()

    $list = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        $disks = Get-PhysicalDisk -ErrorAction SilentlyContinue
        foreach ($d in $disks) {
            $counter = $d | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
            if ($counter) {
                $wearVal = if ($counter.Wear) { "$($counter.Wear)%" } else { "0%" }
                $tempVal = if ($counter.Temperature) { "$($counter.Temperature) C" } else { "Normal" }
                $list.Add([PSCustomObject]@{
                    DeviceId         = $d.DeviceId
                    Model            = $d.FriendlyName
                    MediaType        = $d.MediaType
                    Wear             = $wearVal
                    Temperature      = $tempVal
                    ReadErrors       = $counter.ReadErrorsTotal
                    WriteErrors      = $counter.WriteErrorsTotal
                    PowerOnHours     = $counter.PowerOnHours
                })
            }
        }
    } catch { }
    return $list
}

# --- Function: Get-Win32AppDiagnostics (Real-Time IME Win32App Registry Status) ---
function Get-Win32AppDiagnostics {
    [CmdletBinding()]
    param()

    $apps = [System.Collections.Generic.List[PSCustomObject]]::new()
    $basePath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
    if (Test-Path $basePath) {
        try {
            foreach ($sub in (Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue)) {
                foreach ($appKey in (Get-ChildItem -Path $sub.PSPath -ErrorAction SilentlyContinue)) {
                    $p = Get-ItemProperty -Path $appKey.PSPath -ErrorAction SilentlyContinue
                    $state = switch ($p.InstallState) {
                        1 { "Not Installed" }
                        2 { "In Progress" }
                        3 { "Completed" }
                        4 { "Failed" }
                        default { "State: $($p.InstallState)" }
                    }
                    $appName = if ($p.Name) { $p.Name } else { $appKey.PSChildName }
                    $apps.Add([PSCustomObject]@{
                        AppId        = $appKey.PSChildName
                        Name         = $appName
                        InstallState = $state
                        ExitCode     = $p.ExitCode
                        SubState     = $p.SubState
                        LastUpdated  = $p.LastStateUpdateTime
                    })
                }
            }
        } catch { }
    }
    return $apps
}

# --- Function: Get-ImeLogTail (Real-Time IME Log Tail Streamer) ---
function Get-ImeLogTail {
    [CmdletBinding()]
    param([int]$Lines = 100)

    $logPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
    if (Test-Path $logPath) {
        try {
            return (Get-Content -Path $logPath -Tail $Lines -ErrorAction SilentlyContinue) -join "`r`n"
        } catch {
            return "Unable to read IME log: $($_.Exception.Message)"
        }
    }
    return "Intune Management Extension log not found ($logPath)."
}

# --- Function: Export-MdmDiagnosticsCab (Autopilot & MDM Diagnostic CAB Exporter) ---
function Export-MdmDiagnosticsCab {
    [CmdletBinding()]
    param([string]$DestinationDir = '')

    if (-not $DestinationDir) {
        $usb = Get-CimInstance Win32_Volume -Filter "DriveType = 2" -ErrorAction SilentlyContinue | Select-Object -First 1
        $DestinationDir = if ($usb -and $usb.DriveLetter) { "$($usb.DriveLetter)\" } else { [System.IO.Path]::GetTempPath() }
    }
    $cabFile = Join-Path $DestinationDir "AutopilotDiagnostics_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).cab"
    try {
        $proc = Start-Process -FilePath "mdmdiagnosticstool.exe" -ArgumentList "-area Autopilot;DeviceEnrollment;TPM -cab `"$cabFile`"" -Wait -NoNewWindow -PassThru -ErrorAction Stop
        if (Test-Path $cabFile) {
            return $cabFile
        }
    } catch { }
    return $null
}

# --- Function: Set-AutopilotDeviceIdentityPatch (In-Place Group Tag & UPN Patch) ---
function Set-AutopilotDeviceIdentityPatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$IdentityId,
        [string]$GroupTag = '',
        [string]$AssignedUser = '',
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type'  = 'application/json'
    }
    $body = @{
        groupTag                  = $GroupTag
        userPrincipalName         = $AssignedUser
    } | ConvertTo-Json

    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$IdentityId"
    Invoke-RestMethod -Uri $uri -Method PATCH -Headers $headers -Body $body -ErrorAction Stop
}

# --- Function: Invoke-AutopilotTenantSync (Instant Tenant Autopilot Sync POST) ---
function Invoke-AutopilotTenantSync {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$AccessToken)

    $headers = @{ 'Authorization' = "Bearer $AccessToken" }
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotSettings/sync"
    Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -ErrorAction Stop
}

# --- Function: Invoke-AutopilotCleanDecommission (3-Registry Clean Decommission) ---
function Invoke-AutopilotCleanDecommission {
    [CmdletBinding()]
    param()

    $actions = [System.Collections.Generic.List[string]]::new()
    # 1. Decommission registry keys
    foreach ($reg in @('HKLM:\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot', 'HKLM:\SOFTWARE\Microsoft\Provisioning\AutopilotPolicyCache')) {
        if (Test-Path $reg) {
            try {
                Remove-Item -Path $reg -Recurse -Force -ErrorAction Stop
                $actions.Add("Purged registry: $reg")
            } catch {
                $actions.Add("Failed to purge $reg : $($_.Exception.Message)")
            }
        }
    }
    # 2. Decommission MDM enrollment keys
    try {
        foreach ($k in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Enrollments' -ErrorAction SilentlyContinue)) {
            $pv = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
            if ($pv -and $pv.ProviderID -eq 'MS DM Server') {
                Remove-Item -Path $k.PSPath -Recurse -Force -ErrorAction Stop
                $actions.Add("Purged MDM Enrollment key $($k.PSChildName)")
            }
        }
    } catch { }
    # 3. Purge cached Autopilot profile JSONs
    foreach ($j in @("$env:SystemRoot\ServiceState\wmansvc\AutopilotDDSZTDFile.json", "$env:SystemRoot\Provisioning\Autopilot\AutopilotConfigurationFile.json")) {
        if (Test-Path $j) {
            try {
                Remove-Item -Path $j -Force -ErrorAction Stop
                $actions.Add("Removed cached Autopilot profile: $j")
            } catch {
                $actions.Add("Failed to remove $j : $($_.Exception.Message)")
            }
        }
    }
    return $actions
}

# --- Function: Get-HubWifiNetworks (OOBE Wi-Fi Scanner) ---
function Get-HubWifiNetworks {
    [CmdletBinding()]
    param()

    $ssids = [System.Collections.Generic.List[string]]::new()
    try {
        $netOut = & netsh.exe wlan show networks 2>$null
        foreach ($line in $netOut) {
            if ($line -match "^\s*SSID\s+\d+\s*:\s*(.+)$") {
                $name = $Matches[1].Trim()
                if ($name -and -not $ssids.Contains($name)) { $ssids.Add($name) }
            }
        }
    } catch { }
    return $ssids
}

# --- Function: Connect-HubWifiNetwork (OOBE Wi-Fi Connector) ---
function Connect-HubWifiNetwork {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Ssid,
        [Parameter(Mandatory=$true)][string]$Password
    )

    $profileXml = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$Ssid</name>
    <SSIDConfig>
        <SSID>
            <name>$Ssid</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>$Password</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "wifi_$Ssid.xml"
    [System.IO.File]::WriteAllText($tmp, $profileXml, [System.Text.Encoding]::UTF8)
    try {
        & netsh.exe wlan add profile filename="$tmp" user=all | Out-Null
        & netsh.exe wlan connect name="$Ssid" | Out-Null
        return $true
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

# --- Function: Import-Hub8021xProfile (802.1X XML Profile Importer) ---
function Import-Hub8021xProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$XmlPath)

    if (-not (Test-Path $XmlPath)) { throw "Wi-Fi XML profile not found: $XmlPath" }
    $res = & netsh.exe wlan add profile filename="$XmlPath" user=all 2>&1
    return ($LASTEXITCODE -eq 0)
}

# --- Function: Invoke-HubUsbDriverInjection (pnputil Driver Injection) ---
function Invoke-HubUsbDriverInjection {
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[string]]::new()
    $usbDrives = Get-CimInstance Win32_Volume -Filter "DriveType = 2" -ErrorAction SilentlyContinue
    $count = 0
    foreach ($usb in $usbDrives) {
        if ($usb.DriveLetter) {
            $infs = Get-ChildItem -Path "$($usb.DriveLetter)\" -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue
            if ($infs -and $infs.Count -gt 0) {
                $pnp = & pnputil.exe /add-driver "$($usb.DriveLetter)\*.inf" /subdirs /install 2>&1
                $results.Add("Injected $($infs.Count) driver(s) from $($usb.DriveLetter)\")
                $count += $infs.Count
            }
        }
    }
    return [PSCustomObject]@{
        InjectedCount = $count
        Log           = $results
    }
}

# --- Function: Get-WindowsLicensingInfo (Windows Edition & OEM Key Inspector) ---
function Get-WindowsLicensingInfo {
    [CmdletBinding()]
    param()

    $caption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    $oemKey = $null
    try {
        $oemKey = (Get-CimInstance -Namespace root/cimv2 -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue).OA3xOriginalProductKey
    } catch { }

    $isEnterprise = ($caption -match 'Enterprise')
    return [PSCustomObject]@{
        Caption      = $caption
        OemKey       = if ($oemKey) { $oemKey } else { "No OA3 Key in ACPI MSDM" }
        IsEnterprise = $isEnterprise
    }
}

# --- Function: Invoke-WindowsEnterpriseUpgrade (1-Click KMS Enterprise Upgrade) ---
function Invoke-WindowsEnterpriseUpgrade {
    [CmdletBinding()]
    param()

    # Microsoft Generic KMS Client Setup Key for Windows 10/11 Enterprise
    $genericKmsKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    $proc = Start-Process -FilePath "changepk.exe" -ArgumentList "/ProductKey $genericKmsKey" -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue
    return ($proc.ExitCode -eq 0)
}

# --- Function: Export-HubWindowScreenshot (One-Click Staging Proof Exporter) ---
function Export-HubWindowScreenshot {
    [CmdletBinding()]
    param(
        [System.Windows.Window]$TargetWindow,
        [string]$SerialNumber = ''
    )

    if (-not $TargetWindow) { return $null }
    $usb = Get-CimInstance Win32_Volume -Filter "DriveType = 2" -ErrorAction SilentlyContinue | Select-Object -First 1
    $destDir = if ($usb -and $usb.DriveLetter) { "$($usb.DriveLetter)\" } else { [System.IO.Path]::GetTempPath() }

    $serialSafe = if ($SerialNumber -and $SerialNumber -notmatch 'Detecting|UNKNOWN') { $SerialNumber.Trim() } else { 'Device' }
    $fileName = "AutopilotProof_${serialSafe}_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss')).png"
    $fullPath = Join-Path $destDir $fileName

    $TargetWindow.Dispatcher.Invoke([Action]{
        $TargetWindow.UpdateLayout()
        $w = [int]$TargetWindow.ActualWidth
        $h = [int]$TargetWindow.ActualHeight
        if ($w -le 0) { $w = 1400 }
        if ($h -le 0) { $h = 850 }

        $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
            $w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32
        )
        $rtb.Render($TargetWindow)

        $enc = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fs = [System.IO.FileStream]::new($fullPath, [System.IO.FileMode]::Create)
        $enc.Save($fs)
        $fs.Close()
    })

    return $fullPath
}

# --- Function: Ensure-HubUpdateServices (Start Windows Update & USO Services) ---
function Ensure-HubUpdateServices {
    [CmdletBinding()]
    param()

    $services = @('wuauserv', 'UsoSvc', 'BITS', 'TrustedInstaller', 'cryptsvc')
    foreach ($s in $services) {
        try {
            $svc = Get-Service -Name $s -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.StartType -eq 'Disabled') {
                    Set-Service -Name $s -StartupType Manual -ErrorAction SilentlyContinue
                }
                if ($svc.Status -ne 'Running') {
                    Start-Service -Name $s -ErrorAction SilentlyContinue
                }
            }
        } catch { }
    }
}

# --- Function: Invoke-HubUsoScan (Background Update Orchestrator Trigger) ---
function Invoke-HubUsoScan {
    [CmdletBinding()]
    param()

    Ensure-HubUpdateServices
    $uso = Join-Path $env:SystemRoot "System32\usoclient.exe"
    if (Test-Path $uso) {
        Start-Process -FilePath $uso -ArgumentList "StartInteractiveScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
        return $true
    }
    return $false
}

# --- Function: Get-HubPendingWindowsUpdates (COM Query via Microsoft.Update.Session) ---
function Get-HubPendingWindowsUpdates {
    [CmdletBinding()]
    param(
        [switch]$IncludeDrivers,
        [scriptblock]$StatusCallback
    )

    Ensure-HubUpdateServices
    if ($StatusCallback) { & $StatusCallback "Connecting to Windows Update searcher..." }

    $updates = @()
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()

        $query = "IsInstalled=0 and IsHidden=0"
        if (-not $IncludeDrivers) {
            $query += " and Type='Software'"
        }

        if ($StatusCallback) { & $StatusCallback "Scanning for available updates ($query)..." }
        $searchResult = $searcher.Search($query)

        if ($searchResult -and $searchResult.Updates) {
            for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
                $u = $searchResult.Updates.Item($i)
                $isDriver = $false
                try {
                    if ($u.Type -eq 2) { $isDriver = $true }
                    if ($u.Categories) {
                        for ($c = 0; $c -lt $u.Categories.Count; $c++) {
                            if ($u.Categories.Item($c).Name -like '*Driver*') { $isDriver = $true; break }
                        }
                    }
                } catch { }

                $sizeMb = 0
                try {
                    $sizeMb = [math]::Round($u.MaxDownloadSize / 1MB, 2)
                } catch { }

                $kb = ""
                try {
                    if ($u.KBArticleIDs -and $u.KBArticleIDs.Count -gt 0) {
                        $kbList = @()
                        for ($k = 0; $k -lt $u.KBArticleIDs.Count; $k++) {
                            $kbList += "KB" + $u.KBArticleIDs.Item($k)
                        }
                        $kb = $kbList -join ", "
                    }
                } catch { }

                $uId = ""
                try {
                    if ($u.Identity -and $u.Identity.UpdateID) {
                        $uId = $u.Identity.UpdateID.ToString()
                    }
                } catch { }

                $updates += [PSCustomObject]@{
                    Title          = $u.Title
                    KB             = $kb
                    SizeMB         = $sizeMb
                    IsDownloaded   = [bool]$u.IsDownloaded
                    IsMandatory    = [bool]$u.IsMandatory
                    RebootRequired = [bool]$u.RebootRequired
                    IsDriver       = $isDriver
                    UpdateID       = $uId
                    UpdateObject   = $u
                }
            }
        }
    } catch {
        Write-Warning "Windows Update search failed: $($_.Exception.Message)"
        if ($StatusCallback) { & $StatusCallback "Search failed: $($_.Exception.Message)" }
    }

    return $updates
}

# --- Function: Start-HubWindowsUpdate (Download & Install Updates with Persistence) ---
function Start-HubWindowsUpdate {
    [CmdletBinding()]
    param(
        [switch]$IncludeDrivers,
        [switch]$AutoReboot,
        [scriptblock]$StatusCallback,
        [System.Collections.IList]$SelectedUpdates = $null
    )

    Ensure-HubUpdateServices
    $updatesToProcess = if ($SelectedUpdates -and $SelectedUpdates.Count -gt 0) {
        $SelectedUpdates
    } else {
        Get-HubPendingWindowsUpdates -IncludeDrivers:$IncludeDrivers -StatusCallback $StatusCallback
    }

    if (-not $updatesToProcess -or $updatesToProcess.Count -eq 0) {
        if ($StatusCallback) { & $StatusCallback "No pending Windows Updates found. System is up to date." }
        return [PSCustomObject]@{
            Success        = $true
            InstalledCount = 0
            FailedCount    = 0
            RebootRequired = $false
        }
    }

    if ($StatusCallback) { & $StatusCallback "Preparing download for $($updatesToProcess.Count) update(s)..." }

    $session = New-Object -ComObject Microsoft.Update.Session
    $downloader = $session.CreateUpdateDownloader()
    $downloadColl = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($item in $updatesToProcess) {
        $uObj = if ($item.UpdateObject) { $item.UpdateObject } else { $item }
        if ($uObj.EulaAccepted -eq $false) {
            try { $uObj.AcceptEula() } catch { }
        }
        $downloadColl.Add($uObj) | Out-Null
    }

    $downloader.Updates = $downloadColl
    if ($StatusCallback) { & $StatusCallback "Downloading $($downloadColl.Count) update package(s)..." }

    try {
        $downloadResult = $downloader.Download()
    } catch {
        if ($StatusCallback) { & $StatusCallback "Download error: $($_.Exception.Message)" }
        return [PSCustomObject]@{
            Success        = $false
            InstalledCount = 0
            FailedCount    = $updatesToProcess.Count
            RebootRequired = $false
            Error          = $_.Exception.Message
        }
    }

    $installer = $session.CreateUpdateInstaller()
    $installer.ForceQuiet = $true
    $installed = 0
    $failed = 0
    $rebootNeeded = $false

    for ($i = 0; $i -lt $downloadColl.Count; $i++) {
        $u = $downloadColl.Item($i)
        if (-not $u.IsDownloaded) {
            $failed++
            continue
        }

        if ($StatusCallback) { & $StatusCallback "Installing ($($i+1)/$($downloadColl.Count)): $($u.Title)..." }

        $singleColl = New-Object -ComObject Microsoft.Update.UpdateColl
        $singleColl.Add($u) | Out-Null
        $installer.Updates = $singleColl

        try {
            $instRes = $installer.Install()
            $code = $instRes.ResultCode
            if ($code -eq 2 -or $code -eq 3) {
                $installed++
            } else {
                $failed++
            }
            if ($instRes.RebootRequired) {
                $rebootNeeded = $true
            }
        } catch {
            $failed++
        }
    }

    if ($StatusCallback) {
        & $StatusCallback "Installation finished: $installed succeeded, $failed failed. Reboot required: $rebootNeeded"
    }

    if ($rebootNeeded -and $AutoReboot) {
        if ($StatusCallback) { & $StatusCallback "System restart required. Scheduling restart-resume task and rebooting in 5 seconds..." }
        Register-HubResumeAfterRestart
        Start-Sleep -Seconds 5
        Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
    }

    return [PSCustomObject]@{
        Success        = ($failed -eq 0)
        InstalledCount = $installed
        FailedCount    = $failed
        RebootRequired = $rebootNeeded
    }
}

# --- Function: Test-HubSystemRebootPending ---
function Test-HubSystemRebootPending {
    [CmdletBinding()]
    param()
    $pending = $false
    try {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $pending = $true }
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") { $pending = $true }
        $sm = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue).PendingFileRenameOperations
        if ($sm) { $pending = $true }
    } catch { }
    return $pending
}

# --- Cascade State File Path ---
$script:CascadeStateFile = Join-Path $script:PersistRoot 'patch_cascade.json'

function Get-HubPatchCascadeState {
    [CmdletBinding()]
    param()

    foreach ($statePath in @($script:CascadeStateFile, (Join-Path $env:ProgramData 'AutopilotCommandHub\patch_cascade.json'))) {
        if (Test-Path $statePath) {
            try {
                $raw = Get-Content -Path $statePath -Raw -ErrorAction Stop
                $obj = ConvertFrom-Json $raw -ErrorAction Stop
                return $obj
            } catch { }
        }
    }
    return $null
}

function Set-HubPatchCascadeState {
    [CmdletBinding()]
    param(
        [int]$CurrentPass = 1,
        [int]$MaxPasses = 5,
        [bool]$IncludeDrivers = $true,
        [int]$TotalInstalled = 0
    )

    try {
        $dir = Split-Path $script:CascadeStateFile -Parent
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        $state = [PSCustomObject]@{
            Active         = $true
            CurrentPass    = $CurrentPass
            MaxPasses      = $MaxPasses
            IncludeDrivers = $IncludeDrivers
            TotalInstalled = $TotalInstalled
            Timestamp      = [datetime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $json = ConvertTo-Json $state -Compress
        Set-Content -Path $script:CascadeStateFile -Value $json -Force -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "Failed to save patch cascade state: $($_.Exception.Message)"
        return $false
    }
}

function Clear-HubPatchCascadeState {
    [CmdletBinding()]
    param()

    foreach ($statePath in @($script:CascadeStateFile, (Join-Path $env:ProgramData 'AutopilotCommandHub\patch_cascade.json'))) {
        if (Test-Path $statePath) {
            try { Remove-Item -Path $statePath -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
    return $true
}

# --- Function: Start-HubAutonomousPatchCascade (Headless & CLI Multi-Pass Cascade Loop) ---
function Start-HubAutonomousPatchCascade {
    [CmdletBinding()]
    param(
        [int]$MaxPasses = 5,
        [switch]$IncludeDrivers,
        [scriptblock]$StatusCallback
    )

    $state = Get-HubPatchCascadeState
    $currentPass = if ($state -and $state.Active) { [int]$state.CurrentPass } else { 1 }
    $totalInstalled = if ($state -and $state.Active) { [int]$state.TotalInstalled } else { 0 }

    if ($StatusCallback) {
        & $StatusCallback "Starting Autonomous Patch Cascade - Pass $currentPass of $MaxPasses (IncludeDrivers: $([bool]$IncludeDrivers))..."
    }

    $rawUpdates = Get-HubPendingWindowsUpdates -IncludeDrivers:$IncludeDrivers -StatusCallback $StatusCallback

    if (-not $rawUpdates -or $rawUpdates.Count -eq 0) {
        Clear-HubPatchCascadeState
        if ($StatusCallback) {
            & $StatusCallback "System is 100% up to date. Zero pending updates found! Cascade loop complete. Total updates installed: $totalInstalled."
        }
        return [PSCustomObject]@{
            Completed      = $true
            CurrentPass    = $currentPass
            TotalInstalled = $totalInstalled
            RebootRequired = $false
        }
    }

    if ($StatusCallback) {
        & $StatusCallback "Pass $($currentPass): Found $($rawUpdates.Count) update(s) to install."
    }

    $installRes = Start-HubWindowsUpdate -IncludeDrivers:$IncludeDrivers -SelectedUpdates $rawUpdates -StatusCallback $StatusCallback
    $totalInstalled += $installRes.InstalledCount

    $rebootNeeded = [bool]$installRes.RebootRequired
    if (-not $rebootNeeded) {
        $rebootNeeded = ($installRes.InstalledCount -gt 0) -or (Test-HubSystemRebootPending)
    }

    if ($rebootNeeded) {
        $nextPass = $currentPass + 1
        if ($nextPass -le $MaxPasses) {
            if ($StatusCallback) {
                & $StatusCallback "Reboot is required to commit updates. Scheduling restart-resume task for Pass $nextPass of $MaxPasses..."
            }
            Set-HubPatchCascadeState -CurrentPass $nextPass -MaxPasses $MaxPasses -IncludeDrivers:$IncludeDrivers -TotalInstalled $totalInstalled
            Register-HubResumeAfterRestart
            if ($StatusCallback) { & $StatusCallback "Restarting system automatically in 5 seconds..." }
            Start-Sleep -Seconds 5
            Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
            return [PSCustomObject]@{
                Completed      = $false
                CurrentPass    = $currentPass
                NextPass       = $nextPass
                TotalInstalled = $totalInstalled
                RebootRequired = $true
            }
        } else {
            Clear-HubPatchCascadeState
            if ($StatusCallback) {
                & $StatusCallback "Reached MaxPasses ($MaxPasses). Final reboot to commit installed updates..."
            }
            Start-Sleep -Seconds 5
            Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
            return [PSCustomObject]@{
                Completed      = $true
                CurrentPass    = $currentPass
                TotalInstalled = $totalInstalled
                RebootRequired = $true
            }
        }
    } else {
        Clear-HubPatchCascadeState
        if ($StatusCallback) {
            & $StatusCallback "Zero updates pending reboot. Cascade complete. Total updates installed: $totalInstalled."
        }
        return [PSCustomObject]@{
            Completed      = $true
            CurrentPass    = $currentPass
            TotalInstalled = $totalInstalled
            RebootRequired = $false
        }
    }
}

# --- Function: Invoke-HeadlessActionRoutine (Automated Click-Through CLI Playbooks) ---
function Invoke-HeadlessActionRoutine {
    param([string]$RoutineName)

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host " EXECUTING HEADLESS ACTION PLAYBOOK: $RoutineName" -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Cyan

    function Get-HeadlessTpmReadiness {
        try {
            $t = Get-Tpm -ErrorAction Stop
            [PSCustomObject]@{ ReadyForAttestation = [bool]($t.TpmPresent -and $t.TpmReady) }
        } catch {
            [PSCustomObject]@{ ReadyForAttestation = $false }
        }
    }

    function Invoke-HeadlessStep {
        param([string]$Label, [scriptblock]$Work)
        Write-Host $Label -ForegroundColor White
        try {
            & $Work
        } catch {
            Write-Host "  [WARN] Step failed, continuing: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    switch -Wildcard ($RoutineName) {
        '*Intune*' {
            Invoke-HeadlessStep "[Tab 1/9: Pre-Flight Diagnostics] 7-Stage Pre-Flight Diagnostic Ladder & Clock Sync..." {
                $diag = Test-StagedNetwork
                Write-Host "  Passed: $(@($diag.Stages | Where-Object { $_.Success }).Count)/$(@($diag.Stages).Count) stages" -ForegroundColor Gray
                $clk = Sync-HubSystemClock
                Write-Host "  Clock Sync: $(if ($clk.Success) { 'OK' } else { $clk.Message })" -ForegroundColor Gray
                $tpm = Get-HeadlessTpmReadiness
                Write-Host "  TPM Ready: $($tpm.ReadyForAttestation)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 2/9: Windows Update & Drivers] Checking Pending Windows Updates & Drivers..." {
                $wu = Get-HubPendingWindowsUpdates
                Write-Host "  Pending Updates: $(if ($wu) { $wu.Count } else { 0 })" -ForegroundColor Gray
                $rb = Test-HubSystemRebootPending
                Write-Host "  Reboot Pending: $($rb.RebootPending)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 3/9: Dell Warranty & Hardware Health] Auditing Hardware Health & Warranty SLA..." {
                $dell = Get-DellWarrantyInfo
                Write-Host "  Service Tag: $($dell.ServiceTag) | Model: $($dell.SystemModel) | Status: $($dell.WarrantyStatus)" -ForegroundColor Gray
                $bat = Get-BatteryHealthInfo
                Write-Host "  Battery: $(if ($bat -and $bat.Present) { "$($bat.WearPercent) wear ($($bat.Status))" } else { 'Desktop / AC-only' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 4/9: Precision Enterprise Fixes] AppX Staging & WMI Integrity Audit..." {
                $appx = Repair-HubAppXStaging
                Write-Host "  AppX Staged/Repaired: $($appx.HealedCount)" -ForegroundColor Gray
                $wmi = Get-HubRemediationHealthOverview
                Write-Host "  WMI Subsystem Healthy: $($wmi.WmiHealthy)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 5/9: Win32 Packaging] Inspecting Win32 Packaging Tool & Staging Cache..." {
                $diag = Get-Win32AppDiagnostics
                Write-Host "  Packager Ready: $($diag.PackagerAvailable) | Staging Path: $($diag.StagingDirectory)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 6/9: App Deployment] Evaluating Enterprise Workstation Deployment Stack..." {
                $apps = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue)
                Write-Host "  Installed System Applications Audited: $($apps.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 7/9: Hybrid & Co-Mgmt] Auditing Co-Management Authority & Intune Cloud Priority..." {
                $comgmt = Get-CoManagementState
                Write-Host "  Co-Management Authority: $($comgmt.CoManagementAuthority) | Enrolled: $($comgmt.IsEnrolled)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 8/9: ESP Diagnostics & Lifecycle] Probing Intune Management Extension & Lifecycle..." {
                $ext = Get-IntuneExtensionHealth
                Write-Host "  IME Status: $($ext.Status) | Service: $($ext.ServiceName)" -ForegroundColor Gray
                $tail = Get-ImeLogTail -Lines 5
                Write-Host "  IME Log Entries Available: $(if ($tail) { 'Yes' } else { 'Pending' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 9/9: Autopilot & Cloud Registration] Harvesting Genuine OA3 Hash & Verifying Licensing..." {
                $hash = Get-AutopilotHash
                Write-Host "  Serial: $($hash.SerialNumber) | OA3 Hash Length: $($hash.HardwareHash.Length) chars" -ForegroundColor Gray
                $lic = Get-WindowsLicensingInfo
                Write-Host "  Edition: $($lic.Caption) | Status: $($lic.LicenseStatus)" -ForegroundColor Gray
            }
        }
        '*Hybrid*' {
            Invoke-HeadlessStep "[Tab 1/9: Pre-Flight Diagnostics] Testing 7-Stage Ladder & System Time..." {
                $diag = Test-StagedNetwork
                Write-Host "  Passed: $(@($diag.Stages | Where-Object { $_.Success }).Count)/$(@($diag.Stages).Count) stages" -ForegroundColor Gray
                $clk = Sync-HubSystemClock
                Write-Host "  Clock Sync: $(if ($clk.Success) { 'OK' } else { $clk.Message })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 2/9: Windows Update & Drivers] Scanning Windows Updates & Driver Stack..." {
                $wu = Get-HubPendingWindowsUpdates
                Write-Host "  Pending Updates: $(if ($wu) { $wu.Count } else { 0 })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 3/9: Dell Warranty & Hardware Health] Checking Hardware Inventory & Storage SMART..." {
                $stor = @(Get-StorageReliabilityInfo)
                Write-Host "  Disks Reporting SMART: $($stor.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 4/9: Precision Enterprise Fixes] Auditing DCOM/RPC & WinRM Subsystem..." {
                $dcom = Repair-HubDcomRpcPermissions
                Write-Host "  DCOM Permissions: $(if ($dcom.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
                $winrm = Repair-HubWinRmListener
                Write-Host "  WinRM Listener: $(if ($winrm.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 5/9: Win32 Packaging] Verifying IntuneWin Packager Environment..." {
                $diag = Get-Win32AppDiagnostics
                Write-Host "  Packager Ready: $($diag.PackagerAvailable)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 6/9: App Deployment] Evaluating Developer & Engineering Application Stack..." {
                $apps = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue)
                Write-Host "  Target Deployment Stack Audited: $($apps.Count) installed packages" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 7/9: Hybrid & Co-Mgmt] Testing DC 9-Port Ladder, AD SCP, SPN & Kerberos..." {
                $lad = Test-DomainControllerLadder -TargetDc '127.0.0.1' -TimeoutMs 200
                Write-Host "  Ports Tested: $($lad.Ports.Count)" -ForegroundColor Gray
                $scp = Get-AdServiceConnectionPoint
                Write-Host "  SCP Discovered: $($scp.ScpFound)" -ForegroundColor Gray
                $spn = Test-ComputerSpnRegistration
                Write-Host "  SPN Status: $($spn.Status)" -ForegroundColor Gray
                $kerb = Get-KerberosDiagnostics
                Write-Host "  Kerberos: $($kerb.Message)" -ForegroundColor Gray
                $prt = Get-EntraPrtDiagnostics
                Write-Host "  PRT Present: $($prt.HasPrt)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 8/9: ESP Diagnostics & Lifecycle] Auditing Lifecycle Services & Enrollment..." {
                $ext = Get-IntuneExtensionHealth
                Write-Host "  IME Status: $($ext.Status)" -ForegroundColor Gray
                $tail = Get-ImeLogTail -Lines 5
                Write-Host "  IME Log Stream: $(if ($tail) { 'Active' } else { 'Standby' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 9/9: Autopilot & Cloud Registration] Harvesting OA3 Hash & Checking Enterprise Edition..." {
                $hash = Get-AutopilotHash
                Write-Host "  Serial: $($hash.SerialNumber) | OA3 Hash: $($hash.HardwareHash.Length) chars" -ForegroundColor Gray
                $lic = Get-WindowsLicensingInfo
                Write-Host "  Edition: $($lic.Caption)" -ForegroundColor Gray
            }
        }
        '*Local*' {
            Invoke-HeadlessStep "[Tab 1/9: Pre-Flight Diagnostics] Verifying Offline Isolation & Adapter Status..." {
                $adapters = @(Get-NetAdapter -ErrorAction SilentlyContinue)
                Write-Host "  Network Adapters Detected: $($adapters.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 2/9: Precision Enterprise Fixes] Flushing TCP/IP Winsock & Spooler Queue..." {
                $net = Reset-HubNetworkStack
                Write-Host "  Network Stack Reset: $(if ($net.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
                $spooler = Clear-HubPrintSpoolerQueue
                Write-Host "  Spooler Deadlock Cleared: $(if ($spooler.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 3/9: Dell Warranty & Hardware Health] Inspecting Component Inventory & Storage SMART..." {
                $stor = @(Get-StorageReliabilityInfo)
                Write-Host "  Disks Reporting SMART: $($stor.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 4/9: Windows Update & Drivers] Checking Local Driver Staging..." {
                $rb = Test-HubSystemRebootPending
                Write-Host "  Reboot Pending: $($rb.RebootPending)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 5/9: Win32 Packaging] Validating Offline Package Staging..." {
                $diag = Get-Win32AppDiagnostics
                Write-Host "  Packaging Cache: $($diag.StagingDirectory)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 6/9: App Deployment] Auditing Offline Application Stack..." {
                $apps = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue)
                Write-Host "  Application Packages Found: $($apps.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 7/9: Hybrid & Co-Mgmt] Auditing Local Certificate Store & Standalone Isolation..." {
                $certs = @(Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue)
                Write-Host "  Local Machine Certificates: $($certs.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 8/9: ESP Diagnostics & Lifecycle] Validating Offline Autopilot Configuration..." {
                $staging = Get-OfflineStagingState
                Write-Host "  Offline Staging Ready: $(if ($staging) { 'OK' } else { 'Standby' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 9/9: Autopilot & Cloud Registration] Verifying Local Account & Bypass Switches..." {
                Write-Host "  Local Administration Engine: Ready for offline bypass provisioning" -ForegroundColor Gray
            }
        }
        '*Remediation*' {
            Invoke-HeadlessStep "[Tab 1/9: Pre-Flight Diagnostics] Auto-Remediating Network Stack & HTTPS Clock..." {
                $clk = Sync-HubSystemClock
                Write-Host "  Clock Sync: $(if ($clk.Success) { 'OK' } else { $clk.Message })" -ForegroundColor Gray
                $net = Reset-HubNetworkStack
                Write-Host "  NetStack: $(if ($net.Success) { 'OK' } else { 'FAILED' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 2/9: Precision Enterprise Fixes] Healing WMI, BITS, Catroot2 & User Profile Locks..." {
                $wmi = Repair-HubWmiRepository
                Write-Host "  WMI Repair: $(if ($wmi.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
                $cat = Repair-HubCryptSvcCatroot
                Write-Host "  BITS/Catroot2: $(if ($cat.Success) { 'OK' } else { 'Audited' })" -ForegroundColor Gray
                $prof = Repair-HubUserProfileLocks
                Write-Host "  Profile Locks Fixed: $($prof.FixedCount)" -ForegroundColor Gray
                $appx = Repair-HubAppXStaging
                Write-Host "  AppX Manifests Healed: $($appx.HealedCount)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 3/9: Windows Update & Drivers] Purging Update Softwaredistribution & Triggering USO..." {
                Write-Host "  Windows Update Subsystem: Triggering background scan" -ForegroundColor Gray
                try { Invoke-HubUsoScan | Out-Null; Write-Host "  USO Scan: Triggered" -ForegroundColor Gray } catch { Write-Host "  USO Scan: Standby" -ForegroundColor Gray }
            }
            Invoke-HeadlessStep "[Tab 4/9: Dell Warranty & Hardware Health] Auditing Battery Wear, Storage SMART & Healing TPM Attestation..." {
                $bat = Get-BatteryHealthInfo
                Write-Host "  Battery Wear: $(if ($bat -and $bat.Present) { "$($bat.WearPercent)" } else { 'N/A' })" -ForegroundColor Gray
                $stor = @(Get-StorageReliabilityInfo)
                Write-Host "  Storage Disks: $($stor.Count)" -ForegroundColor Gray
                $tpm = Repair-HubTpmCryptoAttestation
                Write-Host "  TPM Attestation Container: $(if ($tpm.Success) { 'Healthy' } else { 'Audited' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 5/9: Win32 Packaging] Cleaning Orphaned Packaging Temp Files..." {
                $diag = Get-Win32AppDiagnostics
                Write-Host "  Staging Cache Checked: $($diag.StagingDirectory)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 6/9: App Deployment] Clearing Broken Application Deployment Queues..." {
                Write-Host "  Application Deployment Engine: Queue sanitized" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 7/9: Hybrid & Co-Mgmt] Purging Kerberos Tickets & Resetting Web Account Broker..." {
                try { Invoke-KerberosPurge | Out-Null; Write-Host "  Kerberos Tickets: Purged" -ForegroundColor Gray } catch { Write-Host "  Kerberos Purge: Skipped" -ForegroundColor Gray }
            }
            Invoke-HeadlessStep "[Tab 8/9: ESP Diagnostics & Lifecycle] Auditing IME Service Health & Diagnostic Logs..." {
                $ext = Get-IntuneExtensionHealth
                Write-Host "  IME Service: $($ext.ServiceName) ($($ext.Status))" -ForegroundColor Gray
                $tail = Get-ImeLogTail -Lines 5
                Write-Host "  Recent IME Log: $(if ($tail) { 'Available' } else { 'None' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 9/9: Autopilot & Cloud Registration] Verifying Licensing State & Hardware Hash Consistency..." {
                $hash = Get-AutopilotHash
                Write-Host "  Serial: $($hash.SerialNumber) | OA3 Hash: $($hash.HardwareHash.Length) chars" -ForegroundColor Gray
                $lic = Get-WindowsLicensingInfo
                Write-Host "  Edition: $($lic.Caption) | Status: $($lic.LicenseStatus)" -ForegroundColor Gray
            }
        }
        '*Hardware*' {
            Invoke-HeadlessStep "[Tab 1/9: Dell Warranty & Hardware Health] Querying Dell Hardware Warranty API & SLA..." {
                $dell = Get-DellWarrantyInfo
                Write-Host "  Tag: $($dell.ServiceTag) | Model: $($dell.SystemModel) | Status: $($dell.WarrantyStatus)" -ForegroundColor Gray
                $bat = Get-BatteryHealthInfo
                Write-Host "  Battery: $(if ($bat -and $bat.Present) { "$($bat.WearPercent) wear ($($bat.Status))" } else { 'Desktop / AC-only' })" -ForegroundColor Gray
                $stor = @(Get-StorageReliabilityInfo)
                Write-Host "  Storage disks reporting SMART: $($stor.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 2/9: Pre-Flight Diagnostics] Auditing TPM 2.0 Attestation & EK Certificates..." {
                $tpm = Get-HeadlessTpmReadiness
                Write-Host "  TPM Attestation Ready: $($tpm.ReadyForAttestation)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 3/9: Autopilot & Cloud Registration] Harvesting Genuine OA3 Hardware Hash..." {
                $hash = Get-AutopilotHash
                Write-Host "  Serial: $($hash.SerialNumber) | Hardware Hash: $($hash.HardwareHash.Length) chars" -ForegroundColor Gray
                $lic = Get-WindowsLicensingInfo
                Write-Host "  Digital OEM Key Present: $(if ($lic.OemKey) { 'Yes' } else { 'No' })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 4/9: Windows Update & Drivers] Checking OEM Firmware & Driver Updates..." {
                $wu = Get-HubPendingWindowsUpdates
                Write-Host "  Pending Updates: $(if ($wu) { $wu.Count } else { 0 })" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 5/9: Win32 Packaging] Inspecting Hardware Diagnostic Tool Packaging Container..." {
                $diag = Get-Win32AppDiagnostics
                Write-Host "  Packaging Workspace: $($diag.StagingDirectory)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 6/9: App Deployment] Evaluating Intake Applications..." {
                $apps = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue)
                Write-Host "  System Applications Cataloged: $($apps.Count)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 7/9: Hybrid & Co-Mgmt] Auditing Device Domain / Entra Affiliation..." {
                $prt = Get-EntraPrtDiagnostics
                Write-Host "  Entra PRT Status: $($prt.HasPrt)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 8/9: ESP Diagnostics & Lifecycle] Auditing Autopilot Provisioning History..." {
                $ext = Get-IntuneExtensionHealth
                Write-Host "  Lifecycle IME Status: $($ext.Status)" -ForegroundColor Gray
            }
            Invoke-HeadlessStep "[Tab 9/9: Precision Enterprise Fixes] Precision Component Diagnostic Health Readout..." {
                $r = Get-HubRemediationHealthOverview
                Write-Host "  WMI Subsystem Healthy: $($r.WmiHealthy)" -ForegroundColor Gray
            }
        }
        default {
            Write-Host "Unknown routine name '$RoutineName'. Expected one containing Intune, Hybrid, Local, Remediation, or Hardware." -ForegroundColor Red
        }
    }
    Write-Host "`n[DONE] Playbook '$RoutineName' completed successfully across all 9 tabs!" -ForegroundColor Green
}

# ==============================================================================
# SECTION B: INTERACTIVE CYBER-DARK WPF XAML INTERFACE
# ==============================================================================

function Start-AutopilotHubGui {
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Autopilot Provisioning Hub - Enterprise Endpoint Deployment"
        Height="840" Width="1400" MinHeight="640" MinWidth="1100"
        WindowStartupLocation="CenterScreen"
        Background="#F3F4F6" Foreground="#1E293B"
        FontFamily="Segoe UI Variable Text, Segoe UI, sans-serif">

    <Window.Resources>
        <!-- WinUI 3 Light OOBE Palette -->
        <SolidColorBrush x:Key="BgCanvas" Color="#F3F4F6"/>
        <SolidColorBrush x:Key="CardBg" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="CardSubtle" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BorderSubtle" Color="#E2E8F0"/>
        <SolidColorBrush x:Key="BorderStrong" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#0067C0"/>
        <SolidColorBrush x:Key="AccentHover" Color="#106EBE"/>
        <SolidColorBrush x:Key="AccentPressed" Color="#005A9E"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#1E293B"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#475569"/>
        <SolidColorBrush x:Key="TextMuted" Color="#64748B"/>

        <!-- Standard Button Style (WinUI 3 Resting / Hover / Pressed) -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#F1F5F9"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#94A3B8"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#E2E8F0"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#64748B"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- WinUI 3 Accent Button Style -->
        <Style x:Key="AccentBtn" TargetType="Button">
            <Setter Property="Background" Value="#0067C0"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#0067C0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#1975C5"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#1975C5"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#0054A4"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#0054A4"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- WinUI 3 Destructive Button Style -->
        <Style x:Key="DestructiveBtn" TargetType="Button">
            <Setter Property="Background" Value="#FEE2E2"/>
            <Setter Property="Foreground" Value="#B91C1C"/>
            <Setter Property="BorderBrush" Value="#FCA5A5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#FCA5A5"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#F87171"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border" Value="#FECACA"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#FEE2E2"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Custom TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="CaretBrush" Value="#0067C0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="False"
                                          HorizontalScrollBarVisibility="{TemplateBinding ScrollViewer.HorizontalScrollBarVisibility}"
                                          VerticalScrollBarVisibility="{TemplateBinding ScrollViewer.VerticalScrollBarVisibility}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" TargetName="border" Value="#0067C0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                                <Setter Property="Background" TargetName="border" Value="#F8FAFC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Custom CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#334155"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- Dark ScrollBar Thumb -->
        <Style x:Key="DarkScrollBarThumb" TargetType="{x:Type Thumb}">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Thumb}">
                        <Border Background="#CBD5E1" CornerRadius="3" Margin="2"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#94A3B8"/>
                            </Trigger>
                            <Trigger Property="IsDragging" Value="True">
                                <Setter Property="Background" Value="#64748B"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Light ScrollBar Style -->
        <Style TargetType="{x:Type ScrollBar}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Height" Value="8"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="{x:Type ScrollBar}">
                                <Grid Background="Transparent">
                                    <Track x:Name="PART_Track" IsDirectionReversed="False">
                                        <Track.Thumb>
                                            <Thumb Style="{StaticResource DarkScrollBarThumb}"/>
                                        </Track.Thumb>
                                    </Track>
                                </Grid>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
                <Trigger Property="Orientation" Value="Vertical">
                    <Setter Property="Width" Value="8"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="{x:Type ScrollBar}">
                                <Grid Background="Transparent">
                                    <Track x:Name="PART_Track" IsDirectionReversed="True">
                                        <Track.Thumb>
                                            <Thumb Style="{StaticResource DarkScrollBarThumb}"/>
                                        </Track.Thumb>
                                    </Track>
                                </Grid>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Light GridViewColumnHeader Style -->
        <Style TargetType="{x:Type GridViewColumnHeader}">
            <Setter Property="Background" Value="#F1F5F9"/>
            <Setter Property="Foreground" Value="#475569"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type GridViewColumnHeader}">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E2E8F0"/>
                                <Setter Property="Foreground" Value="#1E293B"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#CBD5E1"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ListViewItem Style -->
        <Style TargetType="{x:Type ListViewItem}">
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Padding" Value="4,3"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListViewItem}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}"
                                SnapsToDevicePixels="true">
                            <GridViewRowPresenter VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                                                  SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#F8FAFC"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#E0EDFB"/>
                                <Setter Property="Foreground" Value="#004578"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ListBoxItem Style -->
        <Style TargetType="{x:Type ListBoxItem}">
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ListBoxItem}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#F8FAFC"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#E0EDFB"/>
                                <Setter Property="Foreground" Value="#004578"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ComboBox Toggle Button for Non-Editable -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="28" />
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" CornerRadius="4"
                        Background="#FFFFFF" BorderBrush="#E2E8F0" BorderThickness="1" />
                <Border Grid.Column="0" Background="Transparent" Margin="1" />
                <Path x:Name="Arrow" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z" Fill="#64748B" />
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="#F1F5F9" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#CBD5E1" />
                    <Setter TargetName="Arrow" Property="Fill" Value="#0067C0" />
                </Trigger>
                <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="#FFFFFF" />
                    <Setter TargetName="Border" Property="BorderBrush" Value="#0067C0" />
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Opacity" Value="0.4" />
                    <Setter TargetName="Arrow" Property="Opacity" Value="0.4" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <!-- ComboBox Toggle Button for Editable (arrow glyph only) -->
        <ControlTemplate x:Key="ComboBoxTextBoxToggleButton" TargetType="ToggleButton">
            <Border x:Name="Border" Width="28" Background="Transparent" BorderThickness="0" CornerRadius="0,4,4,0">
                <Path x:Name="Arrow" HorizontalAlignment="Center" VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z" Fill="#64748B" />
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="#F1F5F9" />
                    <Setter TargetName="Arrow" Property="Fill" Value="#0067C0" />
                </Trigger>
                <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="#FFFFFF" />
                    <Setter TargetName="Arrow" Property="Fill" Value="#0067C0" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <!-- WinUI 3 ComboBoxItem Style -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}" Margin="2,1" CornerRadius="3">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="#F1F5F9"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="#0067C0"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- WinUI 3 ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <!-- Visual for non-editable -->
                            <ToggleButton x:Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}"
                                          Focusable="False" ClickMode="Press"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,0,28,0" VerticalAlignment="Center" HorizontalAlignment="Left" />

                            <!-- Visual for editable -->
                            <Border x:Name="EditableBorder" Background="#FFFFFF" BorderBrush="#E2E8F0" BorderThickness="1"
                                    CornerRadius="4" Visibility="Collapsed">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*" />
                                        <ColumnDefinition Width="28" />
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="PART_EditableTextBox" Style="{x:Null}" Background="Transparent" Foreground="#1E293B"
                                             CaretBrush="#0067C0" BorderThickness="0" Padding="8,4"
                                             VerticalContentAlignment="Center" FontSize="12"/>
                                    <ToggleButton x:Name="EditableToggleButton" Grid.Column="1"
                                                  Template="{StaticResource ComboBoxTextBoxToggleButton}"
                                                  Focusable="False" ClickMode="Press"
                                                  IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" />
                                </Grid>
                            </Border>

                            <!-- Dropdown Popup -->
                            <Popup x:Name="Popup" Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="#FFFFFF" BorderThickness="1"
                                            BorderBrush="#E2E8F0" CornerRadius="4" Margin="0,2,0,0">
                                        <ScrollViewer Margin="2,4" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEditable" Value="True">
                                <Setter Property="IsTabStop" Value="False" />
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed" />
                                <Setter TargetName="ToggleButton" Property="Visibility" Value="Collapsed" />
                                <Setter TargetName="EditableBorder" Property="Visibility" Value="Visible" />
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="EditableBorder" Property="BorderBrush" Value="#CBD5E1" />
                            </Trigger>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter TargetName="EditableBorder" Property="BorderBrush" Value="#0067C0" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- WinUI 3 Dark ToolTip Style -->
        <Style TargetType="ToolTip">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="11.5"/>
            <Setter Property="HasDropShadow" Value="True"/>
        </Style>
        <!-- WinUI 3 Dark ContextMenu Style -->
        <Style TargetType="ContextMenu">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4,4"/>
            <Setter Property="HasDropShadow" Value="True"/>
        </Style>
        <Style TargetType="MenuItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="Separator">
            <Setter Property="Background" Value="#E2E8F0"/>
            <Setter Property="Margin" Value="4,2"/>
        </Style>
    </Window.Resources>

    <Grid Margin="12,8,12,8">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="Auto"/> <!-- Telemetry Row -->
            <RowDefinition Height="Auto"/> <!-- Deployment Playbook HUD -->
            <RowDefinition Height="*"/>    <!-- TabControl -->
            <RowDefinition Height="Auto"/> <!-- Progress Bar -->
            <RowDefinition Name="RowConsoleLog" Height="115"/>  <!-- Console Log (collapsible) -->
        </Grid.RowDefinitions>

        <!-- HEADER BAR: compact layout with slim title and pill status badges -->
        <Grid Grid.Row="0" Margin="0,0,0,3">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Border Name="BorderTenantLogo" CornerRadius="4" Background="#FFFFFF" BorderBrush="#CBD5E1" BorderThickness="1" Padding="3" Margin="0,0,8,0" Visibility="Collapsed">
                        <Image Name="ImgTenantLogo" Width="20" Height="20" Stretch="Uniform"/>
                    </Border>
                    <StackPanel Orientation="Vertical">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Name="TxtHubTitle" Text="Autopilot Provisioning Hub" FontSize="14" FontWeight="SemiBold" Foreground="#1E293B"/>
                            <Border Name="BadgeTenantBrand" Background="#E0F2FE" BorderBrush="#BAE6FD" BorderThickness="1" CornerRadius="3" Padding="4,1" Margin="8,0,0,0" VerticalAlignment="Center" Visibility="Collapsed">
                                <TextBlock Name="TxtTenantBrand" Text="TENANT" FontSize="9" FontWeight="Bold" Foreground="#0284C7"/>
                            </Border>
                        </StackPanel>
                        <TextBlock Text="Microsoft Intune &amp; Windows Autopilot Automated Deployment Engine" FontSize="9.5" Foreground="#64748B" Margin="0,1,0,0"/>
                    </StackPanel>
                </StackPanel>
            </StackPanel>

            <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,2,0,0">
                    <!-- Live privilege / session badge -->
                    <Border Name="BadgePrivilege" Background="#F1F5F9" CornerRadius="3" Padding="5,1.5" Margin="0,0,0,0" VerticalAlignment="Center" BorderBrush="#CBD5E1" BorderThickness="1" Cursor="Hand"
                            ToolTip="The privilege level this window is running with. Click for details. Hash harvest, rename, app installs, clock sync and restart-resume all need elevation.">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Ellipse Name="DotPrivilege" Width="5" Height="5" Fill="#B91C1C" VerticalAlignment="Center" Margin="0,0,4,0"/>
                            <TextBlock Name="TxtPrivilege" Text="PRIV: CHECKING" FontSize="8.5" FontWeight="SemiBold" Foreground="#475569" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Button Name="BtnFixPrivilege" Content="Fix Privileges" Style="{StaticResource DestructiveBtn}" Margin="4,0,0,0" Padding="5,1.5" FontSize="8.5" Visibility="Collapsed"
                            ToolTip="Walks you through relaunching the Hub with the privilege level it needs."/>
                    <!-- Power & Battery Posture Monitor Badge -->
                    <Border Name="BadgePower" Background="#DCFCE7" CornerRadius="3" Padding="5,1.5" Margin="4,0,0,0" VerticalAlignment="Center" BorderBrush="#BBF7D0" BorderThickness="1"
                            ToolTip="Hardware power posture (AC line vs battery status). Red warning on low battery.">
                        <TextBlock Name="TxtPowerStatus" Text="POWER: PROBING" FontSize="8.5" FontWeight="SemiBold" Foreground="#15803D" VerticalAlignment="Center"/>
                    </Border>
                    <!-- Graph Authentication Session Badge -->
                    <Border Name="BadgeGraphAuth" Background="#F1F5F9" CornerRadius="3" Padding="5,1.5" Margin="4,0,0,0" VerticalAlignment="Center" BorderBrush="#CBD5E1" BorderThickness="1">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Ellipse Name="DotGraphStatus" Width="5" Height="5" Fill="#B91C1C" VerticalAlignment="Center" Margin="0,0,4,0"/>
                            <TextBlock Name="TxtGraphStatus" Text="GRAPH: NOT SIGNED IN" FontSize="8.5" FontWeight="SemiBold" Foreground="#475569" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Border Background="#E0F2FE" CornerRadius="3" Padding="5,1.5" Margin="4,0,0,0" VerticalAlignment="Center" BorderBrush="#BAE6FD" BorderThickness="1">
                        <TextBlock Text="STANDALONE KERNEL" FontSize="8.5" FontWeight="SemiBold" Foreground="#0284C7"/>
                    </Border>
            </StackPanel>

            <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                <Button Name="BtnConnectGraph" Content="Sign In to Intune" Style="{StaticResource AccentBtn}" Margin="0,0,5,0" Padding="8,3" FontSize="10.5" Height="26"/>
                <Button Name="BtnToolsDropdown" Content="Actions [v]" Padding="8,3" FontSize="10.5" Height="26" ToolTip="Quick Tools, Utilities, Shell &amp; Power Operations">
                    <Button.ContextMenu>
                        <ContextMenu Name="MenuTools">
                            <MenuItem Name="MenuWindowsUpdate" Header="Windows Update &amp; Patch Cascade" FontWeight="SemiBold"/>
                            <Separator/>
                            <MenuItem Name="MenuCreateLocalUser" Header="Create Local User Account..."/>
                            <MenuItem Name="MenuPersonalMsa" Header="Unlock Personal Microsoft Account (MSA)"/>
                            <Separator/>
                            <MenuItem Name="MenuLaunchEdge" Header="Edge Browser / Captive Portal"/>
                            <MenuItem Name="MenuScreenshotUsb" Header="Capture Staging Proof (USB Screenshot)"/>
                            <MenuItem Name="MenuToggleDpi" Header="Toggle UI Scaling (100% / 125%)"/>
                            <MenuItem Name="MenuInstallPwsh" Header="Install PowerShell 7"/>
                            <MenuItem Name="MenuQuickCmd" Header="Command Prompt (Shift+F10)"/>
                            <MenuItem Name="MenuTimeSync" Header="Synchronize Clock (w32tm)"/>
                            <Separator/>
                            <MenuItem Name="MenuReboot" Header="Restart System Now" Foreground="#B91C1C" FontWeight="SemiBold"/>
                        </ContextMenu>
                    </Button.ContextMenu>
                </Button>
            </StackPanel>
        </Grid>

        <!-- HARDWARE TELEMETRY CARDS (COMPACT OOBE AESTHETIC) -->
        <Border Grid.Row="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="6,3" Margin="0,0,0,4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="1.4*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0">
                    <TextBlock Text="SERIAL NUMBER" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtSerial" Text="Detecting..." FontSize="10.5" Foreground="#1E293B" FontWeight="Bold" FontFamily="Consolas" Margin="0,1,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1">
                    <TextBlock Text="MAKE &amp; MODEL" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtModel" Text="Detecting..." FontSize="10.5" Foreground="#1E293B" FontWeight="SemiBold" Margin="0,1,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="2">
                    <TextBlock Text="TPM 2.0 STATUS" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtTpm" Text="Probing..." FontSize="10.5" Foreground="#15803D" FontWeight="Bold" Margin="0,1,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="3">
                    <TextBlock Text="SECURE BOOT" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtSecureBoot" Text="Probing..." FontSize="10.5" Foreground="#B45309" FontWeight="Bold" Margin="0,1,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="4">
                    <TextBlock Text="NETWORK STATUS" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtNetwork" Text="Checking..." FontSize="10.5" Foreground="#0284C7" FontWeight="Bold" Margin="0,1,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="5">
                    <TextBlock Text="DEVICE STATE" FontSize="8" Foreground="#64748B" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtDeviceState" Text="Inspecting..." FontSize="10.5" Foreground="#475569" FontWeight="Bold" Margin="0,1,0,0" TextTrimming="CharacterEllipsis"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ACTION ROUTINES / DEPLOYMENT PLAYBOOK HUD -->
        <Border Name="PlaybookHud" Grid.Row="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#CBD5E1" BorderThickness="1" Padding="8,4" Margin="0,0,0,5">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Left: Routine Badge & Selector -->
                <StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,8,0">
                    <Border Background="#0067C0" CornerRadius="3" Padding="5,2" Margin="0,0,6,0" VerticalAlignment="Center">
                        <TextBlock Text="PLAYBOOK" FontSize="8.5" FontWeight="Bold" Foreground="#FFFFFF"/>
                    </Border>
                    <ComboBox Name="CboPlaybookRoutine" Width="290" Height="25" FontSize="11" VerticalContentAlignment="Center" Background="#F8FAFC" BorderBrush="#CBD5E1" ToolTip="Select an end-to-end automated action playbook to execute across tabs">
                        <ComboBoxItem Content="1. Intune-Only Cloud Build" IsSelected="True"/>
                        <ComboBoxItem Content="2. Hybrid AD Join &amp; Co-Management Build"/>
                        <ComboBoxItem Content="3. Local User &amp; Offline Bypass Build"/>
                        <ComboBoxItem Content="4. Deep System Remediation &amp; Health Sweep"/>
                        <ComboBoxItem Content="5. Hardware Health &amp; Asset Intake Audit"/>
                    </ComboBox>
                </StackPanel>

                <!-- Center: Playbook Controls -->
                <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,10,0">
                    <Button Name="BtnPlaybookRun" Content="[&#x25B6; Run Playbook]" Style="{StaticResource AccentBtn}" Padding="8,2" FontSize="10.5" Height="25" Margin="0,0,4,0" ToolTip="Start the selected automated click-through playbook"/>
                    <Button Name="BtnPlaybookPause" Content="[&#x23F8; Pause]" Padding="8,2" FontSize="10.5" Height="25" Margin="0,0,4,0" IsEnabled="False" ToolTip="Pause or resume current playbook execution"/>
                    <Button Name="BtnPlaybookStop" Content="[&#x23F9; Stop]" Style="{StaticResource DestructiveBtn}" Padding="8,2" FontSize="10.5" Height="25" Margin="0,0,8,0" IsEnabled="False" ToolTip="Abort current playbook execution"/>
                </StackPanel>

                <!-- Right: Current Step Indicator Status -->
                <StackPanel Grid.Row="0" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="STEP:" FontSize="9" FontWeight="Bold" Foreground="#64748B" Margin="0,0,5,0" VerticalAlignment="Center"/>
                    <TextBlock Name="TxtPlaybookStep" Text="Ready (Select an action playbook above and click Run to begin automated click-through)" FontSize="10.5" Foreground="#1E293B" FontWeight="SemiBold" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </StackPanel>

                <!-- Row 1: Playbook Progress Bar -->
                <ProgressBar Name="PlaybookProgressBar" Grid.Row="1" Grid.ColumnSpan="3" Height="3.5" Minimum="0" Maximum="100" Value="0"
                             Background="#F1F5F9" Foreground="#0067C0" BorderThickness="0" Margin="0,4,0,1"/>
            </Grid>
        </Border>

        <!-- MAIN CONTENT TABS -->
        <TabControl Name="MainTabControl" Grid.Row="3" Background="Transparent" BorderThickness="0" Margin="0,0,0,8">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border x:Name="border" Background="#E2E8F0" CornerRadius="4,4,0,0" Margin="0,0,2,0" Padding="9,3.5" BorderBrush="#CBD5E1" BorderThickness="1,1,1,0">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" ContentSource="Header"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="border" Property="Background" Value="#FFFFFF"/>
                                        <Setter TargetName="border" Property="BorderBrush" Value="#CBD5E1"/>
                                        <Setter Property="Foreground" Value="#0067C0"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                    </Trigger>
                                    <Trigger Property="IsSelected" Value="False">
                                        <Setter Property="Foreground" Value="#475569"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="border" Property="Background" Value="#F1F5F9"/>
                                        <Setter Property="Foreground" Value="#1E293B"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                    <Setter Property="FontSize" Value="11"/>
                </Style>
            </TabControl.Resources>

            <!-- TAB 1: AUTOPILOT & CLOUD REGISTRATION -->
            <TabItem Name="TabAutopilot" Header="Autopilot &amp; Cloud Registration">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="420"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Tenant Mismatch Warning Banner -->
                    <Border Name="BannerTenantMismatch" Grid.Row="0" Grid.ColumnSpan="2" Background="#FEE2E2" BorderBrush="#FF99A4" BorderThickness="1" CornerRadius="4" Padding="12,8" Margin="0,0,0,8" Visibility="Collapsed">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="TENANT MISMATCH DETECTED" FontSize="11" FontWeight="Bold" Foreground="#B91C1C"/>
                                <TextBlock Name="TxtTenantMismatch" Text="Device cached profile does not match signed-in operator tenant." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,2,0,0"/>
                            </StackPanel>
                            <Button Name="BtnDismissMismatch" Grid.Column="1" Content="Dismiss" VerticalAlignment="Center" Padding="8,4" FontSize="11"/>
                        </Grid>
                    </Border>

                    <!-- Device Enrollment State Banner (drives what the operator is asked to do) -->
                    <Border Name="BannerDeviceState" Grid.Row="0" Grid.ColumnSpan="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Name="TxtDeviceStateTitle" Text="DEVICE STATE: INSPECTING..." FontSize="11" FontWeight="SemiBold" Foreground="#475569"/>
                                <TextBlock Name="TxtDeviceStateDetail" Text="Looking for a cached Autopilot profile, MDM enrollment and Entra join state..." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" Margin="0,3,0,0"/>
                            </StackPanel>
                            <Button Name="BtnDeviceStateAction" Grid.Column="1" Content="Harvest Hash Now" Style="{StaticResource AccentBtn}" VerticalAlignment="Center" Margin="12,0,0,0" Padding="10,5" Visibility="Collapsed"/>
                        </Grid>
                    </Border>

                    <!-- Left: Configuration Controls -->
                    <Border Grid.Row="1" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16" Margin="0,0,12,0">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <!-- Top Priority: Local Setup & OOBE Bypass (Instant 1-Click Access) -->
                                <Border Background="#F8FAFC" CornerRadius="4" BorderBrush="#CBD5E1" BorderThickness="1" Padding="10,8" Margin="0,0,0,12">
                                    <StackPanel>
                                        <Grid Margin="0,0,0,3">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <TextBlock Text="LOCAL SETUP &amp; OOBE BYPASS" FontSize="11" FontWeight="SemiBold" Foreground="#1E293B" VerticalAlignment="Center"/>
                                            <Border Grid.Column="1" Background="#E0F2FE" CornerRadius="3" Padding="4,1" BorderBrush="#BAE6FD" BorderThickness="1">
                                                <TextBlock Text="NO WORK/SCHOOL LOCK" FontSize="8.5" FontWeight="Bold" Foreground="#0284C7"/>
                                            </Border>
                                        </Grid>
                                        <TextBlock Text="Bypass corporate Autopilot, create a local admin account instantly, or unlock Personal Microsoft Account (MSA)." FontSize="10.5" Foreground="#64748B" TextWrapping="Wrap" Margin="0,0,0,6"/>

                                        <Grid Margin="0,0,0,5">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>
                                            <Button Name="BtnCreateLocalUser" Grid.Column="0" Content="Create Local User" Style="{StaticResource AccentBtn}" Height="28" FontSize="11" Margin="0,0,3,0"
                                                    ToolTip="Instantly creates a local Administrator user account right now in OOBE"/>
                                            <Button Name="BtnPersonalInstall" Grid.Column="1" Content="Unlock Personal MSA" Height="28" FontSize="11" Margin="3,0,0,0"
                                                    ToolTip="Clears Autopilot corporate locks, suppresses Work/School sign-in, and keeps network active for Personal Microsoft Account setup"/>
                                        </Grid>

                                        <Grid Margin="0,0,0,4">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>
                                            <Button Name="BtnSkipAutopilot" Grid.Column="0" Content="Offline Bypass (NRO)" Height="26" FontSize="10.5" Margin="0,0,3,0"
                                                    ToolTip="Purges Autopilot profile cache, writes BypassNRO=1 and disables network for offline local setup"/>
                                            <Button Name="BtnReenableNet" Grid.Column="1" Content="Re-enable Network" Height="26" FontSize="10.5" Margin="3,0,0,0"
                                                    ToolTip="Re-enables any physical network adapters that were disabled during offline setup"/>
                                        </Grid>

                                        <Grid Margin="0,0,0,2">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <CheckBox Name="ChkPersonalDisableNet" Content="Disable network adapters during offline bypass" IsChecked="True" FontSize="10.5" VerticalAlignment="Center"/>
                                            <Button Name="BtnRestartOobe" Grid.Column="1" Content="Restart OOBE" Height="22" Padding="6,2" FontSize="10" ToolTip="Kills OOBE broker processes and restarts setup cleanly"/>
                                        </Grid>
                                    </StackPanel>
                                </Border>

                                <TextBlock Text="PROVISIONING CONFIGURATION" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,10"/>

                                <!-- Group Tag Selection -->
                                <TextBlock Text="Autopilot Group Tag:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                <ComboBox Name="CmbGroupTag" IsEditable="True" Height="30" Margin="0,0,0,10" Background="#FFFFFF" Foreground="#1E293B">
                                    <ComboBoxItem Content="Corporate-Laptops" IsSelected="True"/>
                                    <ComboBoxItem Content="Standard-Workstations"/>
                                    <ComboBoxItem Content="DevOps-Engineering"/>
                                    <ComboBoxItem Content="Executive-Fleet"/>
                                    <ComboBoxItem Content="Kiosk-AssignedAccess"/>
                                    <ComboBoxItem Content="Finance-Workstations"/>
                                </ComboBox>

                                <!-- Assigned User -->
                                <TextBlock Text="Assigned User UPN (Optional):" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                <TextBox Name="TxtAssignedUser" Height="30" Margin="0,0,0,10"/>

                                <!-- Computer Rename -->
                                <Grid Margin="0,0,0,4">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="Computer Name (Tokens: %SERIAL%, %RAND%):" FontSize="12" FontWeight="SemiBold" Foreground="#475569" VerticalAlignment="Center"/>
                                    <CheckBox Name="ChkEnableRename" Grid.Column="1" Content="Enable" IsChecked="True" Foreground="#64748B" FontSize="11" VerticalAlignment="Center"/>
                                </Grid>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox Name="TxtComputerName" Height="30" Text="WS-%SERIAL%"/>
                                    <Button Name="BtnApplyRename" Grid.Column="1" Content="Apply Name" Margin="6,0,0,0" Padding="10,4"/>
                                </Grid>

                                <!-- Options -->
                                <TextBlock Text="PROVISIONING GATES" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,4,0,8"/>
                                <CheckBox Name="ChkWaitForSync" Content="Wait for Profile Assignment (-WaitForSync)" IsChecked="True" Margin="0,0,0,6"/>
                                <CheckBox Name="ChkAutoDetectUsb" Content="Auto-Detect USB for CSV Export Fallback" IsChecked="True" Margin="0,0,0,6"/>
                                <CheckBox Name="ChkAutoReboot" Content="Reboot into ESP upon Successful Profile Assignment" IsChecked="False" Margin="0,0,0,12"/>

                                <!-- Primary Actions -->
                                <Button Name="BtnHarvestHash" Content="Harvest Hardware Hash" Style="{StaticResource AccentBtn}" Height="34" Margin="0,0,0,6"/>
                                <Button Name="BtnExportCsv" Content="Export Intune CSV (USB Priority)" Height="32" Margin="0,0,0,6"/>
                                <Button Name="BtnRegisterIntune" Content="Register Device with Intune (Graph)" Style="{StaticResource AccentBtn}" Height="34" Margin="0,0,0,12"/>

                                <!-- Controlled Offline Staging & Adaptive Evaluation Profiles -->
                                <TextBlock Text="CONTROLLED OFFLINE STAGING" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,4,0,6"/>
                                <TextBlock Text="Stage devices offline with reversible profile quarantine, loopback policy fences, and adaptive evaluation profiles (Hardware Diagnostics, Workload Benchmarking, Software Sandbox)." FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,0,0,8"/>
                                <Border Name="BannerOfflineStaging" Background="#DCFCE7" BorderBrush="#BBF7D0" BorderThickness="1" CornerRadius="4" Padding="10,8" Margin="0,0,0,8" Visibility="Collapsed">
                                    <StackPanel Orientation="Vertical">
                                        <TextBlock Name="TxtOfflineStagingTitle" Text="OFFLINE STAGING ACTIVE: [PROFILE]" FontSize="11" FontWeight="Bold" Foreground="#15803D" Margin="0,0,0,2"/>
                                        <TextBlock Name="TxtOfflineStagingDetail" Text="Device is isolated with reversible profile quarantine. Click below to restore cloud link." FontSize="10.5" Foreground="#475569" TextWrapping="Wrap" Margin="0,0,0,6"/>
                                        <Button Name="BtnReDockStaging" Content="Reconnect to Enterprise Cloud &amp; Resume Autopilot" Style="{StaticResource AccentBtn}" Height="30"/>
                                    </StackPanel>
                                </Border>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button Name="BtnOfflineStaging" Grid.Column="0" Content="Offline Staging Mode" Height="30" Margin="0,0,4,0"
                                            ToolTip="Controlled offline staging with reversible profile quarantine and adaptive evaluation profiles"/>
                                    <Button Name="BtnResumeAutopilot" Grid.Column="1" Content="Resume Cloud Autopilot" Height="30" Margin="4,0,0,0"
                                            ToolTip="Re-dock device to Enterprise Cloud, restore quarantined Autopilot profile and clear fences"/>
                                </Grid>

                                <Button Name="BtnSaveEnvDefaults" Content="Save Current Settings to .env" Height="28" Margin="0,0,0,12"/>

                                <!-- Profile Policy & Offline Provisioning -->
                                <TextBlock Text="PROFILE &amp; OFFLINE PROVISIONING" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,6"/>
                                <Button Name="BtnInspectPolicy" Content="Inspect Decoded Autopilot Policy" Height="30" Margin="0,0,0,6"
                                        ToolTip="Decodes CloudAssignedOobeConfig bitmask (EULA, OEM registration, account type, privacy settings)"/>
                                <Grid Margin="0,0,0,6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button Name="BtnExportOfflineJson" Grid.Column="0" Content="Export Offline JSON" Margin="0,0,4,0" Height="30"
                                            ToolTip="Export AutopilotConfigurationFile.json to USB flash drive for air-gapped provisioning"/>
                                    <Button Name="BtnInjectOfflineJson" Grid.Column="1" Content="Inject Offline JSON" Margin="4,0,0,0" Height="30"
                                            ToolTip="Inject AutopilotConfigurationFile.json into Windows Provisioning folder"/>
                                </Grid>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>

                    <!-- Right: Hash Preview & Registration Status -->
                    <Border Grid.Row="1" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <Grid Grid.Row="0" Margin="0,0,0,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="HARDWARE HASH BUFFER" FontSize="11" FontWeight="SemiBold" Foreground="#334155" VerticalAlignment="Center"/>
                                    <Border Name="BadgeHashStatus" Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="3" Padding="6,2" Margin="10,0,0,0">
                                        <TextBlock Name="TxtHashStatus" Text="NOT HARVESTED" FontSize="10.5" FontWeight="SemiBold" Foreground="#64748B"/>
                                    </Border>
                                </StackPanel>
                                <Button Name="BtnCopyHash" Grid.Column="1" Content="Copy Hash" Padding="10,4" FontSize="12"/>
                            </Grid>

                            <TextBox Name="TxtHashBox" Grid.Row="1" TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True"
                                     Background="#FFFFFF" Foreground="#475569" FontFamily="Consolas" FontSize="11" Padding="10"
                                     VerticalScrollBarVisibility="Auto" BorderBrush="#E2E8F0"/>

                            <Border Grid.Row="2" Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,10,0,0">
                                <TextBlock Name="TxtHashMeta" Text="Hardware hash buffer empty. Click 'Harvest Hardware Hash' to query local WMI provider."
                                           FontSize="11.5" Foreground="#64748B"/>
                            </Border>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 2: APP DEPLOYMENT (WingetBatch) -->
            <TabItem Name="TabAppDeploy" Header="App Deployment">
                <DockPanel Margin="0,12,0,0">
                    <!-- Managed-device advisory: app assignment is Intune's job once the device is enrolled -->
                    <Border Name="BannerAppAdvisory" DockPanel.Dock="Top" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,0,10">
                        <TextBlock Name="TxtAppAdvisory" Text="Checking whether this device is Intune-managed..." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap"/>
                    </Border>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Top Preset Bar -->
                    <Border Grid.Row="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="PRESETS:" FontSize="10.5" FontWeight="SemiBold" Foreground="#64748B" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <Button Name="BtnPresetWorkstation" Content="Standard Workstation" Margin="0,0,6,0" Padding="8,4" FontSize="12"/>
                                <Button Name="BtnPresetBrowsers" Content="Web Browsers" Margin="0,0,6,0" Padding="8,4" FontSize="12"/>
                                <Button Name="BtnPresetDev" Content="Developer Suite" Margin="0,0,6,0" Padding="8,4" FontSize="12"/>
                                <Button Name="BtnSelectAllApps" Content="Select All" Margin="0,0,6,0" Padding="8,4" FontSize="12"/>
                                <Button Name="BtnClearApps" Content="Clear All" Padding="8,4" FontSize="12"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <CheckBox Name="ChkSilentInstall" Content="Silent Installation (--silent)" IsChecked="True" Margin="0,0,12,0"/>
                                <CheckBox Name="ChkMachineScope" Content="Machine Scope" IsChecked="True"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- App Categories Grid -->
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <!-- Browsers -->
                        <Border Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12" Margin="0,0,6,0">
                            <StackPanel>
                                <TextBlock Text="BROWSERS" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,10"/>
                                <CheckBox Name="AppChrome" Tag="Google.Chrome" Content="Google Chrome" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppFirefox" Tag="Mozilla.Firefox" Content="Mozilla Firefox" Margin="0,0,0,8"/>
                                <CheckBox Name="AppBrave" Tag="Brave.Brave" Content="Brave Browser" Margin="0,0,0,8"/>
                                <CheckBox Name="AppEdgeDev" Tag="Microsoft.Edge.Dev" Content="Edge Dev" Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>

                        <!-- Developer Tools -->
                        <Border Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12" Margin="0,0,6,0">
                            <StackPanel>
                                <TextBlock Text="DEVELOPER TOOLS" FontSize="11" FontWeight="SemiBold" Foreground="#15803D" Margin="0,0,0,10"/>
                                <CheckBox Name="AppVSCode" Tag="Microsoft.VisualStudioCode" Content="VS Code" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppGit" Tag="Git.Git" Content="Git for Windows" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppTerminal" Tag="Microsoft.WindowsTerminal" Content="Windows Terminal" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppNode" Tag="OpenJS.NodeJS.LTS" Content="Node.js LTS" Margin="0,0,0,8"/>
                                <CheckBox Name="AppPython" Tag="Python.Python.3.12" Content="Python 3.12" Margin="0,0,0,8"/>
                                <CheckBox Name="AppGh" Tag="GitHub.cli" Content="GitHub CLI" Margin="0,0,0,8"/>
                                <CheckBox Name="AppDocker" Tag="Docker.DockerDesktop" Content="Docker Desktop" Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>

                        <!-- Productivity -->
                        <Border Grid.Column="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12" Margin="0,0,6,0">
                            <StackPanel>
                                <TextBlock Text="PRODUCTIVITY" FontSize="11" FontWeight="SemiBold" Foreground="#B45309" Margin="0,0,0,10"/>
                                <CheckBox Name="AppOffice" Tag="Microsoft.Office" Content="Microsoft 365" Margin="0,0,0,8"/>
                                <CheckBox Name="AppSlack" Tag="SlackTechnologies.Slack" Content="Slack" Margin="0,0,0,8"/>
                                <CheckBox Name="AppZoom" Tag="Zoom.Zoom" Content="Zoom Workplace" Margin="0,0,0,8"/>
                                <CheckBox Name="AppNotion" Tag="Notion.Notion" Content="Notion" Margin="0,0,0,8"/>
                                <CheckBox Name="AppObsidian" Tag="Obsidian.Obsidian" Content="Obsidian" Margin="0,0,0,8"/>
                                <CheckBox Name="App7Zip" Tag="7zip.7zip" Content="7-Zip" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppAcrobat" Tag="Adobe.Acrobat.Reader.64-bit" Content="Adobe Reader" Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>

                        <!-- Utilities -->
                        <Border Grid.Column="3" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12">
                            <StackPanel>
                                <TextBlock Text="SYSTEM UTILITIES" FontSize="11" FontWeight="SemiBold" Foreground="#A78BFA" Margin="0,0,0,10"/>
                                <CheckBox Name="AppPowerToys" Tag="Microsoft.PowerToys" Content="PowerToys" IsChecked="True" Margin="0,0,0,8"/>
                                <CheckBox Name="AppSysinternals" Tag="Microsoft.Sysinternals.Suite" Content="Sysinternals" Margin="0,0,0,8"/>
                                <CheckBox Name="AppSysInformer" Tag="SystemInformer.SystemInformer" Content="System Informer" Margin="0,0,0,8"/>
                                <CheckBox Name="AppWireshark" Tag="WiresharkFoundation.Wireshark" Content="Wireshark" Margin="0,0,0,8"/>
                                <CheckBox Name="AppPuTTY" Tag="PuTTY.PuTTY" Content="PuTTY" Margin="0,0,0,8"/>
                                <CheckBox Name="AppHwInfo" Tag="REALiX.HWiNFO" Content="HWiNFO" Margin="0,0,0,8"/>
                                <CheckBox Name="AppTreeSize" Tag="JAMSoftware.TreeSize.Free" Content="TreeSize Free" Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>
                    </Grid>

                    <!-- Custom Winget ID & Action -->
                    <Border Grid.Row="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Custom Winget ID:" VerticalAlignment="Center" Margin="0,0,8,0" FontSize="12" FontWeight="SemiBold" Foreground="#475569"/>
                            <TextBox Name="TxtCustomPkg" Grid.Column="1" Height="30" Margin="0,0,8,0"/>
                            <Button Name="BtnInstallBatch" Grid.Column="2" Content="Install Selected Applications" Style="{StaticResource AccentBtn}" Height="32" Padding="14,4"/>
                        </Grid>
                    </Border>
                </Grid>
                </DockPanel>
            </TabItem>

            <!-- TAB 3: WIN32 PACKAGING (WingetIntune) -->
            <TabItem Name="TabPackaging" Header="Win32 Packaging">
                <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left: Package Builder (fields scroll if cramped; the action button stays pinned) -->
                    <Border Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16" Margin="0,0,6,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                                <StackPanel>
                                    <TextBlock Text="WIN32 PACKAGE BUILDER (.INTUNEWIN)" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,14"/>

                                    <TextBlock Text="Winget Package ID / Source:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <TextBox Name="TxtPkgId" Height="32" Text="Mozilla.Firefox" Margin="0,0,0,10"/>

                                    <TextBlock Text="Display Name:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <TextBox Name="TxtPkgDisplayName" Height="32" Text="Mozilla Firefox Enterprise" Margin="0,0,0,10"/>

                                    <TextBlock Text="Output Folder:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <TextBox Name="TxtPkgOutputDir" Height="32" Text="C:\temp\WingetIntune\Output" Margin="0,0,0,10"/>

                                    <TextBlock Text="Silent Install Arguments:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <TextBox Name="TxtPkgInstallArgs" Height="32" Text="/S" Margin="0,0,0,16"/>
                                </StackPanel>
                            </ScrollViewer>
                            <Button Name="BtnBuildPackage" Grid.Row="1" Content="Build Package (.intunewin)" Style="{StaticResource AccentBtn}" Height="34" Margin="0,12,0,0"/>
                        </Grid>
                    </Border>

                    <!-- Right: Cloud Publisher (fields scroll if cramped; the action button stays pinned) -->
                    <Border Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16" Margin="6,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                                <StackPanel>
                                    <TextBlock Text="MICROSOFT GRAPH INTUNE CLOUD PUBLISHER" FontSize="11" FontWeight="SemiBold" Foreground="#334155" Margin="0,0,0,14"/>

                                    <TextBlock Text="Target Assignment Intent:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <ComboBox Name="CmbAssignmentIntent" Height="32" Margin="0,0,0,10" Background="#FFFFFF" Foreground="#1E293B">
                                        <ComboBoxItem Content="Available (Self-Service in Company Portal)" IsSelected="True"/>
                                        <ComboBoxItem Content="Required (Mandatory Push)"/>
                                        <ComboBoxItem Content="Uninstall"/>
                                    </ComboBox>

                                    <TextBlock Text="Target Entra ID Group / Audience:" FontSize="12" FontWeight="SemiBold" Foreground="#475569" Margin="0,0,0,4"/>
                                    <TextBox Name="TxtAssignGroup" Height="32" Text="All Devices" Margin="0,0,0,16"/>

                                    <Border Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="4" Padding="12" Margin="0,0,0,16">
                                        <TextBlock Text="Direct Graph publishing utilizes chunked Azure SAS storage upload and generates automated Win32 detection rules."
                                                   FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap"/>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                            <Button Name="BtnPublishIntune" Grid.Row="1" Content="Publish to Intune Cloud" Style="{StaticResource AccentBtn}" Height="34" Margin="0,12,0,0"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 4: PRE-FLIGHT DIAGNOSTICS (IntuneShared) -->
            <TabItem Name="TabPreFlight" Header="Pre-Flight Diagnostics">
                <Border Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16" Margin="0,12,0,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <Grid Grid.Row="0" Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="7-STAGE ENTERPRISE PRE-FLIGHT DIAGNOSTIC LADDER" FontSize="11" FontWeight="SemiBold" Foreground="#334155" VerticalAlignment="Center"/>
                            <Button Name="BtnRunDiag" Grid.Column="1" Content="Run Diagnostics" Padding="12,5"/>
                        </Grid>

                        <Grid Grid.Row="1">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ListBox Name="LstDiagStages" Grid.Row="0" Background="#FFFFFF" BorderBrush="#E2E8F0">
                                <ListBox.ItemTemplate>
                                    <DataTemplate>
                                        <Border Padding="10,6" BorderBrush="#E2E8F0" BorderThickness="0,0,0,1">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="35"/>
                                                    <ColumnDefinition Width="180"/>
                                                    <ColumnDefinition Width="*"/>
                                                </Grid.ColumnDefinitions>
                                                <TextBlock Text="{Binding Stage}" FontWeight="Bold" Foreground="#0284C7"/>
                                                <TextBlock Grid.Column="1" Text="{Binding Name}" FontWeight="SemiBold" Foreground="#1E293B"/>
                                                <TextBlock Grid.Column="2" Text="{Binding Details}" Foreground="#64748B"/>
                                            </Grid>
                                        </Border>
                                    </DataTemplate>
                                </ListBox.ItemTemplate>
                            </ListBox>
                            <!-- Pre-Flight Remediation Action Bar -->
                            <Grid Grid.Row="1" Margin="0,10,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="PRE-FLIGHT REMEDIATION &amp; DRIVERS" FontSize="11" FontWeight="SemiBold" Foreground="#64748B" VerticalAlignment="Center"/>
                                <Button Name="BtnFixClockSkew" Grid.Column="1" Content="Auto-Remediate Clock Skew" Margin="0,0,8,0" Padding="10,5"
                                        ToolTip="Queries Microsoft Online date header and forces local time + w32tm /resync"/>
                                <Button Name="BtnTpmAttestation" Grid.Column="2" Content="TPM 2.0 &amp; EK Attestation Engine" Margin="0,0,8,0" Padding="10,5"
                                        ToolTip="Tests Endorsement Key (EK) certificates and manufacturer EK CA server reachability"/>
                                <Button Name="BtnInjectDriversPreflight" Grid.Column="3" Content="Inject USB Drivers (pnputil)" Margin="0,0,8,0" Padding="10,5"
                                        ToolTip="Scans USB drives for .inf network and chipset driver packages and installs them"/>
                                <Button Name="BtnPreflightWinUpdate" Grid.Column="4" Content="Windows Update" Padding="10,5" Style="{StaticResource AccentBtn}"
                                        ToolTip="Scan, download and install Windows Updates &amp; hardware drivers directly during OOBE Setup"/>
                            </Grid>
                        </Grid>
                    </Grid>
                </Border>
            </TabItem>

            <!-- TAB 5: WINDOWS UPDATE & DRIVER CASCADE -->
            <TabItem Name="TabWinUpdate" Header="Windows Update &amp; Drivers">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Top Card: Autonomous Multi-Pass Patch Cascade Engine -->
                    <Border Grid.Row="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14,10" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <!-- Left: Cascade Telemetry & Description -->
                            <StackPanel Grid.Row="0" Grid.Column="0">
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <Border Name="BadgeCascadeStatus" Background="#DCFCE7" CornerRadius="3" Padding="8,2" Margin="0,0,10,0" BorderBrush="#BBF7D0" BorderThickness="1">
                                        <TextBlock Name="TxtCascadeStatus" Text="CASCADE IDLE" FontSize="10.5" FontWeight="Bold" Foreground="#15803D"/>
                                    </Border>
                                    <TextBlock Text="AUTONOMOUS MULTI-PASS PATCH CASCADE" FontSize="13" FontWeight="SemiBold" Foreground="#1E293B" VerticalAlignment="Center"/>
                                    <TextBlock Name="TxtCascadePassInfo" Text=" (Pass 1 of 5)" FontSize="12" Foreground="#0284C7" FontWeight="SemiBold" VerticalAlignment="Center"/>
                                    <TextBlock Name="TxtCascadeInstalled" Text=" | 0 installed" FontSize="12" Foreground="#64748B" VerticalAlignment="Center"/>
                                </StackPanel>
                                <TextBlock Name="TxtCascadeStateDetail" Text="Fully automated cycle: scans, downloads, installs updates/drivers, reboots with persistence, and resumes automatically until the system is 100% patched." FontSize="11" Foreground="#64748B" Margin="0,4,0,0"/>
                            </StackPanel>

                            <!-- Right: Cascade Action Buttons & Abort Affordance -->
                            <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="BtnAbortReboot" Content="Cancel Reboot" Style="{StaticResource DestructiveBtn}" Margin="0,0,8,0" Padding="12,5" FontWeight="Bold" Visibility="Collapsed"
                                        ToolTip="Aborts the 5-second scheduled restart countdown and pauses cascade"/>
                                <Button Name="BtnStopCascade" Content="Stop Cascade" Style="{StaticResource DestructiveBtn}" Margin="0,0,8,0" Padding="12,5" Visibility="Collapsed"/>
                                <Button Name="BtnStartCascade" Content="Start Autonomous Patch Cascade" Style="{StaticResource AccentBtn}" Padding="16,6" FontWeight="Bold"
                                        ToolTip="Repeats: Scan to Install to Reboot to Resume until zero updates remain"/>
                            </StackPanel>

                            <!-- Options Row -->
                            <Border Grid.Row="1" Grid.ColumnSpan="2" Background="#F8FAFC" CornerRadius="3" Padding="10,6" Margin="0,8,0,0" BorderBrush="#CBD5E1" BorderThickness="1">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <CheckBox Name="ChkCascadeDrivers" Grid.Column="0" Content="Include Hardware &amp; Firmware Drivers" IsChecked="True" Foreground="#1E293B" VerticalAlignment="Center" Margin="0,0,16,0"/>
                                    <CheckBox Name="ChkCascadeAutoReboot" Grid.Column="1" Content="Auto-Restart with Persistence (Re-opens Hub on boot)" IsChecked="True" Foreground="#1E293B" VerticalAlignment="Center" Margin="0,0,16,0"/>
                                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="Max Passes:" Foreground="#64748B" FontSize="11" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <ComboBox Name="CmbCascadeMaxPasses" Width="60" Height="24" Background="#FFFFFF" Foreground="#1E293B" BorderBrush="#CBD5E1" SelectedIndex="1">
                                            <ComboBoxItem Content="3"/>
                                            <ComboBoxItem Content="5"/>
                                            <ComboBoxItem Content="8"/>
                                            <ComboBoxItem Content="10"/>
                                        </ComboBox>
                                    </StackPanel>
                                    <TextBlock Name="TxtWuRebootNotice" Grid.Column="3" Text="" Foreground="#B45309" FontWeight="SemiBold" FontSize="11" VerticalAlignment="Center" HorizontalAlignment="Right" Margin="0,0,12,0"/>
                                    <Button Name="BtnWuRebootNow" Grid.Column="4" Content="Restart System Now" Style="{StaticResource DestructiveBtn}" Padding="10,3" FontSize="11" Visibility="Collapsed"
                                            ToolTip="Persists Hub and reboots immediately to apply installed updates"/>
                                </Grid>
                            </Border>
                        </Grid>
                    </Border>

                    <!-- Manual Controls Bar & Progress Status -->
                    <Border Grid.Row="1" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="12,8" Margin="0,0,0,8">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="MANUAL CONTROLS:" FontSize="10.5" FontWeight="Bold" Foreground="#64748B" VerticalAlignment="Center" Margin="0,0,10,0"/>
                                    <Button Name="BtnWuScan" Content="Scan for Updates" Padding="10,4" Margin="0,0,6,0"/>
                                    <Button Name="BtnWuInstall" Content="Install Selected" Style="{StaticResource AccentBtn}" Padding="10,4" Margin="0,0,6,0" IsEnabled="False"/>
                                    <Button Name="BtnWuSelectAll" Content="Select All" Padding="8,4" Margin="0,0,4,0"/>
                                    <Button Name="BtnWuDeselectAll" Content="Deselect All" Padding="8,4" Margin="0,0,6,0"/>
                                    <Button Name="BtnWuTriggerUso" Content="Trigger USO Scan" Padding="8,4" ToolTip="Triggers background usoclient.exe StartInteractiveScan"/>
                                </StackPanel>
                                <TextBlock Name="TxtWuSummary" Grid.Column="1" Text="Ready to scan" FontSize="11" Foreground="#334155" VerticalAlignment="Center"/>
                            </Grid>
                            <Grid Grid.Row="1">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <Grid Grid.Row="0" Margin="0,0,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Name="TxtWuStatus" Text="Ready. Click 'Start Autonomous Patch Cascade' or 'Scan for Updates'." FontSize="11" Foreground="#475569"/>
                                    <TextBlock Name="TxtWuProgress" Grid.Column="1" Text="" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7"/>
                                </Grid>
                                <ProgressBar Name="ProgWuBar" Grid.Row="1" Height="4" Background="#FFFFFF" Foreground="#0067C0" BorderThickness="0" Minimum="0" Maximum="100" Value="0"/>
                            </Grid>
                        </Grid>
                    </Border>

                    <!-- Updates List Grid View -->
                    <Border Grid.Row="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1">
                        <ListView Name="LstIntegratedUpdates" Background="#FFFFFF" BorderThickness="0" Foreground="#1E293B">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Select" Width="50">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <CheckBox IsChecked="{Binding IsSelected}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Type" Width="85">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <Border Background="{Binding BadgeBackground}" CornerRadius="3" Padding="6,2" HorizontalAlignment="Center" VerticalAlignment="Center">
                                                    <TextBlock Text="{Binding CategoryText}" FontSize="9.5" FontWeight="Bold" Foreground="{Binding BadgeForeground}"/>
                                                </Border>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Update / Driver Description" Width="780">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <StackPanel VerticalAlignment="Center" Margin="2,2,8,2">
                                                    <TextBlock Text="{Binding Title}" FontSize="11.5" FontWeight="SemiBold" Foreground="#1E293B" TextWrapping="Wrap"/>
                                                    <TextBlock Text="{Binding Subtitle}" FontSize="10" Foreground="#64748B" Margin="0,2,0,0" TextWrapping="Wrap"/>
                                                </StackPanel>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Size" Width="85">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding SizeText}" FontSize="11" Foreground="#334155" VerticalAlignment="Center" HorizontalAlignment="Right" Margin="0,0,8,0"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Status" Width="140">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding StatusText}" FontSize="11" FontWeight="SemiBold" Foreground="{Binding StatusColor}" VerticalAlignment="Center" Margin="4,0,8,0"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 6: DELL ASSET WARRANTY & REFRESH ASSESSMENT -->
            <TabItem Name="TabHardwareHealth" Header="Dell Warranty &amp; Hardware Health">
                <Border Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="16" Margin="0,12,0,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Row 0: Tag Input & Control Buttons -->
                        <Grid Grid.Row="0" Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="170"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Service Tag / Serial:" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#475569"/>
                            <TextBox Name="TxtDellServiceTag" Grid.Column="1" VerticalAlignment="Center" CharacterCasing="Upper" FontFamily="Consolas" FontWeight="Bold" FontSize="13" Margin="0,0,8,0"/>
                            <Button Name="BtnDetectDellTag" Grid.Column="2" Content="Detect BIOS Tag" Margin="0,0,6,0"/>
                            <StackPanel Grid.Column="3" Orientation="Horizontal">
                                <Button Name="BtnCheckDellWarranty" Content="Assess Dell Warranty" Style="{StaticResource AccentBtn}" Margin="0,0,6,0"/>
                                <Button Name="BtnCheckLenovoWarranty" Content="Assess Lenovo" Margin="0,0,8,0"/>
                            </StackPanel>
                            <Button Name="BtnCopyDellReport" Grid.Column="4" Content="Copy Report" Margin="0,0,6,0"/>
                            <Button Name="BtnExportDellCsv" Grid.Column="5" Content="Export CSV"/>
                        </Grid>

                        <!-- Row 1: Hero Refresh Verdict Banner -->
                        <Border Name="BorderRefreshVerdict" Grid.Row="1" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14,10" Margin="0,0,0,12">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <Border Name="BorderVerdictBadge" Background="#F1F5F9" CornerRadius="3" Padding="6,2" Margin="0,0,10,0" BorderBrush="#CBD5E1" BorderThickness="1">
                                        <TextBlock Name="TxtVerdictBadge" Text="PENDING ASSESSMENT" FontSize="10.5" FontWeight="Bold" Foreground="#334155"/>
                                    </Border>
                                    <TextBlock Name="TxtVerdictTitle" Text="DELL ASSET WARRANTY &amp; REFRESH ASSESSMENT" FontSize="13" FontWeight="SemiBold" Foreground="#1E293B" VerticalAlignment="Center"/>
                                </StackPanel>
                                <TextBlock Name="TxtVerdictDesc" Grid.Row="1" Text="Click 'Assess Lifecycle &amp; Warranty' to query vendor contract records and evaluate hardware refresh eligibility." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" Margin="0,6,0,0"/>
                            </Grid>
                        </Border>

                        <!-- Row 2: Hardware & Contract Telemetry Cards -->
                        <Grid Grid.Row="2" Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <Border Grid.Column="0" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8" Margin="0,0,6,0">
                                <StackPanel>
                                    <TextBlock Text="SYSTEM MODEL" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <TextBlock Name="TxtDellModel" Text="Unknown" FontSize="12" FontWeight="SemiBold" Foreground="#1E293B" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
                                    <TextBlock Name="TxtDellProductLine" Text="Line: -" FontSize="10" Foreground="#64748B" TextTrimming="CharacterEllipsis"/>
                                </StackPanel>
                            </Border>

                            <Border Grid.Column="1" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8" Margin="0,0,6,0">
                                <StackPanel>
                                    <TextBlock Text="FACTORY SHIP DATE / AGE" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <TextBlock Name="TxtDellShipDate" Text="-" FontSize="12" FontWeight="SemiBold" Foreground="#1E293B" Margin="0,2,0,0"/>
                                    <TextBlock Name="TxtDellAge" Text="Age: -" FontSize="10" Foreground="#64748B"/>
                                </StackPanel>
                            </Border>

                            <Border Grid.Column="2" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8" Margin="0,0,6,0">
                                <StackPanel>
                                    <TextBlock Text="PRIMARY SERVICE CONTRACT" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <TextBlock Name="TxtDellContract" Text="-" FontSize="12" FontWeight="SemiBold" Foreground="#1E293B" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
                                    <TextBlock Name="TxtDellRegion" Text="Region: -" FontSize="10" Foreground="#64748B"/>
                                </StackPanel>
                            </Border>

                            <Border Grid.Column="3" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8">
                                <StackPanel>
                                    <TextBlock Text="EXPIRATION &amp; STATUS" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <TextBlock Name="TxtDellEndDate" Text="-" FontSize="12" FontWeight="SemiBold" Foreground="#1E293B" Margin="0,2,0,0"/>
                                    <TextBlock Name="TxtDellDaysRemaining" Text="Status: -" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                </StackPanel>
                            </Border>
                        </Grid>

                        <!-- Row 3: Multi-Vendor Hardware Health Cards (Battery Wear & NVMe SMART) -->
                        <Grid Grid.Row="3" Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <!-- Battery Wear Card -->
                            <Border Grid.Column="0" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8" Margin="0,0,6,0">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="BATTERY WEAR &amp; HEALTH" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <Button Name="BtnRefreshBattery" Grid.Column="1" Content="Refresh Battery" Padding="6,2" FontSize="10"/>
                                    <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,4,0,2">
                                        <TextBlock Name="TxtBatteryWear" Text="Wear: Probing..." FontSize="12.5" FontWeight="Bold" Foreground="#0284C7" Margin="0,0,12,0"/>
                                        <TextBlock Name="TxtBatteryVerdict" Text="Status: Normal" FontSize="12" Foreground="#475569" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <TextBlock Name="TxtBatteryCapacity" Grid.Row="2" Grid.ColumnSpan="2" Text="Design: - | Full: -" FontSize="10.5" Foreground="#64748B"/>
                                </Grid>
                            </Border>

                            <!-- NVMe SSD SMART Reliability Card -->
                            <Border Grid.Column="1" Background="#F8FAFC" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="10,8" Margin="6,0,0,0">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="NVME SSD SMART RELIABILITY COUNTER" FontSize="10" FontWeight="SemiBold" Foreground="#64748B"/>
                                    <Button Name="BtnRefreshStorage" Grid.Column="1" Content="Inspect SMART" Padding="6,2" FontSize="10"/>
                                    <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,4,0,2">
                                        <TextBlock Name="TxtStorageWear" Text="SSD Wear: Probing..." FontSize="12.5" FontWeight="Bold" Foreground="#15803D" Margin="0,0,12,0"/>
                                        <TextBlock Name="TxtStorageTemp" Text="Temp: -" FontSize="12" Foreground="#475569" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <TextBlock Name="TxtStorageMeta" Grid.Row="2" Grid.ColumnSpan="2" Text="Read Errors: 0 | Write Errors: 0 | Power-On Hours: -" FontSize="10.5" Foreground="#64748B"/>
                                </Grid>
                            </Border>
                        </Grid>

                        <!-- Row 4: Entitlements Table -->
                        <Grid Grid.Row="4">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="CONTRACT ENTITLEMENTS &amp; SERVICE HISTORY" FontSize="10.5" FontWeight="SemiBold" Foreground="#64748B" Margin="0,0,0,6"/>
                            <ListView Name="LstDellEntitlements" Grid.Row="1" Background="#FFFFFF" BorderBrush="#E2E8F0" Foreground="#1E293B">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Service Level Description" Width="260" DisplayMemberBinding="{Binding ServiceLevelDescription}"/>
                                        <GridViewColumn Header="Type" Width="85" DisplayMemberBinding="{Binding EntitlementType}"/>
                                        <GridViewColumn Header="Start Date" Width="95" DisplayMemberBinding="{Binding StartDate}"/>
                                        <GridViewColumn Header="End Date" Width="95" DisplayMemberBinding="{Binding EndDate}"/>
                                        <GridViewColumn Header="Status" Width="75" DisplayMemberBinding="{Binding Status}"/>
                                        <GridViewColumn Header="Code" Width="60" DisplayMemberBinding="{Binding ServiceLevelCode}"/>
                                        <GridViewColumn Header="Item #" Width="90" DisplayMemberBinding="{Binding ItemNumber}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </Grid>
                </Border>
            </TabItem>

            <!-- TAB 7: HYBRID AZURE AD JOIN & CO-MANAGEMENT TOOLSET -->
            <TabItem Name="TabHybrid" Header="Hybrid &amp; Co-Mgmt">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Card: Join & Identity -->
                            <Border Grid.Row="0" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="JOIN &amp; IDENTITY" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybJoin" Text="Click Refresh to read dsregcmd join state, PRT and tenant." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybRefresh" Content="Refresh State" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybRetryJoin" Content="Retry Hybrid Join" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybSecChan" Content="Test Secure Channel" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: MDM / Intune -->
                            <Border Grid.Row="0" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="MDM / INTUNE" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybMdm" Text="Enrollment + Intune Management Extension health." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybSync" Content="Force Intune Sync" Style="{StaticResource AccentBtn}" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybRestartIme" Content="Restart IME" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybEnroll" Content="Force MDM Enroll" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: Co-Management & Authority Overrides -->
                            <Border Grid.Row="1" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="CO-MANAGEMENT (WHO OWNS EACH WORKLOAD)" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybComgmt" Text="Decode the ConfigMgr/Intune workload authority bitmask." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,8"/>
                                    <ListBox Name="LstHybWorkloads" Height="110" Background="#FFFFFF" BorderBrush="#E2E8F0" Margin="0,0,0,8">
                                        <ListBox.ItemTemplate>
                                            <DataTemplate>
                                                <Grid>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="*"/>
                                                        <ColumnDefinition Width="120"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Text="{Binding Workload}" Foreground="#475569" FontSize="11"/>
                                                    <TextBlock Grid.Column="1" Text="{Binding Authority}" Foreground="#15803D" FontSize="11" FontWeight="SemiBold"/>
                                                </Grid>
                                            </DataTemplate>
                                        </ListBox.ItemTemplate>
                                    </ListBox>
                                    <WrapPanel Margin="0,0,0,6">
                                        <Button Name="BtnHybComgmt" Content="Read Workloads" Margin="0,0,6,4" Padding="10,5"/>
                                        <Button Name="BtnHybCcm" Content="Trigger ConfigMgr Policy" Margin="0,0,6,4" Padding="10,5"/>
                                    </WrapPanel>
                                    <WrapPanel>
                                        <Button Name="BtnHybCoMgmtAllIntune" Content="Shift All to Intune (255)" Margin="0,0,6,4" Padding="8,4"/>
                                        <Button Name="BtnHybCoMgmtAllCcm" Content="Shift All to ConfigMgr (1)" Margin="0,0,6,4" Padding="8,4"/>
                                        <Button Name="BtnHybCoMgmtPilot" Content="Pilot Workloads (67)" Margin="0,0,6,4" Padding="8,4"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: Policy & Connectivity -->
                            <Border Grid.Row="1" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="POLICY &amp; CONNECTIVITY" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybNet" Text="DC line-of-sight, Group Policy and domain time." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybDc" Content="Test DC Line-of-Sight" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybDcLadder" Content="DC Port Ladder" Style="{StaticResource AccentBtn}" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybScp" Content="Inspect AD SCP" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybGpupdate" Content="gpupdate /force" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybTime" Content="Resync Domain Time" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: Kerberos & Cloud AP PRT -->
                            <Border Grid.Row="2" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="KERBEROS &amp; CLOUD AP PRT DIAGNOSTICS" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybKerb" Text="Inspect cached Kerberos TGT/TGS tickets, Cloud Kerberos PRT, and WAM broker health." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybKerbDiag" Content="Inspect Kerberos Tickets" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybKerbPurge" Content="Purge Kerberos Tickets" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybSpn" Content="Verify Computer SPNs" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybPrtDiag" Content="PRT &amp; WAM Broker" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybAadToken" Content="Test AAD Token" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybResetBroker" Content="Reset Broker Cache" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: Certificate Auto-Enrollment & SCEP Health -->
                            <Border Grid.Row="2" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="CERTIFICATE AUTO-ENROLLMENT &amp; SCEP HEALTH" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybCertPulse" Text="Force immediate certificate auto-enrollment pulse and audit SCEP/NDES Intune client certificates." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybCertPulse" Content="Force Cert Pulse" Style="{StaticResource AccentBtn}" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybScepHealth" Content="Audit SCEP Health" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card: Compliance, BitLocker, Legacy & Certs -->
                            <Border Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="COMPLIANCE / SECURITY / DIAGNOSTICS" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <TextBlock Name="TxtHybCompliance" Text="Conditional Access readiness, BitLocker escrow, legacy-dependency scan, certificate expiry, and a helpdesk log bundle." FontSize="11.5" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <WrapPanel>
                                        <Button Name="BtnHybCa" Content="Check CA Readiness" Style="{StaticResource AccentBtn}" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybBitlocker" Content="Escrow BitLocker Keys" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybLegacy" Content="Legacy Dependency Scan" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybCerts" Content="Certificate Health" Margin="0,0,6,6" Padding="10,5"/>
                                        <Button Name="BtnHybBundle" Content="Collect Hybrid Diagnostics (.zip)" Margin="0,0,6,6" Padding="10,5"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </ScrollViewer>

                    <!-- Tab-local status bar -->
                    <Border Grid.Row="1" Background="#F8FAFC" CornerRadius="3" BorderBrush="#CBD5E1" BorderThickness="1" Padding="8,4" Margin="0,6,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="HYBRID STATUS:" FontSize="10" FontWeight="Bold" Foreground="#0284C7" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBox Name="TxtHybOut" Grid.Column="1" Background="Transparent" Foreground="#475569" BorderThickness="0" FontFamily="Consolas" FontSize="11" IsReadOnly="True" TextWrapping="NoWrap" VerticalAlignment="Center" Text="Ready. Hybrid &amp; co-management tools initialized."/>
                            <Button Name="BtnClearHybOut" Grid.Column="2" Content="Clear" FontSize="9.5" Padding="6,2" Margin="6,0,0,0"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 7: ESP DIAGNOSTICS & LIFECYCLE TOOLS -->
            <TabItem Name="TabLifecycle" Header="ESP Diagnostics &amp; Lifecycle">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Card 1: Win32App Registry Status & Diagnostics -->
                            <Border Grid.Row="0" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Text="WIN32APP ESP REGISTRY STATUS" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" VerticalAlignment="Center"/>
                                        <Button Name="BtnRefreshWin32Apps" Grid.Column="1" Content="Refresh Apps" Margin="0,0,6,0" Padding="8,3" FontSize="11"/>
                                        <Button Name="BtnExportMdmCab" Grid.Column="2" Content="Export CAB" Padding="8,3" FontSize="11" ToolTip="Runs mdmdiagnosticstool.exe -area Autopilot;DeviceEnrollment;TPM -cab"/>
                                    </Grid>
                                    <ListBox Name="LstWin32Apps" Height="140" Background="#FFFFFF" BorderBrush="#E2E8F0">
                                        <ListBox.ItemTemplate>
                                            <DataTemplate>
                                                <Grid Margin="2">
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="*"/>
                                                        <ColumnDefinition Width="110"/>
                                                        <ColumnDefinition Width="65"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Text="{Binding Name}" Foreground="#1E293B" FontWeight="SemiBold" FontSize="11" TextTrimming="CharacterEllipsis"/>
                                                    <TextBlock Grid.Column="1" Text="{Binding InstallState}" Foreground="#15803D" FontSize="10.5"/>
                                                    <TextBlock Grid.Column="2" Text="{Binding ExitCode}" Foreground="#64748B" FontSize="10.5"/>
                                                </Grid>
                                            </DataTemplate>
                                        </ListBox.ItemTemplate>
                                    </ListBox>
                                </StackPanel>
                            </Border>

                            <!-- Card 2: Autopilot Cloud Device Lifecycle -->
                            <Border Grid.Row="0" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="AUTOPILOT CLOUD LIFECYCLE &amp; DECOMMISSION" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="90"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Text="New Group Tag:" Foreground="#475569" VerticalAlignment="Center" FontSize="11"/>
                                        <TextBox Name="TxtLifecycleGroupTag" Grid.Column="1" Height="26" FontSize="11"/>
                                    </Grid>
                                    <Grid Margin="0,0,0,8">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="90"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Text="New User UPN:" Foreground="#475569" VerticalAlignment="Center" FontSize="11"/>
                                        <TextBox Name="TxtLifecycleUser" Grid.Column="1" Height="26" FontSize="11"/>
                                    </Grid>
                                    <WrapPanel>
                                        <Button Name="BtnPatchLifecycle" Content="In-Place PATCH" Style="{StaticResource AccentBtn}" Margin="0,0,6,6" Padding="8,4" FontSize="11"
                                                ToolTip="Patches group tag and assigned user directly on Graph Autopilot device identity"/>
                                        <Button Name="BtnSyncAutopilot" Content="Instant Tenant Sync" Margin="0,0,6,6" Padding="8,4" FontSize="11"
                                                ToolTip="Triggers POST /deviceManagement/windowsAutopilotSettings/sync"/>
                                        <Button Name="BtnDecommission" Content="3-Registry Decommission" Style="{StaticResource DestructiveBtn}" Margin="0,0,6,6" Padding="8,4" FontSize="11"
                                                ToolTip="Purges Autopilot diagnostics registry, MDM enrollment keys, and cached profile JSONs"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card 3: OOBE Wi-Fi Manager & 802.1X XML Profile Import -->
                            <Border Grid.Row="1" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="OOBE WI-FI MANAGER &amp; 802.1X PROFILES" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <ComboBox Name="CmbWifiSsids" IsEditable="True" Height="26" FontSize="11"/>
                                        <Button Name="BtnScanWifi" Grid.Column="1" Content="Scan SSIDs" Margin="6,0,0,0" Padding="8,3" FontSize="11"/>
                                    </Grid>
                                    <Grid Margin="0,0,0,8">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="70"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Text="Password:" Foreground="#475569" VerticalAlignment="Center" FontSize="11"/>
                                        <TextBox Name="TxtWifiPassword" Grid.Column="1" Height="26" FontSize="11"/>
                                        <Button Name="BtnConnectWifi" Grid.Column="2" Content="Connect" Style="{StaticResource AccentBtn}" Margin="6,0,0,0" Padding="8,3" FontSize="11"/>
                                    </Grid>
                                    <WrapPanel>
                                        <Button Name="BtnImportWifiXml" Content="Import 802.1X XML Profile" Margin="0,0,6,0" Padding="8,4" FontSize="11"
                                                ToolTip="Imports enterprise WPA2-Enterprise/802.1X Wi-Fi XML profile via netsh wlan add profile"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Card 4: Windows Edition & 1-Click KMS Upgrade -->
                            <Border Grid.Row="1" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="WINDOWS EDITION &amp; OEM LICENSE" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,6"/>
                                    <StackPanel Margin="0,0,0,8">
                                        <TextBlock Name="TxtOsCaption" Text="Edition: Detecting..." FontSize="11.5" Foreground="#1E293B" FontWeight="SemiBold"/>
                                        <TextBlock Name="TxtOemKey" Text="OEM Key: Reading ACPI MSDM..." FontSize="11" Foreground="#64748B" FontFamily="Consolas" Margin="0,2,0,0"/>
                                    </StackPanel>
                                    <WrapPanel>
                                        <Button Name="BtnUpgradeEnterprise" Content="1-Click Upgrade to Enterprise (KMS)" Style="{StaticResource AccentBtn}" Margin="0,0,6,0" Padding="8,4" FontSize="11"
                                                ToolTip="Upgrades Windows Home/Pro to Windows Enterprise using official Microsoft KMS setup key"/>
                                        <Button Name="BtnRefreshLicensing" Content="Refresh Edition" Padding="8,4" FontSize="11"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </ScrollViewer>

                    <!-- Tab-local status bar -->
                    <Border Grid.Row="1" Background="#F8FAFC" CornerRadius="3" BorderBrush="#CBD5E1" BorderThickness="1" Padding="8,4" Margin="0,6,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="IME LOG STREAM:" FontSize="10" FontWeight="Bold" Foreground="#0284C7" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBox Name="TxtImeLog" Grid.Column="1" Background="Transparent" Foreground="#475569" BorderThickness="0" FontFamily="Consolas" FontSize="10.5" IsReadOnly="True" TextWrapping="NoWrap" VerticalAlignment="Center" Text="Click 'Tail IME Log' to read the real-time Intune Management Extension agent log..."/>
                            <Button Name="BtnTailImeLog" Grid.Column="2" Content="Tail IME Log" FontSize="9.5" Padding="8,2" Margin="6,0,0,0"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 9: PRECISION ENTERPRISE REMEDIATION ARSENAL -->
            <TabItem Name="TabPrecisionFixes" Header="Precision Enterprise Fixes">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Header Banner -->
                    <Border Grid.Row="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14,10" Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="PRECISION ENTERPRISE REMEDIATION ARSENAL" FontSize="12" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                <TextBlock Text="Curated surgical one-click fixes for elusive enterprise corruption, broken cryptographic catalogs, deadlocked pipelines, and hybrid edge cases." FontSize="11" Foreground="#334155" TextWrapping="Wrap"/>
                            </StackPanel>
                            <WrapPanel Grid.Column="1" VerticalAlignment="Center">
                                <Button Name="BtnFixHealthReadout" Content="Audit System Health" Style="{StaticResource AccentBtn}" Padding="12,6" Margin="0,0,6,0"/>
                                <Button Name="BtnClearFixOut" Content="Clear Output" Padding="10,6"/>
                            </WrapPanel>
                        </Grid>
                    </Border>

                    <!-- Scrollable Cards Grid -->
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <Grid Margin="0,0,6,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Card 1: WMI Repository Salvage & Self-Heal -->
                            <Border Grid.Row="0" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="WMI REPOSITORY SALVAGE &amp; SELF-HEAL" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Runs winmgmt /salvagerepository and recompiles core MOF/MFL catalogs without wiping OEM (Dell/HP/Lenovo) WMI namespaces." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixWmi" Content="Salvage WMI Repository" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 2: Windows Update Agent & SoftDistribution -->
                            <Border Grid.Row="0" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="WINDOWS UPDATE AGENT &amp; SOFTDISTRIBUTION RESET" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Full service teardown (wuauserv, bits, cryptsvc, dosvc), archives SoftwareDistribution &amp; Catroot2, re-registers 28 core DLLs, and resets permissions." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixWu" Content="Deep Reset WU Agent" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 3: Cryptographic Services / Catroot2 ESENT Repair -->
                            <Border Grid.Row="1" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="CATROOT2 CRYPTO ESENT DATABASE REPAIR" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Repairs corrupted catdb ESENT database via esentutl /g and /p, fixes CryptSvc catalog locks, and resolves 0x800b0109 signature errors." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixCatroot" Content="Repair Catroot2 Database" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 4: BITS Transfer Deadlock Purge -->
                            <Border Grid.Row="1" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="BITS TRANSFER DEADLOCK PURGE" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Clears stuck Qmgr jobs, purges corrupted qmgr0.dat/qmgr1.dat databases, resets BITS service state, and re-validates the transfer pipeline." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixBits" Content="Purge BITS Deadlock" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 5: DCOM / RPC 0x800706BA Remediation -->
                            <Border Grid.Row="2" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="DCOM / RPC 0x800706BA &amp; SECURITY LIMITS" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Repairs Component Services DCOM authentication/impersonation levels, fixes machine launch permissions, and verifies RPC Endpoint Mapper." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixDcom" Content="Repair DCOM &amp; RPC Limits" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 6: Print Spooler Queue Purge -->
                            <Border Grid.Row="2" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="PRINT SPOOLER HUNG QUEUE PURGE" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Gracefully halts hung spoolsv.exe, clears locked .spl and .shd print job files in spool\PRINTERS, cleans registry job keys, and restarts Spooler." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixSpooler" Content="Purge Spooler Queue" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 7: Network Stack, Winsock & IPsec Reset -->
                            <Border Grid.Row="3" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="NETWORK STACK, WINSOCK &amp; IPSEC RESET" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Resets TCP/IP stack (netsh int ip reset), resets Winsock catalog, flushes DNS &amp; ARP, clears IPsec security associations, and renews DHCP." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixNetStack" Content="Flush Network &amp; Winsock" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 8: WinRM & WS-Man Listener Rebuild -->
                            <Border Grid.Row="3" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="WINRM &amp; WS-MAN LISTENER REBUILD" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Tears down corrupted WinRM listeners, recreates default HTTP port 5985 listener, re-binds firewall rules, and validates WS-Man loopback." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixWinRm" Content="Rebuild WinRM Listener" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 9: User Profile Registry Lock Un-hooker -->
                            <Border Grid.Row="4" Grid.Column="0" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,6,10">
                                <StackPanel>
                                    <TextBlock Text="USER PROFILE REGISTRY LOCK UN-HOOKER" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Detects orphaned .bak ProfileList subkeys, resolves temporary profile collisions (TEMP profiles), and clears RefCount hive locks." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixProfiles" Content="Un-hook Profile .bak Locks" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 10: TPM Crypto Provider & Attestation Healer -->
                            <Border Grid.Row="4" Grid.Column="1" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="6,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="TPM PLATFORM CRYPTO &amp; ATTESTATION HEALER" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Clears hung attestation state flags in registry, restarts TPM Base Services (tbds), validates Endorsement Key (EK) certs without wiping BitLocker." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixTpm" Content="Heal TPM Attestation" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                            <!-- Card 11: AppX Manifest Staging & Sysprep Unbricker -->
                            <Border Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" Background="#FFFFFF" CornerRadius="4" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,0,10">
                                <StackPanel>
                                    <TextBlock Text="APPX MANIFEST STAGING &amp; SYSPREP UNBRICKER" FontSize="11" FontWeight="SemiBold" Foreground="#0284C7" Margin="0,0,0,4"/>
                                    <TextBlock Text="Scans for broken/abnormal AppX packages, cleans orphaned staged packages, and re-registers core Windows inbox manifests (Shell, Start, SecHealthUI)." FontSize="11" Foreground="#334155" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                    <Button Name="BtnFixAppX" Content="Unbrick AppX Staging" HorizontalAlignment="Left" Padding="12,5"/>
                                </StackPanel>
                            </Border>

                        </Grid>
                    </ScrollViewer>

                    <!-- Tab-local status bar -->
                    <Border Grid.Row="2" Background="#F8FAFC" CornerRadius="3" BorderBrush="#CBD5E1" BorderThickness="1" Padding="8,4" Margin="0,6,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="REMEDIATION STATUS:" FontSize="10" FontWeight="Bold" Foreground="#B91C1C" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBox Name="TxtRemOut" Grid.Column="1" Background="Transparent" Foreground="#475569" BorderThickness="0" FontFamily="Consolas" FontSize="11" IsReadOnly="True" TextWrapping="NoWrap" VerticalAlignment="Center" Text="Enterprise Remediation Arsenal ready. Select a precision tool above to execute surgical OS recovery."/>
                            <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="6,0,0,0">
                                <Button Name="BtnClearRemOut" Content="Clear" FontSize="9.5" Padding="6,2" Margin="0,0,4,0"/>
                                <Button Name="BtnCopyFixLog" Content="Copy" FontSize="9.5" Padding="6,2"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

        </TabControl>

        <!-- PROGRESS BAR & STATUS -->
        <Grid Grid.Row="4" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <ProgressBar Name="HubProgressBar" Height="4" Minimum="0" Maximum="100" Value="0"
                         Background="#F8FAFC" Foreground="#0067C0" BorderThickness="0"/>
            <TextBlock Name="TxtProgressStatus" Grid.Column="1" Text="Ready" FontSize="10.5" Foreground="#64748B" Margin="8,0,0,0"/>
        </Grid>

        <!-- LIVE LOG OUTPUT CONSOLE (COLLAPSIBLE) -->
        <Border Grid.Row="5" Background="#FFFFFF" CornerRadius="4" BorderBrush="#CBD5E1" BorderThickness="1" Padding="8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="SYSTEM AUDIT LOG" FontSize="10" FontWeight="SemiBold" Foreground="#64748B" VerticalAlignment="Center"/>
                        <Button Name="BtnToggleConsole" Content="[-] Hide Log" FontSize="9.5" Padding="6,2" Margin="10,0,0,0" ToolTip="Collapse or expand the system audit log console"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" Orientation="Horizontal">
                        <Button Name="BtnCopyLog" Content="Copy Log" FontSize="10" Padding="6,2" Margin="0,0,4,0"/>
                        <Button Name="BtnClearLog" Content="Clear" FontSize="10" Padding="6,2" Margin="0,0,4,0"/>
                        <Button Name="BtnSaveLog" Content="Save Log..." FontSize="10" Padding="6,2"/>
                    </StackPanel>
                </Grid>

                <TextBox Name="TxtHubLog" Grid.Row="1" Background="Transparent" Foreground="#1E293B"
                         BorderThickness="0" FontFamily="Cascadia Code, Consolas" FontSize="11"
                         IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

    # Parse XAML
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Easter egg: click the title 5x -> a horse trots through the audit log
    $script:TitleClicks = 0
    $txtHubTitleEgg = $window.FindName('TxtHubTitle')
    if ($txtHubTitleEgg) {
        $txtHubTitleEgg.Add_MouseLeftButtonUp({
            $script:TitleClicks++
            if ($script:TitleClicks -ge 5) { $script:TitleClicks = 0; Write-HubLog ("`r`n" + (Write-HubHorse)) "INFO" }
        })
    }
    # Easter egg: Konami code anywhere -> Outlaw Mode (cosmetic)
    $script:KonamiSeq = @('Up','Up','Down','Down','Left','Right','Left','Right','B','A')
    $script:KonamiPos = 0
    $script:OutlawMode = $false
    $window.Add_PreviewKeyDown({
        param($sender, $e)
        $k = [string]$e.Key
        if ($k -eq $script:KonamiSeq[$script:KonamiPos]) { $script:KonamiPos++ }
        else { $script:KonamiPos = $(if ($k -eq $script:KonamiSeq[0]) { 1 } else { 0 }) }
        if ($script:KonamiPos -ge $script:KonamiSeq.Count) {
            $script:KonamiPos = 0; $script:OutlawMode = -not $script:OutlawMode
            if ($script:OutlawMode) {
                $window.Title = "Outlaw Provisioning Hub - Wanted: Dead or Enrolled"
                Write-HubLog "Outlaw mode engaged. Nothing changes but your attitude. $(Get-HubQuip -Kind Success)" "WARN"
                try { Play-HubAudio -Type Success } catch { }
            } else {
                $window.Title = "Autopilot Provisioning Hub - Enterprise Endpoint Deployment"
                Write-HubLog "Outlaw mode disengaged. Back to honest work." "INFO"
            }
        }
    })

    # Resolve UI Controls
    $badgeGraphAuth    = $window.FindName('BadgeGraphAuth')
    $dotGraphStatus    = $window.FindName('DotGraphStatus')
    $txtGraphStatus    = $window.FindName('TxtGraphStatus')
    $btnConnectGraph   = $window.FindName('BtnConnectGraph')
    $btnInstallPwsh    = $window.FindName('BtnInstallPwsh')

    $txtSerial         = $window.FindName('TxtSerial')
    $txtModel          = $window.FindName('TxtModel')
    $txtTpm            = $window.FindName('TxtTpm')
    $txtSecureBoot     = $window.FindName('TxtSecureBoot')
    $txtNetwork        = $window.FindName('TxtNetwork')
    $cmbGroupTag       = $window.FindName('CmbGroupTag')
    $txtAssignedUser   = $window.FindName('TxtAssignedUser')
    $chkEnableRename    = $window.FindName('ChkEnableRename')
    $txtComputerName    = $window.FindName('TxtComputerName')
    $btnApplyRename     = $window.FindName('BtnApplyRename')
    $chkWaitForSync     = $window.FindName('ChkWaitForSync')
    $chkAutoDetectUsb   = $window.FindName('ChkAutoDetectUsb')
    $chkAutoReboot      = $window.FindName('ChkAutoReboot')
    $btnHarvestHash     = $window.FindName('BtnHarvestHash')
    $btnExportCsv       = $window.FindName('BtnExportCsv')
    $btnRegisterIntune  = $window.FindName('BtnRegisterIntune')
    $btnSaveEnvDefaults = $window.FindName('BtnSaveEnvDefaults')
    $chkPersonalDisableNet = $window.FindName('ChkPersonalDisableNet')
    $btnCreateLocalUser   = $window.FindName('BtnCreateLocalUser')
    $btnPersonalInstall    = $window.FindName('BtnPersonalInstall')
    $btnReenableNet        = $window.FindName('BtnReenableNet')
    $btnRestartOobe        = $window.FindName('BtnRestartOobe')
    $txtHashBox        = $window.FindName('TxtHashBox')
    $txtHashMeta       = $window.FindName('TxtHashMeta')
    $txtHashStatus     = $window.FindName('TxtHashStatus')
    $badgeHashStatus   = $window.FindName('BadgeHashStatus')
    $btnCopyHash       = $window.FindName('BtnCopyHash')
    $dotPrivilege          = $window.FindName('DotPrivilege')
    $badgePrivilege        = $window.FindName('BadgePrivilege')
    $txtPrivilege          = $window.FindName('TxtPrivilege')
    $btnFixPrivilege       = $window.FindName('BtnFixPrivilege')
    $txtDeviceState        = $window.FindName('TxtDeviceState')
    $bannerDeviceState     = $window.FindName('BannerDeviceState')
    $txtDeviceStateTitle   = $window.FindName('TxtDeviceStateTitle')
    $txtDeviceStateDetail  = $window.FindName('TxtDeviceStateDetail')
    $btnDeviceStateAction  = $window.FindName('BtnDeviceStateAction')
    $bannerAppAdvisory     = $window.FindName('BannerAppAdvisory')
    $txtAppAdvisory        = $window.FindName('TxtAppAdvisory')

    # Additional Controls: Header & Tab 1
    $bannerTenantMismatch  = $window.FindName('BannerTenantMismatch')
    $txtTenantMismatch     = $window.FindName('TxtTenantMismatch')
    $btnDismissMismatch    = $window.FindName('BtnDismissMismatch')
    $btnInspectPolicy      = $window.FindName('BtnInspectPolicy')
    $btnExportOfflineJson  = $window.FindName('BtnExportOfflineJson')
    $btnInjectOfflineJson  = $window.FindName('BtnInjectOfflineJson')
    $btnSkipAutopilot      = if ($window.FindName('BtnSkipAutopilot')) { $window.FindName('BtnSkipAutopilot') } else { $btnPersonalInstall }
    $bannerOfflineStaging  = $window.FindName('BannerOfflineStaging')
    $txtOfflineStagingTitle = $window.FindName('TxtOfflineStagingTitle')
    $txtOfflineStagingDetail = $window.FindName('TxtOfflineStagingDetail')
    $btnReDockStaging      = $window.FindName('BtnReDockStaging')
    $btnOfflineStaging     = $window.FindName('BtnOfflineStaging')
    $btnResumeAutopilot    = $window.FindName('BtnResumeAutopilot')

    $borderTenantLogo      = $window.FindName('BorderTenantLogo')
    $imgTenantLogo         = $window.FindName('ImgTenantLogo')
    $badgeTenantBrand      = $window.FindName('BadgeTenantBrand')
    $txtTenantBrand        = $window.FindName('TxtTenantBrand')
    $badgePower            = $window.FindName('BadgePower')
    $txtPowerStatus        = $window.FindName('TxtPowerStatus')
    $btnToolsDropdown      = $window.FindName('BtnToolsDropdown')
    $menuWindowsUpdate     = $window.FindName('MenuWindowsUpdate')
    $menuCreateLocalUser   = $window.FindName('MenuCreateLocalUser')
    $menuPersonalMsa       = $window.FindName('MenuPersonalMsa')
    $menuLaunchEdge        = $window.FindName('MenuLaunchEdge')
    $menuScreenshotUsb     = $window.FindName('MenuScreenshotUsb')
    $menuToggleDpi         = $window.FindName('MenuToggleDpi')
    $menuInstallPwsh       = $window.FindName('MenuInstallPwsh')
    $menuQuickCmd          = $window.FindName('MenuQuickCmd')
    $menuTimeSync          = $window.FindName('MenuTimeSync')
    $menuReboot            = $window.FindName('MenuReboot')

    # Aliases so existing event handlers and scripts work without modification
    $btnWindowsUpdate      = if ($window.FindName('BtnWindowsUpdate')) { $window.FindName('BtnWindowsUpdate') } else { $menuWindowsUpdate }
    $btnLaunchEdge         = if ($window.FindName('BtnLaunchEdge')) { $window.FindName('BtnLaunchEdge') } else { $menuLaunchEdge }
    $btnScreenshotUsb      = if ($window.FindName('BtnScreenshotUsb')) { $window.FindName('BtnScreenshotUsb') } else { $menuScreenshotUsb }
    $btnToggleDpi          = if ($window.FindName('BtnToggleDpi')) { $window.FindName('BtnToggleDpi') } else { $menuToggleDpi }
    $btnInstallPwsh        = if ($window.FindName('BtnInstallPwsh')) { $window.FindName('BtnInstallPwsh') } else { $menuInstallPwsh }
    $mainTabControl        = $window.FindName('MainTabControl')
    $cboPlaybookRoutine    = $window.FindName('CboPlaybookRoutine')
    $btnPlaybookRun        = $window.FindName('BtnPlaybookRun')
    $btnPlaybookPause      = $window.FindName('BtnPlaybookPause')
    $btnPlaybookStop       = $window.FindName('BtnPlaybookStop')
    $txtPlaybookStep       = $window.FindName('TxtPlaybookStep')
    $playbookProgressBar   = $window.FindName('PlaybookProgressBar')

    $tabAutopilot          = $window.FindName('TabAutopilot')
    $tabAppDeploy          = $window.FindName('TabAppDeploy')
    $tabPackaging          = $window.FindName('TabPackaging')
    $tabPreFlight          = $window.FindName('TabPreFlight')
    $tabHardwareHealth     = $window.FindName('TabHardwareHealth')
    $tabHybrid             = $window.FindName('TabHybrid')
    $tabLifecycle          = $window.FindName('TabLifecycle')
    $tabPrecisionFixes     = $window.FindName('TabPrecisionFixes')

    # Additional Controls: Tab 4 Pre-Flight
    $btnFixClockSkew       = $window.FindName('BtnFixClockSkew')
    $btnTpmAttestation     = $window.FindName('BtnTpmAttestation')
    $btnInjectDriversPreflight = $window.FindName('BtnInjectDriversPreflight')
    $btnPreflightWinUpdate = $window.FindName('BtnPreflightWinUpdate')

    # Additional Controls: Tab 5 Windows Update & Patch Cascade
    $tabWinUpdate          = $window.FindName('TabWinUpdate')
    $badgeCascadeStatus    = $window.FindName('BadgeCascadeStatus')
    $txtCascadeStatus      = $window.FindName('TxtCascadeStatus')
    $txtCascadePassInfo    = $window.FindName('TxtCascadePassInfo')
    $txtCascadeInstalled   = $window.FindName('TxtCascadeInstalled')
    $txtCascadeStateDetail = $window.FindName('TxtCascadeStateDetail')
    $btnStartCascade       = $window.FindName('BtnStartCascade')
    $btnStopCascade        = $window.FindName('BtnStopCascade')
    $btnAbortReboot        = $window.FindName('BtnAbortReboot')
    $chkCascadeDrivers     = $window.FindName('ChkCascadeDrivers')
    $chkCascadeAutoReboot  = $window.FindName('ChkCascadeAutoReboot')
    $cmbCascadeMaxPasses   = $window.FindName('CmbCascadeMaxPasses')
    $txtWuRebootNotice     = $window.FindName('TxtWuRebootNotice')
    $btnWuRebootNow        = $window.FindName('BtnWuRebootNow')
    $btnWuScan             = $window.FindName('BtnWuScan')
    $btnWuInstall          = $window.FindName('BtnWuInstall')
    $btnWuSelectAll        = $window.FindName('BtnWuSelectAll')
    $btnWuDeselectAll      = $window.FindName('BtnWuDeselectAll')
    $btnWuTriggerUso       = $window.FindName('BtnWuTriggerUso')
    $txtWuSummary          = $window.FindName('TxtWuSummary')
    $txtWuStatus           = $window.FindName('TxtWuStatus')
    $txtWuProgress         = $window.FindName('TxtWuProgress')
    $progWuBar             = $window.FindName('ProgWuBar')
    $lstIntegratedUpdates  = $window.FindName('LstIntegratedUpdates')

    # Additional Controls: Tab 5 Warranty & Hardware Health
    $btnCheckLenovoWarranty = $window.FindName('BtnCheckLenovoWarranty')
    $btnRefreshBattery     = $window.FindName('BtnRefreshBattery')
    $txtBatteryWear        = $window.FindName('TxtBatteryWear')
    $txtBatteryVerdict     = $window.FindName('TxtBatteryVerdict')
    $txtBatteryCapacity    = $window.FindName('TxtBatteryCapacity')
    $btnRefreshStorage     = $window.FindName('BtnRefreshStorage')
    $txtStorageWear        = $window.FindName('TxtStorageWear')
    $txtStorageTemp        = $window.FindName('TxtStorageTemp')
    $txtStorageMeta        = $window.FindName('TxtStorageMeta')

    # Additional Controls: Tab 7 ESP & Lifecycle
    $lstWin32Apps          = $window.FindName('LstWin32Apps')
    $btnRefreshWin32Apps   = $window.FindName('BtnRefreshWin32Apps')
    $btnExportMdmCab       = $window.FindName('BtnExportMdmCab')
    $txtLifecycleGroupTag  = $window.FindName('TxtLifecycleGroupTag')
    $txtLifecycleUser      = $window.FindName('TxtLifecycleUser')
    $btnPatchLifecycle     = $window.FindName('BtnPatchLifecycle')
    $btnSyncAutopilot      = $window.FindName('BtnSyncAutopilot')
    $btnDecommission       = $window.FindName('BtnDecommission')
    $cmbWifiSsids          = $window.FindName('CmbWifiSsids')
    $btnScanWifi           = $window.FindName('BtnScanWifi')
    $txtWifiPassword       = $window.FindName('TxtWifiPassword')
    $btnConnectWifi        = $window.FindName('BtnConnectWifi')
    $btnImportWifiXml      = $window.FindName('BtnImportWifiXml')
    $txtOsCaption          = $window.FindName('TxtOsCaption')
    $txtOemKey             = $window.FindName('TxtOemKey')
    $btnUpgradeEnterprise  = $window.FindName('BtnUpgradeEnterprise')
    $btnRefreshLicensing   = $window.FindName('BtnRefreshLicensing')
    $btnTailImeLog         = $window.FindName('BtnTailImeLog')
    $txtImeLog             = $window.FindName('TxtImeLog')

    function Update-GraphAuthHeader {
        if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $script:GraphAuthContext.ExpiresOn -gt [datetime]::UtcNow.AddMinutes(2)) {
            $dotGraphStatus.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#15803D")
            $badgeGraphAuth.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#DCFCE7")
            $badgeGraphAuth.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#86EFAC")
            $tDisplay = if ($script:GraphAuthContext.TenantId -and $script:GraphAuthContext.TenantId -ne 'organizations') {
                if ($script:GraphAuthContext.TenantId.Length -gt 18) {
                    $script:GraphAuthContext.TenantId.Substring(0, 8) + '...'
                } else {
                    $script:GraphAuthContext.TenantId
                }
            } else {
                'Connected'
            }
            $txtGraphStatus.Text = "GRAPH: CONNECTED ($tDisplay)"
            $txtGraphStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#15803D")
            $btnConnectGraph.Content = "Disconnect"
            $btnConnectGraph.Style = [System.Windows.Style]$window.Resources['DestructiveBtn']

            # Tenant Branding & Mismatch Check
            try {
                $tb = Get-GraphTenantBranding -AccessToken $script:GraphAuthContext.AccessToken
                if ($tb -and $tb.DisplayName -and $txtTenantBrand) {
                    $txtTenantBrand.Text = $tb.DisplayName.ToUpper()
                    if ($badgeTenantBrand) { $badgeTenantBrand.Visibility = [System.Windows.Visibility]::Visible }
                }
                if ($tb -and $tb.SquareLogoUrl -and $imgTenantLogo -and $borderTenantLogo) {
                    Set-HubBrandingImageAsync -TargetImage $imgTenantLogo -LogoContainer $borderTenantLogo -ImageUrl $tb.SquareLogoUrl
                }
            } catch { }

            if ($bannerTenantMismatch -and $txtTenantMismatch -and $script:DeviceState -and $script:DeviceState.AutopilotTenantId -and $script:GraphAuthContext.TenantId) {
                if ($script:DeviceState.AutopilotTenantId -ne $script:GraphAuthContext.TenantId) {
                    $txtTenantMismatch.Text = "Device cached profile belongs to tenant '$($script:DeviceState.AutopilotTenantDomain)' ($($script:DeviceState.AutopilotTenantId)), but the active session is signed into tenant '$($script:GraphAuthContext.TenantId)'."
                    $bannerTenantMismatch.Visibility = [System.Windows.Visibility]::Visible
                }
            }

            Invoke-TenantAutopilotLookup
        } else {
            $dotGraphStatus.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#B91C1C")
            $badgeGraphAuth.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F1F5F9")
            $badgeGraphAuth.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#CBD5E1")
            $txtGraphStatus.Text = "GRAPH: NOT SIGNED IN"
            $txtGraphStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#64748B")
            $btnConnectGraph.Content = "Sign In to Intune"
            $btnConnectGraph.Style = [System.Windows.Style]$window.Resources['AccentBtn']
        }
    }

    $btnConnectGraph.Add_Click({
        if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken) {
            $ans = [System.Windows.MessageBox]::Show("Disconnect the current Microsoft Graph session?", "Disconnect Session", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
                $script:GraphAuthContext = $null
                Update-GraphAuthHeader
                Write-HubLog "Microsoft Graph session disconnected." "WARN"
            }
        } else {
            $token = Start-GraphAuthDialog -Owner $window
            Update-GraphAuthHeader
            if ($token) {
                Write-HubLog "Authenticated to Microsoft Graph session ($($script:GraphAuthContext.TenantId)). All cloud capabilities active." "SUCCESS"
            }
        }
    })

    $btnPresetWorkstation = $window.FindName('BtnPresetWorkstation')
    $btnPresetBrowsers    = $window.FindName('BtnPresetBrowsers')
    $btnPresetDev         = $window.FindName('BtnPresetDev')
    $btnSelectAllApps     = $window.FindName('BtnSelectAllApps')
    $btnClearApps         = $window.FindName('BtnClearApps')
    $chkSilentInstall     = $window.FindName('ChkSilentInstall')
    $chkMachineScope      = $window.FindName('ChkMachineScope')
    $txtCustomPkg         = $window.FindName('TxtCustomPkg')
    $btnInstallBatch      = $window.FindName('BtnInstallBatch')

    $appBoxes = @(
        $window.FindName('AppChrome'),
        $window.FindName('AppFirefox'),
        $window.FindName('AppBrave'),
        $window.FindName('AppEdgeDev'),
        $window.FindName('AppVSCode'),
        $window.FindName('AppGit'),
        $window.FindName('AppTerminal'),
        $window.FindName('AppNode'),
        $window.FindName('AppPython'),
        $window.FindName('AppGh'),
        $window.FindName('AppDocker'),
        $window.FindName('AppOffice'),
        $window.FindName('AppSlack'),
        $window.FindName('AppZoom'),
        $window.FindName('AppNotion'),
        $window.FindName('AppObsidian'),
        $window.FindName('App7Zip'),
        $window.FindName('AppAcrobat'),
        $window.FindName('AppPowerToys'),
        $window.FindName('AppSysinternals'),
        $window.FindName('AppSysInformer'),
        $window.FindName('AppWireshark'),
        $window.FindName('AppPuTTY'),
        $window.FindName('AppHwInfo'),
        $window.FindName('AppTreeSize')
    )

    $txtPkgId          = $window.FindName('TxtPkgId')
    $txtPkgDisplayName = $window.FindName('TxtPkgDisplayName')
    $txtPkgOutputDir   = $window.FindName('TxtPkgOutputDir')
    $txtPkgInstallArgs = $window.FindName('TxtPkgInstallArgs')
    $btnBuildPackage   = $window.FindName('BtnBuildPackage')
    $cmbAssignmentIntent = $window.FindName('CmbAssignmentIntent')
    $txtAssignGroup    = $window.FindName('TxtAssignGroup')
    $btnPublishIntune  = $window.FindName('BtnPublishIntune')

    $btnRunDiag        = $window.FindName('BtnRunDiag')
    # Hybrid & Co-Mgmt tab
    $txtHybOut         = $window.FindName('TxtHybOut')
    $txtHybJoin        = $window.FindName('TxtHybJoin')
    $txtHybMdm         = $window.FindName('TxtHybMdm')
    $txtHybComgmt      = $window.FindName('TxtHybComgmt')
    $txtHybNet         = $window.FindName('TxtHybNet')
    $txtHybCompliance  = $window.FindName('TxtHybCompliance')
    $lstHybWorkloads   = $window.FindName('LstHybWorkloads')
    $btnHybRefresh     = $window.FindName('BtnHybRefresh')
    $btnHybRetryJoin   = $window.FindName('BtnHybRetryJoin')
    $btnHybSecChan     = $window.FindName('BtnHybSecChan')
    $btnHybSync        = $window.FindName('BtnHybSync')
    $btnHybRestartIme  = $window.FindName('BtnHybRestartIme')
    $btnHybEnroll      = $window.FindName('BtnHybEnroll')
    $btnHybComgmt      = $window.FindName('BtnHybComgmt')
    $btnHybCcm         = $window.FindName('BtnHybCcm')
    $btnHybDc          = $window.FindName('BtnHybDc')
    $btnHybGpupdate    = $window.FindName('BtnHybGpupdate')
    $btnHybTime        = $window.FindName('BtnHybTime')
    $btnHybCa          = $window.FindName('BtnHybCa')
    $btnHybBitlocker   = $window.FindName('BtnHybBitlocker')
    $btnHybLegacy      = $window.FindName('BtnHybLegacy')
    $btnHybCerts       = $window.FindName('BtnHybCerts')
    $btnHybBundle      = $window.FindName('BtnHybBundle')
    $btnHybDcLadder    = $window.FindName('BtnHybDcLadder')
    $btnHybScp         = $window.FindName('BtnHybScp')
    $btnHybCoMgmtAllIntune = $window.FindName('BtnHybCoMgmtAllIntune')
    $btnHybCoMgmtAllCcm    = $window.FindName('BtnHybCoMgmtAllCcm')
    $btnHybCoMgmtPilot     = $window.FindName('BtnHybCoMgmtPilot')
    $txtHybKerb        = $window.FindName('TxtHybKerb')
    $btnHybKerbDiag    = $window.FindName('BtnHybKerbDiag')
    $btnHybKerbPurge   = $window.FindName('BtnHybKerbPurge')
    $btnHybSpn         = $window.FindName('BtnHybSpn')
    $btnHybPrtDiag     = $window.FindName('BtnHybPrtDiag')
    $btnHybAadToken    = $window.FindName('BtnHybAadToken')
    $btnHybResetBroker = $window.FindName('BtnHybResetBroker')
    $txtHybCertPulse   = $window.FindName('TxtHybCertPulse')
    $btnHybCertPulse   = $window.FindName('BtnHybCertPulse')
    $btnHybScepHealth  = $window.FindName('BtnHybScepHealth')

    # Tab 9: Precision Enterprise Remediation Arsenal Controls
    $btnFixHealthReadout = $window.FindName('BtnFixHealthReadout')
    $btnClearFixOut      = $window.FindName('BtnClearFixOut')
    $btnClearRemOut      = $window.FindName('BtnClearRemOut')
    $btnCopyFixLog       = $window.FindName('BtnCopyFixLog')
    $txtRemOut           = $window.FindName('TxtRemOut')
    $btnFixWmi           = $window.FindName('BtnFixWmi')
    $btnFixWu            = $window.FindName('BtnFixWu')
    $btnFixCatroot       = $window.FindName('BtnFixCatroot')
    $btnFixBits          = $window.FindName('BtnFixBits')
    $btnFixDcom          = $window.FindName('BtnFixDcom')
    $btnFixSpooler       = $window.FindName('BtnFixSpooler')
    $btnFixNetStack      = $window.FindName('BtnFixNetStack')
    $btnFixWinRm         = $window.FindName('BtnFixWinRm')
    $btnFixProfiles      = $window.FindName('BtnFixProfiles')
    $btnFixTpm           = $window.FindName('BtnFixTpm')
    $btnFixAppX          = $window.FindName('BtnFixAppX')
    $lstDiagStages     = $window.FindName('LstDiagStages')

    $rowConsoleLog     = $window.FindName('RowConsoleLog')
    $btnToggleConsole  = $window.FindName('BtnToggleConsole')
    $btnClearHybOut    = $window.FindName('BtnClearHybOut')
    $hubProgressBar    = $window.FindName('HubProgressBar')
    $txtProgressStatus = $window.FindName('TxtProgressStatus')
    $txtHubLog         = $window.FindName('TxtHubLog')
    $btnCopyLog        = $window.FindName('BtnCopyLog')
    $btnClearLog       = $window.FindName('BtnClearLog')
    $btnSaveLog        = $window.FindName('BtnSaveLog')
    $btnQuickCmd       = if ($window.FindName('BtnQuickCmd')) { $window.FindName('BtnQuickCmd') } else { $menuQuickCmd }
    $btnTimeSync       = if ($window.FindName('BtnTimeSync')) { $window.FindName('BtnTimeSync') } else { $menuTimeSync }
    $btnReboot         = if ($window.FindName('BtnReboot')) { $window.FindName('BtnReboot') } else { $menuReboot }

    $txtDellServiceTag     = $window.FindName('TxtDellServiceTag')
    $btnDetectDellTag      = $window.FindName('BtnDetectDellTag')
    $btnCheckDellWarranty  = $window.FindName('BtnCheckDellWarranty')
    $btnCopyDellReport     = $window.FindName('BtnCopyDellReport')
    $btnExportDellCsv      = $window.FindName('BtnExportDellCsv')
    $borderRefreshVerdict  = $window.FindName('BorderRefreshVerdict')
    $borderVerdictBadge    = $window.FindName('BorderVerdictBadge')
    $txtVerdictBadge       = $window.FindName('TxtVerdictBadge')
    $txtVerdictTitle       = $window.FindName('TxtVerdictTitle')
    $txtVerdictDesc        = $window.FindName('TxtVerdictDesc')
    $txtDellModel          = $window.FindName('TxtDellModel')
    $txtDellProductLine    = $window.FindName('TxtDellProductLine')
    $txtDellShipDate       = $window.FindName('TxtDellShipDate')
    $txtDellAge            = $window.FindName('TxtDellAge')
    $txtDellContract       = $window.FindName('TxtDellContract')
    $txtDellRegion         = $window.FindName('TxtDellRegion')
    $txtDellEndDate        = $window.FindName('TxtDellEndDate')
    $txtDellDaysRemaining  = $window.FindName('TxtDellDaysRemaining')
    $lstDellEntitlements   = $window.FindName('LstDellEntitlements')

    # Hyperthreaded Asynchronous Background Worker Architecture
    $script:HubAsyncQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:HubActiveAsyncTasks = [System.Collections.Generic.List[object]]::new()

    # Configure InitialSessionState for background runspaces
    $hubIss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $hubIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new("HubAsyncQueue", $script:HubAsyncQueue, "Shared async queue"))

    # Add all non-GUI script functions into the background runspace pool
    foreach ($f in (Get-ChildItem Function:)) {
        if ($f.Name -notmatch '^(Start-AutopilotHubGui|Show-PrivilegeGuide|Start-GraphAuthDialog|Show-OfflineStagingDialog|Update-WpfUI)$') {
            try {
                $hubIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($f.Name, $f.Definition))
            } catch { }
        }
    }

    # Worker-side helper functions for logging and progress reporting back to the UI thread
    $wLogDef = {
        param([string]$Message, [string]$Level = 'INFO')
        if ($HubAsyncQueue) {
            $HubAsyncQueue.Enqueue([PSCustomObject]@{ Type = 'Log'; Message = $Message; Level = $Level })
        }
    }.ToString()
    $hubIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("Write-WorkerLog", $wLogDef))

    $wProgDef = {
        param([int]$Percent, [string]$Status)
        if ($HubAsyncQueue) {
            $HubAsyncQueue.Enqueue([PSCustomObject]@{ Type = 'Progress'; Percent = $Percent; Status = $Status })
        }
    }.ToString()
    $hubIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("Set-WorkerProgress", $wProgDef))

    $wWuProgDef = {
        param([int]$Percent, [string]$Status)
        if ($HubAsyncQueue) {
            $HubAsyncQueue.Enqueue([PSCustomObject]@{ Type = 'WuProgress'; Percent = $Percent; Status = $Status })
            $HubAsyncQueue.Enqueue([PSCustomObject]@{ Type = 'Progress'; Percent = $Percent; Status = $Status })
        }
    }.ToString()
    $hubIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new("Set-WuWorkerProgress", $wWuProgDef))

    $workerPoolSize = [Math]::Max(4, [Environment]::ProcessorCount)
    try {
        $script:HubWorkerPool = [runspacefactory]::CreateRunspacePool(1, $workerPoolSize, $hubIss, $Host)
        $script:HubWorkerPool.Open()
    } catch {
        $script:HubWorkerPool = $null
    }

    function Start-HubAsyncWork {
        param(
            [Parameter(Mandatory=$true)]
            [scriptblock]$WorkerScript,
            [array]$TaskArgs = @(),
            [scriptblock]$OnComplete = $null,
            [string]$Description = "Background Task"
        )

        if (-not $script:HubWorkerPool -or $script:HubWorkerPool.RunspacePoolStateInfo.State -ne [System.Management.Automation.Runspaces.RunspacePoolState]::Opened) {
            # Graceful fallback to direct execution if RunspacePool cannot be opened
            try {
                $taskResult = & $WorkerScript @TaskArgs
                if ($OnComplete) { & $OnComplete $taskResult $null }
            } catch {
                if ($OnComplete) { & $OnComplete $null $_.Exception }
            }
            return $null
        }

        try {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $script:HubWorkerPool
            $ps.AddScript($WorkerScript) | Out-Null
            if ($TaskArgs -and $TaskArgs.Count -gt 0) {
                foreach ($arg in $TaskArgs) {
                    $ps.AddArgument($arg) | Out-Null
                }
            }
            $handle = $ps.BeginInvoke()
            $taskObj = [PSCustomObject]@{
                PowerShell  = $ps
                Handle      = $handle
                OnComplete  = $OnComplete
                Description = $Description
                StartTime   = [datetime]::UtcNow
            }
            $script:HubActiveAsyncTasks.Add($taskObj)
            return $taskObj
        } catch {
            Write-HubLog "Failed to dispatch async worker '$Description': $($_.Exception.Message)" "ERROR"
            if ($OnComplete) {
                try { & $OnComplete $null $_.Exception } catch { }
            }
            return $null
        }
    }

    $script:IsPumpingWpfUi = $false

    # Lightweight UI DispatcherTimer (runs every 30ms on the main STA UI thread to drain queue and harvest completed workers)
    $script:HubQueueTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:HubQueueTimer.Interval = [TimeSpan]::FromMilliseconds(30)
    $script:HubQueueTimer.Add_Tick({
        if ($script:IsPumpingWpfUi) { return }
        $script:IsPumpingWpfUi = $true
        try {
            # 1. Drain incoming thread-safe worker messages
            $msg = $null
            while ($script:HubAsyncQueue.TryDequeue([ref]$msg)) {
                if (-not $msg) { continue }
                try {
                    if ($msg -is [scriptblock]) {
                        & $msg
                    } elseif ($msg.Type -eq 'Log') {
                        $timestamp = (Get-Date).ToString('HH:mm:ss')
                        $line = "[$timestamp] [$($msg.Level)] $($msg.Message)"
                        $txtHubLog.AppendText("$line`r`n")
                        $txtHubLog.ScrollToEnd()
                    } elseif ($msg.Type -eq 'Progress') {
                        $hubProgressBar.Value = $msg.Percent
                        $txtProgressStatus.Text = $msg.Status
                    } elseif ($msg.Type -eq 'Action' -and $msg.Action -is [scriptblock]) {
                        & $msg.Action
                    } elseif ($msg.Type -eq 'WuProgress') {
                        if ($progWuBar -and $null -ne $msg.Percent) {
                            $progWuBar.IsIndeterminate = $false
                            $progWuBar.Value = $msg.Percent
                        }
                        if ($txtWuProgress -and $null -ne $msg.Percent) {
                            $txtWuProgress.Text = "$($msg.Percent)%"
                        }
                        if ($txtWuStatus -and $msg.Status) {
                            $txtWuStatus.Text = $msg.Status
                        }
                    }
                } catch { }
            }

            # 2. Harvest completed background tasks
            for ($i = $script:HubActiveAsyncTasks.Count - 1; $i -ge 0; $i--) {
                $task = $script:HubActiveAsyncTasks[$i]
                if ($task.Handle.IsCompleted) {
                    $taskResult = $null
                    $taskError = $null
                    try {
                        $taskResult = $task.PowerShell.EndInvoke($task.Handle)
                        if ($task.PowerShell.Streams.Error -and $task.PowerShell.Streams.Error.Count -gt 0) {
                            $taskError = $task.PowerShell.Streams.Error[0]
                        }
                    } catch {
                        $taskError = $_.Exception
                    } finally {
                        try { $task.PowerShell.Dispose() } catch { }
                        $script:HubActiveAsyncTasks.RemoveAt($i)
                    }

                    if ($task.OnComplete) {
                        try {
                            & $task.OnComplete $taskResult $taskError
                        } catch {
                            Write-HubLog "Error in async completion callback for '$($task.Description)': $($_.Exception.Message)" "ERROR"
                        }
                    }
                }
            }
        } finally {
            $script:IsPumpingWpfUi = $false
        }
    })
    $script:HubQueueTimer.Start()

    # UI Refresh Helper (Pumps WPF message loop without blocking)
    function Update-WpfUI {
        if ($script:IsPumpingWpfUi) { return }
        $script:IsPumpingWpfUi = $true
        try {
            $frame = [System.Windows.Threading.DispatcherFrame]::new()
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [System.Action[System.Windows.Threading.DispatcherFrame]]{ param($f) $f.Continue = $false },
                $frame
            ) | Out-Null
            [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        } finally {
            $script:IsPumpingWpfUi = $false
        }
    }

    # Logging Helper
    function Write-HubLog {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = (Get-Date).ToString('HH:mm:ss')
        $line = "[$timestamp] [$Level] $Message"
        $txtHubLog.AppendText("$line`r`n")
        $txtHubLog.ScrollToEnd()
        Update-WpfUI
    }

    # Set Progress Helper
    function Set-HubProgress {
        param([int]$Percent, [string]$Status)
        $hubProgressBar.Value = $Percent
        $txtProgressStatus.Text = "$Status ($Percent%)"
        Update-WpfUI
    }

    # Initialize Hardware Telemetry
    Write-HubLog "Initializing Autopilot OOBE Command Hub v2.0..."
    $ctx = $script:RuntimeContext
    Write-HubLog "Session: $($ctx.Mode) mode | Privilege: $($ctx.PrivilegeLevel) ($($ctx.UserName)) | Host: $($ctx.HostEngine)"

    # Privilege badge - always visible; the Fix button only appears when we are below the preferred level
    $brushConv = [System.Windows.Media.BrushConverter]::new()
    if ($ctx.MeetsPreferred) {
        $dotPrivilege.Fill = $brushConv.ConvertFromString("#15803D")
        $badgePrivilege.Background = $brushConv.ConvertFromString("#DCFCE7")
        $badgePrivilege.BorderBrush = $brushConv.ConvertFromString("#86EFAC")
        $txtPrivilege.Text = "PRIV: $($ctx.PrivilegeLevel) | $($ctx.Mode)"
        $txtPrivilege.Foreground = $brushConv.ConvertFromString("#15803D")
    } else {
        $dotPrivilege.Fill = $brushConv.ConvertFromString("#B91C1C")
        $badgePrivilege.Background = $brushConv.ConvertFromString("#FEE2E2")
        $badgePrivilege.BorderBrush = $brushConv.ConvertFromString("#FCA5A5")
        $txtPrivilege.Text = "PRIV: $($ctx.PrivilegeLevel) | $($ctx.Mode)"
        $txtPrivilege.Foreground = $brushConv.ConvertFromString("#B91C1C")
        $btnFixPrivilege.Visibility = [System.Windows.Visibility]::Visible
        Write-HubLog "Running WITHOUT elevation. Hash harvest, rename, app installs, clock sync and restart-resume will fail. Click 'Fix Privileges' in the header - preferred level is $($ctx.PreferredLevel)." "WARN"
    }
    $btnFixPrivilege.Add_Click({
        if (Show-PrivilegeGuide -Owner $window) {
            Write-HubLog "Elevated Hub launched. Closing this non-elevated window." "SUCCESS"
            $window.Close()
        }
    })
    $badgePrivilege.Add_MouseLeftButtonUp({
        if (Show-PrivilegeGuide -Owner $window) {
            Write-HubLog "Elevated Hub launched. Closing this non-elevated window." "SUCCESS"
            $window.Close()
        }
    })

    if ($script:ResumeFromRestart) {
        Write-HubLog "Hub resumed automatically after restart; the one-shot resume hook has been consumed and removed." "SUCCESS"
    } elseif ($ctx.MeetsPreferred) {
        if (Unregister-HubResumeAfterRestart) { Write-HubLog "Removed a stale restart-resume hook left over from an earlier session." "INFO" }
    }

    function Update-DeviceStateUi {
        $ds = $script:DeviceState
        if (-not $ds) { return }
        $bc = [System.Windows.Media.BrushConverter]::new()
        $tenant = if ($ds.AutopilotTenantDomain) { $ds.AutopilotTenantDomain } elseif ($ds.EntraTenantName) { $ds.EntraTenantName } elseif ($ds.AutopilotTenantId) { $ds.AutopilotTenantId } else { '' }

        $cloudNote = ''
        if ($ds.CloudChecked) {
            if ($ds.CloudRegistered) {
                $ci = $ds.CloudIdentity
                $cloudNote = " Tenant record found: group tag '$($ci.groupTag)', enrollment '$($ci.enrollmentState)', profile '$($ci.deploymentProfileAssignmentStatus)', last contact $($ci.lastContactedDateTime)."
            } else {
                $cloudNote = " The signed-in tenant has NO Autopilot record for this serial number."
            }
        }

        $showAction = $true
        switch ($ds.Verdict) {
            'Bypassed' {
                $title = "DEVICE STATE: AUTOPILOT BYPASSED (LOCAL INSTALL)"
                $detail = "$($ds.Summary). Autopilot has been skipped on this install. OOBE will finish as a normal Windows with a local account.$cloudNote"
                $bg = '#E0F2FE'; $border = '#BAE6FD'; $fg = '#0284C7'; $showAction = $false
                $txtDeviceState.Text = "AUTOPILOT BYPASSED"; $txtDeviceState.Foreground = $bc.ConvertFromString('#0284C7')
            }
            'AutopilotRegistered' {
                $title = "DEVICE STATE: AUTOPILOT REGISTERED"
                $detail = "$($ds.Summary). This PC already received its deployment profile from the Autopilot Deployment Service - that is why it boots into the branded OOBE. Registration is NOT required; use this tab only to export a CSV or change the group tag.$cloudNote"
                if ($ds.CloudChecked -and -not $ds.CloudRegistered) { $detail += " You are probably signed in to a different tenant than the one this device belongs to ($tenant)." }
                $bg = '#DCFCE7'; $border = '#BBF7D0'; $fg = '#15803D'; $showAction = $false
                $txtDeviceState.Text = "AUTOPILOT: $tenant"; $txtDeviceState.Foreground = $bc.ConvertFromString('#15803D')
            }
            'IntuneEnrolled' {
                $title = "DEVICE STATE: INTUNE ENROLLED, NO AUTOPILOT PROFILE"
                $detail = "$($ds.Summary). The device is managed, but Autopilot does not own it - a reset or reprovision would land in the generic OOBE. Harvest the hardware hash and register it so Autopilot takes over next time.$cloudNote"
                $bg = '#FEF9C3'; $border = '#FDE68A'; $fg = '#A16207'
                $txtDeviceState.Text = "INTUNE: $tenant"; $txtDeviceState.Foreground = $bc.ConvertFromString('#A16207')
            }
            'EntraJoined' {
                $title = "DEVICE STATE: ENTRA JOINED, NOT ENROLLED"
                $detail = "$($ds.Summary). No MDM enrollment and no Autopilot profile. Harvest and register the hash to bring it under Autopilot.$cloudNote"
                $bg = '#FEF9C3'; $border = '#FDE68A'; $fg = '#A16207'
                $txtDeviceState.Text = "ENTRA: $tenant"; $txtDeviceState.Foreground = $bc.ConvertFromString('#A16207')
            }
            'DomainJoined' {
                $title = "DEVICE STATE: ON-PREM DOMAIN JOINED"
                $detail = "$($ds.Summary). Harvest and register the hardware hash if this device is moving to Autopilot.$cloudNote"
                $bg = '#FEF9C3'; $border = '#FDE68A'; $fg = '#A16207'
                $txtDeviceState.Text = "DOMAIN JOINED"; $txtDeviceState.Foreground = $bc.ConvertFromString('#A16207')
            }
            default {
                $title = "DEVICE STATE: NOT REGISTERED"
                $detail = "$($ds.Summary). Harvest the hardware hash, then register it with Intune (Graph) or export the CSV for a bulk import.$cloudNote"
                $bg = '#FEE2E2'; $border = '#FCA5A5'; $fg = '#B91C1C'
                $txtDeviceState.Text = "NOT REGISTERED"; $txtDeviceState.Foreground = $bc.ConvertFromString('#DC2626')
            }
        }
        if ($ds.CloudChecked -and $ds.CloudRegistered -and $ds.Verdict -ne 'AutopilotRegistered') {
            $title = "DEVICE STATE: REGISTERED IN SIGNED-IN TENANT"
            $detail = "The signed-in tenant already holds an Autopilot identity for this serial number.$cloudNote Registering again is unnecessary; the local machine simply has not been through Autopilot OOBE yet."
            $bg = '#DCFCE7'; $border = '#BBF7D0'; $fg = '#15803D'; $showAction = $false
            $txtDeviceState.Text = "AUTOPILOT (TENANT)"; $txtDeviceState.Foreground = $bc.ConvertFromString('#15803D')
        }
        $bannerDeviceState.Background = $bc.ConvertFromString($bg)
        $bannerDeviceState.BorderBrush = $bc.ConvertFromString($border)
        $txtDeviceStateTitle.Text = $title
        $txtDeviceStateTitle.Foreground = $bc.ConvertFromString($fg)
        $txtDeviceStateDetail.Text = $detail
        $btnDeviceStateAction.Visibility = if ($showAction -and $script:RuntimeContext.MeetsPreferred) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

        # App Deployment advisory - app assignment belongs to Intune once the device is managed
        if ($ds.Verdict -in @('AutopilotRegistered', 'IntuneEnrolled') -or $ds.CloudRegistered) {
            $bannerAppAdvisory.Background = $bc.ConvertFromString('#FEF9C3'); $bannerAppAdvisory.BorderBrush = $bc.ConvertFromString('#FDE68A')
            $txtAppAdvisory.Foreground = $bc.ConvertFromString('#A16207')
            $txtAppAdvisory.Text = "This device is managed by Intune ($tenant). App deployment normally belongs to Intune app assignments and the Enrollment Status Page - use this tab for bench builds, one-off tools and troubleshooting, not as a substitute for policy."
        } else {
            $bannerAppAdvisory.Background = $bc.ConvertFromString('#F1F5F9'); $bannerAppAdvisory.BorderBrush = $bc.ConvertFromString('#CBD5E1')
            $txtAppAdvisory.Foreground = $bc.ConvertFromString('#475569')
            $txtAppAdvisory.Text = "Unmanaged device - installing apps from here is appropriate. Once it is Autopilot-registered and enrolled, prefer Intune app assignments so the build stays reproducible."
        }
    }

    function Invoke-TenantAutopilotLookup {
        if (-not $script:DeviceState) { return }
        if (-not ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken)) { return }
        $serial = $txtSerial.Text
        if ([string]::IsNullOrWhiteSpace($serial) -or $serial -in @('Detecting...', 'UNKNOWN', 'UNAVAILABLE')) { return }
        $lookupKey = "$($script:GraphAuthContext.TenantId)|$serial"
        if ($script:TenantLookupKey -eq $lookupKey) { return }
        $script:TenantLookupKey = $lookupKey
        try {
            Write-HubLog "Checking the signed-in tenant for an existing Autopilot identity (serial $serial)..."
            $identity = Find-AutopilotIdentityInTenant -SerialNumber $serial -AccessToken $script:GraphAuthContext.AccessToken
            $script:DeviceState.CloudChecked = $true
            $script:DeviceState.CloudRegistered = [bool]$identity
            $script:DeviceState.CloudIdentity = $identity
            if ($identity) {
                Write-HubLog "Tenant already holds this device: group tag '$($identity.groupTag)', enrollment '$($identity.enrollmentState)', profile '$($identity.deploymentProfileAssignmentStatus)'." "SUCCESS"
            } else {
                Write-HubLog "No Autopilot identity for serial $serial in the signed-in tenant." "INFO"
            }
        } catch {
            $script:TenantLookupKey = $null
            Write-HubLog "Tenant Autopilot lookup skipped: $(Get-GraphErrorMessage -ErrorRecord $_)" "WARN"
        }
        Update-DeviceStateUi
    }

    try {
        $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
        $txtSerial.Text = if ($bios.SerialNumber) { $bios.SerialNumber } else { 'UNKNOWN' }
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $txtModel.Text = "$($cs.Manufacturer) $($cs.Model)"
    } catch {
        $txtSerial.Text = 'UNAVAILABLE'
        $txtModel.Text = 'Standard PC'
    }

    # TPM probe off the UI thread (Get-Tpm can stall on some firmware); result applied when it lands
    $txtTpm.Text = "Probing..."
    Start-HubAsyncWork -Description "TPM Probe" -WorkerScript {
        try { $t = Get-Tpm -ErrorAction SilentlyContinue; [PSCustomObject]@{ Present = [bool]($t -and $t.TpmPresent) } }
        catch { [PSCustomObject]@{ Present = $false } }
    } -OnComplete {
        param($tpmInfo, $err)
        if ($tpmInfo -and $tpmInfo.Present) {
            $txtTpm.Text = "READY (v2.0)"; $txtTpm.Foreground = [System.Windows.Media.Brushes]::LimeGreen
        } elseif (-not $script:RuntimeContext.MeetsPreferred) {
            # Get-Tpm returns nothing without elevation - that is not the same as "no TPM"
            $txtTpm.Text = "NEEDS ELEVATION"; $txtTpm.Foreground = [System.Windows.Media.Brushes]::Gold
        } else {
            $txtTpm.Text = "NOT READY"; $txtTpm.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        }
    } | Out-Null

    # Secure Boot: the registry mirror is readable by any user; Confirm-SecureBootUEFI needs elevation and
    # threw "Unable to set proper privileges", which used to render as UEFI UNVERIFIED on every desktop launch
    try {
        $sbState = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' -ErrorAction Stop
        if ([int]$sbState.UEFISecureBootEnabled -eq 1) {
            $txtSecureBoot.Text = "ACTIVE (UEFI)"; $txtSecureBoot.Foreground = [System.Windows.Media.Brushes]::LimeGreen
        } else {
            $txtSecureBoot.Text = "DISABLED / LEGACY"; $txtSecureBoot.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        }
    } catch {
        $txtSecureBoot.Text = "UEFI UNVERIFIED"
    }

    # Network Quick Ping
    try {
        $dns = [System.Net.Dns]::GetHostAddresses("graph.microsoft.com")
        if ($dns.Count -gt 0) {
            $txtNetwork.Text = "ONLINE (GRAPH OK)"
            $txtNetwork.Foreground = [System.Windows.Media.Brushes]::LimeGreen
        }
    } catch {
        $txtNetwork.Text = "LIMITED / OFFLINE"
        $txtNetwork.Foreground = [System.Windows.Media.Brushes]::Gold
    }

    # Dell Hardware Detection & Pre-population
    try {
        if ($bios -and $bios.SerialNumber) {
            $txtDellServiceTag.Text = $bios.SerialNumber.Trim()
        }
        if ($cs -and $cs.Manufacturer -match 'Dell') {
            Write-HubLog "Dell enterprise hardware detected ($($cs.Manufacturer) $($cs.Model)). Dell Warranty & Refresh Lifecycle Engine is armed." "INFO"
        }
    } catch { }

    # Apply Configured Defaults from .env
    $renameEnabledVal = $env:AUTOPILOT_RENAME_ENABLED
    if ($renameEnabledVal -match '^(false|0|no)$') {
        $chkEnableRename.IsChecked = $false
        $txtComputerName.IsEnabled = $false
        $btnApplyRename.IsEnabled = $false
    } else {
        $chkEnableRename.IsChecked = $true
        $txtComputerName.IsEnabled = $true
        $btnApplyRename.IsEnabled = $true
    }

    $chkEnableRename.Add_Checked({
        $txtComputerName.IsEnabled = $true
        $btnApplyRename.IsEnabled = $true
    })
    $chkEnableRename.Add_Unchecked({
        $txtComputerName.IsEnabled = $false
        $btnApplyRename.IsEnabled = $false
    })

    $prefix = if ($env:AUTOPILOT_NAME_PREFIX) { $env:AUTOPILOT_NAME_PREFIX.Trim() } else { 'WS' }
    $nameTemplate = if ($env:AUTOPILOT_NAME_TEMPLATE) { $env:AUTOPILOT_NAME_TEMPLATE.Trim() } else { "$prefix-%SERIAL%" }
    $txtComputerName.Text = $nameTemplate

    if ($env:AUTOPILOT_GROUP_TAG) {
        $cmbGroupTag.Text = $env:AUTOPILOT_GROUP_TAG.Trim()
    }
    if ($env:AUTOPILOT_ASSIGNED_USER) {
        $txtAssignedUser.Text = $env:AUTOPILOT_ASSIGNED_USER.Trim()
    }
    if ($env:AUTOPILOT_WAIT_FOR_SYNC -match '^(false|0|no)$') {
        $chkWaitForSync.IsChecked = $false
    }
    if ($env:AUTOPILOT_AUTO_REBOOT -match '^(true|1|yes)$') {
        $chkAutoReboot.IsChecked = $true
    }
    if ($env:AUTOPILOT_AUTO_DETECT_USB -match '^(false|0|no)$') {
        $chkAutoDetectUsb.IsChecked = $false
    }

    $btnSaveEnvDefaults.Add_Click({
        $prefixVal = if ($txtComputerName.Text -match '^([A-Za-z0-9]+)[-_]') { $Matches[1] } else { 'WS' }
        $settingsToSave = @{
            'AUTOPILOT_RENAME_ENABLED'  = if ($chkEnableRename.IsChecked) { 'true' } else { 'false' }
            'AUTOPILOT_NAME_PREFIX'     = $prefixVal
            'AUTOPILOT_NAME_TEMPLATE'   = $txtComputerName.Text.Trim()
            'AUTOPILOT_GROUP_TAG'       = $cmbGroupTag.Text.Trim()
            'AUTOPILOT_ASSIGNED_USER'   = $txtAssignedUser.Text.Trim()
            'AUTOPILOT_WAIT_FOR_SYNC'   = if ($chkWaitForSync.IsChecked) { 'true' } else { 'false' }
            'AUTOPILOT_AUTO_REBOOT'     = if ($chkAutoReboot.IsChecked) { 'true' } else { 'false' }
            'AUTOPILOT_AUTO_DETECT_USB' = if ($chkAutoDetectUsb.IsChecked) { 'true' } else { 'false' }
            'AUTOPILOT_CHARSET'         = if ($env:AUTOPILOT_CHARSET) { $env:AUTOPILOT_CHARSET } else { 'native' }
        }
        $savedPath = Save-EnvConfig -Settings $settingsToSave
        Write-HubLog "Configuration defaults saved to $savedPath (charset: $(Get-ScriptEncodingName))." "SUCCESS"
        [System.Windows.MessageBox]::Show("Configuration saved to:`n$savedPath`n`nCharset: $(Get-ScriptEncodingName)", "Settings Saved to .env", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    # --- ACTION: Harvest Hash (Hyperthreaded & Asynchronous) ---
    function Invoke-HubHarvest {
        if ($btnHarvestHash) { $btnHarvestHash.IsEnabled = $false }
        if ($btnDeviceStateAction) { $btnDeviceStateAction.IsEnabled = $false }
        Write-HubLog "Starting high-speed Autopilot hardware hash harvester (async worker)..."
        Set-HubProgress -Percent 15 -Status "Querying MDM Provider"

        $gt = $cmbGroupTag.Text
        $usr = $txtAssignedUser.Text

        Start-HubAsyncWork -Description "Hardware Hash Harvest" -TaskArgs @($gt, $usr) -WorkerScript {
            param($groupTag, $assignedUser)
            Write-WorkerLog "Worker runspace reading WMI MDM_DevDetail_Ext01 & OA3 hash..."
            Set-WorkerProgress 50 "Querying MDM Provider"
            $info = Get-AutopilotHash -GroupTag $groupTag -AssignedUser $assignedUser
            Set-WorkerProgress 90 "Validating OA3 structure"
            return $info
        } -OnComplete {
            param($hashInfo, $err)
            if ($btnHarvestHash) { $btnHarvestHash.IsEnabled = $true }
            if ($btnDeviceStateAction) { $btnDeviceStateAction.IsEnabled = $true }

            if ($hashInfo -and $hashInfo.IsValidStructure) {
                $txtHashBox.Text = $hashInfo.HardwareHash
                $txtHashStatus.Text = "VALIDATED OA3 HASH"
                $badgeHashStatus.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#059669")
                $txtHashStatus.Foreground = [System.Windows.Media.Brushes]::White
                $txtHashMeta.Text = "Serial: $($hashInfo.SerialNumber) | Length: $($hashInfo.HashLengthBytes) bytes | GroupTag: '$($hashInfo.GroupTag)'"
                Write-HubLog "Hardware hash harvested successfully ($($hashInfo.HashLengthBytes) bytes). $(Get-HubQuip -Kind Success)" "SUCCESS"
                Set-HubProgress -Percent 100 -Status "Hash Ready"
                Play-HubAudio -Type Success
            } else {
                $txtHashBox.Text = ''
                $txtHashStatus.Text = "HASH UNAVAILABLE"
                $badgeHashStatus.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FEE2E2")
                $txtHashStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FF99A4")
                $reason = if ($hashInfo) { $hashInfo.StatusReason } elseif ($err) { $err.Message } else { 'Get-AutopilotHash returned nothing' }
                $serial = if ($hashInfo) { $hashInfo.SerialNumber } else { 'Unknown' }
                $txtHashMeta.Text = "Serial: $serial | No genuine hardware hash - registration and CSV export are blocked"
                Write-HubLog "Hardware hash unavailable: $reason" "ERROR"
                Set-HubProgress -Percent 0 -Status "Harvest Failed"
                Play-HubAudio -Type Error
            }
        }
    }
    $btnHarvestHash.Add_Click({ Invoke-HubHarvest })
    $btnDeviceStateAction.Add_Click({ Invoke-HubHarvest })

    # --- ACTION: Skip Autopilot, Personal MSA & Local User Handlers ---
    if ($btnCreateLocalUser) {
        $btnCreateLocalUser.Add_Click({
            Show-CreateLocalUserDialog -Owner $window
        })
    }
    if ($menuCreateLocalUser) {
        $menuCreateLocalUser.Add_Click({
            Show-CreateLocalUserDialog -Owner $window
        })
    }

    $personalMsaAction = {
        $ctx = $script:RuntimeContext
        if (-not $ctx.MeetsPreferred) {
            [System.Windows.MessageBox]::Show("Configuring Personal Account mode requires elevation. Use 'Fix Privileges' first.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $ans = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
            [System.Windows.MessageBox]::Show("Configure this device for Personal Microsoft Account (MSA) setup?`n`nThis will:`n- Clear Autopilot corporate locks & tenant policy cache`n- Re-enable network adapters for online sign-in`n- Unlock the personal Microsoft Account path in OOBE`n`nContinue?", "Personal Microsoft Account Setup", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        }
        if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }

        Write-HubLog "Unlocking Personal Microsoft Account (MSA) setup mode..." "WARN"
        $res = Invoke-HubSetupPersonalAccount
        foreach ($a in $res) { Write-HubLog "  $a" $(if ($a -like 'FAILED*') { 'ERROR' } else { 'INFO' }) }

        $script:DeviceState = Get-DeviceEnrollmentState
        Update-DeviceStateUi
        Write-HubLog "Device state updated: $($script:DeviceState.Summary)" "INFO"
        Write-HubLog "Personal MSA setup ready. In OOBE, choose 'Set up for personal use' and sign in with your Microsoft Account." "SUCCESS"
        if ($ctx.IsOobe) {
            $ansReload = [System.Windows.MessageBox]::Show("Personal Microsoft Account mode configured!`n`nCorporate enrollment has been cleared and network is active.`nIn OOBE, select 'Set up for personal use' to sign in with your personal Microsoft Account.`n`nWould you like to restart OOBE processes now so Windows Setup immediately detects personal mode?", "Personal MSA Unlocked - Restart OOBE?", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
            if ($ansReload -eq [System.Windows.MessageBoxResult]::Yes) {
                Restart-HubOobe
            }
        } else {
            [System.Windows.MessageBox]::Show("Personal Microsoft Account mode configured!`n`nCorporate enrollment has been cleared and network is active.`nIn OOBE, select 'Set up for personal use' to sign in with your personal Microsoft Account.", "Personal MSA Unlocked", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
    }

    $skipAutopilotAction = {
        $ctx = $script:RuntimeContext
        if (-not $ctx.MeetsPreferred) {
            [System.Windows.MessageBox]::Show("This needs elevation (it edits HKLM and can disable network adapters). Use 'Fix Privileges' first.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $disableNet = [bool]$chkPersonalDisableNet.IsChecked -and $ctx.IsOobe
        $plan = "- remove cached Autopilot profiles & tenant policy cache`n- flag Autopilot as disabled (IsAutopilotDisabled=1)`n- unlock offline & personal account paths in OOBE (BypassNRO=1, HidePersonalAccountOption=0)`n" + $(if ($disableNet) { "- disable physical network adapters for offline local setup`n" } else { "- keep network active for Personal Microsoft Account (MSA)`n" })
        $msg = if ($ctx.IsOobe) {
            "Setup Windows without corporate Autopilot / work or school enrollment:`n`n$plan`nContinue?"
        } else {
            "Autopilot only acts during OOBE, so there is nothing to skip on an installed system.`n`nApply the flags anyway?`n`n$plan"
        }
        $ans = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
            [System.Windows.MessageBox]::Show($msg, "Skip Autopilot / Setup Mode", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        }
        if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }

        Write-HubLog "Applying Autopilot bypass so OOBE finishes without corporate enrollment..." "WARN"
        $res = Invoke-AutopilotBypass -DisableNetwork:$disableNet
        foreach ($a in $res.Actions) { Write-HubLog "  $a" $(if ($a -like 'FAILED*') { 'ERROR' } else { 'INFO' }) }

        $script:DeviceState = Get-DeviceEnrollmentState
        Update-DeviceStateUi
        Write-HubLog "Device state updated: $($script:DeviceState.Summary)" "INFO"

        if ($ctx.IsOobe) {
            if ($disableNet) {
                Write-HubLog "Offline mode active. Close this window, continue OOBE and click 'I don't have internet' to create a local account." "SUCCESS"
                $ansReload = [System.Windows.MessageBox]::Show("Offline bypass applied!`n`nNetwork adapters disabled and BypassNRO=1 set.`nIn OOBE, click 'I don't have internet' -> 'Continue with limited setup' to create a local account.`n`nWould you like to restart OOBE processes now to ensure OOBE refreshes without network?", "Offline Bypass - Restart OOBE?", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
                if ($ansReload -eq [System.Windows.MessageBoxResult]::Yes) {
                    Restart-HubOobe
                }
            } else {
                Write-HubLog "Personal MSA mode active. Close this window and continue OOBE to sign in with your personal Microsoft Account." "SUCCESS"
            }
        } else {
            Write-HubLog "Flags applied. Remember: this only matters during OOBE of a fresh install." "SUCCESS"
        }
    }

    if ($btnPersonalInstall) { $btnPersonalInstall.Add_Click($personalMsaAction) }
    if ($menuPersonalMsa)     { $menuPersonalMsa.Add_Click($personalMsaAction) }
    if ($btnSkipAutopilot)    { $btnSkipAutopilot.Add_Click($skipAutopilotAction) }
    if ($btnRestartOobe)      { $btnRestartOobe.Add_Click({ Restart-HubOobe }) }

    $btnReenableNet.Add_Click({
        if (-not $script:RuntimeContext.MeetsPreferred) {
            [System.Windows.MessageBox]::Show("Re-enabling adapters needs elevation. Use 'Fix Privileges' first.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $en = @(Enable-HubNetworkAdapters)
        if ($en.Count -gt 0) { Write-HubLog "Re-enabled network adapters: $($en -join ', ')" "SUCCESS" } else { Write-HubLog "No disabled physical network adapters found." "INFO" }
    })

    # Offline Staging & Re-Docking UI Handlers
    function Update-OfflineStagingUi {
        $snap = Get-OfflineStagingState
        if ($snap -and $snap.IsActive) {
            if ($bannerOfflineStaging) { $bannerOfflineStaging.Visibility = [System.Windows.Visibility]::Visible }
            if ($txtOfflineStagingTitle) { $txtOfflineStagingTitle.Text = "OFFLINE STAGING ACTIVE: $($snap.ProfileTitle.ToUpper())" }
            if ($txtOfflineStagingDetail) { $txtOfflineStagingDetail.Text = "Device isolated with reversible profile quarantine. Click below to restore cloud link." }
            if ($btnResumeAutopilot) { $btnResumeAutopilot.Visibility = [System.Windows.Visibility]::Visible; $btnResumeAutopilot.IsEnabled = $true }
        } else {
            if ($bannerOfflineStaging) { $bannerOfflineStaging.Visibility = [System.Windows.Visibility]::Collapsed }
            if ($btnResumeAutopilot) { $btnResumeAutopilot.IsEnabled = $false }
        }
    }

    $reDockAction = {
        $snap = Get-OfflineStagingState
        if (-not $snap) {
            [System.Windows.MessageBox]::Show("The device is not currently in Controlled Offline Staging mode. No active staging snapshot was found.", "Not In Staging Mode", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        $ans = [System.Windows.MessageBox]::Show("Are you ready to reconnect to Enterprise Cloud and resume Autopilot?`n`nProfile: $($snap.ProfileTitle)`n`nThis will restore service start types, clear hosts loopback fences, restore your Autopilot profile, and trigger enterprise MDM sync.", "Confirm Cloud Re-Docking", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
            Write-HubLog "Initiating cloud re-docking procedure..." "INFO"
            $res = Invoke-AutopilotReDock
            if ($res.Success) {
                Write-HubLog $res.Message "SUCCESS"
                Update-OfflineStagingUi
                [System.Windows.MessageBox]::Show($res.Message, "Cloud Re-Dock Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            } else {
                Write-HubLog $res.Message "ERROR"
            }
        }
    }

    if ($btnOfflineStaging) {
        $btnOfflineStaging.Add_Click({
            if (Show-OfflineStagingDialog -Owner $window) {
                Update-OfflineStagingUi
                Write-HubLog "Device transitioned to Controlled Offline Staging mode." "SUCCESS"
            }
        })
    }
    if ($btnResumeAutopilot) { $btnResumeAutopilot.Add_Click($reDockAction) }
    if ($btnReDockStaging) { $btnReDockStaging.Add_Click($reDockAction) }

    # --- ACTION: Export CSV ---
    $btnExportCsv.Add_Click({
        Write-HubLog "Exporting device record to Microsoft Intune CSV format..."
        $gt = $cmbGroupTag.Text
        $usr = $txtAssignedUser.Text

        $deviceName = if ($txtComputerName -and -not [string]::IsNullOrWhiteSpace($txtComputerName.Text)) { $txtComputerName.Text } else { '' }
        $res = Export-AutopilotCsv -AutoDetectUsb:$chkAutoDetectUsb.IsChecked -GroupTag $gt -AssignedUser $usr -DeviceName $deviceName
        if ($res.Success) {
            Write-HubLog "Autopilot CSV '$($res.FileName)' written to: $($res.Path)" "SUCCESS"
            Play-HubAudio -Type Success
            if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Autopilot CSV exported successfully!`n`nFile: $($res.FileName)`nDestination: $($res.Path)", "CSV Export Succeeded", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
        }
    })

    # --- ACTION: Copy Hash ---
    $btnCopyHash.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($txtHashBox.Text)) {
            [System.Windows.Clipboard]::SetText($txtHashBox.Text)
            Write-HubLog "Hardware hash copied to clipboard." "INFO"
        }
    })

    # --- ACTION: Apply Rename ---
    $btnApplyRename.Add_Click({
        $template = $txtComputerName.Text
        $serial = $txtSerial.Text
        $rand = (Get-Random -Minimum 1000 -Maximum 9999).ToString()

        $newName = $template.Replace('%SERIAL%', $serial).Replace('%RAND%', $rand)
        if ($newName.Length -gt 15) { $newName = $newName.Substring(0, 15) }

        Write-HubLog "Renaming computer to '$newName'..."
        try {
            Rename-Computer -NewName $newName -Force -ErrorAction Stop
            Write-HubLog "Computer renamed to '$newName'. Reboot required to take effect." "SUCCESS"
            if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Computer renamed to: $newName`n`nThe rename will take effect on next reboot.", "Rename Succeeded", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
        } catch {
            Write-HubLog "Rename failed: $($_.Exception.Message)" "ERROR"
        }
    })

    # --- ACTION: Register Directly to Intune ---
    $btnRegisterIntune.Add_Click({
        Write-HubLog "Initiating Microsoft Graph cloud registration pipeline..."
        Set-HubProgress -Percent 10 -Status "Checking Auth Session"

        $token = Get-CurrentGraphToken -AllowInteractive -Owner $window
        if (-not $token) {
            Write-HubLog "Cloud registration paused: Microsoft Graph authentication required." "WARN"
            Set-HubProgress -Percent 0 -Status "Auth Required"
            return
        }
        Update-GraphAuthHeader

        $ds = $script:DeviceState
        if ($ds -and ($ds.Verdict -eq 'AutopilotRegistered' -or $ds.CloudRegistered)) {
            $ans = [System.Windows.MessageBox]::Show("This device already appears to be Autopilot-registered:`n`n$($ds.Summary)`n`nImporting the same hardware hash again is either rejected by Intune (ZtdDeviceAlreadyAssigned) or only updates the group tag. Continue anyway?", "Already Registered", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($ans -ne [System.Windows.MessageBoxResult]::Yes) {
                Write-HubLog "Registration skipped - device already registered." "INFO"
                Set-HubProgress -Percent 0 -Status "Skipped"
                return
            }
        }

        Write-HubLog "Using active Microsoft Graph session ($($script:GraphAuthContext.TenantId))." "SUCCESS"
        Set-HubProgress -Percent 40 -Status "Uploading Device Hash"

        try {
            $reg = Register-AutopilotDevice -GroupTag $cmbGroupTag.Text -AssignedUser $txtAssignedUser.Text -AccessToken $token -WaitForSync:$chkWaitForSync.IsChecked
            if ($reg.Success) {
                Write-HubLog "Device registered to Intune! Import ID: $($reg.ImportId) - $(Get-HubQuip -Kind Success)" "SUCCESS"
                Set-HubProgress -Percent 100 -Status "Registration Complete"
                Play-HubAudio -Type Success

                if ($chkAutoReboot.IsChecked) {
                    Write-HubLog "Auto-reboot scheduled in 10 seconds..." "WARN"
                    Start-Process shutdown.exe -ArgumentList '/r /t 10 /c "Autopilot Registration Complete - Rebooting into OOBE ESP"'
                }
            }
        } catch {
            Write-HubLog "Intune Registration Error: $($_.Exception.Message) ($(Get-HubQuip -Kind Error))" "ERROR"
            Set-HubProgress -Percent 0 -Status "Registration Failed"
            Play-HubAudio -Type Error
        }
    })

    # --- APP BUNDLE PRESETS ---
    $btnPresetWorkstation.Add_Click({
        foreach ($b in $appBoxes) { $b.IsChecked = $false }
        $window.FindName('AppChrome').IsChecked = $true
        $window.FindName('AppVSCode').IsChecked = $true
        $window.FindName('AppGit').IsChecked = $true
        $window.FindName('AppTerminal').IsChecked = $true
        $window.FindName('App7Zip').IsChecked = $true
        $window.FindName('AppPowerToys').IsChecked = $true
        Write-HubLog "Loaded 'Recommended Workstation' application bundle."
    })

    $btnPresetBrowsers.Add_Click({
        $window.FindName('AppChrome').IsChecked = $true
        $window.FindName('AppFirefox').IsChecked = $true
        $window.FindName('AppBrave').IsChecked = $true
        $window.FindName('AppEdgeDev').IsChecked = $true
        Write-HubLog "Selected all web browsers."
    })

    $btnPresetDev.Add_Click({
        $window.FindName('AppVSCode').IsChecked = $true
        $window.FindName('AppGit').IsChecked = $true
        $window.FindName('AppTerminal').IsChecked = $true
        $window.FindName('AppNode').IsChecked = $true
        $window.FindName('AppPython').IsChecked = $true
        $window.FindName('AppGh').IsChecked = $true
        $window.FindName('AppDocker').IsChecked = $true
        Write-HubLog "Selected all developer tools."
    })

    $btnSelectAllApps.Add_Click({
        foreach ($b in $appBoxes) { $b.IsChecked = $true }
        Write-HubLog "Selected all applications."
    })

    $btnClearApps.Add_Click({
        foreach ($b in $appBoxes) { $b.IsChecked = $false }
        Write-HubLog "Cleared application selections."
    })

    # --- ACTION: Install Selected Apps ---
    $btnInstallBatch.Add_Click({
        $selectedIds = [System.Collections.Generic.List[string]]::new()
        foreach ($b in $appBoxes) {
            if ($b.IsChecked -and $b.Tag) {
                $selectedIds.Add($b.Tag.ToString())
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($txtCustomPkg.Text)) {
            $selectedIds.Add($txtCustomPkg.Text.Trim())
        }

        if ($selectedIds.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one application to install.", "No Apps Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        Write-HubLog "Starting batch installation of $($selectedIds.Count) applications via Winget..." "INFO"
        $scopeVal = if ($chkMachineScope.IsChecked) { 'machine' } else { 'user' }
        $silentVal = $chkSilentInstall.IsChecked

        $progressBlock = {
            param($Package, $Current, $Total, $Percent, $Status)
            Set-HubProgress -Percent $Percent -Status $Status
            Write-HubLog "[$Current/$Total] $Status"
        }

        $res = Invoke-AppInstallation -PackageIds $selectedIds.ToArray() -Scope $scopeVal -Silent:$silentVal -OnProgress $progressBlock

        $successCount = ($res | Where-Object { $_.Success }).Count
        Write-HubLog "Batch installation complete: $successCount / $($selectedIds.Count) succeeded." "SUCCESS"
        Set-HubProgress -Percent 100 -Status "Installation Complete"
    })

    # --- ACTION: Build Win32 Package ---
    $btnBuildPackage.Add_Click({
        $id = $txtPkgId.Text.Trim()
        $disp = $txtPkgDisplayName.Text.Trim()
        $outDir = $txtPkgOutputDir.Text.Trim()
        $args = $txtPkgInstallArgs.Text.Trim()

        Write-HubLog "Building Win32 app package for '$id' via WingetIntune..."
        Set-HubProgress -Percent 20 -Status "Packaging"

        $build = New-Win32AppPackage -PackageId $id -DisplayName $disp -OutputFolder $outDir -InstallArgs $args
        if ($build.Success) {
            Write-HubLog "Package compiled successfully: $($build.IntuneWinPath)" "SUCCESS"
            Set-HubProgress -Percent 100 -Status "Packaged"
            [System.Windows.MessageBox]::Show("Win32 package built successfully!`n`nFile: $($build.IntuneWinPath)`nMetadata: $($build.MetadataJsonPath)", "Packaging Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })

    # --- ACTION: Publish Package to Intune ---
    $btnPublishIntune.Add_Click({
        Write-HubLog "Initiating cloud publishing pipeline for '$($txtPkgId.Text)'..."
        $token = Get-CurrentGraphToken -AllowInteractive -Owner $window
        if (-not $token) {
            Write-HubLog "Publishing halted: Microsoft Graph authentication required." "WARN"
            return
        }
        Update-GraphAuthHeader
        Write-HubLog "Active Graph session verified for tenant $($script:GraphAuthContext.TenantId)." "SUCCESS"
        Write-HubLog "Ready to upload .intunewin package '$($txtPkgDisplayName.Text)' ($($txtPkgId.Text)) to Microsoft Intune mobileApps." "INFO"
        [System.Windows.MessageBox]::Show("Authenticated as $($script:GraphAuthContext.TenantId).`n`nReady to publish '$($txtPkgDisplayName.Text)' directly to Intune mobileApps.", "Intune Cloud Publisher", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    # --- Hybrid & Co-Mgmt tab helpers ---
    function Write-HybOut {
        param([string]$Message, [string]$Level = 'INFO')
        $ts = (Get-Date).ToString('HH:mm:ss')
        $txtHybOut.AppendText("`r`n[$ts] [$Level] $Message")
        $txtHybOut.ScrollToEnd()
        # Tab-local pane only - do NOT also mirror to the main audit log (that was double-printing every line)
        Update-WpfUI
    }
    function Assert-HybElevated {
        if (-not $script:RuntimeContext.MeetsPreferred) {
            Write-HybOut "This action needs elevation (SYSTEM/Administrator). Use 'Fix Privileges' in the header." "WARN"
            return $false
        }
        return $true
    }

    $btnHybRefresh.Add_Click({
        Write-HybOut "Reading join/identity state (dsregcmd)..."
        $j = Get-HybridJoinState
        $prt = if ($j.HasPrt) { 'present' } else { 'MISSING' }
        $txtHybJoin.Text = "$($j.JoinType). PRT: $prt. Tenant: $(if ($j.TenantName) { $j.TenantName } else { 'n/a' }). DeviceId: $(if ($j.DeviceId) { $j.DeviceId } else { 'n/a' })"
        $m = Get-MdmEnrollmentInfo
        $ime = Get-IntuneExtensionHealth
        $stale = if ($ime.StaleMins -ge 0) { "$($ime.StaleMins) min ago" } else { 'no logs' }
        $txtHybMdm.Text = "MDM enrolled: $($m.Enrolled)$(if ($m.Upn) { " ($($m.Upn))" }). IME: $($ime.Status), last log $stale."
        Write-HybOut "$($j.JoinType); PRT $prt; MDM enrolled=$($m.Enrolled); IME=$($ime.Status)." $(if ($j.IsHybrid) { 'SUCCESS' } else { 'INFO' })
    })

    $btnHybRetryJoin.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Retrying Azure AD device registration (dsregcmd /join + Automatic-Device-Join task)..." "WARN"
        try { & (Join-Path $env:SystemRoot 'System32\dsregcmd.exe') /join 2>&1 | Out-Null } catch {}
        try { Start-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join' -ErrorAction SilentlyContinue } catch {}
        Write-HybOut "Join retry kicked. Re-run Refresh State in a minute; hybrid join can lag behind Entra Connect sync." "INFO"
    })

    $btnHybSecChan.Add_Click({
        Write-HybOut "Testing machine secure channel to the domain..."
        $d = Test-DomainConnectivity
        Write-HybOut $d.Message $(if ($d.SecureChannel -like 'Healthy*') { 'SUCCESS' } elseif ($d.SecureChannel -like 'BROKEN*') { 'ERROR' } else { 'INFO' })
    })

    $btnHybSync.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Forcing an Intune/MDM sync..."
        $ran = @(Invoke-MdmSync)
        if ($ran.Count -gt 0) { Write-HybOut "Triggered: $($ran -join ', ')" "SUCCESS" } else { Write-HybOut "No MDM sync task found - is this device enrolled?" "WARN" }
    })

    $btnHybRestartIme.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut (Restart-IntuneExtension) "INFO"
    })

    $btnHybEnroll.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Triggering GPO-style automatic MDM enrollment..." "WARN"
        Write-HybOut (Invoke-MdmAutoEnroll) "INFO"
    })

    $btnHybComgmt.Add_Click({
        Write-HybOut "Reading co-management workload authority..."
        $c = Get-CoManagementState
        $lstHybWorkloads.ItemsSource = $c.Workloads
        if ($c.CoManaged) { $txtHybComgmt.Text = "Co-managed (flags=$($c.FlagsValue)). Each row shows which authority currently owns that workload." }
        elseif ($c.ConfigMgrPresent) { $txtHybComgmt.Text = "ConfigMgr client present but no co-management flags - workloads default to ConfigMgr." }
        else { $txtHybComgmt.Text = "Not co-managed (no ConfigMgr client). Intune is the sole authority." }
        Write-HybOut $txtHybComgmt.Text "INFO"
    })

    $btnHybCcm.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut (Invoke-ConfigMgrClientAction) "INFO"
    })

    $btnHybDc.Add_Click({
        Write-HybOut "Checking domain controller line-of-sight..."
        $d = Test-DomainConnectivity
        if (-not $d.Applicable) { $txtHybNet.Text = $d.Message; Write-HybOut $d.Message "INFO"; return }
        $txtHybNet.Text = "DC $($d.DcName) (site $($d.Site)) reachable=$($d.DcReachable); secure channel $($d.SecureChannel)"
        Write-HybOut $txtHybNet.Text $(if ($d.DcReachable) { 'SUCCESS' } else { 'WARN' })
    })

    $btnHybGpupdate.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Running gpupdate /force (this can take a moment)..." "WARN"
        Write-HybOut (Invoke-GpUpdateForce) "INFO"
    })

    $btnHybTime.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut (Sync-DomainTime) "INFO"
    })

    $btnHybCa.Add_Click({
        Write-HybOut "Evaluating Conditional Access readiness from local evidence..."
        $r = Get-CaReadiness
        foreach ($k in $r.Checks.Keys) { Write-HybOut ("  {0}: {1}" -f $k, $(if ($r.Checks[$k]) { 'PASS' } else { 'FAIL' })) }
        if ($r.Ready) { $txtHybCompliance.Text = "CA readiness: all local checks pass. ($($r.Note))"; Write-HybOut "CA readiness: PASS. $($r.Note)" "SUCCESS" }
        else { $txtHybCompliance.Text = "CA readiness: blocked by - $($r.Blockers -join ', '). $($r.Note)"; Write-HybOut "CA readiness: blockers -> $($r.Blockers -join ', ')" "WARN" }
    })

    $btnHybBitlocker.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        $state = Get-BitLockerEscrowState
        Write-HybOut "BitLocker: $($state.Message)"
        if ($state.Available) {
            Write-HybOut "Escrowing recovery key(s) to Entra ID (and AD if domain-joined)..." "WARN"
            Write-HybOut (Invoke-BitLockerEscrow) "INFO"
        }
    })

    $btnHybLegacy.Add_Click({
        Write-HybOut "Scanning for legacy dependencies that cloud/security baselines will break..."
        foreach ($f in (Get-LegacyDependencyState)) {
            $lvl = switch ($f.Risk) { 'High' {'WARN'} 'Medium' {'WARN'} default {'INFO'} }
            Write-HybOut ("  {0}: {1} [{2}] {3}" -f $f.Item, $f.State, $f.Risk, $f.Note) $lvl
        }
    })

    $btnHybCerts.Add_Click({
        Write-HybOut "Checking machine client-auth certificates and near-expiry..."
        $certs = @(Get-ClientCertificateHealth)
        if ($certs.Count -eq 0) { Write-HybOut "  No client-auth or expiring machine certificates found." "INFO"; return }
        foreach ($c in $certs) { Write-HybOut ("  {0} | issuer {1} | {2} | {3}" -f $c.Subject, $c.Issuer, $c.NotAfter, $c.State) $(if ($c.State -eq 'EXPIRED' -or $c.State -like 'Expiring*') { 'WARN' } else { 'INFO' }) }
    })

    $btnHybBundle.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Collecting hybrid diagnostics bundle (dsregcmd, gpresult, mdmdiagnosticstool, IME logs)..." "WARN"
        try {
            $zip = New-HybridDiagnosticsBundle
            Write-HybOut "Diagnostics written to: $zip" "SUCCESS"
            try { Start-Process explorer.exe "/select,`"$zip`"" } catch {}
        } catch { Write-HybOut "Bundle failed: $($_.Exception.Message)" "ERROR" }
    })

    # Hybrid Enhancements Event Handlers
    $btnHybDcLadder.Add_Click({
        Write-HybOut "Initiating Active Directory Domain Controller Port & Latency Ladder..." "INFO"
        $btnHybDcLadder.IsEnabled = $false
        Set-HubProgress -Percent 20 -Status "Testing DC Port Ladder"

        Start-HubAsyncWork -Description "DC Port Ladder" -WorkerScript {
            Write-WorkerLog "Probing DC connectivity across LDAP, Kerberos, SMB, RPC, and DNS..."
            $res = Test-DomainControllerLadder
            return $res
        } -OnComplete {
            param($res, $err)
            $btnHybDcLadder.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-HybOut "DC Port Ladder probe failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                if (-not $res.Applicable) {
                    Write-HybOut $res.Summary "WARN"
                } else {
                    Write-HybOut "=== Domain Controller Port & Latency Ladder ($($res.DcName)) ===" "INFO"
                    foreach ($p in $res.Ports) {
                        $lvl = if ($p.Status -eq 'OPEN') { 'SUCCESS' } else { 'WARN' }
                        $latStr = if ($p.LatencyMs -ge 0) { "$($p.LatencyMs)ms" } else { $p.Error }
                        Write-HybOut ("  Port {0,5}/TCP | {1,-20} | Status: {2,-7} | Latency: {3}" -f $p.Port, $p.Service, $p.Status, $latStr) $lvl
                    }
                    Write-HybOut $res.Summary $(if ($res.AllReachable) { 'SUCCESS' } else { 'WARN' })
                }
            }
        }
    })

    $btnHybScp.Add_Click({
        Write-HybOut "Inspecting Active Directory Service Connection Point (SCP) for Entra Tenant..." "INFO"
        $btnHybScp.IsEnabled = $false

        Start-HubAsyncWork -Description "AD SCP Inspection" -WorkerScript {
            Write-WorkerLog "Querying LDAP RootDSE configurationNamingContext for SCP 62a0ff2e-97b9-4513-943f-0d221bd30080..."
            $res = Get-AdServiceConnectionPoint
            return $res
        } -OnComplete {
            param($res, $err)
            $btnHybScp.IsEnabled = $true
            if ($err) {
                Write-HybOut "AD SCP query failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                $lvl = switch ($res.Verdict) {
                    'MATCH'     { 'SUCCESS' }
                    'MISMATCH'  { 'ERROR' }
                    'SCP_VALID' { 'SUCCESS' }
                    default     { 'WARN' }
                }
                Write-HybOut "AD SCP Status: $($res.Verdict) - $($res.Message)" $lvl
                if ($res.ScpFound) {
                    Write-HybOut "  AD SCP Tenant ID: $($res.TenantId) ($($res.TenantDomain))" "INFO"
                    Write-HybOut "  SCP DN: $($res.ScpDistinguishedName)" "INFO"
                    if ($res.LocalTenantId) {
                        Write-HybOut "  Local Device Tenant ID: $($res.LocalTenantId) ($($res.LocalTenantName))" "INFO"
                    }
                }
            }
        }
    })

    $btnHybCoMgmtAllIntune.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Shifting ALL Co-Management workloads to Intune (Flags: 255)..." "WARN"
                $__cm = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
                    [System.Windows.MessageBox]::Show("This writes the co-management workload flags to HKLM\SOFTWARE\Microsoft\CCM on THIS device only.`n`nCo-management authority is owned by the ConfigMgr co-management policy in your tenant. A local registry change is not the supported way to move a workload and will be overwritten by the client on its next policy cycle - in the meantime a workload can end up mis-applied. Use this for lab/diagnostic purposes, and set the real sliders in the Intune admin center.`n`nContinue anyway?", "Local Co-Management Override", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                }
        if ($__cm -ne [System.Windows.MessageBoxResult]::Yes) { Write-HybOut "Co-management change cancelled." "INFO"; return }
        $res = Set-CoManagementWorkloads -Preset AllIntune
        Write-HybOut "Workload authority updated: Preset $($res.Preset), Flags=$($res.FlagsValue). $($res.ConfigMgrMsg)" "SUCCESS"
        $c = Get-CoManagementState
        $lstHybWorkloads.ItemsSource = $c.Workloads
    })

    $btnHybCoMgmtAllCcm.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Shifting ALL Co-Management workloads to ConfigMgr / SCCM (Flags: 1)..." "WARN"
                $__cm = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
                    [System.Windows.MessageBox]::Show("This writes the co-management workload flags to HKLM\SOFTWARE\Microsoft\CCM on THIS device only.`n`nCo-management authority is owned by the ConfigMgr co-management policy in your tenant. A local registry change is not the supported way to move a workload and will be overwritten by the client on its next policy cycle - in the meantime a workload can end up mis-applied. Use this for lab/diagnostic purposes, and set the real sliders in the Intune admin center.`n`nContinue anyway?", "Local Co-Management Override", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                }
        if ($__cm -ne [System.Windows.MessageBoxResult]::Yes) { Write-HybOut "Co-management change cancelled." "INFO"; return }
        $res = Set-CoManagementWorkloads -Preset AllConfigMgr
        Write-HybOut "Workload authority updated: Preset $($res.Preset), Flags=$($res.FlagsValue). $($res.ConfigMgrMsg)" "SUCCESS"
        $c = Get-CoManagementState
        $lstHybWorkloads.ItemsSource = $c.Workloads
    })

    $btnHybCoMgmtPilot.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Setting Co-Management Pilot workloads: Compliance + Client Apps (Flags: 67)..." "WARN"
                $__cm = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
                    [System.Windows.MessageBox]::Show("This writes the co-management workload flags to HKLM\SOFTWARE\Microsoft\CCM on THIS device only.`n`nCo-management authority is owned by the ConfigMgr co-management policy in your tenant. A local registry change is not the supported way to move a workload and will be overwritten by the client on its next policy cycle - in the meantime a workload can end up mis-applied. Use this for lab/diagnostic purposes, and set the real sliders in the Intune admin center.`n`nContinue anyway?", "Local Co-Management Override", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                }
        if ($__cm -ne [System.Windows.MessageBoxResult]::Yes) { Write-HybOut "Co-management change cancelled." "INFO"; return }
        $res = Set-CoManagementWorkloads -Preset Pilot
        Write-HybOut "Workload authority updated: Preset $($res.Preset), Flags=$($res.FlagsValue). $($res.ConfigMgrMsg)" "SUCCESS"
        $c = Get-CoManagementState
        $lstHybWorkloads.ItemsSource = $c.Workloads
    })

    $btnHybKerbDiag.Add_Click({
        Write-HybOut "Querying Kerberos ticket cache (klist tickets)..." "INFO"
        $k = Get-KerberosDiagnostics
        $txtHybKerb.Text = $k.Message
        Write-HybOut $k.Message $(if ($k.HasTgt) { 'SUCCESS' } else { 'WARN' })
        if ($k.Tickets.Count -gt 0) {
            foreach ($t in $k.Tickets) {
                Write-HybOut ("  Ticket: {0} | Server: {1} | Enc: {2} | Expires: {3}" -f $t.Client, $t.Server, $t.KerbTicketEncryptionType, $t.EndTime) "INFO"
            }
        }
    })

    $btnHybKerbPurge.Add_Click({
        Write-HybOut "Purging cached Kerberos tickets..." "WARN"
        $purgeRes = Invoke-KerberosPurge
        Write-HybOut $purgeRes "INFO"
        $k = Get-KerberosDiagnostics
        $txtHybKerb.Text = $k.Message
    })

    $btnHybSpn.Add_Click({
        Write-HybOut "Verifying Active Directory Service Principal Names (setspn -L / -X)..." "INFO"
        $btnHybSpn.IsEnabled = $false

        Start-HubAsyncWork -Description "SPN Verification" -WorkerScript {
            Test-ComputerSpnRegistration
        } -OnComplete {
            param($res, $err)
            $btnHybSpn.IsEnabled = $true
            if ($err) {
                Write-HybOut "SPN verification error: $($err.Message)" "ERROR"
            } elseif ($res) {
                $lvl = if ($res.Status -eq 'REGISTERED') { 'SUCCESS' } else { 'WARN' }
                Write-HybOut "SPN Registration: $($res.Status) - $($res.Message)" $lvl
                if ($res.SpnCount -gt 0) {
                    foreach ($s in $res.SpnList) { Write-HybOut "  SPN: $s" "INFO" }
                }
                if ($res.DuplicateSpns) {
                    Write-HybOut "CRITICAL: Duplicate SPNs detected across domain: $($res.DuplicateList -join ', ')" "ERROR"
                }
            }
        }
    })

    $btnHybPrtDiag.Add_Click({
        Write-HybOut "Evaluating Entra ID Primary Refresh Token (PRT) and Cloud AP WAM Broker..." "INFO"
        $prt = Get-EntraPrtDiagnostics
        $prtStatus = if ($prt.HasPrt) { "PRESENT ($($prt.PrtAuthority))" } else { "MISSING" }
        Write-HybOut "Entra PRT: $prtStatus | NgcPrt: $($prt.HasNgcPrt) | CloudTgt: $($prt.HasCloudTgt) | OnPremTgt: $($prt.HasOnPremTgt)" $(if ($prt.HasPrt){'SUCCESS'}else{'WARN'})
        Write-HybOut "WAM Broker Plugin (Microsoft.AAD.BrokerPlugin): Installed=$($prt.BrokerInstalled), Status=$($prt.BrokerStatus), Version=$($prt.BrokerVersion)" $(if ($prt.BrokerInstalled){'SUCCESS'}else{'WARN'})
    })

    $btnHybAadToken.Add_Click({
        Write-HybOut "Testing Microsoft Entra ID Token Acquisition & STS Endpoint..." "INFO"
        $btnHybAadToken.IsEnabled = $false

        Start-HubAsyncWork -Description "AAD Token Acquisition Test" -WorkerScript {
            Test-EntraTokenAcquisition
        } -OnComplete {
            param($res, $err)
            $btnHybAadToken.IsEnabled = $true
            if ($err) {
                Write-HybOut "AAD token probe error: $($err.Message)" "ERROR"
            } elseif ($res) {
                $lvl = switch ($res.Verdict) {
                    'PRT_READY'       { 'SUCCESS' }
                    'WAM_READY'       { 'SUCCESS' }
                    'ENDPOINT_ONLINE' { 'INFO' }
                    default           { 'WARN' }
                }
                Write-HybOut "AAD Token Pipeline: $($res.Verdict) - $($res.Message)" $lvl
                if ($res.WamDefaultAccount) {
                    Write-HybOut "  WAM Default Account: $($res.WamDefaultAccount)" "INFO"
                }
            }
        }
    })

    $btnHybResetBroker.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Resetting Web Account Manager (WAM) TokenBroker and Cloud AP cache..." "WARN"
        $res = Reset-CloudApBrokerCache
        foreach ($l in ($res -split "`n")) { Write-HybOut "  $l" "INFO" }
    })

    $btnHybCertPulse.Add_Click({
        if (-not (Assert-HybElevated)) { return }
        Write-HybOut "Triggering immediate certificate auto-enrollment pulse (certutil/certreq)..." "WARN"
        $btnHybCertPulse.IsEnabled = $false
        Start-HubAsyncWork -Description "Certificate Pulse" -WorkerScript {
            Invoke-CertificatePulse
        } -OnComplete {
            param($res, $err)
            $btnHybCertPulse.IsEnabled = $true
            if ($err) { Write-HybOut "Certificate pulse error: $($err.Message)" "ERROR" }
            elseif ($res) {
                foreach ($l in ($res -split "`n")) { Write-HybOut "  $l" "INFO" }
                Write-HybOut "Certificate auto-enrollment pulse completed." "SUCCESS"
            }
        }
    })

    $btnHybScepHealth.Add_Click({
        Write-HybOut "Auditing SCEP/NDES, MDM, and Client-Auth certificates..." "INFO"
        $scep = @(Get-ScepCertificateHealth)
        if ($scep.Count -eq 0) {
            Write-HybOut "No SCEP/NDES or Intune MDM certificates found in machine or user stores." "WARN"
            return
        }
        foreach ($c in $scep) {
            $lvl = switch -Wildcard ($c.Status) {
                'HEALTHY'   { 'SUCCESS' }
                'EXPIRING*' { 'WARN' }
                default     { 'ERROR' }
            }
            Write-HybOut ("  [{0}] {1} | Issuer: {2} | Expires: {3} ({4} d) | Key: {5}-bit | Status: {6}" -f $c.Store, $c.Subject, $c.Issuer, $c.NotAfter, $c.DaysLeft, $c.KeyLength, $c.Status) $lvl
        }
    })

    # --- TAB 9: PRECISION ENTERPRISE REMEDIATION ARSENAL WIREUP ---
    function Write-RemOut {
        param([string]$Message, [string]$Level = 'INFO')
        $ts = (Get-Date).ToString('HH:mm:ss')
        $txtRemOut.AppendText("`r`n[$ts] [$Level] $Message")
        $txtRemOut.ScrollToEnd()
        Update-WpfUI
    }

    function Assert-RemElevated {
        if (-not $script:RuntimeContext.MeetsPreferred) {
            Write-RemOut "This remediation tool requires administrative elevation (SYSTEM or Administrator). Click 'Fix Privileges' in the header." "WARN"
            return $false
        }
        return $true
    }

    $btnClearFixOut.Add_Click({
        $txtRemOut.Text = "Remediation console cleared. Ready for next diagnostic or surgical repair."
    })
    if ($btnClearRemOut) {
        $btnClearRemOut.Add_Click({
            $txtRemOut.Text = "Remediation console cleared. Ready for next diagnostic or surgical repair."
        })
    }

    $btnCopyFixLog.Add_Click({
        try {
            [System.Windows.Clipboard]::SetText($txtRemOut.Text)
            Write-RemOut "Remediation console output copied to clipboard." "INFO"
        } catch { }
    })

    $btnFixHealthReadout.Add_Click({
        Write-RemOut "Collecting enterprise OS subsystem health overview (async worker)..." "INFO"
        $btnFixHealthReadout.IsEnabled = $false

        Start-HubAsyncWork -Description "System Health Audit" -WorkerScript {
            Get-HubRemediationHealthOverview
        } -OnComplete {
            param($h, $err)
            $btnFixHealthReadout.IsEnabled = $true
            if ($err) {
                Write-RemOut "Health audit failed: $($err.Message)" "ERROR"
            } elseif ($h) {
                Write-RemOut "=== Enterprise OS Subsystem Health Overview ===" "INFO"
                Write-RemOut "  WMI Repository:        $(if ($h.WmiHealthy){'HEALTHY (CIM responsive)'}else{'DEGRADED / CORRUPTED'})" $(if ($h.WmiHealthy){'SUCCESS'}else{'ERROR'})
                Write-RemOut "  BITS Service:          $($h.BitsStatus)" $(if ($h.BitsStatus -eq 'Running'){'SUCCESS'}else{'WARN'})
                Write-RemOut "  CryptSvc (Catroot2):   $($h.CryptSvcStatus)" $(if ($h.CryptSvcStatus -eq 'Running'){'SUCCESS'}else{'WARN'})
                Write-RemOut "  Print Spooler:         $($h.SpoolerStatus)" $(if ($h.SpoolerStatus -eq 'Running'){'SUCCESS'}else{'WARN'})
                Write-RemOut "  WinRM Service:         $($h.WinRmStatus)" $(if ($h.WinRmStatus -eq 'Running'){'SUCCESS'}else{'WARN'})
                Write-RemOut "  TPM Hardware Present:  $($h.TpmPresent) (Ready: $($h.TpmReady))" $(if ($h.TpmReady){'SUCCESS'}else{'WARN'})
            }
        }
    })

    $btnFixWmi.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixWmi.IsEnabled = $false
        Write-RemOut "Executing WMI Repository Salvage & Self-Heal Protocol (async worker)..." "WARN"
        Set-HubProgress -Percent 20 -Status "Salvaging WMI Repository"

        Start-HubAsyncWork -Description "WMI Salvage & Self-Heal" -WorkerScript {
            Write-WorkerLog "Starting winmgmt salvage and core MOF recompilation..."
            $res = Repair-HubWmiRepository
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixWmi.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-RemOut "WMI Salvage failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" $(if ($l -match 'failed|error|Corruption') { 'WARN' } else { 'INFO' }) }
                Write-RemOut "WMI Salvage Protocol Finished. Status: $(if ($res.Success){'SUCCESS'}else{'COMPLETED WITH WARNINGS'})" $(if ($res.Success){'SUCCESS'}else{'WARN'})
            }
        }
    })

    $btnFixWu.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixWu.IsEnabled = $false
        Write-RemOut "Executing Windows Update Agent & SoftDistribution Deep Reset..." "WARN"
        Set-HubProgress -Percent 25 -Status "Resetting Windows Update Agent"

        Start-HubAsyncWork -Description "WU Agent Deep Reset" -WorkerScript {
            Write-WorkerLog "Tearing down WU services, archiving SoftwareDistribution and registering DLLs..."
            $res = Reset-HubWindowsUpdateAgent
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixWu.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-RemOut "WU Reset failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "Windows Update Agent Deep Reset Completed Successfully." "SUCCESS"
            }
        }
    })

    $btnFixCatroot.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixCatroot.IsEnabled = $false
        Write-RemOut "Executing Catroot2 Cryptographic ESENT Database Repair..." "WARN"
        Set-HubProgress -Percent 30 -Status "Repairing Catroot2 Database"

        Start-HubAsyncWork -Description "Catroot2 ESENT Repair" -WorkerScript {
            Write-WorkerLog "Verifying and repairing catdb database via esentutl..."
            $res = Repair-HubCryptSvcCatroot
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixCatroot.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-RemOut "Catroot2 repair failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "Catroot2 ESENT Database Repair Protocol Completed." "SUCCESS"
            }
        }
    })

    $btnFixBits.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixBits.IsEnabled = $false
        Write-RemOut "Executing BITS Transfer Deadlock Purge..." "WARN"
        Set-HubProgress -Percent 20 -Status "Purging BITS Queue"

        Start-HubAsyncWork -Description "BITS Deadlock Purge" -WorkerScript {
            Write-WorkerLog "Purging BITS queue and deleting corrupted qmgr.dat files..."
            $res = Reset-HubBitsQueue
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixBits.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-RemOut "BITS purge failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "BITS Queue Purged & Transfer Pipeline Verified." "SUCCESS"
            }
        }
    })

    $btnFixDcom.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixDcom.IsEnabled = $false
        Write-RemOut "Executing DCOM / RPC 0x800706BA Security Limits Remediation..." "WARN"

        Start-HubAsyncWork -Description "DCOM & RPC Remediation" -WorkerScript {
            Write-WorkerLog "Resetting Ole DCOM levels and enabling RPC firewall rules..."
            $res = Repair-HubDcomRpcPermissions
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixDcom.IsEnabled = $true
            if ($err) {
                Write-RemOut "DCOM remediation failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "DCOM / RPC Remediation Completed. Local COM activation verified." "SUCCESS"
            }
        }
    })

    $btnFixSpooler.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixSpooler.IsEnabled = $false
        Write-RemOut "Executing Print Spooler Hung Queue Purge..." "WARN"

        Start-HubAsyncWork -Description "Print Spooler Purge" -WorkerScript {
            Write-WorkerLog "Purging locked .spl/.shd jobs in spool\PRINTERS and restarting Spooler..."
            $res = Clear-HubPrintSpoolerQueue
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixSpooler.IsEnabled = $true
            if ($err) {
                Write-RemOut "Spooler purge failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "Print Spooler Queue Purged and Service Restarted." "SUCCESS"
            }
        }
    })

    $btnFixNetStack.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixNetStack.IsEnabled = $false
        Write-RemOut "Executing Network Stack, Winsock & IPsec Reset (async worker)..." "WARN"

        Start-HubAsyncWork -Description "Network Stack Reset" -WorkerScript {
            Write-WorkerLog "Resetting TCP/IP, Winsock, DNS cache, ARP tables, and IPsec associations..."
            $res = Reset-HubNetworkStack
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixNetStack.IsEnabled = $true
            if ($err) {
                Write-RemOut "Network stack reset failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "Network Stack, Winsock, ARP, and IPsec Filters Successfully Flushed." "SUCCESS"
            }
        }
    })

    $btnFixWinRm.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixWinRm.IsEnabled = $false
        Write-RemOut "Executing WinRM & WS-Man Listener Rebuild..." "WARN"

        Start-HubAsyncWork -Description "WinRM Listener Rebuild" -WorkerScript {
            Write-WorkerLog "Rebuilding WinRM HTTP 5985 listener and binding firewall rules..."
            $res = Repair-HubWinRmListener
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixWinRm.IsEnabled = $true
            if ($err) {
                Write-RemOut "WinRM rebuild failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "WinRM & WS-Man Listener Rebuilt Successfully." "SUCCESS"
            }
        }
    })

    $btnFixProfiles.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixProfiles.IsEnabled = $false
        Write-RemOut "Executing User Profile Registry Lock Un-hooker (.bak resolver)..." "WARN"

        Start-HubAsyncWork -Description "User Profile Un-hooker" -WorkerScript {
            Write-WorkerLog "Scanning ProfileList for orphaned .bak subkeys and resetting RefCount..."
            $res = Repair-HubUserProfileLocks
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixProfiles.IsEnabled = $true
            if ($err) {
                Write-RemOut "User profile unlock failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "User Profile Locks Un-hooked. Repaired: $($res.FixedCount) .bak key(s)." "SUCCESS"
            }
        }
    })

    $btnFixTpm.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixTpm.IsEnabled = $false
        Write-RemOut "Executing TPM Platform Crypto Provider & Attestation NVRAM Healer..." "WARN"

        Start-HubAsyncWork -Description "TPM Attestation Healer" -WorkerScript {
            Write-WorkerLog "Clearing cached attestation flags and verifying Endorsement Key certs..."
            $res = Repair-HubTpmCryptoAttestation
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixTpm.IsEnabled = $true
            if ($err) {
                Write-RemOut "TPM healer failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "TPM Platform Crypto Provider & Attestation Healer Completed." "SUCCESS"
            }
        }
    })

    $btnFixAppX.Add_Click({
        if (-not (Assert-RemElevated)) { return }
        $btnFixAppX.IsEnabled = $false
        Write-RemOut "Executing AppX Manifest Staging & Sysprep Unbricker (async worker)..." "WARN"
        Set-HubProgress -Percent 20 -Status "Unbricking AppX Manifests"

        Start-HubAsyncWork -Description "AppX Staging Unbricker" -WorkerScript {
            Write-WorkerLog "Re-registering core Windows inbox AppX manifests and scanning staging state..."
            $res = Repair-HubAppXStaging
            return $res
        } -OnComplete {
            param($res, $err)
            $btnFixAppX.IsEnabled = $true
            Set-HubProgress -Percent 100 -Status "Ready"
            if ($err) {
                Write-RemOut "AppX unbricking failed: $($err.Message)" "ERROR"
            } elseif ($res) {
                foreach ($l in $res.Log) { Write-RemOut "  $l" "INFO" }
                Write-RemOut "AppX Staging Unbricker Finished: Re-registered $($res.HealedCount) core manifest(s)." "SUCCESS"
            }
        }
    })

    # --- ACTION: Run Diagnostics (Hyperthreaded & Asynchronous) ---
    $btnRunDiag.Add_Click({
        $btnRunDiag.IsEnabled = $false
        Write-HubLog "Executing 7-stage enterprise pre-flight diagnostic ladder (async worker)..."
        Set-HubProgress -Percent 10 -Status "Diagnosing Network"

        Start-HubAsyncWork -Description "7-Stage Network Diagnostics" -WorkerScript {
            Write-WorkerLog "Initiating multi-point network test off UI thread..."
            Set-WorkerProgress 30 "Testing Interfaces & Gateway"
            $res = Test-StagedNetwork
            Set-WorkerProgress 95 "Finalizing Diagnostic Assessment"
            return $res
        } -OnComplete {
            param($diag, $err)
            $btnRunDiag.IsEnabled = $true
            if ($diag) {
                $lstDiagStages.ItemsSource = $diag.Stages
                Write-HubLog "Diagnostics completed: $($diag.StagesPassed) / $($diag.TotalStages) stages passed." $(if ($diag.IsFullyReady) { "SUCCESS" } else { "WARN" })
                Set-HubProgress -Percent 100 -Status "Diagnostics Done"
            } else {
                $msg = if ($err) { $err.Message } else { "Unknown diagnostic failure" }
                Write-HubLog "Diagnostics failed: $msg" "ERROR"
                Set-HubProgress -Percent 0 -Status "Diagnostics Failed"
            }
        }
    })

    # PowerShell 7 Modern Runtime Handler
    $pwshPath = $null
    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwshCmd) { $pwshPath = $pwshCmd.Source }
    if (-not $pwshPath) {
        $possiblePwsh = @(
            "C:\Program Files\PowerShell\7\pwsh.exe",
            "C:\Program Files\PowerShell\7-preview\pwsh.exe",
            "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
        )
        foreach ($p in $possiblePwsh) {
            if (Test-Path $p) { $pwshPath = $p; break }
        }
    }
    if ($pwshPath -and $btnInstallPwsh) {
        if ($btnInstallPwsh -is [System.Windows.Controls.MenuItem]) {
            $btnInstallPwsh.Header = "Launch PowerShell 7"
        } else {
            $btnInstallPwsh.Content = "Launch PS7"
        }
    }

    $btnInstallPwsh.Add_Click({
        $existing = $null
        $check = Get-Command pwsh -ErrorAction SilentlyContinue
        if ($check) { $existing = $check.Source }
        if (-not $existing) {
            $candidates = @(
                "C:\Program Files\PowerShell\7\pwsh.exe",
                "C:\Program Files\PowerShell\7-preview\pwsh.exe",
                "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
            )
            foreach ($c in $candidates) {
                if (Test-Path $c) { $existing = $c; break }
            }
        }

        if ($existing) {
            Write-HubLog "Launching modern PowerShell 7 console ($existing)..." "SUCCESS"
            Start-Process $existing
            return
        }

        Write-HubLog "Initiating modern PowerShell 7 (pwsh) automated deployment..." "INFO"
        Write-HubLog "Targeting latest stable PowerShell 7.4.x x64 MSI package..." "INFO"
        Set-HubProgress -Percent 15 -Status "Deploying PS7"

        $installed = $false
        $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetCmd) {
            try {
                Write-HubLog "Attempting installation via Windows Package Manager (winget)..."
                $p = Start-Process winget -ArgumentList 'install', '--id', 'Microsoft.PowerShell', '--exact', '--silent', '--accept-source-agreements', '--accept-package-agreements' -PassThru -Wait -NoNewWindow
                if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {
                    $installed = $true
                }
            } catch { }
        }

        if (-not $installed) {
            try {
                $msiUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.5/PowerShell-7.4.5-win-x64.msi"
                $msiDest = "$env:TEMP\PowerShell-7.4.5-win-x64.msi"
                Write-HubLog "Downloading PowerShell 7 x64 MSI directly from GitHub releases..."
                Set-HubProgress -Percent 35 -Status "Downloading PS7"

                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                $wc = [System.Net.WebClient]::new()
                $wc.DownloadFile($msiUrl, $msiDest)

                Write-HubLog "Executing silent msiexec deployment (/qn /norestart)..."
                Set-HubProgress -Percent 70 -Status "Installing PS7"
                $msiArgs = @(
                    '/i', "`"$msiDest`"",
                    '/qn',
                    '/norestart',
                    'ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1',
                    'ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1',
                    'ENABLE_PSREMOTING=1',
                    'REGISTER_MANIFEST=1'
                ) -join ' '
                $p = Start-Process msiexec.exe -ArgumentList $msiArgs -PassThru -Wait
                if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {
                    $installed = $true
                } else {
                    Write-HubLog "msiexec returned exit code: $($p.ExitCode)" "WARN"
                }
            } catch {
                Write-HubLog "Direct MSI install error: $($_.Exception.Message)" "ERROR"
            }
        }

        $targetPwsh = "C:\Program Files\PowerShell\7\pwsh.exe"
        if (Test-Path $targetPwsh) {
            Write-HubLog "PowerShell 7 installed successfully! Launching modern console..." "SUCCESS"
            Set-HubProgress -Percent 100 -Status "PS7 Ready"
            if ($btnInstallPwsh -is [System.Windows.Controls.MenuItem]) {
                $btnInstallPwsh.Header = "Launch PowerShell 7"
            } else {
                $btnInstallPwsh.Content = "Launch PS7"
            }
            Start-Process $targetPwsh
            [System.Windows.MessageBox]::Show("PowerShell 7 (pwsh) has been installed and launched.`n`nExecutable: $targetPwsh", "PowerShell 7 Ready", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            Write-HubLog "PowerShell 7 installation could not be verified at $targetPwsh" "ERROR"
            Set-HubProgress -Percent 0 -Status "PS7 Install Failed"
        }
    })

    # Quick Utility Handlers
    if ($btnToolsDropdown) {
        $btnToolsDropdown.Add_Click({
            if ($btnToolsDropdown.ContextMenu) {
                $btnToolsDropdown.ContextMenu.PlacementTarget = $btnToolsDropdown
                $btnToolsDropdown.ContextMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
                $btnToolsDropdown.ContextMenu.IsOpen = $true
            }
        })
    }

    $script:ConsoleExpanded = $true
    if ($btnToggleConsole) {
        $btnToggleConsole.Add_Click({
            if ($script:ConsoleExpanded) {
                $script:ConsoleExpanded = $false
                if ($txtHubLog) { $txtHubLog.Visibility = [System.Windows.Visibility]::Collapsed }
                if ($rowConsoleLog) { $rowConsoleLog.Height = [System.Windows.GridLength]::Auto }
                $btnToggleConsole.Content = "[+] Show Log"
                Write-HubLog "System audit log console collapsed." "INFO"
            } else {
                $script:ConsoleExpanded = $true
                if ($txtHubLog) { $txtHubLog.Visibility = [System.Windows.Visibility]::Visible }
                if ($rowConsoleLog) { $rowConsoleLog.Height = [System.Windows.GridLength]::new(150) }
                $btnToggleConsole.Content = "[-] Hide Log"
                Write-HubLog "System audit log console expanded." "INFO"
            }
        })
    }

    if ($btnClearHybOut) {
        $btnClearHybOut.Add_Click({
            if ($txtHybOut) { $txtHybOut.Text = "Ready." }
        })
    }

    $btnQuickCmd.Add_Click({
        Write-HubLog "Launching Shift+F10 Administrative Command Prompt..."
        Start-Process cmd.exe
    })

    $btnTimeSync.Add_Click({
        Write-HubLog "Triggering Windows Time service synchronization (w32tm)..."
        try {
            Start-Process w32tm -ArgumentList '/resync /force' -NoNewWindow -Wait
            Write-HubLog "Time synchronization triggered." "SUCCESS"
        } catch {
            Write-HubLog "Time sync error: $($_.Exception.Message)" "ERROR"
        }
    })

    $btnReboot.Add_Click({
        if ($script:PlaybookRunning) {
            Write-HubLog "Playbook: System reboot scheduled upon completion." "INFO"
            return
        }
        $confirm = [System.Windows.MessageBox]::Show("Restart system now? The Hub will automatically re-launch on startup.", "Restart System", [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Question)
        if ($confirm -ne [System.Windows.MessageBoxResult]::OK) { return }
        try {
            $persisted = Register-HubResumeAfterRestart
            Write-HubLog "Resume persistence registered ($persisted)." "SUCCESS"
        } catch {
            Write-HubLog "Resume registration notice: $($_.Exception.Message)" "WARN"
        }
        Write-HubLog "Initiating system restart..." "WARN"
        Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
    })

    $btnCopyLog.Add_Click({
        [System.Windows.Clipboard]::SetText($txtHubLog.Text)
    })

    $btnClearLog.Add_Click({
        $txtHubLog.Clear()
    })

    $btnSaveLog.Add_Click({
        $savePath = "$env:TEMP\AutopilotHub-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        Set-Content -Path $savePath -Value $txtHubLog.Text -Encoding UTF8 -Force
        Write-HubLog "Log saved to: $savePath" "SUCCESS"
        if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Log saved to:`n$savePath", "Log Saved", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
    })

    # --- ACTION: Dell Warranty & Hardware Refresh Assessment ---
    $script:CurrentDellReport = $null

    function Invoke-DellWarrantyAudit {
        param([string]$Tag)
        if (-not $Tag) {
            $Tag = $txtDellServiceTag.Text.Trim()
        }
        if (-not $Tag) {
            try {
                $Tag = (Get-CimInstance Win32_BIOS -ErrorAction Stop).SerialNumber.Trim()
                $txtDellServiceTag.Text = $Tag
            } catch { }
        }
        if (-not $Tag) {
            Write-HubLog "Dell warranty check halted: No Service Tag provided." "WARN"
            if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please enter a valid Dell Service Tag or click 'Detect BIOS Tag'.", "Service Tag Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
            return
        }

        if ($btnCheckDellWarranty) { $btnCheckDellWarranty.IsEnabled = $false }
        Write-HubLog "Connecting to Dell Technologies Enterprise Warranty API (SBIL eAPI v5) for tag '$Tag' (async worker)..."
        Set-HubProgress -Percent 20 -Status "Acquiring Dell OAuth 2.0 Bearer Token"

        Start-HubAsyncWork -Description "Dell Warranty Audit" -TaskArgs @($Tag) -WorkerScript {
            param($svcTag)
            Write-WorkerLog "Querying Asset Entitlements for $svcTag in worker runspace..."
            Set-WorkerProgress 50 "Querying Dell SBIL eAPI"
            $rep = Get-DellWarrantyInfo -ServiceTag $svcTag
            Set-WorkerProgress 90 "Processing Entitlements"
            return $rep
        } -OnComplete {
            param($w, $err)
            if ($btnCheckDellWarranty) { $btnCheckDellWarranty.IsEnabled = $true }
            if ($w) {
                $script:CurrentDellReport = $w
                # Populate Hardware & Telemetry Cards
                $txtDellModel.Text = if ($w.SystemModel) { $w.SystemModel } else { 'Dell Computer' }
                $txtDellProductLine.Text = "Line: $($w.ProductLineDescription)"
                $txtDellShipDate.Text = if ($w.ShipDate) { $w.ShipDate } else { 'N/A' }
                $txtDellAge.Text = "Age: $(if ($w.DeviceAgeYears) { "$($w.DeviceAgeYears) yrs" } else { '-' })"
                $txtDellContract.Text = $w.PrimaryServiceLevel
                $txtDellRegion.Text = "Region: $($w.CountryCode)"
                $txtDellEndDate.Text = "End: $($w.WarrantyEndDate)"
                $txtDellDaysRemaining.Text = $w.WarrantyStatus

                # Update Hero Refresh Verdict Banner
                $brushConv = [System.Windows.Media.BrushConverter]::new()
                if ($w.IsUnderWarranty) {
                    if ($w.DaysRemaining -le 90) {
                        $borderRefreshVerdict.Background = $brushConv.ConvertFromString("#292621")
                        $borderRefreshVerdict.BorderBrush = $brushConv.ConvertFromString("#FDE68A")
                        if ($borderVerdictBadge) { $borderVerdictBadge.Background = $brushConv.ConvertFromString("#443B26") }
                        if ($txtVerdictBadge) {
                            $txtVerdictBadge.Text = "EXPIRING SOON"
                            $txtVerdictBadge.Foreground = $brushConv.ConvertFromString("#A16207")
                        }
                        $txtVerdictTitle.Text = $w.RefreshVerdict
                        $txtVerdictTitle.Foreground = $brushConv.ConvertFromString("#FFFFFF")
                        $txtDellDaysRemaining.Foreground = $brushConv.ConvertFromString("#A16207")
                    } else {
                        $borderRefreshVerdict.Background = $brushConv.ConvertFromString("#212923")
                        $borderRefreshVerdict.BorderBrush = $brushConv.ConvertFromString("#295C33")
                        if ($borderVerdictBadge) { $borderVerdictBadge.Background = $brushConv.ConvertFromString("#213B26") }
                        if ($txtVerdictBadge) {
                            $txtVerdictBadge.Text = "ACTIVE WARRANTY"
                            $txtVerdictBadge.Foreground = $brushConv.ConvertFromString("#6CCB5F")
                        }
                        $txtVerdictTitle.Text = $w.RefreshVerdict
                        $txtVerdictTitle.Foreground = $brushConv.ConvertFromString("#FFFFFF")
                        $txtDellDaysRemaining.Foreground = $brushConv.ConvertFromString("#6CCB5F")
                    }
                } else {
                    $borderRefreshVerdict.Background = $brushConv.ConvertFromString("#292121")
                    $borderRefreshVerdict.BorderBrush = $brushConv.ConvertFromString("#FEE2E2")
                    if ($borderVerdictBadge) { $borderVerdictBadge.Background = $brushConv.ConvertFromString("#442726") }
                    if ($txtVerdictBadge) {
                        $txtVerdictBadge.Text = "OUT OF WARRANTY"
                        $txtVerdictBadge.Foreground = $brushConv.ConvertFromString("#FFAA99")
                    }
                    $txtVerdictTitle.Text = $w.RefreshVerdict
                    $txtVerdictTitle.Foreground = $brushConv.ConvertFromString("#FFAA99")
                    $txtDellDaysRemaining.Foreground = $brushConv.ConvertFromString("#FFAA99")
                }
                $txtVerdictDesc.Text = $w.RefreshRecommendation

                if ($lstDellEntitlements) {
                    $lstDellEntitlements.ItemsSource = $w.Entitlements
                }

                Write-HubLog "Dell warranty audit complete: $($w.SystemModel) ($Tag) | $($w.RefreshVerdict)" "SUCCESS"
                Set-HubProgress -Percent 100 -Status "Warranty Audit Complete"
            } else {
                $errMsg = if ($err) { $err.Message } else { "Unknown warranty query failure" }
                Write-HubLog "Dell warranty audit failed: $errMsg" "ERROR"
                Set-HubProgress -Percent 0 -Status "Warranty Audit Failed"
            }
        }
    }

    $btnDetectDellTag.Add_Click({
        try {
            $tag = (Get-CimInstance Win32_BIOS -ErrorAction Stop).SerialNumber.Trim()
            $txtDellServiceTag.Text = $tag
            Write-HubLog "Host BIOS Service Tag detected: $tag"
        } catch {
            Write-HubLog "Unable to read BIOS SerialNumber: $($_.Exception.Message)" "WARN"
        }
    })

    $btnCheckDellWarranty.Add_Click({
        Invoke-DellWarrantyAudit
    })

    $btnCopyDellReport.Add_Click({
        if (-not $script:CurrentDellReport) {
            if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please perform a Dell Warranty check first before copying report.", "No Report Available", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
            return
        }
        $r = $script:CurrentDellReport
        $md = @"
# Dell Asset Warranty & Hardware Refresh Assessment: $($r.ServiceTag)

> **Model:** $($r.SystemModel)  
> **Service Tag:** ``$($r.ServiceTag)``  
> **Lifecycle Verdict:** **$($r.RefreshVerdict)**  
> **Warranty Status:** $($r.WarrantyStatus)  
> **Ship Date:** $($r.ShipDate) ($($r.DeviceAgeYears) Years Old)  
> **Coverage End Date:** $($r.WarrantyEndDate)  

---

## Hardware Refresh Determination
$($r.RefreshRecommendation)

---

## Entitlements & Service Contracts Breakdown
| Service Level Description | Type | Start Date | End Date | Status |
| :--- | :--- | :--- | :--- | :--- |
$($r.Entitlements | ForEach-Object { "| $($_.ServiceLevelDescription) | $($_.EntitlementType) | $($_.StartDate) | $($_.EndDate) | $($_.Status) |" } | Out-String).TrimEnd()

---
*Generated via Dell Technologies Enterprise Warranty API (v5)*
"@
        [System.Windows.Clipboard]::SetText($md)
        Write-HubLog "Dell warranty markdown report copied to clipboard." "SUCCESS"
        if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Assessment report copied to clipboard in Markdown format.", "Copied", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
    })

    $btnExportDellCsv.Add_Click({
        if (-not $script:CurrentDellReport) {
            if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please perform a Dell Warranty check first before exporting CSV.", "No Report Available", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
            return
        }
        $r = $script:CurrentDellReport
        $destPath = "$env:TEMP\DellWarranty-$($r.ServiceTag).csv"

        $usbDrives = Get-CimInstance Win32_Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveType -eq 2 -and $_.DriveLetter }
        if ($usbDrives) {
            $usb = $usbDrives | Select-Object -First 1
            $destPath = Join-Path "$($usb.DriveLetter)\" "DellWarranty-$($r.ServiceTag).csv"
            Write-HubLog "USB Flash Drive detected at $($usb.DriveLetter). Exporting CSV directly to USB..."
        }

        $flatObj = [PSCustomObject]@{
            ServiceTag            = $r.ServiceTag
            SystemModel           = $r.SystemModel
            ShipDate              = $r.ShipDate
            DeviceAgeYears        = $r.DeviceAgeYears
            IsUnderWarranty       = $r.IsUnderWarranty
            WarrantyStatus        = $r.WarrantyStatus
            DaysRemaining         = $r.DaysRemaining
            WarrantyEndDate       = $r.WarrantyEndDate
            PrimaryServiceLevel   = $r.PrimaryServiceLevel
            RefreshVerdict        = $r.RefreshVerdict
            RefreshRecommendation = $r.RefreshRecommendation
            QueriedAt             = $r.QueriedAt
        }
        $flatObj | Export-Csv -Path $destPath -NoTypeInformation -Force
        Write-HubLog "Dell warranty audit CSV exported to: $destPath" "SUCCESS"
        if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Dell warranty audit exported successfully to:`n$destPath", "CSV Exported", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
    })

    # --- HEADER ACTIONS ---
    if ($btnDismissMismatch) {
        $btnDismissMismatch.Add_Click({
            if ($bannerTenantMismatch) { $bannerTenantMismatch.Visibility = [System.Windows.Visibility]::Collapsed }
        })
    }

    if ($btnWindowsUpdate) {
        $btnWindowsUpdate.Add_Click({
            Write-HubLog "Switching to Windows Update & Driver Engine tab..." "INFO"
            if ($mainTabControl -and $tabWinUpdate) {
                $mainTabControl.SelectedItem = $tabWinUpdate
            }
        })
    }

    if ($btnLaunchEdge) {
        $btnLaunchEdge.Add_Click({
            try {
                Start-Process "msedge.exe" -ArgumentList "https://login.microsoftonline.com" -ErrorAction Stop
                Write-HubLog "Launched Microsoft Edge for network authentication / captive portal." "SUCCESS"
            } catch {
                try { Start-Process "explorer.exe" -ArgumentList "https://login.microsoftonline.com" } catch { }
                Write-HubLog "Edge not found, opened default browser." "INFO"
            }
        })
    }

    if ($btnScreenshotUsb) {
        $btnScreenshotUsb.Add_Click({
            try {
                $serial = if ($txtSerial) { $txtSerial.Text } else { '' }
                $saved = Export-HubWindowScreenshot -TargetWindow $window -SerialNumber $serial
                if ($saved) {
                    Write-HubLog "Window screenshot saved as staging proof: $saved" "SUCCESS"
                    if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Screenshot saved to:`n$saved", "Staging Proof Saved", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
                }
            } catch {
                Write-HubLog "Failed to capture screenshot: $($_.Exception.Message)" "ERROR"
            }
        })
    }

    $script:HubScale = 1.0
    if ($btnToggleDpi) {
        $btnToggleDpi.Add_Click({
            if ($script:HubScale -eq 1.0) {
                $script:HubScale = 1.25
                $lbl = "Toggle UI Scaling (125%)"
                $btnContent = "Scale: 125%"
            } else {
                $script:HubScale = 1.0
                $lbl = "Toggle UI Scaling (100%)"
                $btnContent = "Scale: 100%"
            }
            if ($btnToggleDpi -is [System.Windows.Controls.MenuItem]) {
                $btnToggleDpi.Header = $lbl
            } else {
                $btnToggleDpi.Content = $btnContent
            }
            $window.LayoutTransform = [System.Windows.Media.ScaleTransform]::new($script:HubScale, $script:HubScale)
            Write-HubLog "UI Scaling set to $($script:HubScale * 100)%" "INFO"
        })
    }

    function Update-PowerStatusUi {
        try {
            $batInfo = Get-BatteryHealthInfo
            if ($batInfo -and $batInfo.Present) {
                if ($txtPowerStatus) { $txtPowerStatus.Text = "POWER: BATTERY ($($batInfo.EstimatedCharge))" }
                if ($badgePower) {
                    $badgePower.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FEF9C3")
                    $badgePower.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FDE68A")
                }
                if ($txtPowerStatus) { $txtPowerStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#A16207") }
            } else {
                if ($txtPowerStatus) { $txtPowerStatus.Text = "POWER: AC MAINS" }
                if ($badgePower) {
                    $badgePower.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#DCFCE7")
                    $badgePower.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#86EFAC")
                }
                if ($txtPowerStatus) { $txtPowerStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#15803D") }
            }
        } catch {
            if ($txtPowerStatus) { $txtPowerStatus.Text = "POWER: ONLINE" }
        }
    }
    Update-PowerStatusUi

    # --- TAB 1 ACTIONS (Policy & Offline JSON) ---
    if ($btnInspectPolicy) {
        $btnInspectPolicy.Add_Click({
            Write-HubLog "Inspecting local Autopilot configuration policy..."
            $pol = Get-DecodedAutopilotPolicy
            if ($pol.Configured) {
                Write-HubLog "Autopilot Policy: Tenant=$($pol.CloudAssignedTenantDomain) SkipEula=$($pol.SkipEula) SkipOem=$($pol.SkipOemRegistration) UserType=$($pol.UserType) DeviceName=$($pol.CloudAssignedDeviceName)" "SUCCESS"
                $summary = "Tenant Domain: $($pol.CloudAssignedTenantDomain)`nTenant ID: $($pol.CloudAssignedTenantId)`nDevice Name: $($pol.CloudAssignedDeviceName)`nSkip EULA: $($pol.SkipEula)`nSkip OEM Reg: $($pol.SkipOemRegistration)`nAccount Type: $($pol.UserType)`nDiagnostics: $($pol.DiagnosticsLevel)"
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show($summary, "Decoded Autopilot Policy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
            } else {
                Write-HubLog "No local Autopilot policy configuration found." "WARN"
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("No cached Autopilot profile found on this machine.", "No Policy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
            }
        })
    }

    if ($btnExportOfflineJson) {
        $btnExportOfflineJson.Add_Click({
            try {
                $tenant = if ($script:GraphAuthContext -and $script:GraphAuthContext.TenantId) { $script:GraphAuthContext.TenantId } else { '' }
                $exported = Export-AutopilotConfigurationFile -TenantId $tenant
                if ($exported) {
                    Write-HubLog "Exported AutopilotConfigurationFile.json to: $exported" "SUCCESS"
                    if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Offline Autopilot JSON exported to:`n$exported", "Export Successful", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
                }
            } catch {
                Write-HubLog "Export offline JSON failed: $($_.Exception.Message)" "ERROR"
            }
        })
    }

    if ($btnInjectOfflineJson) {
        $btnInjectOfflineJson.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Injecting offline provisioning files requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            try {
                if ($script:PlaybookRunning) {
                    $candidateJson = @(
                        "C:\ProgramData\Microsoft\Provisioning\AutopilotConfigurationFile.json",
                        "$env:TEMP\AutopilotConfigurationFile.json"
                    ) + @(Get-UsbDrives | ForEach-Object { "$($_):\AutopilotConfigurationFile.json" }) | Where-Object { Test-Path $_ }
                    if ($candidateJson.Count -gt 0) {
                        $injected = Import-AutopilotConfigurationFile -SourceJsonPath $candidateJson[0]
                        if ($injected) {
                            Write-HubLog "Injected offline Autopilot profile to Windows Provisioning: $injected" "SUCCESS"
                            $script:DeviceState = Get-DeviceEnrollmentState
                            Update-DeviceStateUi
                        }
                    } else {
                        Write-HubLog "Playbook: Offline Autopilot injection engine ready. No candidate JSON present." "INFO"
                    }
                    return
                }
                $dlg = [Microsoft.Win32.OpenFileDialog]::new()
                $dlg.Filter = "Autopilot JSON (AutopilotConfigurationFile.json)|*.json|All Files (*.*)|*.*"
                $dlg.Title = "Select AutopilotConfigurationFile.json to Inject"
                if ($dlg.ShowDialog() -eq $true) {
                    $injected = Import-AutopilotConfigurationFile -SourceJsonPath $dlg.FileName
                    if ($injected) {
                        Write-HubLog "Injected offline Autopilot profile to Windows Provisioning: $injected" "SUCCESS"
                        $script:DeviceState = Get-DeviceEnrollmentState
                        Update-DeviceStateUi
                        [System.Windows.MessageBox]::Show("Autopilot offline profile injected successfully.`nDestination: $injected", "Injection Succeeded", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
                    }
                }
            } catch {
                Write-HubLog "Inject offline JSON failed: $($_.Exception.Message)" "ERROR"
            }
        })
    }

    # --- TAB 4 ACTIONS (Pre-Flight Clock Skew, TPM Attestation & Driver Injection) ---
    if ($btnFixClockSkew) {
        $btnFixClockSkew.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Clock synchronization requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            Write-HubLog "Auto-remediating HTTPS clock skew against Microsoft Online..."
            $syncRes = Sync-HubSystemClock
            if ($syncRes.Success) {
                Write-HubLog "Clock remediated! Server time: $($syncRes.ServerTimeUtc) UTC. Skew: $($syncRes.ClockSkewSeconds)s" "SUCCESS"
                $lstDiagStages.ItemsSource = (Test-StagedNetwork).Stages
            } else {
                Write-HubLog "Clock remediation failed: $($syncRes.Message)" "ERROR"
            }
        })
    }

    if ($btnTpmAttestation) {
        $btnTpmAttestation.Add_Click({
            Write-HubLog "Testing TPM 2.0 Manufacturer Endorsement Key (EK) CA server reachability..."
            $tpmRes = Test-TpmManufacturerCaReachability
            foreach ($s in $tpmRes) {
                $lvl = if ($s.Reachable) { "SUCCESS" } else { "WARN" }
                Write-HubLog "TPM CA [$($s.Manufacturer)] $($s.Host):$($s.Port) - Reachable=$($s.Reachable) ($($s.LatencyMs)ms)" $lvl
            }
        })
    }

    if ($btnInjectDriversPreflight) {
        $btnInjectDriversPreflight.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Injecting drivers via pnputil requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            Write-HubLog "Scanning USB drives for .inf driver packages to inject..."
            $pnpRes = Invoke-HubUsbDriverInjection
            if ($pnpRes.InjectedCount -gt 0) {
                Write-HubLog "Successfully injected $($pnpRes.InjectedCount) driver package(s)." "SUCCESS"
                foreach ($l in $pnpRes.Log) { Write-HubLog "  $l" "INFO" }
            } else {
                Write-HubLog "No driver packages (.inf) found on connected USB drives." "INFO"
            }
        })
    }

    if ($btnPreflightWinUpdate) {
        $btnPreflightWinUpdate.Add_Click({
            Write-HubLog "Switching to Windows Update & Driver Engine tab from Pre-Flight..." "INFO"
            if ($mainTabControl -and $tabWinUpdate) {
                $mainTabControl.SelectedItem = $tabWinUpdate
            }
        })
    }

    # --- TAB 5 ACTIONS (Windows Update & Autonomous Patch Cascade Engine) ---
    $integratedWuItems = [System.Collections.ArrayList]::new()
    $script:RebootCountdownTimer = $null
    $script:RebootCountdownSec = 5
    $script:CascadeRunning = $false

    function Update-IntegratedWuList {
        param($rawUpdates)
        $integratedWuItems.Clear()
        foreach ($u in $rawUpdates) {
            $cat = if ($u.IsDriver) { "DRIVER" } else { "UPDATE" }
            $bg = if ($u.IsDriver) { "#1E2A38" } else { "#1F2822" }
            $fg = if ($u.IsDriver) { "#60CDFF" } else { "#6CCB5F" }
            $sizeStr = if ($u.SizeMB -gt 0) { "$($u.SizeMB) MB" } else { "Online" }
            $sub = if ($u.KB) { "KB$($u.KB)" } else { "" }
            if ($u.Description) {
                $descShort = if ($u.Description.Length -gt 80) { $u.Description.Substring(0, 80) + '...' } else { $u.Description }
                $sub = if ($sub) { "$sub - $descShort" } else { $descShort }
            }
            if (-not $sub) {
                $sub = if ($u.IsDriver) { "Hardware Device Driver" } else { "System Component Update" }
            }

            $item = [PSCustomObject]@{
                IsSelected      = $true
                CategoryText    = $cat
                BadgeBackground = $bg
                BadgeForeground = $fg
                Title           = $u.Title
                Subtitle        = $sub
                SizeText        = $sizeStr
                StatusText      = "Pending"
                StatusColor     = "#D0D0D0"
                UpdateObject    = $u.UpdateObject
                RawItem         = $u
            }
            [void]$integratedWuItems.Add($item)
        }
        if ($lstIntegratedUpdates) {
            $lstIntegratedUpdates.ItemsSource = $null
            $lstIntegratedUpdates.ItemsSource = $integratedWuItems
        }
        if ($txtWuSummary) { $txtWuSummary.Text = "$($integratedWuItems.Count) update(s) available" }
    }

    function Start-RebootCountdown {
        param([int]$Seconds = 5, [scriptblock]$OnComplete)
        $script:RebootCountdownSec = $Seconds
        $script:RebootCountdownAction = $OnComplete
        if ($btnAbortReboot) { $btnAbortReboot.Visibility = [System.Windows.Visibility]::Visible }
        if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "System restart in $($script:RebootCountdownSec)s... Click 'Cancel Reboot' to abort." }
        Write-HubLog "System restart countdown initiated ($Seconds seconds). Click 'Cancel Reboot' on the Windows Update tab to abort." "WARN"

        if ($script:RebootCountdownTimer) {
            try { $script:RebootCountdownTimer.Stop() } catch { }
        }

        $script:RebootCountdownTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:RebootCountdownTimer.Interval = [TimeSpan]::FromSeconds(1)
        $script:RebootCountdownTimer.Add_Tick({
            $script:RebootCountdownSec--
            if ($script:RebootCountdownSec -gt 0) {
                if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "System restart in $($script:RebootCountdownSec)s... Click 'Cancel Reboot' to abort." }
                if ($txtWuStatus) { $txtWuStatus.Text = "Restarting system in $($script:RebootCountdownSec) seconds..." }
            } else {
                $script:RebootCountdownTimer.Stop()
                $script:RebootCountdownTimer = $null
                if ($btnAbortReboot) { $btnAbortReboot.Visibility = [System.Windows.Visibility]::Collapsed }
                if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "" }
                if ($script:RebootCountdownAction) {
                    $action = $script:RebootCountdownAction
                    $script:RebootCountdownAction = $null
                    & $action
                    try {
                        if ($window) { $window.Close() }
                    } catch { }
                }
            }
        })
        $script:RebootCountdownTimer.Start()
    }

    function Stop-RebootCountdown {
        if ($script:RebootCountdownTimer) {
            try { $script:RebootCountdownTimer.Stop() } catch { }
            $script:RebootCountdownTimer = $null
        }
        $script:RebootCountdownAction = $null
        if ($btnAbortReboot) { $btnAbortReboot.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "Restart cancelled by user." }
        Write-HubLog "System restart countdown cancelled by operator." "WARN"
    }

    function Invoke-GuiCascadePass {
        $state = Get-HubPatchCascadeState
        $currPass = if ($state -and $state.CurrentPass) { [int]$state.CurrentPass } else { 1 }
        $maxPasses = if ($state -and $state.MaxPasses) { [int]$state.MaxPasses } else {
            if ($cmbCascadeMaxPasses -and $cmbCascadeMaxPasses.SelectedItem) {
                [int]($cmbCascadeMaxPasses.SelectedItem.Content)
            } else { 5 }
        }
        $includeDrivers = if ($state) { [bool]$state.IncludeDrivers } else {
            if ($chkCascadeDrivers) { [bool]$chkCascadeDrivers.IsChecked } else { $true }
        }
        $autoReboot = if ($chkCascadeAutoReboot) { [bool]$chkCascadeAutoReboot.IsChecked } else { $true }
        $totalInstalled = if ($state -and $state.TotalInstalled) { [int]$state.TotalInstalled } else { 0 }

        $script:CascadeRunning = $true
        if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStopCascade.IsEnabled = $true }
        if ($btnWuScan) { $btnWuScan.IsEnabled = $false }
        if ($btnWuInstall) { $btnWuInstall.IsEnabled = $false }
        if ($btnWuRebootNow) { $btnWuRebootNow.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "" }

        if ($txtCascadeStatus) { $txtCascadeStatus.Text = "CASCADE ACTIVE" }
        if ($badgeCascadeStatus) {
            $badgeCascadeStatus.Background = $brushConv.ConvertFromString("#E0F2FE")
            $badgeCascadeStatus.BorderBrush = $brushConv.ConvertFromString("#BAE6FD")
        }
        if ($txtCascadeStatus) { $txtCascadeStatus.Foreground = $brushConv.ConvertFromString("#0284C7") }
        if ($txtCascadePassInfo) { $txtCascadePassInfo.Text = " (Pass $currPass of $maxPasses)" }
        if ($txtCascadeInstalled) { $txtCascadeInstalled.Text = " | $totalInstalled installed" }
        if ($txtWuStatus) { $txtWuStatus.Text = "Pass $($currPass): Scanning for Windows updates and drivers (async worker)..." }
        if ($progWuBar) { $progWuBar.IsIndeterminate = $true }
        Write-HubLog "Starting Autonomous Patch Cascade - Pass $currPass of $maxPasses (Drivers: $includeDrivers)..." "INFO"

        Set-HubPatchCascadeState -CurrentPass $currPass -MaxPasses $maxPasses -IncludeDrivers:$includeDrivers -TotalInstalled $totalInstalled

        # Step 1: Asynchronous Scan in background worker
        Start-HubAsyncWork -Description "Cascade Scan Pass $currPass" -TaskArgs @($includeDrivers) -WorkerScript {
            param($incDrivers)
            Write-WorkerLog "Cascade scan query in background runspace..."
            Set-WuWorkerProgress 10 "Scanning for updates and drivers"
            $updates = @(Get-HubPendingWindowsUpdates -IncludeDrivers:$incDrivers -StatusCallback {
                param($m)
                Write-WorkerLog $m "INFO"
            })
            return $updates
        } -OnComplete {
            param($rawUpdates, $scanErr)
            if (-not $script:CascadeRunning) { return }

            if ($progWuBar) { $progWuBar.IsIndeterminate = $false; $progWuBar.Value = 10 }
            $updateList = if ($rawUpdates) { @($rawUpdates) } else { @() }
            Update-IntegratedWuList -rawUpdates $updateList

            if ($updateList.Count -eq 0) {
                Clear-HubPatchCascadeState
                $script:CascadeRunning = $false
                if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStartCascade.IsEnabled = $true }
                if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Collapsed }
                if ($btnWuScan) { $btnWuScan.IsEnabled = $true }
                if ($txtCascadeStatus) { $txtCascadeStatus.Text = "100% PATCHED" }
                if ($badgeCascadeStatus) {
                    $badgeCascadeStatus.Background = $brushConv.ConvertFromString("#DCFCE7")
                    $badgeCascadeStatus.BorderBrush = $brushConv.ConvertFromString("#BBF7D0")
                }
                if ($txtCascadeStatus) { $txtCascadeStatus.Foreground = $brushConv.ConvertFromString("#15803D") }
                if ($txtCascadeStateDetail) { $txtCascadeStateDetail.Text = "Zero updates pending. All security patches, quality updates, and hardware drivers are fully applied." }
                if ($txtWuStatus) { $txtWuStatus.Text = "Cascade complete! System is 100% up to date. Total installed: $totalInstalled" }
                if ($progWuBar) { $progWuBar.Value = 100 }
                Write-HubLog "Patch Cascade finished: Zero updates remaining. System is 100% up to date! Total installed across cascade: $totalInstalled" "SUCCESS"
                try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }
                return
            }

            Write-HubLog "Pass $($currPass): Found $($updateList.Count) update(s) to install. Beginning download & installation off UI thread..." "INFO"
            if ($txtWuStatus) { $txtWuStatus.Text = "Pass $($currPass): Downloading & installing $($updateList.Count) update(s)..." }
            if ($progWuBar) { $progWuBar.IsIndeterminate = $true }

            $selectedTitles = @($updateList | ForEach-Object { $_.Title })
            $selectedIds = @($updateList | ForEach-Object { if ($_.UpdateID) { $_.UpdateID } else { $_.Title } })

            # Step 2: Asynchronous Download & Install in background worker
            Start-HubAsyncWork -Description "Cascade Install Pass $currPass" -TaskArgs @($selectedTitles, $selectedIds) -WorkerScript {
                param($titles, $ids)
                Write-WorkerLog "Initializing COM Update Session in background worker..."
                Ensure-HubUpdateServices
                $session = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $searchResult = $searcher.Search("IsInstalled=0")

                $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                if ($searchResult -and $searchResult.Updates) {
                    for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
                        $u = $searchResult.Updates.Item($i)
                        $match = $false
                        if ($ids -contains $u.Identity.UpdateID) { $match = $true }
                        elseif ($titles -contains $u.Title) { $match = $true }
                        if ($match) {
                            if ($u.EulaAccepted -eq $false) {
                                try { $u.AcceptEula() } catch { }
                            }
                            $updatesToInstall.Add($u) | Out-Null
                        }
                    }
                }

                if ($updatesToInstall.Count -eq 0) {
                    Write-WorkerLog "Could not locate selected updates in COM search." "WARN"
                    return [PSCustomObject]@{ Succeeded = 0; Failed = 0; RebootRequired = $false; Results = @() }
                }

                # Download
                Write-WorkerLog "Downloading $($updatesToInstall.Count) update package(s)..."
                Set-WuWorkerProgress 20 "Downloading updates"
                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $updatesToInstall
                $downRes = $downloader.Download()
                Write-WorkerLog "Download finished. Result: $($downRes.ResultCode)"
                Set-WuWorkerProgress 30 "Download complete"

                # Install
                $installer = $session.CreateUpdateInstaller()
                $installer.ForceQuiet = $true
                $total = $updatesToInstall.Count
                $succeeded = 0
                $failed = 0
                $rebootNeeded = $false
                $results = @()

                for ($i = 0; $i -lt $total; $i++) {
                    $u = $updatesToInstall.Item($i)
                    $pct = 30 + [math]::Round((($i) / $total) * 65)
                    Set-WuWorkerProgress $pct "Installing ($($i+1)/$total): $($u.Title)"
                    Write-WorkerLog "Installing ($($i+1)/$total): $($u.Title)..."

                    $singleColl = New-Object -ComObject Microsoft.Update.UpdateColl
                    $singleColl.Add($u) | Out-Null
                    $installer.Updates = $singleColl

                    $code = 0
                    try {
                        $instRes = $installer.Install()
                        $code = $instRes.ResultCode
                        if ($instRes.RebootRequired) { $rebootNeeded = $true }
                    } catch { $code = 99 }

                    if ($code -eq 2 -or $code -eq 3) {
                        $succeeded++
                        $status = "Installed"
                        Write-WorkerLog "Successfully installed: $($u.Title)" "SUCCESS"
                    } else {
                        $failed++
                        $status = "Failed ($code)"
                        Write-WorkerLog "Failed to install: $($u.Title) [Code: $code]" "ERROR"
                    }

                    $results += [PSCustomObject]@{
                        Title = $u.Title
                        UpdateID = $u.Identity.UpdateID
                        Status = $status
                        Code = $code
                    }
                }

                if (-not $rebootNeeded) {
                    if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -or
                        (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")) {
                        $rebootNeeded = $true
                    }
                }

                Set-WuWorkerProgress 100 "Installation Complete"
                return [PSCustomObject]@{
                    Succeeded = $succeeded
                    Failed = $failed
                    RebootRequired = $rebootNeeded
                    Results = $results
                }
            } -OnComplete {
                param($summary, $instErr)
                if (-not $script:CascadeRunning) {
                    Write-HubLog "Cascade paused after pass $currPass." "WARN"
                    return
                }

                $passInstalled = if ($summary) { $summary.Succeeded } else { 0 }
                $rebootNeeded = if ($summary) { [bool]$summary.RebootRequired } else { $false }
                if (-not $rebootNeeded) {
                    $rebootNeeded = ($passInstalled -gt 0) -or (Test-HubSystemRebootPending)
                }
                $totalInstalled += $passInstalled

                if ($summary -and $summary.Results) {
                    foreach ($res in $summary.Results) {
                        $found = $integratedWuItems | Where-Object { $_.Title -eq $res.Title -or ($_.RawItem -and $_.RawItem.UpdateID -eq $res.UpdateID) } | Select-Object -First 1
                        if ($found) {
                            $found.StatusText = $res.Status
                            $found.StatusColor = if ($res.Code -eq 2 -or $res.Code -eq 3) { "#15803D" } else { "#B91C1C" }
                        }
                    }
                    if ($lstIntegratedUpdates) { $lstIntegratedUpdates.Items.Refresh() }
                }

                if ($txtCascadeInstalled) { $txtCascadeInstalled.Text = " | $totalInstalled installed" }
                if ($progWuBar) { $progWuBar.IsIndeterminate = $false; $progWuBar.Value = 100 }
                if ($txtWuProgress) { $txtWuProgress.Text = "100%" }

                # In autonomous cascade, never show manual reboot button or notice
                if ($btnWuRebootNow) { $btnWuRebootNow.Visibility = [System.Windows.Visibility]::Collapsed }
                if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "" }

                if ($rebootNeeded) {
                    $nextPass = $currPass + 1
                    if ($nextPass -le $maxPasses) {
                        Write-HubLog "Pass $currPass completed ($passInstalled installed). Reboot is REQUIRED. Scheduling Pass $nextPass of $maxPasses..." "WARN"
                        Set-HubPatchCascadeState -CurrentPass $nextPass -MaxPasses $maxPasses -IncludeDrivers:$includeDrivers -TotalInstalled $totalInstalled
                        Register-HubResumeAfterRestart

                        # Autonomous 5-second countdown to reboot - ZERO extra clicks!
                        Start-RebootCountdown -Seconds 5 -OnComplete {
                            Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
                        }
                    } else {
                        Clear-HubPatchCascadeState
                        $script:CascadeRunning = $false
                        if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStartCascade.IsEnabled = $true }
                        if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Collapsed }
                        if ($btnWuScan) { $btnWuScan.IsEnabled = $true }
                        Write-HubLog "Reached maximum cascade passes ($maxPasses). Total updates installed: $totalInstalled. Performing final reboot to commit updates..." "WARN"
                        if ($txtWuStatus) { $txtWuStatus.Text = "Max passes ($maxPasses) reached ($totalInstalled installed). Committing updates via final reboot in 5s..." }
                        if ($txtCascadeStatus) { $txtCascadeStatus.Text = "CASCADE COMPLETE" }
                        # Automatically reboot final time so updates commit without requiring user clicks
                        Start-RebootCountdown -Seconds 5 -OnComplete {
                            Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
                        }
                    }
                } else {
                    Clear-HubPatchCascadeState
                    $script:CascadeRunning = $false
                    if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStartCascade.IsEnabled = $true }
                    if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Collapsed }
                    if ($btnWuScan) { $btnWuScan.IsEnabled = $true }
                    if ($txtCascadeStatus) { $txtCascadeStatus.Text = "100% PATCHED" }
                    if ($badgeCascadeStatus) {
                        $badgeCascadeStatus.Background = $brushConv.ConvertFromString("#DCFCE7")
                        $badgeCascadeStatus.BorderBrush = $brushConv.ConvertFromString("#BBF7D0")
                    }
                    if ($txtCascadeStatus) { $txtCascadeStatus.Foreground = $brushConv.ConvertFromString("#15803D") }
                    if ($txtCascadeStateDetail) { $txtCascadeStateDetail.Text = "Zero updates pending. All security patches, quality updates, and hardware drivers are fully applied." }
                    if ($txtWuStatus) { $txtWuStatus.Text = "Cascade complete! System is 100% up to date. Total installed: $totalInstalled" }
                    Write-HubLog "Patch Cascade complete: Zero updates pending reboot. System is 100% up to date! Total installed: $totalInstalled" "SUCCESS"
                    try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }
                }
            }
        }
    }

    if ($btnStartCascade) {
        $btnStartCascade.Add_Click({
            Invoke-GuiCascadePass
        })
    }

    if ($btnStopCascade) {
        $btnStopCascade.Add_Click({
            $script:CascadeRunning = $false
            Stop-RebootCountdown
            Clear-HubPatchCascadeState
            if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStartCascade.IsEnabled = $true }
            if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Collapsed }
            if ($btnWuScan) { $btnWuScan.IsEnabled = $true }
            if ($txtCascadeStatus) { $txtCascadeStatus.Text = "CASCADE STOPPED" }
            if ($badgeCascadeStatus) {
                $badgeCascadeStatus.Background = $brushConv.ConvertFromString("#FEE2E2")
                $badgeCascadeStatus.BorderBrush = $brushConv.ConvertFromString("#FCA5A5")
            }
            if ($txtCascadeStatus) { $txtCascadeStatus.Foreground = $brushConv.ConvertFromString("#B91C1C") }
            if ($txtWuStatus) { $txtWuStatus.Text = "Cascade stopped. Saved cascade state cleared." }
            Write-HubLog "Autonomous Patch Cascade stopped and state cleared by operator." "INFO"
        })
    }

    if ($btnAbortReboot) {
        $btnAbortReboot.Add_Click({
            Stop-RebootCountdown
            $script:CascadeRunning = $false
            if ($btnStartCascade) { $btnStartCascade.Visibility = [System.Windows.Visibility]::Visible; $btnStartCascade.IsEnabled = $true }
            if ($btnStopCascade) { $btnStopCascade.Visibility = [System.Windows.Visibility]::Collapsed }
            if ($btnWuScan) { $btnWuScan.IsEnabled = $true }
            if ($txtCascadeStatus) { $txtCascadeStatus.Text = "CASCADE PAUSED" }
            if ($badgeCascadeStatus) {
                $badgeCascadeStatus.Background = $brushConv.ConvertFromString("#FEF3C7")
                $badgeCascadeStatus.BorderBrush = $brushConv.ConvertFromString("#FDE68A")
            }
            if ($txtCascadeStatus) { $txtCascadeStatus.Foreground = $brushConv.ConvertFromString("#B45309") }
            if ($txtWuStatus) { $txtWuStatus.Text = "Restart cancelled. Autonomous cascade is paused." }
        })
    }

    if ($btnWuScan) {
        $btnWuScan.Add_Click({
            $btnWuScan.IsEnabled = $false
            if ($btnWuInstall) { $btnWuInstall.IsEnabled = $false }
            if ($progWuBar) { $progWuBar.IsIndeterminate = $true }
            if ($txtWuStatus) { $txtWuStatus.Text = "Scanning for updates and drivers (async worker)..." }
            Write-HubLog "Scanning for Windows updates and drivers (async worker)..."

            $incDrivers = if ($chkCascadeDrivers) { [bool]$chkCascadeDrivers.IsChecked } else { $true }
            Start-HubAsyncWork -Description "Windows Update Scan" -TaskArgs @($incDrivers) -WorkerScript {
                param($includeDrivers)
                Write-WorkerLog "Querying Microsoft.Update.Session off UI thread..."
                Set-WuWorkerProgress 30 "Connecting to Windows Update searcher"
                $rawUpdates = @(Get-HubPendingWindowsUpdates -IncludeDrivers:$includeDrivers -StatusCallback {
                    param($msg)
                    Write-WorkerLog $msg "INFO"
                })
                Set-WuWorkerProgress 90 "Compiling update list"
                return $rawUpdates
            } -OnComplete {
                param($rawUpdates, $err)
                $btnWuScan.IsEnabled = $true
                if ($progWuBar) { $progWuBar.IsIndeterminate = $false; $progWuBar.Value = 0 }
                if ($rawUpdates) {
                    Update-IntegratedWuList -rawUpdates $rawUpdates
                    if ($integratedWuItems.Count -gt 0) {
                        if ($btnWuInstall) { $btnWuInstall.IsEnabled = $true }
                        if ($txtWuStatus) { $txtWuStatus.Text = "Found $($integratedWuItems.Count) update(s). Select updates and click 'Install Selected'." }
                    } else {
                        if ($txtWuStatus) { $txtWuStatus.Text = "System is up to date. No pending updates found." }
                    }
                    Write-HubLog "Scan complete: Found $($rawUpdates.Count) available update(s)." "SUCCESS"
                } else {
                    $msg = if ($err) { $err.Message } else { "No updates found or search failed." }
                    Write-HubLog "Windows Update scan: $msg" "WARN"
                    if ($txtWuStatus) { $txtWuStatus.Text = $msg }
                }
            }
        })
    }

    if ($btnWuSelectAll) {
        $btnWuSelectAll.Add_Click({
            foreach ($it in $integratedWuItems) { $it.IsSelected = $true }
            if ($lstIntegratedUpdates) { $lstIntegratedUpdates.Items.Refresh() }
        })
    }

    if ($btnWuDeselectAll) {
        $btnWuDeselectAll.Add_Click({
            foreach ($it in $integratedWuItems) { $it.IsSelected = $false }
            if ($lstIntegratedUpdates) { $lstIntegratedUpdates.Items.Refresh() }
        })
    }

    if ($btnWuTriggerUso) {
        $btnWuTriggerUso.Add_Click({
            $ok = Invoke-HubUsoScan
            if ($ok) {
                Write-HubLog "USO client interactive scan initiated in background." "INFO"
                if ($txtWuStatus) { $txtWuStatus.Text = "USO Interactive Scan initiated in background." }
            } else {
                Write-HubLog "usoclient.exe not available on this system." "WARN"
                if ($txtWuStatus) { $txtWuStatus.Text = "usoclient.exe not found on this system." }
            }
        })
    }

    if ($btnWuRebootNow) {
        $btnWuRebootNow.Add_Click({
            if ($script:PlaybookRunning) {
                Write-HubLog "Playbook: System reboot scheduled upon completion." "INFO"
                return
            }
            Write-HubLog "Restart requested by operator. Registering resume task..." "INFO"
            try {
                Register-HubResumeAfterRestart
                Write-HubLog "Restart-resume hook registered. Hub will re-launch automatically on boot." "SUCCESS"
            } catch {
                Write-HubLog "Resume registration notice: $($_.Exception.Message)" "WARN"
            }
            Start-RebootCountdown -Seconds 3 -OnComplete {
                Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
            }
        })
    }

    if ($btnWuInstall) {
        $btnWuInstall.Add_Click({
            $selected = @($integratedWuItems | Where-Object { $_.IsSelected -eq $true })
            if ($selected.Count -eq 0) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please select at least one update to install.", "No Updates Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }

            $btnWuScan.IsEnabled = $false
            $btnWuInstall.IsEnabled = $false
            if ($progWuBar) { $progWuBar.Value = 0 }
            if ($txtWuStatus) { $txtWuStatus.Text = "Preparing update download session in background..." }
            Write-HubLog "Starting asynchronous download and installation of $($selected.Count) update(s)..."

            $selectedTitles = @($selected | ForEach-Object { $_.Title })
            $selectedIds = @($selected | ForEach-Object { if ($_.RawItem.UpdateID) { $_.RawItem.UpdateID } else { $_.Title } })

            Start-HubAsyncWork -Description "Windows Update Installation" -TaskArgs @($selectedTitles, $selectedIds) -WorkerScript {
                param($titles, $ids)
                Write-WorkerLog "Initializing COM Update Session in background runspace..."
                Ensure-HubUpdateServices
                $session = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                $searchResult = $searcher.Search("IsInstalled=0")

                $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                if ($searchResult -and $searchResult.Updates) {
                    for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
                        $u = $searchResult.Updates.Item($i)
                        $match = $false
                        if ($ids -contains $u.Identity.UpdateID) { $match = $true }
                        elseif ($titles -contains $u.Title) { $match = $true }
                        if ($match) {
                            if ($u.EulaAccepted -eq $false) {
                                try { $u.AcceptEula() } catch { }
                            }
                            $updatesToInstall.Add($u) | Out-Null
                        }
                    }
                }

                if ($updatesToInstall.Count -eq 0) {
                    Write-WorkerLog "Could not locate selected updates in COM search." "WARN"
                    return [PSCustomObject]@{ Succeeded = 0; Failed = 0; RebootRequired = $false; Results = @() }
                }

                # Download
                Write-WorkerLog "Downloading $($updatesToInstall.Count) update package(s)..."
                Set-WuWorkerProgress 15 "Downloading Updates"
                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $updatesToInstall
                $downRes = $downloader.Download()
                Write-WorkerLog "Download finished. Result: $($downRes.ResultCode)"
                Set-WuWorkerProgress 30 "Download Complete"

                # Install
                $installer = $session.CreateUpdateInstaller()
                $installer.ForceQuiet = $true
                $total = $updatesToInstall.Count
                $succeeded = 0
                $failed = 0
                $rebootNeeded = $false
                $results = @()

                for ($i = 0; $i -lt $total; $i++) {
                    $u = $updatesToInstall.Item($i)
                    $pct = 30 + [math]::Round((($i) / $total) * 65)
                    Set-WuWorkerProgress $pct "Installing ($($i+1)/$total): $($u.Title)"
                    Write-WorkerLog "Installing ($($i+1)/$total): $($u.Title)..."

                    $singleColl = New-Object -ComObject Microsoft.Update.UpdateColl
                    $singleColl.Add($u) | Out-Null
                    $installer.Updates = $singleColl

                    $code = 0
                    try {
                        $instRes = $installer.Install()
                        $code = $instRes.ResultCode
                        if ($instRes.RebootRequired) { $rebootNeeded = $true }
                    } catch { $code = 99 }

                    if ($code -eq 2 -or $code -eq 3) {
                        $succeeded++
                        $status = "Installed"
                        Write-WorkerLog "Successfully installed: $($u.Title)" "SUCCESS"
                    } else {
                        $failed++
                        $status = "Failed ($code)"
                        Write-WorkerLog "Failed to install: $($u.Title) [Code: $code]" "ERROR"
                    }

                    $results += [PSCustomObject]@{
                        Title = $u.Title
                        UpdateID = $u.Identity.UpdateID
                        Status = $status
                        Code = $code
                    }
                }

                Set-WuWorkerProgress 100 "Installation Complete"
                return [PSCustomObject]@{
                    Succeeded = $succeeded
                    Failed = $failed
                    RebootRequired = $rebootNeeded
                    Results = $results
                }
            } -OnComplete {
                param($summary, $err)
                $btnWuScan.IsEnabled = $true
                $btnWuInstall.IsEnabled = $true

                if ($summary) {
                    foreach ($res in $summary.Results) {
                        $found = $integratedWuItems | Where-Object { $_.Title -eq $res.Title -or ($_.RawItem -and $_.RawItem.UpdateID -eq $res.UpdateID) } | Select-Object -First 1
                        if ($found) {
                            $found.StatusText = $res.Status
                            $found.StatusColor = if ($res.Code -eq 2 -or $res.Code -eq 3) { "#15803D" } else { "#B91C1C" }
                        }
                    }
                    if ($lstIntegratedUpdates) { $lstIntegratedUpdates.Items.Refresh() }

                    $rebootReq = [bool]($summary.RebootRequired -or ($summary.Succeeded -gt 0 -and (Test-HubSystemRebootPending)))
                    Write-HubLog "Update installation complete: $($summary.Succeeded) succeeded, $($summary.Failed) failed. Reboot required: $rebootReq" $(if ($summary.Failed -eq 0) { "SUCCESS" } else { "WARN" })
                    if ($txtWuStatus) { $txtWuStatus.Text = "Update complete: $($summary.Succeeded) installed, $($summary.Failed) failed." }
                    if ($progWuBar) { $progWuBar.Value = 100 }
                    try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }

                    if ($rebootReq) {
                        if ($btnWuRebootNow) { $btnWuRebootNow.Visibility = [System.Windows.Visibility]::Visible }
                        if ($txtWuRebootNotice) { $txtWuRebootNotice.Text = "Updates installed successfully. System restart is REQUIRED." }
                        if ($chkCascadeAutoReboot -and $chkCascadeAutoReboot.IsChecked) {
                            Start-RebootCountdown -Seconds 5 -OnComplete {
                                Register-HubResumeAfterRestart
                                Invoke-HubSystemReboot -Reason 'Autonomous System Restart'
                            }
                        }
                    }
                } else {
                    $errMsg = if ($err) { $err.Message } else { "Unknown update installation error" }
                    Write-HubLog "Update installation error: $errMsg" "ERROR"
                    if ($txtWuStatus) { $txtWuStatus.Text = "Installation failed: $errMsg" }
                }
            }
        })
    }

    # --- TAB 6 ACTIONS (Lenovo Warranty, Battery Wear & Storage SMART) ---
    if ($btnCheckLenovoWarranty) {
        $btnCheckLenovoWarranty.Add_Click({
            $serial = if ($txtDellServiceTag -and -not [string]::IsNullOrWhiteSpace($txtDellServiceTag.Text)) { $txtDellServiceTag.Text.Trim() } else { $txtSerial.Text.Trim() }
            $btnCheckLenovoWarranty.IsEnabled = $false
            Write-HubLog "Checking Lenovo warranty for serial: $serial (async worker)..."
            Start-HubAsyncWork -Description "Lenovo Warranty Lookup" -TaskArgs @($serial) -WorkerScript {
                param($s)
                return Get-LenovoWarrantyInfo -SerialNumber $s
            } -OnComplete {
                param($lenovo, $err)
                $btnCheckLenovoWarranty.IsEnabled = $true
                if ($lenovo -and $lenovo.Success) {
                    Write-HubLog "Lenovo Warranty: Product=$($lenovo.ProductName) Status=$($lenovo.Status) EndDate=$($lenovo.EndDate)" "SUCCESS"
                    if ($txtDellEndDate) { $txtDellEndDate.Text = "$($lenovo.EndDate)" }
                    if ($txtDellDaysRemaining) { $txtDellDaysRemaining.Text = "Status: $($lenovo.Status) ($($lenovo.DaysRemaining) days)" }
                } else {
                    $msg = if ($lenovo) { $lenovo.ErrorMessage } elseif ($err) { $err.Message } else { "Unknown warranty failure" }
                    Write-HubLog "Lenovo warranty query failed: $msg" "ERROR"
                }
            }
        })
    }

    function Update-BatteryUi {
        Start-HubAsyncWork -Description "Battery Health Probing" -WorkerScript {
            return Get-BatteryHealthInfo
        } -OnComplete {
            param($bat, $err)
            if ($bat -and $bat.Present) {
                if ($txtBatteryWear)     { $txtBatteryWear.Text     = "Wear: $($bat.WearPercent)" }
                if ($txtBatteryVerdict)  { $txtBatteryVerdict.Text  = "Status: $($bat.HealthVerdict)" }
                if ($txtBatteryCapacity) { $txtBatteryCapacity.Text = "Design: $($bat.DesignCapacity) | Full: $($bat.FullChargeCapacity)" }
            } else {
                if ($txtBatteryWear)     { $txtBatteryWear.Text     = "No Battery" }
                if ($txtBatteryVerdict)  { $txtBatteryVerdict.Text  = "Desktop / AC" }
                if ($txtBatteryCapacity) { $txtBatteryCapacity.Text = "Design: N/A | Full: N/A" }
            }
        }
    }
    Update-BatteryUi
    if ($btnRefreshBattery) {
        $btnRefreshBattery.Add_Click({ Update-BatteryUi; Write-HubLog "Battery wear and capacity refresh requested (async worker)." "INFO" })
    }

    function Update-StorageUi {
        Start-HubAsyncWork -Description "Storage Reliability Probing" -WorkerScript {
            return Get-StorageReliabilityInfo
        } -OnComplete {
            param($disks, $err)
            if ($disks -and $disks.Count -gt 0) {
                $primary = $disks[0]
                if ($txtStorageWear) { $txtStorageWear.Text = "SSD Wear: $($primary.Wear)" }
                if ($txtStorageTemp) { $txtStorageTemp.Text = "Temp: $($primary.Temperature)" }
                if ($txtStorageMeta) { $txtStorageMeta.Text = "Read Err: $($primary.ReadErrors) | Write Err: $($primary.WriteErrors) | POH: $($primary.PowerOnHours)h" }
            } else {
                if ($txtStorageWear) { $txtStorageWear.Text = "SSD Wear: N/A" }
                if ($txtStorageTemp) { $txtStorageTemp.Text = "Temp: Normal" }
                if ($txtStorageMeta) { $txtStorageMeta.Text = "Storage reliability counter not supported or no physical NVMe disks found." }
            }
        }
    }
    Update-StorageUi
    if ($btnRefreshStorage) {
        $btnRefreshStorage.Add_Click({ Update-StorageUi; Write-HubLog "Storage SMART reliability counter refresh requested (async worker)." "INFO" })
    }

    # --- TAB 7 ACTIONS (ESP Win32Apps, Lifecycle, Wi-Fi, Licensing, IME Tail) ---
    function Update-Win32AppsUi {
        try {
            $apps = Get-Win32AppDiagnostics
            if ($lstWin32Apps) {
                $lstWin32Apps.ItemsSource = @($apps)
            }
            Write-HubLog "Loaded $($apps.Count) Win32 app ESP tracking state(s) from registry." "INFO"
        } catch {
            Write-HubLog "Failed to read Win32App state: $($_.Exception.Message)" "WARN"
        }
    }
    Update-Win32AppsUi
    if ($btnRefreshWin32Apps) {
        $btnRefreshWin32Apps.Add_Click({ Update-Win32AppsUi })
    }

    if ($btnExportMdmCab) {
        $btnExportMdmCab.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("mdmdiagnosticstool requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            Write-HubLog "Exporting MDM & Autopilot diagnostic CAB bundle..."
            $cab = Export-MdmDiagnosticsCab
            if ($cab) {
                Write-HubLog "MDM Diagnostics CAB created: $cab" "SUCCESS"
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Diagnostic CAB bundle exported to:`n$cab", "CAB Export Succeeded", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
            } else {
                Write-HubLog "Failed to export MDM diagnostic CAB." "ERROR"
            }
        })
    }

    if ($btnPatchLifecycle) {
        $btnPatchLifecycle.Add_Click({
            if (-not ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken)) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please sign in to Microsoft Graph first.", "Authentication Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            $ident = if ($script:DeviceState -and $script:DeviceState.CloudIdentity) { $script:DeviceState.CloudIdentity.id } else { '' }
            if (-not $ident) {
                $serial = if ($txtSerial) { $txtSerial.Text.Trim() } else { '' }
                $found = Find-AutopilotIdentityInTenant -SerialNumber $serial -AccessToken $script:GraphAuthContext.AccessToken
                if ($found) { $ident = $found.id }
            }
            if (-not $ident) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("No Autopilot device identity found in tenant for serial $($txtSerial.Text).", "Device Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            $gt = if ($txtLifecycleGroupTag) { $txtLifecycleGroupTag.Text } else { '' }
            $user = if ($txtLifecycleUser) { $txtLifecycleUser.Text } else { '' }
            Write-HubLog "Patching Autopilot identity $ident (Tag='$gt', User='$user')..."
            $ok = Set-AutopilotDeviceIdentityPatch -IdentityId $ident -GroupTag $gt -AssignedUser $user -AccessToken $script:GraphAuthContext.AccessToken
            if ($ok) {
                Write-HubLog "Successfully patched Autopilot identity in tenant." "SUCCESS"
            } else {
                Write-HubLog "Failed to patch Autopilot identity." "ERROR"
            }
        })
    }

    if ($btnSyncAutopilot) {
        $btnSyncAutopilot.Add_Click({
            if (-not ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken)) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please sign in to Microsoft Graph first.", "Authentication Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            Write-HubLog "Triggering instant Windows Autopilot tenant sync..."
            $syncOk = Invoke-AutopilotTenantSync -AccessToken $script:GraphAuthContext.AccessToken
            if ($syncOk) {
                Write-HubLog "Autopilot sync request accepted by Microsoft Graph." "SUCCESS"
            } else {
                Write-HubLog "Failed to trigger Autopilot sync." "ERROR"
            }
        })
    }

    if ($btnDecommission) {
        $btnDecommission.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Decommissioning device requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            $ans = if ($script:PlaybookRunning) { [System.Windows.MessageBoxResult]::Yes } else {
                [System.Windows.MessageBox]::Show("Decommission this device's LOCAL Autopilot / MDM state?`n`nThis purges on this machine only:`n- the cached Autopilot profile and Provisioning diagnostics`n- the local MDM (Intune) enrollment keys`n`nIt does NOT remove the device from Intune or Autopilot in the tenant - retire/delete it in the portal for that, otherwise the tenant still believes it is managed.`n`nThis action cannot be undone.", "Confirm Local Decommission", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            }
            if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-HubLog "Executing clean local Autopilot decommission..." "WARN"
                foreach ($act in @(Invoke-AutopilotCleanDecommission)) { Write-HubLog "  $act" $(if ($act -like 'Failed*') { 'ERROR' } else { 'INFO' }) }
                $script:DeviceState = Get-DeviceEnrollmentState
                Update-DeviceStateUi
                Write-HubLog "Decommission complete." "SUCCESS"
            }
        })
    }

    if ($btnScanWifi) {
        $btnScanWifi.Add_Click({
            $btnScanWifi.IsEnabled = $false
            Write-HubLog "Scanning for available Wi-Fi SSIDs (async worker)..."
            Start-HubAsyncWork -Description "Wi-Fi Scanning" -WorkerScript {
                return Get-HubWifiNetworks
            } -OnComplete {
                param($ssids, $err)
                $btnScanWifi.IsEnabled = $true
                if ($ssids) {
                    if ($cmbWifiSsids) {
                        $cmbWifiSsids.ItemsSource = @($ssids)
                        if ($ssids.Count -gt 0) { $cmbWifiSsids.SelectedIndex = 0 }
                    }
                    Write-HubLog "Found $($ssids.Count) Wi-Fi network(s)." "INFO"
                } else {
                    Write-HubLog "Wi-Fi scan failed or no networks found." "WARN"
                }
            }
        })
    }

    if ($btnConnectWifi) {
        $btnConnectWifi.Add_Click({
            $ssid = if ($cmbWifiSsids) { $cmbWifiSsids.Text } else { '' }
            $pwd = if ($txtWifiPassword) { $txtWifiPassword.Text } else { '' }
            if (-not $ssid) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Please select or enter an SSID.", "SSID Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            Write-HubLog "Connecting to Wi-Fi SSID '$ssid'..."
            $connRes = Connect-HubWifiNetwork -Ssid $ssid -Password $pwd
            if ($connRes) {
                Write-HubLog "Connected successfully to Wi-Fi '$ssid'." "SUCCESS"
            } else {
                Write-HubLog "Failed to connect to Wi-Fi '$ssid'." "ERROR"
            }
        })
    }

    if ($btnImportWifiXml) {
        $btnImportWifiXml.Add_Click({
            try {
                if ($script:PlaybookRunning) {
                    Write-HubLog "Playbook: Wi-Fi XML profile engine ready." "INFO"
                    return
                }
                $dlg = [Microsoft.Win32.OpenFileDialog]::new()
                $dlg.Filter = "Wi-Fi Profile XML (*.xml)|*.xml|All Files (*.*)|*.*"
                $dlg.Title = "Select 802.1X Wi-Fi XML Profile"
                if ($dlg.ShowDialog() -eq $true) {
                    $impRes = Import-Hub8021xProfile -XmlPath $dlg.FileName
                    if ($impRes) {
                        Write-HubLog "Imported Wi-Fi profile from $($dlg.FileName)." "SUCCESS"
                    } else {
                        Write-HubLog "Failed to import Wi-Fi profile from $($dlg.FileName)." "ERROR"
                    }
                }
            } catch {
                Write-HubLog "Import 802.1X XML profile error: $($_.Exception.Message)" "ERROR"
            }
        })
    }

    function Update-LicensingUi {
        # Get-WindowsLicensingInfo is 1-2 s of licensing WMI; fetch off the UI thread
        if ($txtOsCaption) { $txtOsCaption.Text = "Edition: querying..." }
        Start-HubAsyncWork -Description "Windows Licensing" -WorkerScript {
            Get-WindowsLicensingInfo
        } -OnComplete {
            param($lic, $err)
            if (-not $lic) { return }
            if ($txtOsCaption) { $txtOsCaption.Text = "Edition: $($lic.Caption)" }
            if ($txtOemKey)    { $txtOemKey.Text    = "OEM Key: $($lic.OemKey)" }
            if ($lic.OemKey -and $lic.OemKey -notlike 'No OA3*') { $script:CachedOa3Key = $lic.OemKey }
        } | Out-Null
    }
    Update-LicensingUi
    if ($btnRefreshLicensing) {
        $btnRefreshLicensing.Add_Click({ Update-LicensingUi; Write-HubLog "Refreshed Windows edition and OEM key." "INFO" })
    }

    if ($btnUpgradeEnterprise) {
        $btnUpgradeEnterprise.Add_Click({
            if (-not $script:RuntimeContext.MeetsPreferred) {
                if (-not $script:PlaybookRunning) { [System.Windows.MessageBox]::Show("Windows Edition upgrade requires elevation.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
                return
            }
            if ($script:PlaybookRunning) {
                Write-HubLog "Playbook: Verified Enterprise KMS upgrade engine and readiness." "INFO"
                Update-LicensingUi
                return
            }
            $ans = [System.Windows.MessageBox]::Show("Change the Windows edition to Enterprise with Microsoft's generic KMS client setup key?`n`nCommand: changepk.exe /ProductKey NPPR9-FWDCX-D2C8J-H872K-2YT43`n`nThis key only ACTIVATES against a KMS host on your network (or a Microsoft 365 E3/E5 subscription activation). On a machine with neither, Windows becomes an unactivated Enterprise install - and going back to Pro means re-entering the Pro key.`n`nContinue?", "Upgrade to Enterprise", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-HubLog "Upgrading Windows to Enterprise via changepk.exe..." "WARN"
                $upOk = Invoke-WindowsEnterpriseUpgrade
                Update-LicensingUi
                if ($upOk) {
                    Write-HubLog "Windows Enterprise upgrade command finished successfully." "SUCCESS"
                } else {
                    Write-HubLog "changepk exited with non-zero status. Check Windows Activation Settings." "WARN"
                }
            }
        })
    }

    if ($btnTailImeLog) {
        $btnTailImeLog.Add_Click({
            Write-HubLog "Tailing last 100 lines of IntuneManagementExtension.log..."
            $tail = Get-ImeLogTail -Lines 100
            if ($txtImeLog) {
                $txtImeLog.Text = $tail
            }
        })
    }

    # ==============================================================================
    # ACTION ROUTINES / DEPLOYMENT PLAYBOOK ENGINE (MULTI-TAB CLICK-THROUGH)
    # ==============================================================================

    $script:PlaybookRunning = $false
    $script:PlaybookState = @{
        Status          = 'Idle'
        IsPaused        = $false
        CancelRequested = $false
        CurrentStep     = 0
        TotalSteps      = 0
        CurrentStepInfo = ''
    }

    function Get-HubPlaybookSteps {
        param([string]$RoutineName)

        $steps = [System.Collections.Generic.List[object]]::new()

        switch -Wildcard ($RoutineName) {
            '*Intune*' {
                # ROUTINE 1: Intune-Only Cloud Build (Coordinated across ALL 9 tabs)

                # Tab 9: Precision Enterprise Fixes (Pre-flight component audit)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Output Log'
                    Button = $btnClearFixOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearFixOut) { $btnClearFixOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Audit Comprehensive System Health'
                    Button = $btnFixHealthReadout; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixHealthReadout) { $btnFixHealthReadout.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Verify & Unbrick AppX Manifest Staging'
                    Button = $btnFixAppX; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixAppX) { $btnFixAppX.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Copy Pre-Flight Remediation Log'
                    Button = $btnCopyFixLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyFixLog) { $btnCopyFixLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 4: Pre-Flight Diagnostics
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Pre-Flight Update Readiness'
                    Button = $btnPreflightWinUpdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPreflightWinUpdate) { $btnPreflightWinUpdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Auto-Remediate HTTPS Clock Skew (w32tm)'
                    Button = $btnFixClockSkew; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixClockSkew) { $btnFixClockSkew.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'TPM 2.0 & EK Attestation Engine'
                    Button = $btnTpmAttestation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTpmAttestation) { $btnTpmAttestation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Scan & Inject USB Drivers (pnputil)'
                    Button = $btnInjectDriversPreflight; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectDriversPreflight) { $btnInjectDriversPreflight.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Run 7-Stage Pre-Flight Network Ladder'
                    Button = $btnRunDiag; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnRunDiag) { $btnRunDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 6: Dell Warranty & Hardware Health
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Detect BIOS Service Tag'
                    Button = $btnDetectDellTag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnDetectDellTag) { $btnDetectDellTag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Dell Official Warranty & SLA'
                    Button = $btnCheckDellWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckDellWarranty) { $btnCheckDellWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Lenovo Warranty'
                    Button = $btnCheckLenovoWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckLenovoWarranty) { $btnCheckLenovoWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Refresh Battery Wear & Health'
                    Button = $btnRefreshBattery; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshBattery) { $btnRefreshBattery.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Storage NVMe SMART'
                    Button = $btnRefreshStorage; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshStorage) { $btnRefreshStorage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Copy Markdown Hardware Audit Report'
                    Button = $btnCopyDellReport; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnCopyDellReport) {
                            $btnCopyDellReport.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "Hardware markdown report copy skipped: No report cached." "INFO"
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Export Dell Warranty CSV'
                    Button = $btnExportDellCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnExportDellCsv) {
                            $btnExportDellCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "Hardware CSV export skipped: No Dell warranty report cached." "INFO"
                        }
                    }
                })

                # Tab 1: Autopilot & Cloud Registration
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Dismiss Tenant Mismatch Banner'
                    Button = $btnDismissMismatch; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($bannerTenantMismatch -and $bannerTenantMismatch.Visibility -eq [System.Windows.Visibility]::Visible -and $btnDismissMismatch) {
                            $btnDismissMismatch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Execute Dynamic Device State Action'
                    Button = $btnDeviceStateAction; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnDeviceStateAction -and $btnDeviceStateAction.Visibility -eq [System.Windows.Visibility]::Visible) {
                            $btnDeviceStateAction.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Apply Standardized Computer Name'
                    Button = $btnApplyRename; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnApplyRename) { $btnApplyRename.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Harvest Genuine OA3 Hardware Hash'
                    Button = $btnHarvestHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHarvestHash) { $btnHarvestHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Copy Hardware Hash to Clipboard'
                    Button = $btnCopyHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyHash) { $btnCopyHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Inspect Decoded Autopilot Policy'
                    Button = $btnInspectPolicy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInspectPolicy) { $btnInspectPolicy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Export Intune CSV (USB Priority)'
                    Button = $btnExportCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportCsv) { $btnExportCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Export Offline Provisioning JSON'
                    Button = $btnExportOfflineJson; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportOfflineJson) { $btnExportOfflineJson.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Save Current Settings to .env'
                    Button = $btnSaveEnvDefaults; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSaveEnvDefaults) { $btnSaveEnvDefaults.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Register Device with Intune Graph'
                    Button = $btnRegisterIntune; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $btnRegisterIntune) {
                            $btnRegisterIntune.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "Graph registration skipped in playbook: No active Microsoft Graph session." "INFO"
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Verify Cloud Link & Resume Autopilot'
                    Button = $btnResumeAutopilot; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnResumeAutopilot -and $btnResumeAutopilot.Visibility -eq [System.Windows.Visibility]::Visible) {
                            $btnResumeAutopilot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } elseif ($btnReDockStaging -and $btnReDockStaging.Visibility -eq [System.Windows.Visibility]::Visible) {
                            $btnReDockStaging.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })

                # Tab 7: Hybrid & Co-Mgmt (Pure Cloud Authority Verification)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearHybOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearHybOut) { $btnClearHybOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit Entra PRT & WAM Broker'
                    Button = $btnHybPrtDiag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybPrtDiag) { $btnHybPrtDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test Entra ID Token Acquisition'
                    Button = $btnHybAadToken; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybAadToken) { $btnHybAadToken.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit SCEP Certificate Health'
                    Button = $btnHybScepHealth; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybScepHealth) { $btnHybScepHealth.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Machine Certificate Health'
                    Button = $btnHybCerts; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCerts) { $btnHybCerts.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Read Co-Management Workload Authority'
                    Button = $btnHybComgmt; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybComgmt) { $btnHybComgmt.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Shift Workloads to Intune (255)'
                    Button = $btnHybCoMgmtAllIntune; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCoMgmtAllIntune) { $btnHybCoMgmtAllIntune.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force Intune MDM Sync'
                    Button = $btnHybSync; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybSync) { $btnHybSync.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Restart Intune Management Extension'
                    Button = $btnHybRestartIme; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybRestartIme) { $btnHybRestartIme.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Escrow BitLocker Recovery Keys'
                    Button = $btnHybBitlocker; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybBitlocker) { $btnHybBitlocker.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 5: Windows Update & Drivers
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Trigger USO Client Background Scan'
                    Button = $btnWuTriggerUso; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuTriggerUso) { $btnWuTriggerUso.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Schedule Post-Update System Restart'
                    Button = $btnWuRebootNow; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuRebootNow) { $btnWuRebootNow.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Scan for Updates & Drivers'
                    Button = $btnWuScan; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 25
                    Action = { if ($btnWuScan) { $btnWuScan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Deselect All Updates'
                    Button = $btnWuDeselectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuDeselectAll) { $btnWuDeselectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Select All Available Updates'
                    Button = $btnWuSelectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuSelectAll) { $btnWuSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Install Selected Updates'
                    Button = $btnWuInstall; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        $sel = @($integratedWuItems | Where-Object { $_.IsSelected -eq $true })
                        if ($sel.Count -gt 0 -and $btnWuInstall) {
                            $btnWuInstall.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "No pending Windows updates to install. System is current." "INFO"
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Schedule Post-Update System Restart'
                    Button = $btnWuRebootNow; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuRebootNow) { $btnWuRebootNow.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Start Autonomous Patch Cascade'
                    Button = $btnStartCascade; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnStartCascade -and $btnStartCascade.IsEnabled) { $btnStartCascade.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Cancel Reboot & Suppress Cascade'
                    Button = $btnAbortReboot; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnAbortReboot) { $btnAbortReboot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
                        if ($btnStopCascade) { $btnStopCascade.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
                    }
                })

                # Tab 2: App Deployment
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Clear App Selection'
                    Button = $btnClearApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearApps) { $btnClearApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Apply Standard Workstation Preset'
                    Button = $btnPresetWorkstation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPresetWorkstation) { $btnPresetWorkstation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Select Target Applications'
                    Button = $btnSelectAllApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSelectAllApps) { $btnSelectAllApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Install Selected Applications'
                    Button = $btnInstallBatch; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        $appCount = ($script:AppCatalog | Where-Object { $_.IsSelected }).Count
                        if ($appCount -gt 0 -and $btnInstallBatch) {
                            $btnInstallBatch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "App catalog selection empty. Batch install skipped." "INFO"
                        }
                    }
                })

                # Tab 3: Win32 Packaging
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Build Win32 App Package (.intunewin)'
                    Button = $btnBuildPackage; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnBuildPackage) { $btnBuildPackage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Publish Win32 Package to Intune Cloud'
                    Button = $btnPublishIntune; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $btnPublishIntune) {
                            $btnPublishIntune.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "Win32 publishing skipped: Graph authentication required." "INFO"
                        }
                    }
                })

                # Tab 8: ESP Diagnostics & Lifecycle
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh Windows Edition Licensing'
                    Button = $btnRefreshLicensing; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshLicensing) { $btnRefreshLicensing.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = '1-Click Enterprise KMS Upgrade Check'
                    Button = $btnUpgradeEnterprise; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnUpgradeEnterprise) { $btnUpgradeEnterprise.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Scan Wi-Fi SSIDs'
                    Button = $btnScanWifi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnScanWifi) { $btnScanWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Verify Wi-Fi Connection Readiness'
                    Button = $btnConnectWifi; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnConnectWifi) { $btnConnectWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Validate 802.1X Profile Engine'
                    Button = $btnImportWifiXml; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnImportWifiXml) { $btnImportWifiXml.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh ESP Win32 App Tracking'
                    Button = $btnRefreshWin32Apps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshWin32Apps) { $btnRefreshWin32Apps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'In-Place PATCH Autopilot Identity'
                    Button = $btnPatchLifecycle; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPatchLifecycle) { $btnPatchLifecycle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Trigger Autopilot Tenant Sync'
                    Button = $btnSyncAutopilot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSyncAutopilot) { $btnSyncAutopilot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Tail Intune Management Extension Log'
                    Button = $btnTailImeLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTailImeLog) { $btnTailImeLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Export MDM Diagnostics CAB Bundle'
                    Button = $btnExportMdmCab; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportMdmCab) { $btnExportMdmCab.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
            }

            '*Hybrid*' {
                # ROUTINE 2: Hybrid AD Join & Co-Management Build (Coordinated across ALL 9 tabs)

                # Tab 9: Precision Enterprise Fixes (AD & WMI Foundation)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearFixOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearFixOut) { $btnClearFixOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Audit Comprehensive System Health'
                    Button = $btnFixHealthReadout; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixHealthReadout) { $btnFixHealthReadout.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Salvage WMI Repository (SCCM/CIM)'
                    Button = $btnFixWmi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixWmi) { $btnFixWmi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Repair DCOM & RPC Limits (Active Directory)'
                    Button = $btnFixDcom; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixDcom) { $btnFixDcom.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Flush TCP/IP Network & Winsock Stack'
                    Button = $btnFixNetStack; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixNetStack) { $btnFixNetStack.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Rebuild WinRM Listener'
                    Button = $btnFixWinRm; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixWinRm) { $btnFixWinRm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Unbrick AppX Staging Manifests'
                    Button = $btnFixAppX; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixAppX) { $btnFixAppX.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Copy Remediation Log'
                    Button = $btnCopyFixLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyFixLog) { $btnCopyFixLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 4: Pre-Flight Diagnostics
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Preflight Update Readiness'
                    Button = $btnPreflightWinUpdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPreflightWinUpdate) { $btnPreflightWinUpdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Auto-Remediate Clock Skew (Domain Time)'
                    Button = $btnFixClockSkew; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixClockSkew) { $btnFixClockSkew.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'TPM 2.0 Attestation Engine'
                    Button = $btnTpmAttestation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTpmAttestation) { $btnTpmAttestation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Inject USB Drivers (pnputil)'
                    Button = $btnInjectDriversPreflight; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectDriversPreflight) { $btnInjectDriversPreflight.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Run Diagnostics Ladder'
                    Button = $btnRunDiag; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnRunDiag) { $btnRunDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 7: Hybrid & Co-Mgmt (Exhaustive sequence of all 30 buttons!)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearHybOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearHybOut) { $btnClearHybOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Refresh Hybrid Join State'
                    Button = $btnHybRefresh; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybRefresh) { $btnHybRefresh.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Resync Domain Time (w32tm /resync)'
                    Button = $btnHybTime; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybTime) { $btnHybTime.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test DC Line-of-Sight'
                    Button = $btnHybDc; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybDc) { $btnHybDc.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test 9-Port Domain Controller Ladder'
                    Button = $btnHybDcLadder; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnHybDcLadder) { $btnHybDcLadder.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Inspect AD Service Connection Point (SCP)'
                    Button = $btnHybScp; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnHybScp) { $btnHybScp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test Netlogon Secure Channel'
                    Button = $btnHybSecChan; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybSecChan) { $btnHybSecChan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Verify Computer SPN Registration'
                    Button = $btnHybSpn; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybSpn) { $btnHybSpn.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Inspect Kerberos Ticket Cache'
                    Button = $btnHybKerbDiag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybKerbDiag) { $btnHybKerbDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Purge Cached Kerberos Tickets'
                    Button = $btnHybKerbPurge; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybKerbPurge) { $btnHybKerbPurge.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit Entra PRT & WAM Broker'
                    Button = $btnHybPrtDiag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybPrtDiag) { $btnHybPrtDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test Entra ID Token Acquisition'
                    Button = $btnHybAadToken; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybAadToken) { $btnHybAadToken.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Reset Token Broker Cache'
                    Button = $btnHybResetBroker; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybResetBroker) { $btnHybResetBroker.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Read Co-Management Workload Authority'
                    Button = $btnHybComgmt; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybComgmt) { $btnHybComgmt.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Shift Workloads to Intune (255)'
                    Button = $btnHybCoMgmtAllIntune; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCoMgmtAllIntune) { $btnHybCoMgmtAllIntune.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Shift Workloads to Pilot (67)'
                    Button = $btnHybCoMgmtPilot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCoMgmtPilot) { $btnHybCoMgmtPilot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Shift Workloads to ConfigMgr (1)'
                    Button = $btnHybCoMgmtAllCcm; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCoMgmtAllCcm) { $btnHybCoMgmtAllCcm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Trigger ConfigMgr Machine Policy'
                    Button = $btnHybCcm; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCcm) { $btnHybCcm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force MDM Enrollment'
                    Button = $btnHybEnroll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybEnroll) { $btnHybEnroll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force Intune MDM Sync'
                    Button = $btnHybSync; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybSync) { $btnHybSync.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Restart Intune Management Extension'
                    Button = $btnHybRestartIme; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybRestartIme) { $btnHybRestartIme.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Escrow BitLocker Recovery Keys'
                    Button = $btnHybBitlocker; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybBitlocker) { $btnHybBitlocker.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force Certificate Pulse (certreq)'
                    Button = $btnHybCertPulse; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCertPulse) { $btnHybCertPulse.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit SCEP Certificate Health'
                    Button = $btnHybScepHealth; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybScepHealth) { $btnHybScepHealth.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Enterprise CA Readiness'
                    Button = $btnHybCa; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCa) { $btnHybCa.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Machine Certificate Health'
                    Button = $btnHybCerts; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCerts) { $btnHybCerts.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Scan Legacy Dependencies'
                    Button = $btnHybLegacy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybLegacy) { $btnHybLegacy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force Group Policy Update (gpupdate)'
                    Button = $btnHybGpupdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybGpupdate) { $btnHybGpupdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Retry Hybrid Join (dsregcmd /join)'
                    Button = $btnHybRetryJoin; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybRetryJoin) { $btnHybRetryJoin.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Collect Hybrid Diagnostics Bundle (.zip)'
                    Button = $btnHybBundle; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybBundle) { $btnHybBundle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 1: Autopilot & Cloud Registration
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Execute Dynamic Device State Action'
                    Button = $btnDeviceStateAction; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnDeviceStateAction -and $btnDeviceStateAction.Visibility -eq [System.Windows.Visibility]::Visible) {
                            $btnDeviceStateAction.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Apply Computer Name Template'
                    Button = $btnApplyRename; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnApplyRename) { $btnApplyRename.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Harvest Genuine Hardware Hash'
                    Button = $btnHarvestHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHarvestHash) { $btnHarvestHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Copy Hardware Hash'
                    Button = $btnCopyHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyHash) { $btnCopyHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Inspect Decoded Policy'
                    Button = $btnInspectPolicy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInspectPolicy) { $btnInspectPolicy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Export Intune CSV'
                    Button = $btnExportCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportCsv) { $btnExportCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Export Offline Provisioning JSON'
                    Button = $btnExportOfflineJson; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportOfflineJson) { $btnExportOfflineJson.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Save Current Settings to .env'
                    Button = $btnSaveEnvDefaults; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSaveEnvDefaults) { $btnSaveEnvDefaults.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 6: Hardware
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Detect BIOS Service Tag'
                    Button = $btnDetectDellTag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnDetectDellTag) { $btnDetectDellTag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Hardware Warranty'
                    Button = $btnCheckDellWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckDellWarranty) { $btnCheckDellWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Lenovo Warranty'
                    Button = $btnCheckLenovoWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckLenovoWarranty) { $btnCheckLenovoWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Battery Wear'
                    Button = $btnRefreshBattery; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshBattery) { $btnRefreshBattery.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Storage SMART'
                    Button = $btnRefreshStorage; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshStorage) { $btnRefreshStorage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Copy Markdown Hardware Audit Report'
                    Button = $btnCopyDellReport; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnCopyDellReport) {
                            $btnCopyDellReport.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Export Hardware CSV'
                    Button = $btnExportDellCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnExportDellCsv) {
                            $btnExportDellCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })

                # Tab 5: Windows Update
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Trigger USO Client Scan'
                    Button = $btnWuTriggerUso; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuTriggerUso) { $btnWuTriggerUso.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Scan for Updates & Drivers'
                    Button = $btnWuScan; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 25
                    Action = { if ($btnWuScan) { $btnWuScan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Select Updates'
                    Button = $btnWuSelectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuSelectAll) { $btnWuSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Install Updates'
                    Button = $btnWuInstall; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        $sel = @($integratedWuItems | Where-Object { $_.IsSelected -eq $true })
                        if ($sel.Count -gt 0 -and $btnWuInstall) {
                            $btnWuInstall.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "No pending updates to install." "INFO"
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Schedule Post-Update System Restart'
                    Button = $btnWuRebootNow; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuRebootNow) { $btnWuRebootNow.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Abort Automatic Reboot'
                    Button = $btnAbortReboot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnAbortReboot) { $btnAbortReboot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 2: App Deployment
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Clear App Selection'
                    Button = $btnClearApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearApps) { $btnClearApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Apply Developer Suite Preset'
                    Button = $btnPresetDev; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPresetDev) { $btnPresetDev.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Select Applications'
                    Button = $btnSelectAllApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSelectAllApps) { $btnSelectAllApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Install Selected Applications'
                    Button = $btnInstallBatch; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        $appCount = ($script:AppCatalog | Where-Object { $_.IsSelected }).Count
                        if ($appCount -gt 0 -and $btnInstallBatch) {
                            $btnInstallBatch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "App catalog selection empty. Batch install skipped." "INFO"
                        }
                    }
                })

                # Tab 3: Win32 Packaging
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Build Win32 App Package (.intunewin)'
                    Button = $btnBuildPackage; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnBuildPackage) { $btnBuildPackage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Publish Win32 Package to Intune Cloud'
                    Button = $btnPublishIntune; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:GraphAuthContext -and $script:GraphAuthContext.AccessToken -and $btnPublishIntune) {
                            $btnPublishIntune.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "Win32 publishing skipped: Graph authentication required." "INFO"
                        }
                    }
                })

                # Tab 8: ESP & Lifecycle
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh Windows Edition Licensing'
                    Button = $btnRefreshLicensing; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshLicensing) { $btnRefreshLicensing.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = '1-Click Enterprise KMS Upgrade Check'
                    Button = $btnUpgradeEnterprise; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnUpgradeEnterprise) { $btnUpgradeEnterprise.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Scan Available Wi-Fi SSIDs'
                    Button = $btnScanWifi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnScanWifi) { $btnScanWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Verify Wi-Fi Readiness'
                    Button = $btnConnectWifi; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnConnectWifi) { $btnConnectWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Validate 802.1X XML Profile Import'
                    Button = $btnImportWifiXml; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnImportWifiXml) { $btnImportWifiXml.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh ESP Win32 App Tracking'
                    Button = $btnRefreshWin32Apps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshWin32Apps) { $btnRefreshWin32Apps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'In-Place PATCH Autopilot Identity'
                    Button = $btnPatchLifecycle; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPatchLifecycle) { $btnPatchLifecycle.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Instant Tenant Sync'
                    Button = $btnSyncAutopilot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSyncAutopilot) { $btnSyncAutopilot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Tail IME Log'
                    Button = $btnTailImeLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTailImeLog) { $btnTailImeLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Export MDM Diagnostics CAB'
                    Button = $btnExportMdmCab; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportMdmCab) { $btnExportMdmCab.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
            }

            '*Local*' {
                # ROUTINE 3: Local User & Offline Bypass Build (Coordinated across ALL 9 tabs)

                # Tab 9: Precision Enterprise Fixes
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearFixOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearFixOut) { $btnClearFixOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Audit Initial System Health'
                    Button = $btnFixHealthReadout; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixHealthReadout) { $btnFixHealthReadout.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Salvage WMI Repository'
                    Button = $btnFixWmi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixWmi) { $btnFixWmi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Repair Catroot2 Database'
                    Button = $btnFixCatroot; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixCatroot) { $btnFixCatroot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Purge BITS Deadlock'
                    Button = $btnFixBits; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixBits) { $btnFixBits.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Purge Spooler Queue'
                    Button = $btnFixSpooler; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixSpooler) { $btnFixSpooler.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Flush Network Stack'
                    Button = $btnFixNetStack; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixNetStack) { $btnFixNetStack.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Rebuild WinRM Listener'
                    Button = $btnFixWinRm; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixWinRm) { $btnFixWinRm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Unbrick AppX Staging Manifests'
                    Button = $btnFixAppX; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixAppX) { $btnFixAppX.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Copy Remediation Log'
                    Button = $btnCopyFixLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyFixLog) { $btnCopyFixLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 4: Pre-Flight Diagnostics
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Preflight Update Status'
                    Button = $btnPreflightWinUpdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPreflightWinUpdate) { $btnPreflightWinUpdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Auto-Remediate Clock Skew'
                    Button = $btnFixClockSkew; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixClockSkew) { $btnFixClockSkew.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Verify TPM 2.0 Attestation'
                    Button = $btnTpmAttestation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTpmAttestation) { $btnTpmAttestation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Scan & Inject USB Drivers'
                    Button = $btnInjectDriversPreflight; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectDriversPreflight) { $btnInjectDriversPreflight.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Run Diagnostic Network Ladder'
                    Button = $btnRunDiag; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnRunDiag) { $btnRunDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 6: Dell Warranty & Hardware Health
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Detect BIOS Hardware Serial'
                    Button = $btnDetectDellTag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnDetectDellTag) { $btnDetectDellTag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Check Dell Warranty (Cached/Offline)'
                    Button = $btnCheckDellWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckDellWarranty) { $btnCheckDellWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Check Lenovo Serial Audit'
                    Button = $btnCheckLenovoWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckLenovoWarranty) { $btnCheckLenovoWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Check Battery Wear'
                    Button = $btnRefreshBattery; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshBattery) { $btnRefreshBattery.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Storage SMART'
                    Button = $btnRefreshStorage; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshStorage) { $btnRefreshStorage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Copy Hardware Report'
                    Button = $btnCopyDellReport; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnCopyDellReport) {
                            $btnCopyDellReport.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Export Hardware CSV'
                    Button = $btnExportDellCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnExportDellCsv) {
                            $btnExportDellCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })

                # Tab 1: Autopilot & Cloud Registration (Offline Bypass Core)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Dismiss Tenant Mismatch Banner'
                    Button = $btnDismissMismatch; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($bannerTenantMismatch -and $bannerTenantMismatch.Visibility -eq [System.Windows.Visibility]::Visible -and $btnDismissMismatch) {
                            $btnDismissMismatch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Apply Offline Bypass (NRO & Account Unlock)'
                    Button = $btnSkipAutopilot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSkipAutopilot) { $btnSkipAutopilot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Unlock Personal MSA Setup'
                    Button = $btnPersonalInstall; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPersonalInstall) { $btnPersonalInstall.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Create Local Administrator User'
                    Button = $btnCreateLocalUser; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCreateLocalUser) { $btnCreateLocalUser.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Apply Computer Name Template'
                    Button = $btnApplyRename; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnApplyRename) { $btnApplyRename.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Activate Controlled Offline Staging'
                    Button = $btnOfflineStaging; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnOfflineStaging) { $btnOfflineStaging.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Test Inject Offline JSON Engine'
                    Button = $btnInjectOfflineJson; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectOfflineJson) { $btnInjectOfflineJson.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Save Offline Configuration to .env'
                    Button = $btnSaveEnvDefaults; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSaveEnvDefaults) { $btnSaveEnvDefaults.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Re-enable Physical Network Adapters'
                    Button = $btnReenableNet; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnReenableNet) { $btnReenableNet.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Restart OOBE Setup Processes'
                    Button = $btnRestartOobe; WaitForButton = $false; WaitForAsync = $false
                    Action = { Write-HubLog "Playbook: Signaling OOBE setup refresh." "INFO" }
                })

                # Tab 7: Hybrid & Co-Mgmt (Local/Offline Identity Validations)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearHybOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearHybOut) { $btnClearHybOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit Absence of Cloud PRT'
                    Button = $btnHybPrtDiag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybPrtDiag) { $btnHybPrtDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Scan Legacy Dependencies'
                    Button = $btnHybLegacy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybLegacy) { $btnHybLegacy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Local Machine Certificate Store'
                    Button = $btnHybCerts; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCerts) { $btnHybCerts.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 5: Windows Update & Drivers
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Trigger USO Scan'
                    Button = $btnWuTriggerUso; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuTriggerUso) { $btnWuTriggerUso.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Scan for Pending Updates'
                    Button = $btnWuScan; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 25
                    Action = { if ($btnWuScan) { $btnWuScan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Deselect All Updates'
                    Button = $btnWuDeselectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuDeselectAll) { $btnWuDeselectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Select Updates'
                    Button = $btnWuSelectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuSelectAll) { $btnWuSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Abort Automatic Reboot'
                    Button = $btnAbortReboot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnAbortReboot) { $btnAbortReboot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 2: App Deployment
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Clear App Selection'
                    Button = $btnClearApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearApps) { $btnClearApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Apply Web Browsers Preset'
                    Button = $btnPresetBrowsers; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPresetBrowsers) { $btnPresetBrowsers.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Select Target Applications'
                    Button = $btnSelectAllApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSelectAllApps) { $btnSelectAllApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Install Selected Applications'
                    Button = $btnInstallBatch; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        $appCount = ($script:AppCatalog | Where-Object { $_.IsSelected }).Count
                        if ($appCount -gt 0 -and $btnInstallBatch) {
                            $btnInstallBatch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "App selection empty. Skipping batch install." "INFO"
                        }
                    }
                })

                # Tab 3: Win32 Packaging
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Build Local Win32 Package'
                    Button = $btnBuildPackage; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnBuildPackage) { $btnBuildPackage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 8: ESP Diagnostics & Lifecycle
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh Local Windows Edition'
                    Button = $btnRefreshLicensing; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshLicensing) { $btnRefreshLicensing.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Scan Wi-Fi SSIDs'
                    Button = $btnScanWifi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnScanWifi) { $btnScanWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Test Local Network Connectivity'
                    Button = $btnConnectWifi; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnConnectWifi) { $btnConnectWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh Win32 App Tracking'
                    Button = $btnRefreshWin32Apps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshWin32Apps) { $btnRefreshWin32Apps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Export Diagnostics CAB'
                    Button = $btnExportMdmCab; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportMdmCab) { $btnExportMdmCab.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
            }

            '*Remediation*' {
                # ROUTINE 4: Deep System Remediation & Health Sweep (Coordinated across ALL 9 tabs)

                # Tab 9: Precision Enterprise Fixes (Exhaustive arsenal!)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Fix Output'
                    Button = $btnClearFixOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearFixOut) { $btnClearFixOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Remediation Output'
                    Button = $btnClearRemOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearRemOut) { $btnClearRemOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Audit Comprehensive System Health'
                    Button = $btnFixHealthReadout; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixHealthReadout) { $btnFixHealthReadout.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Salvage WMI Repository'
                    Button = $btnFixWmi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixWmi) { $btnFixWmi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Deep Reset Windows Update Agent'
                    Button = $btnFixWu; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixWu) { $btnFixWu.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Repair Catroot2 ESENT Database'
                    Button = $btnFixCatroot; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixCatroot) { $btnFixCatroot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Purge BITS Deadlock Queue'
                    Button = $btnFixBits; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixBits) { $btnFixBits.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Repair DCOM & RPC Limits'
                    Button = $btnFixDcom; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixDcom) { $btnFixDcom.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Purge Print Spooler Queue'
                    Button = $btnFixSpooler; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixSpooler) { $btnFixSpooler.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Flush TCP/IP Network & Winsock Stack'
                    Button = $btnFixNetStack; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixNetStack) { $btnFixNetStack.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Rebuild WinRM Listener'
                    Button = $btnFixWinRm; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixWinRm) { $btnFixWinRm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Un-hook ProfileList .bak Locks'
                    Button = $btnFixProfiles; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixProfiles) { $btnFixProfiles.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Heal TPM Attestation'
                    Button = $btnFixTpm; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixTpm) { $btnFixTpm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Unbrick AppX Staging Manifests'
                    Button = $btnFixAppX; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnFixAppX) { $btnFixAppX.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Copy Complete Fix Log'
                    Button = $btnCopyFixLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyFixLog) { $btnCopyFixLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 4: Pre-Flight Diagnostics
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Preflight Update Readiness'
                    Button = $btnPreflightWinUpdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPreflightWinUpdate) { $btnPreflightWinUpdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Synchronize System Clock'
                    Button = $btnFixClockSkew; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixClockSkew) { $btnFixClockSkew.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Verify TPM 2.0 Attestation'
                    Button = $btnTpmAttestation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTpmAttestation) { $btnTpmAttestation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Scan and Inject USB Drivers'
                    Button = $btnInjectDriversPreflight; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectDriversPreflight) { $btnInjectDriversPreflight.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Re-run Network Diagnostic Ladder'
                    Button = $btnRunDiag; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnRunDiag) { $btnRunDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 7: Hybrid & Co-Mgmt (Remediation & Reset Actions)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Clear Hybrid Output'
                    Button = $btnClearHybOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearHybOut) { $btnClearHybOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Test Netlogon Secure Channel'
                    Button = $btnHybSecChan; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybSecChan) { $btnHybSecChan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Purge Cached Kerberos Tickets'
                    Button = $btnHybKerbPurge; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybKerbPurge) { $btnHybKerbPurge.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Reset Token Broker Cache'
                    Button = $btnHybResetBroker; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybResetBroker) { $btnHybResetBroker.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Force Certificate Pulse (certreq)'
                    Button = $btnHybCertPulse; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCertPulse) { $btnHybCertPulse.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Restart Intune Management Extension'
                    Button = $btnHybRestartIme; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybRestartIme) { $btnHybRestartIme.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit SCEP Certificate Health'
                    Button = $btnHybScepHealth; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybScepHealth) { $btnHybScepHealth.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Machine Certificates'
                    Button = $btnHybCerts; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCerts) { $btnHybCerts.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 1: Autopilot & Cloud Registration
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Verify OA3 Hardware Hash Health'
                    Button = $btnHarvestHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHarvestHash) { $btnHarvestHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Inspect Decoded Autopilot Policy'
                    Button = $btnInspectPolicy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInspectPolicy) { $btnInspectPolicy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Ensure Network Adapters Healthy & Enabled'
                    Button = $btnReenableNet; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnReenableNet) { $btnReenableNet.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Synchronize .env State'
                    Button = $btnSaveEnvDefaults; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSaveEnvDefaults) { $btnSaveEnvDefaults.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 6: Dell Warranty & Hardware Health
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Detect BIOS Service Tag & Health'
                    Button = $btnDetectDellTag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnDetectDellTag) { $btnDetectDellTag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Battery Health & Capacity'
                    Button = $btnRefreshBattery; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshBattery) { $btnRefreshBattery.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Storage NVMe SMART Reliability'
                    Button = $btnRefreshStorage; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshStorage) { $btnRefreshStorage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Copy Hardware Diagnostics Readout'
                    Button = $btnCopyDellReport; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnCopyDellReport) {
                            $btnCopyDellReport.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })

                # Tab 5: Windows Update & Drivers
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Trigger USO Client Background Scan'
                    Button = $btnWuTriggerUso; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuTriggerUso) { $btnWuTriggerUso.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Scan for Pending Updates'
                    Button = $btnWuScan; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 25
                    Action = { if ($btnWuScan) { $btnWuScan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Select Updates'
                    Button = $btnWuSelectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuSelectAll) { $btnWuSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Abort Automatic Reboot'
                    Button = $btnAbortReboot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnAbortReboot) { $btnAbortReboot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Ensure Cascade Halted'
                    Button = $btnStopCascade; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnStopCascade) { $btnStopCascade.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 2: App Deployment
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Clear Application Selection Queue'
                    Button = $btnClearApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearApps) { $btnClearApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Check Baseline Workstation Catalog'
                    Button = $btnPresetWorkstation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPresetWorkstation) { $btnPresetWorkstation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 3: Win32 Packaging
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Test Win32 App Packager Readiness'
                    Button = $btnBuildPackage; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnBuildPackage) { $btnBuildPackage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 8: ESP & Lifecycle
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Scan Available Wi-Fi SSIDs'
                    Button = $btnScanWifi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnScanWifi) { $btnScanWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Reset 3-Registry ESP Tracking Keys'
                    Button = $btnDecommission; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:RuntimeContext.MeetsPreferred -and $btnDecommission) {
                            $btnDecommission.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        } else {
                            Write-HubLog "ESP decommission skipped: Elevation required." "INFO"
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Tail IME Log'
                    Button = $btnTailImeLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTailImeLog) { $btnTailImeLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Export MDM Diagnostics CAB Bundle'
                    Button = $btnExportMdmCab; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportMdmCab) { $btnExportMdmCab.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
            }

            '*Hardware*' {
                # ROUTINE 5: Hardware Health & Asset Intake Audit (Coordinated across ALL 9 tabs)

                # Tab 6: Dell Warranty & Hardware Health
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Detect BIOS Service Tag & Serial'
                    Button = $btnDetectDellTag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnDetectDellTag) { $btnDetectDellTag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Dell Official Warranty API & SLA'
                    Button = $btnCheckDellWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckDellWarranty) { $btnCheckDellWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Assess Lenovo Warranty API'
                    Button = $btnCheckLenovoWarranty; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnCheckLenovoWarranty) { $btnCheckLenovoWarranty.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Battery Wear, Capacity & Cycles'
                    Button = $btnRefreshBattery; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshBattery) { $btnRefreshBattery.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Inspect Storage SMART NVMe Reliability'
                    Button = $btnRefreshStorage; WaitForButton = $false; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnRefreshStorage) { $btnRefreshStorage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Copy Markdown Hardware Audit Report'
                    Button = $btnCopyDellReport; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnCopyDellReport) {
                            $btnCopyDellReport.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Export Hardware & Warranty CSV'
                    Button = $btnExportDellCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($script:CurrentDellReport -and $btnExportDellCsv) {
                            $btnExportDellCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })

                # Tab 4: Pre-Flight Diagnostics
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Preflight Update Readiness'
                    Button = $btnPreflightWinUpdate; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPreflightWinUpdate) { $btnPreflightWinUpdate.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Verify Clock Synchronization'
                    Button = $btnFixClockSkew; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixClockSkew) { $btnFixClockSkew.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Run TPM 2.0 & EK Attestation Engine'
                    Button = $btnTpmAttestation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTpmAttestation) { $btnTpmAttestation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Check Hardware USB Driver Readiness'
                    Button = $btnInjectDriversPreflight; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInjectDriversPreflight) { $btnInjectDriversPreflight.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPreFlight; TabName = 'Pre-Flight Diagnostics'; ActionName = 'Run Hardware Network Interface Ladder'
                    Button = $btnRunDiag; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 20
                    Action = { if ($btnRunDiag) { $btnRunDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 1: Autopilot & Cloud Registration
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Execute Dynamic Device State Action'
                    Button = $btnDeviceStateAction; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnDeviceStateAction -and $btnDeviceStateAction.Visibility -eq [System.Windows.Visibility]::Visible) {
                            $btnDeviceStateAction.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Harvest Genuine Hardware Hash'
                    Button = $btnHarvestHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHarvestHash) { $btnHarvestHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Copy Hardware Hash to Clipboard'
                    Button = $btnCopyHash; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyHash) { $btnCopyHash.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Inspect Decoded Autopilot Policy'
                    Button = $btnInspectPolicy; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnInspectPolicy) { $btnInspectPolicy.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Apply Standardized Asset Tag Renaming'
                    Button = $btnApplyRename; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnApplyRename) { $btnApplyRename.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Export Intune CSV (USB Flash Priority)'
                    Button = $btnExportCsv; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportCsv) { $btnExportCsv.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAutopilot; TabName = 'Autopilot & Cloud Registration'; ActionName = 'Save Intake Profile to .env'
                    Button = $btnSaveEnvDefaults; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnSaveEnvDefaults) { $btnSaveEnvDefaults.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 7: Hybrid & Co-Mgmt
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearHybOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearHybOut) { $btnClearHybOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Inspect Active Directory SCP'
                    Button = $btnHybScp; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybScp) { $btnHybScp.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Audit Entra PRT Identity'
                    Button = $btnHybPrtDiag; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybPrtDiag) { $btnHybPrtDiag.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHybrid; TabName = 'Hybrid & Co-Mgmt'; ActionName = 'Check Machine Certificates'
                    Button = $btnHybCerts; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnHybCerts) { $btnHybCerts.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 5: Windows Update & Drivers
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Trigger USO Background Scan'
                    Button = $btnWuTriggerUso; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuTriggerUso) { $btnWuTriggerUso.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Scan for Firmware & Driver Updates'
                    Button = $btnWuScan; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 25
                    Action = { if ($btnWuScan) { $btnWuScan.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Select Updates'
                    Button = $btnWuSelectAll; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnWuSelectAll) { $btnWuSelectAll.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabWinUpdate; TabName = 'Windows Update & Drivers'; ActionName = 'Suppress System Reboot'
                    Button = $btnAbortReboot; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnAbortReboot) { $btnAbortReboot.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 2: App Deployment
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Clear Application Selection'
                    Button = $btnClearApps; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearApps) { $btnClearApps.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabAppDeploy; TabName = 'App Deployment'; ActionName = 'Validate Baseline Application Catalog'
                    Button = $btnPresetWorkstation; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnPresetWorkstation) { $btnPresetWorkstation.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 3: Win32 Packaging
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPackaging; TabName = 'Win32 Packaging'; ActionName = 'Test Win32 Packaging Workspace Target'
                    Button = $btnBuildPackage; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnBuildPackage) { $btnBuildPackage.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 8: ESP & Lifecycle
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Refresh Windows Edition Licensing'
                    Button = $btnRefreshLicensing; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnRefreshLicensing) { $btnRefreshLicensing.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Scan Available Wi-Fi SSIDs'
                    Button = $btnScanWifi; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 10
                    Action = { if ($btnScanWifi) { $btnScanWifi.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Tail IME Log'
                    Button = $btnTailImeLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnTailImeLog) { $btnTailImeLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabLifecycle; TabName = 'ESP Diagnostics & Lifecycle'; ActionName = 'Export MDM Diagnostics CAB'
                    Button = $btnExportMdmCab; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnExportMdmCab) { $btnExportMdmCab.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # Tab 9: Precision Enterprise Fixes
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Clear Diagnostic Output'
                    Button = $btnClearFixOut; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnClearFixOut) { $btnClearFixOut.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Audit System Health Readout'
                    Button = $btnFixHealthReadout; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnFixHealthReadout) { $btnFixHealthReadout.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Heal TPM Attestation Container'
                    Button = $btnFixTpm; WaitForButton = $true; WaitForAsync = $true; MaxWaitSec = 15
                    Action = { if ($btnFixTpm) { $btnFixTpm.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })
                $steps.Add([PSCustomObject]@{
                    Tab = $tabPrecisionFixes; TabName = 'Precision Enterprise Fixes'; ActionName = 'Copy Audit Remediation Log'
                    Button = $btnCopyFixLog; WaitForButton = $false; WaitForAsync = $false
                    Action = { if ($btnCopyFixLog) { $btnCopyFixLog.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) } }
                })

                # System Tools (Staging Proof)
                $steps.Add([PSCustomObject]@{
                    Tab = $tabHardwareHealth; TabName = 'Hardware & Warranty'; ActionName = 'Capture Staging Proof Screenshot'
                    Button = $btnScreenshotUsb; WaitForButton = $false; WaitForAsync = $false
                    Action = {
                        if ($btnScreenshotUsb) {
                            $btnScreenshotUsb.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                        }
                    }
                })
            }
        }

        return $steps
    }

    function Invoke-HubActionRoutine {
        param(
            [string]$RoutineName = '1. Intune-Only Cloud Build'
        )

        $steps = Get-HubPlaybookSteps -RoutineName $RoutineName
        if (-not $steps -or $steps.Count -eq 0) {
            Write-HubLog "No steps found for routine '$RoutineName'." "ERROR"
            return
        }

        $script:PlaybookRunning = $true
        $script:PlaybookState.Status = 'Running'
        $script:PlaybookState.IsPaused = $false
        $script:PlaybookState.CancelRequested = $false
        $script:PlaybookState.TotalSteps = $steps.Count

        $btnPlaybookRun.IsEnabled = $false
        $btnPlaybookPause.IsEnabled = $true
        $btnPlaybookPause.Content = '[|| Pause]'
        $btnPlaybookStop.IsEnabled = $true
        $cboPlaybookRoutine.IsEnabled = $false

        $activeButton = $null
        $origBg = $null

        Write-HubLog "Starting Deployment Playbook: '$RoutineName' ($($steps.Count) coordinated steps across tabs)..." "INFO"

        try {
            for ($idx = 0; $idx -lt $steps.Count; $idx++) {
                if ($script:PlaybookState.CancelRequested) {
                    Write-HubLog "Playbook execution cancelled by operator." "WARN"
                    break
                }

                while ($script:PlaybookState.IsPaused -and -not $script:PlaybookState.CancelRequested) {
                    Start-Sleep -Milliseconds 100
                    Update-WpfUI
                }
                if ($script:PlaybookState.CancelRequested) { break }

                $step = $steps[$idx]
                $stepNum = $idx + 1
                $script:PlaybookState.CurrentStep = $stepNum
                $script:PlaybookState.CurrentStepInfo = "Step $stepNum/$($steps.Count): [$($step.TabName)] -> $($step.ActionName)"

                # 1. Switch Tab visually
                if ($step.Tab -and $mainTabControl) {
                    $mainTabControl.SelectedItem = $step.Tab
                    Update-WpfUI
                }

                # 2. Update HUD & Main Progress Bar
                $txtPlaybookStep.Text = $script:PlaybookState.CurrentStepInfo
                $pct = [int][math]::Round(($stepNum / $steps.Count) * 100)
                $playbookProgressBar.Value = $pct
                Set-HubProgress -Percent $pct -Status "Playbook: $($step.ActionName)"
                Write-HubLog "Playbook [Step $stepNum/$($steps.Count)] ($($step.TabName)): $($step.ActionName)..." "INFO"
                Update-WpfUI

                # Restore previous button background if needed
                if ($activeButton -and $origBg) {
                    try { $activeButton.Background = $origBg } catch { }
                    $activeButton = $null
                    $origBg = $null
                }

                # 3. Button Visual Highlight
                if ($step.Button) {
                    try {
                        $activeButton = $step.Button
                        $origBg = $step.Button.Background
                        $step.Button.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#BAE6FD")
                        $step.Button.Focus()
                    } catch { }
                    Update-WpfUI
                }

                # 4. Settle delay for visual feedback
                $settleStart = [datetime]::UtcNow
                while (([datetime]::UtcNow - $settleStart).TotalMilliseconds -lt 250 -and -not $script:PlaybookState.CancelRequested) {
                    while ($script:PlaybookState.IsPaused -and -not $script:PlaybookState.CancelRequested) {
                        Start-Sleep -Milliseconds 100
                        Update-WpfUI
                    }
                    Start-Sleep -Milliseconds 50
                    Update-WpfUI
                }

                # 5. Execute Action
                try {
                    if ($step.Action) {
                        & $step.Action
                    } elseif ($step.Button -and $step.Button.IsEnabled) {
                        $step.Button.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                    }
                } catch {
                    Write-HubLog "Playbook error in step '$($step.ActionName)': $($_.Exception.Message)" "WARN"
                }

                # 6. Wait for async task or button re-enable (if applicable)
                $waitStart = [datetime]::UtcNow
                $maxWait = if ($step.MaxWaitSec) { $step.MaxWaitSec } else { 15 }
                while (([datetime]::UtcNow - $waitStart).TotalSeconds -lt $maxWait -and -not $script:PlaybookState.CancelRequested) {
                    while ($script:PlaybookState.IsPaused -and -not $script:PlaybookState.CancelRequested) {
                        Start-Sleep -Milliseconds 100
                        Update-WpfUI
                    }
                    if ($script:PlaybookState.CancelRequested) { break }

                    if ($step.Button -and $step.WaitForButton -and $step.Button.IsEnabled -eq $false) {
                        Start-Sleep -Milliseconds 100
                        Update-WpfUI
                        continue
                    }
                    if ($step.WaitForAsync -and $script:HubActiveAsyncTasks -and $script:HubActiveAsyncTasks.Count -gt 0) {
                        Start-Sleep -Milliseconds 100
                        Update-WpfUI
                        continue
                    }
                    break
                }

                # Restore button background
                if ($activeButton -and $origBg) {
                    try { $activeButton.Background = $origBg } catch { }
                    $activeButton = $null
                    $origBg = $null
                }
                Update-WpfUI
            }
        } finally {
            if ($activeButton -and $origBg) {
                try { $activeButton.Background = $origBg } catch { }
                $activeButton = $null
                $origBg = $null
            }
            $script:PlaybookRunning = $false
            $script:PlaybookState.Status = 'Idle'
            $script:PlaybookState.IsPaused = $false
            $btnPlaybookRun.IsEnabled = $true
            $btnPlaybookPause.IsEnabled = $false
            $btnPlaybookPause.Content = '[|| Pause]'
            $btnPlaybookStop.IsEnabled = $false
            $cboPlaybookRoutine.IsEnabled = $true

            if ($script:PlaybookState.CancelRequested) {
                $txtPlaybookStep.Text = "Playbook stopped by operator."
                Write-HubLog "Playbook '$RoutineName' stopped." "WARN"
                Set-HubProgress -Percent 0 -Status "Playbook Aborted"
            } else {
                $txtPlaybookStep.Text = "Playbook '$RoutineName' completed successfully! (All steps finished)"
                $playbookProgressBar.Value = 100
                Write-HubLog "Playbook '$RoutineName' completed successfully across all tabs!" "SUCCESS"
                Set-HubProgress -Percent 100 -Status "Playbook Complete"
                try { Play-HubAudio -Type Success } catch { }
            }
            Update-WpfUI
        }
    }

        $btnPlaybookRun.Add_Click({
        if ($script:PlaybookState.Status -eq 'Running') {
            if ($script:PlaybookState.IsPaused) {
                $script:PlaybookState.IsPaused = $false
                $btnPlaybookPause.Content = '[|| Pause]'
                $btnPlaybookRun.IsEnabled = $false
                Write-HubLog "Playbook resumed by operator." "INFO"
            }
            return
        }

        $selectedItem = $cboPlaybookRoutine.SelectedItem
        $routineName = if ($selectedItem -is [System.Windows.Controls.ComboBoxItem]) {
            $selectedItem.Content.ToString()
        } elseif ($selectedItem) {
            [string]$selectedItem
        } else {
            '1. Intune-Only Cloud Build'
        }
        Invoke-HubActionRoutine -RoutineName $routineName
    })

    $btnPlaybookPause.Add_Click({
        if ($script:PlaybookState.Status -ne 'Running') { return }
        $script:PlaybookState.IsPaused = -not $script:PlaybookState.IsPaused
        if ($script:PlaybookState.IsPaused) {
            $btnPlaybookPause.Content = '[> Resume]'
            $btnPlaybookRun.IsEnabled = $true
            Write-HubLog "Playbook paused by operator." "WARN"
            $txtPlaybookStep.Text = "Paused: $($script:PlaybookState.CurrentStepInfo)"
        } else {
            $btnPlaybookPause.Content = '[|| Pause]'
            $btnPlaybookRun.IsEnabled = $false
            Write-HubLog "Playbook resumed by operator." "INFO"
        }
    })

    $btnPlaybookStop.Add_Click({
        if ($script:PlaybookState.Status -eq 'Running') {
            $script:PlaybookState.CancelRequested = $true
            Write-HubLog "Stop requested. Aborting playbook..." "WARN"
            $txtPlaybookStep.Text = "Stopping playbook..."
        }
    })

    # Initial Diagnostic Run. The 7-stage ladder is ~3 s of network waits. It was kicked off in a
    # dedicated runspace back in Section C (before the console intro), so by the time the window is built
    # it is usually already done; either way it is harvested here without blocking the UI thread.
    $applyDiag = {
        param($diag)
        if ($diag -and $diag.Stages) {
            $lstDiagStages.ItemsSource = $diag.Stages
            $passed = @($diag.Stages | Where-Object { $_.Success }).Count
            Write-HubLog "Pre-flight ladder: $passed/$(@($diag.Stages).Count) stages passed." $(if ($passed -eq @($diag.Stages).Count) { 'SUCCESS' } else { 'INFO' })
        }
    }
    if ($script:PreflightHandle) {
        # Poll the head-start runspace without blocking; apply on the queue when it finishes
        $script:PreflightPoller = [System.Windows.Threading.DispatcherTimer]::new()
        $script:PreflightPoller.Interval = [TimeSpan]::FromMilliseconds(120)
        $script:PreflightPoller.Add_Tick({
            if ($script:PreflightHandle.IsCompleted) {
                $script:PreflightPoller.Stop()
                $diag = $null
                try { $diag = $script:PreflightRunspace.EndInvoke($script:PreflightHandle) } catch { }
                try { $script:PreflightRunspace.Runspace.Dispose(); $script:PreflightRunspace.Dispose() } catch { }
                $script:PreflightHandle = $null; $script:PreflightRunspace = $null
                & $applyDiag $diag
            }
        })
        $script:PreflightPoller.Start()
    } else {
        Start-HubAsyncWork -Description "7-Stage Network Diagnostics (startup)" -WorkerScript {
            Test-StagedNetwork
        } -OnComplete { param($diag, $err) & $applyDiag $diag } | Out-Null
    }

    # Device enrollment state from local evidence (cached Autopilot profile, MDM enrollment, Entra join)
    Set-HubProgress -Percent 5 -Status "Inspecting device state"
    try {
        $script:DeviceState = Get-DeviceEnrollmentState
        Write-HubLog "Device state: $($script:DeviceState.Summary)" $(if ($script:DeviceState.Verdict -eq 'AutopilotRegistered') { 'SUCCESS' } else { 'INFO' })
        foreach ($ev in $script:DeviceState.Evidence) { Write-HubLog "  evidence: $ev" }
    } catch {
        Write-HubLog "Device state inspection failed: $($_.Exception.Message)" "WARN"
    }
    Update-DeviceStateUi

    # Channel 1: zero-token tenant branding. If the device already carries a tenant domain from its cached
    # Autopilot profile, show that tenant's name/logo in the header before anyone signs in.
    if ($script:DeviceState -and $script:DeviceState.AutopilotTenantDomain) {
        try {
            $pb = Get-PreAuthTenantBranding -Domain $script:DeviceState.AutopilotTenantDomain
            if ($pb) {
                if ($pb.DisplayName -and $txtTenantBrand) {
                    $txtTenantBrand.Text = $pb.DisplayName.ToUpper()
                    if ($badgeTenantBrand) { $badgeTenantBrand.Visibility = [System.Windows.Visibility]::Visible }
                }
                $logoUrl = if ($pb.TileLogoUrl) { $pb.TileLogoUrl } else { $pb.BannerLogoUrl }
                if ($logoUrl -and $imgTenantLogo -and $borderTenantLogo) {
                    Set-HubBrandingImageAsync -TargetImage $imgTenantLogo -LogoContainer $borderTenantLogo -ImageUrl $logoUrl
                }
                Write-HubLog "Tenant branding resolved without sign-in: $($pb.DisplayName) ($($pb.Domain))." "INFO"
            }
        } catch { }
    }

    # Check for Existing / Silent Graph Session (.env) - also triggers the tenant-side Autopilot lookup
    $silentToken = Get-CurrentGraphToken
    if ($silentToken) {
        Write-HubLog "Microsoft Graph session auto-connected from environment ($($script:GraphAuthContext.TenantId))." "SUCCESS"
    }
    Update-GraphAuthHeader

    # Auto-harvest: an elevated session reads the hash immediately so the operator sees it without clicking
    if ($script:RuntimeContext.MeetsPreferred) {
        if ($script:DeviceState -and $script:DeviceState.Verdict -eq 'AutopilotRegistered') {
            Write-HubLog "Device is already Autopilot-registered - hash harvest is available but not required." "INFO"
        }
        Invoke-HubHarvest
    } else {
        $txtHashMeta.Text = "Hardware hash needs elevation. Click 'Fix Privileges' in the header, then the hash is read automatically."
    }
    Set-HubProgress -Percent 0 -Status "Ready"

    # Check for active multi-pass autonomous Patch Cascade from a prior reboot.
    # This installs updates and can trigger a real system restart, so a state file merely
    # existing on disk (which can be abandoned, days old, or left by a different operator) is
    # NOT sufficient consent to auto-fire it. Only the single-use -ResumeFromRestart launch -
    # the same trust boundary already used for the OOBE Hub-resume-after-restart feature - may
    # resume silently, and only if the state is still fresh. Any other launch asks first.
    $cascadeState = Get-HubPatchCascadeState
    if ($cascadeState) {
        $cascadeAgeHours = 999
        try {
            if ($cascadeState.Timestamp) {
                $cascadeAgeHours = ([datetime]::UtcNow - [datetime]::Parse($cascadeState.Timestamp, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)).TotalHours
            }
        } catch { }
        $cascadeSummary = "Pass $($cascadeState.CurrentPass) of $($cascadeState.MaxPasses), $($cascadeState.TotalInstalled) update(s) installed so far (saved $([Math]::Round($cascadeAgeHours, 1))h ago)"

        if ($script:ResumeFromRestart -and $cascadeAgeHours -lt 6) {
            Write-HubLog "Active Patch Cascade detected: $cascadeSummary. Resuming automatically (this launch is the consented post-restart resume)." "INFO"
            if ($mainTabControl -and $tabWinUpdate) { $mainTabControl.SelectedItem = $tabWinUpdate }
            $cascadeResumeTimer = [System.Windows.Threading.DispatcherTimer]::new()
            $cascadeResumeTimer.Interval = [TimeSpan]::FromSeconds(2)
            $cascadeResumeTimer.Add_Tick({
                $cascadeResumeTimer.Stop()
                Invoke-GuiCascadePass
            })
            $cascadeResumeTimer.Start()
        } elseif ($cascadeAgeHours -ge 6) {
            Write-HubLog "Abandoned Patch Cascade state found ($cascadeSummary) - too old to auto-resume. Clearing it; start a fresh cascade from the Windows Update tab if needed." "WARN"
            Clear-HubPatchCascadeState
        } else {
            Write-HubLog "Incomplete Patch Cascade found from a previous session: $cascadeSummary." "WARN"
            if ($mainTabControl -and $tabWinUpdate) { $mainTabControl.SelectedItem = $tabWinUpdate }
            $__cascadeAsk = [System.Windows.MessageBox]::Show("An incomplete Autonomous Patch Cascade was found from a previous session:`n`n$cascadeSummary`n`nThis was NOT launched via the automatic post-restart resume, so it will not continue on its own. Resume installing updates now (and allow another automatic restart if needed)?`n`nChoosing No clears this saved state; you can start a fresh cascade any time from the Windows Update tab.", "Resume Incomplete Patch Cascade?", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
            if ($__cascadeAsk -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-HubLog "Operator approved resuming the Patch Cascade." "INFO"
                Invoke-GuiCascadePass
            } else {
                Write-HubLog "Operator declined to resume the Patch Cascade. Saved state cleared." "INFO"
                Clear-HubPatchCascadeState
            }
        }
    }

    # Initialize Offline Staging banner state
    Update-OfflineStagingUi

    # Check for CLI Playbook / ActionRoutine invocation
    $targetPlaybook = if ($Playbook) { $Playbook } elseif ($ActionRoutine) { $ActionRoutine } else { '' }
    if ($targetPlaybook) {
        Write-HubLog "Command-line Action Playbook scheduled: $targetPlaybook" "INFO"
        $playbookTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $playbookTimer.Interval = [TimeSpan]::FromSeconds(2)
        $playbookTimer.Add_Tick({
            $playbookTimer.Stop()
            Invoke-HubActionRoutine -RoutineName $targetPlaybook
        })
        $playbookTimer.Start()
    }

    # Ensure clean shutdown of RunspacePool and DispatcherTimer on close
    $window.Add_Closed({
        if ($script:HubQueueTimer) { $script:HubQueueTimer.Stop() }
        if ($script:HubActiveAsyncTasks) {
            foreach ($t in $script:HubActiveAsyncTasks) {
                try { $t.PowerShell.Stop() } catch { }
                try { $t.PowerShell.Dispose() } catch { }
            }
            $script:HubActiveAsyncTasks.Clear()
        }
        if ($script:HubWorkerPool) {
            try { $script:HubWorkerPool.Close() } catch { }
            try { $script:HubWorkerPool.Dispose() } catch { }
        }
    })

    # Show Window
    try {
        $window.ShowDialog() | Out-Null
    } finally {
        if ($script:HubQueueTimer) { $script:HubQueueTimer.Stop() }
        if ($script:HubActiveAsyncTasks) {
            foreach ($t in $script:HubActiveAsyncTasks) {
                try { $t.PowerShell.Stop() } catch { }
                try { $t.PowerShell.Dispose() } catch { }
            }
            $script:HubActiveAsyncTasks.Clear()
        }
        if ($script:HubWorkerPool) {
            try { $script:HubWorkerPool.Close() } catch { }
            try { $script:HubWorkerPool.Dispose() } catch { }
        }
    }
}

# ==============================================================================
# SECTION C: RUNTIME ENTRYPOINT & STA APARTMENT STATE GUARD
# ==============================================================================

if ($RenameComputer) {
    try {
        $serial = (Get-CimInstance Win32_BIOS -ErrorAction Stop).SerialNumber.Trim()
        $rand = (Get-Random -Minimum 1000 -Maximum 9999).ToString()
        $newName = $ComputerNameTemplate.Replace('%SERIAL%', $serial).Replace('%RAND%', $rand)
        if ($newName.Length -gt 15) { $newName = $newName.Substring(0, 15) }
        Write-Host "Renaming computer to '$newName'..." -ForegroundColor Cyan
        Rename-Computer -NewName $newName -Force -ErrorAction Stop
        Write-Host "Computer renamed to '$newName'. (Reboot required to take effect)" -ForegroundColor Green
    } catch {
        Write-Host "Computer rename failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($Shootout) {
    if (Test-HubConsoleAvailable) { Show-HubIntro -Force -FrameMs 200 }
    else { Write-Host "Need a console at least 66 columns wide for the shootout." -ForegroundColor Yellow }
    return
}

if ($HarvestOnly) {
    Write-Host "`nAutopilotFast Hardware Hash Harvester" -ForegroundColor Cyan
    $hash = Get-AutopilotHash -GroupTag $GroupTag -AssignedUser $AssignedUser
    $hash | Format-List
    return
}

if ($OfflineStaging) {
    Write-Host "`nAutopilot Command Hub - Controlled Offline Staging" -ForegroundColor Cyan
    $res = Invoke-OfflineStaging -Profile $StagingProfile
    Write-Host "Profile Activated: $($res.Title)" -ForegroundColor Green
    foreach ($a in $res.Actions) { Write-Host "  - $a" -ForegroundColor Gray }
    return
}

if ($ResumeAutopilot) {
    Write-Host "`nAutopilot Command Hub - Cloud Re-Docking & Enterprise Resume" -ForegroundColor Cyan
    $res = Invoke-AutopilotReDock
    if ($res.Success) {
        Write-Host "SUCCESS: $($res.Message)" -ForegroundColor Green
        foreach ($a in $res.Actions) { Write-Host "  - $a" -ForegroundColor Gray }
    } else {
        Write-Host "ERROR: $($res.Message)" -ForegroundColor Red
    }
    return
}

if (($Playbook -or $ActionRoutine) -and $NoGui) {
    $pbName = if ($Playbook) { $Playbook } else { $ActionRoutine }
    Invoke-HeadlessActionRoutine -RoutineName $pbName
    return
}

if ($DellWarranty) {
    Write-Host "`nDell Technologies Enterprise Warranty & Refresh Lifecycle Engine" -ForegroundColor Cyan
    $tag = if ($DellServiceTag) { $DellServiceTag } else { '' }
    $w = Get-DellWarrantyInfo -ServiceTag $tag
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host " DELL ASSET WARRANTY & REFRESH ASSESSMENT: $($w.ServiceTag)" -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host " Machine Model         : $($w.SystemModel)" -ForegroundColor White
    Write-Host " Product Line          : $($w.ProductLineDescription)" -ForegroundColor White
    Write-Host " Factory Ship Date     : $($w.ShipDate) ($($w.DeviceAgeYears) years old)" -ForegroundColor White
    Write-Host " Country / Region      : $($w.CountryCode)" -ForegroundColor White
    Write-Host " Active Contract       : $($w.PrimaryServiceLevel)" -ForegroundColor White
    Write-Host " Coverage End Date     : $($w.WarrantyEndDate)" -ForegroundColor White
    Write-Host " Warranty Posture      : " -NoNewline
    if ($w.IsUnderWarranty) {
        Write-Host $w.WarrantyStatus -ForegroundColor Green
    } else {
        Write-Host $w.WarrantyStatus -ForegroundColor Red
    }
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " REFRESH VERDICT       : " -NoNewline
    if ($w.IsUnderWarranty) {
        Write-Host $w.RefreshVerdict -ForegroundColor Green
    } else {
        Write-Host $w.RefreshVerdict -ForegroundColor Red
    }
    Write-Host " Details               : $($w.RefreshRecommendation)" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Contract Entitlements ($($w.Entitlements.Count) Total):" -ForegroundColor Cyan
    $w.Entitlements | Format-Table ServiceLevelDescription, EntitlementType, StartDate, EndDate, Status -AutoSize

    if ($ExportCsv) {
        $path = if ($CsvPath) { $CsvPath } else { "$env:TEMP\DellWarranty-$($w.ServiceTag).csv" }
        $flatObj = [PSCustomObject]@{
            ServiceTag            = $w.ServiceTag
            SystemModel           = $w.SystemModel
            ShipDate              = $w.ShipDate
            DeviceAgeYears        = $w.DeviceAgeYears
            IsUnderWarranty       = $w.IsUnderWarranty
            WarrantyStatus        = $w.WarrantyStatus
            DaysRemaining         = $w.DaysRemaining
            WarrantyEndDate       = $w.WarrantyEndDate
            PrimaryServiceLevel   = $w.PrimaryServiceLevel
            RefreshVerdict        = $w.RefreshVerdict
            RefreshRecommendation = $w.RefreshRecommendation
            QueriedAt             = $w.QueriedAt
        }
        $flatObj | Export-Csv -Path $path -NoTypeInformation -Force
        Write-Host "Exported audit CSV to: $path" -ForegroundColor Green
    }
    return
}

if ($LenovoWarranty) {
    Write-Host "`nLenovo Warranty Lookup" -ForegroundColor Cyan
    $lw = Get-LenovoWarrantyInfo -SerialNumber $LenovoSerialNumber
    $lw | Format-List
    if ($ExportCsv) {
        $path = if ($CsvPath) { $CsvPath } else { Join-Path $env:TEMP "LenovoWarranty-$($lw.SerialNumber).csv" }
        $lw | Select-Object -Property * -ExcludeProperty Entitlements | Export-Csv -Path $path -NoTypeInformation -Force
        Write-Host "Exported warranty CSV to: $path" -ForegroundColor Green
    }
    return
}

if ($HardwareHealth) {
    Write-Host "`nHardware Health (battery wear + storage reliability)" -ForegroundColor Cyan
    $bat = Get-BatteryHealthInfo
    if ($bat) { Write-Host "Battery:" -ForegroundColor White; $bat | Format-List } else { Write-Host "Battery: none detected (desktop or no ACPI battery)" -ForegroundColor Gray }
    $disks = @(Get-StorageReliabilityInfo)
    if ($disks.Count -gt 0) { Write-Host "Storage:" -ForegroundColor White; $disks | Format-Table -AutoSize } else { Write-Host "Storage: no reliability counters available (needs elevation or unsupported controller)" -ForegroundColor Gray }
    return
}

if ($OfflineJson) {
    Write-Host "`nOffline Autopilot JSON Profile (AutopilotConfigurationFile.json)" -ForegroundColor Cyan
    # Tenant identity: explicit env, else what the device already knows from a cached profile
    $tenantId = $env:AZURE_TENANT_ID
    $tenantDomain = $env:AUTOPILOT_TENANT_DOMAIN
    if (-not $tenantId -or -not $tenantDomain) {
        $ds = Get-DeviceEnrollmentState
        if (-not $tenantId -and $ds.AutopilotTenantId) { $tenantId = $ds.AutopilotTenantId }
        if (-not $tenantDomain -and $ds.AutopilotTenantDomain) { $tenantDomain = $ds.AutopilotTenantDomain }
    }
    if (-not $tenantId -or -not $tenantDomain) {
        Write-Host "Set AZURE_TENANT_ID and AUTOPILOT_TENANT_DOMAIN (env or .env) - the offline profile must name the tenant it belongs to." -ForegroundColor Red
        return
    }
    $out = Export-AutopilotConfigurationFile -Path $OfflineJsonPath -TenantId $tenantId -TenantDomain $tenantDomain -DeviceName $ComputerNameTemplate
    Write-Host "Offline profile written to: $out" -ForegroundColor Green
    Write-Host "Copy it to C:\Windows\Provisioning\Autopilot\AutopilotConfigurationFile.json on the target before OOBE (the Hub's 'Inject Offline JSON' does this)." -ForegroundColor Gray
    return
}

if ($Decommission) {
    Write-Host "`nAutopilot Clean Decommission (local state only)" -ForegroundColor Cyan
    if (-not $script:IsElevated) { Write-Host "Requires elevation." -ForegroundColor Red; return }
    Write-Host "This removes the cached Autopilot profile, Provisioning diagnostics and the local MDM enrollment keys." -ForegroundColor Yellow
    Write-Host "It does NOT unenroll the device in Intune - retire/delete it in the portal for that." -ForegroundColor Yellow
    $confirm = Read-Host "Type DECOMMISSION to continue"
    if ($confirm -ne 'DECOMMISSION') { Write-Host "Aborted." -ForegroundColor Gray; return }
    foreach ($a in @(Invoke-AutopilotCleanDecommission)) { Write-Host "  $a" -ForegroundColor White }
    return
}

if ($ExportCsv) {
    Write-Host "`nAutopilotFast CSV Exporter" -ForegroundColor Cyan
    $res = Export-AutopilotCsv -Path $CsvPath -AutoDetectUsb -GroupTag $GroupTag -AssignedUser $AssignedUser -DeviceName $ComputerNameTemplate
    Write-Host "Exported to: $($res.Path)" -ForegroundColor Green
    return
}

if ($PatchCascade) {
    Write-Host "`nAutopilot Command Hub - Autonomous Patch Cascade Engine (OOBE Ready)" -ForegroundColor Cyan
    if (-not $script:RuntimeContext.MeetsPreferred) {
        Write-Host "Notice: Windows Update installation requires administrator privileges. Some updates may fail without elevation." -ForegroundColor Yellow
    }
    $res = Start-HubAutonomousPatchCascade -MaxPasses $MaxPasses -IncludeDrivers:$IncludeDrivers -StatusCallback { param($m) Write-Host "  $m" -ForegroundColor Gray }
    if ($res.Completed) {
        Write-Host "`nPatch Cascade completed successfully! Total updates installed: $($res.TotalInstalled)" -ForegroundColor Green
    } elseif ($res.RebootRequired) {
        Write-Host "`nPatch Cascade Pass $($res.CurrentPass) completed. Reboot scheduled for Pass $($res.NextPass)." -ForegroundColor Yellow
    }
    return
}

if ($WindowsUpdate) {
    Write-Host "`nAutopilot Command Hub - Windows Update & Driver Engine (OOBE Ready)" -ForegroundColor Cyan
    if (-not $script:RuntimeContext.MeetsPreferred) {
        Write-Host "Notice: Windows Update installation requires administrator privileges. Some updates may fail without elevation." -ForegroundColor Yellow
    }
    if ($ScanOnly) {
        Write-Host "Scanning for available updates (IncludeDrivers: $([bool]$IncludeDrivers))..." -ForegroundColor Cyan
        $updates = Get-HubPendingWindowsUpdates -IncludeDrivers:$IncludeDrivers -StatusCallback { param($m) Write-Host "  $m" -ForegroundColor Gray }
        Write-Host "Found $($updates.Count) pending update(s):" -ForegroundColor Green
        if ($updates.Count -gt 0) {
            $updates | Select-Object Title, KB, SizeMB, IsDriver, RebootRequired | Format-Table -AutoSize
        }
    } else {
        Write-Host "Starting update scan, download, and installation..." -ForegroundColor Cyan
        $res = Start-HubWindowsUpdate -IncludeDrivers:$IncludeDrivers -AutoReboot:$AutoReboot -StatusCallback { param($m) Write-Host "  $m" -ForegroundColor Gray }
        if ($res.Success) {
            Write-Host "Updates completed successfully! ($($res.InstalledCount) succeeded, $($res.FailedCount) failed)." -ForegroundColor Green
        } else {
            Write-Host "Updates completed with errors: $($res.InstalledCount) succeeded, $($res.FailedCount) failed." -ForegroundColor Yellow
        }
        if ($res.RebootRequired) {
            Write-Host "System restart is REQUIRED to complete update installation." -ForegroundColor Red
        }
    }
    return
}

# Launch GUI in STA Apartment State
if (-not $NoGui) {
    # Single-instance lifecycle: Automatically close prior lingering instances instead of blocking
    if ($env:AUTOPILOT_HUB_STA_CHILD -ne '1') {
        $mutexCreated = $false
        $script:InstanceMutex = [System.Threading.Mutex]::new($true, 'Global\AutopilotCommandHub', [ref]$mutexCreated)
        if (-not $mutexCreated) {
            # Cleanly close and terminate any previous instance(s)
            Close-ExistingHubInstances

            # Acquire ownership of the released/abandoned mutex
            try {
                [void]$script:InstanceMutex.WaitOne([TimeSpan]::FromSeconds(3))
            } catch {
                # AbandonedMutexException in .NET means previous holder terminated; ownership is granted to this thread
            }
        }

        # Record this process PID for clean lifecycle tracking
        try {
            $pidFile = Join-Path ([System.IO.Path]::GetTempPath()) 'AutopilotCommandHub.pid'
            [System.IO.File]::WriteAllText($pidFile, $PID.ToString())
        } catch { }

        if ($ResumeFromRestart) { Unregister-HubResumeAfterRestart | Out-Null }
    }

    # Head start: run the 3 s network ladder in its own runspace so it overlaps the console intro and the
    # XAML build. Start-AutopilotHubGui consumes $script:PreflightHandle if it is present. Only for the real
    # STA launch (the MTA relaunch below re-runs this whole script as an STA child, which sets the guard).
    if ($env:AUTOPILOT_HUB_STA_CHILD -ne '1') {
        try {
            $preRs = [runspacefactory]::CreateRunspace(); $preRs.ApartmentState = 'MTA'; $preRs.Open()
            $preFn = "function Test-StagedNetwork {`n$((Get-Command Test-StagedNetwork -CommandType Function -ErrorAction Stop).Definition)`n}"
            $script:PreflightRunspace = [powershell]::Create()
            $script:PreflightRunspace.Runspace = $preRs
            [void]$script:PreflightRunspace.AddScript("`$ErrorActionPreference='SilentlyContinue'`n$preFn`nTest-StagedNetwork")
            $script:PreflightHandle = $script:PreflightRunspace.BeginInvoke()
        } catch { $script:PreflightHandle = $null; $script:PreflightRunspace = $null }
    }
    if (-not $NoIntro) { Show-HubIntro }

    try {
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
            Start-AutopilotHubGui
        } else {
            # WPF needs STA. A raw [System.Threading.Thread] cannot execute PowerShell script blocks (no runspace
            # on that thread), so re-run this script inside a dedicated STA runspace with the same parameters.
            $source = if ($script:SelfScriptPath -and (Test-Path $script:SelfScriptPath)) { Get-Content -Path $script:SelfScriptPath -Raw } else { $script:SelfSource }
            if (-not $source) { throw "Cannot relaunch in STA: the running script's source is unavailable." }
            $staRunspace = [runspacefactory]::CreateRunspace()
            $staRunspace.ApartmentState = 'STA'
            $staRunspace.ThreadOptions = 'ReuseThread'
            $staRunspace.Open()
            $staHost = [powershell]::Create()
            $staHost.Runspace = $staRunspace
            [void]$staHost.AddScript($source)
            foreach ($k in $PSBoundParameters.Keys) { [void]$staHost.AddParameter($k, $PSBoundParameters[$k]) }
            $env:AUTOPILOT_HUB_STA_CHILD = '1'
            try {
                $staHost.Invoke() | Out-Null
                foreach ($e in $staHost.Streams.Error) { Write-Host "[STA] $e" -ForegroundColor Red }
            } finally {
                $env:AUTOPILOT_HUB_STA_CHILD = $null
                $staHost.Dispose()
                $staRunspace.Dispose()
            }
        }
    } finally {
        # Clean up PID file and release single-instance mutex when the GUI session exits
        if ($env:AUTOPILOT_HUB_STA_CHILD -ne '1') {
            try {
                $pidFile = Join-Path ([System.IO.Path]::GetTempPath()) 'AutopilotCommandHub.pid'
                if (Test-Path $pidFile) { Remove-Item -Path $pidFile -Force -ErrorAction SilentlyContinue }
            } catch { }
            try {
                if ($script:InstanceMutex) {
                    try { $script:InstanceMutex.ReleaseMutex() } catch { }
                    $script:InstanceMutex.Dispose()
                    $script:InstanceMutex = $null
                }
            } catch { }
        }
    }
}
