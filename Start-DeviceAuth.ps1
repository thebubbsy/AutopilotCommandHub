# Start-DeviceAuth.ps1
# Interactive Microsoft Graph & Intune Device Code Flow Authenticator
[CmdletBinding()]
param(
    [string]$TenantId = 'organizations',
    [string]$ClientId = '1950a258-227b-4e31-a9cf-717495945fc2'
)

$ErrorActionPreference = 'Stop'

function Set-ClipboardSafe([string]$text) {
    try {
        Set-Clipboard -Value $text -ErrorAction Stop
    } catch {
        try {
            Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue
            [System.Windows.Clipboard]::SetText($text)
        } catch { }
    }
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " INITIALIZING MICROSOFT GRAPH & INTUNE DEVICE AUTH ENGINE " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Initialize embedded core engine
$hubScript = Join-Path $PSScriptRoot "autopilot.ps1"
if (-not (Test-Path $hubScript)) {
    throw "autopilot.ps1 not found in $PSScriptRoot"
}
. $hubScript -NoGui
Write-Host "[1/4] Embedded IntuneShared & Autopilot modules initialized." -ForegroundColor Green

# 2. Copy Device Login URL to clipboard first
$loginUrl = "https://login.microsoft.com/device"
Set-ClipboardSafe $loginUrl
Write-Host "[2/4] Device Login URL copied to clipboard: $loginUrl" -ForegroundColor Green

# 3. Request AAD Device Code
Write-Host "[3/4] Requesting AAD Device Code from Microsoft Entra ID..." -ForegroundColor Yellow
$dc = Connect-GraphToken -TenantId $TenantId -ClientId $ClientId -InteractiveDeviceCode

if (-not $dc -or -not $dc.UserCode) {
    throw "Failed to obtain device authorization code from Entra ID."
}

# Copy the authorization User Code to clipboard so user can immediately paste
Set-ClipboardSafe $dc.UserCode

Write-Host "`n----------------------------------------------------------" -ForegroundColor DarkCyan
Write-Host "  DEVICE CODE:       " -NoNewline
Write-Host " $($dc.UserCode) " -ForegroundColor Black -BackgroundColor Cyan
Write-Host "  VERIFICATION URL:  " -NoNewline
Write-Host "$($dc.VerificationUrl)" -ForegroundColor Yellow
Write-Host "  EXPIRES IN:        $($dc.ExpiresIn) seconds ($([Math]::Round($dc.ExpiresIn/60, 1)) minutes)" -ForegroundColor Gray
Write-Host "  STATUS:            Code copied to clipboard! Ready to paste." -ForegroundColor Green
Write-Host "----------------------------------------------------------`n" -ForegroundColor DarkCyan

# Launch browser automatically
try {
    Start-Process $dc.VerificationUrl
    Write-Host "Browser launched to $($dc.VerificationUrl)." -ForegroundColor Green
} catch {
    Write-Warning "Could not launch browser automatically. Please navigate to $($dc.VerificationUrl) manually."
}

# 4. Handle and Poll for Graph API Token
Write-Host "`n[4/4] Listening for browser authorization from Microsoft Identity Platform..." -ForegroundColor Yellow
Write-Host "Paste the code '$($dc.UserCode)' into your browser and approve permissions.`n" -ForegroundColor Cyan

$interval = if ($dc.Interval) { [Math]::Max(3, [int]$dc.Interval) } else { 5 }
$maxAttempts = [Math]::Floor($dc.ExpiresIn / $interval)
$attempt = 0
$token = $null

while ($attempt -lt $maxAttempts) {
    $attempt++
    Start-Sleep -Seconds $interval
    try {
        $token = Poll-GraphDeviceCodeToken -DeviceCodeContext $dc
        if ($token) {
            Write-Host "`n`n==========================================================" -ForegroundColor Green
            Write-Host " AUTHENTICATED SUCCESSFULLY TO MICROSOFT GRAPH & INTUNE! " -ForegroundColor Green
            Write-Host "==========================================================" -ForegroundColor Green
            Write-Host "Tenant ID:     $($script:GraphAuthContext.TenantId)" -ForegroundColor Cyan
            Write-Host "Token Expiry:  $($script:GraphAuthContext.ExpiresOn) UTC" -ForegroundColor Cyan
            
            # Save token to persistent cache for session reuse across tools
            $tokenCache = Join-Path $env:TEMP 'GraphTokenCache.json'
            $script:GraphAuthContext | ConvertTo-Json | Set-Content -Path $tokenCache -Force
            Write-Host "Token Cached:  $tokenCache" -ForegroundColor DarkGray
            Write-Host "Graph API session is active and ready for requests.`n" -ForegroundColor Green
            break
        }
        Write-Host -NoNewline "."
    } catch {
        Write-Host "`n"
        Write-Error "Device authentication failed: $($_.Exception.Message)"
        exit 1
    }
}

if (-not $token) {
    Write-Warning "`nTimed out waiting for device authorization."
}
