<#
.SYNOPSIS
    Autopilot Command Hub - Fleet Plane Telemetry Ingestion Server & Dashboard
.DESCRIPTION
    Lightweight, self-contained HTTP server that ingests deployment telemetry JSON payloads
    from Autopilot Command Hub benches, computes cross-fleet cohort baselines, flags statistical
    outliers, and serves a real-time Fluent Dark fleet dashboard.

    Endpoints:
    - POST /api/telemetry  : Ingest deployment metrics & signed provisioning receipts
    - GET  /api/deployments: Query all captured fleet deployments
    - GET  /api/cohorts    : Query cross-fleet cohort analytics by hardware model
    - GET  /api/outliers   : Query statistical outlier alerts across active deployments
    - GET  /api/health     : Health probe and deployments counter
    - GET  /               : Interactive Single-Page Application (SPA) Fleet Dashboard

    Invocation:
        .\Start-HubFleetServer.ps1 -Port 8443
        .\Start-HubFleetServer.ps1 -Port 8080 -DataDir "D:\FleetData" -OpenBrowser

    (c) 2026 Matthew Bubb (thebubbsy) | https://onyachamp.com
#>

[CmdletBinding()]
param(
    [int]$Port = 8443,
    [string]$DataDir = 'C:\AutopilotLogs\FleetData',
    # Interface to bind. Default 'localhost' is safe and needs no admin, but ONLY accepts connections from
    # this machine. For real multi-bench ingestion, pass the server's LAN IP or '+' (all interfaces) -
    # that needs an elevated shell or a one-time 'netsh http add urlacl', and you should set -AuthToken too.
    [string]$BindAddress = 'localhost',
    # Shared secret required in the X-Hub-Token header on every request when set (else defaults to
    # $env:HUB_FLEET_TOKEN). Strongly recommended whenever BindAddress is not localhost.
    [string]$AuthToken = $env:HUB_FLEET_TOKEN,
    [switch]$OpenBrowser
)

$ErrorActionPreference = 'Stop'

function Get-ScriptEncoding {
    return [System.Text.UTF8Encoding]::new($false)
}

function ConvertTo-SvgSafe {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
}

function Get-HubCohortBaseline {
    param([array]$Records, [string]$Model = '')
    $matching = if ($Model) {
        @($Records | Where-Object { ($_.Device -and $_.Device.Model -like "*$Model*") -or ($_.Baseline -and $_.Baseline.Model -like "*$Model*") })
    } else {
        @($Records)
    }
    if ($matching.Count -eq 0) { return $null }

    $durations = [System.Collections.Generic.List[double]]::new()
    $downloads = [System.Collections.Generic.List[double]]::new()
    $iopsArr   = [System.Collections.Generic.List[double]]::new()

    foreach ($r in $matching) {
        if ($r.TotalSeconds) { $durations.Add([double]$r.TotalSeconds) }
        elseif ($r.Metrics -and $r.Metrics.TotalSeconds) { $durations.Add([double]$r.Metrics.TotalSeconds) }

        if ($r.Metrics -and $r.Metrics.Wu -and $r.Metrics.Wu.AvgDownloadMBps) {
            $downloads.Add([double]$r.Metrics.Wu.AvgDownloadMBps)
        }
        if ($r.Metrics -and $r.Metrics.Disk -and $r.Metrics.Disk.AvgIops) {
            $iopsArr.Add([double]$r.Metrics.Disk.AvgIops)
        }
    }

    function _calcStats([System.Collections.Generic.List[double]]$nums) {
        if (-not $nums -or $nums.Count -eq 0) { return [PSCustomObject]@{ N = 0; Mean = 0; StdDev = 0; Median = 0; P5 = 0; P95 = 0 } }
        $n = $nums.Count
        $sorted = @($nums | Sort-Object)
        $sum = 0
        foreach ($x in $sorted) { $sum += $x }
        $mean = $sum / [double]$n
        $variance = 0.0
        if ($n -gt 1) {
            $sumSq = 0.0
            foreach ($x in $sorted) { $sumSq += [math]::Pow($x - $mean, 2) }
            $variance = $sumSq / [double]($n - 1)
        }
        $stdDev = [math]::Sqrt($variance)
        $medIdx = [int][math]::Floor($n / 2)
        $median = if ($n % 2 -eq 0 -and $medIdx -gt 0) { ($sorted[$medIdx - 1] + $sorted[$medIdx]) / 2.0 } else { $sorted[$medIdx] }
        $p5Idx = [math]::Max(0, [int][math]::Floor($n * 0.05))
        $p95Idx = [math]::Min($n - 1, [int][math]::Floor($n * 0.95))
        return [PSCustomObject]@{
            N      = $n
            Mean   = [math]::Round($mean, 1)
            StdDev = [math]::Round($stdDev, 1)
            Median = [math]::Round($median, 1)
            P5     = [math]::Round($sorted[$p5Idx], 1)
            P95    = [math]::Round($sorted[$p95Idx], 1)
        }
    }

    return [PSCustomObject]@{
        Model         = if ($Model) { $Model } else { 'All Models' }
        SampleCount   = $matching.Count
        DurationStats = (_calcStats $durations)
        DownloadStats = (_calcStats $downloads)
        DiskIopsStats = (_calcStats $iopsArr)
        GeneratedUtc  = [datetime]::UtcNow.ToString('o')
    }
}

function Test-HubCohortOutlier {
    param($Metrics, $CohortBaseline)
    if (-not $CohortBaseline -or $CohortBaseline.DurationStats.N -lt 3) { return @() }
    $outliers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $dur = 0.0
    if ($Metrics.TotalSeconds) { $dur = [double]$Metrics.TotalSeconds }
    elseif ($Metrics.Metrics -and $Metrics.Metrics.TotalSeconds) { $dur = [double]$Metrics.Metrics.TotalSeconds }
    $dStats = $CohortBaseline.DurationStats

    if ($dur -gt 0 -and $dStats.StdDev -gt 0) {
        $zScore = ($dur - $dStats.Mean) / $dStats.StdDev
        $zThresh = if ($dStats.N -le 8) { 1.65 } else { 2.0 }
        if ($zScore -gt $zThresh) {
            $outliers.Add([PSCustomObject]@{
                Metric   = 'Total Deployment Duration'
                Severity = 'ALERT'
                Value    = "$([math]::Round($dur, 0)) s"
                Expected = "$($dStats.Mean) s (+/- $($dStats.StdDev) s)"
                ZScore   = [math]::Round($zScore, 2)
                Message  = "Unit took $([math]::Round($dur,0))s (+$( [math]::Round($zScore, 1) ) sigma vs $($dStats.Mean)s cohort average) - 98th percentile slow"
            })
        }
    }

    $dl = 0.0
    if ($Metrics.Wu -and $Metrics.Wu.AvgDownloadMBps) { $dl = [double]$Metrics.Wu.AvgDownloadMBps }
    elseif ($Metrics.Metrics -and $Metrics.Metrics.Wu -and $Metrics.Metrics.Wu.AvgDownloadMBps) { $dl = [double]$Metrics.Metrics.Wu.AvgDownloadMBps }
    if ($dl -gt 0 -and $CohortBaseline.DownloadStats.N -ge 3) {
        $dlStats = $CohortBaseline.DownloadStats
        if ($dlStats.StdDev -gt 0) {
            $zScoreDl = ($dl - $dlStats.Mean) / $dlStats.StdDev
            $zDlThresh = if ($dlStats.N -le 8) { -1.5 } else { -1.8 }
            if ($zScoreDl -lt $zDlThresh) {
                $outliers.Add([PSCustomObject]@{
                    Metric   = 'WU Network Download Speed'
                    Severity = 'WARN'
                    Value    = "$([math]::Round($dl, 1)) MB/s"
                    Expected = "$($dlStats.Mean) MB/s"
                    ZScore   = [math]::Round($zScoreDl, 2)
                    Message  = "Download throughput $([math]::Round($dl,1)) MB/s is significantly below cohort average ($($dlStats.Mean) MB/s)"
                })
            }
        }
    }

    $diskIops = 0.0
    if ($Metrics.Disk -and $Metrics.Disk.AvgIops) { $diskIops = [double]$Metrics.Disk.AvgIops }
    elseif ($Metrics.Metrics -and $Metrics.Metrics.Disk -and $Metrics.Metrics.Disk.AvgIops) { $diskIops = [double]$Metrics.Metrics.Disk.AvgIops }
    if ($diskIops -gt 0 -and $CohortBaseline.DiskIopsStats -and $CohortBaseline.DiskIopsStats.N -ge 3) {
        $iStats = $CohortBaseline.DiskIopsStats
        if ($iStats.StdDev -gt 0) {
            $zScoreIops = ($diskIops - $iStats.Mean) / $iStats.StdDev
            $zIopsThresh = if ($iStats.N -le 8) { -1.5 } else { -1.8 }
            if ($zScoreIops -lt $zIopsThresh) {
                $outliers.Add([PSCustomObject]@{
                    Metric   = 'Storage Disk IOPS'
                    Severity = 'WARN'
                    Value    = "$([math]::Round($diskIops, 0)) IOPS"
                    Expected = "$($iStats.Mean) IOPS"
                    ZScore   = [math]::Round($zScoreIops, 2)
                    Message  = "Storage IOPS $([math]::Round($diskIops,0)) is significantly degraded below cohort average ($($iStats.Mean) IOPS)"
                })
            }
        }
    }

    return @($outliers)
}

try {
    if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
    $depDir = Join-Path $DataDir 'deployments'
    if (-not (Test-Path $depDir)) { New-Item -ItemType Directory -Path $depDir -Force | Out-Null }
} catch {
    $DataDir = Join-Path $env:TEMP 'AutopilotFleetData'
    $depDir = Join-Path $DataDir 'deployments'
    if (-not (Test-Path $depDir)) { New-Item -ItemType Directory -Path $depDir -Force | Out-Null }
}

$listener = [System.Net.HttpListener]::new()
$prefix = "http://${BindAddress}:${Port}/"
$listener.Prefixes.Add($prefix)
if ($BindAddress -ne 'localhost' -and $BindAddress -ne '127.0.0.1') {
    if ([string]::IsNullOrWhiteSpace($AuthToken)) {
        Write-Host "[WARN] Server is bound to '$BindAddress' (network-reachable) with NO -AuthToken. Anyone who can reach this port can post telemetry and read the dashboard. Set -AuthToken or `$env:HUB_FLEET_TOKEN." -ForegroundColor Yellow
    }
}

try {
    $listener.Start()
} catch {
    Write-Host "[ERROR] Could not bind Fleet Ingestion Server to $prefix : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try specifying an alternate port: .\Start-HubFleetServer.ps1 -Port 8088" -ForegroundColor Yellow
    exit 1
}

Write-Host "`n==========================================================================" -ForegroundColor Cyan
Write-Host " AUTOPILOT FLEET PLANE INGESTION SERVER ACTIVE" -ForegroundColor Green
Write-Host " Listening URL:    $prefix" -ForegroundColor White
Write-Host " Ingestion API:    POST ${prefix}api/telemetry" -ForegroundColor White
Write-Host " Fleet Dashboard:  GET  $prefix" -ForegroundColor White
Write-Host " Storage Root:     $DataDir" -ForegroundColor White
Write-Host " Press CTRL+C to terminate server." -ForegroundColor Yellow
Write-Host "==========================================================================`n" -ForegroundColor Cyan

if ($OpenBrowser) {
    try { Start-Process $prefix } catch { }
}

function Build-FleetDashboardHtml {
    param([array]$Deployments)

    $total = $Deployments.Count
    $verified = @($Deployments | Where-Object { $_.Status -eq 'Success' -or ($_.Receipt -and $_.Receipt.Verdict -eq 'VERIFIED') }).Count
    $passRate = if ($total -gt 0) { [math]::Round(($verified / [double]$total) * 100, 1) } else { 100.0 }

    $durations = [System.Collections.Generic.List[double]]::new()
    $modelMap = @{}
    $siteMap = @{}

    foreach ($d in $Deployments) {
        $sec = 0.0
        if ($d.TotalSeconds) { $sec = [double]$d.TotalSeconds }
        elseif ($d.Metrics -and $d.Metrics.TotalSeconds) { $sec = [double]$d.Metrics.TotalSeconds }
        if ($sec -gt 0) { $durations.Add($sec) }

        $mName = if ($d.Device -and $d.Device.Model) { [string]$d.Device.Model } elseif ($d.Baseline -and $d.Baseline.Model) { [string]$d.Baseline.Model } else { 'Unknown Model' }
        if (-not $modelMap.ContainsKey($mName)) { $modelMap[$mName] = [System.Collections.Generic.List[object]]::new() }
        $modelMap[$mName].Add($d)

        $sName = if ($d.Site) { [string]$d.Site } else { 'Default' }
        if (-not $siteMap.ContainsKey($sName)) { $siteMap[$sName] = [System.Collections.Generic.List[object]]::new() }
        $siteMap[$sName].Add($d)
    }

    $sortedDur = @($durations | Sort-Object)
    $medDur = if ($sortedDur.Count -gt 0) { [math]::Round($sortedDur[[int][math]::Floor($sortedDur.Count / 2)], 0) } else { 0 }

    # Cohort breakdown
    $cohortRows = ''
    $cohortMap = @{}
    foreach ($mk in $modelMap.Keys) {
        $units = $modelMap[$mk]
        $cBase = Get-HubCohortBaseline -Records $units -Model $mk
        $cohortMap[$mk] = $cBase
        $med = if ($cBase -and $cBase.DurationStats) { "$($cBase.DurationStats.Median) s" } else { '-' }
        $meanStd = if ($cBase -and $cBase.DurationStats) { "$($cBase.DurationStats.Mean) s (+/- $($cBase.DurationStats.StdDev)s)" } else { '-' }
        $avgDl = if ($cBase -and $cBase.DownloadStats -and $cBase.DownloadStats.Mean) { "$($cBase.DownloadStats.Mean) MB/s" } else { '-' }
        $cohortRows += "<tr><td class='font-bold'>$(ConvertTo-SvgSafe $mk)</td><td>$($units.Count)</td><td>$med</td><td>$meanStd</td><td>$avgDl</td><td><span class='badge-pass'>BENCHMARK OK</span></td></tr>`n"
    }

    # Evaluate Outliers across all deployments
    $allOutliers = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $Deployments) {
        $mName = if ($d.Device -and $d.Device.Model) { [string]$d.Device.Model } elseif ($d.Baseline -and $d.Baseline.Model) { [string]$d.Baseline.Model } else { 'Unknown Model' }
        $cohort = if ($cohortMap.ContainsKey($mName)) { $cohortMap[$mName] } else { $null }
        if ($cohort) {
            $outs = Test-HubCohortOutlier -Metrics $d -CohortBaseline $cohort
            foreach ($o in $outs) {
                $allOutliers.Add([PSCustomObject]@{
                    DeploymentId = if ($d.DeploymentId) { $d.DeploymentId } else { 'N/A' }
                    SerialNumber = if ($d.Device -and $d.Device.SerialNumber) { $d.Device.SerialNumber } else { 'N/A' }
                    Model        = $mName
                    Metric       = $o.Metric
                    Severity     = $o.Severity
                    Value        = $o.Value
                    Expected     = $o.Expected
                    ZScore       = $o.ZScore
                    Message      = $o.Message
                })
            }
        }
    }

    $outlierRows = ''
    foreach ($oa in $allOutliers) {
        $bClass = if ($oa.Severity -eq 'ALERT') { 'badge-fail' } else { 'badge-warn' }
        $outlierRows += "<tr><td class='mono'>$(ConvertTo-SvgSafe $oa.DeploymentId)</td><td class='mono'>$(ConvertTo-SvgSafe $oa.SerialNumber)</td><td>$(ConvertTo-SvgSafe $oa.Model)</td><td><span class='$bClass'>$($oa.Severity)</span></td><td>$($oa.Metric)</td><td>$($oa.Value)</td><td>$($oa.Expected)</td><td>Z=$($oa.ZScore)</td><td style='font-size:12px;'>$(ConvertTo-SvgSafe $oa.Message)</td></tr>`n"
    }

    $outlierSection = if ($allOutliers.Count -gt 0) {
        @"
<div class='card' style='border:1px solid var(--red);background:rgba(248,113,113,0.06);'>
<h2 style='color:var(--red);'>[!] Active Statistical Outliers &amp; Fleet Anomaly Alerts ($($allOutliers.Count))</h2>
<table><thead><tr><th>Deployment ID</th><th>Serial</th><th>Model</th><th>Severity</th><th>Metric</th><th>Observed</th><th>Cohort Average</th><th>Z-Score</th><th>Alert Details</th></tr></thead>
<tbody>$outlierRows</tbody></table></div>
"@
    } else {
        @"
<div class='card' style='border:1px solid rgba(74,222,128,0.3);background:rgba(74,222,128,0.05);'>
<h2 style='color:var(--green);'>[OK] Fleet Cohort Health: Normal Statistical Variance</h2>
<div style='color:var(--text);font-size:12.5px;'>All active units are provisioning within expected duration (&plusmn;2&sigma;) and network throughput thresholds.</div></div>
"@
    }

    # Site breakdown
    $siteRows = ''
    foreach ($sk in $siteMap.Keys) {
        $sUnits = $siteMap[$sk]
        $siteRows += "<tr><td class='font-bold'>$(ConvertTo-SvgSafe $sk)</td><td>$($sUnits.Count)</td><td><span class='badge-pass'>ACTIVE</span></td></tr>`n"
    }

    # Deployments list
    $deployRows = ''
    foreach ($d in ($Deployments | Select-Object -First 50)) {
        $depId = if ($d.DeploymentId) { $d.DeploymentId } else { '-' }
        $sn = if ($d.Device -and $d.Device.SerialNumber) { $d.Device.SerialNumber } else { '-' }
        $md = if ($d.Device -and $d.Device.Model) { $d.Device.Model } else { '-' }
        $site = if ($d.Site) { $d.Site } else { 'Default' }
        $tech = if ($d.Technician) { $d.Technician } else { '-' }
        $sec = if ($d.TotalSeconds) { [math]::Round([double]$d.TotalSeconds, 0) } elseif ($d.Metrics -and $d.Metrics.TotalSeconds) { [math]::Round([double]$d.Metrics.TotalSeconds, 0) } else { 0 }
        $status = if ($d.Status) { $d.Status } else { 'Success' }
        $badgeClass = if ($status -eq 'Success') { 'badge-pass' } else { 'badge-fail' }
        $deployRows += "<tr><td class='mono'>$(ConvertTo-SvgSafe $depId)</td><td class='mono'>$(ConvertTo-SvgSafe $sn)</td><td>$(ConvertTo-SvgSafe $md)</td><td>$(ConvertTo-SvgSafe $site)</td><td>$(ConvertTo-SvgSafe $tech)</td><td>${sec}s</td><td><span class='$badgeClass'>$status</span></td></tr>`n"
    }

    $genTime = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
    return @"
<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Autopilot Fleet Plane Dashboard</title>
<style>
:root{--bg:#0b1329;--surface:#152238;--surface2:#1c2d4a;--border:#2a4066;--text:#f1f5f9;--muted:#94a3b8;--accent:#38bdf8;--green:#4ade80;--amber:#fbbf24;--red:#f87171;}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:'Segoe UI Variable Text','Segoe UI',system-ui,sans-serif;font-size:13.5px}
.wrap{max-width:1250px;margin:0 auto;padding:24px}
header{display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--border);padding-bottom:16px;margin-bottom:20px}
header h1{margin:0;font-size:22px;letter-spacing:-0.5px}header .sub{color:var(--muted);font-size:12px;margin-top:3px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:24px}
.kpi-card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px 20px}
.kpi-val{font-size:28px;font-weight:700;color:var(--accent)}
.kpi-lbl{font-size:11px;color:var(--muted);text-transform:uppercase;margin-top:4px;font-weight:600}
.grid-2{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:24px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:20px;margin-bottom:24px}
.card h2{margin:0 0 14px;font-size:15px;border-bottom:1px solid var(--border);padding-bottom:8px}
table{width:100%;border-collapse:collapse;font-size:13px}
table th{text-align:left;padding:8px 10px;background:var(--surface2);color:var(--muted);font-size:11px;text-transform:uppercase;font-weight:600}
table td{padding:8px 10px;border-bottom:1px solid var(--border)}
.mono{font-family:Consolas,monospace;font-size:11.5px}
.font-bold{font-weight:600}
.badge-pass{background:rgba(74,222,128,0.15);color:var(--green);border:1px solid rgba(74,222,128,0.3);border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600}
.badge-warn{background:rgba(251,191,36,0.15);color:var(--amber);border:1px solid rgba(251,191,36,0.3);border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600}
.badge-fail{background:rgba(248,113,113,0.15);color:var(--red);border:1px solid rgba(248,113,113,0.3);border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600}
footer{text-align:center;color:var(--muted);font-size:11px;margin-top:24px}
</style></head><body><div class='wrap'>
<header>
<div><h1>Autopilot Fleet Plane Dashboard</h1><div class='sub'>Real-Time Fleet Provisioning Telemetry &middot; Generated $genTime</div></div>
<div class='mono' style='color:var(--accent);background:rgba(56,189,248,0.1);padding:4px 10px;border-radius:4px;'>[INGESTION LISTENER ACTIVE]</div>
</header>
<div class='kpi-row'>
<div class='kpi-card'><div class='kpi-val'>$total</div><div class='kpi-lbl'>Total Deployments</div></div>
<div class='kpi-card'><div class='kpi-val' style='color:var(--green);'>$passRate%</div><div class='kpi-lbl'>Fleet Success Rate</div></div>
<div class='kpi-card'><div class='kpi-val'>$($medDur)s</div><div class='kpi-lbl'>Median Provisioning Time</div></div>
<div class='kpi-card'><div class='kpi-val'>$($modelMap.Keys.Count)</div><div class='kpi-lbl'>Cohort Models Tracked</div></div>
<div class='kpi-card'><div class='kpi-val' style='color:$(if ($allOutliers.Count -gt 0) { "var(--red)" } else { "var(--green)" });'>$($allOutliers.Count)</div><div class='kpi-lbl'>Outlier Anomalies</div></div>
</div>
$outlierSection
<div class='grid-2'>
<div class='card'><h2>Model Cohort Baselines</h2>
<table><thead><tr><th>Hardware Model</th><th>Units (N)</th><th>Median Duration</th><th>Mean (+/- 1s)</th><th>Avg Download</th><th>Cohort Health</th></tr></thead>
<tbody>$cohortRows</tbody></table></div>
<div class='card'><h2>Site &amp; Bench Activity</h2>
<table><thead><tr><th>Provisioning Site</th><th>Deployments</th><th>Status</th></tr></thead>
<tbody>$siteRows</tbody></table></div>
</div>
<div class='card'><h2>Recent Fleet Deployments Feed</h2>
<table><thead><tr><th>Deployment ID</th><th>Serial</th><th>Model</th><th>Site</th><th>Tech</th><th>Duration</th><th>Status</th></tr></thead>
<tbody>$deployRows</tbody></table></div>
<footer>Autopilot Command Hub &middot; https://onyachamp.com</footer>
</div></body></html>
"@
}

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response

        $path = $req.Url.AbsolutePath
        $method = $req.HttpMethod

        if (-not [string]::IsNullOrWhiteSpace($AuthToken)) {
            $provided = $req.Headers['X-Hub-Token']
            if ($provided -ne $AuthToken) {
                $den = [System.Text.Encoding]::UTF8.GetBytes('{"error":"unauthorized"}')
                $res.StatusCode = 401
                $res.ContentType = 'application/json'
                $res.ContentLength64 = $den.Length
                $res.OutputStream.Write($den, 0, $den.Length)
                $res.Close()
                Write-Host "[AUTH] Rejected $method $path from $($req.RemoteEndPoint) (bad/missing X-Hub-Token)" -ForegroundColor DarkYellow
                continue
            }
        }

        if ($method -eq 'POST' -and $path -match '^/api/(telemetry|ingest)') {
            $reader = [System.IO.StreamReader]::new($req.InputStream, [System.Text.Encoding]::UTF8)
            $body = $reader.ReadToEnd()
            $reader.Close()

            $parsed = $null
            try { $parsed = $body | ConvertFrom-Json } catch { }
            $depId = if ($parsed -and $parsed.DeploymentId) { $parsed.DeploymentId } else { "DEP-$([guid]::NewGuid().ToString('N').Substring(0,8))" }
            $savePath = Join-Path $depDir "Deployment_${depId}.json"
            [System.IO.File]::WriteAllText($savePath, $body, (Get-ScriptEncoding))

            $respData = [PSCustomObject]@{ status = 'ok'; deploymentId = $depId; receivedUtc = [datetime]::UtcNow.ToString('o') }
            $buf = [System.Text.Encoding]::UTF8.GetBytes(($respData | ConvertTo-Json))
            $res.ContentType = 'application/json'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
            Write-Host "[TELEMETRY INGESTED] Deployment: $depId" -ForegroundColor Green
        } elseif ($method -eq 'GET' -and $path -eq '/api/deployments') {
            $files = Get-ChildItem -Path $depDir -Filter "Deployment_*.json" -ErrorAction SilentlyContinue
            $records = [System.Collections.Generic.List[object]]::new()
            if ($files) {
                foreach ($f in $files) {
                    try { $records.Add((Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json)) } catch { }
                }
            }
            $j = if ($records -and $records.Count -gt 0) { ($records | ConvertTo-Json -Depth 6) } else { '[]' }
            $buf = [System.Text.Encoding]::UTF8.GetBytes($j)
            $res.ContentType = 'application/json'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
        } elseif ($method -eq 'GET' -and $path -eq '/api/cohorts') {
            $files = Get-ChildItem -Path $depDir -Filter "Deployment_*.json" -ErrorAction SilentlyContinue
            $records = [System.Collections.Generic.List[object]]::new()
            if ($files) {
                foreach ($f in $files) {
                    try { $records.Add((Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json)) } catch { }
                }
            }
            $modelGroups = @{}
            foreach ($d in $records) {
                $mName = if ($d.Device -and $d.Device.Model) { [string]$d.Device.Model } elseif ($d.Baseline -and $d.Baseline.Model) { [string]$d.Baseline.Model } else { 'Generic Model' }
                if (-not $modelGroups.ContainsKey($mName)) { $modelGroups[$mName] = [System.Collections.Generic.List[object]]::new() }
                $modelGroups[$mName].Add($d)
            }
            $cohorts = [System.Collections.Generic.List[object]]::new()
            foreach ($mk in $modelGroups.Keys) {
                $cb = Get-HubCohortBaseline -Records $modelGroups[$mk] -Model $mk
                if ($cb) { $cohorts.Add($cb) }
            }
            $j = if ($cohorts -and $cohorts.Count -gt 0) { ($cohorts | ConvertTo-Json -Depth 5) } else { '[]' }
            $buf = [System.Text.Encoding]::UTF8.GetBytes($j)
            $res.ContentType = 'application/json'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
        } elseif ($method -eq 'GET' -and $path -eq '/api/outliers') {
            $files = Get-ChildItem -Path $depDir -Filter "Deployment_*.json" -ErrorAction SilentlyContinue
            $records = [System.Collections.Generic.List[object]]::new()
            if ($files) {
                foreach ($f in $files) {
                    try { $records.Add((Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json)) } catch { }
                }
            }
            $modelGroups = @{}
            foreach ($d in $records) {
                $mName = if ($d.Device -and $d.Device.Model) { [string]$d.Device.Model } elseif ($d.Baseline -and $d.Baseline.Model) { [string]$d.Baseline.Model } else { 'Generic Model' }
                if (-not $modelGroups.ContainsKey($mName)) { $modelGroups[$mName] = [System.Collections.Generic.List[object]]::new() }
                $modelGroups[$mName].Add($d)
            }
            $cohortMap = @{}
            foreach ($mk in $modelGroups.Keys) {
                $cohortMap[$mk] = Get-HubCohortBaseline -Records $modelGroups[$mk] -Model $mk
            }
            $outliers = [System.Collections.Generic.List[object]]::new()
            foreach ($d in $records) {
                $mName = if ($d.Device -and $d.Device.Model) { [string]$d.Device.Model } elseif ($d.Baseline -and $d.Baseline.Model) { [string]$d.Baseline.Model } else { 'Generic Model' }
                $cohort = if ($cohortMap.ContainsKey($mName)) { $cohortMap[$mName] } else { $null }
                if ($cohort) {
                    $outs = Test-HubCohortOutlier -Metrics $d -CohortBaseline $cohort
                    foreach ($o in $outs) {
                        $outliers.Add([PSCustomObject]@{
                            DeploymentId = if ($d.DeploymentId) { $d.DeploymentId } else { 'N/A' }
                            SerialNumber = if ($d.Device -and $d.Device.SerialNumber) { $d.Device.SerialNumber } else { 'N/A' }
                            Model        = $mName
                            Metric       = $o.Metric
                            Severity     = $o.Severity
                            Value        = $o.Value
                            Expected     = $o.Expected
                            ZScore       = $o.ZScore
                            Message      = $o.Message
                        })
                    }
                }
            }
            $j = if ($outliers -and $outliers.Count -gt 0) { ($outliers | ConvertTo-Json -Depth 5) } else { '[]' }
            $buf = [System.Text.Encoding]::UTF8.GetBytes($j)
            $res.ContentType = 'application/json'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
        } elseif ($method -eq 'GET' -and $path -eq '/api/health') {
            $files = Get-ChildItem -Path $depDir -Filter "Deployment_*.json" -ErrorAction SilentlyContinue
            $hObj = [PSCustomObject]@{ status = 'healthy'; uptime = 'active'; deploymentsCount = if ($files) { $files.Count } else { 0 } }
            $buf = [System.Text.Encoding]::UTF8.GetBytes(($hObj | ConvertTo-Json))
            $res.ContentType = 'application/json'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
        } else {
            # Serve Fleet Dashboard
            $files = Get-ChildItem -Path $depDir -Filter "Deployment_*.json" -ErrorAction SilentlyContinue
            $records = [System.Collections.Generic.List[object]]::new()
            if ($files) {
                foreach ($f in $files) {
                    try { $records.Add((Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json)) } catch { }
                }
            }
            $html = Build-FleetDashboardHtml -Deployments $records
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $res.ContentType = 'text/html; charset=utf-8'
            $res.StatusCode = 200
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.Close()
        }
    } catch {
        if (-not $listener.IsListening) { break }
    }
}
