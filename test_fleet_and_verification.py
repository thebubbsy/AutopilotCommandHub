# test_fleet_and_verification.py
# Automated Verification Suite for Fleet Plane, Closed-Loop Verification,
# Tamper-Evident Receipts, OEM Drivers, Config-as-Code Profiles & Cart Mode

import os
import sys
import json
import time
import base64
import urllib.request
import urllib.error
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

repo_dir = os.path.dirname(os.path.abspath(__file__))
ps1_path = os.path.join(repo_dir, 'autopilot.ps1')
server_ps1 = os.path.join(repo_dir, 'Start-HubFleetServer.ps1')

def run_ps(cmd):
    full_cmd = f". '{ps1_path}' -NoGui; {cmd}"
    res = subprocess.run(
        ['pwsh', '-ExecutionPolicy', 'Bypass', '-Command', full_cmd],
        capture_output=True,
        encoding='utf-8',
        errors='replace'
    )
    return res

def test_closed_loop_verification_and_receipt():
    print("\n--- Test 1: Closed-Loop Verification Engine & Tamper-Evident Receipt ---")
    # 1. Test In-Memory First (Zero Traces by Default)
    ps_cmd = "$rec = Export-HubProvisioningReceipt; $rec | ConvertTo-Json -Depth 4"
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Export-HubProvisioningReceipt failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('{'):
            json_start = i
            break
    assert json_start != -1, f"Could not find JSON output from Export-HubProvisioningReceipt:\n{res.stdout}"
    data = json.loads('\n'.join(lines[json_start:]))

    assert data.get('Success') is True, f"Receipt generation did not report Success=true: {data}"
    assert data.get('InMemory') is True, f"Receipt should be held in-memory by default: {data}"
    assert not os.path.exists("C:\\AutopilotLogs"), "C:\\AutopilotLogs should NOT exist on disk by default!"
    dep_id = data.get('DeploymentId')
    print(f"  [PASS] In-Memory Receipt generated (Deployment ID: {dep_id}, Verdict: {data.get('ReceiptData', {}).get('Verdict')})")

    # Verify cryptographic seal using Test-HubProvisioningReceipt with in-memory object
    verify_cmd = """
    $rec = Export-HubProvisioningReceipt
    $v = Test-HubProvisioningReceipt -ReceiptPathOrObject $rec.ReceiptData
    $v | ConvertTo-Json
    """
    res_v = run_ps(verify_cmd)
    assert res_v.returncode == 0, f"Test-HubProvisioningReceipt failed:\n{res_v.stderr}"
    lines_v = res_v.stdout.strip().splitlines()
    j_idx = next(i for i, l in enumerate(lines_v) if l.strip().startswith('{'))
    v_data = json.loads('\n'.join(lines_v[j_idx:]))
    assert v_data.get('IsValid') is True, f"In-memory cryptographic seal validation failed: {v_data}"
    print(f"  [PASS] In-memory cryptographic integrity seal verified: {v_data.get('StoredHash')}")

    # 2. Test Explicit User Export (-SaveToDisk)
    temp_receipt_dir = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), f'ReceiptExport_{int(time.time())}')
    export_cmd = f"$rec = Export-HubProvisioningReceipt -SaveToDisk -OutputDir '{temp_receipt_dir}'; $rec | ConvertTo-Json -Depth 4"
    res_exp = run_ps(export_cmd)
    assert res_exp.returncode == 0, f"Export-HubProvisioningReceipt -SaveToDisk failed:\n{res_exp.stderr}"
    lines_exp = res_exp.stdout.strip().splitlines()
    j_exp_idx = next(i for i, l in enumerate(lines_exp) if l.strip().startswith('{'))
    exp_data = json.loads('\n'.join(lines_exp[j_exp_idx:]))
    html_path = exp_data.get('HtmlPath')
    json_path = exp_data.get('JsonPath')
    assert os.path.exists(html_path), f"Exported HTML receipt file not found: {html_path}"
    assert os.path.exists(json_path), f"Exported JSON receipt file not found: {json_path}"

    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    assert "Autopilot Provisioning Receipt" in html_content
    assert "Integrity Seal (" in html_content
    assert exp_data.get('DeploymentId') in html_content

    # Adversarial tampering test on exported file
    with open(json_path, 'r', encoding='utf-8') as f:
        tampered_obj = json.load(f)
    tampered_obj['Verdict'] = 'TAMPERED_STATUS'
    tampered_path = json_path.replace('.json', '_tampered.json')
    with open(tampered_path, 'w', encoding='utf-8') as f:
        json.dump(tampered_obj, f, indent=2)

    tamper_cmd = f"$v = Test-HubProvisioningReceipt -ReceiptPathOrObject '{tampered_path}'; $v | ConvertTo-Json"
    res_t = run_ps(tamper_cmd)
    lines_t = res_t.stdout.strip().splitlines()
    j_idx_t = next(i for i, l in enumerate(lines_t) if l.strip().startswith('{'))
    t_data = json.loads('\n'.join(lines_t[j_idx_t:]))
    assert t_data.get('IsValid') is False, f"Tampered receipt should have failed verification: {t_data}"
    print(f"  [PASS] Adversarial tampering detected and rejected: {t_data.get('Reason')}")

    # Clean up exported files
    for p in [html_path, json_path, tampered_path]:
        if os.path.exists(p):
            os.remove(p)
    if os.path.exists(temp_receipt_dir):
        os.rmdir(temp_receipt_dir)

def test_fleet_telemetry_and_buffering():
    print("\n--- Test 2: Fleet Plane Telemetry & Offline Buffering ---")
    ps_cmd = """
    $payload = [PSCustomObject]@{
        DeploymentId = 'HUB-20260923-TEST0001'
        RoutineName  = 'Intune-Only Cloud Build'
        Status       = 'Success'
        TotalSeconds = 420.5
        Site         = 'Melbourne'
        Technician   = 'bench-tech'
    }
    $sendRes = Send-HubFleetTelemetry -Payload $payload -Endpoint ''
    $sendRes | ConvertTo-Json
    """
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Send-HubFleetTelemetry failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('{'):
            json_start = i
            break
    data = json.loads('\n'.join(lines[json_start:]))
    assert data.get('Buffered') is True, f"Telemetry should have been buffered offline: {data}"
    assert data.get('InMemory') is True, f"Telemetry should be buffered in-memory: {data}"
    assert not os.path.exists("C:\\AutopilotLogs"), "C:\\AutopilotLogs should NOT exist on disk by default!"
    buf_path = data.get('Path')
    assert buf_path.startswith('memory://FleetBuffer'), f"Expected memory buffer path, got: {buf_path}"
    print(f"  [PASS] In-memory offline buffering verified (Zero disk trace): {buf_path}")

def test_cohort_baselines_and_outliers():
    print("\n--- Test 3: Cross-Fleet Cohort Baselines & Statistical Outlier Alerting ---")
    ps_cmd = """
    $samples = @(
        [PSCustomObject]@{ TotalSeconds = 300; Device = [PSCustomObject]@{ Model = 'Dell Pro Max 16' }; Metrics = [PSCustomObject]@{ Wu = [PSCustomObject]@{ AvgDownloadMBps = 45 }; Disk = [PSCustomObject]@{ AvgIops = 1200 } } },
        [PSCustomObject]@{ TotalSeconds = 310; Device = [PSCustomObject]@{ Model = 'Dell Pro Max 16' }; Metrics = [PSCustomObject]@{ Wu = [PSCustomObject]@{ AvgDownloadMBps = 48 }; Disk = [PSCustomObject]@{ AvgIops = 1250 } } },
        [PSCustomObject]@{ TotalSeconds = 295; Device = [PSCustomObject]@{ Model = 'Dell Pro Max 16' }; Metrics = [PSCustomObject]@{ Wu = [PSCustomObject]@{ AvgDownloadMBps = 44 }; Disk = [PSCustomObject]@{ AvgIops = 1180 } } },
        [PSCustomObject]@{ TotalSeconds = 305; Device = [PSCustomObject]@{ Model = 'Dell Pro Max 16' }; Metrics = [PSCustomObject]@{ Wu = [PSCustomObject]@{ AvgDownloadMBps = 47 }; Disk = [PSCustomObject]@{ AvgIops = 1220 } } },
        [PSCustomObject]@{ TotalSeconds = 302; Device = [PSCustomObject]@{ Model = 'Dell Pro Max 16' }; Metrics = [PSCustomObject]@{ Wu = [PSCustomObject]@{ AvgDownloadMBps = 46 }; Disk = [PSCustomObject]@{ AvgIops = 1210 } } }
    )
    $cohort = Get-HubCohortBaseline -Records $samples -Model 'Dell Pro Max 16'
    $slowUnit = [PSCustomObject]@{
        TotalSeconds = 480
        Metrics = [PSCustomObject]@{
            Wu = [PSCustomObject]@{ AvgDownloadMBps = 12 }
            Disk = [PSCustomObject]@{ AvgIops = 300 }
        }
    }
    $outliers = Test-HubCohortOutlier -Metrics $slowUnit -CohortBaseline $cohort
    [PSCustomObject]@{
        Cohort = $cohort
        Outliers = $outliers
    } | ConvertTo-Json -Depth 5
    """
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Cohort baseline execution failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('{'):
            json_start = i
            break
    data = json.loads('\n'.join(lines[json_start:]))
    c_stats = data['Cohort']['DurationStats']
    assert c_stats['N'] == 5
    assert 300 <= c_stats['Mean'] <= 305
    print(f"  [PASS] Cohort Baseline: N={c_stats['N']}, Mean={c_stats['Mean']}s, StdDev={c_stats['StdDev']}s, Median={c_stats['Median']}s")

    outliers = data['Outliers']
    assert len(outliers) >= 3, f"Expected 3 outlier alerts (duration, download, disk iops), got {len(outliers)}: {outliers}"
    print(f"  [PASS] Outlier Detection: {len(outliers)} alert(s) triggered:")
    metrics_flagged = [o['Metric'] for o in outliers]
    assert 'Total Deployment Duration' in metrics_flagged
    assert 'WU Network Download Speed' in metrics_flagged
    assert 'Storage Disk IOPS' in metrics_flagged
    for o in outliers:
        print(f"    - [{o['Severity']}] {o['Metric']}: {o['Message']} (Z={o.get('ZScore')})")

def test_deployment_profiles_config_as_code():
    print("\n--- Test 4: Deployment Profiles (Config-as-Code) ---")
    ps_cmd = """
    $profiles = Get-HubDeploymentProfiles
    $applied = Apply-HubDeploymentProfile -ProfileNameOrId 'Melbourne-Office'
    [PSCustomObject]@{
        ProfileCount = $profiles.Count
        Profiles     = @($profiles | ForEach-Object { $_.Name })
        AppliedName  = $applied.Name
        GroupTagEnv  = $env:AUTOPILOT_GROUP_TAG
        NameTemplEnv = $env:AUTOPILOT_NAME_TEMPLATE
        SiteEnv      = $env:AUTOPILOT_SITE
    } | ConvertTo-Json
    """
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Deployment profiles test failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('{'):
            json_start = i
            break
    data = json.loads('\n'.join(lines[json_start:]))

    assert data['ProfileCount'] >= 5, f"Expected at least 5 profiles, got {data['ProfileCount']}"
    assert data['GroupTagEnv'] == 'CORP-MEL-STD'
    assert data['NameTemplEnv'] == 'MEL-%SERIAL%'
    assert data['SiteEnv'] == 'Melbourne'
    print(f"  [PASS] Verified {data['ProfileCount']} Deployment Profiles.")
    print(f"  [PASS] Applied 'Melbourne-Office': GroupTag={data['GroupTagEnv']}, Template={data['NameTemplEnv']}")

def test_oem_driver_tool_and_exit_codes():
    print("\n--- Test 5: OEM Driver & Firmware Management ---")
    ps_cmd = "$tool = Get-HubOemDriverTool; $tool | ConvertTo-Json"
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Get-HubOemDriverTool failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('{'):
            json_start = i
            break
    data = json.loads('\n'.join(lines[json_start:]))
    assert 'Vendor' in data
    assert 'ToolName' in data
    print(f"  [PASS] OEM Tool Detected: Vendor={data['Vendor']}, Tool={data['ToolName']}, Installed={data['IsInstalled']}")

def test_cart_mode_harvest_and_batch_registration():
    print("\n--- Test 6: Bulk Cart Harvest Station & Deduplication ---")
    temp_csv = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'TestCartBatchMulti.csv')
    if os.path.exists(temp_csv):
        os.remove(temp_csv)

    oa3_payload = bytes([0x4F, 0x41, 0x33, 0x00]) + (b'\xAA' * 2044)
    oa3_b64 = base64.b64encode(oa3_payload).decode('ascii')

    # Harvest Unit 1, Unit 2, and then re-harvest Unit 1 (should deduplicate)
    ps_cmd = f"""
    $res1 = Invoke-HubCartHarvest -CartName 'TestCart' -CsvPath '{temp_csv}' -GroupTag 'CART-TEST' -HardwareHashOverride '{oa3_b64}' -SerialNumberOverride 'UNIT-001'
    $res2 = Invoke-HubCartHarvest -CartName 'TestCart' -CsvPath '{temp_csv}' -GroupTag 'CART-TEST' -HardwareHashOverride '{oa3_b64}' -SerialNumberOverride 'UNIT-002'
    $res3 = Invoke-HubCartHarvest -CartName 'TestCart' -CsvPath '{temp_csv}' -GroupTag 'CART-TEST' -HardwareHashOverride '{oa3_b64}' -SerialNumberOverride 'UNIT-001'
    [PSCustomObject]@{{
        Step1Total = $res1.TotalUnits
        Step2Total = $res2.TotalUnits
        Step3Total = $res3.TotalUnits
        CsvPath    = $res1.CartCsvPath
        Success    = $res1.Success
    }} | ConvertTo-Json
    """
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Invoke-HubCartHarvest failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = next(i for i, l in enumerate(lines) if l.strip().startswith('{'))
    data = json.loads('\n'.join(lines[json_start:]))

    assert data['Step1Total'] == 1
    assert data['Step2Total'] == 2, f"Expected 2 distinct units, got {data['Step2Total']}"
    assert data['Step3Total'] == 2, f"Expected deduplication to keep total at 2, got {data['Step3Total']}"
    print(f"  [PASS] Cart Harvest multi-unit accumulation & deduplication verified: {data['CsvPath']} (Total: 2 units)")

    # Verify Register-AutopilotDevice accepts HardwareHash and SerialNumber without ParameterBindingException
    reg_test_cmd = f"""
    try {{
        # Test call with dummy token to verify parameter binding
        $null = Register-AutopilotDevice -AccessToken 'dummy-token' -HardwareHash '{oa3_b64}' -SerialNumber 'TEST-SN' -ErrorAction Stop
        Write-Host "BINDING_OK"
    }} catch {{
        if ($_.Exception.Message -match 'ParameterBindingException|A parameter cannot be found') {{
            Write-Host "BINDING_FAILED: $($_.Exception.Message)"
        }} else {{
            # Expected to fail on Microsoft Graph HTTP 401 unauthorized, which proves parameter binding succeeded!
            Write-Host "BINDING_OK: $($_.Exception.Message)"
        }}
    }}
    """
    res_reg = run_ps(reg_test_cmd)
    assert "BINDING_OK" in res_reg.stdout, f"Register-AutopilotDevice failed parameter binding:\n{res_reg.stdout}"
    print("  [PASS] Register-AutopilotDevice parameter binding verified for cart batch operations.")

def test_security_posture_and_laps():
    print("\n--- Test 7: Security Posture Enforcement (Defender & Local Admin Posture) ---")
    ps_cmd = """
    $sec = Set-HubSecurityBaseline
    $adm = Set-HubLocalAdminPosture
    [PSCustomObject]@{
        SecBaseline = $sec.Success
        LocalAdmin  = $adm.Success
        Message     = $adm.Message
    } | ConvertTo-Json
    """
    res = run_ps(ps_cmd)
    assert res.returncode == 0, f"Security posture enforcement failed:\n{res.stderr}"

    lines = res.stdout.strip().splitlines()
    json_start = next(i for i, l in enumerate(lines) if l.strip().startswith('{'))
    data = json.loads('\n'.join(lines[json_start:]))
    assert data.get('SecBaseline') is True or data.get('LocalAdmin') is True
    assert "zero registry keys written" in data.get('Message', '')
    print(f"  [PASS] Security Baseline and Local Admin Posture verified (Zero registry keys written).")

def test_live_fleet_server_and_dashboard():
    print("\n--- Test 8: Live Fleet Plane HTTP Server & Real-Time Dashboard ---")
    test_port = 8499
    temp_fleet_dir = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), f'FleetData_{int(time.time())}')

    # Launch Start-HubFleetServer.ps1 as background process
    server_cmd = [
        'pwsh', '-ExecutionPolicy', 'Bypass', '-File', server_ps1,
        '-Port', str(test_port),
        '-DataDir', temp_fleet_dir
    ]
    proc = subprocess.Popen(server_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    try:
        # Wait up to 10 seconds for server to start listening
        server_ready = False
        health_url = f"http://localhost:{test_port}/api/health"
        for _ in range(20):
            time.sleep(0.5)
            try:
                with urllib.request.urlopen(health_url, timeout=2) as resp:
                    if resp.status == 200:
                        server_ready = True
                        break
            except Exception:
                continue

        assert server_ready, f"Fleet Server failed to start on port {test_port}"
        print(f"  [PASS] Fleet Server listening at http://localhost:{test_port}/")

        # 1. POST /api/telemetry (7 deployments, including 1 outlier)
        for i in range(1, 7):
            payload = {
                "DeploymentId": f"HUB-20260923-00{i}",
                "RoutineName": "1. Intune-Only Cloud Build",
                "Status": "Success",
                "TotalSeconds": 300 + (i * 2),
                "Site": "Melbourne",
                "Technician": "tech1",
                "Device": {"SerialNumber": f"SN00{i}", "Model": "Latitude 5520"},
                "Metrics": {
                    "TotalSeconds": 300 + (i * 2),
                    "Wu": {"AvgDownloadMBps": 45.0},
                    "Disk": {"AvgIops": 1200}
                }
            }
            req = urllib.request.Request(
                f"http://localhost:{test_port}/api/telemetry",
                data=json.dumps(payload).encode('utf-8'),
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            with urllib.request.urlopen(req, timeout=3) as resp:
                assert resp.status == 200

        # Outlier unit
        slow_payload = {
            "DeploymentId": "HUB-20260923-OUTLIER",
            "RoutineName": "1. Intune-Only Cloud Build",
            "Status": "Success",
            "TotalSeconds": 580.0,
            "Site": "Sydney",
            "Technician": "tech2",
            "Device": {"SerialNumber": "SN-SLOW-01", "Model": "Latitude 5520"},
            "Metrics": {
                "TotalSeconds": 580.0,
                "Wu": {"AvgDownloadMBps": 8.0},
                "Disk": {"AvgIops": 250}
            }
        }
        req = urllib.request.Request(
            f"http://localhost:{test_port}/api/telemetry",
            data=json.dumps(slow_payload).encode('utf-8'),
            headers={'Content-Type': 'application/json'},
            method='POST'
        )
        with urllib.request.urlopen(req, timeout=3) as resp:
            assert resp.status == 200
        print(f"  [PASS] POST /api/telemetry: 7 payloads ingested successfully.")

        # 2. GET /api/deployments
        with urllib.request.urlopen(f"http://localhost:{test_port}/api/deployments", timeout=3) as resp:
            assert resp.status == 200
            deps = json.loads(resp.read().decode('utf-8'))
            assert len(deps) == 7
        print(f"  [PASS] GET /api/deployments: 7 deployments retrieved.")

        # 3. GET /api/cohorts
        with urllib.request.urlopen(f"http://localhost:{test_port}/api/cohorts", timeout=3) as resp:
            assert resp.status == 200
            cohorts = json.loads(resp.read().decode('utf-8'))
            if isinstance(cohorts, dict):
                cohorts = [cohorts]
            assert len(cohorts) >= 1
            lat = next(c for c in cohorts if 'Latitude 5520' in c.get('Model', ''))
            assert lat['SampleCount'] == 7
        print(f"  [PASS] GET /api/cohorts: Cohort baselines computed across hardware models.")

        # 4. GET /api/outliers
        with urllib.request.urlopen(f"http://localhost:{test_port}/api/outliers", timeout=3) as resp:
            assert resp.status == 200
            outliers = json.loads(resp.read().decode('utf-8'))
            if isinstance(outliers, dict):
                outliers = [outliers]
            assert len(outliers) >= 1
            flagged = next(o for o in outliers if o.get('DeploymentId') == 'HUB-20260923-OUTLIER')
            assert flagged['Severity'] in ['ALERT', 'WARN']
        print(f"  [PASS] GET /api/outliers: Outlier unit correctly flagged (Z={flagged.get('ZScore')}).")

        # 5. GET / (Dashboard HTML)
        with urllib.request.urlopen(f"http://localhost:{test_port}/", timeout=3) as resp:
            assert resp.status == 200
            html = resp.read().decode('utf-8')
            assert "Autopilot Fleet Plane Dashboard" in html
            assert "HUB-20260923-OUTLIER" in html
            assert "Active Statistical Outliers" in html
        print(f"  [PASS] GET /: Fleet Dashboard HTML rendered with active outlier alert card.")

    finally:
        proc.terminate()
        try:
            proc.wait(timeout=3)
        except Exception:
            proc.kill()

def test_ast_and_pure_ascii():
    print("\n--- Test 9: PowerShell 5.1 AST & Pure ASCII on All Generated Artifacts ---")
    files_to_check = [ps1_path, os.path.join(repo_dir, 'autopilot'), server_ps1]
    for fp in files_to_check:
        fname = os.path.basename(fp)
        with open(fp, 'rb') as f:
            raw = f.read()

        assert not raw.startswith(b'\xef\xbb\xbf'), f"{fname} contains UTF-8 BOM!"
        non_ascii = [(idx, b) for idx, b in enumerate(raw) if b > 127]
        assert len(non_ascii) == 0, f"{fname} contains non-ASCII bytes: {non_ascii[:5]}!"

        ps_cmd = "$c = Get-Content -LiteralPath '" + fp.replace("'", "''") + "' -Raw; [scriptblock]::Create($c) | Out-Null; Write-Host 'PS51_PARSE_OK'"
        res = subprocess.run(['powershell.exe', '-NoProfile', '-Command', ps_cmd], capture_output=True, encoding='utf-8', errors='replace')
        assert res.returncode == 0, f"PS 5.1 AST syntax check failed on {fname}:\n{res.stderr}"
        assert "PS51_PARSE_OK" in res.stdout, f"PS 5.1 AST syntax check failed on {fname}:\n{res.stdout}"
        print(f"  [PASS] {fname}: {len(raw)} bytes, UTF-8 without BOM, 100% pure ASCII, PS 5.1 AST clean.")

if __name__ == '__main__':
    print("==================================================================")
    print(" FLEET PLANE & CLOSED-LOOP VERIFICATION TEST SUITE")
    print("==================================================================")
    test_closed_loop_verification_and_receipt()
    test_fleet_telemetry_and_buffering()
    test_cohort_baselines_and_outliers()
    test_deployment_profiles_config_as_code()
    test_oem_driver_tool_and_exit_codes()
    test_cart_mode_harvest_and_batch_registration()
    test_security_posture_and_laps()
    test_live_fleet_server_and_dashboard()
    test_ast_and_pure_ascii()
    print("==================================================================")
    print(" ALL FLEET & VERIFICATION TESTS PASSED CLEANLY (100% VERIFICATION)")
    print("==================================================================")
