# test_dell_warranty.py
# Automated Verification Test Suite for Dell Asset Warranty & Hardware Refresh Lifecycle Engine

import os
import sys
import subprocess
import json
import xml.etree.ElementTree as ET

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def test_xaml_validity():
    print("\n--- Test 1: XAML Syntax & XML Tree Integrity ---")
    ps1_path = os.path.join(os.path.dirname(__file__), 'autopilot.ps1')
    assert os.path.exists(ps1_path), "autopilot.ps1 missing"

    with open(ps1_path, 'r', encoding='utf-8') as f:
        content = f.read()

    start = content.find('$xaml = @\'\n') + len('$xaml = @\'\n')
    end = content.find('\n\'@', start)
    assert start > len('$xaml = @\'\n') and end > start, "Could not extract XAML block"

    xaml = content[start:end]
    tree = ET.fromstring(xaml)
    assert 'Window' in tree.tag, f"Root element expected Window, got {tree.tag}"
    elements = list(tree.iter())
    print(f"  [PASS] XAML XML successfully parsed: {len(elements)} nodes.")

    # Verify Tab 5 exists in XAML
    tab_headers = [elem.attrib.get('Header', '') for elem in tree.iter() if elem.tag.endswith('TabItem')]
    assert any('Dell' in h for h in tab_headers), f"Dell tab not found in TabHeaders: {tab_headers}"
    print(f"  [PASS] Dell Warranty Tab found in XAML: {tab_headers}")

def test_standalone_get_dell_warranty():
    print("\n--- Test 2: Get-DellWarranty.ps1 Invocation (BIOS Service Tag) ---")
    script_path = os.path.join(os.path.dirname(__file__), 'Get-DellWarranty.ps1')
    assert os.path.exists(script_path), "Get-DellWarranty.ps1 missing"

    cmd = [
        'pwsh', '-ExecutionPolicy', 'Bypass', '-Command',
        f'& "{script_path}" -ServiceTag "6BYQJW2" -OutputFormat Json'
    ]
    res = subprocess.run(cmd, capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0, f"Get-DellWarranty.ps1 failed with code {res.returncode}: {res.stderr}"

    data = json.loads(res.stdout.strip())
    assert data['ServiceTag'] == '6BYQJW2', f"Unexpected service tag: {data.get('ServiceTag')}"
    assert 'Dell' in data['SystemModel'], f"Unexpected system model: {data.get('SystemModel')}"
    assert data['WarrantyEndDate'] == '2024-01-14', f"Unexpected warranty end date: {data.get('WarrantyEndDate')}"
    assert data['IsUnderWarranty'] is False, f"Expected out of warranty, got {data.get('IsUnderWarranty')}"
    assert 'REFRESH RECOMMENDED' in data['RefreshVerdict'], f"Unexpected verdict: {data.get('RefreshVerdict')}"
    assert len(data['Entitlements']) >= 3, f"Expected at least 3 entitlements, got {len(data.get('Entitlements', []))}"
    print(f"  [PASS] Get-DellWarranty verified: {data['SystemModel']} | Status: {data['WarrantyStatus']} | Verdict: {data['RefreshVerdict']}")

def test_autopilot_cli_dell_warranty():
    print("\n--- Test 3: autopilot.ps1 -DellWarranty CLI Mode ---")
    script_path = os.path.join(os.path.dirname(__file__), 'autopilot.ps1')
    csv_path = os.path.join(os.environ.get('TEMP', 'C:\\temp'), 'TestDellWarranty.csv')

    if os.path.exists(csv_path):
        os.remove(csv_path)

    cmd = [
        'pwsh', '-ExecutionPolicy', 'Bypass', '-Command',
        f'& "{script_path}" -DellWarranty -DellServiceTag "6BYQJW2" -ExportCsv -CsvPath "{csv_path}"'
    ]
    res = subprocess.run(cmd, capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0, f"autopilot.ps1 failed: {res.stderr}"
    assert "DELL ASSET WARRANTY & REFRESH ASSESSMENT" in res.stdout, "Assessment header not in stdout"
    assert "REFRESH RECOMMENDED" in res.stdout, "Verdict not in stdout"
    assert os.path.exists(csv_path), f"CSV export file not created: {csv_path}"

    with open(csv_path, 'r', encoding='utf-8') as f:
        csv_lines = f.readlines()
    assert len(csv_lines) >= 2, f"CSV had fewer lines than expected: {csv_lines}"
    assert '6BYQJW2' in csv_lines[1], "Service tag not in CSV row"
    print(f"  [PASS] autopilot.ps1 -DellWarranty -ExportCsv verified: {csv_path} ({os.path.getsize(csv_path)} bytes)")

def test_env_and_gateway():
    print("\n--- Test 4: Dynamic Gateway Detection & .env Defaults Engine ---")
    script_path = os.path.join(os.path.dirname(__file__), 'autopilot.ps1')
    
    # Test Test-StagedNetwork Stage 2 dynamic resolution and .env loading
    ps_code = f"""
    . '{script_path}' -NoGui
    $diag = Test-StagedNetwork
    $s2 = $diag.Stages | Where-Object {{ $_.Stage -eq 2 }}
    $encName = Get-ScriptEncodingName
    [PSCustomObject]@{{
        GatewayDiscovered = $s2.Details
        GatewaySuccess    = $s2.Success
        DefaultEncoding   = $encName
        LoadedEnv         = $script:LoadedEnvPath
    }} | ConvertTo-Json
    """
    cmd = ['pwsh', '-ExecutionPolicy', 'Bypass', '-Command', ps_code]
    res = subprocess.run(cmd, capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0, f"Network/Env test failed: {res.stderr}"
    
    data = json.loads(res.stdout.strip())
    print(f"  [PASS] Dynamic Gateway attestation: {data['GatewayDiscovered']}")
    print(f"  [PASS] Native Charset resolution: {data['DefaultEncoding']}")
    assert "Gateway" in data['GatewayDiscovered'], "Gateway stage output not found"

def test_custom_env_options():
    print("\n--- Test 5: Custom .env Defaults (Rename Toggle, Prefix, Native Charset) ---")
    script_path = os.path.join(os.path.dirname(__file__), 'autopilot.ps1')
    temp_env = os.path.join(os.environ.get('TEMP', 'C:\\temp'), 'test_custom.env')
    
    with open(temp_env, 'w', encoding='utf-8') as f:
        f.write("AUTOPILOT_RENAME_ENABLED=true\n")
        f.write("AUTOPILOT_NAME_PREFIX=LT\n")
        f.write("AUTOPILOT_GROUP_TAG=Fleet-Dev\n")
        f.write("AUTOPILOT_CHARSET=native\n")
        f.write("DEFAULT_GATEWAY=192.168.88.1\n")
        
    ps_code = f"""
    . '{script_path}' -NoGui -EnvFile '{temp_env}'
    [PSCustomObject]@{{
        Prefix         = $ComputerNamePrefix
        Template       = $ComputerNameTemplate
        RenameEnabled  = [bool]$RenameComputer
        GroupTag       = $GroupTag
        EncodingName   = (Get-ScriptEncodingName)
    }} | ConvertTo-Json
    """
    cmd = ['pwsh', '-ExecutionPolicy', 'Bypass', '-Command', ps_code]
    res = subprocess.run(cmd, capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0, f"Custom env test failed: {res.stderr}"
    
    start = res.stdout.find('{')
    end = res.stdout.rfind('}')
    assert start != -1 and end != -1, f"JSON block not found in stdout: {res.stdout}"
    data = json.loads(res.stdout[start:end+1])
    assert data['Prefix'] == 'LT', f"Expected prefix 'LT', got {data.get('Prefix')}"
    assert data['Template'] == 'LT-%SERIAL%', f"Expected template 'LT-%SERIAL%', got {data.get('Template')}"
    assert data['RenameEnabled'] is True, f"Expected rename enabled, got {data.get('RenameEnabled')}"
    assert data['GroupTag'] == 'Fleet-Dev', f"Expected GroupTag 'Fleet-Dev', got {data.get('GroupTag')}"
    assert data['EncodingName'] == 'Default', f"Expected native charset 'Default', got {data.get('EncodingName')}"
    print(f"  [PASS] Custom .env loaded: Prefix={data['Prefix']} | Template={data['Template']} | Rename={data['RenameEnabled']} | Charset={data['EncodingName']}")

if __name__ == '__main__':
    print("==================================================================")
    print(" DELL WARRANTY & REFRESH ENGINE AUTOMATED VERIFICATION SUITE")
    print("==================================================================")
    test_xaml_validity()
    test_standalone_get_dell_warranty()
    test_autopilot_cli_dell_warranty()
    test_env_and_gateway()
    test_custom_env_options()
    print("\n==================================================================")
    print(" ALL 5 TEST SUITES PASSED CLEANLY (100% EMPIRICAL VERIFICATION)")
    print("==================================================================")
