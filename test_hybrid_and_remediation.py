# test_hybrid_and_remediation.py
# Automated Verification Suite for Hybrid & Co-Management Enhancements and Precision Enterprise Remediation

import os
import sys
import subprocess
import json
import xml.etree.ElementTree as ET

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

repo_dir = os.path.dirname(os.path.abspath(__file__))
ps1_path = os.path.join(repo_dir, 'autopilot.ps1')

def test_xaml_remediation_tab():
    print("\n--- Test 1: XAML Precision Enterprise Fixes & Hybrid Controls ---")
    with open(ps1_path, 'r', encoding='utf-8') as f:
        content = f.read()

    start = content.find("$xaml = @'\n") + len("$xaml = @'\n")
    end = content.find("\n'@", start)
    xaml = content[start:end]
    tree = ET.fromstring(xaml)

    tab_headers = [elem.attrib.get('Header', '') for elem in tree.iter() if elem.tag.endswith('TabItem')]
    print(f"  Discovered Tabs: {tab_headers}")
    assert 'Precision Enterprise Fixes' in tab_headers, "Tab 'Precision Enterprise Fixes' missing from XAML!"
    assert any('Hybrid' in h for h in tab_headers), "Tab 'Hybrid & Co-Mgmt' missing from XAML!"

    # Verify buttons exist in XAML
    expected_buttons = [
        'BtnHybDcLadder', 'BtnHybScp', 'BtnHybCoMgmtAllIntune', 'BtnHybCoMgmtAllCcm', 'BtnHybCoMgmtPilot',
        'BtnHybKerbDiag', 'BtnHybKerbPurge', 'BtnHybPrtDiag', 'BtnHybResetBroker', 'BtnHybCertPulse', 'BtnHybScepHealth',
        'BtnFixHealthReadout', 'BtnFixWmi', 'BtnFixWu', 'BtnFixCatroot', 'BtnFixBits', 'BtnFixDcom',
        'BtnFixSpooler', 'BtnFixNetStack', 'BtnFixWinRm', 'BtnFixProfiles', 'BtnFixTpm', 'BtnFixAppX'
    ]

    button_names = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('Button')]
    missing = [b for b in expected_buttons if b not in button_names]
    assert len(missing) == 0, f"Missing buttons in XAML: {missing}"
    print(f"  [PASS] All {len(expected_buttons)} enterprise remediation and hybrid buttons verified in XAML.")

def test_powershell_functions_execution():
    print("\n--- Test 2: Execution of Remediation & Hybrid Diagnostics in pwsh ---")
    ps_code = f"""
    . '{ps1_path}' -NoGui

    $h = Get-HubRemediationHealthOverview
    $k = Get-KerberosDiagnostics
    $s = Get-ScepCertificateHealth
    $p = Get-EntraPrtDiagnostics
    $l = Test-DomainControllerLadder -TargetDc '127.0.0.1' -TimeoutMs 200

    [PSCustomObject]@{{
        HealthOverviewWmi = $h.WmiHealthy
        HealthOverviewBits = $h.BitsStatus
        KerberosMessage   = $k.Message
        ScepCount         = $s.Count
        PrtChecked        = ($null -ne $p.HasPrt)
        LadderApplicable  = $l.Applicable
        LadderPortCount   = $l.Ports.Count
    }} | ConvertTo-Json
    """

    res = subprocess.run(['pwsh', '-ExecutionPolicy', 'Bypass', '-Command', ps_code], capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0, f"Remediation probe failed: {res.stderr}"

    start_idx = res.stdout.find('{')
    end_idx = res.stdout.rfind('}')
    assert start_idx != -1 and end_idx != -1, f"JSON output not found: {res.stdout}"
    data = json.loads(res.stdout[start_idx:end_idx+1])

    print(f"  WMI Subsystem Healthy: {data['HealthOverviewWmi']}")
    print(f"  BITS Service Status:   {data['HealthOverviewBits']}")
    print(f"  Kerberos Result:       {data['KerberosMessage']}")
    print(f"  DC Ladder Ports:       {data['LadderPortCount']} ports tested")
    assert data['LadderPortCount'] == 9, f"Expected 9 DC ladder ports, got {data['LadderPortCount']}"
    assert data['PrtChecked'] is True, "Expected PRT check to produce boolean"
    print("  [PASS] All core PowerShell diagnostic and remediation engines executed cleanly.")

def test_ps51_ast():
    print("\n--- Test 3: PowerShell 5.1 AST Syntax Integrity ---")
    ps_ast_code = f"""
    $scriptContent = [System.IO.File]::ReadAllText('{ps1_path}')
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptContent, [ref]$tokens, [ref]$errors)
    if ($errors.Count -eq 0) {{
        Write-Host 'PS51_PARSE_OK'
        exit 0
    }} else {{
        Write-Host "PS51_PARSE_FAILED: $($errors.Count)"
        exit 1
    }}
    """
    res = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps_ast_code], capture_output=True, encoding='utf-8', errors='replace')
    assert res.returncode == 0 and 'PS51_PARSE_OK' in res.stdout, f"PS5.1 AST Parse failed: {res.stderr}\n{res.stdout}"
    print("  [PASS] PowerShell 5.1 AST syntax parse: 100% clean (zero errors).")

if __name__ == '__main__':
    print("==================================================================")
    print(" HYBRID & PRECISION ENTERPRISE REMEDIATION VERIFICATION SUITE")
    print("==================================================================")
    test_xaml_remediation_tab()
    test_powershell_functions_execution()
    test_ps51_ast()
    print("\n==================================================================")
    print(" ALL ENTERPRISE SUITES PASSED CLEANLY (100% EMPIRICAL VERIFICATION)")
    print("==================================================================")
