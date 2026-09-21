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
        'BtnHybKerbDiag', 'BtnHybKerbPurge', 'BtnHybSpn', 'BtnHybPrtDiag', 'BtnHybAadToken', 'BtnHybResetBroker',
        'BtnHybCertPulse', 'BtnHybScepHealth',
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
    $spn = Test-ComputerSpnRegistration
    $tok = Test-EntraTokenAcquisition
    $l = Test-DomainControllerLadder -TargetDc '127.0.0.1' -TimeoutMs 200

    [PSCustomObject]@{{
        HealthOverviewWmi = $h.WmiHealthy
        HealthOverviewBits = $h.BitsStatus
        KerberosMessage   = $k.Message
        ScepCount         = $s.Count
        PrtChecked        = ($null -ne $p.HasPrt)
        SpnChecked        = ($null -ne $spn.Status)
        TokenAcqChecked   = ($null -ne $tok.Verdict)
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
    print(f"  SPN Status Checked:    {data['SpnChecked']}")
    print(f"  Token Acq Checked:     {data['TokenAcqChecked']}")
    print(f"  DC Ladder Ports:       {data['LadderPortCount']} ports tested")
    assert data['LadderPortCount'] == 9, f"Expected 9 DC ladder ports, got {data['LadderPortCount']}"
    assert data['PrtChecked'] is True, "Expected PRT check to produce boolean"
    assert data['SpnChecked'] is True, "Expected SPN check to produce status"
    assert data['TokenAcqChecked'] is True, "Expected token acquisition to produce verdict"
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

def test_encoding_and_ascii():
    print("\n--- Test 4: UTF-8 Without BOM & Pure ASCII Verification ---")
    for fname in ['autopilot.ps1', 'autopilot']:
        fpath = os.path.join(repo_dir, fname)
        with open(fpath, 'rb') as f:
            raw = f.read()

        # Check BOM: UTF-8 BOM is 0xEF, 0xBB, 0xBF. First 4 bytes must be '<#\n.' -> [60, 35, 10, 46]
        prefix = list(raw[:4])
        assert prefix == [60, 35, 10, 46], f"{fname} has invalid magic bytes: {prefix}, expected [60, 35, 10, 46] (no BOM)"

        # Check ASCII
        non_ascii = [(idx, b) for idx, b in enumerate(raw) if b > 127]
        assert len(non_ascii) == 0, f"{fname} contains {len(non_ascii)} non-ASCII bytes! First at {non_ascii[:5]}"
        print(f"  [PASS] {fname}: {len(raw)} bytes, UTF-8 without BOM, 100% pure ASCII.")

if __name__ == '__main__':
    print("==================================================================")
    print(" HYBRID & PRECISION ENTERPRISE REMEDIATION VERIFICATION SUITE")
    print("==================================================================")
    test_xaml_remediation_tab()
    test_powershell_functions_execution()
    test_ps51_ast()
    test_encoding_and_ascii()
    print("\n==================================================================")
    print(" ALL ENTERPRISE SUITES PASSED CLEANLY (100% EMPIRICAL VERIFICATION)")
    print("==================================================================")
