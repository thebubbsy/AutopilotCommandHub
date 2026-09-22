# test_action_routines.py
# Automated Verification Suite for Action Routines & Deployment Playbook Engine

import os
import sys
import subprocess
import xml.etree.ElementTree as ET

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

repo_dir = os.path.dirname(os.path.abspath(__file__))
ps1_path = os.path.join(repo_dir, 'autopilot.ps1')

def test_xaml_playbook_hud():
    print("\n--- Test 1: XAML Deployment Playbook HUD Controls & Tabs ---")
    with open(ps1_path, 'r', encoding='utf-8') as f:
        content = f.read()

    start = content.find("$xaml = @'\n") + len("$xaml = @'\n")
    end = content.find("\n'@", start)
    xaml = content[start:end]
    tree = ET.fromstring(xaml)

    # Check Playbook HUD border
    borders = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('Border')]
    assert 'PlaybookHud' in borders, "Border 'PlaybookHud' missing from XAML!"

    # Check ComboBox
    combos = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('ComboBox')]
    assert 'CboPlaybookRoutine' in combos, "ComboBox 'CboPlaybookRoutine' missing from XAML!"

    # Check ComboBox Items
    combo_elem = [elem for elem in tree.iter() if elem.attrib.get('Name') == 'CboPlaybookRoutine'][0]
    combo_items = [elem.attrib.get('Content', '') for elem in combo_elem.iter() if elem.tag.endswith('ComboBoxItem')]
    print(f"  Playbook Routines in ComboBox: {len(combo_items)}")
    for ci in combo_items:
        print(f"    - {ci}")
    assert len(combo_items) == 5, f"Expected 5 playbook routines, got {len(combo_items)}"

    # Check buttons
    buttons = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('Button')]
    expected_playbook_buttons = ['BtnPlaybookRun', 'BtnPlaybookPause', 'BtnPlaybookStop']
    for b in expected_playbook_buttons:
        assert b in buttons, f"Playbook button {b} missing from XAML!"

    # Check step indicator textblock and progress bar
    tb_names = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('TextBlock')]
    assert 'TxtPlaybookStep' in tb_names, "TextBlock 'TxtPlaybookStep' missing from XAML!"

    pb_names = [elem.attrib.get('Name', '') for elem in tree.iter() if elem.tag.endswith('ProgressBar')]
    assert 'PlaybookProgressBar' in pb_names, "ProgressBar 'PlaybookProgressBar' missing from XAML!"

    # Check tabs
    tab_headers = [elem.attrib.get('Header', '') for elem in tree.iter() if elem.tag.endswith('TabItem') and 'Header' in elem.attrib and 'Flow' not in elem.attrib['Header'] and 'Secret' not in elem.attrib['Header']]
    print(f"  Discovered Tabs ({len(tab_headers)}): {tab_headers}")
    assert len(tab_headers) == 9, f"Expected 9 tabs, got {len(tab_headers)}"

    print("  [PASS] Playbook HUD, all 3 controls, 5 routines, and 9 tabs verified in XAML.")

def test_headless_playbook_execution():
    print("\n--- Test 2: Headless Playbook Execution (All 5 Action Routines) ---")
    routines = [
        "1. Intune-Only Cloud Build",
        "2. Hybrid AD Join & Co-Management Build",
        "3. Local User & Offline Bypass Build",
        "4. Deep System Remediation & Health Sweep",
        "5. Hardware Health & Asset Intake Audit"
    ]

    for routine in routines:
        print(f"  Testing routine: '{routine}'...")
        ps_code = f"& '{ps1_path}' -NoGui -Playbook '{routine}'"
        res = subprocess.run(['pwsh', '-ExecutionPolicy', 'Bypass', '-Command', ps_code], capture_output=True, encoding='utf-8', errors='replace')
        assert res.returncode == 0, f"Routine '{routine}' failed with exit code {res.returncode}:\n{res.stderr}"
        assert "completed successfully" in res.stdout, f"Routine '{routine}' did not report success:\n{res.stdout}"
        print(f"    [PASS] '{routine}' completed cleanly.")

def test_ascii_and_bom():
    print("\n--- Test 3: UTF-8 Without BOM & Pure ASCII Verification ---")
    for fname in ['autopilot.ps1', 'autopilot']:
        fpath = os.path.join(repo_dir, fname)
        with open(fpath, 'rb') as f:
            raw = f.read()

        assert not raw.startswith(b'\xef\xbb\xbf'), f"{fname} contains UTF-8 BOM!"
        non_ascii = [(idx, b) for idx, b in enumerate(raw) if b > 127]
        assert len(non_ascii) == 0, f"{fname} contains non-ASCII bytes!"
        print(f"  [PASS] {fname}: {len(raw)} bytes, UTF-8 without BOM, 100% pure ASCII.")

if __name__ == '__main__':
    print("==================================================================")
    print(" ACTION ROUTINES & PLAYBOOK VERIFICATION SUITE")
    print("==================================================================")
    test_xaml_playbook_hud()
    test_headless_playbook_execution()
    test_ascii_and_bom()
    print("==================================================================")
    print(" ALL ACTION ROUTINE TESTS PASSED CLEANLY (100% EMPIRICAL VERIFICATION)")
    print("==================================================================")
