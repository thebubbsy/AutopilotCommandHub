# render_ui.py
import re
import subprocess
import os

with open('autopilot.ps1', 'r', encoding='utf-8') as f:
    content = f.read()

start = content.find("$xaml = @'\n") + len("$xaml = @'\n")
end = content.find("\n'@", start)
xaml = content[start:end]

with open('C:\\temp\\extracted.xaml', 'w', encoding='utf-8') as f:
    f.write(xaml)

ps_script = r'''
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$reader = [System.Xml.XmlReader]::Create("C:\temp\extracted.xaml")
$w = [System.Windows.Markup.XamlReader]::Load($reader)
$w.Show()
$w.UpdateLayout()

$rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
    [int]$w.ActualWidth, [int]$w.ActualHeight, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32
)
$rtb.Render($w)
$enc = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
$enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
$fs = [System.IO.FileStream]::new("C:\temp\rendered_window.png", [System.IO.FileMode]::Create)
$enc.Save($fs)
$fs.Close()
$w.Close()
Write-Host "RENDER_SUCCESS"
'''

with open('C:\\temp\\render.ps1', 'w', encoding='utf-8') as f:
    f.write(ps_script)

res = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', 'C:\\temp\\render.ps1'], capture_output=True, text=True)
print("STDOUT:", res.stdout)
print("STDERR:", res.stderr)
