param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    Remove-Item -LiteralPath "build", "dist" -Recurse -Force -ErrorAction SilentlyContinue
}

$tkinterDndPath = python -c "import tkinterdnd2, pathlib; print(pathlib.Path(tkinterdnd2.__file__).resolve().parent)"
$tkdndWinX64 = Join-Path $tkinterDndPath "tkdnd\win-x64"

# Use onedir instead of onefile: onefile EXEs must unpack a large temporary
# runtime on every launch, which is especially slow with Windows Defender.
python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name "I-Need-Rest" `
    --distpath "dist" `
    --workpath "build" `
    --specpath "build\spec" `
    --add-data "$tkdndWinX64;tkinterdnd2/tkdnd/win-x64" `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    --hidden-import win32timezone `
    --exclude-module numpy `
    --exclude-module matplotlib `
    --exclude-module psutil `
    bsod_shortcut.py

$portableZip = "dist\I-Need-Rest-portable.zip"
Remove-Item -LiteralPath $portableZip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path "dist\I-Need-Rest" -DestinationPath $portableZip -CompressionLevel Fastest

Write-Host "Build complete: $((Resolve-Path 'dist\I-Need-Rest\I-Need-Rest.exe').Path)"
Write-Host "Portable package: $((Resolve-Path $portableZip).Path)"
