param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    Remove-Item -LiteralPath "build", "dist" -Recurse -Force -ErrorAction SilentlyContinue
}

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "I-Need-Rest" `
    --collect-all tkinterdnd2 `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    --hidden-import win32timezone `
    bsod_shortcut.py

Write-Host "Build complete: $((Resolve-Path 'dist\I-Need-Rest.exe').Path)"
