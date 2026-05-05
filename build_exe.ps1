$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = "python"
$entry = "ticket_printer.py"
$appName = "财务票据打印_V1.0"
$templateName = "template_workbook.xls"
$dataFileName = "ticket_printer_data.json"
$sourceTemplate = Join-Path $root $templateName
$sourceDataFile = Join-Path $root $dataFileName

Write-Host "Cleaning old build output..."
Remove-Item -LiteralPath ".\build" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\dist" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\$appName.spec" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath ".\ticket_printer.spec" -Force -ErrorAction SilentlyContinue

Write-Host "Building exe with PyInstaller..."
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name $appName `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32timezone `
    $entry

Write-Host ""
Write-Host "Build finished:"
Write-Host "  .\dist\$appName.exe"
if (-not (Test-Path $sourceTemplate)) {
    throw "Template file not found: $templateName"
}
Copy-Item -LiteralPath $sourceTemplate -Destination (Join-Path $root "dist\$templateName") -Force
Write-Host "  .\dist\$templateName"
if (-not (Test-Path $sourceDataFile)) {
    throw "Data file not found: $dataFileName"
}
Copy-Item -LiteralPath $sourceDataFile -Destination (Join-Path $root "dist\$dataFileName") -Force
Write-Host "  .\dist\$dataFileName"
