# PowerShell script til at afinstallere Outlook VSTO Add-in

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Outlook Meeting Template Add-in Uninstaller" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check om Outlook kører
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "⚠️  ADVARSEL: Outlook kører!" -ForegroundColor Yellow
    Write-Host "Luk venligst Outlook før afinstallation fortsætter." -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "Vil du lukke Outlook nu? (J/N)"
    if ($response -eq "J" -or $response -eq "j") {
        Stop-Process -Name "OUTLOOK" -Force
        Write-Host "✓ Outlook lukket" -ForegroundColor Green
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Afinstallation afbrudt." -ForegroundColor Red
        exit
    }
}

# Fjern registry nøgle
$registryPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"

try {
    if (Test-Path $registryPath) {
        Remove-Item -Path $registryPath -Recurse -Force
        Write-Host "✓ Registry nøgle fjernet" -ForegroundColor Green
    } else {
        Write-Host "ℹ️  Add-in var ikke installeret" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Afinstallation fuldført!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

} catch {
    Write-Host "❌ Fejl ved afinstallation: $_" -ForegroundColor Red
    exit
}
