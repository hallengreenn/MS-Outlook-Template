# Install Office Development Tools - Run as Administrator
# Dette script åbner Visual Studio Installer og viser dig præcis hvad du skal klikke

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Office Development Tools Installation" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Dette script åbner Visual Studio Installer." -ForegroundColor Yellow
Write-Host "Følg derefter disse simple steps:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Find 'Visual Studio Community 2025'" -ForegroundColor White
Write-Host "2. Klik på 'Modify' knappen" -ForegroundColor White
Write-Host "3. Vælg 'Workloads' tab" -ForegroundColor White
Write-Host "4. Sæt hak ved 'Office/SharePoint development'" -ForegroundColor White
Write-Host "5. Klik 'Modify' nederst til højre" -ForegroundColor White
Write-Host "6. Vent 10-15 minutter" -ForegroundColor White
Write-Host ""

$response = Read-Host "Tryk Enter for at åbne Visual Studio Installer"

# Åbn Visual Studio Installer
Start-Process "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe"

Write-Host ""
Write-Host "Visual Studio Installer åbnet!" -ForegroundColor Green
Write-Host "Følg steps ovenfor..." -ForegroundColor Green
Write-Host ""
Write-Host "Når installation er færdig, fortæl Claude: 'installation færdig'" -ForegroundColor Cyan
