# Enable Meeting Template Add-in i Outlook
# Kør som den bruger der skal bruge add-in'et (IKKE som admin)

Write-Host "===================================" -ForegroundColor Cyan
Write-Host "Enable Meeting Template Add-in" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

# Add-in details
$addinId = "d6ff5248-95b4-4400-8f5d-4ce60c4ed0be"
$manifestUrl = "https://icy-hill-0716f1a03.1.azurestaticapps.net/manifest.xml"
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$addinId"

# Check om Outlook kører
$outlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "ADVARSEL: Outlook kører!" -ForegroundColor Yellow
    Write-Host "Luk Outlook før du fortsætter." -ForegroundColor Yellow
    Write-Host ""
    $continue = Read-Host "Tryk Enter når Outlook er lukket (eller Ctrl+C for at afbryde)"
}

# Opret registry keys
Write-Host "Opretter registry keys..." -ForegroundColor Green

try {
    # Opret path hvis den ikke eksisterer
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Sæt add-in properties
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $regPath -Name "Manifest" -Value $manifestUrl -Type String
    Set-ItemProperty -Path $regPath -Name "Description" -Value "Meeting Template Add-in" -Type String
    Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Meeting Template" -Type String

    Write-Host ""
    Write-Host "SUCCESS! Add-in er konfigureret." -ForegroundColor Green
    Write-Host ""
    Write-Host "Registry keys oprettet:" -ForegroundColor Cyan
    Write-Host "  Path: $regPath" -ForegroundColor Gray
    Write-Host "  LoadBehavior: 3 (Always enabled)" -ForegroundColor Gray
    Write-Host "  Manifest: $manifestUrl" -ForegroundColor Gray
    Write-Host ""

    # Verificer
    Write-Host "Verificerer..." -ForegroundColor Cyan
    $verification = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    if ($verification.LoadBehavior -eq 3) {
        Write-Host "✓ Verified: LoadBehavior = 3" -ForegroundColor Green
    }
    if ($verification.Manifest -eq $manifestUrl) {
        Write-Host "✓ Verified: Manifest URL korrekt" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "NÆSTE SKRIDT:" -ForegroundColor Yellow
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "1. Åbn Outlook" -ForegroundColor White
    Write-Host "2. Gå til Calendar" -ForegroundColor White
    Write-Host "3. Opret New Meeting" -ForegroundColor White
    Write-Host "4. 'Meeting Template' knap skulle være synlig i ribbon!" -ForegroundColor White
    Write-Host ""
    Write-Host "Hvis knappen IKKE vises:" -ForegroundColor Yellow
    Write-Host "  - Check File → Options → Trust Center → Trust Center Settings → Add-ins" -ForegroundColor Gray
    Write-Host "  - Sørg for 'Require Application Extensions...' er UNCHECKED" -ForegroundColor Gray
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "ERROR: Kunne ikke oprette registry keys" -ForegroundColor Red
    Write-Host "Fejl: $_" -ForegroundColor Red
    Write-Host ""
}

Write-Host "Script færdigt!" -ForegroundColor Cyan
Write-Host ""
