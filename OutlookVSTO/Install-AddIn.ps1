# PowerShell script til at installere Outlook VSTO Add-in
# Koer som Administrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Outlook Meeting Template Add-in Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check om Outlook koerer
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Host "ADVARSEL: Outlook koerer!" -ForegroundColor Yellow
    Write-Host "Luk venligst Outlook foer installation fortsaetter." -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "Vil du lukke Outlook nu? (J/N)"
    if ($response -eq "J" -or $response -eq "j") {
        Stop-Process -Name "OUTLOOK" -Force
        Write-Host "OK - Outlook lukket" -ForegroundColor Green
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Installation afbrudt." -ForegroundColor Red
        exit
    }
}

# Find DLL path
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath = Join-Path $scriptPath "bin\Release\OutlookMeetingAddin.dll"
$vstoPath = Join-Path $scriptPath "bin\Release\OutlookMeetingAddin.vsto"

# Check om filen eksisterer
if (-not (Test-Path $dllPath)) {
    Write-Host "FEJL: DLL ikke fundet!" -ForegroundColor Red
    Write-Host "Forventet placering: $dllPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Byg venligst projektet foerst:" -ForegroundColor Yellow
    Write-Host "  1. Aabn projektet i Visual Studio" -ForegroundColor White
    Write-Host "  2. Vaelg 'Release' configuration" -ForegroundColor White
    Write-Host "  3. Byg projektet (F6)" -ForegroundColor White
    exit
}

Write-Host "OK - DLL fundet: $dllPath" -ForegroundColor Green

# Opret registry noegle
$registryPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"

Write-Host "Opretter registry noegle..." -ForegroundColor Yellow

try {
    # Fjern eksisterende noegle hvis den findes
    if (Test-Path $registryPath) {
        Remove-Item -Path $registryPath -Recurse -Force
        Write-Host "OK - Gammel registry noegle fjernet" -ForegroundColor Green
    }

    # Opret ny noegle
    New-Item -Path $registryPath -Force | Out-Null

    # Saet vaerdier
    New-ItemProperty -Path $registryPath -Name "Description" -Value "Meeting Template Add-in" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Moede Skabelon" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

    # Konverter til file:/// format
    $manifestPath = "file:///" + $vstoPath.Replace("\", "/") + "|vstolocal"
    New-ItemProperty -Path $registryPath -Name "Manifest" -Value $manifestPath -PropertyType String -Force | Out-Null

    Write-Host "OK - Registry noegle oprettet" -ForegroundColor Green
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Installation fuldfoert!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Naeste trin:" -ForegroundColor Yellow
    Write-Host "1. Start Outlook" -ForegroundColor White
    Write-Host "2. Opret en ny moedeaftale" -ForegroundColor White
    Write-Host "3. Task pane vises automatisk!" -ForegroundColor White
    Write-Host ""

} catch {
    Write-Host "FEJL ved installation: $_" -ForegroundColor Red
    exit
}

# Tilbyd at starte Outlook
Write-Host ""
$response = Read-Host "Vil du starte Outlook nu? (J/N)"
if ($response -eq "J" -or $response -eq "j") {
    Start-Process "outlook.exe"
    Write-Host "Outlook startet" -ForegroundColor Green
}
