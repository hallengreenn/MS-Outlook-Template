# Microsoft Intune Deployment Guide

Denne guide viser hvordan du deployer Outlook Meeting Template Add-in via Microsoft Intune (Microsoft Endpoint Manager) til alle Windows enheder i din organisation.

## 📋 Forudsætninger

- ✅ Microsoft Intune licens (inkluderet i Microsoft 365 E3/E5)
- ✅ Enheder enrolled i Intune
- ✅ Global Administrator eller Intune Administrator rolle
- ✅ VSTO add-in bygget og klar
- ✅ Microsoft Win32 Content Prep Tool

## 🔧 Del 1: Forbered Installationspakken

### Trin 1: Download IntuneWinAppUtil

```powershell
# Download Microsoft Win32 Content Prep Tool
$url = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
Invoke-WebRequest -Uri $url -OutFile "IntuneWinAppUtil.exe"
```

### Trin 2: Opret Source Folder

```powershell
# Opret folder struktur
New-Item -Path "C:\IntunePackaging\Source" -ItemType Directory -Force
New-Item -Path "C:\IntunePackaging\Output" -ItemType Directory -Force

# Kopier add-in filer til Source
Copy-Item -Path ".\bin\Release\*" -Destination "C:\IntunePackaging\Source\" -Recurse

# Kopier deployment script
Copy-Item -Path ".\Deploy-AddIn-Silent.ps1" -Destination "C:\IntunePackaging\Source\"
```

Din source folder skulle se sådan ud:
```
C:\IntunePackaging\Source\
├── OutlookMeetingAddin.dll
├── OutlookMeetingAddin.dll.manifest
├── OutlookMeetingAddin.vsto
├── Deploy-AddIn-Silent.ps1
└── [andre DLL dependencies]
```

### Trin 3: Opret .intunewin Pakke

```powershell
# Kør IntuneWinAppUtil
.\IntuneWinAppUtil.exe `
    -c "C:\IntunePackaging\Source" `
    -s "Deploy-AddIn-Silent.ps1" `
    -o "C:\IntunePackaging\Output"
```

Dette opretter: `C:\IntunePackaging\Output\Deploy-AddIn-Silent.intunewin`

## 🚀 Del 2: Upload til Intune

### Trin 1: Åbn Microsoft Endpoint Manager Admin Center

1. Gå til https://endpoint.microsoft.com
2. Log ind med admin credentials

### Trin 2: Opret Win32 App

1. **Naviger til:**
   - Apps → Windows → Add → Windows app (Win32)

2. **App information:**
   - Klik "Select app package file"
   - Upload `Deploy-AddIn-Silent.intunewin`
   - Klik OK

3. **Udfyld App information:**
   ```
   Name: Outlook Meeting Template Add-in
   Description: Automatisk møde skabelon til Outlook
   Publisher: [Dit firma navn]
   App Version: 1.0.0
   Category: Productivity
   Show this as a featured app in Company Portal: No
   Information URL: (optional)
   Privacy URL: (optional)
   Developer: [Dit navn]
   Owner: IT Department
   Notes: Auto-deployed via Intune
   ```

4. **Upload logo (optional):**
   - 512x512 PNG billede
   - Klik Next

### Trin 3: Konfigurer Program Settings

```
Install command:
powershell.exe -ExecutionPolicy Bypass -File Deploy-AddIn-Silent.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -File Deploy-AddIn-Silent.ps1 -Uninstall

Install behavior: System

Device restart behavior: App install may force a device restart

Return codes:
0 = Success
1 = Failed
3010 = Soft reboot
1641 = Hard reboot
```

Klik Next.

### Trin 4: Konfigurer Requirements

```
Operating system architecture: 64-bit
Minimum operating system: Windows 10 1607

Additional requirements: (optional)
- Disk space required (MB): 50
- Physical memory required (MB): 100
- Minimum number of processors required: 1
- Minimum CPU speed required (MHz): 1000
```

**Custom requirement rules:**

Tilføj rule for at checke om Outlook er installeret:
```
Rule type: Registry
Key path: HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
Value name: Platform
Operator: Exists
```

Klik Next.

### Trin 5: Konfigurer Detection Rules

**Option 1: Registry Detection (Anbefalet)**

```
Rule type: Registry
Key path: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin
Value name: FriendlyName
Detection method: String comparison
Operator: Equals
Value: Møde Skabelon
Associated with a 32-bit app on 64-bit clients: No
```

**Option 2: File Detection**

```
Rule type: File
Path: C:\Program Files\OutlookMeetingAddin
File or folder: OutlookMeetingAddin.dll
Detection method: File or folder exists
Associated with a 32-bit app on 64-bit clients: No
```

**Option 3: Custom PowerShell Script**

Opret `Detect-AddIn.ps1`:
```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
if (Test-Path $regPath) {
    $loadBehavior = Get-ItemProperty -Path $regPath -Name "LoadBehavior" -ErrorAction SilentlyContinue
    if ($loadBehavior.LoadBehavior -eq 3) {
        Write-Output "Installed"
        exit 0
    }
}
exit 1
```

Upload scriptet og vælg:
```
Run script as 32-bit process on 64-bit clients: No
Enforce script signature check: No
Run script in 64-bit PowerShell Host: Yes
```

Klik Next.

### Trin 6: Konfigurer Dependencies

Hvis VSTO Runtime skal installeres først:

1. Upload VSTO Runtime som separat Win32 app først
2. Tilføj som dependency her

Eller skip dette hvis dit install script håndterer VSTO Runtime.

Klik Next.

### Trin 7: Konfigurer Supersedence

Skip dette (ingen tidligere version).

Klik Next.

### Trin 8: Assignments

**Required (Automatisk installation):**

1. Klik "Add group" under Required
2. Vælg security group (fx "All Users" eller "Outlook Users")
3. Konfigurer assignment:
   ```
   Make this app required for all devices: No
   Notification: Show toast notification
   Filter: None
   ```

**Available for enrolled devices (Valgfri installation):**

Skip dette hvis du vil forced deployment.

**Uninstall:**

Hvis du vil afinstallere fra specifikke grupper, tilføj dem her.

Klik Next.

### Trin 9: Review + Create

1. Gennemgå alle indstillinger
2. Klik "Create"
3. Vent på upload (kan tage 5-10 minutter)

## 📊 Del 3: Monitor Deployment

### Check Deployment Status

1. **I Intune Admin Center:**
   - Apps → Windows → Outlook Meeting Template Add-in
   - Klik "Device install status"
   - Se installation progress

2. **View by user:**
   - Klik "User install status"

### Check på Klient Enhed

**Via Company Portal App:**
- Åbn Company Portal
- Gå til Apps
- Find "Outlook Meeting Template Add-in"
- Status vises

**Via PowerShell:**
```powershell
# Check Intune Management Extension logs
Get-Content "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log" | Select-String "OutlookMeeting"

# Check installation status
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
```

## 🔄 Del 4: Updates

### Deploy Ny Version

1. **Build ny version** med opdateret version nummer

2. **Opret ny .intunewin pakke:**
   ```powershell
   .\IntuneWinAppUtil.exe -c "C:\IntunePackaging\Source" -s "Deploy-AddIn-Silent.ps1" -o "C:\IntunePackaging\Output"
   ```

3. **I Intune:**
   - Gå til din app
   - Klik "Properties"
   - Ved "App package file" klik Edit
   - Upload ny .intunewin fil
   - Opdater version nummer
   - Save

4. **Devices vil automatisk opdatere ved næste sync** (typisk inden for 8 timer)

### Force Sync

På klient maskine:
```
Settings → Accounts → Access work or school → [Din konto] → Info → Sync
```

## 🗑️ Del 5: Afinstallation

### Option 1: Ændr Assignment

1. Gå til app → Assignments
2. Flyt grupper fra "Required" til "Uninstall"
3. Save

### Option 2: Delete App

1. Marker app
2. Klik Delete
3. Confirm

Devices vil automatisk afinstallere ved næste sync.

## 📱 Del 6: Reporting

### Opret Custom Report

```powershell
# Connect to Microsoft Graph
Install-Module Microsoft.Graph.Intune -Scope CurrentUser
Connect-MSGraph

# Get app installation status
$appId = "your-app-id"
$installs = Get-IntuneAppInstallStatus -AppId $appId

# Export to CSV
$installs | Export-Csv -Path "AddIn-Intune-Status.csv" -NoTypeInformation
```

### Monitor via Log Analytics

Konfigurer Intune diagnostics til Log Analytics workspace for avanceret reporting.

## 🎯 Best Practices

1. **Pilot først:**
   - Deploy til pilot gruppe (5-10 brugere)
   - Test i 1-2 uger
   - Gather feedback

2. **Phased rollout:**
   - Uge 1: 10% af brugere
   - Uge 2: 25% af brugere
   - Uge 3: 50% af brugere
   - Uge 4: 100% af brugere

3. **Communication:**
   - Send email til brugere før deployment
   - Inkluder quick start guide
   - Tilbyd support kanal

4. **Monitoring:**
   - Check daily i første uge
   - Review failure reports
   - Address issues promptly

## 🔧 Troubleshooting

### App Installation Fails

**Check logs på klient:**
```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\
```

**Common issues:**
- VSTO Runtime mangler → Deploy separat
- Insufficient permissions → Check script execution policy
- Outlook ikke installeret → Update requirements rule
- Registry writes fejler → Ensure script runs as SYSTEM

### App Shows as Installed but Doesn't Work

1. Check registry manually
2. Check Outlook add-in list (File → Options → Add-ins)
3. Check LoadBehavior value (should be 3)
4. Check VSTO Runtime installation

### Detection Rule Fails

Test detection script manuelt:
```powershell
# Run as SYSTEM context (download PsExec)
PsExec.exe -s -i powershell.exe
# Then run detection script
```

## 📞 Support

Ved problemer:
- Check Intune Admin Center → Troubleshooting
- Review klient logs
- Contact Microsoft Support (Intune ticket)
- Check Microsoft Docs: https://docs.microsoft.com/mem/intune/

---

**Comparison:** Se [ENTERPRISE-DEPLOYMENT.md](ENTERPRISE-DEPLOYMENT.md) for sammenligning med andre deployment metoder.
