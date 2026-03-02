# VBA Enterprise Deployment - Ingen Build Nødvendig! 🚀

**For 400 brugere via Intune - NEMMERE end VSTO!**

## ✅ Fordele over VSTO

| | VBA | VSTO |
|---|---|---|
| **Auto-udfyld** | ✅ JA - Indsætter direkte | ✅ JA - Via taskpane |
| **Build krævet** | ❌ NEJ! | ✅ Ja - Kræver Visual Studio |
| **Installation tid** | 2 min | 3+ timer |
| **Dependencies** | Ingen | VSTO Runtime (30MB) |
| **Intune deployment** | ✅ Meget simpel | ⚠️ Kompleks |
| **Maintenance** | Meget lav | Medium |

## 🎯 Anbefaling for Jeres 400 Brugere

**Brug VBA deployment!** Det opfylder jeres KRITISKE krav (auto-udfyld) og er **10x nemmere** at deploye.

---

## 📦 Komplet VBA Deployment Guide

### Del 1: Opret VBA Macro (5 minutter)

#### Step 1: Eksporter VBA Template

Jeg har allerede lavet koden - du skal bare eksportere den:

1. **Åbn Outlook** på din maskine
2. **Tryk Alt+F11** (åbner VBA Editor)
3. **Dobbeltklik "ThisOutlookSession"** i Project Explorer
4. **Kopier denne kode** (fra `MeetingTemplate.vba`):

```vba
Private WithEvents AppointmentInspectors As Outlook.Inspectors

Private Sub Application_Startup()
    Set AppointmentInspectors = Application.Inspectors
End Sub

Private Sub AppointmentInspectors_NewInspector(ByVal Inspector As Inspector)
    On Error Resume Next
    If TypeOf Inspector.CurrentItem Is AppointmentItem Then
        Dim appt As AppointmentItem
        Set appt = Inspector.CurrentItem
        If Len(Trim(appt.Body)) = 0 Then
            appt.Body = BuildMeetingTemplate()
        End If
    End If
End Sub

Private Function BuildMeetingTemplate() As String
    Dim template As String
    template = "═══════════════════════════════════════════════════" & vbCrLf
    template = template & "           MØDE SKABELON" & vbCrLf
    template = template & "═══════════════════════════════════════════════════" & vbCrLf
    template = template & vbCrLf
    template = template & "📋 FORMÅL MED MØDET" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Kort beskrivelse af, hvorfor vi mødes, og hvad vi skal opnå." & vbCrLf
    template = template & vbCrLf
    template = template & "📝 DAGSORDEN/EMNER" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Punkt 1: " & vbCrLf
    template = template & "• Punkt 2: " & vbCrLf
    template = template & "• Punkt 3: " & vbCrLf
    template = template & vbCrLf
    template = template & "👥 ROLLER OG ANSVAR" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Mødeleder: " & vbCrLf
    template = template & "Referent: " & vbCrLf
    template = template & vbCrLf
    template = template & "✅ BESLUTNINGER OG NÆSTE SKRIDT" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Beslutning 1: " & vbCrLf
    template = template & "• Action item 1: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & vbCrLf
    template = template & "💡 EVT." & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Tid til spørgsmål eller andre punkter." & vbCrLf
    template = template & vbCrLf
    template = template & "═══════════════════════════════════════════════════" & vbCrLf
    BuildMeetingTemplate = template
End Function
```

5. **Gem** (Ctrl+S)
6. **Luk Outlook**
7. **Find VBA fil:**
   - Gå til: `%APPDATA%\Microsoft\Outlook`
   - Find filen `VbaProject.OTM`
8. **Kopier den** til `C:\IntunePackaging\VbaProject.OTM`

---

### Del 2: Opret Intune Deployment (10 minutter)

#### Step 1: Opret PowerShell Deployment Script

Gem denne som `Deploy-VBA.ps1`:

```powershell
#Requires -RunAsAdministrator
# VBA Meeting Template Deployment for Intune

$LogPath = "$env:TEMP\VBA-Deploy.log"
$SourceVBA = "VbaProject.OTM"  # Skal være i samme folder som script
$TargetPath = "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogPath -Value "$timestamp - $Message"
    Write-Host $Message
}

try {
    Write-Log "Starting VBA deployment..."

    # Stop Outlook hvis kører
    $outlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlook) {
        Write-Log "Stopping Outlook..."
        Stop-Process -Name "OUTLOOK" -Force
        Start-Sleep -Seconds 3
    }

    # Ensure Outlook folder exists
    $outlookFolder = Split-Path -Parent $TargetPath
    if (-not (Test-Path $outlookFolder)) {
        New-Item -Path $outlookFolder -ItemType Directory -Force | Out-Null
    }

    # Copy VBA file
    $scriptPath = Split-Path -Parent $PSCommandPath
    $sourceFile = Join-Path $scriptPath $SourceVBA

    if (-not (Test-Path $sourceFile)) {
        throw "VBA file not found: $sourceFile"
    }

    Copy-Item -Path $sourceFile -Destination $TargetPath -Force
    Write-Log "VBA file deployed to $TargetPath"

    # Enable macros (set to prompt)
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "Level" -Value 2 -Type DWord -Force
    Write-Log "Macro security configured"

    Write-Log "Deployment completed successfully!"
    exit 0

} catch {
    Write-Log "ERROR: $_"
    exit 1
}
```

#### Step 2: Pak til Intune

```powershell
# Opret folders
New-Item -Path "C:\IntunePackaging\VBA\Source" -ItemType Directory -Force
New-Item -Path "C:\IntunePackaging\VBA\Output" -ItemType Directory -Force

# Kopier filer
Copy-Item -Path "C:\IntunePackaging\VbaProject.OTM" -Destination "C:\IntunePackaging\VBA\Source\"
Copy-Item -Path "Deploy-VBA.ps1" -Destination "C:\IntunePackaging\VBA\Source\"

# Download IntuneWinAppUtil (hvis ikke allerede gjort)
if (-not (Test-Path "C:\IntunePackaging\IntuneWinAppUtil.exe")) {
    Invoke-WebRequest `
        -Uri "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" `
        -OutFile "C:\IntunePackaging\IntuneWinAppUtil.exe"
}

# Pak til .intunewin
& "C:\IntunePackaging\IntuneWinAppUtil.exe" `
    -c "C:\IntunePackaging\VBA\Source" `
    -s "Deploy-VBA.ps1" `
    -o "C:\IntunePackaging\VBA\Output" `
    -q
```

**Output:** `C:\IntunePackaging\VBA\Output\Deploy-VBA.intunewin`

---

### Del 3: Upload til Intune

#### Intune App Configuration

**App Information:**
```
Name: Outlook Meeting Template (VBA)
Description: Automatisk møde skabelon - auto-udfylder nye møder
Publisher: [Dit firma]
Version: 1.0.0
Category: Productivity
```

**Program:**
```
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File Deploy-VBA.ps1

Uninstall command:
powershell.exe -Command "Remove-Item '$env:APPDATA\Microsoft\Outlook\VbaProject.OTM' -Force"

Install behavior: User
Device restart behavior: No specific action
```

**Requirements:**
```
Operating system: Windows 10 1607+
Architecture: 64-bit

Custom requirement:
Rule type: Registry
Key path: HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
Value name: Platform
Operator: Exists
```

**Detection Rule:**
```
Rule type: File
Path: %APPDATA%\Microsoft\Outlook
File: VbaProject.OTM
Detection method: File or folder exists
```

**Assignments:**
- Required: [Din bruger gruppe - 400 brugere]

---

### Del 4: Bruger Kommunikation

**Pre-deployment Email:**

```
Emne: Nyt værktøj: Automatisk Møde Skabelon

Hej alle,

Vi indfører et nyt værktøj der automatisk indsætter en struktureret
skabelon når du opretter møder i Outlook.

HVAD SKER DER?
Næste gang du logger ind og åbner Outlook, kan du få en sikkerhedsbesked
om macros. Klik "Enable Macros" - det er safe!

HVORDAN VIRKER DET?
1. Opret et nyt møde i Outlook
2. Skabelonen indsættes automatisk i beskrivelsen
3. Udfyld punkterne og send invitation

HVORNÅR?
Installeres automatisk: [Dato]

SPØRGSMÅL?
Kontakt IT: [email/phone]
```

**VIGTIGT:**
Første gang brugere åbner Outlook efter deployment, får de en sikkerhedsbesked:

```
Microsoft Outlook
────────────────────
En macro forsøger at køre.
Vil du aktivere macros?

[Enable Macros]  [Disable Macros]
```

Brugere SKAL klikke **"Enable Macros"**.

**Dette sker kun én gang!**

---

### Del 5: Monitoring

**Check deployment status i Intune:**
1. https://endpoint.microsoft.com
2. Apps → Windows → Outlook Meeting Template (VBA)
3. Device install status

**PowerShell reporting:**
```powershell
# Check på klient maskiner
Invoke-Command -ComputerName PC001,PC002 -ScriptBlock {
    $vbaPath = "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM"
    if (Test-Path $vbaPath) {
        Write-Host "$env:COMPUTERNAME - VBA Deployed ✅"
    } else {
        Write-Host "$env:COMPUTERNAME - NOT Deployed ❌" -ForegroundColor Red
    }
}
```

---

## 🎯 Deployment Timeline for 400 Brugere

**Total tid:** 2-3 timer + 4 ugers phased rollout

### Preparation (1 time)
- [ ] Eksporter VBA template (5 min)
- [ ] Opret deployment script (5 min)
- [ ] Pak til .intunewin (5 min)
- [ ] Upload til Intune (10 min)
- [ ] Konfigurer app (10 min)
- [ ] Test på 1 maskine (30 min)

### Rollout (4 uger)
- **Uge 1:** Pilot (10 brugere) + monitor
- **Uge 2:** Phase 1 (50 brugere)
- **Uge 3:** Phase 2 (150 brugere)
- **Uge 4:** Phase 3 (190 brugere)

---

## ⚠️ Forskel på VBA vs VSTO

### VBA
- ✅ Auto-udfylder body direkte
- ✅ Simpel deployment
- ⚠️ Kræver "Enable Macros" første gang
- ✅ Ingen dependencies

### VSTO
- ✅ Task pane med knap (mere "professionelt")
- ✅ Ingen macro security prompt
- ❌ Kræver build proces
- ❌ Kræver VSTO Runtime (30MB)
- ❌ Mere kompleks deployment

**Min anbefaling:** For 400 brugere, **VBA er bedre** fordi:
1. Hurtigere at deploye (i dag vs næste uge)
2. Færre dependencies
3. Nemmere at vedligeholde
4. Opfylder jeres KRITISKE krav (auto-udfyld)

---

## 🚀 Næste Skridt

**Vælg én af to veje:**

### Option A: VBA (ANBEFALET)
1. Eksporter VBA template (5 min) - se Step 1 ovenfor
2. Pak til Intune (10 min)
3. Deploy i dag! 🎉

### Option B: VSTO
1. Installer Office Development Tools manuelt (se `Install-OfficeTools.ps1`)
2. Vent 15 min på installation
3. Build projekt
4. Pak til Intune
5. Deploy

**Hvad foretrækker du?**
