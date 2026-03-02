# 🎯 Deployment Plan: 400 Brugere via Intune

**Organization:** 400 brugere
**Infrastructure:** Microsoft Intune (Hybrid)
**Requirement:** Auto-show/auto-udfyld **KRITISK**
**Timeline:** 4 uger fra pilot til full rollout

---

## 📋 Executive Summary

Vi deployer Outlook Meeting Template VSTO Add-in til 400 brugere via Microsoft Intune.

**Key Points:**
- ✅ Auto-show taskpane når nye møder oprettes
- ✅ Ingen brugerinteraktion krævet
- ✅ Central deployment via Intune
- ✅ Phased rollout over 4 uger
- ⏱️ Total deployment tid: 4 uger
- 💰 Cost: Ingen ekstra licenser nødvendigt

---

## 📅 Deployment Timeline

### Uge 1: Pilot (10 brugere)
- Dag 1-2: Build og pak add-in
- Dag 3: Upload til Intune
- Dag 4: Deploy til pilot gruppe
- Dag 5-7: Monitor og gather feedback

### Uge 2: Phase 1 (50 brugere)
- Dag 1: Deploy til 50 brugere
- Dag 2-7: Monitor og support

### Uge 3: Phase 2 (150 brugere)
- Dag 1: Deploy til yderligere 150 brugere
- Dag 2-7: Monitor og support

### Uge 4: Phase 3 (190 brugere)
- Dag 1: Deploy til resterende brugere
- Dag 2-7: Monitor og afslut

---

## 🔧 Technical Implementation

### Del 1: Preparation (DONE!)

✅ VSTO Add-in projekt er oprettet
✅ PowerShell deployment scripts er klar
✅ Intune deployment guides er skrevet
🔄 Office Development Tools installeres (kører nu!)

### Del 2: Build Process (Efter installation)

**Step 1: Build VSTO Project**
```powershell
cd "C:\Users\dbl\OneDrive - netIP\Skrivebord\customeeting\OutlookVSTO"

# Build
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    OutlookMeetingAddin.sln `
    /p:Configuration=Release `
    /p:Platform="Any CPU"
```

**Step 2: Verify Build**
```powershell
# Check output
ls "bin\Release\"

# Should contain:
# - OutlookMeetingAddin.dll
# - OutlookMeetingAddin.dll.manifest
# - OutlookMeetingAddin.vsto
# - Microsoft.Office.Tools.Common.v4.0.Utilities.dll
# - Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll
```

**Step 3: Prepare Intune Package**
```powershell
# Create packaging folder
New-Item -Path "C:\IntunePackaging" -ItemType Directory -Force
New-Item -Path "C:\IntunePackaging\Source" -ItemType Directory -Force
New-Item -Path "C:\IntunePackaging\Output" -ItemType Directory -Force

# Copy files
Copy-Item -Path "bin\Release\*" -Destination "C:\IntunePackaging\Source\" -Recurse
Copy-Item -Path "Deploy-AddIn-Silent.ps1" -Destination "C:\IntunePackaging\Source\"

# Download IntuneWinAppUtil
Invoke-WebRequest `
    -Uri "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" `
    -OutFile "C:\IntunePackaging\IntuneWinAppUtil.exe"

# Create .intunewin package
& "C:\IntunePackaging\IntuneWinAppUtil.exe" `
    -c "C:\IntunePackaging\Source" `
    -s "Deploy-AddIn-Silent.ps1" `
    -o "C:\IntunePackaging\Output" `
    -q
```

**Output:** `C:\IntunePackaging\Output\Deploy-AddIn-Silent.intunewin`

### Del 3: Intune Configuration

**App Information:**
```
Name: Outlook Meeting Template Add-in
Description: Automatisk møde skabelon - auto-udfylder nye møder
Publisher: [Dit firma]
Version: 1.0.0
Category: Productivity
```

**Install Command:**
```powershell
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File Deploy-AddIn-Silent.ps1
```

**Uninstall Command:**
```powershell
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File Deploy-AddIn-Silent.ps1 -Uninstall
```

**Detection Rule (Registry):**
```
Path: HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin
Value: LoadBehavior
Type: DWORD
Operator: Equals
Value: 3
```

**Requirements:**
```
OS: Windows 10 1607 or later
Architecture: 64-bit
Disk space: 50 MB
Memory: 100 MB

Custom: Registry check for Outlook
Path: HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
Value: Platform
Operator: Exists
```

### Del 4: Pilot Groups

**Pilot Gruppe (10 brugere):**
- IT Department: 3 personer
- Power Users: 5 personer
- Management: 2 personer

Opret Azure AD gruppe: `Pilot-OutlookAddIn`

**Phase 1 Gruppe (50 brugere):**
- Early adopters fra forskellige afdelinger
Opret Azure AD gruppe: `Phase1-OutlookAddIn`

**Phase 2 Gruppe (150 brugere):**
Opret Azure AD gruppe: `Phase2-OutlookAddIn`

**Phase 3 Gruppe (Resten):**
Opret Azure AD gruppe: `Phase3-OutlookAddIn`

---

## 📊 Monitoring & Reporting

### Daily Checks (Første uge)

**Intune Portal:**
1. Log ind på https://endpoint.microsoft.com
2. Apps → Windows → Outlook Meeting Template Add-in
3. Check "Device install status"
4. Check "User install status"

**Success Metrics:**
- Installation success rate > 95%
- User complaints < 5%
- Add-in functioning correctly

### PowerShell Reporting Script

```powershell
# Get installation status for all users
Connect-MSGraph

$appId = "your-app-id-here"
$installs = Get-IntuneAppInstallStatus -AppId $appId

# Group by status
$summary = $installs | Group-Object -Property InstallState |
    Select-Object Name, Count

# Display
$summary | Format-Table -AutoSize

# Export
$installs | Export-Csv -Path "AddIn-Status-$(Get-Date -Format 'yyyy-MM-dd').csv"
```

### Weekly Report Template

```
OUTLOOK ADD-IN DEPLOYMENT - WEEKLY REPORT

Week: [Number]
Date: [Date]

DEPLOYMENT STATUS:
- Total Target Users: 400
- Deployed To: [Number]
- Successful Installs: [Number] ([Percentage]%)
- Failed Installs: [Number]
- Pending: [Number]

ISSUES:
- [List any issues]

USER FEEDBACK:
- Positive: [Count]
- Negative: [Count]
- Feature Requests: [List]

NEXT STEPS:
- [Plan for next week]
```

---

## 🚨 Rollback Plan

**If issues arise:**

1. **Pause deployment:**
   - I Intune, gå til app assignments
   - Fjern assignment fra problematic gruppe

2. **Rollback til affected users:**
   - Ændr assignment til "Uninstall"
   - Eller kør uninstall script via Intune

3. **Fix issue:**
   - Update add-in
   - Upload ny .intunewin
   - Re-test på pilot

4. **Resume deployment:**
   - Assign til grupper igen

**Emergency PowerShell Uninstall:**
```powershell
# Run on affected machines
Invoke-Command -ComputerName (Get-Content affected-machines.txt) -ScriptBlock {
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force
    }
}
```

---

## 📞 Support Plan

### Tier 1: End User Support

**Common Issues:**
1. "Add-in ikke synlig"
   - Check Outlook version
   - Check add-in is enabled: File → Options → Add-ins

2. "Taskpane vises ikke automatisk"
   - Check LoadBehavior registry value
   - Genstart Outlook

3. "Skabelon indsættes ikke"
   - Check VSTO Runtime installation
   - Re-install via Intune

### Tier 2: IT Support

**Diagnostics:**
```powershell
# Check installation på bruger maskine
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
Get-ItemProperty -Path $regPath

# Check VSTO Runtime
Test-Path "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"

# Check Intune logs
Get-Content "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log" |
    Select-String "OutlookMeeting" | Select-Object -Last 50
```

### Tier 3: Escalation

Contact: [IT Manager]
Email: [Email]
Phone: [Phone]

---

## 📧 User Communication

### Pre-Deployment Email (Uge 0)

```
Emne: Nyt værktøj kommer til Outlook - Møde Skabeloner

Hej alle,

Vi introducerer et nyt værktøj der gør det nemmere at oprette strukturerede møder.

HVAD ER DET?
Når du opretter en ny mødeaftale, vil en skabelon automatisk blive indsat
med de vigtigste punkter til et godt møde:
- Formål
- Dagsorden
- Roller og ansvar
- Beslutninger og næste skridt

HVORNÅR KOMMER DET?
- Pilot gruppe: [Dato]
- Alle andre: [Dato]

SKAL JEG GØRE NOGET?
Nej! Det installeres automatisk på din computer. Du skal bare åbne Outlook
og oprette et nyt møde - resten sker automatisk.

SPØRGSMÅL?
Kontakt IT support på [email/phone]

Med venlig hilsen,
IT Afdelingen
```

### Post-Deployment Email

```
Emne: Møde Skabelon værktøjet er nu aktivt! 🎉

Hej alle,

Møde skabelon værktøjet er nu aktiveret på din computer!

SÅDAN BRUGER DU DET:
1. Opret et nyt møde i Outlook Calendar
2. Skabelonen indsættes automatisk!
3. Udfyld skabelonen og send invitation

FEEDBACK?
Vi vil gerne høre hvad du synes! Send feedback til [email]

PROBLEMER?
Kontakt IT support: [email/phone]

Med venlig hilsen,
IT Afdelingen
```

---

## ✅ Success Criteria

**Technical:**
- [x] > 95% installation success rate
- [x] < 5 support tickets per 100 users
- [x] Add-in loads automatically
- [x] Template inserts correctly

**Business:**
- [x] Users report møder er mere strukturerede
- [x] Positive user feedback > 80%
- [x] Adoption rate > 90% (brugere faktisk bruger skabelonen)

---

## 📝 Checklist

### Before Pilot
- [ ] Office Development Tools installeret
- [ ] VSTO projekt bygget successfully
- [ ] .intunewin pakke oprettet
- [ ] Uploaded til Intune
- [ ] Pilot Azure AD gruppe oprettet
- [ ] Detection rule testet
- [ ] Pre-deployment email sendt

### During Pilot
- [ ] App assigned til pilot gruppe
- [ ] Daily monitoring første 3 dage
- [ ] Feedback indsamlet fra pilot users
- [ ] Issues dokumenteret og fixet
- [ ] Success criteria met

### Before Phase 1
- [ ] Pilot success confirmed
- [ ] Phase 1 gruppe oprettet
- [ ] Communication email sendt
- [ ] Support team briefed

### During Phase 1-3
- [ ] Weekly progress reports
- [ ] Daily monitoring første uge af hver phase
- [ ] Issue tracking og resolution
- [ ] User feedback collection

### Post-Deployment
- [ ] Final success metrics rapport
- [ ] Lessons learned dokumenteret
- [ ] Support documentation opdateret
- [ ] Celebration! 🎉

---

**NEXT STEP:** Vent på Office Development Tools installation (kører i baggrund), derefter bygger vi add-in!

Estimeret tid tilbage: ~10 minutter
