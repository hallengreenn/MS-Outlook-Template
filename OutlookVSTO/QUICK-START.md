# 🚀 Quick Start Guide - Dansk

Denne guide hjælper dig med at komme i gang med Outlook Meeting Template VSTO Add-in.

## ⚡ Hurtig Installation (5 minutter)

### Trin 1: Installer Visual Studio (hvis ikke allerede installeret)

1. Download [Visual Studio 2022 Community](https://visualstudio.microsoft.com/downloads/) (gratis)
2. Under installation, vælg workload: **"Office/SharePoint development"**
3. Vent til installation er færdig

### Trin 2: Byg Projektet

**Metode A: Via Visual Studio (Anbefalet)**
1. Dobbeltklik på `OutlookMeetingAddin.csproj`
2. Når Visual Studio åbner, tryk **F6** (eller gå til Build → Build Solution)
3. Vent til du ser "Build succeeded" i Output vinduet

**Metode B: Via PowerShell**
```powershell
cd "C:\Users\dbl\OneDrive - netIP\Skrivebord\customeeting\OutlookVSTO"

# For Visual Studio 2022:
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookMeetingAddin.csproj /p:Configuration=Release
```

### Trin 3: Installer Add-in

**⚠️ VIGTIGT: Luk Outlook først!**

Derefter kør én af følgende:

**Metode A: PowerShell Script (Nemmest)**
```powershell
# Højreklik på PowerShell og vælg "Run as Administrator"
cd "C:\Users\dbl\OneDrive - netIP\Skrivebord\customeeting\OutlookVSTO"
.\Install-AddIn.ps1
```

**Metode B: Via Visual Studio**
1. I Visual Studio, tryk **F5**
2. Outlook åbner automatisk med add-in'et installeret

### Trin 4: Test Det!

1. Start Outlook (hvis ikke allerede startet)
2. Gå til **Calendar**
3. Klik **"New Meeting"** eller **"New Appointment"**
4. 🎉 **Task pane vises automatisk på højre side!**
5. Klik på knappen **"INDSÆT MØDE SKABELON"**
6. Skabelonen indsættes i mødet!

## ✅ Sådan ser det ud

```
┌─────────────────────────────────────────┐
│ Outlook - Ny Mødeaftale                 │
├─────────────────────────────────────────┤
│                              ┌──────────┐│
│  Emne: [..................] │ 📋 Møde  ││
│  Sted:  [..................] │ Skabelon ││
│  Start: [Date] [Time]        │          ││
│                              │ Preview: ││
│  Beskrivelse:                │          ││
│  ┌─────────────────────────┐ │ • Formål││
│  │                         │ │ • Agenda││
│  │                         │ │ • Roller││
│  │                         │ │          ││
│  │                         │ │ [INDSÆT] ││
│  └─────────────────────────┘ └──────────┘│
└─────────────────────────────────────────┘
         ↑                          ↑
    Møde indhold          Task pane (vises AUTO!)
```

## 🎯 Hvad Gør Dette Add-in?

1. **Overvåger** når du opretter nye møder
2. **Viser automatisk** en task pane med skabelon preview
3. **Indsætter** dansk møde skabelon med ét klik:
   - Formål med mødet
   - Dagsorden/emner
   - Roller og ansvar
   - Beslutninger og næste skridt
   - Evt.

## 🔧 Fejlfinding

### "Build failed" fejl

**Problem:** Mangler Office development tools

**Løsning:**
1. Åbn Visual Studio Installer
2. Klik "Modify"
3. Vælg "Office/SharePoint development"
4. Klik "Modify" for at installere

### Add-in vises ikke i Outlook

**Løsning 1:** Check om det er aktiveret
1. I Outlook: File → Options → Add-ins
2. Find "Møde Skabelon"
3. Hvis det er under "Disabled", aktivér det

**Løsning 2:** Installer VSTO Runtime
- Download: https://aka.ms/VSTORuntimeDownload

**Løsning 3:** Geninstaller
```powershell
.\Uninstall-AddIn.ps1
.\Install-AddIn.ps1
```

### Task pane vises ikke automatisk

**Check:**
1. Gå til Outlook ribbon
2. Find add-in under "Add-ins" tab
3. Manuelt klik på det for at åbne
4. Check at du opretter en **APPOINTMENT** og ikke bare en task

## 📝 Tilpas Skabelonen

Åbn `MeetingTaskPane.cs` og find linjen:

```csharp
private const string MEETING_TEMPLATE = @"
<p><strong>Formål med mødet</strong></p>
<p>Kort beskrivelse...</p>
...
```

Rediger teksten, gem, og byg projektet igen (F6).

## 🗑️ Afinstallation

Kør:
```powershell
.\Uninstall-AddIn.ps1
```

Eller manuelt:
1. Luk Outlook
2. Slet registry nøglen:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin
   ```

## 💡 Tips

1. **Pin task pane:** Task pane forbliver åben mellem møder
2. **Rediger efter indsættelse:** Skabelonen er bare en start - rediger frit
3. **Virker med invitations:** Fungerer både for appointments og meeting requests

## 📞 Har Du Problemer?

Check disse ting:
- [ ] Er Visual Studio Office Development Tools installeret?
- [ ] Er projektet bygget uden fejl?
- [ ] Er Outlook lukket når du installerer?
- [ ] Er VSTO Runtime installeret?
- [ ] Bruger du Desktop Outlook (ikke Web)?

## 🎓 Næste Trin

- Læs den fulde [README.md](README.md) for flere detaljer
- Tilpas skabelonen til dine behov
- Installer på flere maskiner i din organisation

---

**God fornøjelse! 🎉**
