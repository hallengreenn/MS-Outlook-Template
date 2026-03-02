# Outlook VSTO Meeting Template Add-in

Dette er et **COM-baseret VSTO add-in** til Outlook der automatisk viser en møde skabelon task pane når du opretter nye kalenderbegivenheder.

## ✨ Funktioner

- ✅ **Automatisk visning** - Task pane vises automatisk når du opretter en ny mødeaftale
- ✅ **Indsæt skabelon** - Klik på en knap for at indsætte møde skabelonen
- ✅ **Dansk skabelon** - Formål, dagsorden, roller, beslutninger, evt.
- ✅ **Native Windows** - Fuld integration med Outlook via COM

## 📋 Krav

For at bygge og installere dette add-in skal du have:

1. **Visual Studio 2019 eller 2022** (Community, Professional, eller Enterprise)
2. **Office Development Tools** (installeres via Visual Studio Installer)
3. **Microsoft Office 2016 eller nyere** (Desktop version - IKKE Office 365 Web)
4. **.NET Framework 4.7.2** eller nyere

## 🛠️ Installation af Visual Studio komponenter

1. Åbn **Visual Studio Installer**
2. Klik på **Modify** ved din Visual Studio installation
3. Under **Workloads**, vælg:
   - ✅ **Office/SharePoint development**
4. Klik **Modify** for at installere

## 🚀 Sådan bygger du projektet

### Metode 1: Via Visual Studio GUI (ANBEFALET)

1. Dobbeltklik på `OutlookMeetingAddin.csproj` for at åbne projektet i Visual Studio
2. Når projektet er åbnet, gå til **Build** → **Build Solution** (eller tryk F6)
3. Vent til build er færdig (check Output vinduet)

### Metode 2: Via Command Line

```cmd
cd "C:\Users\dbl\OneDrive - netIP\Skrivebord\customeeting\OutlookVSTO"

"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookMeetingAddin.csproj /p:Configuration=Release
```

Hvis du har Visual Studio 2019:
```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookMeetingAddin.csproj /p:Configuration=Release
```

## 📦 Installation af Add-in

### Automatisk Installation (via Visual Studio)

1. Sæt projektet som startup projekt
2. Tryk **F5** for at debugge
3. Outlook vil åbne med add-in'et installeret

### Manuel Installation via Registry

**⚠️ VIGTIGT: Luk Outlook før du kører dette!**

Opret en `.reg` fil med følgende indhold:

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin]
"Description"="Meeting Template Add-in"
"FriendlyName"="Møde Skabelon"
"LoadBehavior"=dword:00000003
"Manifest"="file:///C:/Users/dbl/OneDrive - netIP/Skrivebord/customeeting/OutlookVSTO/bin/Release/OutlookMeetingAddin.vsto|vstolocal"
```

**Dobbeltklik på .reg filen** for at tilføje til registry.

### Manuel Installation via ClickOnce

1. Byg projektet i **Release** mode
2. Gå til `bin\Release\` mappen
3. Find `OutlookMeetingAddin.vsto` filen
4. Dobbeltklik på den for at installere

## 🧪 Test af Add-in

1. **Start Outlook**
2. **Opret en ny mødeaftale**:
   - Gå til Calendar
   - Klik "New Meeting" eller "New Appointment"
3. **Task pane vises automatisk!**
   - Du skulle nu se "Møde Skabelon" task pane på højre side
4. **Klik på knappen** "INDSÆT MØDE SKABELON"
5. **Skabelonen indsættes** i møde beskrivelsen

## 🔧 Fejlfinding

### Add-in vises ikke i Outlook

1. **Check om add-in er indlæst:**
   - I Outlook: Gå til **File** → **Options** → **Add-ins**
   - Find "Møde Skabelon" i listen
   - Hvis den er under "Disabled Items", skal du aktivere den

2. **Check LoadBehavior i Registry:**
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin
   LoadBehavior = 3
   ```
   - Hvis den er 0 eller 2, sæt den til 3

3. **Check VSTO Runtime:**
   - Download og installer: [Microsoft Visual Studio Tools for Office Runtime](https://aka.ms/VSTORuntimeDownload)

### Build fejl

**Fejl: "Could not find Office PIA"**

Løsning:
1. Åbn Visual Studio Installer
2. Modify din installation
3. Vælg "Office/SharePoint development" workload
4. Geninstaller

**Fejl: "Target framework not installed"**

Løsning:
1. Download .NET Framework 4.7.2 Developer Pack
2. Eller rediger `.csproj` og skift `TargetFrameworkVersion` til den version du har

### Task pane vises ikke automatisk

1. Check at koden i `ThisAddIn.cs` har:
   ```csharp
   taskPane.Visible = true;
   ```

2. Debug i Visual Studio og sæt breakpoint i `Inspectors_NewInspector` metoden

## 📝 Sådan redigerer du skabelonen

Åbn `MeetingTaskPane.cs` og rediger `MEETING_TEMPLATE` konstanten:

```csharp
private const string MEETING_TEMPLATE = @"
<p><strong>Dit nye afsnit</strong></p>
<p>Din nye tekst</p>
<br>
";
```

Gem og byg projektet igen.

## 🆚 Forskelle fra Web Add-in

| Feature | Web Add-in | VSTO COM Add-in |
|---------|-----------|-----------------|
| Automatisk visning | ❌ Nej | ✅ Ja |
| Platform support | Windows/Mac/Web | Kun Windows |
| Installation | Centraliseret | Lokal |
| Opdateringer | Automatisk | Manuel |
| Teknologi | HTML/JS | C#/WinForms |

## 📂 Projekt Struktur

```
OutlookVSTO/
├── OutlookMeetingAddin.csproj    # Projekt fil
├── ThisAddIn.cs                   # Hoved add-in logik
├── MeetingTaskPane.cs             # Task pane UI logik
├── MeetingTaskPane.Designer.cs    # WinForms designer kode
├── MeetingTaskPane.resx           # Ressourcer
├── Properties/
│   └── AssemblyInfo.cs           # Assembly metadata
└── README.md                      # Denne fil
```

## 🔐 Sikkerhed

VSTO add-ins kræver:
- Trusted location eller signeret manifest
- VSTO Runtime installeret
- .NET Framework

## 📞 Support

Hvis du har problemer:
1. Check Visual Studio Output vinduet for fejlmeddelelser
2. Check Windows Event Viewer under "Application"
3. Aktivér VSTO logging i registry

## 📜 Licens

Dette projekt er open source og kan modificeres frit.
