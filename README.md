# Outlook Meeting Template Add-in

Et professionelt Outlook add-in der automatisk indsætter møde-skabeloner når du opretter nye kalenderbegivenheder.

## Indholdsfortegnelse

- [Om projektet](#om-projektet)
- [Funktioner](#funktioner)
- [To løsninger](#to-løsninger)
- [Kom hurtigt i gang](#kom-hurtigt-i-gang)
- [Projekt struktur](#projekt-struktur)
- [Dokumentation](#dokumentation)
- [Deployment](#deployment)

## Om projektet

Dette projekt gør det nemt at standardisere møde-dokumentation i din organisation ved automatisk at indsætte en struktureret skabelon i alle nye mødeaftaler i Outlook.

**Standard skabelon indeholder:**
- Formål med mødet
- Dagsorden/emner
- Roller og ansvar
- Beslutninger og næste skridt
- Evt.

## Funktioner

- **Automatisk aktivering** - Skabelonen indsættes automatisk når du opretter en ny mødeaftale
- **Event-baseret** - Bruger Office.js LaunchEvent API for hurtig respons
- **Fuld HTML support** - Skabelonen understøtter formatering, lister, tabeller osv.
- **Cross-platform** - Virker i Outlook Desktop, Web og Mac (når deployed til HTTPS)
- **Enterprise-klar** - Kan deployes centraliseret via Microsoft 365 Admin Center eller Intune

## To løsninger

Dette repository indeholder to forskellige implementeringer:

### 1. Web Add-in (Anbefalet) ✅

**Placering:** Rod-mappen + `/deploy`
**Teknologi:** Office.js (HTML/JavaScript)

**Fordele:**
- Cross-platform (Windows, Mac, Web, Mobile)
- Automatiske opdateringer
- Centraliseret deployment
- Nemmere at vedligeholde

**Brug denne løsning hvis:**
- Du vil have cross-platform support
- Du vil deploye til mange brugere
- Du vil have automatiske opdateringer

### 2. VSTO COM Add-in (Legacy)

**Placering:** `/OutlookVSTO`
**Teknologi:** C# VSTO/COM

**Fordele:**
- Dybere Windows integration
- Mere kontrol over Outlook funktioner
- Virker offline

**Brug denne løsning hvis:**
- Du kun har Windows Desktop Outlook
- Du har specifikke COM integration behov
- Du har legacy requirements

**Se `/OutlookVSTO/README.md` for VSTO dokumentation.**

## Kom hurtigt i gang

### Forudsætninger

- Node.js installeret (https://nodejs.org)
- Outlook Desktop (Microsoft 365 eller Office 2016+)
- Code editor (VS Code anbefales)

### Lokal test i 3 trin

1. **Installer dependencies**
   ```bash
   npm install
   ```

2. **Start lokal server**
   ```bash
   node server.js
   ```
   Serveren kører nu på http://localhost:3000

3. **Tilføj add-in til Outlook**
   - Åbn Outlook Desktop
   - Gå til **File** → **Get Add-ins** → **My Add-ins**
   - Vælg **Add a custom add-in** → **Add from file...**
   - Vælg `manifest.xml` fra dette projekt
   - Klik **Install**

4. **Test det!**
   - Gå til Calendar i Outlook
   - Klik **New Meeting**
   - Skabelonen indsættes automatisk

**Se `docs/QUICKSTART.md` for detaljeret guide.**

## Projekt struktur

```
customeeting/
│
├── manifest.xml                 # Office Add-in manifest (aktiv)
├── index.html                   # Landing page for add-in
├── commands.html                # Commands page (event handler)
├── commands.js                  # Event-based activation logic
├── taskpane.html                # Task pane UI
├── taskpane.js                  # Task pane logic
├── staticwebapp.config.json     # Azure Static Web App konfiguration
├── server.js                    # Lokal development server
├── package.json                 # NPM dependencies
│
├── assets/                      # Ikoner (PNG format)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
│
├── deploy/                      # Klar til Azure deployment
│   ├── manifest.xml             # Manifest med Azure URLs
│   ├── index.html
│   ├── commands.html
│   ├── commands.js
│   ├── taskpane.html
│   ├── taskpane.js
│   ├── staticwebapp.config.json
│   └── assets/                  # Ikoner
│
├── docs/                        # Dokumentation
│   ├── QUICKSTART.md            # Quick start guide
│   ├── DEPLOYMENT.md            # Fuld deployment guide
│   ├── AZURE-DEPLOYMENT.md      # Azure specifik guide
│   ├── HURTIG-AZURE-GUIDE.md    # Hurtig Azure guide
│   ├── MANUAL-UPLOAD-GUIDE.md   # Manuel Azure upload
│   ├── BRUGER-GUIDE.md          # Guide til slutbrugere
│   ├── PINNING-GUIDE.md         # Guide til at pinne add-in
│   ├── INTUNE-POLICY-QUICK.md   # Intune deployment guide
│   ├── FIX-GUIDE.md             # Fejlfinding
│   └── FILER-TIL-AZURE.txt      # Liste over Azure filer
│
├── archive/                     # Gamle/test filer (ignorer)
│   ├── manifest-*.xml           # Test manifest versioner
│   ├── commands-FIXED.js
│   ├── create-icons.html
│   └── icon-base64.txt
│
├── OutlookVSTO/                 # VSTO COM Add-in (alternativ løsning)
│   ├── OutlookMeetingAddin.sln
│   ├── OutlookMeetingAddin.csproj
│   ├── ThisAddIn.cs
│   ├── MeetingTaskPane.cs
│   └── README.md                # VSTO dokumentation
│
└── .github/                     # GitHub Actions workflows
    └── workflows/
```

## Dokumentation

### For udviklere

- **docs/QUICKSTART.md** - Kom hurtigt i gang med lokal udvikling
- **docs/DEPLOYMENT.md** - Komplet deployment guide
- **docs/AZURE-DEPLOYMENT.md** - Azure Static Web Apps deployment
- **docs/HURTIG-AZURE-GUIDE.md** - Hurtig Azure guide
- **docs/MANUAL-UPLOAD-GUIDE.md** - Manuel upload til Azure
- **docs/FIX-GUIDE.md** - Fejlfinding og problemer

### For slutbrugere

- **docs/BRUGER-GUIDE.md** - Guide til slutbrugere (hvordan man bruger add-in)
- **docs/PINNING-GUIDE.md** - Sådan pinner du add-in i Outlook

### For IT administrators

- **docs/INTUNE-POLICY-QUICK.md** - Deploy via Intune til alle brugere
- **Enable-MeetingTemplateAddin.ps1** - PowerShell script til centraliseret deployment

## Deployment

### Lokal test (Development)

For at teste lokalt:

```bash
npm install
node server.js
```

Tilføj add-in manuelt i Outlook via **File → Get Add-ins → My Add-ins → Add from file**.

### Azure Static Web Apps (Production)

1. **Upload til Azure**
   - Kopier filer fra `/deploy` mappen til Azure Static Web App
   - Se `docs/AZURE-DEPLOYMENT.md` for fuld guide

2. **Opdater manifest URLs**
   - Manifest i `/deploy` er allerede konfigureret til Azure
   - Skift URL hvis du bruger en anden hosting provider

3. **Deploy til brugere**
   - Via Microsoft 365 Admin Center: **Settings → Integrated apps → Upload custom app**
   - Via Intune: Se `docs/INTUNE-POLICY-QUICK.md`

### Enterprise deployment

For at deploye til 400+ brugere:

1. **Host add-in på Azure** (se docs/AZURE-DEPLOYMENT.md)
2. **Centraliseret deployment via:**
   - **Microsoft 365 Admin Center** (anbefalet)
   - **Intune Policy** (for device management)
   - **PowerShell script** (Enable-MeetingTemplateAddin.ps1)

Se detaljeret guide i `/OutlookVSTO/ENTERPRISE-DEPLOYMENT.md`

## Tilpas skabelonen

Skabelonen kan tilpasses i `commands.js`:

```javascript
const MEETING_TEMPLATE = `
<p><strong>🎯 Formål med mødet</strong></p>
<p>Hvorfor mødes vi?</p>
<br>

<p><strong>📋 Dagsorden/emner</strong></p>
<ul>
  <li>Emne 1</li>
  <li>Emne 2</li>
  <li>Emne 3</li>
</ul>
<br>

<p><strong>👥 Roller og ansvar</strong></p>
<ul>
  <li>Ordstyrer: </li>
  <li>Referent: </li>
  <li>Tidsansvarlig: </li>
</ul>
<br>

<p><strong>✅ Beslutninger og næste skridt</strong></p>
<ul>
  <li>Beslutning 1: </li>
  <li>Handling 1: </li>
</ul>
<br>

<p><strong>💬 Evt.</strong></p>
<p></p>
`;
```

**Efter ændringer:**
1. Gem filen
2. Push til GitHub (hvis du bruger GitHub Actions deployment)
3. Eller upload manuelt til Azure

## Tekniske detaljer

### Office.js API

Add-in bruger følgende Office.js APIs:
- **LaunchEvent API** - Event-based activation (OnNewAppointmentOrganizer)
- **Mailbox API** - Læs/skriv til møde body
- **Body API** - HTML formatering af body content

### Browser kompatibilitet

- Edge (Chromium)
- Chrome
- Safari (Mac)
- IE11 (legacy support via polyfills)

### Office versioner

- Microsoft 365 (Current Channel)
- Office 2021
- Office 2019
- Office 2016 (med opdateringer)

## Fejlfinding

### Add-in vises ikke i Outlook

1. Check at serveren kører (for lokal test)
2. Tjek at manifest.xml er korrekt installeret
3. Genstart Outlook
4. Check Event Viewer for fejl

### Skabelonen indsættes ikke automatisk

1. Åbn Developer Tools (F12) i mødevinduet
2. Check Console for JavaScript fejl
3. Verificer at LaunchEvent er registreret korrekt

### Får 404 fejl fra Azure

1. Check at alle filer er uploaded til Azure
2. Verificer staticwebapp.config.json indstillinger
3. Check CORS settings i Azure

Se `docs/FIX-GUIDE.md` for flere fejlfinding tips.

## Sikkerhed

- Add-in bruger HTTPS for produktion
- Ingen følsomme data gemmes
- Kører i sandboxed environment
- Følger Microsoft security best practices

## Support

For spørgsmål eller problemer:
1. Check dokumentation i `/docs` mappen
2. Se fejlfinding guide: `docs/FIX-GUIDE.md`
3. Opret et issue i GitHub repository

## Licens

Dette projekt er udviklet til intern brug i organisationen.

## Bidrag

For at bidrage til projektet:
1. Fork repository
2. Opret feature branch
3. Commit dine ændringer
4. Push til branch
5. Opret Pull Request

## Version historik

- **v1.1.0** - Event-based activation med LaunchEvent API
- **v1.0.0** - Initial release med manuel activation

## Kontakt

For support kontakt IT-afdelingen
