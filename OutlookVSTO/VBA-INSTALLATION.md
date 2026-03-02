# 🚀 VBA Løsning - Installation Guide (2 minutter!)

Denne løsning giver dig **automatisk møde skabelon indsættelse** UDEN at skulle bygge noget!

## ✅ Fordele

- ✅ Ingen build proces
- ✅ Ingen installation
- ✅ Virker med det samme
- ✅ Automatisk indsættelse
- ✅ Let at dele med kolleger

## 📋 Step-by-Step Installation

### Trin 1: Aktivér Developer Tab i Outlook

1. Åbn **Outlook**
2. Gå til **File** → **Options**
3. Vælg **Customize Ribbon**
4. I højre side, find og **sæt hak ved "Developer"**
5. Klik **OK**

Du skulle nu kunne se en "Developer" tab i Outlook ribbon.

### Trin 2: Aktivér Macros

1. I Outlook, gå til **File** → **Options** → **Trust Center**
2. Klik **Trust Center Settings**
3. Vælg **Macro Security**
4. Vælg **"Notifications for all macros"** (eller "Enable all macros" hvis du stoler på dig selv)
5. Klik **OK** to gange

### Trin 3: Åbn VBA Editor

1. I Outlook, tryk **Alt+F11** (eller gå til Developer tab → Visual Basic)
2. VBA Editor åbner

### Trin 4: Indsæt Koden

1. I **Project Explorer** (venstre side), find:
   ```
   Project1
   └── Microsoft Outlook Objects
       └── ThisOutlookSession
   ```
2. **Dobbeltklik** på "ThisOutlookSession"
3. En editor åbner på højre side
4. **Åbn filen** `MeetingTemplate.vba` i Notepad
5. **Kopier HELE indholdet** (Ctrl+A, Ctrl+C)
6. **Indsæt i VBA Editor** (Ctrl+V)
7. **Gem** (Ctrl+S eller File → Save)
8. **Luk VBA Editor**

### Trin 5: Genstart Outlook

1. **Luk Outlook fuldstændigt**
2. **Start Outlook igen**
3. Hvis du får en sikkerhedsadvarsel om macros, klik **"Enable Macros"**

## 🎉 Test Det!

1. Gå til **Calendar** i Outlook
2. Klik **New Meeting** eller **New Appointment**
3. **BOOM!** 💥 Skabelonen indsættes automatisk!

## 📸 Sådan ser det ud

```
═══════════════════════════════════════════════════
           MØDE SKABELON
═══════════════════════════════════════════════════

📋 FORMÅL MED MØDET
───────────────────────────────────────────────────
Kort beskrivelse af, hvorfor vi mødes, og hvad vi skal opnå.


📝 DAGSORDEN/EMNER
───────────────────────────────────────────────────
• Punkt 1:
• Punkt 2:
• Punkt 3:


👥 ROLLER OG ANSVAR
───────────────────────────────────────────────────
Mødeleder:
Referent:


✅ BESLUTNINGER OG NÆSTE SKRIDT
───────────────────────────────────────────────────
• Beslutning 1:
• Action item 1: [Ansvarlig] [Deadline]
• Action item 2: [Ansvarlig] [Deadline]


💡 EVT.
───────────────────────────────────────────────────
Tid til spørgsmål eller andre punkter.


═══════════════════════════════════════════════════
```

## 🎨 Tilpas Skabelonen

Hvis du vil ændre skabelonen:

1. Tryk **Alt+F11** i Outlook
2. Find `BuildMeetingTemplate` funktionen
3. Rediger teksten efter behov
4. Gem (Ctrl+S)
5. Genstart Outlook

## 📤 Del Med Kolleger

### Metode 1: Export VBA Project

1. I Outlook, tryk Alt+F11
2. Højreklik på "Project1" → **Export File**
3. Gem som `MeetingTemplate.bas`
4. Send til kolleger
5. De kan **importere** ved at højreklikke → Import File

### Metode 2: Manual Copy-Paste

Send dem bare `MeetingTemplate.vba` filen og denne installations guide!

## 🔧 Fejlfinding

### "Macro er deaktiveret"

**Løsning:**
1. File → Options → Trust Center → Trust Center Settings
2. Macro Security → Vælg "Notifications for all macros"
3. Genstart Outlook

### Skabelonen indsættes ikke

**Løsning:**
1. Check at koden er i "ThisOutlookSession" (ikke et nyt module)
2. Genstart Outlook
3. Check at macros er aktiveret

### "Application_Startup" køres ikke

**Løsning:**
- Denne event kører kun ved Outlook opstart
- Luk Outlook fuldstændigt og start igen
- Check Task Manager at OUTLOOK.EXE ikke kører

## 🆚 VBA vs VSTO

| Feature | VBA (Denne løsning) | VSTO |
|---------|---------------------|------|
| Installation | 2 minutter | Kræver Visual Studio |
| Build proces | Ingen | Kompleks |
| Deployment | Copy-paste kode | Installer package |
| Automatisk indsættelse | ✅ Ja | ✅ Ja |
| Custom UI | Begrænset | Fuld WinForms |
| Sikkerhed | Kræver macro tillid | Digital signatur mulig |

## 💡 Pro Tips

1. **Kun indsæt i nye møder:** Koden checker om body er tom først
2. **Virker med både Appointment og Meeting Requests**
3. **Kan tilpasses** - rediger `BuildMeetingTemplate()` funktionen
4. **Export din customization** og gem som backup

## 🎓 Næste Skridt

- Tilpas skabelonen til dit team
- Exporter og distribuer til kolleger
- Tilføj flere templates for forskellige møde typer
- Opret custom ribbon buttons til forskellige skabeloner

---

**Spørgsmål?** Lad mig vide hvis noget ikke virker! 🚀
