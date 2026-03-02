# Manual Installation af Office Development Tools

Visual Studio Installer kunne ikke køres automatisk. Her er hvordan du gør det manuelt (5 minutter):

## 📋 Step 1: Åbn Visual Studio Installer

1. **Tryk Windows-tast**
2. **Søg efter:** "Visual Studio Installer"
3. **Klik på den** (eller kør: `C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe`)

## 🔧 Step 2: Modificer Visual Studio 18

1. I Visual Studio Installer, find **Visual Studio Community 2025 (version 18)**
2. Klik **"Modify"** knappen

## ✅ Step 3: Vælg Office Workload

1. Under **"Workloads"** tab, scroll ned til:
   - ✅ **"Office/SharePoint development"**

2. Sæt hak ved denne workload

3. (Valgfrit) Under "Individual components", check at disse er valgt:
   - Office Developer Tools
   - VSTO (Visual Studio Tools for Office)
   - Office 365 Developer Tools

## 💾 Step 4: Installer

1. Klik **"Modify"** knappen nederst til højre
2. Vent på download og installation (**5-15 minutter**)
3. Du får en notifikation når det er færdigt

## ✅ Step 5: Verificer Installation

Når installationen er færdig:

```powershell
# Check at Office tools er installeret
Test-Path "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\PublicAssemblies\Microsoft.Office.Tools.*.dll"
```

Skulle returnere `True`.

## 🚀 Step 6: Build VSTO Projekt

Nu kan du bygge projektet:

```powershell
cd "C:\Users\dbl\OneDrive - netIP\Skrivebord\customeeting\OutlookVSTO"

& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    OutlookMeetingAddin.sln `
    /p:Configuration=Release
```

---

## ⚡ Hurtigere Alternativ: Kommandolinje Installation

Hvis du vil køre det via kommandolinje (kræver admin):

```cmd
"C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe" ^
    modify ^
    --installPath "C:\Program Files\Microsoft Visual Studio\18\Community" ^
    --add Microsoft.VisualStudio.Workload.Office ^
    --passive
```

Dette åbner et vindue med progress, men kræver ingen interaktion.

---

**EFTER INSTALLATION:** Gå tilbage til main deployment guide: [DEPLOYMENT-PLAN-400-USERS.md](DEPLOYMENT-PLAN-400-USERS.md)
