# Enterprise Deployment Guide - Central Distribution

Dette dokument beskriver hvordan du distribuerer Outlook Meeting Template Add-in til mange maskiner uden brugerinteraktion.

## 🎯 Deployment Muligheder

### ⭐ ANBEFALET: Web Add-in via Microsoft 365 Admin Center

**Fordele:**
- ✅ Centraliseret deployment fra admin center
- ✅ Ingen installation på klient maskiner
- ✅ Automatiske opdateringer
- ✅ Ingen brugerinteraktion
- ✅ Virker på Windows, Mac, Web, Mobile

**Ulemper:**
- ❌ Kan IKKE auto-show taskpane (Microsoft begrænsning)
- ❌ Bruger skal klikke på add-in i ribbon

**Sådan deployes:**
1. Upload dit web add-in manifest til Microsoft 365 Admin Center
2. Gå til **Settings** → **Integrated apps** → **Add-ins**
3. Upload `manifest.xml`
4. Vælg "Deploy to entire organization" eller specifikke grupper
5. Add-in rulles automatisk ud til alle brugere

**Dette er din nuværende løsning!** Men den kan ikke auto-show.

---

### 🔧 Alternativ 1: VSTO via Group Policy (GPO)

For Windows domæner med Active Directory.

**Fordele:**
- ✅ Automatisk deployment
- ✅ Kan auto-show taskpane
- ✅ Central administration

**Kræver:**
- Windows Server med Active Directory
- VSTO Runtime på alle maskiner
- Netværks share til deployment

**Deployment Steps:**
1. Build VSTO add-in som MSI installer
2. Placer MSI på network share
3. Opret GPO der installerer MSI
4. Link GPO til relevante OUs

Se [GPO-DEPLOYMENT.md](GPO-DEPLOYMENT.md) for detaljer.

---

### 📱 Alternativ 2: Intune Deployment

For moderne cloud-managed maskiner.

**Fordele:**
- ✅ Cloud-baseret deployment
- ✅ Ingen netværks share nødvendigt
- ✅ Modern management

**Deployment Steps:**
1. Pakke add-in som `.intunewin` fil
2. Upload til Intune
3. Assign til user/device groups
4. Deployment sker automatisk

Se [INTUNE-DEPLOYMENT.md](INTUNE-DEPLOYMENT.md) for detaljer.

---

### 🚀 Alternativ 3: PowerShell Silent Installation

For ad-hoc deployment eller kombineret med andre metoder.

**Fordele:**
- ✅ Kan køres remote via PSRemoting
- ✅ Kan integreres i eksisterende deployment pipelines
- ✅ Fuld kontrol

**Brug:**
```powershell
# Kør på remote maskine
Invoke-Command -ComputerName PC001,PC002,PC003 -FilePath .\Deploy-AddIn.ps1
```

Se [POWERSHELL-DEPLOYMENT.md](POWERSHELL-DEPLOYMENT.md) for detaljer.

---

## 🎓 Sammenligning

| Metode | Auto-Show | User Interaction | Platform Support | Kompleksitet |
|--------|-----------|------------------|------------------|--------------|
| Web Add-in (M365 Admin) | ❌ Nej | ❌ Ingen | Windows/Mac/Web/Mobile | ⭐ Lav |
| VSTO via GPO | ✅ Ja | ❌ Ingen | Kun Windows | ⭐⭐⭐ Høj |
| VSTO via Intune | ✅ Ja | ❌ Ingen | Kun Windows | ⭐⭐⭐ Høj |
| PowerShell Script | ✅ Ja | ❌ Ingen | Kun Windows | ⭐⭐ Medium |
| VBA Macro | ✅ Ja | ✅ Ja (manual) | Kun Windows | ⭐ Lav |

## 💡 Min Anbefaling For Dig

### Hvis I bruger Microsoft 365 (cloud):
**Brug Web Add-in + "Quick Tip" til brugerne**
- Deploy via M365 Admin Center
- Send en kort email: "Klik på 'Meeting Template' knappen når du opretter møder"
- Simpelt og vedligeholdelsesvenligt

### Hvis I SKAL have auto-show:

**Option A: VSTO via Intune** (moderne organisationer)
- Bedst for cloud-first organisationer
- Modern device management
- Se INTUNE-DEPLOYMENT.md

**Option B: VSTO via GPO** (traditionelle domæner)
- Bedst for on-premise AD domæner
- Proven technology
- Se GPO-DEPLOYMENT.md

### Hvis I vil teste hurtigt:
**PowerShell Script**
- Deploy til en pilot gruppe først
- Se POWERSHELL-DEPLOYMENT.md

## 📋 Næste Skridt

Vælg en deployment metode baseret på din infrastruktur:

1. **Har I Microsoft 365 tenant?** → Web Add-in via Admin Center
2. **Har I Intune?** → VSTO via Intune
3. **Har I Active Directory med GPO?** → VSTO via GPO
4. **Vil I teste først?** → PowerShell Script

Fortæl mig hvilken infrastruktur I har, så kan jeg lave en komplet deployment pakke!
