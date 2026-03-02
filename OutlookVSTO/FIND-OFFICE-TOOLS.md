# 📍 Sådan Finder Du Office Tools Installation

## 🎯 Hurtig Metode (2 klik)

### Via Windows Start Menu:

1. **Tryk Windows-tast**
2. **Skriv:** `Visual Studio Installer`
3. **Klik på:** Visual Studio Installer appen

**ELLER**

### Via File Explorer:

1. **Tryk Windows + R**
2. **Skriv:** `"C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe"`
3. **Tryk Enter**

---

## 📋 Når Visual Studio Installer Åbner

Du vil se noget lignende dette:

```
┌─────────────────────────────────────────┐
│  Visual Studio Installer                │
├─────────────────────────────────────────┤
│                                         │
│  📦 Visual Studio Community 2025        │
│     Version 18.x.xxxxx                  │
│                                         │
│     [Modify]  [More ▼]                  │
│                                         │
└─────────────────────────────────────────┘
```

### Trin for Trin:

**1. Klik på "Modify" knappen**

**2. Vælg "Workloads" tab (skulle være default)**

**3. Scroll ned og find:**
```
☐ Office/SharePoint development
  Build Office and SharePoint add-ins and solutions
```

**4. Sæt HAK i boksen**

**5. Klik "Modify" nederst til højre**

**6. Vent 10-15 minutter**

Det er det! 🎉

---

## ⚠️ ALTERNATIV: Du behøver faktisk IKKE Office Tools!

**Hvis du vil deploye VBA i stedet:**

✅ **INGEN Visual Studio nødvendigt**
✅ **INGEN build proces**
✅ **INGEN Office Tools**

**Bare:**
1. Åbn Outlook
2. Eksporter VBA template
3. Deploy via Intune

Se: `VBA-ENTERPRISE-DEPLOYMENT.md`

---

## 🎯 Så Hvad Nu?

Du har nu læst om sikkerhed. **Hvad er din beslutning?**

### Option A: VBA (Min anbefaling)
**Sikkerhed:** ✅ Acceptabel med proper deployment
**Tid:** ⏱️ Klar i dag
**Næste step:** Læs `VBA-ENTERPRISE-DEPLOYMENT.md`

### Option B: VSTO
**Sikkerhed:** ✅ Marginal bedre (hvis digitalt signeret)
**Tid:** ⏱️ Klar i morgen
**Næste step:** Installer Office Tools (guide ovenfor)

---

**Fortæl mig:** "VBA" eller "VSTO"? 🚀
