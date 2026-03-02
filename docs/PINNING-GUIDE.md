# 📌 Auto-Pin Outlook Add-in Guide

Sådan sikrer du at Meeting Template add-in automatisk er aktivt for alle brugere.

## 🎯 Metode 1: Via Microsoft 365 Admin Center (Nemmest)

### Trin 1: Find dit deployed add-in

1. Gå til **https://admin.microsoft.com/#/Settings/IntegratedApps**
2. Find **"Meeting Template"** i listen
3. Klik på add-in'et

### Trin 2: Rediger Deployment Settings

1. Klik på **"Edit"** eller **"Manage"**
2. Find **"User access"** eller **"Availability"** sektionen
3. Ændre settings:
   - **Availability:** Sæt til **"Optional, always enabled"**
   - ELLER: **"Required"** (tvinger det på for alle)

4. Klik **"Save"**

### Trin 3: Vent på propagation

- Ændringer tager **2-6 timer** at propagere
- Brugerne kan genstart Outlook for hurtigere effekt

---

## 🎯 Metode 2: Via Intune Configuration Policy

### Trin 1: Opret Configuration Profile

1. Gå til **https://endpoint.microsoft.com**
2. **Devices** → **Configuration profiles** → **Create profile**
3. Platform: **Windows 10 and later**
4. Profile type: **Templates** → **Administrative Templates**
5. Klik **Create**

### Trin 2: Konfigurer Add-in Policy

1. **Configuration settings** → **Add**
2. Find: **Microsoft Outlook 2016**
3. Søg efter: **"List of managed add-ins"**
4. **Enable** policy
5. Klik **Show...** ved "List of managed add-ins"

### Trin 3: Tilføj dit add-in

Tilføj en række med:

**Value name:**
```
d6ff5248-95b4-4400-8f5d-4ce60c4ed0be
```

**Value:**
```
3;https://icy-hill-0716f1a03.1.azurestaticapps.net/manifest.xml
```

**Forklaring af value:**
- `3` = Always enabled (kan ikke deaktiveres af bruger)
- `2` = Optional (bruger kan deaktivere)
- `1` = Mandatory (altid aktivt, kan ikke fjernes)

### Trin 4: Assign policy

1. **Assignments** tab
2. Assign til **All Users** eller specifikke grupper
3. Klik **Create**

---

## 🎯 Metode 3: Via Group Policy (On-Premises AD)

### Hvis du bruger On-Premises Active Directory:

1. Download Office Administrative Template files:
   https://www.microsoft.com/en-us/download/details.aspx?id=49030

2. Åbn **Group Policy Management**

3. Edit policy → **User Configuration** → **Administrative Templates** → **Microsoft Outlook 2016** → **Miscellaneous**

4. Find: **"List of managed add-ins"**

5. **Enable** og tilføj:
   ```
   d6ff5248-95b4-4400-8f5d-4ce60c4ed0be=3;https://icy-hill-0716f1a03.1.azurestaticapps.net/manifest.xml
   ```

6. **gpupdate /force** på klient maskiner

---

## 🎯 Metode 4: Registry Key (Direkte på maskiner)

### For hurtig test på enkelt maskine:

**Registry path:**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\d6ff5248-95b4-4400-8f5d-4ce60c4ed0be
```

**Værdier:**
```
"Description" = "Meeting Template"
"FriendlyName" = "Meeting Template"
"LoadBehavior" = dword:00000003
"Manifest" = "https://icy-hill-0716f1a03.1.azurestaticapps.net/manifest.xml"
```

**ELLER via PowerShell:**

```powershell
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\d6ff5248-95b4-4400-8f5d-4ce60c4ed0be"

New-Item -Path $regPath -Force
Set-ItemProperty -Path $regPath -Name "Description" -Value "Meeting Template"
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Meeting Template"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "https://icy-hill-0716f1a03.1.azurestaticapps.net/manifest.xml"
```

**LoadBehavior værdier:**
- `0` = Unloaded
- `1` = Connected (loaded on demand)
- `2` = Load on startup, but disabled
- `3` = Load on startup (ANBEFALET)

---

## 📋 Verificer at det virker

**Efter konfiguration, test:**

1. **Genstart Outlook** helt (luk alle vinduer)
2. **Åbn Calendar**
3. **New Meeting**
4. **Check ribbon** - "Meeting Template" knappen skulle automatisk være synlig
5. Brugeren skal IKKE klikke på "Get Add-ins" eller aktivere noget

---

## ⚙️ Troubleshooting

### Add-in vises stadig ikke automatisk:

**Check 1: Er add-in'et deployed korrekt?**
```
Admin Center → Integrated Apps → Meeting Template → Status = "Deployed"
```

**Check 2: Har brugeren modtaget deployment?**
- Vent 24 timer efter deployment
- Genstart Outlook

**Check 3: Check Office version**
- Outlook Desktop version 2016 eller nyere påkrævet
- Microsoft 365 Apps virker bedst

**Check 4: Check Outlook add-in trust center**
- File → Options → Trust Center → Trust Center Settings
- Add-ins → Require Application Extensions to be signed by Trusted Publisher (skal være UNchecked for custom add-ins)

---

## 🎯 Anbefalet Approach

**For KOMBIT deployment:**

1. **Brug Metode 1** (Admin Center) først - nemmest
2. Hvis det ikke virker, **brug Metode 2** (Intune policy) - mest kontrol
3. For test maskiner, **brug Metode 4** (Registry) - hurtigst at teste

---

## 📞 Support til brugere

**Hvis en bruger STADIG skal aktivere add-in'et manuelt:**

Det betyder policy'en ikke er applied korrekt. Check:
1. Er brugeren i den gruppe der har fået assigned add-in'et?
2. Har policy'en synced (Intune: Sync device i Endpoint Manager)
3. Har brugeren genstartet Outlook siden deployment?

---

**Efter konfiguration skulle add-in'et AUTOMATISK være synligt i ribbon!** ✅
