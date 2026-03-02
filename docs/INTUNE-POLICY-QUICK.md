# 🚀 Intune Policy - Auto-Enable Meeting Template Add-in

Quick guide til at lave Intune policy der auto-enabler add-in'et.

## 📋 Trin-for-trin:

### 1. Gå til Intune Admin Center
```
https://endpoint.microsoft.com
```

### 2. Opret Configuration Profile

- Klik **Devices** (venstre menu)
- Klik **Configuration profiles**
- Klik **Create profile**

### 3. Vælg Platform og Type

- **Platform:** Windows 10 and later
- **Profile type:** Settings catalog
- Klik **Create**

### 4. Basis information

- **Name:** `Auto-Enable Meeting Template Add-in`
- **Description:** `Automatically enables Meeting Template add-in for all Outlook users`
- Klik **Next**

### 5. Configuration Settings

1. Klik **Add settings**
2. I søgefeltet, søg: **"Office"**
3. Expand: **Microsoft Office Outlook**
4. Find og check: **"List of managed add-ins"**
5. Klik **Close**

### 6. Konfigurer Add-in

Nu skulle du se "List of managed add-ins" i settings listen:

1. Klik på dropdown ved "List of managed add-ins"
2. Klik **Add**
3. Indtast:

**Add-in ID:**
```
d6ff5248-95b4-4400-8f5d-4ce60c4ed0be
```

**Value (vælg EN af disse):**

**Option A - Always enabled (anbefalet):**
```
3
```

**Option B - Optional but shown:**
```
2
```

**Option C - Required (kan ikke fjernes):**
```
1
```

**Manifest URL (hvis der er et felt til det):**
```
https://nice-beach-04e97be03.1.azurestaticapps.net/manifest.xml
```

4. Klik **OK**

### 7. Assignments

- Klik **Next**
- Under "Assign to": Vælg **All users** eller **All devices**
- ELLER: Vælg specific groups (fx "KOMBIT Users")
- Klik **Next**

### 8. Applicability Rules (Optional)

- Klik **Next** (lad være blank medmindre du har specifikke krav)

### 9. Review + Create

- Gennemgå settings
- Klik **Create**

---

## ⏱️ Propagation Time

- **Intune sync:** 8 timer (standard)
- **Force sync:** Gå til device → **Sync** knap
- **Outlook restart:** Påkrævet efter policy applied

---

## ✅ Verificer Policy Applied

### På bruger maskine:

**Check 1: Intune sync status**
```
Settings → Accounts → Access work or school → Info → Sync
```

**Check 2: Registry (efter sync)**
```powershell
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\d6ff5248-95b4-4400-8f5d-4ce60c4ed0be" -ErrorAction SilentlyContinue
```

Skulle vise add-in settings hvis policy er applied.

**Check 3: Outlook**
- Genstart Outlook
- Åbn Calendar → New Meeting
- "Meeting Template" knap skulle være synlig i ribbon UDEN at klikke "Get Add-ins"

---

## 🔧 Alternative: Settings Catalog vs Administrative Templates

Hvis du IKKE kan finde "List of managed add-ins" i Settings Catalog:

### Brug Administrative Templates i stedet:

**Trin 3 (alternativ):**
- **Profile type:** Templates
- **Template name:** Administrative Templates
- Klik **Create**

**Trin 5 (alternativ):**
- Computer Configuration → **Microsoft Outlook 2016** → **Miscellaneous**
- Find: **"List of managed add-ins"**
- Enable policy
- Add entry:
  ```
  Value name: d6ff5248-95b4-4400-8f5d-4ce60c4ed0be
  Value data: 3
  ```

---

## 📱 Value Forklaring

**LoadBehavior Values:**
- `0` = Not loaded
- `1` = Loaded (required, cannot be disabled)
- `2` = Load on demand (user can disable)
- `3` = Load at startup (always enabled) ⭐ **ANBEFALET**
- `9` = Load on demand and connected
- `16` = Loaded once

**For auto-enable uden user interaction: Brug `3`**

---

## 🆘 Troubleshooting

### Policy ikke applied efter 24 timer:

**Check 1: Policy status**
```
Endpoint Manager → Devices → Monitor → Configuration → Find din policy → Status
```

Skal vise "Succeeded" for devices.

**Check 2: Conflict resolution**
Hvis en anden policy sætter samme setting, kan der være konflikt.

**Check 3: Office version**
Policy virker kun med:
- Office 2016 or later
- Microsoft 365 Apps

**Check 4: User vs Device policy**
Outlook add-ins er typisk **User** configuration, ikke Device.
Sørg for at assigne til **Users**, ikke kun **Devices**.

---

## 💡 Quick Test

For at teste på EN maskine før rollout:

**PowerShell som admin:**
```powershell
# Simuler Intune policy manuelt
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\d6ff5248-95b4-4400-8f5d-4ce60c4ed0be"
New-Item -Path $regPath -Force
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "https://nice-beach-04e97be03.1.azurestaticapps.net/manifest.xml"
Set-ItemProperty -Path $regPath -Name "Description" -Value "Meeting Template"
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Meeting Template"

Write-Host "Add-in registry keys created. Restart Outlook to test."
```

Genstart Outlook og test om knappen vises automatisk.

---

**Efter Intune policy er applied, skulle add-in'et automatisk være synligt! ✅**
