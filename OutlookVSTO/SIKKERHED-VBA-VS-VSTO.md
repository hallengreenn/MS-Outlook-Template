# 🔒 Sikkerhedsanalyse: VBA vs VSTO

**For 400 brugere - Enterprise Security Perspektiv**

---

## 🎯 Executive Summary

**Kort svar:** VSTO er mere sikkert end VBA, MEN begge kan gøres sikre med korrekt deployment.

**Anbefaling for 400 brugere:** VBA via Intune med korrekt Group Policy konfiguration er **acceptabelt sikkert** for intern brug.

---

## 🔒 Sikkerhedssammenligning

### VBA Makro Sikkerhed

#### ⚠️ Risici:
1. **Macro Security Prompt**
   - Brugere kan enable macros fra ondsindede kilder
   - Social engineering risiko
   - Brugere kan ikke skelne mellem god/ond macro

2. **Kode Er Synlig**
   - VBA kode kan læses og modificeres af brugere
   - Ikke compilet/obfuscated
   - Kan kopieres til andre VBA projekter

3. **Ingen Code Signing**
   - Standard VBA macros kan ikke digitalt signeres nemt
   - Svært at verificere oprindelse

4. **Privilegie Eskalering**
   - Macros kører i brugerens security context
   - Har samme rettigheder som brugeren
   - Kan potentielt tilgå brugerens filer

#### ✅ Mitigering:
1. **Controlled Deployment via Intune**
   - Centralt deployed = kontrolleret kilde
   - Ikke bruger-downloaded
   - IT kontrollerer hvad der deployes

2. **Group Policy Macro Settings**
   ```
   Disable all macros except digitally signed macros
   - Trust Center → Trusted Locations
   - Tilføj %APPDATA%\Microsoft\Outlook som trusted
   ```

3. **Read-Only Deployment**
   - Deploy OTM fil som read-only
   - Bruger kan ikke modificere

4. **Audit Logging**
   - Log macro execution via Group Policy
   - Monitor for uautoriseret macro brug

5. **Limited Scope**
   - Vores VBA kun indsætter text
   - Ingen fil access, ingen network calls
   - Minimal attack surface

---

### VSTO Add-in Sikkerhed

#### ✅ Fordele:
1. **Compiled Code**
   - .NET assembly er compiled
   - Sværere at reverse engineer end VBA
   - Kan obfusceres

2. **Digital Signing**
   - Kan signeres med code signing certificate
   - Windows verificerer signatur ved installation
   - Trusted publisher mechanism

3. **ClickOnce Security**
   - Sandbox execution model
   - Permission model (kan begrænses)
   - Certificate-based trust

4. **No Macro Prompt**
   - Ingen scary "Enable Macros" besked
   - Bedre user experience
   - Mindre social engineering risiko

5. **Deployment Control**
   - Via Intune/GPO = centralt kontrolleret
   - Version management
   - Kan force updates

#### ⚠️ Risici:
1. **Større Attack Surface**
   - Fuld .NET framework tilgængelig
   - Kan lave netværks calls, fil access, etc.
   - Mere kompleks kode = flere potentielle bugs

2. **Dependencies**
   - VSTO Runtime dependency
   - Flere komponenter = større attack surface
   - Version compatibility issues

3. **Privilege Issues**
   - Kører også i bruger context
   - Samme fil/netværk access som brugeren
   - Kan potentielt misbruges

---

## 🎯 Specifik Risikovurdering for Jeres Add-in

### Vores VBA Kode:
```vba
' Hele koden gør dette:
1. Lytter til nye appointment events
2. Checker om body er tom
3. Indsætter text template
```

**Security Assessment:**
- ✅ **Ingen fil access** - Læser/skriver ikke filer
- ✅ **Ingen network calls** - Ingen internet forbindelser
- ✅ **Ingen registry changes** - Modificerer ikke registry
- ✅ **Ingen external processes** - Starter ikke andre programmer
- ✅ **Read-only data** - Indsætter kun præ-defined text
- ✅ **Minimal privileges** - Bruger kun Outlook API

**Risiko Niveau:** 🟢 **LAV**

**Hvorfor lav risiko:**
- Koden er transparent og simpel
- Ingen ekstern kommunikation
- Ingen bruger data indsamling
- Ingen admin privileges nødvendigt

---

### Vores VSTO Kode:
```csharp
// Samme funktionalitet som VBA
1. Lytter til nye appointment events (via COM interop)
2. Viser task pane
3. Indsætter text template når bruger klikker
```

**Security Assessment:**
- ✅ **Samme sikkerhedsprofil som VBA**
- ✅ **Compiled = mindre transparent**
- ✅ **Kan digitalt signeres**

**Risiko Niveau:** 🟢 **LAV** (samme som VBA)

---

## 🛡️ Enterprise Security Best Practices

### For VBA Deployment:

#### 1. Group Policy Configuration
```
Computer Configuration → Administrative Templates →
Microsoft Outlook 2016 → Security → Trust Center

Setting: VBA Macro Notification Settings
Value: Disable all except digitally signed macros

Setting: Trusted Locations
Add: %APPDATA%\Microsoft\Outlook
```

#### 2. Deploy via Intune (Controlled Source)
```powershell
# Intune deployment betyder:
- IT kontrollerer kilde
- Ikke bruger-downloaded
- Centralt management
- Kan revoke/update
```

#### 3. File Permissions
```powershell
# Deploy OTM som read-only
$vbaPath = "$env:APPDATA\Microsoft\Outlook\VbaProject.OTM"
Set-ItemProperty -Path $vbaPath -Name IsReadOnly -Value $true
```

#### 4. Audit Logging
```
Enable Audit Logging for Macros:
HKCU\Software\Microsoft\Office\16.0\Outlook\Security
Value: EnableLog = 1
```

---

### For VSTO Deployment:

#### 1. Code Signing Certificate
```powershell
# Sign assembly med company certificate
signtool sign /sha1 [cert-thumbprint] /t http://timestamp.digicert.com OutlookMeetingAddin.dll
```

#### 2. Trusted Publisher
```
Add code signing certificate to Trusted Publishers
via Group Policy:
Computer Configuration → Windows Settings →
Security Settings → Public Key Policies → Trusted Publishers
```

#### 3. ClickOnce Security
```xml
<!-- I manifest.xml -->
<trustInfo>
  <security>
    <applicationRequestMinimum>
      <PermissionSet class="System.Security.PermissionSet"
                     version="1"
                     ID="Custom"
                     SameSite="site">
        <!-- Minimal permissions -->
      </PermissionSet>
    </applicationRequestMinimum>
  </security>
</trustInfo>
```

---

## 🎓 Sammenligning: Sikkerhed i Kontekst

### VBA med Proper Deployment:
```
✅ Centralt deployed via Intune
✅ Trusted location (Group Policy)
✅ Read-only file
✅ Audit logging enabled
✅ Minimal code scope
= 🟢 ACCEPTABELT for intern brug
```

### VSTO med Proper Deployment:
```
✅ Centralt deployed via Intune
✅ Digitalt signeret
✅ Trusted publisher (Group Policy)
✅ ClickOnce security model
✅ Samme minimal code scope
= 🟢 BEDRE, men ikke nødvendigvis kritisk bedre
```

---

## 💼 Compliance Perspektiv

### Hvis I har compliance krav (ISO 27001, SOC 2, osv.):

**VBA:**
- ⚠️ Kan kræve ekstra dokumentation
- ⚠️ "Macros" kan være red flag i audit
- ✅ Mitigeres med: Controlled deployment + logging

**VSTO:**
- ✅ Bedre story for auditors
- ✅ "Managed add-in" lyder mere professionelt
- ✅ Digital signing = clear audit trail

**Hvis I IKKE har strenge compliance krav:**
- VBA er fint med proper deployment

---

## 🚨 Reelle Trusler

### Trusler VBA IKKE beskytter mod:
1. **Kompromitteret udvikler maskine**
   - Hvis min/din maskine er hacket under development
   - Gælder BÅDE VBA og VSTO

2. **Insider threat**
   - Ondsindet IT admin der deployer malicious code
   - Gælder BÅDE VBA og VSTO

3. **Supply chain attack**
   - Kompromitteret build environment
   - Mindre relevant for VBA (ingen build)
   - Gælder VSTO

### Trusler VBA beskytter mod (med proper deployment):
1. ✅ **Phishing macros** - Vores er centralt deployed
2. ✅ **User error** - Brugere downloader ikke selv
3. ✅ **Unauthorized modification** - Read-only + Intune control

---

## 📊 Risk Matrix

| Trussel | Sandsynlighed | Impact | VBA Risiko | VSTO Risiko |
|---------|---------------|--------|------------|-------------|
| User downloads malicious macro | Lav (centralt deployed) | Høj | 🟢 Lav | 🟢 Lav |
| Code modification by user | Medium | Medium | 🟡 Medium | 🟢 Lav |
| Privilege escalation | Lav (limited scope) | Medium | 🟢 Lav | 🟢 Lav |
| Data exfiltration | Meget lav (no network) | Høj | 🟢 Lav | 🟢 Lav |
| Audit compliance issues | Medium | Medium | 🟡 Medium | 🟢 Lav |
| Social engineering | Lav (controlled deployment) | Medium | 🟡 Medium | 🟢 Lav |

**Samlet Risiko:**
- **VBA:** 🟡 LOW-MEDIUM (acceptabelt med mitigations)
- **VSTO:** 🟢 LOW (bedre, men marginal forskel for vores use case)

---

## 🎯 Min Anbefaling for 400 Brugere

### Hvis I HAR strenge security/compliance krav:
→ **VSTO med digital signing**
- Bedre audit story
- Professionel appearance
- Clear chain of trust

### Hvis I IKKE har strenge compliance krav:
→ **VBA med proper Intune deployment**
- Hurtigere implementation
- Acceptabel sikkerhed med mitigations
- Nemmere at vedligeholde

### Security Checklist (Uanset hvilken):
- [ ] Deploy via Intune (controlled source)
- [ ] Enable audit logging
- [ ] Document deployment for audit trail
- [ ] Regular security reviews
- [ ] Version control of source code
- [ ] Limited scope (kun meeting template)
- [ ] No external network calls
- [ ] No sensitive data collection

---

## 🔐 Bonus: Hvordan Vi Gør VBA Mere Sikkert

### 1. Digital Signing af VBA (Avanceret)
```
1. Køb code signing certificate
2. I Outlook VBA Editor:
   Tools → Digital Signature
3. Sign VBA project
4. Deploy signed OTM
5. Group Policy: Trust only signed macros
```

### 2. Hash Verification
```powershell
# I Intune deployment script, verificer hash
$expectedHash = "ABC123..."
$actualHash = Get-FileHash VbaProject.OTM -Algorithm SHA256
if ($actualHash.Hash -ne $expectedHash) {
    Write-Error "Hash mismatch! Potential tampering!"
    exit 1
}
```

### 3. Monitoring
```powershell
# Log alle macro executions
# Intune kan indsamle disse logs
```

---

## 📞 Spørgsmål til at hjælpe jer med at beslutte

1. **Har I compliance krav (ISO, SOC 2, GDPR audits)?**
   - Ja → VSTO (bedre audit story)
   - Nej → VBA (hurtigere)

2. **Er jeres IT security team bekymrede for macros?**
   - Ja → VSTO
   - Nej → VBA

3. **Har I code signing certificate?**
   - Ja → VSTO (brug det!)
   - Nej → VBA (billigere)

4. **Er "Enable Macros" prompt en sikkerhedsbekymring?**
   - Ja → VSTO
   - Nej → VBA

---

## ✅ Bundlinje

**For jeres specifikke use case:**
- ✅ Intern brug (ikke customer-facing)
- ✅ Simpel funktionalitet (indsæt text)
- ✅ Ingen data collection
- ✅ Ingen external network calls
- ✅ Centralt deployed via Intune

**VBA er sikkerhedsmæssigt acceptabelt!**

Men hvis I har tid og vil have "gold standard", så VSTO med digital signing.

**Min anbefaling:** Start med VBA, upgrade til VSTO senere hvis security audit kræver det.

---

**Hvad er jeres security posture?** Fortæl mig hvis I har specifikke compliance krav, så kan jeg justere anbefalingen.
