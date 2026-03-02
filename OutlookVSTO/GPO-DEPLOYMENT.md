# Group Policy (GPO) Deployment Guide

Denne guide viser hvordan du deployer Outlook Meeting Template Add-in via Active Directory Group Policy til alle maskiner uden brugerinteraktion.

## 📋 Forudsætninger

- ✅ Windows Server med Active Directory
- ✅ Group Policy Management Console (GPMC)
- ✅ Netværks share tilgængelig for alle klient maskiner
- ✅ VSTO add-in bygget og klar
- ✅ Administrative rettigheder

## 🚀 Deployment Steps

### Trin 1: Forbered Netværks Share

1. **Opret et netværks share:**
   ```powershell
   # På file server
   New-Item -Path "C:\Shares\OutlookAddIn" -ItemType Directory
   New-SmbShare -Name "OutlookAddIn" -Path "C:\Shares\OutlookAddIn" -ReadAccess "Domain Computers"
   ```

2. **Kopier add-in filer til share:**
   ```
   \\server\OutlookAddIn\
   ├── OutlookMeetingAddin.dll
   ├── OutlookMeetingAddin.dll.manifest
   ├── OutlookMeetingAddin.vsto
   ├── Deploy-AddIn-Silent.ps1
   └── vstor_redist.exe (VSTO Runtime installer)
   ```

### Trin 2: Opret Deployment Script

Placer `Deploy-AddIn-Silent.ps1` på netværks share.

Test scriptet manuelt først:
```powershell
# Test på én maskine
\\server\OutlookAddIn\Deploy-AddIn-Silent.ps1 -NetworkShare "\\server\OutlookAddIn"
```

### Trin 3: Opret Group Policy Object

1. **Åbn Group Policy Management Console**
   - På domain controller: Start → Administrative Tools → Group Policy Management

2. **Opret ny GPO:**
   - Højreklik på din OU → Create a GPO in this domain, and Link it here
   - Navn: "Deploy Outlook Meeting Template Add-in"

3. **Rediger GPO:**
   - Højreklik på GPO → Edit

### Trin 4: Konfigurer Computer Startup Script

1. **Naviger til:**
   ```
   Computer Configuration
   └── Policies
       └── Windows Settings
           └── Scripts (Startup/Shutdown)
   ```

2. **Dobbeltklik på "Startup"**

3. **Klik "Add":**
   - Script Name: `\\server\OutlookAddIn\Deploy-AddIn-Silent.ps1`
   - Script Parameters: `-NetworkShare "\\server\OutlookAddIn"`

4. **Klik OK**

### Trin 5: Konfigurer GPO Scope

1. **Gå tilbage til GPMC**

2. **Vælg din GPO**

3. **Under "Scope":**
   - Security Filtering: Tilføj "Domain Computers" eller specifik gruppe
   - Link til relevante OUs

### Trin 6: Test Deployment

**Test på én maskine først:**

1. Tilføj test maskine til target OU
2. På test maskinen, kør:
   ```cmd
   gpupdate /force
   ```
3. Genstart maskinen
4. Startup script kører ved boot
5. Log ind og åbn Outlook
6. Opret ny mødeaftale → Task pane vises automatisk!

**Check logs:**
```powershell
# På klient maskinen
Get-Content "$env:TEMP\OutlookAddIn-Install.log"
```

### Trin 7: Rollout

Når test er succesfuld:
1. Link GPO til produktions OUs
2. Maskiner vil installere add-in ved næste genstart
3. Monitor Event Logs for fejl

## 🔍 Monitoring og Troubleshooting

### Check Installation Status

**Via PowerShell (kør remote):**
```powershell
Invoke-Command -ComputerName PC001,PC002,PC003 -ScriptBlock {
    $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
    if (Test-Path $regPath) {
        Write-Host "$env:COMPUTERNAME - Installeret"
        Get-ItemProperty -Path $regPath
    } else {
        Write-Host "$env:COMPUTERNAME - IKKE installeret" -ForegroundColor Red
    }
}
```

### Check VSTO Runtime

```powershell
Invoke-Command -ComputerName PC001 -ScriptBlock {
    Test-Path "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
}
```

### View Installation Logs

```powershell
# På klient maskine
Get-Content "$env:TEMP\OutlookAddIn-Install.log" -Tail 50
```

### Common Issues

**Problem: Script kører ikke**
- Check GPO link status
- Check Security Filtering
- Check script path er korrekt UNC path
- Check klient har adgang til network share

**Problem: Access Denied**
- Check share permissions ("Domain Computers" skal have Read)
- Check NTFS permissions på filer

**Problem: VSTO Runtime installation fejler**
- Deploy VSTO Runtime separat via GPO software installation
- Eller brug MSI installer til add-in der bundler VSTO Runtime

## 📦 Alternativ: Software Installation GPO

I stedet for startup script, kan du pakke add-in som MSI:

### Opret MSI Installer

1. **Brug WiX Toolset eller Visual Studio Installer Projects**

2. **MSI skal:**
   - Installere VSTO Runtime (eller check om det findes)
   - Kopiere add-in filer til Program Files
   - Oprette registry keys
   - Sætte permissions

3. **Deploy via GPO:**
   ```
   Computer Configuration
   └── Policies
       └── Software Settings
           └── Software installation
               → Right-click → New → Package
               → Browse to MSI on network share
               → Choose "Assigned"
   ```

### Fordele ved MSI Deployment

- ✅ Native Windows Installer support
- ✅ Automatisk rollback ved fejl
- ✅ Windows Update integration
- ✅ Reporting via GPMC

## 🔄 Opdateringer

**Deploy ny version:**

1. Upload nye filer til network share
2. Opdater version number i script/MSI
3. GPO registrerer automatisk ny version
4. Installer ved næste reboot/login

**Eller brug samme script:**
```powershell
# På klient via GPO startup script
Deploy-AddIn-Silent.ps1 -NetworkShare "\\server\OutlookAddIn" -Force
```

## 🗑️ Afinstallation

**Via GPO:**

1. Opret afinstallations GPO:
   - Startup script: `Deploy-AddIn-Silent.ps1 -Uninstall`

2. Eller disable/fjern installations GPO

3. Kør cleanup script:
   ```powershell
   Invoke-Command -ComputerName (Get-ADComputer -Filter *) -ScriptBlock {
       $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
       if (Test-Path $regPath) {
           Remove-Item -Path $regPath -Recurse -Force
       }
   }
   ```

## 📊 Reporting

**Check deployment status:**

```powershell
# Get alle domain computers
$computers = (Get-ADComputer -Filter * -SearchBase "OU=Workstations,DC=domain,DC=com").Name

# Check installation på hver
$results = @()
foreach ($computer in $computers) {
    $status = Invoke-Command -ComputerName $computer -ScriptBlock {
        Test-Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookMeetingAddin"
    } -ErrorAction SilentlyContinue

    $results += [PSCustomObject]@{
        Computer = $computer
        Installed = $status
    }
}

# Vis resultat
$results | Format-Table -AutoSize

# Export til CSV
$results | Export-Csv -Path "AddIn-Deployment-Status.csv" -NoTypeInformation
```

## 🎓 Best Practices

1. **Test først** på pilot gruppe
2. **Stage rollout** - deploy gradvist
3. **Monitor logs** for fejl
4. **Kommuniker** med brugere om ny funktionalitet
5. **Document** deployment detaljer
6. **Backup** GPO før changes

## 📞 Support

Ved problemer, check:
- GPO replication status
- Network share accessibility
- Klient event logs
- VSTO Runtime installation
- Outlook version compatibility

---

**Next:** Se [INTUNE-DEPLOYMENT.md](INTUNE-DEPLOYMENT.md) for cloud-baseret deployment alternativ.
