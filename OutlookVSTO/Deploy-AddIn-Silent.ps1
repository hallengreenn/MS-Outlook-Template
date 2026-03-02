#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Silent deployment script for Outlook Meeting Template Add-in

.DESCRIPTION
    Dette script installerer Outlook Meeting Template Add-in uden brugerinteraktion.
    Kan køres via GPO startup script, Intune, SCCM, eller manuelt.

.PARAMETER AddInPath
    Sti til add-in DLL filen (hvis lokal installation)

.PARAMETER NetworkShare
    UNC sti til netværks share hvor add-in filer ligger

.EXAMPLE
    # Lokal installation
    .\Deploy-AddIn-Silent.ps1 -AddInPath "C:\Temp\OutlookMeetingAddin.dll"

.EXAMPLE
    # Fra netværks share
    .\Deploy-AddIn-Silent.ps1 -NetworkShare "\\server\share\OutlookAddIn"

.EXAMPLE
    # Remote installation til flere maskiner
    Invoke-Command -ComputerName PC001,PC002,PC003 -FilePath .\Deploy-AddIn-Silent.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$AddInPath,

    [Parameter(Mandatory=$false)]
    [string]$NetworkShare,

    [Parameter(Mandatory=$false)]
    [switch]$Uninstall
)

# ============================================================================
# KONFIGURATION
# ============================================================================

$AddInName = "OutlookMeetingAddin"
$AddInFriendlyName = "Møde Skabelon"
$AddInDescription = "Automatisk møde skabelon add-in"
$TargetPath = "$env:ProgramFiles\$AddInName"
$RegistryPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$AddInName"

# Log fil
$LogPath = "$env:TEMP\OutlookAddIn-Install.log"

# ============================================================================
# FUNKTIONER
# ============================================================================

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Host $logMessage
    Add-Content -Path $LogPath -Value $logMessage
}

function Test-OutlookRunning {
    $outlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    return ($null -ne $outlook)
}

function Stop-OutlookProcess {
    Write-Log "Lukker Outlook..."
    Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

function Install-VSTORuntime {
    Write-Log "Checker VSTO Runtime..."

    $vstoInstalled = Test-Path "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"

    if (-not $vstoInstalled) {
        Write-Log "VSTO Runtime ikke fundet. Installerer..."

        # Download VSTO Runtime
        $vstoUrl = "https://download.microsoft.com/download/E/3/1/E315A5C5-9AB2-4E67-ADB0-4D299C633B3E/vstor_redist.exe"
        $vstoInstaller = "$env:TEMP\vstor_redist.exe"

        try {
            Invoke-WebRequest -Uri $vstoUrl -OutFile $vstoInstaller
            Start-Process -FilePath $vstoInstaller -ArgumentList "/quiet" -Wait
            Write-Log "VSTO Runtime installeret"
        } catch {
            Write-Log "FEJL: Kunne ikke installere VSTO Runtime: $_"
            throw
        }
    } else {
        Write-Log "VSTO Runtime allerede installeret"
    }
}

function Copy-AddInFiles {
    param([string]$SourcePath)

    Write-Log "Kopierer add-in filer til $TargetPath..."

    # Opret target directory
    if (-not (Test-Path $TargetPath)) {
        New-Item -Path $TargetPath -ItemType Directory -Force | Out-Null
    }

    # Kopier filer
    Copy-Item -Path "$SourcePath\*" -Destination $TargetPath -Recurse -Force

    Write-Log "Filer kopieret"
}

function Register-AddIn {
    Write-Log "Registrerer add-in i registry..."

    # Opret registry nøgle
    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    # Sæt værdier
    Set-ItemProperty -Path $RegistryPath -Name "Description" -Value $AddInDescription
    Set-ItemProperty -Path $RegistryPath -Name "FriendlyName" -Value $AddInFriendlyName
    Set-ItemProperty -Path $RegistryPath -Name "LoadBehavior" -Value 3 -Type DWord

    $manifestPath = "file:///$($TargetPath -replace '\\','/')/OutlookMeetingAddin.vsto|vstolocal"
    Set-ItemProperty -Path $RegistryPath -Name "Manifest" -Value $manifestPath

    Write-Log "Add-in registreret"
}

function Unregister-AddIn {
    Write-Log "Afregistrerer add-in..."

    if (Test-Path $RegistryPath) {
        Remove-Item -Path $RegistryPath -Recurse -Force
        Write-Log "Registry nøgle fjernet"
    }

    if (Test-Path $TargetPath) {
        Remove-Item -Path $TargetPath -Recurse -Force
        Write-Log "Filer fjernet"
    }

    Write-Log "Add-in afinstalleret"
}

# ============================================================================
# MAIN SCRIPT
# ============================================================================

Write-Log "========================================="
Write-Log "Outlook Meeting Template Add-in Deployment"
Write-Log "========================================="

try {
    # Check hvis Outlook kører
    if (Test-OutlookRunning) {
        Write-Log "ADVARSEL: Outlook kører. Lukker processen..."
        Stop-OutlookProcess
    }

    if ($Uninstall) {
        # AFINSTALLATION
        Unregister-AddIn
        Write-Log "Afinstallation fuldført"
        exit 0
    }

    # INSTALLATION

    # Bestem source path
    if ($NetworkShare) {
        $sourcePath = $NetworkShare
        Write-Log "Bruger netværks share: $NetworkShare"
    } elseif ($AddInPath) {
        $sourcePath = Split-Path -Parent $AddInPath
        Write-Log "Bruger lokal sti: $sourcePath"
    } else {
        # Antag at script kører fra samme folder som add-in
        $sourcePath = Split-Path -Parent $PSCommandPath
        Write-Log "Bruger script directory: $sourcePath"
    }

    # Verify source files exist
    if (-not (Test-Path $sourcePath)) {
        throw "Source path ikke fundet: $sourcePath"
    }

    # Install VSTO Runtime
    Install-VSTORuntime

    # Copy files
    Copy-AddInFiles -SourcePath $sourcePath

    # Register add-in
    Register-AddIn

    Write-Log "========================================="
    Write-Log "Installation fuldført!"
    Write-Log "========================================="
    Write-Log "Add-in vil være tilgængeligt næste gang Outlook startes"

    exit 0

} catch {
    Write-Log "FEJL: $_"
    Write-Log $_.ScriptStackTrace
    exit 1
}
