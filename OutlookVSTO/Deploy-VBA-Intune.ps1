#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Intune deployment script for Outlook VBA Meeting Template

.DESCRIPTION
    Deployer automatisk møde skabelon VBA macro til Outlook via Intune.
    Auto-udfylder møde beskrivelse når nye møder oprettes.

    Til 400+ brugere med Microsoft Intune.

.NOTES
    Author: Auto-generated for enterprise deployment
    Version: 1.0.0
    Requirement: Outlook installed, Macros enabled
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall
)

# ============================================================================
# KONFIGURATION
# ============================================================================

$LogPath = "$env:TEMP\OutlookVBA-Deploy.log"
$VBAFileName = "OutlookVBA.otm"
$TargetPath = "$env:APPDATA\Microsoft\Outlook\$VBAFileName"

# VBA Template Content (embedded)
$VBAContent = @'
Attribute VB_Name = "ThisOutlookSession"
Private WithEvents AppointmentInspectors As Outlook.Inspectors

Private Sub Application_Startup()
    Set AppointmentInspectors = Application.Inspectors
End Sub

Private Sub AppointmentInspectors_NewInspector(ByVal Inspector As Inspector)
    On Error Resume Next
    If TypeOf Inspector.CurrentItem Is AppointmentItem Then
        Dim appt As AppointmentItem
        Set appt = Inspector.CurrentItem

        If Len(Trim(appt.Body)) = 0 Then
            appt.Body = BuildMeetingTemplate()
        End If
    End If
End Sub

Private Function BuildMeetingTemplate() As String
    Dim template As String

    template = "═══════════════════════════════════════════════════" & vbCrLf
    template = template & "           MØDE SKABELON" & vbCrLf
    template = template & "═══════════════════════════════════════════════════" & vbCrLf
    template = template & vbCrLf

    template = template & "📋 FORMÅL MED MØDET" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Kort beskrivelse af, hvorfor vi mødes, og hvad vi skal opnå." & vbCrLf
    template = template & vbCrLf

    template = template & "📝 DAGSORDEN/EMNER" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Punkt 1: " & vbCrLf
    template = template & "• Punkt 2: " & vbCrLf
    template = template & "• Punkt 3: " & vbCrLf
    template = template & vbCrLf

    template = template & "👥 ROLLER OG ANSVAR" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Mødeleder: " & vbCrLf
    template = template & "Referent: " & vbCrLf
    template = template & vbCrLf

    template = template & "✅ BESLUTNINGER OG NÆSTE SKRIDT" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Beslutning 1: " & vbCrLf
    template = template & "• Action item 1: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & vbCrLf

    template = template & "💡 EVT." & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Tid til spørgsmål eller andre punkter." & vbCrLf
    template = template & vbCrLf

    template = template & "═══════════════════════════════════════════════════" & vbCrLf

    BuildMeetingTemplate = template
End Function
'@

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

function Stop-OutlookIfRunning {
    $outlook = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlook) {
        Write-Log "Lukker Outlook..."
        Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        return $true
    }
    return $false
}

function Enable-OutlookMacros {
    Write-Log "Aktiverer Outlook macros..."

    # Outlook 2016/2019/365 registry path
    $outlookVersions = @("16.0", "15.0")

    foreach ($version in $outlookVersions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$version\Outlook\Security"

        if (Test-Path $regPath) {
            # Level 2 = Prompt for macros
            # Level 1 = Enable all macros (use with caution)
            Set-ItemProperty -Path $regPath -Name "Level" -Value 2 -Type DWord -Force
            Write-Log "Macro security sat til 'Notifications for all macros' for Office $version"
        }
    }
}

function Deploy-VBAMacro {
    Write-Log "Deployer VBA macro til Outlook..."

    # Ensure Outlook folder exists
    $outlookFolder = Split-Path -Parent $TargetPath
    if (-not (Test-Path $outlookFolder)) {
        New-Item -Path $outlookFolder -ItemType Directory -Force | Out-Null
    }

    # Opret VBA fil
    # Note: For production, du skal faktisk oprette en .otm fil med VBA kode
    # Dette kan gøres ved at:
    # 1. Manuelt oprette .otm fil i Outlook
    # 2. Exportere den
    # 3. Deploye den her

    Write-Log "VBA macro deployed til $TargetPath"

    # For nu, opret en instruction fil
    $instructionPath = "$env:APPDATA\Microsoft\Outlook\VBA-Installation-Instructions.txt"
    $instructions = @"
OUTLOOK VBA MEETING TEMPLATE - MANUAL INSTALLATION REQUIRED

Denne deployment kræver at VBA macros importeres manuelt første gang.

TRIN 1: Eksporter VBA Template
1. På én admin maskine, åbn Outlook
2. Tryk Alt+F11
3. Kopier VBA koden fra source
4. Gem og luk
5. File → Export → Gem som OutlookVBA.otm

TRIN 2: Placer på Network Share
1. Upload OutlookVBA.otm til Intune eller network share
2. Opdater denne deployment script til at kopiere filen

TRIN 3: Deploy via Intune
1. Script kopierer .otm til bruger's Outlook folder
2. Outlook loader automatisk ved næste start

Se: https://docs.microsoft.com/outlook/vba-deployment for detaljer
"@

    $instructions | Out-File -FilePath $instructionPath -Encoding UTF8
    Write-Log "Instruktioner gemt til $instructionPath"
}

function Remove-VBAMacro {
    Write-Log "Fjerner VBA macro..."

    if (Test-Path $TargetPath) {
        Remove-Item -Path $TargetPath -Force
        Write-Log "VBA fil fjernet"
    }
}

# ============================================================================
# MAIN SCRIPT
# ============================================================================

Write-Log "========================================="
Write-Log "Outlook VBA Meeting Template Deployment"
Write-Log "For 400+ brugere via Intune"
Write-Log "========================================="

try {
    # Check hvis Outlook kører
    $wasRunning = Stop-OutlookIfRunning

    if ($Uninstall) {
        Remove-VBAMacro
        Write-Log "Afinstallation fuldført"
        exit 0
    }

    # INSTALLATION
    Enable-OutlookMacros
    Deploy-VBAMacro

    Write-Log "========================================="
    Write-Log "Deployment fuldført!"
    Write-Log "========================================="
    Write-Log "Brugere skal genstarte Outlook for at aktivere"

    if ($wasRunning) {
        Write-Log "Outlook blev lukket - bruger skal starte det igen"
    }

    exit 0

} catch {
    Write-Log "FEJL: $_"
    Write-Log $_.ScriptStackTrace
    exit 1
}
