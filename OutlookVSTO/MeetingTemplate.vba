' ==============================================================================
' Outlook VBA Macro - Automatisk Møde Skabelon
' ==============================================================================
'
' INSTALLATION:
' 1. Åbn Outlook
' 2. Tryk Alt+F11 for at åbne VBA Editor
' 3. I Project Explorer, find "Project1" → "Microsoft Outlook Objects" → "ThisOutlookSession"
' 4. Dobbeltklik på "ThisOutlookSession"
' 5. Kopier HELE denne kode ind
' 6. Gem (Ctrl+S)
' 7. Luk VBA Editor
' 8. Genstart Outlook
'
' BRUG:
' - Opret en ny mødeaftale
' - Skabelonen indsættes automatisk hvis body er tom!
'
' ==============================================================================

Private WithEvents AppointmentInspectors As Outlook.Inspectors

' Kør ved Outlook start
Private Sub Application_Startup()
    Set AppointmentInspectors = Application.Inspectors
End Sub

' Overvåg nye vinduer
Private Sub AppointmentInspectors_NewInspector(ByVal Inspector As Inspector)
    ' Check om det er en appointment
    If TypeOf Inspector.CurrentItem Is AppointmentItem Then
        Dim appt As AppointmentItem
        Set appt = Inspector.CurrentItem

        ' Indsæt skabelon hvis body er tom (ny møde)
        If Len(Trim(appt.Body)) = 0 Then
            appt.Body = BuildMeetingTemplate()
        End If
    End If
End Sub

' Funktion der bygger møde skabelonen
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
    template = template & vbCrLf

    template = template & "📝 DAGSORDEN/EMNER" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Punkt 1: " & vbCrLf
    template = template & "• Punkt 2: " & vbCrLf
    template = template & "• Punkt 3: " & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "👥 ROLLER OG ANSVAR" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Mødeleder: " & vbCrLf
    template = template & "Referent: " & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "✅ BESLUTNINGER OG NÆSTE SKRIDT" & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "• Beslutning 1: " & vbCrLf
    template = template & "• Action item 1: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & "• Action item 2: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "💡 EVT." & vbCrLf
    template = template & "───────────────────────────────────────────────────" & vbCrLf
    template = template & "Tid til spørgsmål eller andre punkter." & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "═══════════════════════════════════════════════════" & vbCrLf

    BuildMeetingTemplate = template
End Function

' Optional: Tilføj en knap i Ribbon til manuelt at indsætte skabelon
Public Sub InsertMeetingTemplate()
    On Error Resume Next

    Dim Inspector As Outlook.Inspector
    Dim appt As AppointmentItem

    Set Inspector = Application.ActiveInspector

    If Not Inspector Is Nothing Then
        If TypeOf Inspector.CurrentItem Is AppointmentItem Then
            Set appt = Inspector.CurrentItem
            appt.Body = BuildMeetingTemplate()
            MsgBox "Møde skabelon indsat!", vbInformation, "Succes"
        Else
            MsgBox "Dette er ikke en mødeaftale.", vbExclamation, "Fejl"
        End If
    Else
        MsgBox "Åbn venligst en mødeaftale først.", vbExclamation, "Fejl"
    End If
End Sub
