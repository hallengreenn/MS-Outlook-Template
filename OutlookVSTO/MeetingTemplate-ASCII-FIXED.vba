Attribute VB_Name = "ThisOutlookSession"
' ==============================================================================
' Outlook VBA Macro - Automatisk Moede Skabelon (ASCII Version)
' ==============================================================================
' FIXED: Kun ASCII karakterer - ingen unicode problemer
' ==============================================================================

Private WithEvents AppointmentInspectors As Outlook.Inspectors

' Koer ved Outlook start
Private Sub Application_Startup()
    Set AppointmentInspectors = Application.Inspectors
End Sub

' Overvaag nye vinduer
Private Sub AppointmentInspectors_NewInspector(ByVal Inspector As Inspector)
    ' Check om det er en appointment
    If TypeOf Inspector.CurrentItem Is AppointmentItem Then
        Dim appt As AppointmentItem
        Set appt = Inspector.CurrentItem

        ' Indsaet skabelon hvis body er tom (ny moede)
        If Len(Trim(appt.Body)) = 0 Then
            appt.Body = BuildMeetingTemplate()
        End If
    End If
End Sub

' Funktion der bygger moede skabelonen (KUN ASCII!)
Private Function BuildMeetingTemplate() As String
    Dim template As String

    template = "======================================================" & vbCrLf
    template = template & "              MOEDE SKABELON" & vbCrLf
    template = template & "======================================================" & vbCrLf
    template = template & vbCrLf

    template = template & "FORMAAL MED MOEDET" & vbCrLf
    template = template & "------------------------------------------------------" & vbCrLf
    template = template & "Kort beskrivelse af, hvorfor vi moedes, og hvad vi skal opnaa." & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "DAGSORDEN/EMNER" & vbCrLf
    template = template & "------------------------------------------------------" & vbCrLf
    template = template & "* Punkt 1: " & vbCrLf
    template = template & "* Punkt 2: " & vbCrLf
    template = template & "* Punkt 3: " & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "ROLLER OG ANSVAR" & vbCrLf
    template = template & "------------------------------------------------------" & vbCrLf
    template = template & "Moeideleder: " & vbCrLf
    template = template & "Referent: " & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "BESLUTNINGER OG NAESTE SKRIDT" & vbCrLf
    template = template & "------------------------------------------------------" & vbCrLf
    template = template & "* Beslutning 1: " & vbCrLf
    template = template & "* Action item 1: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & "* Action item 2: [Ansvarlig] [Deadline]" & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "EVT." & vbCrLf
    template = template & "------------------------------------------------------" & vbCrLf
    template = template & "Tid til spoergsmaal eller andre punkter." & vbCrLf
    template = template & vbCrLf
    template = template & vbCrLf

    template = template & "======================================================" & vbCrLf

    BuildMeetingTemplate = template
End Function
