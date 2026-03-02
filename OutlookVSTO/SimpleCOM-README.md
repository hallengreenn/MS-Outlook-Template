# Alternativ Løsning: Simpel COM Add-in

## ⚠️ Problemet med VSTO

VSTO (Visual Studio Tools for Office) projekter kræver:
- Office Developer Tools installeret i Visual Studio
- VSTO Runtime
- Kompleks setup og deployment
- ClickOnce eller Windows Installer

Dette gør det besværligt at bygge uden fuld Visual Studio installation med Office workload.

## ✅ Bedre Løsning: Form Region Add-in

I stedet for VSTO's Custom Task Pane, kan vi bruge **Outlook Form Regions** som er:
- Native Outlook funktionalitet
- Lettere at deploye
- Automatisk visning er indbygget
- Kræver ikke VSTO runtime

## 🎯 Tre Tilgange (Rangeret efter Lettilgængelighed)

### Option 1: Outlook Form Region (ANBEFALET FOR DIG!)

**Fordele:**
- ✅ Bygges direkte i Outlook (ingen Visual Studio nødvendigt!)
- ✅ Vises automatisk når møde oprettes
- ✅ Ingen installation - bare import af en fil
- ✅ Kan distribueres via `.OFS` fil

**Ulemper:**
- ❌ Design værktøj er basalt (Outlook's form designer)
- ❌ Kan ikke have kompleks logik
- ❌ Begrænset til Outlook's form controls

**Sådan oprettes det:**
1. Åbn Outlook
2. Gå til Developer tab (aktivér det først under File → Options → Customize Ribbon)
3. Klik "Design a Form"
4. Vælg "Appointment"
5. Tilføj en ny region med skabelon tekst
6. Gem som `.OFS` fil
7. Distribuer til brugere

### Option 2: VBA Macro Add-in

**Fordele:**
- ✅ Bygget direkte i Outlook
- ✅ Ingen kompilering nødvendig
- ✅ Kan automatisk indsætte tekst når møde oprettes

**Ulemper:**
- ❌ Kræver at macros er aktiveret
- ❌ Sikkerhedsadvarsler

**Kode eksempel:**
```vb
' I Outlook VBA Editor (Alt+F11)
Private WithEvents AppointmentInspectors As Outlook.Inspectors

Private Sub Application_Startup()
    Set AppointmentInspectors = Application.Inspectors
End Sub

Private Sub AppointmentInspectors_NewInspector(ByVal Inspector As Inspector)
    If TypeOf Inspector.CurrentItem Is AppointmentItem Then
        Dim appt As AppointmentItem
        Set appt = Inspector.CurrentItem

        ' Auto-insert template hvis body er tom
        If appt.Body = "" Then
            appt.Body = "Formål med mødet" & vbCrLf & _
                       "Kort beskrivelse..." & vbCrLf & vbCrLf & _
                       "Dagsorden/emner" & vbCrLf & _
                       "Liste over punkter..." & vbCrLf & vbCrLf & _
                       "Roller og ansvar" & vbCrLf & _
                       "Hvem er mødeleder..." & vbCrLf & vbCrLf & _
                       "Beslutninger og næste skridt" & vbCrLf & _
                       "Afslut med opsummering..." & vbCrLf & vbCrLf & _
                       "Evt." & vbCrLf & _
                       "Tid til spørgsmål..."
        End If
    End If
End Sub
```

### Option 3: VSTO Add-in (OPRINDELIG PLAN)

Dette er hvad jeg har oprettet, men kræver:
- Visual Studio med Office Development workload
- VSTO Runtime
- Kompleks deployment

**Hvis du vil fortsætte med VSTO:**
1. Installer Office Development Tools i Visual Studio 18
2. Genåbn projektet
3. Build vil fungere automatisk

## 🚀 Min Anbefaling

**For dig vil jeg foreslå Option 2 (VBA Macro)**

Hvorfor?
1. Du kan implementere det NU - ingen build nødvendigt
2. Det vil automatisk indsætte skabelonen
3. Det er nemt at distribuere (bare en .otm fil)
4. Du kan teste det med det samme

## 📝 Vil du have mig til at guide dig gennem VBA løsningen?

Den er meget simplere og du kan have det kørende på under 5 minutter!

Alternativt kan vi:
- Oprette en Form Region (kræver Outlook Developer tab)
- Fortsætte med VSTO (kræver mere setup)
- Lave en PowerShell script der overvåger Outlook via COM (avanceret)

Hvad foretrækker du?
