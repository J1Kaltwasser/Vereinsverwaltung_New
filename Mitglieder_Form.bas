Attribute VB_Name = "Mitglieder_Form"
' ========================================
' FORMULAR: Mitglieder_Form
' ========================================

Private Sub Form_Load()
    ' Standard-Status setzen
    Me.Status.DefaultValue = """Aktiv"""
End Sub

Private Sub Name1_BeforeUpdate(Cancel As Integer)
    ' Name1 (Nachname) ist erforderlich
    If IsNull(Me.Name1) Or Me.Name1 = "" Then
        MsgBox "Name1 (Nachname) ist erforderlich!", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Beitrittsdatum_BeforeUpdate(Cancel As Integer)
    ' Beitrittsdatum validieren
    If IsNull(Me.Beitrittsdatum) Then
        MsgBox "Beitrittsdatum ist erforderlich!", vbExclamation
        Cancel = True
    ElseIf Me.Beitrittsdatum > Date Then
        MsgBox "Beitrittsdatum kann nicht in der Zukunft liegen!", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub AusscheidungsDatum_AfterUpdate()
    ' Ausscheidungsfelder anzeigen/verstecken
    If Not IsNull(Me.AusscheidungsDatum) Then
        Me.AusscheidungsgrundID.Visible = True
        Me.AusscheidungsDatum.BackColor = RGB(200, 255, 200)
    Else
        Me.AusscheidungsgrundID.Visible = False
        Me.KündigungsDatum.Visible = False
        Me.Kündigungswirksam_ab.Visible = False
        Me.FusionÜbernehmer_Name.Visible = False
        Me.FusionÜbernehmer_ID.Visible = False
    End If
End Sub

Private Sub AusscheidungsgrundID_AfterUpdate()
    ' Je nach Grund unterschiedliche Felder anzeigen
    Me.KündigungsDatum.Visible = False
    Me.Kündigungswirksam_ab.Visible = False
    Me.FusionÜbernehmer_Name.Visible = False
    Me.FusionÜbernehmer_ID.Visible = False

    Select Case Me.AusscheidungsgrundID
        Case 3 ' Kündigung
            Me.KündigungsDatum.Visible = True
            Me.Kündigungswirksam_ab.Visible = True
        Case 5 ' Fusion
            Me.FusionÜbernehmer_Name.Visible = True
            Me.FusionÜbernehmer_ID.Visible = True
    End Select
End Sub

Private Sub Kündigungswirksam_ab_BeforeUpdate(Cancel As Integer)
    ' Kündigung nur zum Jahresende mit 6 Monaten Vorlauf
    Dim WirksamDatum As Date
    Dim MinimalDate As Date

    If Not IsNull(Me.Kündigungswirksam_ab) Then
        WirksamDatum = Me.Kündigungswirksam_ab

        ' Nur Jahresende (31.12.)
        If Month(WirksamDatum) <> 12 Or Day(WirksamDatum) <> 31 Then
            MsgBox "Kündigung nur zum Jahresende (31.12.) möglich!", vbExclamation
            Cancel = True
            Exit Sub
        End If

        ' Mindestens 6 Monate Vorlauf
        MinimalDate = DateAdd("m", -6, WirksamDatum)

        If Not IsNull(Me.KündigungsDatum) Then
            If Me.KündigungsDatum > MinimalDate Then
                MsgBox "Kündigungsfrist: mindestens 6 Monate vor dem Jahresende erforderlich!" & vbCrLf & _
                       "Frühestes Kündigungsdatum: " & Format(MinimalDate, "dd.mm.yyyy"), vbExclamation
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub FusionÜbernehmer_ID_BeforeUpdate(Cancel As Integer)
    ' Fusion - Übernehmer validieren
    If Not IsNull(Me.FusionÜbernehmer_ID) And Me.FusionÜbernehmer_ID <> "" Then
        If DCount("Mitgliedsnummer", "Mitglieder", _
            "Mitgliedsnummer='" & Me.FusionÜbernehmer_ID & "'") = 0 Then
            MsgBox "Mitgliedsnummer des Übernehmers existiert nicht!", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Gesamtvalidierung vor Speichern
    If Not IsNull(Me.AusscheidungsDatum) Then
        If IsNull(Me.AusscheidungsgrundID) Then
            MsgBox "Ausscheidungsgrund ist erforderlich!", vbExclamation
            Cancel = True
            Exit Sub
        End If

        If Me.AusscheidungsgrundID = 3 Then
            If IsNull(Me.KündigungsDatum) Or IsNull(Me.Kündigungswirksam_ab) Then
                MsgBox "Bei Kündigung sind Kündigungsdatum und Wirksamkeitsdatum erforderlich!", vbExclamation
                Cancel = True
            End If
        End If

        If Me.AusscheidungsgrundID = 5 Then
            If IsNull(Me.FusionÜbernehmer_Name) Or Me.FusionÜbernehmer_Name = "" Then
                MsgBox "Bei Fusion muss der Name des Übernehmers eingegeben werden!", vbExclamation
                Cancel = True
            End If
        End If
    End If
End Sub
