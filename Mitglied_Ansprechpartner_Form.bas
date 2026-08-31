Attribute VB_Name = "Mitglied_Ansprechpartner_Form"
' ========================================
' FORMULAR: Mitglied_Ansprechpartner_Form
' ========================================

Private Sub Form_Load()
    Me.ÄnderungsDatum = Now()
    Me.ÄnderungsBenutzer = Environ("USERNAME")
    Me.GültigAb = Date
End Sub

Private Sub Vorname_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.Vorname) Or Me.Vorname = "" Then
        MsgBox "Vorname ist erforderlich!", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Nachname_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.Nachname) Or Me.Nachname = "" Then
        MsgBox "Nachname ist erforderlich!", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Email_BeforeUpdate(Cancel As Integer)
    Dim Email As String
    If Not IsNull(Me.Email) And Me.Email <> "" Then
        Email = Me.Email
        If InStr(Email, "@") = 0 Or InStr(Email, ".") = 0 Then
            MsgBox "Ungültiges E-Mail-Format!", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub IstHauptansprechpartner_AfterUpdate()
    If Me.IstHauptansprechpartner = True Then
        Dim db As Object
        Set db = CurrentDb()
        db.Execute "UPDATE Mitglied_Ansprechpartner SET IstHauptansprechpartner = False " & _
            "WHERE Mitgliedsnummer='" & Me.Mitgliedsnummer & "' " & _
            "AND AnsprechpartnerID <> " & Me.AnsprechpartnerID & " " & _
            "AND GültigBis IS NULL"
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.Mitgliedsnummer) Or IsNull(Me.GeschlechtID) Or _
       IsNull(Me.Vorname) Or IsNull(Me.Nachname) Or _
       IsNull(Me.StellungsID) Or IsNull(Me.GültigAb) Then
        MsgBox "Bitte alle Pflichtfelder ausfüllen!", vbExclamation
        Cancel = True
    End If
End Sub
