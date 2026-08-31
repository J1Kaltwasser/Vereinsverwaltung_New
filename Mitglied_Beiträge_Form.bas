Attribute VB_Name = "Mitglied_Beiträge_Form"
' ========================================
' FORMULAR: Mitglied_Beiträge_Form
' ========================================

Private Sub Form_Load()
    If Me.NewRecord Then
        Me.Beitragsjahr = Year(Date)
        Me.GültigAb = DateSerial(Me.Beitragsjahr, 1, 1)
        Me.ÄnderungsDatum = Now()
        Me.ÄnderungsBenutzer = Environ("USERNAME")
    End If
End Sub

Private Sub Beitragsjahr_AfterUpdate()
    If Not IsNull(Me.Beitragsjahr) Then
        Me.GültigAb = DateSerial(Me.Beitragsjahr, 1, 1)
    End If
End Sub

Private Sub Betrag_BeforeUpdate(Cancel As Integer)
    Dim MitgliederTypID As Integer
    Dim MinBeitrag As Currency

    If Not IsNull(Me.Mitgliedsnummer) Then
        MitgliederTypID = DLookup("MitgliederTypID", "Mitglieder", _
            "Mitgliedsnummer='" & Me.Mitgliedsnummer & "'")

        If MitgliederTypID = 3 Then
            MinBeitrag = 30
        Else
            MinBeitrag = 200
        End If

        If Not IsNull(Me.Betrag) And Me.Betrag < MinBeitrag Then
            MsgBox "Beitrag zu niedrig! Mindestbeitrag: " & Format(MinBeitrag, "0.00") & " €", _
                   vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.Mitgliedsnummer) Or IsNull(Me.Beitragsjahr) Or _
       IsNull(Me.Betrag) Or IsNull(Me.GültigAb) Then
        MsgBox "Bitte alle Pflichtfelder ausfüllen!", vbExclamation
        Cancel = True
    End If
End Sub
