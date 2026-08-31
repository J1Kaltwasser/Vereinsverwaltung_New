Attribute VB_Name = "Mitglied_Adressen_Form"
' ========================================
' FORMULAR: Mitglied_Adressen_Form
' ========================================

Private Sub Form_Load()
    ' Beim Laden nichts nötig - Standardwerte kommen aus der Tabelle
End Sub

Private Sub Form_Current()
    ' Bei Datensatzwechsel: LandID vorwählen
    If Me.NewRecord Then
        If IsNull(Me.LandID) Then
            Me.LandID = 1 ' Deutschland
        End If
    End If
End Sub

Private Sub AdressTypID_AfterUpdate()
    ' Bedingte Anzeige: Straße oder Postfach
    Select Case Me.AdressTypID
        Case 1 ' Straßenadresse
            Me.Straße.Visible = True
            Me.Hausnummer.Visible = True
            Me.Postfachnummer.Visible = False
            Me.Postfachnummer.Value = Null
        Case 2 ' Postfachadresse
            Me.Straße.Visible = False
            Me.Hausnummer.Visible = False
            Me.Straße.Value = Null
            Me.Hausnummer.Value = Null
            Me.Postfachnummer.Visible = True
    End Select
End Sub

Private Sub Straße_BeforeUpdate(Cancel As Integer)
    If Me.AdressTypID = 1 Then
        If IsNull(Me.Straße) Or Me.Straße = "" Then
            MsgBox "Straße ist erforderlich!", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub Postfachnummer_BeforeUpdate(Cancel As Integer)
    If Me.AdressTypID = 2 Then
        If IsNull(Me.Postfachnummer) Or Me.Postfachnummer = "" Then
            MsgBox "Postfachnummer ist erforderlich!", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub Postleitzahl_BeforeUpdate(Cancel As Integer)
    Dim PLZ As String
    Dim Ländercode As String

    If Not IsNull(Me.Postleitzahl) And Me.Postleitzahl <> "" Then
        PLZ = Me.Postleitzahl
        Ländercode = DLookup("Ländercode", "Länder", "LandID=" & Me.LandID)

        Select Case Ländercode
            Case "DE"
                If Len(PLZ) <> 5 Or Not IsNumeric(PLZ) Then
                    MsgBox "Deutsche PLZ: 5 Ziffern!", vbExclamation
                    Cancel = True
                End If
            Case "AT"
                If Len(PLZ) <> 4 Or Not IsNumeric(PLZ) Then
                    MsgBox "Österreichische PLZ: 4 Ziffern!", vbExclamation
                    Cancel = True
                End If
            Case "CH"
                If Len(PLZ) <> 4 Or Not IsNumeric(PLZ) Then
                    MsgBox "Schweizer PLZ: 4 Ziffern!", vbExclamation
                    Cancel = True
                End If
        End Select
    End If
End Sub

Private Sub Ortsname_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.Ortsname) Or Me.Ortsname = "" Then
        MsgBox "Ortsname ist erforderlich!", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Alle Pflichtfelder prüfen
    If IsNull(Me.Mitgliedsnummer) Or IsNull(Me.AdressTypID) Or _
       IsNull(Me.VerwendungsID) Or IsNull(Me.Postleitzahl) Or _
       IsNull(Me.Ortsname) Or IsNull(Me.GültigAb) Then
        MsgBox "Bitte alle Pflichtfelder ausfüllen!", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub

Function IsNumeric(strValue As String) As Boolean
    On Error GoTo ErrorHandler
    Dim dblValue As Double
    dblValue = CDbl(strValue)
    IsNumeric = True
    Exit Function
ErrorHandler:
    IsNumeric = False
End Function
