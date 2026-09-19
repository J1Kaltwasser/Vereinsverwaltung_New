Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================
' frm_SQL_Importer (FORMULARMODUL)
' ============================================

Option Compare Database
Option Explicit

Private Sub Form_Load()
    ' DragAcceptFiles entfernt - funktioniert nicht in Access
    Me.txtFilePath = "C:\SQL\"
    Me.chkImportQueries.Value = -1  ' Checked
    Me.lstLog.RowSource = ""
End Sub
' ============================================
' cmdSelectFile_Click (EMPFOHLEN)
' ============================================

Private Sub cmdSelectFile_Click()

    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim useDialog As VbMsgBoxResult

    ' Frage Benutzer: Dialog oder Pfad eingeben
    useDialog = MsgBox( _
        "Wie möchten Sie die Datei auswählen?" & vbCrLf & vbCrLf & _
        "Ja = Datei-Dialog (empfohlen)" & vbCrLf & _
        "Nein = Pfad manuell eingeben", _
        vbYesNo + vbQuestion, "Dateiauswahl")

    If useDialog = vbYes Then
        ' Versuchen Sie Datei-Dialog
        filePath = TryFileDialog()

        If filePath = "" Then
            ' Dialog fehlgeschlagen - InputBox verwenden
            Me.lstLog.AddItem "? Datei-Dialog fehlgeschlagen, verwende manuelle Eingabe"
            filePath = InputBox( _
                "Bitte geben Sie den Pfad zur SQL-Datei ein:" & vbCrLf & _
                "Beispiel: C:\SQL\test.sql", _
                "SQL-Datei auswählen")
        End If
    Else
        ' Benutzer bevorzugt manuelle Eingabe
        filePath = InputBox( _
            "Bitte geben Sie den Pfad zur SQL-Datei ein:" & vbCrLf & _
            "Beispiel: C:\SQL\test.sql", _
            "SQL-Datei auswählen")
    End If

    If filePath = "" Then
        Me.lstLog.AddItem "? Dateiauswahl abgebrochen"
        Exit Sub
    End If

    ' Datei prüfen
    If Dir(filePath) = "" Then
        MsgBox "Datei nicht gefunden: " & filePath, vbExclamation
        Me.lstLog.AddItem "? Datei nicht gefunden: " & filePath
        Exit Sub
    End If

    ' Datei-Endung prüfen
    If LCase(Right(filePath, 4)) <> ".sql" Then
        MsgBox "Bitte nur SQL-Dateien (*.sql) auswählen.", vbExclamation
        Me.lstLog.AddItem "? Falsche Datei-Endung"
        Exit Sub
    End If

    Me.txtFilePath = filePath
    LoadSQLPreview filePath
    Me.lstLog.AddItem "? Datei gewählt: " & filePath

    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
    Me.lstLog.AddItem "? Fehler: " & Err.Description

End Sub

' ============================================
' HILFSFUNKTION: Datei-Dialog via Excel
' ============================================

Private Function TryFileDialog() As String

    Dim xl As Object
    Dim filePath As String

    On Error GoTo ErrorHandler

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    filePath = xl.FileDialog(4, "SQL-Datei auswählen", , 2).FileName

    xl.Quit
    Set xl = Nothing

    TryFileDialog = filePath
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not xl Is Nothing Then xl.Quit
    TryFileDialog = ""

End Function

Private Sub cmdImportSQL_Click()

    On Error GoTo ErrorHandler

    Dim path As String
    path = Nz(Me.txtFilePath, "")

    If path = "" Then
        MsgBox "Bitte zuerst eine SQL-Datei auswählen.", vbExclamation
        Exit Sub
    End If

    If Dir(path) = "" Then
        MsgBox "Datei existiert nicht: " & path, vbExclamation
        Exit Sub
    End If

    ' ListBox leeren
    Me.lstLog.RowSource = ""

    ' Import starten
    ImportSQLFromFile_GUI path, Me.lstLog, Me.chkImportTables, Me.chkImportQueries

    Exit Sub

ErrorHandler:
    Me.lstLog.AddItem "? FEHLER: " & Err.Description
    MsgBox "Fehler: " & Err.Description, vbCritical

End Sub


Public Sub LoadSQLPreview(filePath As String)

    On Error GoTo ErrorHandler

    Dim raw As String
    raw = ReadTextFileRobust(filePath)

    If Len(raw) > 5000 Then
        ' Text kürzen zur Anzeige
        Me.txtPreview = Left$(raw, 5000) & vbCrLf & "... (Text gekürzt)"
    Else
        Me.txtPreview = raw
    End If

    Exit Sub

ErrorHandler:
    Me.txtPreview = "FEHLER beim Laden der Vorschau: " & Err.Description

End Sub

' ============================================
' cmdClearLog_Click Event
' ============================================

Private Sub cmdClearLog_Click()

    On Error GoTo ErrorHandler

    ' ListBox leeren
    Me.lstLog.RowSource = ""

    ' Bestätigung
    Me.lstLog.AddItem "Log gelöscht."

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Leeren des Logs: " & Err.Description, vbCritical

End Sub

' ============================================
' cmdOpenInDesign_Click Event
' ============================================

Private Sub cmdOpenInDesign_Click()

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim latestName As String
    Dim i As Long

    Set db = CurrentDb()

    ' Letzte importierte Abfrage suchen
    latestName = ""

    For i = 0 To db.QueryDefs.Count - 1
        If InStr(db.QueryDefs(i).name, "qry_SQL_") > 0 Then
            latestName = db.QueryDefs(i).name
        End If
    Next i

    If latestName = "" Then
        MsgBox "Keine importierte Abfrage gefunden." & vbCrLf & _
               "Bitte führen Sie erst einen Import durch.", vbExclamation
        Me.lstLog.AddItem "? Keine importierte Abfrage vorhanden"
        Exit Sub
    End If

    ' Abfrage in Entwurfsansicht öffnen
    DoCmd.OpenQuery latestName, acViewDesign

    Me.lstLog.AddItem "? Abfrage geöffnet: " & latestName

    Exit Sub

ErrorHandler:
    Me.lstLog.AddItem "? Fehler: " & Err.Description
    MsgBox "Fehler beim Öffnen der Abfrage: " & Err.Description, vbCritical

End Sub
