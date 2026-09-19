Attribute VB_Name = "modSQLImporter_Final"
' ============================================
' modSQLImporter_Final.bas (OPTIMIERT)
' ============================================

Option Compare Database
Option Explicit

' ============================================
' API DECLARATIONS
' ============================================

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ============================================
' TEXTDATEI LESEN (MIT UTF-8)
' ============================================

Public Function ReadTextFileRobust(path As String) As String

    Dim stream As Object
    Dim result As String

    On Error GoTo ErrorHandler

    Set stream = CreateObject("ADODB.Stream")

    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    result = stream.ReadText()
    stream.Close

    ReadTextFileRobust = result
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not (stream Is Nothing) Then stream.Close
    ReadTextFileRobust = ""

End Function

' ============================================
' KOMMENTARE ENTFERNEN (EINFACH)
' ============================================

Public Function RemoveCommentsFromSQL(SQL As String) As String

    Dim lines() As String
    Dim cleanLines() As String
    Dim cleanCount As Long
    Dim i As Long
    Dim line As String
    Dim finalSQL As String

    On Error GoTo ErrorHandler

    lines = Split(SQL, vbCrLf)
    ReDim cleanLines(UBound(lines))
    cleanCount = 0

    For i = LBound(lines) To UBound(lines)
        line = lines(i)

        ' Kommentare entfernen
        If Left(Trim(line), 2) = "--" Then
            GoTo NextLine
        End If

        ' Inline-Kommentare entfernen
        If InStr(line, "--") > 0 Then
            line = Left(line, InStr(line, "--") - 1)
        End If

        ' Nur nicht-leere Zeilen speichern
        If Len(Trim(line)) > 0 Then
            cleanLines(cleanCount) = Trim(line)
            cleanCount = cleanCount + 1
        End If

NextLine:
    Next i

    ' Array auf richtige Groesse trimmen
    If cleanCount > 0 Then
        ReDim Preserve cleanLines(cleanCount - 1)
        finalSQL = Join(cleanLines, " ")
    Else
        finalSQL = SQL
    End If

    ' Mehrfache Leerzeichen entfernen
    While InStr(finalSQL, "  ") > 0
        finalSQL = Replace(finalSQL, "  ", " ")
    Wend

    RemoveCommentsFromSQL = Trim(finalSQL)
    Exit Function

ErrorHandler:
    RemoveCommentsFromSQL = SQL

End Function

' ============================================
' JOIN PARSER (VEREINFACHT)
' ============================================

Public Function ParseJoins(SQL As String) As Collection
    Dim col As New Collection
    Dim workSQL As String
    Dim pos As Long
    Dim j As clsParsedJoin
    Dim joinCount As Long

    ' SQL normalisieren: Zeilenumbrueche durch Leerzeichen ersetzen
    workSQL = Replace(SQL, vbCrLf, " ")
    workSQL = Replace(workSQL, vbLf, " ")
    workSQL = Replace(workSQL, vbCr, " ")

    ' Mehrfache Leerzeichen entfernen
    While InStr(workSQL, "  ") > 0
        workSQL = Replace(workSQL, "  ", " ")
    Wend

    workSQL = LCase(workSQL)
    pos = InStr(workSQL, " join ")
    joinCount = 0

    Do While pos > 0 And joinCount < 50 ' Max 50 JOINs zur Sicherheit
        Set j = New clsParsedJoin

        j.JoinType = DetectJoinType(workSQL, pos)
        j.LeftTable = DetectLeftTable(workSQL, pos)
        j.RightTable = DetectRightTable(workSQL, pos)

        col.Add j
        joinCount = joinCount + 1

        pos = InStr(pos + 5, workSQL, " join ")
    Loop

    Set ParseJoins = col
End Function

Private Function DetectJoinType(SQL As String, pos As Long) As String
    Dim s As String
    Dim startPos As Long

    startPos = IIf(pos - 20 < 1, 1, pos - 20)
    s = Mid(SQL, startPos, 20)

    If InStr(s, "inner join") > 0 Then
        DetectJoinType = "INNER JOIN"
    ElseIf InStr(s, "left join") > 0 Then
        DetectJoinType = "LEFT JOIN"
    ElseIf InStr(s, "right join") > 0 Then
        DetectJoinType = "RIGHT JOIN"
    Else
        DetectJoinType = "JOIN"
    End If
End Function

Private Function DetectLeftTable(SQL As String, pos As Long) As String
    Dim s As String
    Dim w() As String
    Dim i As Long

    s = Trim(Left(SQL, pos - 1))
    w = Split(s, " ")

    For i = UBound(w) To 0 Step -1
        If Len(Trim(w(i))) > 1 Then
            DetectLeftTable = Trim(w(i))
            Exit Function
        End If
    Next i

    DetectLeftTable = ""
End Function

Private Function DetectRightTable(SQL As String, joinPos As Long) As String
    Dim s As String
    Dim w() As String
    Dim onPos As Long

    s = Mid(SQL, joinPos)
    onPos = InStr(s, " on ")

    If onPos > 0 Then
        s = Left(s, onPos - 1)
    End If

    w = Split(s, " ")

    If UBound(w) >= 1 Then
        DetectRightTable = Trim(w(UBound(w)))
    Else
        DetectRightTable = ""
    End If
End Function

' ============================================
' SQL HIGHLIGHTING
' ============================================

Public Function HighlightSQLPlain(SQL As String) As String
    Dim txt As String
    txt = " " & SQL & " "

    txt = Replace(txt, " select ", " SELECT ")
    txt = Replace(txt, " from ", " FROM ")
    txt = Replace(txt, " where ", " WHERE ")
    txt = Replace(txt, " join ", " JOIN ")
    txt = Replace(txt, " inner join ", " INNER JOIN ")
    txt = Replace(txt, " left join ", " LEFT JOIN ")
    txt = Replace(txt, " right join ", " RIGHT JOIN ")
    txt = Replace(txt, " on ", " ON ")
    txt = Replace(txt, " group by ", " GROUP BY ")
    txt = Replace(txt, " order by ", " ORDER BY ")

    HighlightSQLPlain = Trim(txt)
End Function

' ============================================
' LOG-AUSGABE (OHNE UMLAUTE)
' ============================================

Public Sub AddLogEntry(LogList As ListBox, message As String)

    On Error Resume Next
    LogList.AddItem message

End Sub

' ============================================
' HAUPTFUNKTION: SQL IMPORT
' ============================================

Public Sub ImportSQLFromFile_GUI(SQLFilePath As String, _
                                 LogList As ListBox, _
                                 chkTables As CheckBox, _
                                 chkQueries As CheckBox)

    On Error GoTo ErrorHandler

    Dim doTables As Boolean
    Dim doQueries As Boolean
    Dim raw As String
    Dim cleanSQL As String

    doTables = chkTables.Value <> 0
    doQueries = chkQueries.Value <> 0

    LogList.RowSource = ""

    AddLogEntry LogList, "===================================================="
    AddLogEntry LogList, "Starte SQL-Import"
    AddLogEntry LogList, "Datei: " & SQLFilePath
    AddLogEntry LogList, "===================================================="

    ' Datei pruefen
    If Dir(SQLFilePath) = "" Then
        Err.Raise 1001, "ImportSQLFromFile_GUI", "Datei nicht gefunden"
    End If

    AddLogEntry LogList, "OK: Datei gefunden"

    ' Datei lesen
    raw = ReadTextFileRobust(SQLFilePath)

    If Len(raw) = 0 Then
        Err.Raise 1002, "ImportSQLFromFile_GUI", "Datei ist leer"
    End If

    AddLogEntry LogList, "OK: Datei gelesen (" & Len(raw) & " Zeichen)"

    ' Kommentare entfernen
    cleanSQL = RemoveCommentsFromSQL(raw)

    AddLogEntry LogList, "OK: Kommentare entfernt"
    AddLogEntry LogList, "OK: Bereinigtes SQL (" & Len(cleanSQL) & " Zeichen)"

    ' SQL-Preview (erste 150 Zeichen)
    AddLogEntry LogList, ""
    AddLogEntry LogList, "SQL-Preview:"
    AddLogEntry LogList, Left(cleanSQL, 150) & "..."
    AddLogEntry LogList, ""

    ' JOINs analysieren
    Dim joins As Collection
    Set joins = ParseJoins(cleanSQL)

    AddLogEntry LogList, "OK: JOINs gefunden: " & joins.Count

    Dim i As Long
    Dim j As clsParsedJoin
    For i = 1 To joins.Count
        Set j = joins.Item(i)
        AddLogEntry LogList, "   -> " & j.JoinType & ": " & j.LeftTable & " zu " & j.RightTable
    Next i

    ' Abfrage erstellen
    If doQueries Then
        AddLogEntry LogList, ""
        AddLogEntry LogList, "Erstelle Abfrage..."
        CreateAccessQuery cleanSQL, LogList
    End If

    AddLogEntry LogList, ""
    AddLogEntry LogList, "===================================================="
    AddLogEntry LogList, "OK: Import erfolgreich abgeschlossen"
    AddLogEntry LogList, "===================================================="

    Exit Sub

ErrorHandler:
    AddLogEntry LogList, ""
    AddLogEntry LogList, "FEHLER: " & Err.Description
    AddLogEntry LogList, "Fehler-Nummer: " & Err.Number
    MsgBox "Fehler beim Import: " & Err.Description, vbCritical

End Sub

' ============================================
' ABFRAGE ERSTELLEN
' ============================================

Public Sub CreateAccessQuery(SQL As String, LogList As ListBox)

    On Error GoTo ErrHandler

    Dim qdf As DAO.QueryDef
    Dim name As String
    Dim db As DAO.Database
    Dim cleanSQL As String

    Set db = CurrentDb()

    cleanSQL = Trim(SQL)

    ' Pruefung: SQL muss mit SELECT, INSERT, UPDATE, DELETE beginnen
    If InStr(UCase(cleanSQL), "SELECT") = 0 And _
       InStr(UCase(cleanSQL), "INSERT") = 0 And _
       InStr(UCase(cleanSQL), "UPDATE") = 0 And _
       InStr(UCase(cleanSQL), "DELETE") = 0 Then
        Err.Raise 1003, "CreateAccessQuery", "SQL muss mit SELECT, INSERT, UPDATE oder DELETE beginnen"
    End If

    name = "qry_SQL_" & Format(Now(), "yyyymmdd_hhnnss")

    AddLogEntry LogList, "Erstelle Abfrage: " & name
    AddLogEntry LogList, "SQL-Laenge: " & Len(cleanSQL) & " Zeichen"

    ' Alte Abfrage loeschen falls existiert
    On Error Resume Next
    db.QueryDefs.Delete name
    On Error GoTo ErrHandler

    ' Neue Abfrage erstellen
    Set qdf = db.CreateQueryDef(name, cleanSQL)

    AddLogEntry LogList, "OK: Abfrage erfolgreich erstellt!"
    AddLogEntry LogList, "   Name: " & name
    AddLogEntry LogList, "   Oeffnen Sie die Abfrage in der Entwurfsansicht:"
    AddLogEntry LogList, "   - Im Navigationsbereich (F11)"
    AddLogEntry LogList, "   - Unter 'Abfragen'"
    AddLogEntry LogList, "   - Rechtsklick -> Entwurfsansicht"

    Exit Sub

ErrHandler:
    AddLogEntry LogList, "FEHLER beim Erstellen der Abfrage: " & Err.Description
    AddLogEntry LogList, "Fehler-Nummer: " & Err.Number
    AddLogEntry LogList, "SQL-Anfang: " & Left(SQL, 100) & "..."
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub
