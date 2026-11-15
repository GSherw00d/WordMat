Attribute VB_Name = "WMTest"
' word_math test mode: automated test entry point for the .docx test documents, plus the switch that stops
' WordMat's interactive UI from blocking an automated run.
'
' Called by C:\Dev\word_math\tests\run_docx_tests.ps1 from a private, hidden Word:
'     Application.Run "WMTest_RunDocument", docPath, resultPath
' Test document format and Notes keywords: C:\Dev\word_math\tests\cases\README.md.
'
' While WMTestMode is True:
'   - MsgBox and InputBox (shadowed below for the whole project), MsgBox2, the UserFormError paths,
'     the sign question in CMaxima.AnswerQuestion and SaveBackup do not show anything; each message is
'     recorded with WMTestRecord and appears in the result file.
'   - MsgBox returns OK for OK/OKCancel and No for Yes/No questions (so e.g. "try GeoGebra?" is declined).
'   - InputBox returns its default value.
' With WMTestMode False (normal use) all of this passes straight through to the original behaviour.
Option Explicit

Public WMTestMode As Boolean            ' word_math test mode

Private mMessages As String
Private mScratch As Document
Private mResult As String

' ---------------------------------------------------------------------------------------------------------
' Message capture
' ---------------------------------------------------------------------------------------------------------
Public Sub WMTestRecord(ByVal Source As String, ByVal text As String)
    text = Replace(Replace(Replace(text, vbCrLf, " / "), vbCr, " / "), vbLf, " / ")
    If Len(mMessages) > 0 Then mMessages = mMessages & " || "
    mMessages = mMessages & Source & ": " & Trim$(text)
End Sub

Public Function WMTestDefaultAnswer(ByVal Buttons As Long) As VbMsgBoxResult
    Select Case Buttons And 7
        Case vbOKCancel: WMTestDefaultAnswer = vbOK
        Case vbAbortRetryIgnore: WMTestDefaultAnswer = vbAbort
        Case vbYesNoCancel: WMTestDefaultAnswer = vbNo
        Case vbYesNo: WMTestDefaultAnswer = vbNo
        Case vbRetryCancel: WMTestDefaultAnswer = vbCancel
        Case Else: WMTestDefaultAnswer = vbOK
    End Select
End Function

' Shadows VBA.MsgBox for every unqualified MsgBox call in this project. Outside test mode it simply calls
' VBA.MsgBox with the same arguments.
Public Function MsgBox(Prompt As Variant, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As Variant, _
                       Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult
    If WMTestMode Then
        WMTestRecord "MsgBox", CStr(Prompt)
        MsgBox = WMTestDefaultAnswer(Buttons)
    ElseIf Not IsMissing(HelpFile) Then
        MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    ElseIf Not IsMissing(Title) Then
        MsgBox = VBA.MsgBox(Prompt, Buttons, Title)
    Else
        MsgBox = VBA.MsgBox(Prompt, Buttons)
    End If
End Function

' Shadows VBA.InputBox in the same way. In test mode it returns the default value.
Public Function InputBox(Prompt As Variant, Optional Title As Variant, Optional Default As Variant, _
                         Optional XPos As Variant, Optional YPos As Variant, Optional HelpFile As Variant, _
                         Optional Context As Variant) As String
    If WMTestMode Then
        WMTestRecord "InputBox", CStr(Prompt)
        If Not IsMissing(Default) Then InputBox = CStr(Default)
    ElseIf Not (IsMissing(XPos) And IsMissing(YPos) And IsMissing(HelpFile)) Then
        InputBox = VBA.InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
    ElseIf Not IsMissing(Default) Then
        InputBox = VBA.InputBox(Prompt, Title, Default)
    ElseIf Not IsMissing(Title) Then
        InputBox = VBA.InputBox(Prompt, Title)
    Else
        InputBox = VBA.InputBox(Prompt)
    End If
End Function

' ---------------------------------------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------------------------------------
Public Function WMTest_RunDocument(ByVal docPath As String, ByVal resultPath As String) As String
' Runs every row of the first ID | Input | Expected | Notes table in docPath (which must be a copy: it is
' changed and saved) and writes one tab-separated line per case to resultPath (UTF-8).
' Returns "OK <n> cases" or "ERROR <description>".
    Dim doc As Document, tbl As Table, r As Long, nCases As Long
    Dim cId As Long, cIn As Long, cExp As Long, cNotes As Long, c As Long, h As String
    Dim sv As Variant, found As Table

    On Error GoTo fejl
    WMTestMode = True
    RunFirst                      ' loads WordMat's settings, so the snapshot below sees the real values
    sv = SaveTestSettings()
    CASengine = 0                 ' Maxima
    CASengineTempOnly = 0
    DebugWM = False

    mResult = "ID" & vbTab & "Action" & vbTab & "Input" & vbTab & "Expected" & vbTab & "Notes" & vbTab & _
              "Actual" & vbTab & "Output" & vbTab & "Messages" & vbTab & "ElapsedMs" & vbCrLf
    WriteResult resultPath

    Set doc = Documents.Open(FileName:=docPath, AddToRecentFiles:=False, Visible:=True)
    doc.Activate
    Set mScratch = Documents.Add(Visible:=False)

    For Each tbl In doc.Tables
        cId = 0: cIn = 0: cExp = 0: cNotes = 0
        For c = 1 To tbl.Columns.Count
            h = LCase$(CellText(tbl.Cell(1, c)))
            If h = "id" Then cId = c
            If h = "input" Then cIn = c
            If h = "expected" Then cExp = c
            If h = "notes" Then cNotes = c
        Next
        If cId > 0 And cIn > 0 And cExp > 0 Then Set found = tbl: Exit For
    Next
    Set tbl = found
    If tbl Is Nothing Then Err.Raise vbObjectError + 1, , "no table with ID | Input | Expected | Notes header"

    For r = 2 To tbl.Rows.Count
        If Len(CellText(tbl.Cell(r, cId))) > 0 Then
            RunCase doc, tbl, r, cId, cIn, cExp, cNotes
            nCases = nCases + 1
            WriteResult resultPath
        End If
    Next

    WMTest_RunDocument = "OK " & nCases & " cases"
    GoTo slut
fejl:
    WMTest_RunDocument = "ERROR " & Err.Number & ": " & Err.Description
slut:
    On Error Resume Next
    RestoreTestSettings sv
    If Not mScratch Is Nothing Then mScratch.Close wdDoNotSaveChanges
    Set mScratch = Nothing
    If Not doc Is Nothing Then
        doc.Save
        doc.Close wdDoNotSaveChanges
    End If
    WMTestMode = False
End Function

Private Sub RunCase(doc As Document, tbl As Table, ByVal r As Long, ByVal cId As Long, ByVal cIn As Long, _
                    ByVal cExp As Long, ByVal cNotes As Long)
    Dim id As String, notes As String, action As String, var As String, inputText As String
    Dim expected As String, actual As String, output As String, t As Single
    Dim cellR As Range, om As OMath, i As Long, s As String, first As String

    On Error GoTo fejl
    mMessages = ""
    t = Timer
    id = CellText(tbl.Cell(r, cId))
    If cNotes > 0 Then notes = CellText(tbl.Cell(r, cNotes))
    expected = RangeLinear(CellRange(tbl.Cell(r, cExp)))
    Set cellR = CellRange(tbl.Cell(r, cIn))
    inputText = RangeLinear(cellR)
    action = LCase$(NoteValue(notes, "action"))
    var = NoteValue(notes, "var")

    If action = "" Then
        If InStr(inputText, "~~") > 0 Then
            action = "freetext"
        ElseIf Right$(Replace(inputText, " ", ""), 2) = "=?" Then
            action = "agreed"
        End If
    End If

    ApplyCaseSettings notes

    Select Case action
    Case "evaluate", "solve"
        If cellR.OMaths.Count = 0 Then
            WMTestRecord "WMTest", "action=" & action & " needs a Word equation in the Input cell"
        ElseIf action = "solve" And var = "" Then
            WMTestRecord "WMTest", "action=solve needs var=<name> in Notes (the variable picker cannot be shown)"
        Else
            Set om = cellR.OMaths(1)
            first = RangeLinear(om.Range)
            om.Range.Select
            If action = "evaluate" Then
                beregn
            Else
                MaximaSolvePar var
            End If
            Set cellR = CellRange(tbl.Cell(r, cIn))
            If action = "evaluate" Then
                s = RangeLinear(cellR.OMaths(1).Range)
                If Left$(s, Len(first)) = first Then s = Mid$(s, Len(first) + 1)
                s = Trim$(s)
                If Left$(s, 1) = "=" Or Left$(s, 1) = ChrW$(8776) Then s = Trim$(Mid$(s, 2))
                actual = s
            ElseIf cellR.OMaths.Count > 1 Then
                actual = RangeLinear(cellR.OMaths(cellR.OMaths.Count).Range)
            End If
        End If
    Case "freetext"
        WMTestRecord "WMTest", "free-typed ~~...~~ equations are not implemented yet"
    Case "agreed"
        WMTestRecord "WMTest", "the '= ?' syntax is not implemented yet"
    Case Else
        WMTestRecord "WMTest", "unknown or missing action '" & action & "' (use action=evaluate or action=solve var=x)"
    End Select

    Set cellR = CellRange(tbl.Cell(r, cIn))
    For i = 1 To cellR.OMaths.Count
        If i > 1 Then output = output & " | "
        output = output & RangeLinear(cellR.OMaths(i).Range)
    Next
    If cellR.OMaths.Count = 0 Then output = RangeLinear(cellR)
    GoTo slut
fejl:
    WMTestRecord "WMTest", "VBA error " & Err.Number & ": " & Err.Description
    Resume slut
slut:
    On Error Resume Next
    WMTestMode = True             ' in case something switched it off
    If action = "" Then action = "?"
    mResult = mResult & Clean(id) & vbTab & Clean(action) & vbTab & Clean(inputText) & vbTab & Clean(expected) & _
              vbTab & Clean(notes) & vbTab & Clean(actual) & vbTab & Clean(output) & vbTab & Clean(mMessages) & _
              vbTab & CLng((Timer - t) * 1000) & vbCrLf
End Sub

' ---------------------------------------------------------------------------------------------------------
' Settings: everything a case may change is saved first and restored at the end (WordMat writes these
' properties to HKCU\Software\WordMat).
' ---------------------------------------------------------------------------------------------------------
Private Function SaveTestSettings() As Variant
    SaveTestSettings = Array(MaximaUnits, OutUnits, CASengine, MaximaExact, MaximaCifre)
End Function

Private Sub RestoreTestSettings(sv As Variant)
    On Error Resume Next
    If IsEmpty(sv) Then Exit Sub
    MaximaUnits = sv(0)
    OutUnits = sv(1)
    CASengine = sv(2)
    MaximaExact = sv(3)
    MaximaCifre = sv(4)
    CASengineTempOnly = sv(2)
End Sub

Private Sub ApplyCaseSettings(ByVal notes As String)
' units=on|off (default off), outunits=km,h (default none), exact=auto|exact|num, digits=n
    Dim v As String
    v = LCase$(NoteValue(notes, "units"))
    MaximaUnits = (v = "on" Or v = "true" Or v = "1")
    OutUnits = NoteValue(notes, "outunits")
    v = LCase$(NoteValue(notes, "exact"))
    If v = "auto" Then MaximaExact = 0
    If v = "exact" Then MaximaExact = 1
    If v = "num" Then MaximaExact = 2
    v = NoteValue(notes, "digits")
    If Len(v) > 0 Then MaximaCifre = CInt(v)
End Sub

Private Function NoteValue(ByVal notes As String, ByVal key As String) As String
' Notes look like "action=solve var=a; units=on; xfail: reason". Segments starting with xfail are skipped.
    Dim seg As Variant, tok As Variant, p As Long
    For Each seg In Split(notes, ";")
        If LCase$(Left$(Trim$(seg), 5)) <> "xfail" Then
            For Each tok In Split(Trim$(seg), " ")
                p = InStr(tok, "=")
                If p > 1 Then
                    If LCase$(Trim$(Left$(tok, p - 1))) = LCase$(key) Then
                        NoteValue = Trim$(Mid$(tok, p + 1))
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

' ---------------------------------------------------------------------------------------------------------
' Reading text without touching the test document
' ---------------------------------------------------------------------------------------------------------
Private Function CellRange(c As Cell) As Range
    Dim rg As Range
    Set rg = c.Range
    rg.MoveEnd wdCharacter, -1    ' drop the end-of-cell mark
    Set CellRange = rg
End Function

Private Function CellText(c As Cell) As String
    CellText = Trim$(Replace(Replace(Replace(c.Range.text, Chr$(13) & Chr$(7), ""), Chr$(7), ""), Chr$(13), " "))
End Function

Private Function RangeLinear(rg As Range) As String
' Copies the range into a hidden scratch document, linearises any equations there and returns the text.
' The source range is never modified.
    Dim s As String, om As OMath
    If mScratch Is Nothing Then Set mScratch = Documents.Add(Visible:=False)
    mScratch.Content.Delete
    mScratch.Content.InsertParagraphAfter
    mScratch.Range(0, 0).FormattedText = rg.FormattedText
    For Each om In mScratch.OMaths
        om.Linearize
    Next
    s = mScratch.Content.text
    s = Replace(s, Chr$(7), "")
    Do While Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf
        s = Left$(s, Len(s) - 1)
    Loop
    RangeLinear = Trim$(Replace(Replace(s, vbCr, " "), vbLf, " "))
End Function

Private Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Sub WriteResult(ByVal path As String)
' UTF-8 (with BOM), rewritten after every case so a hang still leaves the finished cases on disk.
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText mResult
    st.SaveToFile path, 2
    st.Close
End Sub
