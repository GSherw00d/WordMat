Attribute VB_Name = "WMTest"
' word_math test mode: automated test entry point for the .docx test documents, plus the switch that stops
' WordMat's interactive UI from blocking an automated run.
'
' Called by C:\Dev\word_math\tests\run_docx_tests.ps1 (via tests\lib\run_docx_worker.ps1) from a private,
' hidden Word:
'     Application.Run "WMTest_RunDocument", docPath, resultPath
' Test document format and Notes keywords: C:\Dev\word_math\tests\cases\README.md.
'
' While WMTestMode is True:
'   - MsgBox and InputBox (shadowed below for the whole project), MsgBox2, the UserFormError paths, the sign
'     questions (CMaxima.AnswerQuestion, CMaxima.RunMaximaFile), the "choose CAS" form after a failed solve,
'     the variable picker of MaximaSolvePar and SaveBackup show nothing; each message is recorded with
'     WMTestRecord and appears in the result file.
'   - MsgBox returns OK for OK/OKCancel, No for Yes/No (so e.g. "try GeoGebra?" is declined), Cancel for
'     Retry/Cancel and Abort for Abort/Retry/Ignore. InputBox returns its default value.
'   - The free-text module's own silent flag (WMFreeTextSilent) is switched on for the run, and its log is
'     added to each case's messages.
' With WMTestMode False (normal use) everything passes straight through to the original behaviour.
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

' Shadows VBA.MsgBox for every unqualified MsgBox call in this project. Outside test mode it calls VBA.MsgBox
' with the same arguments.
Public Function MsgBox(ByVal Prompt As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As Variant, Optional ByVal HelpFile As Variant, _
                       Optional ByVal Context As Variant) As VbMsgBoxResult
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
Public Function InputBox(ByVal Prompt As Variant, Optional ByVal Title As Variant, Optional ByVal Default As Variant, _
                         Optional ByVal XPos As Variant, Optional ByVal YPos As Variant, _
                         Optional ByVal HelpFile As Variant, Optional ByVal Context As Variant) As String
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
' Entry points
' ---------------------------------------------------------------------------------------------------------
Public Function WMTest_Ping() As String
    WMTest_Ping = "WMTest OK"
End Function

Public Function WMTest_RunDocument(ByVal docPath As String, ByVal resultPath As String) As String
' Runs every row of the first ID | Input | Expected | Notes table in docPath (which must be a copy: it is
' changed and saved) and writes one tab-separated line per case to resultPath (UTF-8).
' Returns "OK <n> cases" or "ERROR <description>".
    Dim doc As Document, tbl As Table, r As Long, nCases As Long
    Dim cId As Long, cIn As Long, cExp As Long, cNotes As Long, c As Long, h As String
    Dim sv As Variant, found As Table, prevSilent As Boolean

    On Error GoTo fejl
    WMTestMode = True
    prevSilent = WMFreeTextSilent
    WMFreeTextSilent = True       ' the free-text module's own silent flag (messages go to WMFreeTextLog)
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
    doc.Activate

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
    WMFreeTextSilent = prevSilent
    WMTestMode = False
End Function

Public Function WMTest_ReadCompare(ByVal docPath As String, ByVal resultPath As String) As String
' word_math safe read: reads every OMath of docPath (a copy; it is closed without saving) with the new reader
' (WMSR_ReadOMath) and then the old one (WMSR_ReadOMathUndo), and writes one line per equation:
' index, start, in table, same, new, old, ms new, ms old. Returns "OK <n> equations <same> same <ms new> <ms old>".
    Dim doc As Document, i As Long, n As Long, nSame As Long, om As OMath
    Dim a As String, b As String, t As Single, tn As Double, tOld As Double, dn As Double, dOld As Double
    On Error GoTo fejl
    WMTestMode = True
    mResult = "Index" & vbTab & "Start" & vbTab & "InTable" & vbTab & "Same" & vbTab & "New" & vbTab & "Old" & _
              vbTab & "MsNew" & vbTab & "MsOld" & vbCrLf
    WriteResult resultPath
    Set doc = Documents.Open(FileName:=docPath, AddToRecentFiles:=False, Visible:=True)
    doc.Activate
    n = doc.Range.OMaths.Count
    For i = 1 To n
        Set om = doc.Range.OMaths(i)
        t = Timer: a = WMSR_ReadOMath(om.Range): dn = (Timer - t) * 1000
        Set om = doc.Range.OMaths(i)
        t = Timer: b = WMSR_ReadOMathUndo(om.Range): dOld = (Timer - t) * 1000
        tn = tn + dn: tOld = tOld + dOld
        If a = b Then nSame = nSame + 1
        mResult = mResult & i & vbTab & om.Range.Start & vbTab & om.Range.Information(wdWithInTable) & vbTab & _
                  (a = b) & vbTab & Esc(a) & vbTab & Esc(b) & vbTab & Format$(dn, "0") & vbTab & Format$(dOld, "0") & vbCrLf
        WriteResult resultPath
    Next
    ' the several-OMaths path (ReadSelection / ReadMathBoxes): all OMaths of the document at once
    a = ""
    For i = 1 To n
        a = a & IIf(i > 1, " | ", "") & Replace(WMSR_ReadOMath(doc.Range.OMaths(i).Range), ChrW$(11), "")
    Next
    b = WMSR_ReadOMathsUndo(doc.Content)
    If a = b Then nSame = nSame + 1
    mResult = mResult & "all" & vbTab & 0 & vbTab & "-" & vbTab & (a = b) & vbTab & Esc(a) & vbTab & Esc(b) & vbTab & "-" & vbTab & "-" & vbCrLf
    WriteResult resultPath
    WMTest_ReadCompare = "OK " & n + 1 & " readings " & nSame & " same " & Format$(tn, "0") & " " & Format$(tOld, "0")
    GoTo slut
fejl:
    WMTest_ReadCompare = "ERROR " & Err.Number & ": " & Err.Description & " (equation " & i & ")"
slut:
    On Error Resume Next
    WMSR_CloseScratch
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    WMTestMode = False
End Function

Public Function WMTest_ReadNonDestructive(ByVal docPath As String, ByVal resultPath As String, _
                                          ByVal reader As String) As String
' word_math safe read: types a marker (the known top of the undo stack), reads every OMath of docPath (a copy,
' closed without saving) with reader "new" (the production omax.ReadEquation2) or "old" (WMSR_ReadOMathUndo),
' and checks after each read: full text, Saved, selection; after all: WordOpenXML, then a single Undo must
' remove the marker; finally the same reads with Saved = True. One line per check: name, PASS/FAIL, detail.
    Dim doc As Document, i As Long, n As Long, t0 As String, t1 As String, x1 As String, s1 As Long, e1 As Long
    Dim sv1 As Boolean, bad As String, nFail As Long, x2 As String
    On Error GoTo fejl
    WMTestMode = True
    If omax Is Nothing Then Set omax = New CMaxima
    mResult = "Check" & vbTab & "Result" & vbTab & "Detail" & vbCrLf
    WriteResult resultPath
    Set doc = Documents.Open(FileName:=docPath, AddToRecentFiles:=False, Visible:=True)
    doc.Activate
    t0 = FullText(doc)
    Selection.HomeKey wdStory
    Selection.TypeText "WMSR-MARK "
    t1 = FullText(doc): x1 = NoRsid(doc.Content.WordOpenXML): sv1 = doc.Saved
    s1 = Selection.Start: e1 = Selection.End
    n = doc.Range.OMaths.Count
    For i = 1 To n
        ReadOne doc.Range.OMaths(i).Range, reader
        bad = ""
        If FullText(doc) <> t1 Then bad = bad & "text changed; "
        If doc.Saved <> sv1 Then bad = bad & "Saved changed; "
        If Selection.Start <> s1 Or Selection.End <> e1 Then bad = bad & "selection moved; "
        If Not ActiveDocument Is doc Then bad = bad & "active document changed; "
        NdRow "equation " & i & " (start " & doc.Range.OMaths(i).Range.Start & ")", bad, nFail
    Next
    x2 = NoRsid(doc.Content.WordOpenXML)
    If x2 <> x1 Then WriteText resultPath & ".before.xml", x1: WriteText resultPath & ".after.xml", x2
    NdRow "formatting (WordOpenXML without rsids) after " & n & " reads", IIf(x2 = x1, "", "WordOpenXML changed (see .before.xml/.after.xml)"), nFail
    doc.Undo 1
    NdRow "one Undo removes the typed marker", IIf(FullText(doc) = t0, "", "after one Undo the text is not the original: " & _
          Left$(Clean(FullText(doc)), 80)), nFail
    doc.Saved = True
    For i = 1 To n
        ReadOne doc.Range.OMaths(i).Range, reader
    Next
    NdRow "Saved = True stays True after " & n & " reads", IIf(doc.Saved, "", "Saved became False"), nFail
    WMTest_ReadNonDestructive = "OK " & n & " equations, " & nFail & " failed checks"
    GoTo slut
fejl:
    WMTest_ReadNonDestructive = "ERROR " & Err.Number & ": " & Err.Description & " (equation " & i & ")"
    NdRow "VBA error", WMTest_ReadNonDestructive, nFail
slut:
    On Error Resume Next
    WriteResult resultPath
    WMSR_CloseScratch
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    WMTestMode = False
End Function

Public Function WMTest_ReadNoClose(ByVal docPath As String) As String
' word_math safe read: reads the first OMath of docPath (closed without saving) through omax.ReadEquation2 and
' leaves the scratch document to WMSafeRead's own idle-time close. Returns "<text>|<scratch open?>".
    Dim doc As Document, t As String
    On Error GoTo fejl
    If omax Is Nothing Then Set omax = New CMaxima
    Set doc = Documents.Open(FileName:=docPath, AddToRecentFiles:=False, Visible:=True)
    doc.Activate
    t = omax.ReadEquation2(doc.Range.OMaths(1).Range)
    WMTest_ReadNoClose = t & "|" & WMSR_ScratchIsOpen()
    doc.Close wdDoNotSaveChanges
    Exit Function
fejl:
    WMTest_ReadNoClose = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function WMTest_ReadTiming(ByVal docPath As String, ByVal rounds As Long) As String
' word_math safe read: times CMaxima.FindDefinitions (which reads every earlier definition) with the cursor at
' the end of docPath (closed without saving), alternating the old and the new reader, rounds times each.
' Returns "defs=<n> new=<ms per call> old=<ms per call>" (the first call of each kind is included).
    Dim doc As Document, k As Long, t As Single, tn As Double, tOld As Double, d As String
    On Error GoTo fejl
    WMTestMode = True
    If omax Is Nothing Then Set omax = New CMaxima
    Set doc = Documents.Open(FileName:=docPath, AddToRecentFiles:=False, Visible:=True)
    doc.Activate
    Selection.EndKey wdStory
    For k = 1 To rounds
        WMSR_TestOldReader = True
        t = Timer: d = omax.FindDefinitions: tOld = tOld + (Timer - t) * 1000
        WMSR_TestOldReader = False
        WMSR_CloseScratch                   ' each new-reader call pays for creating the scratch, as in real use
        t = Timer: d = omax.FindDefinitions: tn = tn + (Timer - t) * 1000
    Next
    WMTest_ReadTiming = "defs=" & omax.defindex & " new=" & Format$(tn / rounds, "0") & " old=" & Format$(tOld / rounds, "0")
    GoTo slut
fejl:
    WMTest_ReadTiming = "ERROR " & Err.Number & ": " & Err.Description
slut:
    On Error Resume Next
    WMSR_TestOldReader = False
    WMSR_CloseScratch
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    WMTestMode = False
End Function

Private Sub ReadOne(r As Range, ByVal reader As String)
    If reader = "old" Then WMSR_ReadOMathUndo r Else omax.ReadEquation2 r
End Sub

Private Function FullText(doc As Document) As String
    Dim r As Range
    Set r = doc.Content
    r.TextRetrievalMode.IncludeHiddenText = True
    FullText = r.Text
End Function

Private Sub NdRow(ByVal check As String, ByVal bad As String, nFail As Long)
    If Len(bad) > 0 Then nFail = nFail + 1
    mResult = mResult & check & vbTab & IIf(Len(bad) = 0, "PASS", "FAIL") & vbTab & Clean(bad) & vbCrLf
End Sub

Private Function Esc(ByVal s As String) As String
' Makes control characters visible for the TSV: \r \n \t \v (ChrW 11) and other codes below 32 as \u....
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        Select Case c
        Case 13: o = o & "\r"
        Case 10: o = o & "\n"
        Case 9: o = o & "\t"
        Case 11: o = o & "\v"
        Case 92: o = o & "\\"
        Case 0 To 31: o = o & "\u" & Right$("000" & Hex$(c), 4)
        Case Else: o = o & Mid$(s, i, 1)
        End Select
    Next
    Esc = o
End Function

Private Sub RunCase(doc As Document, tbl As Table, ByVal r As Long, ByVal cId As Long, ByVal cIn As Long, _
                    ByVal cExp As Long, ByVal cNotes As Long)
    Dim id As String, notes As String, action As String, var As String, inputText As String
    Dim expected As String, actual As String, output As String, t As Single
    Dim cellR As Range, om As OMath, i As Long, s As String, first As String, isFree As Boolean
    Dim rg As Range, p As Long

    On Error GoTo fejl
    mMessages = ""
    WMFreeTextLog = ""
    t = Timer
    doc.Activate
    id = CellText(tbl.Cell(r, cId))
    If cNotes > 0 Then notes = CellText(tbl.Cell(r, cNotes))
    expected = RangeLinear(CellRange(tbl.Cell(r, cExp)))
    Set cellR = CellRange(tbl.Cell(r, cIn))
    inputText = RangeLinear(cellR)
    isFree = (cellR.OMaths.Count = 0 And InStr(inputText, "~~") > 0)
    action = LCase$(NoteValue(notes, "action"))
    var = NoteValue(notes, "var")
    If action = "" Then action = "evaluate"

    ApplyCaseSettings notes

    If isFree Then
        ' free-typed ~~...~~: the cursor is put just after the first ~~ and the normal ribbon action runs
        Set rg = cellR.Duplicate
        rg.TextRetrievalMode.IncludeHiddenText = True
        With rg.Find
            .ClearFormatting
            .text = "~~"
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        If Not rg.Find.Execute Then
            WMTestRecord "WMTest", "no ~~ marker found in the Input cell"
        Else
            Select Case action
            Case "evaluate"
                doc.Range(rg.End, rg.End).Select
                beregn
            Case "solve"
                doc.Range(rg.End, rg.End).Select
                MaximaSolvePar var
            Case "define", "none"
                ' nothing to run: the row only supplies a definition for later rows
            Case Else
                WMTestRecord "WMTest", "unknown action '" & action & "' (use evaluate, solve var=x or define)"
            End Select
        End If
        s = RangeLinear(CellRange(tbl.Cell(r, cIn)))
        i = InStr(s, "~~")
        p = InStrRev(s, "~~")
        If i > 0 And p > i Then actual = Trim$(Mid$(s, i + 2, p - i - 2)) Else actual = s
    ElseIf cellR.OMaths.Count = 0 Then
        WMTestRecord "WMTest", "the Input cell holds neither a Word equation nor a ~~...~~ equation"
    Else
        Select Case action
        Case "evaluate", "solve"
            If action = "solve" And var = "" Then
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
                    If Left$(s, Len(first)) = first Then
                        s = Mid$(s, Len(first) + 1)
                    Else
                        p = LastSignPos(s)
                        If p > 0 Then s = Mid$(s, p)
                    End If
                    s = Trim$(s)
                    If Left$(s, 1) = "=" Or Left$(s, 1) = ChrW$(8776) Then s = Trim$(Mid$(s, 2))
                    actual = s
                ElseIf cellR.OMaths.Count > 1 Then
                    actual = RangeLinear(cellR.OMaths(cellR.OMaths.Count).Range)
                End If
            End If
        Case "define", "none"
            ' nothing to run: the equation only supplies a definition for later rows; Actual = the equation itself
            Set cellR = CellRange(tbl.Cell(r, cIn))
            If cellR.OMaths.Count > 0 Then actual = RangeLinear(cellR.OMaths(1).Range)
        Case Else
            WMTestRecord "WMTest", "unknown action '" & action & "' for a Word equation (use evaluate, solve var=x or define)"
        End Select
    End If

    Set cellR = CellRange(tbl.Cell(r, cIn))
    If isFree Or cellR.OMaths.Count = 0 Then
        output = RangeLinear(cellR)
    Else
        For i = 1 To cellR.OMaths.Count
            If i > 1 Then output = output & " | "
            output = output & RangeLinear(cellR.OMaths(i).Range)
        Next
    End If
    GoTo slut
fejl:
    WMTestRecord "WMTest", "VBA error " & Err.Number & ": " & Err.Description
    Resume slut
slut:
    On Error Resume Next
    WMTestMode = True             ' in case something switched it off
    WMFreeTextSilent = True
    If Len(WMFreeTextLog) > 0 Then WMTestRecord "FreeText", WMFreeTextLog
    mResult = mResult & Clean(id) & vbTab & Clean(action) & vbTab & Clean(inputText) & vbTab & Clean(expected) & _
              vbTab & Clean(notes) & vbTab & Clean(actual) & vbTab & Clean(output) & vbTab & Clean(mMessages) & _
              vbTab & CLng((Timer - t) * 1000) & vbCrLf
End Sub

Private Function LastSignPos(ByVal s As String) As Long
' Position of the last top-level "=" or approx sign (outside brackets), 0 if none.
    Dim i As Long, d As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "(" Or ch = "[" Or ch = "{" Then d = d + 1
        If ch = ")" Or ch = "]" Or ch = "}" Then d = d - 1
        If d = 0 And (ch = "=" Or ch = ChrW$(8776)) Then LastSignPos = i
    Next
End Function

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
' Per case, from Notes: units=on|off (default off), outunits=kN (default none), exact=auto|exact|num
' (default auto), digits=n (default: the user's setting)
    Dim v As String
    v = LCase$(NoteValue(notes, "units"))
    If MaximaUnits <> (v = "on" Or v = "true" Or v = "1") Then MaximaUnits = (v = "on" Or v = "true" Or v = "1")
    OutUnits = NoteValue(notes, "outunits")
    v = LCase$(NoteValue(notes, "exact"))
    If v = "exact" Then
        MaximaExact = 1
    ElseIf v = "num" Then
        MaximaExact = 2
    Else
        MaximaExact = 0
    End If
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
' Copies the range into a hidden scratch document, linearises any equations there and returns the text
' (hidden text included). The source range is never modified.
    Dim s As String, om As OMath, sr As Range
    If mScratch Is Nothing Then Set mScratch = Documents.Add(Visible:=False)
    mScratch.Content.Delete
    mScratch.Range(0, 0).FormattedText = rg.FormattedText
    For Each om In mScratch.OMaths
        om.Linearize
    Next
    Set sr = mScratch.Content
    sr.TextRetrievalMode.IncludeHiddenText = True
    s = sr.text
    s = Replace(s, Chr$(7), "")
    Do While Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf
        s = Left$(s, Len(s) - 1)
    Loop
    RangeLinear = Trim$(Replace(Replace(s, vbCr, " "), vbLf, " "))
End Function

Private Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Function NoRsid(ByVal x As String) As String
' WordOpenXML without rsid attributes and <w:rsid> entries: Word assigns new revision-save IDs on export.
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = " w:rsid\w*=""[0-9A-F]+"""
    x = re.Replace(x, "")
    re.Pattern = "<w:rsid w:val=""[0-9A-F]+""/>"
    x = re.Replace(x, "")
    re.Pattern = "<w:rsidRoot w:val=""[0-9A-F]+""/>"
    NoRsid = re.Replace(x, "")
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

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
