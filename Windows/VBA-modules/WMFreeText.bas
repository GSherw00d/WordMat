Attribute VB_Name = "WMFreeText"
' word_math free text
' ---------------------------------------------------------------------------------------------------------
' Free-typed equations in ordinary paragraph text, written between ~~ markers (aim 2 of word_math).
'
'   ~~m := 5 [kg]~~                 definition (":=" or U+2254); applies to later equations of both kinds
'   ~~F = m*a = ?~~                 evaluation; the result replaces the "?"
'   ~~F = m*a = ? [kN]~~            evaluation with a requested output unit
'   ~~F = m*a = ? #(3)~~            "#(n)" numbering is kept and excluded from the maths
'   ~~F = m*a~~                     plain equation: nothing happens on evaluate
'
' The text is normalised into the same kind of linear string that CMaxima.ReadSelection produces for an
' equation-editor object, and then the existing pipeline is reused:
'   omax.Kommando -> CMaxima.beregn (CodeForMaxima -> Maxima -> ConvertToWordSymbols).
' Results are written back as plain, editable text (never OMath objects or images).
'
' Hooks into existing WordMat code (all marked "word_math free text"):
'   Maxima.beregn          -> WMFT_HandleEvaluate
'   Maxima.MaximaSolvePar  -> WMFT_HandleSolve
'   CMaxima.FindDefinitions-> WMFT_BeginDefinitions / WMFT_AddDefinitionsAfter
' Design notes: word_math\research\04_freetext_design.md
' ---------------------------------------------------------------------------------------------------------
Option Explicit

' When True, nothing in this module shows UI; messages are appended to WMFreeTextLog instead.
Public WMFreeTextSilent As Boolean
Public WMFreeTextLog As String

Private Const MARKER As String = "~~"
Private Const MARKER_STYLE As String = "WM Marker"

Private Const KIND_PLAIN As Integer = 0
Private Const KIND_DEF As Integer = 1
Private Const KIND_ASK As Integer = 2       ' unanswered evaluation: "... = ?"
Private Const KIND_ANSWERED As Integer = 3  ' evaluation that already holds a result
Private Const KIND_ERROR As Integer = -1

Private Type FTSpan
    OpenStart As Long       ' document position of the opening "~~"
    ContentStart As Long    ' first character after the opening marker
    ContentEnd As Long      ' position of the closing "~~"
    CloseEnd As Long        ' first position after the closing marker
    Content As String       ' text between the markers (hidden text included)
    Unbalanced As Boolean   ' an opening marker without a partner in the same paragraph
End Type

Private Type FTParsed
    Kind As Integer
    ErrMsg As String
    Expr As String          ' maths to evaluate (raw text) or the definition text
    ReplStart As Long       ' 1-based offset in Content of the first character to replace (the = or ~ sign)
    ReplEnd As Long         ' 1-based offset in Content just after the last character to replace
    Spaced As Boolean       ' the user wrote spaces around the answer's = sign
    OutUnit As String       ' requested output unit, text inside [ ] (empty when none)
End Type

' Free-text definitions found by WMFT_BeginDefinitions, in document order
Private mDefPos() As Long
Private mDefText() As String
Private mDefErr() As String
Private mDefCount As Long
Private mDefIdx As Long

' =========================================================================================================
' Public macros (Alt+F8)
' =========================================================================================================

Public Sub WM_EvaluateAllFreeText()
' Evaluates every free-typed evaluation (unanswered and answered) in the document, top to bottom.
    Dim spans() As FTSpan, n As Long, i As Long, j As Long, delta As Long, msg As String
    Dim nOK As Long, nErr As Long, errList As String, selS As Long, selE As Long
    Dim p As FTParsed
    On Error GoTo fejl
    If Documents.Count = 0 Then Exit Sub
    selS = Selection.Start: selE = Selection.End
    n = FindSpans(ActiveDocument.Content.Start, ActiveDocument.Content.End, spans)
    For i = 1 To n
        If spans(i).Unbalanced Then
            nErr = nErr + 1
            errList = errList & Snippet(spans(i)) & ": " & MsgUnbalanced() & vbCrLf
            LogMsg "ERROR " & Snippet(spans(i)) & ": " & MsgUnbalanced()
        Else
            p = ParseSpan(spans(i).Content)
            If p.Kind = KIND_ERROR Then
                nErr = nErr + 1
                errList = errList & Snippet(spans(i)) & ": " & p.ErrMsg & vbCrLf
                LogMsg "ERROR " & Snippet(spans(i)) & ": " & p.ErrMsg
            ElseIf p.Kind = KIND_ASK Or p.Kind = KIND_ANSWERED Then
                delta = 0
                If EvaluateSpan(spans(i), p, delta, msg, True) Then
                    nOK = nOK + 1
                    For j = i + 1 To n
                        ShiftSpan spans(j), delta
                    Next
                    If selS >= spans(i).ContentStart + p.ReplEnd - 1 Then selS = selS + delta
                    If selE >= spans(i).ContentStart + p.ReplEnd - 1 Then selE = selE + delta
                Else
                    nErr = nErr + 1
                    errList = errList & Snippet(spans(i)) & ": " & msg & vbCrLf
                    LogMsg "ERROR " & Snippet(spans(i)) & ": " & msg
                End If
            End If
        End If
    Next
    On Error Resume Next
    ActiveDocument.Range(selS, selE).Select
    Application.ScreenUpdating = True
    LogMsg "DONE evaluated=" & nOK & " errors=" & nErr
    If Not WMFreeTextSilent Then
        If nErr > 0 Then
            MsgBox2 nOK & " free-typed equation(s) evaluated; " & nErr & " could not be evaluated:" & vbCrLf & vbCrLf & errList, vbOKOnly, TT.Error
        ElseIf nOK = 0 Then
            MsgBox2 "No free-typed evaluations (~~ ... = ?~~) were found in this document.", vbOKOnly, "WordMat"
        End If
    End If
    Exit Sub
fejl:
    ShowMsg "Free-typed equations: unexpected error " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub WM_HideMarkers()
' Applies the hidden character style "WM Marker" (created in the active document only) to every ~~ marker.
    Dim spans() As FTSpan, n As Long, i As Long, st As Style, ur As UndoRecord
    On Error GoTo fejl
    If Documents.Count = 0 Then Exit Sub
    Set st = GetMarkerStyle(True)
    n = FindSpans(ActiveDocument.Content.Start, ActiveDocument.Content.End, spans)
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Hide ~~ markers"
    For i = 1 To n
        ActiveDocument.Range(spans(i).OpenStart, spans(i).OpenStart + 2).Style = st
        If Not spans(i).Unbalanced Then ActiveDocument.Range(spans(i).ContentEnd, spans(i).CloseEnd).Style = st
    Next
    ur.EndCustomRecord
    LogMsg "HIDDEN " & n
    Exit Sub
fejl:
    On Error Resume Next
    If Not ur Is Nothing Then If ur.IsRecordingCustomRecord Then ur.EndCustomRecord
    ShowMsg "Hiding the ~~ markers failed: " & Err.Description
End Sub

Public Sub WM_ShowMarkers()
' Removes the "WM Marker" style from every ~~ marker, so the markers are visible again.
    Dim spans() As FTSpan, n As Long, i As Long, ur As UndoRecord
    On Error GoTo fejl
    If Documents.Count = 0 Then Exit Sub
    n = FindSpans(ActiveDocument.Content.Start, ActiveDocument.Content.End, spans)
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Show ~~ markers"
    For i = 1 To n
        UnhideRange ActiveDocument.Range(spans(i).OpenStart, spans(i).OpenStart + 2)
        If Not spans(i).Unbalanced Then UnhideRange ActiveDocument.Range(spans(i).ContentEnd, spans(i).CloseEnd)
    Next
    ur.EndCustomRecord
    LogMsg "SHOWN " & n
    Exit Sub
fejl:
    On Error Resume Next
    If Not ur Is Nothing Then If ur.IsRecordingCustomRecord Then ur.EndCustomRecord
    ShowMsg "Showing the ~~ markers failed: " & Err.Description
End Sub

' =========================================================================================================
' Hooks called from existing WordMat code
' =========================================================================================================

Public Function WMFT_HandleEvaluate() As Boolean
' Called at the top of Maxima.beregn. Returns True if the free-text route handled the request, False to
' let the equation-editor path run exactly as before.
    Dim spans() As FTSpan, n As Long, i As Long, j As Long, selS As Long, selE As Long
    Dim targets() As Long, nt As Long, unb As Boolean, delta As Long, msg As String
    Dim p As FTParsed, r As Range, evaluated As Long, firstErr As String
    On Error GoTo fejl
    If Documents.Count = 0 Then Exit Function
    If Selection.OMaths.Count > 0 Then Exit Function
    selS = Selection.Start: selE = Selection.End
    Set r = Selection.Range
    r.Expand wdParagraph
    n = FindSpans(r.Start, r.End, spans)
    If n = 0 Then Exit Function
    ReDim targets(1 To n)
    For i = 1 To n
        If spans(i).Unbalanced Then
            If selS = selE Then
                If selS >= spans(i).OpenStart Then unb = True
            ElseIf spans(i).OpenStart < selE And spans(i).OpenStart + 2 > selS Then
                unb = True
            End If
        ElseIf selS = selE Then
            If selS >= spans(i).OpenStart And selS <= spans(i).CloseEnd Then nt = nt + 1: targets(nt) = i
        ElseIf spans(i).OpenStart < selE And spans(i).CloseEnd > selS Then
            nt = nt + 1: targets(nt) = i
        End If
    Next
    If nt = 0 Then
        If unb Then
            ShowMsg MsgUnbalanced()
            WMFT_HandleEvaluate = True
        End If
        Exit Function
    End If
    WMFT_HandleEvaluate = True

    For i = 1 To nt
        p = ParseSpan(spans(targets(i)).Content)
        If p.Kind = KIND_ERROR Then
            If firstErr = "" Then firstErr = p.ErrMsg
        ElseIf p.Kind = KIND_ASK Or p.Kind = KIND_ANSWERED Then
            delta = 0
            If EvaluateSpan(spans(targets(i)), p, delta, msg, False) Then
                evaluated = evaluated + 1
                For j = targets(i) + 1 To n
                    ShiftSpan spans(j), delta
                Next
                If selS >= spans(targets(i)).ContentStart + p.ReplEnd - 1 Then selS = selS + delta
                If selE >= spans(targets(i)).ContentStart + p.ReplEnd - 1 Then selE = selE + delta
            ElseIf msg <> "" Then
                If firstErr = "" Then firstErr = msg
                ' a WordMat error form has possibly been shown already when msg is empty
            Else
                firstErr = "-"
            End If
        ElseIf nt = 1 Then
            If p.Kind = KIND_DEF Then
                firstErr = "This is a definition (:=). There is nothing to evaluate. Definitions are used by the equations after it."
            Else
                firstErr = "There is nothing to evaluate. End the equation with '= ?' to ask for its value, e.g. ~~F = m" & ChrW$(183) & "a = ?~~"
            End If
        End If
    Next
    On Error Resume Next
    If nt = 1 And evaluated = 1 Then
        ' like the equation-editor path: put the cursor after the result
        ActiveDocument.Range(spans(targets(1)).CloseEnd + delta, spans(targets(1)).CloseEnd + delta).Select
    Else
        ActiveDocument.Range(selS, selE).Select
    End If
    Application.ScreenUpdating = True
    If firstErr <> "" And firstErr <> "-" Then ShowMsg firstErr
    Exit Function
fejl:
    WMFT_HandleEvaluate = True
    ShowMsg "Free-typed equations: unexpected error " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
End Function

Public Function WMFT_HandleSolve() As Boolean
' Called at the top of Maxima.MaximaSolvePar. Solving free-typed equations is not supported yet.
    Dim spans() As FTSpan, n As Long, i As Long, selS As Long, selE As Long, r As Range
    On Error GoTo fejl
    If Documents.Count = 0 Then Exit Function
    If Selection.OMaths.Count > 0 Then Exit Function
    selS = Selection.Start: selE = Selection.End
    Set r = Selection.Range
    r.Expand wdParagraph
    n = FindSpans(r.Start, r.End, spans)
    For i = 1 To n
        If selS = selE Then
            If selS >= spans(i).OpenStart And (selS <= spans(i).CloseEnd Or spans(i).Unbalanced) Then WMFT_HandleSolve = True
        ElseIf spans(i).OpenStart < selE And spans(i).CloseEnd > selS Then
            WMFT_HandleSolve = True
        End If
    Next
    If WMFT_HandleSolve Then
        ShowMsg MsgSolveNotSupported(), "WordMat"
    End If
    Exit Function
fejl:
    WMFT_HandleSolve = False
End Function

Public Sub WMFT_BeginDefinitions(ByVal beforePos As Long)
' Called from CMaxima.FindDefinitions. Collects the free-text definitions that end before beforePos.
    Dim spans() As FTSpan, n As Long, i As Long, p As FTParsed, e As String, s As String
    On Error GoTo fejl
    mDefCount = 0
    mDefIdx = 0
    If beforePos <= 0 Then Exit Sub
    n = FindSpans(ActiveDocument.Content.Start, beforePos, spans)
    If n = 0 Then Exit Sub
    ReDim mDefPos(1 To n): ReDim mDefText(1 To n): ReDim mDefErr(1 To n)
    For i = 1 To n
        If Not spans(i).Unbalanced And spans(i).CloseEnd <= beforePos Then
            p = ParseSpan(spans(i).Content)
            If p.Kind = KIND_DEF Then
                e = ""
                s = ToLinear(p.Expr, e)
                mDefCount = mDefCount + 1
                mDefPos(mDefCount) = spans(i).OpenStart
                mDefText(mDefCount) = Replace(s, ":=", ChrW$(8788))
                If e <> "" Then mDefErr(mDefCount) = "Definition " & Snippet(spans(i)) & ": " & e Else mDefErr(mDefCount) = ""
            End If
        End If
    Next
    mDefIdx = mDefCount
    Exit Sub
fejl:
    mDefCount = 0
    mDefIdx = 0
End Sub

Public Function WMFT_AddDefinitionsAfter(cm As CMaxima, ByVal afterPos As Long, ByRef acc As String) As Integer
' Called from CMaxima.FindDefinitions, which walks the equation-editor objects backwards from the cursor.
' Adds (in the same backwards order) every free-text definition that starts at or after afterPos
' (afterPos < 0: all remaining). Returns 0 = OK, 1 = a definition returned "STOP", 2 = error.
    Dim t As String, s As String
    Do While mDefIdx >= 1
        If afterPos >= 0 And mDefPos(mDefIdx) < afterPos Then Exit Do
        If mDefErr(mDefIdx) <> "" Then
            ShowMsg mDefErr(mDefIdx)
            mDefIdx = 0
            WMFT_AddDefinitionsAfter = 2
            Exit Function
        End If
        s = mDefText(mDefIdx)
        mDefIdx = mDefIdx - 1
        t = cm.AddDefinition(s)
        If t = "STOP" Then
            WMFT_AddDefinitionsAfter = 1
            Exit Function
        End If
        acc = t & acc
    Loop
End Function

' =========================================================================================================
' Evaluation
' =========================================================================================================

Private Function EvaluateSpan(sp As FTSpan, p As FTParsed, ByRef delta As Long, ByRef msg As String, ByVal quietMaximaErrors As Boolean) As Boolean
' Evaluates one span and writes the result into the document. On failure the text is left unchanged and
' msg holds the reason (msg is empty when a WordMat error form has already been shown).
    Dim lin As String, e As String, req As String, saveOut As String, changedOut As Boolean
    Dim cmd As String, res As String, sign As String, newText As String, inVars As String
    Dim r As Range, ur As UndoRecord, ED As ErrorDefinition, lone As Boolean
    On Error GoTo fejl
    msg = ""
    delta = 0

    lin = ToLinear(p.Expr, e)
    If e <> "" Then msg = e: GoTo slut
    WMU_Request = p.OutUnit ' word_math units: "= ? [unit]" (or the old answer's unit) goes to the unit engine
    inVars = "," & Join(IdentifierList(OutsideBrackets(p.Expr)), ",") & ","

    ' FindDefinitions uses the cursor position: definitions before the equation apply
    ActiveDocument.Range(sp.OpenStart, sp.OpenStart).Select

    RunFirst
    If CASengine <> 0 Then
        msg = "Free-typed equations need WordMat's Maxima engine (choose Maxima as the CAS in the settings)."
        GoTo slut
    End If
    If Not PrepareMaxima Then
        If msg = "" Then msg = "The definitions before this equation could not be read."
        GoTo slut
    End If
    If Not omax.MaximaInstalled Then msg = "Maxima is not installed.": GoTo slut
    omax.prevspr = ""

    cmd = Trim$(lin)
    If InStr(cmd, ChrW$(8788)) > 0 Then
        msg = "An equation can't both define (:=) and ask for a value (= ?)."
        GoTo slut
    End If
    cmd = Trim$(GetCmdAfterEqualSign(cmd))
    If cmd = "" Then msg = "There is no expression before '= ?'.": GoTo slut
    If GetCountOfChar(cmd, "(") <> GetCountOfChar(cmd, ")") Then
        msg = "The brackets ( ) don't match."
        GoTo slut
    End If
    lone = IsLoneName(p.Expr) ' word_math free-text solving: "~~a = ?~~"
    If lone Then WMU_Request = vbNullString ' first see whether the name has a value (a request could fail on it)
    omax.Kommando = cmd
    omax.beregn
    If omax.StopNow Then msg = "The calculation was stopped.": GoTo slut
    If WMU_LastError <> vbNullString Then msg = WMU_LastError: GoTo slut ' word_math units: nothing is written

    ED = GetErrorDefinition(omax.MaximaOutput, omax.KommentarOutput)
    If ED.Title <> vbNullString Then
        If quietMaximaErrors Or WMFreeTextSilent Then
            msg = ED.Title & ": " & Replace(Replace(ED.Description, vbCr, " "), vbLf, " ")
        Else
            CheckForError ' shows WordMat's own error form
            msg = ""
        End If
        GoTo slut
    End If
    res = omax.MaximaOutput
    If Len(res) = 0 Or Trim$(res) = "=res1" Then
        msg = "Maxima returned no result."
        If omax.KommentarOutput <> "" Then msg = msg & " " & Left$(Replace(omax.KommentarOutput, vbLf, " "), 300)
        RestartMaxima
        GoTo slut
    End If
    If InStr(res, "fejl") > 0 Or res = "=done" Or res = "=false" Then
        msg = "Maxima could not evaluate this expression."
        GoTo slut
    End If

    newText = FormatResult(res, p.Spaced, inVars, e)
    If e <> "" And lone And Left$(e, 7) = "Unknown" Then
        ' "~~a = ?~~" with no definition of a: rearrange the nearest earlier equation that contains a
        ' (option C, decided 2026-11-15)
        e = vbNullString
        If Not SolveForLoneName(sp, p, Trim$(NormaliseChars(p.Expr)), newText, e) Then msg = e: GoTo slut
    ElseIf e = "" And lone And p.OutUnit <> "" Then
        ' the name has a value and a unit was asked for: evaluate again with the request
        ActiveDocument.Range(sp.OpenStart, sp.OpenStart).Select
        If Not PrepareMaxima Then msg = "The definitions before this equation could not be read.": GoTo slut
        omax.prevspr = ""
        WMU_LastError = vbNullString
        WMU_Request = p.OutUnit
        omax.Kommando = cmd
        omax.beregn
        If omax.StopNow Then msg = "The calculation was stopped.": GoTo slut
        If WMU_LastError <> vbNullString Then msg = WMU_LastError: GoTo slut
        newText = FormatResult(omax.MaximaOutput, p.Spaced, inVars, e)
    End If
    If e <> "" Then
        msg = e
        GoTo slut
    End If

    If changedOut Then OutUnits = saveOut: changedOut = False

    ' The document must still hold the same text (reading definitions undoes its temporary changes)
    If RangeText(sp.ContentStart, sp.ContentEnd) <> sp.Content Then
        msg = "The equation changed while it was being evaluated; nothing was written."
        GoTo slut
    End If

    Set r = ActiveDocument.Range(sp.ContentStart + p.ReplStart - 1, sp.ContentStart + p.ReplEnd - 1)
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "WordMat result"
    r.text = newText
    ur.EndCustomRecord
    delta = Len(newText) - (p.ReplEnd - p.ReplStart)
    LogMsg "OK " & Snippet(sp) & " -> " & newText
    EvaluateSpan = True
    If WMU_LastWarning <> vbNullString Then ShowMsg WMU_LastWarning, "WordMat" ' word_math units (result written)
    GoTo slut
fejl:
    msg = "Unexpected error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not ur Is Nothing Then If ur.IsRecordingCustomRecord Then ur.EndCustomRecord
slut:
    On Error Resume Next
    If changedOut Then OutUnits = saveOut
    WMU_Request = vbNullString
    Application.ScreenUpdating = True
End Function

Private Function FormatResult(ByVal res As String, ByVal spaced As Boolean, ByVal inVars As String, ByRef errMsg As String) As String
' Turns WordMat's output ("=5", "~3.33", "=1/3~0.333", "=100000 N") into the plain text that replaces the
' answer: "= 5", "= 100000 [N]". Unknown symbols left over are reported as errors.
    Dim parts() As String, signs() As String, n As Long, i As Long, s As String, c As String
    Dim depth As Long, cur As String, out As String, ids As Variant, k As Long, bad As String

    res = Replace(res, ChrW$(8289), "")          ' function application
    res = Replace(res, ChrW$(9633), "")          ' empty placeholder
    res = Replace(res, ChrW$(12310), "(")        ' invisible brackets become ordinary ones
    res = Replace(res, ChrW$(12311), ")")
    res = Replace(res, vbCr, ""): res = Replace(res, vbLf, "")
    res = Trim$(res)
    If Left$(res, 1) <> "=" And Left$(res, 1) <> ChrW$(8776) Then res = "=" & res

    ' split at top-level = and ~ signs
    ReDim parts(1 To 10): ReDim signs(1 To 10)
    For i = 1 To Len(res)
        c = Mid$(res, i, 1)
        If c = "(" Or c = "[" Or c = "{" Then depth = depth + 1
        If c = ")" Or c = "]" Or c = "}" Then depth = depth - 1
        If depth = 0 And (c = "=" Or c = ChrW$(8776)) Then
            If n > 0 Then parts(n) = cur
            n = n + 1
            If n > UBound(parts) Then ReDim Preserve parts(1 To n + 10): ReDim Preserve signs(1 To n + 10)
            signs(n) = c
            cur = ""
        Else
            cur = cur & c
        End If
    Next
    parts(n) = cur

    For i = 1 To n
        s = Trim$(parts(i))
        If s = "" Then errMsg = "Maxima returned an empty result.": Exit Function
        If MaximaUnits And InStr(res, "[") = 0 Then s = WrapResultUnit(s) ' legacy units mode only
        s = PowersToSuperscript(s)
        If spaced Then
            out = out & IIf(i > 1, " ", "") & signs(i) & " " & s
        Else
            out = out & signs(i) & s
        End If
    Next

    ' unknown symbols
    ids = IdentifierList(OutsideBrackets(out))
    For k = LBound(ids) To UBound(ids)
        If ids(k) <> "" And Not IsKnownName(CStr(ids(k))) Then bad = bad & IIf(bad = "", "", ", ") & ids(k)
    Next
    If MaximaUnits And InStr(res, "[") = 0 Then ' legacy units mode: a variable from the input that comes back inside a unit was never defined
        ids = IdentifierList(out)
        For k = LBound(ids) To UBound(ids)
            If ids(k) <> "" Then
                If InStr(inVars, "," & ids(k) & ",") > 0 And InStr(", " & bad & ",", ", " & ids(k) & ",") = 0 Then bad = bad & IIf(bad = "", "", ", ") & ids(k)
            End If
        Next
    End If
    If bad <> "" Then
        errMsg = "Unknown symbol(s) left in the result: " & bad & ". Define them first, e.g. ~~" & Split(bad, ",")(0) & " := 5~~."
        Exit Function
    End If
    FormatResult = out
End Function

Private Function WrapResultUnit(ByVal s As String) As String
' WordMat's legacy units mode only (units switched on, no [ ] in the equation): "100000 N" -> "100000 [N]",
' "1.5*10^5 kg*m/s^2" -> "1.5*10^5 [kg*m/s^2]". Results of the unit engine already have their [unit].
    Dim i As Long, c As String, numEnd As Long, rest As String, numPart As String, k As Long
    i = 1
    If Mid$(s, 1, 1) = "-" Or Mid$(s, 1, 1) = ChrW$(8722) Then i = 2
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If IsDigitCh(c) Or c = "." Or c = DecSeparator Then i = i + 1 Else Exit Do
    Loop
    numEnd = i - 1
    ' optional power of ten: *10^n, *10^(n), *10 superscript
    If numEnd > 0 Then
        k = numEnd + 1
        c = Mid$(s, k, 1)
        If c = ChrW$(183) Or c = ChrW$(215) Or c = "*" Or c = ChrW$(8901) Then
            If Mid$(s, k + 1, 2) = "10" Then
                k = k + 3
                If Mid$(s, k, 1) = "^" Then
                    k = k + 1
                    If Mid$(s, k, 1) = "(" Then
                        Do While k <= Len(s) And Mid$(s, k, 1) <> ")": k = k + 1: Loop
                        k = k + 1
                    Else
                        If Mid$(s, k, 1) = "-" Or Mid$(s, k, 1) = ChrW$(8722) Then k = k + 1
                        Do While k <= Len(s) And IsDigitCh(Mid$(s, k, 1)): k = k + 1: Loop
                    End If
                    numEnd = k - 1
                ElseIf SupToNormal(Mid$(s, k, 1)) <> "" Then
                    Do While k <= Len(s) And SupToNormal(Mid$(s, k, 1)) <> "": k = k + 1: Loop
                    numEnd = k - 1
                End If
            End If
        End If
    End If
    numPart = Trim$(Left$(s, numEnd))
    rest = Trim$(Mid$(s, numEnd + 1))
    Do While Len(rest) > 0
        c = Left$(rest, 1)
        If c = ChrW$(183) Or c = "*" Or c = ChrW$(8901) Or c = " " Then rest = Mid$(rest, 2) Else Exit Do
    Loop
    If rest = "" Or Not IsUnitText(rest) Then
        WrapResultUnit = s
    ElseIf numPart = "" Then
        WrapResultUnit = "1 [" & rest & "]"
    Else
        WrapResultUnit = numPart & " [" & rest & "]"
    End If
End Function

' =========================================================================================================
' Parsing
' =========================================================================================================

Private Function ParseSpan(ByVal content As String) As FTParsed
' Classifies the text between the markers. Offsets refer to content (1-based).
    Dim p As FTParsed, body As String, bodyLen As Long, q As Long, k As Long, after As String
    Dim i As Long, c As String, depth As Long, nRel As Long, relPos() As Long, tail As String, lhs As String

    ' "#(n)" numbering just before the closing marker
    body = content
    k = InStrRev(body, "#(")
    If k > 0 Then
        If Right$(RTrim$(body), 1) = ")" And InStr(k, body, ")") = Len(RTrim$(body)) Then body = Left$(body, k - 1)
    End If
    bodyLen = Len(RTrim$(body))
    body = Left$(body, bodyLen)
    If Trim$(body) = "" Then p.Kind = KIND_PLAIN: ParseSpan = p: Exit Function

    q = InStr(body, "?")
    If InStr(body, ":=") > 0 Or InStr(body, ChrW$(8788)) > 0 Then
        If q > 0 Then
            p.Kind = KIND_ERROR
            p.ErrMsg = "An equation can't both define (:=) and ask for a value (= ?)."
        Else
            p.Kind = KIND_DEF
            p.Expr = Trim$(Replace(body, ChrW$(8788), ":="))
        End If
        ParseSpan = p
        Exit Function
    End If

    If q > 0 Then
        ' "?" must be the last thing, optionally followed by a [unit], and must follow an = sign
        after = Trim$(Mid$(body, q + 1))
        If after <> "" Then
            If Not (Left$(after, 1) = "[" And Right$(after, 1) = "]" And InStr(2, after, "[") = 0 And InStr(after, "]") = Len(after)) Then
                p.Kind = KIND_ERROR
                p.ErrMsg = "The '?' must come at the end of the equation (optionally followed by an output unit, e.g. = ? [kN])."
                ParseSpan = p: Exit Function
            End If
            p.OutUnit = Trim$(Mid$(after, 2, Len(after) - 2))
            If p.OutUnit = "" Then
                p.Kind = KIND_ERROR: p.ErrMsg = "The output unit [ ] after '?' is empty."
                ParseSpan = p: Exit Function
            End If
        End If
        k = q - 1
        Do While k >= 1
            If Mid$(body, k, 1) <> " " Then Exit Do
            k = k - 1
        Loop
        If k < 1 Then
            p.Kind = KIND_ERROR: p.ErrMsg = "There is no expression before '= ?'."
            ParseSpan = p: Exit Function
        End If
        If Mid$(body, k, 1) <> "=" Then
            p.Kind = KIND_ERROR: p.ErrMsg = "Ask for a value with '= ?' at the end of the equation, e.g. ~~2+3 = ?~~."
            ParseSpan = p: Exit Function
        End If
        p.Expr = Trim$(Left$(body, k - 1))
        If p.Expr = "" Then
            p.Kind = KIND_ERROR: p.ErrMsg = "There is no expression before '= ?'."
            ParseSpan = p: Exit Function
        End If
        p.Kind = KIND_ASK
        p.ReplStart = k
        p.ReplEnd = bodyLen + 1
        p.Spaced = (Mid$(body, k + 1, 1) = " ") Or (Mid$(body, k - 1, 1) = " ")
        ParseSpan = p
        Exit Function
    End If

    ' No "?": an answered evaluation or a plain equation. Collect the top-level = and ~ signs.
    ReDim relPos(1 To Len(body) + 1)
    For i = 1 To Len(body)
        c = Mid$(body, i, 1)
        If c = "(" Or c = "[" Or c = "{" Then depth = depth + 1
        If c = ")" Or c = "]" Or c = "}" Then depth = depth - 1
        If depth = 0 And (c = "=" Or c = ChrW$(8776)) Then nRel = nRel + 1: relPos(nRel) = i
    Next
    p.Kind = KIND_PLAIN
    ' the answer starts at the earliest sign after which everything is "result text"
    For i = 1 To nRel
        tail = Mid$(body, relPos(i) + 1)
        If IsResultText(tail) Then
            lhs = Trim$(Left$(body, relPos(i) - 1))
            If lhs = "" Then Exit For
            ' with only one sign before the answer (no "F = m*a = ..." chain), the left side must be
            ' a pure number expression, so that a plain equation such as "x = 5" is never overwritten
            ' (word_math recalculation, decided 2026-11-16: a lone name followed by a result, "~~a = 2 [m/s^2]~~",
            ' is an answer too, so calculate re-solves or re-evaluates it; accepted trade-off: "~~x = 5~~" typed by
            ' the user is recalculated when something gives x a value)
            ' a name on the left of the FIRST sign: not an answer there, unless it is a lone name and the only
            ' sign; otherwise try the next sign ("y = 2*3 = 99" is answered at the second sign)
            If Not (i = 1 And IdentifierCount(OutsideBrackets(lhs)) > 0 And Not (IsLoneName(lhs) And nRel = 1)) Then
                p.Kind = KIND_ANSWERED
                p.Expr = lhs
                p.ReplStart = relPos(i)
                p.ReplEnd = bodyLen + 1
                p.Spaced = (Mid$(body, relPos(i) + 1, 1) = " ") Or (Mid$(body, relPos(i) - 1, 1) = " ")
                p.OutUnit = LastBracketUnit(tail)
                Exit For
            End If
        End If
    Next
    ParseSpan = p
End Function

Public Function WMFT_AnswerCut(ByVal t As String, ByRef cutPos As Long, ByRef oldUnit As String) As Boolean
' word_math recalculation for Word equations: True if the linear text t already ends with an answer written by
' WordMat, e.g. "F=m*a=10" or "2+3=5" or "1 [bar]*1 [m^2]=100 [kN]". cutPos is the 1-based position of the sign
' where the answer starts; oldUnit is the answer's [unit] (empty if none), used as the unit request so that
' recalculation doubles as a unit check. With only one sign the left side must be pure numbers (a typed equation
' such as "x^2=4" is never treated as answered).
    Dim i As Long, c As String, depth As Long, nRel As Long, relPos() As Long, lhs As String, tail As String
    ReDim relPos(1 To Len(t) + 1)
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c = "(" Or c = "[" Or c = "{" Then depth = depth + 1
        If c = ")" Or c = "]" Or c = "}" Then depth = depth - 1
        If depth = 0 And (c = "=" Or c = ChrW$(8776)) Then nRel = nRel + 1: relPos(nRel) = i
    Next
    For i = 1 To nRel
        tail = Mid$(t, relPos(i) + 1)
        If IsResultText(tail) Then
            lhs = Trim$(Left$(t, relPos(i) - 1))
            If lhs = "" Then Exit Function
            ' a name on the left of the first sign: try the next sign ("y=2*3=99" is answered at the second)
            If Not (i = 1 And IdentifierCount(OutsideBrackets(lhs)) > 0) Then
                cutPos = relPos(i)
                oldUnit = LastBracketUnit(tail)
                WMFT_AnswerCut = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function IsResultText(ByVal t As String) As Boolean
' True if t looks like a result written by WMFreeText: numbers, + - * / ^ ( ), powers of ten, pi, sqrt,
' approximately-equal signs and [unit] groups. Symbolic results are not recognised (see design note).
    Dim i As Long, c As String, j As Long, hasNum As Boolean
    t = Trim$(t)
    If t = "" Then Exit Function
    i = 1
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If c = "[" Then
            j = InStr(i, t, "]")
            If j = 0 Then Exit Function
            i = j + 1
        Else
            If IsDigitCh(c) Or SupToNormal(c) <> "" Or c = ChrW$(960) Or c = ChrW$(8734) Then
                hasNum = True
            ElseIf InStr(" .,+-*/^()=" & ChrW$(183) & ChrW$(215) & ChrW$(8901) & ChrW$(8722) & ChrW$(8211) & ChrW$(8776) & ChrW$(8730) & ChrW$(160), c) > 0 Then
            ElseIf c = "e" Or c = "i" Then
            Else
                Exit Function
            End If
            i = i + 1
        End If
    Loop
    IsResultText = hasNum
End Function

Private Function LastBracketUnit(ByVal t As String) As String
    Dim a As Long, b As Long
    t = Trim$(t)
    If Right$(t, 1) <> "]" Then Exit Function
    a = InStrRev(t, "[")
    b = Len(t)
    If a > 0 Then LastBracketUnit = Trim$(Mid$(t, a + 1, b - a - 1))
End Function

' =========================================================================================================
' Text normalisation into WordMat's linear format
' =========================================================================================================

Private Function ToLinear(ByVal s As String, ByRef errMsg As String) As String
' Free text -> the kind of linear (UnicodeMath-style) string ReadSelection gives for an equation object.
    errMsg = ""
    s = BridgeUnits(s, errMsg)
    If errMsg <> "" Then Exit Function
    s = NormaliseChars(s)
    s = TidySpaces(s)
    ToLinear = s
End Function

Private Function BridgeUnits(ByVal s As String, ByRef errMsg As String) As String
' word_math units: [unit] brackets are passed through unchanged. CMaxima.CodeForMaxima hands them to the unit
' engine (WMUnits.bas -> WMUnits.mac), exactly as for equation-editor equations, so both input modes share one
' path (aim 2). Only the bracket structure is checked here; the engine checks the contents.
    Dim i As Long, depth As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "[" Then
            depth = depth + 1
            If depth > 1 Then errMsg = "Unit brackets [ ] can't be nested.": Exit Function
        ElseIf c = "]" Then
            depth = depth - 1
            If depth < 0 Then errMsg = "A unit bracket ] has no matching [.": Exit Function
        End If
    Next
    If depth > 0 Then errMsg = "A unit bracket [ is not closed with ].": Exit Function
    BridgeUnits = s
End Function

Private Function NormaliseChars(ByVal s As String) As String
    Dim out As String, i As Long, c As String, run As String, w As Long, sup As String, subs As String
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        w = AscW(c)
        sup = SupToNormal(c)
        subs = SubToNormal(c)
        If sup <> "" Then
            run = ""
            Do While i <= Len(s)
                sup = SupToNormal(Mid$(s, i, 1))
                If sup = "" Then Exit Do
                run = run & sup
                i = i + 1
            Loop
            If Len(run) = 1 Then out = out & "^" & run Else out = out & "^(" & run & ")"
        ElseIf subs <> "" Then
            run = ""
            Do While i <= Len(s)
                subs = SubToNormal(Mid$(s, i, 1))
                If subs = "" Then Exit Do
                run = run & subs
                i = i + 1
            Loop
            out = out & "_" & run
        Else
            Select Case w
            Case 8211, 8212, 8722                       ' en dash, em dash, minus sign
                out = out & "-"
            Case 42, 215, 8901, 8729, 8226              ' * x (times) dot operator, bullet operator, bullet
                out = out & ChrW$(183)
            Case 160, 8239, 8201, 8194, 8195, 8202, 9   ' no-break / thin / en / em / hair spaces, tab
                out = out & " "
            Case 173, 8203, 8204, 8205, 65279, 11       ' soft hyphen, zero-width characters, line break
            Case Else
                out = out & c
            End Select
            i = i + 1
        End If
    Loop
    NormaliseChars = out
End Function

Private Function TidySpaces(ByVal s As String) As String
' Removes spaces next to operators and brackets; keeps single spaces between two operands, which WordMat
' reads as implicit multiplication (as it does for equation-editor text).
    Dim out As String, i As Long, c As String, prev As String, nxt As String
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Trim$(s)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = " " Then
            prev = Right$(out, 1)
            nxt = Mid$(s, i + 1, 1)
            If Not (IsOperatorCh(prev) Or IsOperatorCh(nxt) Or nxt = ")" Or prev = "(") Then out = out & c
        Else
            out = out & c
        End If
    Next
    TidySpaces = out
End Function

Private Function PowersToSuperscript(ByVal s As String) As String
' "10^5" -> "10" + superscript 5, "m/s^2" -> superscript 2, "^(-3)" -> superscript minus 3.
' Only integer exponents are converted; anything else stays as "^...".
    Dim out As String, i As Long, j As Long, c As String, ex As String, ok As Boolean
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "^" Then
            ex = "": ok = False
            If Mid$(s, i + 1, 1) = "(" Then
                j = InStr(i + 2, s, ")")
                If j > 0 Then
                    ex = Mid$(s, i + 2, j - i - 2)
                    ok = IsSignedInteger(ex)
                End If
                If ok Then out = out & ToSuperscript(ex): i = j + 1: GoTo nxt
            Else
                j = i + 1
                If Mid$(s, j, 1) = "-" Or Mid$(s, j, 1) = ChrW$(8722) Then j = j + 1
                Do While j <= Len(s)
                    If Not IsDigitCh(Mid$(s, j, 1)) Then Exit Do
                    j = j + 1
                Loop
                ex = Mid$(s, i + 1, j - i - 1)
                If IsSignedInteger(ex) Then out = out & ToSuperscript(ex): i = j: GoTo nxt
            End If
        End If
        out = out & c
        i = i + 1
nxt:
    Loop
    PowersToSuperscript = out
End Function

' =========================================================================================================
' Finding spans in the document
' =========================================================================================================

Private Function FindSpans(ByVal fromPos As Long, ByVal toPos As Long, spans() As FTSpan) As Long
' Finds the ~~...~~ spans in the paragraphs that overlap fromPos..toPos. Markers pair up left to right
' within one paragraph; an opening marker left without a partner is returned with Unbalanced = True.
    Dim rAll As Range, para As Paragraph, pr As Range, t As String, n As Long
    Dim a As Long, b As Long, pos As Long, last As Long, pa As Long, pb As Long, oldMode As Boolean
    ReDim spans(1 To 16)
    If toPos <= fromPos Then toPos = fromPos + 1
    Set rAll = ActiveDocument.Range(fromPos, toPos)
    oldMode = rAll.TextRetrievalMode.IncludeHiddenText
    rAll.TextRetrievalMode.IncludeHiddenText = True
    If InStr(rAll.Paragraphs(1).Range.text & rAll.text, MARKER) = 0 Then
        rAll.TextRetrievalMode.IncludeHiddenText = oldMode
        FindSpans = 0
        Exit Function
    End If
    rAll.TextRetrievalMode.IncludeHiddenText = oldMode
    For Each para In rAll.Paragraphs
        Set pr = para.Range
        oldMode = pr.TextRetrievalMode.IncludeHiddenText
        pr.TextRetrievalMode.IncludeHiddenText = True
        t = pr.text
        If InStr(t, MARKER) > 0 Then
            pos = 1
            last = pr.Start
            Do
                a = InStr(pos, t, MARKER)
                If a = 0 Then Exit Do
                pa = MapMarker(pr, a, last)
                If pa < 0 Then Exit Do
                b = InStr(a + 2, t, MARKER)
                n = n + 1
                If n > UBound(spans) Then ReDim Preserve spans(1 To n * 2)
                spans(n).OpenStart = pa
                spans(n).ContentStart = pa + 2
                If b = 0 Then
                    spans(n).Unbalanced = True
                    spans(n).ContentEnd = pa + 2
                    spans(n).CloseEnd = pa + 2
                    spans(n).content = ""
                    Exit Do
                End If
                pb = MapMarker(pr, b, pa + 2)
                If pb < 0 Then
                    spans(n).Unbalanced = True
                    spans(n).ContentEnd = pa + 2
                    spans(n).CloseEnd = pa + 2
                    Exit Do
                End If
                spans(n).ContentEnd = pb
                spans(n).CloseEnd = pb + 2
                spans(n).content = RangeText(pa + 2, pb)
                last = pb + 2
                pos = b + 2
            Loop
        End If
        pr.TextRetrievalMode.IncludeHiddenText = oldMode
    Next
    ' keep only spans that overlap the requested range
    Dim i As Long, m As Long
    For i = 1 To n
        If spans(i).CloseEnd >= fromPos And spans(i).OpenStart <= toPos Then
            m = m + 1
            If m <> i Then spans(m) = spans(i)
        End If
    Next
    FindSpans = m
End Function

Private Function MapMarker(pr As Range, ByVal offset As Long, ByVal minPos As Long) As Long
' Document position of the marker found at 1-based offset in pr.Text. Offsets normally map 1:1; if the
' paragraph holds fields or other objects they may not, so then the marker is searched for from minPos.
    Dim cand As Long, k As Long
    cand = pr.Start + offset - 1
    If cand >= minPos Then
        If RangeText(cand, cand + 2) = MARKER Then MapMarker = cand: Exit Function
    End If
    For k = minPos To pr.End - 2
        If RangeText(k, k + 2) = MARKER Then MapMarker = k: Exit Function
    Next
    MapMarker = -1
End Function

Private Function RangeText(ByVal a As Long, ByVal b As Long) As String
    Dim r As Range
    Set r = ActiveDocument.Range(a, b)
    r.TextRetrievalMode.IncludeHiddenText = True
    RangeText = r.text
End Function

Private Sub ShiftSpan(sp As FTSpan, ByVal delta As Long)
    sp.OpenStart = sp.OpenStart + delta
    sp.ContentStart = sp.ContentStart + delta
    sp.ContentEnd = sp.ContentEnd + delta
    sp.CloseEnd = sp.CloseEnd + delta
End Sub

Private Function Snippet(sp As FTSpan) As String
    Dim s As String
    s = MARKER & sp.content & IIf(sp.Unbalanced, "", MARKER)
    If sp.Unbalanced Then s = RangeText(sp.OpenStart, ActiveDocument.Range(sp.OpenStart, sp.OpenStart).Paragraphs(1).Range.End - 1)
    If Len(s) > 60 Then s = Left$(s, 57) & "..."
    Snippet = s
End Function

' =========================================================================================================
' Marker style
' =========================================================================================================

Private Function GetMarkerStyle(ByVal create As Boolean) As Style
    Dim st As Style
    On Error Resume Next
    Set st = ActiveDocument.Styles(MARKER_STYLE)
    On Error GoTo 0
    If st Is Nothing And create Then
        Set st = ActiveDocument.Styles.Add(MARKER_STYLE, wdStyleTypeCharacter)
    End If
    If Not st Is Nothing Then
        If create Then st.Font.Hidden = True
    End If
    Set GetMarkerStyle = st
End Function

Private Sub UnhideRange(r As Range)
    r.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
    If r.Font.Hidden <> False Then r.Font.Hidden = False
End Sub

' =========================================================================================================
' Small helpers
' =========================================================================================================

Private Sub ShowMsg(ByVal msg As String, Optional ByVal Title As String = "")
    LogMsg "MSG " & Replace(Replace(msg, vbCr, " "), vbLf, " ")
    If WMFreeTextSilent Then Exit Sub
    If Title = "" Then Title = TT.Error
    Application.ScreenUpdating = True
    MsgBox2 msg, vbOKOnly, Title
End Sub

Private Sub LogMsg(ByVal s As String)
    WMFreeTextLog = WMFreeTextLog & s & vbLf
End Sub

Private Function MsgUnbalanced() As String
    MsgUnbalanced = "The ~~ markers don't pair up in this paragraph. A free-typed equation must start and end with ~~ in the same paragraph."
End Function

Private Function MsgSolveNotSupported() As String
    MsgSolveNotSupported = "The Solve button doesn't work on free-typed equations (~~ ... ~~)." & vbCrLf & vbCrLf & _
        "To solve for an unknown, write the equation, then ask for the unknown on its own:" & vbCrLf & _
        "~~F = m" & ChrW$(183) & "a~~  and later  ~~a = ?~~  (or ~~a = ? [m/s^2]~~)." & vbCrLf & _
        "Every other symbol needs a definition (:=)."
End Function

' ---------------------------------------------------------------------------------------------------------------
' word_math free-text solving (option C, decided 2026-11-15): "~~a = ?~~" where a has no definition rearranges
' the nearest earlier equation (free text or equation editor; not a definition, no "?") that contains a and whose
' other symbols are all defined. The unit engine's WMUnitSolve does the algebra, so units work as for evaluation.
' ---------------------------------------------------------------------------------------------------------------

Private Function SolveForLoneName(sp As FTSpan, p As FTParsed, ByVal v As String, ByRef newText As String, ByRef msg As String) As Boolean
    Dim cands() As String, n As Long, i As Long, ck As String, e As String, firstErr As String, res As String
    n = CollectSourceEquations(sp.OpenStart, v, cands)
    If n = 0 Then
        msg = "No earlier equation contains " & v & ". Write the equation first, e.g. ~~F = m" & ChrW$(183) & "a~~, then ~~" & v & " = ?~~."
        Exit Function
    End If
    For i = 1 To n
        ActiveDocument.Range(sp.OpenStart, sp.OpenStart).Select
        If Not PrepareMaxima Then msg = "The definitions before this equation could not be read.": Exit Function
        omax.prevspr = ""
        WMU_LastError = vbNullString
        WMU_Request = p.OutUnit
        ck = omax.CodeForMaxima(cands(i))
        If ck = vbNullString Then GoTo nxt
        omax.WMUnitsSolve ck, v
        If omax.StopNow Then msg = "The calculation was stopped.": Exit Function
        If WMU_LastError <> vbNullString Then
            If firstErr = vbNullString Then firstErr = WMU_LastError & " (equation: " & cands(i) & ")"
            GoTo nxt
        End If
        res = omax.MaximaOutput
        If InStr(res, ChrW$(8744)) > 0 Then
            msg = "The equation " & cands(i) & " has more than one solution for " & v & ": " & res & _
                  ". Solve it with an equation-editor equation and choose the solution there."
            Exit Function
        End If
        If Left$(res, Len(v)) = v Then res = Mid$(res, Len(v) + 1) ' "a=2 [m/s^2]" -> "=2 [m/s^2]"
        e = vbNullString
        newText = FormatResult(res, p.Spaced, "," & Join(IdentifierList(OutsideBrackets(cands(i))), ",") & ",", e)
        If e = vbNullString Then
            LogMsg "SOLVED " & v & " from " & cands(i)
            SolveForLoneName = True
            Exit Function
        End If
        If Left$(e, 7) <> "Unknown" Then msg = e: Exit Function
nxt:
    Next
    If firstErr <> vbNullString Then
        msg = firstErr
    Else
        msg = "No earlier equation with " & v & " can be solved yet: each one still has symbols without a definition (:=)."
    End If
End Function

Private Function CollectSourceEquations(ByVal beforePos As Long, ByVal v As String, cands() As String) As Long
' Earlier equations that contain the name v, nearest first (at most 30): plain free-text equations and
' equation-editor equations, but no definitions and nothing with "?". Returns the count; cands() holds linear text.
    Dim spans() As FTSpan, n As Long, i As Long, pp As FTParsed, body As String, e As String, lin As String
    Dim pos() As Long, txt() As String, m As Long, j As Long, k As Long, tp As Long, tt As String, om As OMath
    Dim para As Paragraph, pr As Range, prevStart As Long, oldMode As Boolean, t As String, a As Long, b As Long
    Dim sp1 As FTSpan, nOm As Long, iOm As Long
    ReDim pos(1 To 1): ReDim txt(1 To 1)
    ' Walk the paragraphs one by one with .Next (a For Each over a range that ends inside a table can loop
    ' forever in Word), and only look for spans in paragraphs that hold a marker.
    Set para = ActiveDocument.Paragraphs.First
    prevStart = -1
    Do While Not para Is Nothing
        Set pr = para.Range
        If pr.Start >= beforePos Or pr.Start <= prevStart Then Exit Do
        prevStart = pr.Start
        oldMode = pr.TextRetrievalMode.IncludeHiddenText
        pr.TextRetrievalMode.IncludeHiddenText = True
        If InStr(pr.text, MARKER) > 0 Then
            ' read the ~~...~~ pairs straight from the paragraph text: no Range enumeration (which can loop
            ' inside table cells); positions are only used for ordering and the "before" test
            t = pr.text
            a = InStr(1, t, MARKER)
            Do While a > 0
                b = InStr(a + 2, t, MARKER)
                If b = 0 Then Exit Do
                sp1.OpenStart = pr.Start + a - 1
                sp1.ContentStart = sp1.OpenStart + 2
                sp1.ContentEnd = pr.Start + b - 1
                sp1.CloseEnd = sp1.ContentEnd + 2
                sp1.Content = Mid$(t, a + 2, b - a - 2)
                sp1.Unbalanced = False
                CollectOne sp1, beforePos, v, pos, txt, m
                a = InStr(b + 2, t, MARKER)
            Loop
        End If
        pr.TextRetrievalMode.IncludeHiddenText = oldMode
        Set para = para.Next
    Loop
    ' by index, as WordMat's FindDefinitions does (For Each over OMaths can loop forever when one sits in a table)
    nOm = ActiveDocument.Range.OMaths.Count
    For iOm = 1 To nOm
        Set om = ActiveDocument.Range.OMaths(iOm)
        If om.Range.Start >= beforePos Then Exit For
        If om.Range.End <= beforePos Then
            lin = om.Range.text
            If InStr(lin, "=") > 0 And InStr(lin, "?") = 0 And InStr(lin, ChrW$(8788)) = 0 And InStr(lin, ":=") = 0 Then
                lin = omax.ReadEquation2(om.Range)
                If InStr(lin, "?") = 0 And InStr(lin, ChrW$(8788)) = 0 And HasName(OutsideBrackets(NormaliseChars(lin)), v) Then
                    m = m + 1: ReDim Preserve pos(1 To m): ReDim Preserve txt(1 To m)
                    pos(m) = om.Range.Start: txt(m) = lin
                End If
            End If
        End If
    Next
    ' nearest first: sort by position, descending
    For j = 2 To m
        tp = pos(j): tt = txt(j): k = j - 1
        Do While k >= 1
            If pos(k) >= tp Then Exit Do
            pos(k + 1) = pos(k): txt(k + 1) = txt(k): k = k - 1
        Loop
        pos(k + 1) = tp: txt(k + 1) = tt
    Next
    If m > 30 Then m = 30
    If m > 0 Then
        ReDim cands(1 To m)
        For j = 1 To m: cands(j) = txt(j): Next
    End If
    CollectSourceEquations = m
End Function

Private Sub CollectOne(sp As FTSpan, ByVal beforePos As Long, ByVal v As String, pos() As Long, txt() As String, m As Long)
' Adds one free-text span to the candidates if it is a plain equation before beforePos that contains v
    Dim pp As FTParsed, body As String, lin As String, e As String
    If sp.Unbalanced Or sp.CloseEnd > beforePos Then Exit Sub
    pp = ParseSpan(sp.Content)
    If pp.Kind <> KIND_PLAIN Then Exit Sub
    body = PlainBody(sp.Content)
    If InStr(body, "=") = 0 Or InStr(body, "?") > 0 Then Exit Sub
    If Not HasName(OutsideBrackets(body), v) Then Exit Sub
    lin = ToLinear(body, e)
    If e <> vbNullString Or lin = vbNullString Then Exit Sub
    m = m + 1: ReDim Preserve pos(1 To m): ReDim Preserve txt(1 To m)
    pos(m) = sp.OpenStart: txt(m) = lin
End Sub

Private Function PlainBody(ByVal content As String) As String
' The text of a plain free-text equation without a "#(n)" number at the end
    Dim k As Long
    PlainBody = content
    k = InStrRev(PlainBody, "#(")
    If k > 0 Then
        If Right$(RTrim$(PlainBody), 1) = ")" Then PlainBody = Left$(PlainBody, k - 1)
    End If
    PlainBody = Trim$(PlainBody)
End Function

Private Function HasName(ByVal s As String, ByVal v As String) As Boolean
' True if the identifier v occurs in s
    Dim ids As Variant, i As Long
    ids = IdentifierList(NormaliseChars(s))
    For i = LBound(ids) To UBound(ids)
        If ids(i) = v Then HasName = True: Exit Function
    Next
End Function

Private Function IsLoneName(ByVal s As String) As Boolean
' True if s is just one name, e.g. "a" or "v_0"
    Dim ids As Variant
    s = Trim$(NormaliseChars(s))
    ids = IdentifierList(s)
    If UBound(ids) = 0 Then IsLoneName = (ids(0) = s And s <> "")
End Function

Private Function IsDigitCh(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    IsDigitCh = (c >= "0" And c <= "9")
End Function

Private Function IsLetterCh(ByVal c As String) As Boolean
    Dim w As Long
    If Len(c) = 0 Then Exit Function
    w = AscW(c)
    If (w >= 65 And w <= 90) Or (w >= 97 And w <= 122) Then IsLetterCh = True: Exit Function
    If w = 181 Or w = 197 Or w = 198 Or w = 216 Or w = 229 Or w = 230 Or w = 248 Then IsLetterCh = True: Exit Function
    If w >= 913 And w <= 969 And w <> 960 Then IsLetterCh = True ' Greek, except pi (a constant)
End Function

Private Function IsOperatorCh(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsOperatorCh = True: Exit Function
    IsOperatorCh = InStr("+-*/^=<>(,;:!" & ChrW$(183) & ChrW$(8776) & ChrW$(8788) & ChrW$(8800) & ChrW$(8804) & ChrW$(8805) & ChrW$(8730) & ChrW$(247), c) > 0
End Function

Private Function IsUnitText(ByVal u As String) As Boolean
' Unit names, prefixes, * / ^, numbers, ( ) and spaces only, with at least one letter.
    Dim i As Long, c As String, hasLetter As Boolean
    For i = 1 To Len(u)
        c = Mid$(u, i, 1)
        If IsLetterCh(c) Or AscW(c) = 937 Or AscW(c) = 176 Or AscW(c) = 8486 Then
            hasLetter = True
        ElseIf IsDigitCh(c) Or SupToNormal(c) <> "" Then
        ElseIf InStr(" */^()-" & ChrW$(183) & ChrW$(8901) & ChrW$(215) & ChrW$(8722), c) > 0 Then
        Else
            Exit Function
        End If
    Next
    IsUnitText = hasLetter
End Function

Private Function IsSignedInteger(ByVal s As String) As Boolean
    Dim i As Long
    If Left$(s, 1) = "-" Or Left$(s, 1) = ChrW$(8722) Then s = Mid$(s, 2)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Not IsDigitCh(Mid$(s, i, 1)) Then Exit Function
    Next
    IsSignedInteger = True
End Function

Private Function SupToNormal(ByVal c As String) As String
    If Len(c) = 0 Then Exit Function
    Select Case AscW(c)
    Case 8304: SupToNormal = "0"
    Case 185: SupToNormal = "1"
    Case 178: SupToNormal = "2"
    Case 179: SupToNormal = "3"
    Case 8308 To 8313: SupToNormal = CStr(AscW(c) - 8304)
    Case 8314: SupToNormal = "+"
    Case 8315: SupToNormal = "-"
    Case 8317: SupToNormal = "("
    Case 8318: SupToNormal = ")"
    End Select
End Function

Private Function SubToNormal(ByVal c As String) As String
    If Len(c) = 0 Then Exit Function
    Select Case AscW(c)
    Case 8320 To 8329: SubToNormal = CStr(AscW(c) - 8320)
    End Select
End Function

Private Function ToSuperscript(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
        Case "0": out = out & ChrW$(8304)
        Case "1": out = out & ChrW$(185)
        Case "2": out = out & ChrW$(178)
        Case "3": out = out & ChrW$(179)
        Case "4" To "9": out = out & ChrW$(8304 + CInt(c))
        Case "-", ChrW$(8722): out = out & ChrW$(8315)
        End Select
    Next
    ToSuperscript = out
End Function

Private Function OutsideBrackets(ByVal s As String) As String
' s with every [ ... ] group removed
    Dim out As String, i As Long, c As String, depth As Long
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "[" Then
            depth = depth + 1
            out = out & " "
        ElseIf c = "]" Then
            If depth > 0 Then depth = depth - 1
        ElseIf depth = 0 Then
            out = out & c
        End If
    Next
    OutsideBrackets = out
End Function

Private Function IdentifierList(ByVal s As String) As Variant
' Names (a letter followed by letters, digits or _) in s
    Dim out As String, i As Long, c As String, cur As String
    For i = 1 To Len(s) + 1
        c = Mid$(s, i, 1)
        If IsLetterCh(c) Or AscW(c & " ") = 937 Or (cur <> "" And (IsDigitCh(c) Or c = "_")) Then
            cur = cur & c
        Else
            If cur <> "" Then out = out & IIf(out = "", "", Chr$(1)) & cur
            cur = ""
        End If
    Next
    IdentifierList = Split(out, Chr$(1))
    If out = "" Then IdentifierList = Array("")
End Function

Private Function IdentifierCount(ByVal s As String) As Long
    Dim ids As Variant, k As Long
    ids = IdentifierList(s)
    For k = LBound(ids) To UBound(ids)
        If ids(k) <> "" Then IdentifierCount = IdentifierCount + 1
    Next
End Function

Private Function IsKnownName(ByVal s As String) As Boolean
' Constants and function names that may appear in a numeric result
    IsKnownName = InStr(",e,i,ln,log,lg,sin,cos,tan,cot,sec,csc,arcsin,arccos,arctan,asin,acos,atan,sinh,cosh,tanh,exp,sqrt,abs,inf,", "," & s & ",") > 0
End Function

' =========================================================================================================
' Test and diagnostic entry points (used by word_math\sandbox\spikes_freetext\selftest_freetext.ps1)
' =========================================================================================================

Public Function WMFT_Ping() As String
    WMFT_Ping = "WMFreeText OK"
End Function

Public Function WMFT_ClearLog() As String
    WMFreeTextLog = ""
    WMFT_ClearLog = "cleared"
End Function

Public Function WMFT_GetLog() As String
    WMFT_GetLog = WMFreeTextLog
End Function

Public Function WMFT_SetSilent(ByVal onOff As Boolean) As String
    WMFreeTextSilent = onOff
    WMFT_SetSilent = CStr(WMFreeTextSilent)
End Function

Public Function WMFT_Setting(ByVal name As String, Optional ByVal newValue As Variant) As String
' Reads (and optionally sets) one WordMat setting for automated tests; returns the previous value.
' Note: WordMat stores these settings in HKCU\Software\WordMat, so the test restores them afterwards.
    RunFirst
    Select Case LCase$(name)
    Case "units"
        WMFT_Setting = CStr(CInt(MaximaUnits) * -1)
        If Not IsMissing(newValue) Then MaximaUnits = (CInt(newValue) <> 0)
    Case "exact"
        WMFT_Setting = CStr(MaximaExact)
        If Not IsMissing(newValue) Then MaximaExact = CInt(newValue)
    Case "outunits"
        WMFT_Setting = OutUnits
        If Not IsMissing(newValue) Then OutUnits = CStr(newValue)
    Case "decseparator"
        WMFT_Setting = DecSeparator
    End Select
End Function

Public Function WMFT_DebugLinear(ByVal s As String) As String
' free text -> linear text -> Maxima code, for diagnosing the translation
    Dim e As String, lin As String
    lin = ToLinear(s, e)
    If e <> "" Then WMFT_DebugLinear = "ERROR: " & e: Exit Function
    RunFirst
    SetMaxProc
    If omax Is Nothing Then Set omax = New CMaxima
    WMFT_DebugLinear = lin & " => " & omax.CodeForMaxima(lin)
End Function

Public Function WMFT_DebugParse(ByVal content As String) As String
    Dim p As FTParsed
    p = ParseSpan(content)
    WMFT_DebugParse = "kind=" & p.Kind & "|expr=" & p.Expr & "|repl=" & p.ReplStart & "-" & p.ReplEnd & "|unit=" & p.OutUnit & "|err=" & p.ErrMsg
End Function
