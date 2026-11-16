Attribute VB_Name = "WMUnits"
' word_math units
' ---------------------------------------------------------------------------------------------------------
' VBA side of the word_math unit engine (aim 1). The engine itself is Maxima code:
'   Shared\Maxima-files\WMUnits.mac       rules, dimension check, conversion, auto-fill
'   Shared\Maxima-files\WMUnitTable.mac   the unit table (data only)
' Both are copied to %AppData%\WordMat by sandbox\setup_wordmat.ps1 and loaded once per Maxima process
' (CMaxima.GetMaximaSettingsString adds WMU_LoadStatement).
'
' This module only
'   1. recognises [ ... ] spans: WMU_Protect at the very start of CMaxima.CodeForMaxima replaces each one by
'      a placeholder name, so no later CodeForMaxima pass can mangle it (psi -> Greek, m -> variable, the
'      decimal-comma and Greek passes, ...); WMU_Restore at the end turns the placeholders into
'         wmq(5,"kN*m")   for "5 [kN*m]" (a number directly before the bracket: one quantity)
'         wmu("m")        for any other bracket ("x [m]" -> x*wmu("m"))
'      with the unit as an ASCII string (degree signs, Delta, micro, Omega, superscripts mapped);
'   2. supports CMaxima.WMUnitsBeregn, which sends WMUnitResult(...) and formats its answer as "value [unit]";
'   3. extracts the engine's error messages ("WMUNIT: ...").
' Units in [ ] are handled whether or not WordMat's own (legacy) units mode is switched on.
' Design: word_math\research\05_unit_engine_design.md
' ---------------------------------------------------------------------------------------------------------
Option Explicit

Public WMU_Used As Boolean          ' a [unit] was translated since the last WMU_BeginCommand
Public WMU_Request As String        ' "= ? [unit]": the unit as the user typed it ("" = auto-fill)
Public WMU_LastError As String      ' set by CMaxima.WMUnitsBeregn: the message to show (nothing is written)
Public WMU_LastWarning As String    ' set by CMaxima.WMUnitsBeregn: shown after the result is written

Private Const PH As String = "WMUQ"  ' placeholder prefix: upper case, no Greek-letter or unit substrings
Private mCode() As String           ' Maxima code for placeholder i
Private mCount As Long

' ---------------------------------------------------------------------------------------------------------
' Per command
' ---------------------------------------------------------------------------------------------------------
Public Sub WMU_BeginCommand()
' Called from CMaxima.PrepareNewCommand (before the definitions are read).
    mCount = 0
    ReDim mCode(0 To 0)
    WMU_Used = False
    WMU_LastError = vbNullString
    WMU_LastWarning = vbNullString
End Sub

Public Function WMU_LoadStatement() As String
' "errcatch(if wmunits_loaded#true then load(""...WMUnits.mac"")),", for the start of the settings list;
' empty when the file can't be found (bracket units then give a clear error).
    Dim p As String
    p = WMU_MacPath()
    If p = vbNullString Then Exit Function
    WMU_LoadStatement = "errcatch(if wmunits_loaded#true then load(""" & Replace(p, "\", "/") & """)),"
End Function

Public Function WMU_MacPath() As String
    Static cached As String
    Dim c As Variant, p As String
    On Error Resume Next
    If cached <> vbNullString Then
        If Dir(cached) <> vbNullString Then WMU_MacPath = cached: Exit Function
    End If
    For Each c In Array(Environ("AppData") & "\WordMat\", Environ("ProgramFiles") & "\WordMat\", ThisDocument.Path & "\")
        p = CStr(c) & "WMUnits.mac"
        If Len(CStr(c)) > 1 Then
            If Dir(p) <> vbNullString Then cached = p: WMU_MacPath = p: Exit Function
        End If
    Next
End Function

' ---------------------------------------------------------------------------------------------------------
' Input: [ ... ] -> placeholders -> wmq/wmu calls
' ---------------------------------------------------------------------------------------------------------
Public Function WMU_Protect(ByVal s As String) As String
' Called first thing in CMaxima.CodeForMaxima. Brackets holding "," or ";" (lists, intervals), matrices
' (Word's U+25A0 matrix with & and @) and nested brackets are left alone, as before.
    Dim out As String, i As Long, j As Long, c As String, u As String, k As Long, numStart As Long
    Dim numTxt As String, prev As String, a As String, e As String, code As String, sgnPos As Long
    If InStr(s, "[") = 0 Then WMU_Protect = s: Exit Function
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "[" Then
            j = InStr(i + 1, s, "]")
            If j = 0 Then out = out & Mid$(s, i): Exit Do
            u = Mid$(s, i + 1, j - i - 1)
            ' not a unit: nested brackets, lists and intervals (, ;) and matrices in square brackets (U+25A0, & and @)
            If InStr(u, "[") > 0 Or InStr(u, ",") > 0 Or InStr(u, ";") > 0 Or InStr(u, ChrW$(9632)) > 0 Or InStr(u, "&") > 0 Or InStr(u, "@") > 0 Then
                out = out & c
                i = i + 1
                GoTo nxt
            End If
            a = WMU_ToAscii(u, e)
            If e <> vbNullString Then
                code = "wm_err(""" & Replace(Replace(e, "\", "/"), """", "'") & """)" ' word_math security: no \ or " can end the string early
            Else
                code = vbNullString
            End If
            out = RTrim$(out)
            ' a number directly before the bracket forms one quantity with it
            k = Len(out)
            Do While k >= 1
                If IsNumCh(Mid$(out, k, 1)) Then k = k - 1 Else Exit Do
            Loop
            numStart = k + 1
            If numStart <= Len(out) And k > 0 Then
                prev = Mid$(out, k, 1)
                If IsLetterOrDigit(prev) Or prev = "_" Or prev = "^" Or prev = ")" Or prev = "]" Then numStart = Len(out) + 1
            End If
            mCount = mCount + 1
            ReDim Preserve mCode(0 To mCount)
            If numStart <= Len(out) Then
                numTxt = Mid$(out, numStart)
                ' a unary minus belongs to the quantity: -20 [degC] is minus twenty degrees, not -(20 degC)
                sgnPos = numStart - 1
                If sgnPos >= 1 Then
                    If Mid$(out, sgnPos, 1) = "-" Or Mid$(out, sgnPos, 1) = ChrW$(8722) Then
                        prev = Right$(RTrim$(Left$(out, sgnPos - 1)), 1)
                        If prev = "" Or InStr("(=+-*/^,;" & ChrW$(183) & ChrW$(8901) & ChrW$(215) & ChrW$(8776) & ChrW$(8722), prev) > 0 Then
                            numTxt = "-" & numTxt
                            numStart = sgnPos
                        End If
                    End If
                End If
                If DecSeparator <> "." Then numTxt = Replace(numTxt, DecSeparator, ".")
                If code = vbNullString Then code = "wmq(" & numTxt & ",""" & a & """)"
                out = Left$(out, numStart - 1) & PH & Tag(mCount)
            Else
                If code = vbNullString Then code = "wmu(""" & a & """)"
                prev = Right$(out, 1)
                If prev <> "" And Not IsOperatorEnd(prev) Then out = out & "*"
                out = out & PH & Tag(mCount)
            End If
            mCode(mCount) = code
            WMU_Used = True
            ' "5 [m]x" or "[m](a+b)": multiply explicitly, so the placeholder name doesn't run into what follows
            c = Mid$(s, j + 1, 1)
            If IsLetterOrDigit(c) Or c = "(" Or c = ChrW$(12310) Then out = out & "*"
            i = j + 1
        Else
            out = out & c
            i = i + 1
        End If
nxt:
    Loop
    WMU_Protect = out
End Function

Public Function WMU_Restore(ByVal s As String) As String
' Called at the end of CMaxima.CodeForMaxima: placeholders -> Maxima code (longest names first).
    Dim i As Long
    If mCount = 0 Then WMU_Restore = s: Exit Function
    If InStr(s, PH) = 0 Then WMU_Restore = s: Exit Function
    For i = mCount To 1 Step -1
        If InStr(s, PH & Tag(i)) > 0 Then s = Replace(s, PH & Tag(i), mCode(i))
    Next
    WMU_Restore = s
End Function

Private Function Tag(ByVal n As Long) As String
' 1 -> "B", 25 -> "Z", 26 -> "BA": letters only, so no pass sees a number glued to a name
    Dim t As String
    Do
        t = Chr$(65 + (n Mod 26)) & t
        n = n \ 26
    Loop While n > 0
    Tag = t
End Function

Private Function IsNumCh(ByVal c As String) As Boolean
    IsNumCh = (c >= "0" And c <= "9") Or c = "." Or c = DecSeparator
End Function

Private Function IsLetterOrDigit(ByVal c As String) As Boolean
    Dim w As Long
    If c = "" Then Exit Function
    w = AscW(c)
    IsLetterOrDigit = (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (w >= 192 And w <> 215 And w <> 247 And w < 8192)
End Function

Private Function IsOperatorEnd(ByVal c As String) As Boolean
' True when c is an operator or opens a group, i.e. no multiplication sign is needed before a unit
    IsOperatorEnd = InStr("(+-*/^=<>,;:" & ChrW$(183) & ChrW$(8901) & ChrW$(215) & ChrW$(8776) & ChrW$(8722) & ChrW$(8800) & ChrW$(8804) & ChrW$(8805) & ChrW$(8788) & ChrW$(12310), c) > 0
End Function

' ---------------------------------------------------------------------------------------------------------
' Unit text: Word characters <-> the ASCII names of WMUnitTable.mac
' ---------------------------------------------------------------------------------------------------------
Public Function WMU_ToAscii(ByVal u As String, ByRef errMsg As String) As String
    Dim out As String, i As Long, c As String, w As Long, sup As String, run As String, nxt As String
    errMsg = vbNullString
    u = Replace(u, ChrW$(12310), "(")       ' Word's invisible brackets
    u = Replace(u, ChrW$(12311), ")")
    u = Replace(u, ChrW$(9633), "")         ' empty placeholder
    i = 1
    Do While i <= Len(u)
        c = Mid$(u, i, 1)
        w = AscW(c)
        nxt = Mid$(u, i + 1, 1)
        sup = SupDigit(c)
        If sup <> "" Then
            run = ""
            Do While i <= Len(u)
                sup = SupDigit(Mid$(u, i, 1))
                If sup = "" Then Exit Do
                run = run & sup
                i = i + 1
            Loop
            If Len(run) = 1 Then out = out & "^" & run Else out = out & "^(" & run & ")"
            GoTo nxtc
        End If
        Select Case w
        Case 916, 8710                                      ' Delta (Greek capital, increment sign)
            If nxt = ChrW$(176) And (Mid$(u, i + 2, 1) = "C" Or Mid$(u, i + 2, 1) = "F") Then
                out = out & "Ddeg" & Mid$(u, i + 2, 1)
                i = i + 3
                GoTo nxtc
            End If
            errMsg = "[" & u & "]: a Delta in a unit must be followed by " & ChrW$(176) & "C or " & ChrW$(176) & "F."
            Exit Function
        Case 176                                            ' degree sign
            If nxt = "C" Or nxt = "F" Or nxt = "R" Then
                out = out & "deg" & nxt
                i = i + 2
                GoTo nxtc
            End If
            out = out & "deg"
        Case 8451: out = out & "degC"                       ' degree Celsius sign
        Case 8457: out = out & "degF"                       ' degree Fahrenheit sign
        Case 181, 956: out = out & "u"                      ' micro sign, Greek mu
        Case 937, 8486: out = out & "Ohm"                   ' Omega, ohm sign
        Case 183, 8901, 8729, 215, 8226, 8290: out = out & "*"   ' multiplication signs, invisible times
        Case 8722, 8211: out = out & "-"                    ' minus sign, en dash
        Case 160, 8239, 8201, 8194, 8195, 8202, 9: out = out & " "
        Case 8289, 8203, 8204, 8205, 65279, 173             ' invisible characters
        Case Else
            If w < 32 Or w > 126 Or w = 34 Or w = 92 Then
                errMsg = "[" & u & "] is not a valid unit: the character '" & c & "' can't be used in a unit."
                Exit Function
            End If
            out = out & c
        End Select
        i = i + 1
nxtc:
    Loop
    WMU_ToAscii = Trim$(out)
End Function

Private Function SupDigit(ByVal c As String) As String
    If c = "" Then Exit Function
    Select Case AscW(c)
    Case 8304: SupDigit = "0"
    Case 185: SupDigit = "1"
    Case 178: SupDigit = "2"
    Case 179: SupDigit = "3"
    Case 8308 To 8313: SupDigit = CStr(AscW(c) - 8304)
    Case 8315: SupDigit = "-"
    Case 8314: SupDigit = "+"
    End Select
End Function

Public Function WMU_FromAscii(ByVal u As String) As String
' Engine unit text -> what is written into the document: "kN*m" -> kN middle-dot m, "DdegC" -> Delta degree C,
' "uF" -> micro F, "Ohm" -> Omega. Powers stay "^n" (the free-text route turns integer powers into
' superscripts; the equation-editor route builds them up).
    Dim out As String, i As Long, j As Long, c As String, t As String
    i = 1
    Do While i <= Len(u)
        c = Mid$(u, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or c = "_" Then
            j = i
            Do While j <= Len(u)
                c = Mid$(u, j, 1)
                If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or c = "_" Then j = j + 1 Else Exit Do
            Loop
            t = Mid$(u, i, j - i)
            out = out & DisplayName(t)
            i = j
        Else
            If c = "*" Then out = out & ChrW$(183) Else out = out & c
            i = i + 1
        End If
    Loop
    WMU_FromAscii = out
End Function

Private Function DisplayName(ByVal t As String) As String
    Select Case t
    Case "DdegC": DisplayName = ChrW$(916) & ChrW$(176) & "C"
    Case "DdegF": DisplayName = ChrW$(916) & ChrW$(176) & "F"
    Case "degC": DisplayName = ChrW$(176) & "C"
    Case "degF": DisplayName = ChrW$(176) & "F"
    Case "degR": DisplayName = ChrW$(176) & "R"
    Case "deg": DisplayName = ChrW$(176)
    Case "Ohm": DisplayName = ChrW$(937)
    Case Else
        If Len(t) > 3 And Right$(t, 3) = "Ohm" Then
            DisplayName = DisplayName(Left$(t, Len(t) - 3)) & ChrW$(937)   ' kOhm
        ElseIf Len(t) > 1 And Left$(t, 1) = "u" Then
            DisplayName = ChrW$(181) & Mid$(t, 2)                          ' uF -> micro F (no unit starts with u)
        Else
            DisplayName = t
        End If
    End Select
End Function

Public Function WMU_MessageText(ByVal m As String) As String
' Engine message (ASCII) -> display text: unit names inside [ ] are shown as written in the document.
    Dim out As String, i As Long, j As Long
    i = InStr(m, "[")
    Do While i > 0
        j = InStr(i + 1, m, "]")
        If j = 0 Then Exit Do
        out = out & Left$(m, i) & WMU_FromAscii(Mid$(m, i + 1, j - i - 1)) & "]"
        m = Mid$(m, j + 1)
        i = InStr(m, "[")
    Loop
    WMU_MessageText = out & m
End Function

' ---------------------------------------------------------------------------------------------------------
' Output
' ---------------------------------------------------------------------------------------------------------
Public Function WMU_ErrorFromOutput(ByVal raw As String) As String
' The engine's error message (from WMUnitLastError(), "WMUERR@@<message>") in Maxima's output, or "".
' (WordMat's Maxima bridge does not pass Maxima's own error messages on, so the engine keeps them.)
    Dim p As Long, q As Long
    p = InStrRev(raw, "WMUERR@@")
    If p = 0 Then Exit Function
    p = p + 8
    q = InStr(p, raw, """")
    If q = 0 Then q = InStr(p, raw, vbLf)
    If q = 0 Then q = Len(raw) + 1
    WMU_ErrorFromOutput = WMU_MessageText(Trim$(Replace(Mid$(raw, p, q - p), vbCr, "")))
End Function

Public Function WMU_ResultString(ByVal raw As String) As String
' The last "WMU@@..." string in Maxima's output (without quotes), or "".
    Dim p As Long, q As Long
    p = InStrRev(raw, "WMU@@")
    If p = 0 Then Exit Function
    q = InStr(p, raw, """")
    If q = 0 Then q = InStr(p, raw, vbLf)
    If q = 0 Then q = Len(raw) + 1
    WMU_ResultString = Replace(Mid$(raw, p, q - p), vbCr, "")
End Function

' ---------------------------------------------------------------------------------------------------------
' "= ? [unit]" in equation-editor equations
' ---------------------------------------------------------------------------------------------------------
Public Function WMU_TakeAsk(ByRef cmd As String) As Boolean
' If the linear text ends with "= ?" (optionally followed by one [unit]), removes "= ?" and the unit and sets
' WMU_Request. Returns True when there was a "?". The result (with its own = or approximately-equal sign)
' then replaces "= ? [unit]" in the equation (WMU_DeleteAskTail).
    Dim t As String, u As String, k As Long
    t = RTrim$(cmd)
    If Right$(t, 1) = "]" Then
        k = InStrRev(t, "[")
        If k = 0 Then Exit Function
        u = Mid$(t, k + 1, Len(t) - k - 1)
        t = RTrim$(Left$(t, k - 1))
    End If
    If Right$(t, 1) <> "?" Then Exit Function
    t = RTrim$(Left$(t, Len(t) - 1))
    If Right$(t, 1) <> "=" Then Exit Function
    cmd = RTrim$(Left$(t, Len(t) - 1))
    WMU_Request = Trim$(u)
    WMU_TakeAsk = True
End Function

Public Sub WMU_DeleteAnswerTail(om As OMath, ByVal tail As String)
' word_math recalculation: removes the old answer (tail = "=10 [N]" as read, starting with its sign) from the end
' of the equation, so the new result (which starts with its own sign) can be appended. Linearised for this and
' built up again. If the exact tail is not found, the text from the last = or approximately-equal sign is removed.
    Dim r As Range, t As String, k As Long, k2 As Long
    On Error Resume Next
    om.Linearize
    Set r = om.Range
    t = r.text
    k = InStrRev(t, tail)
    If k = 0 Or tail = "" Then
        k = InStrRev(t, "=")
        k2 = InStrRev(t, ChrW$(8776))
        If k2 > k Then k = k2
    End If
    If k > 1 Then ActiveDocument.Range(r.Start + k - 1, r.End).Delete
    om.BuildUp
End Sub
Public Sub WMU_DeleteAskTail(om As OMath)
' Removes the final "= ? [unit]" from the equation, so the result (which starts with its own sign) can be
' appended. The equation is linearised for this and built up again.
    Dim r As Range, k As Long, q As Long
    On Error Resume Next
    om.Linearize
    Set r = om.Range
    q = InStrRev(r.text, "?")
    If q > 0 Then k = InStrRev(r.text, "=", q)
    If k > 0 Then ActiveDocument.Range(r.Start + k - 1, r.End).Delete
    om.BuildUp
End Sub

' ---------------------------------------------------------------------------------------------------------
' Test helpers (Application.Run from tests\lib)
' ---------------------------------------------------------------------------------------------------------
Public Function WMU_Ping() As String
    WMU_Ping = "WMUnits " & WMU_MacPath()
End Function

Public Function WMU_DebugCode(ByVal s As String) As String
' What CodeForMaxima makes of s (placeholders restored).
    RunFirst ' the user's settings (decimal separator, angle mode)
    If omax Is Nothing Then Set omax = New CMaxima
    omax.PrepareNewCommand False ' resets per-command state (also WMU_BeginCommand)
    WMU_DebugCode = omax.CodeForMaxima(s)
End Function

Public Function WMU_DebugMaxima(ByVal cmd As String) As String
' For tests: Maxima's raw output for cmd, sent the way WordMat sends commands (settings string first).
    RunFirst
    If Not PrepareMaxima(False) Then WMU_DebugMaxima = "PrepareMaxima failed": Exit Function
    omax.PrepareNewCommand False
    omax.MaximaInputStreng = cmd
    omax.RunMaxima 0, False, False
    WMU_DebugMaxima = MaxProc.LastMaximaOutput
End Function
' ---------------------------------------------------------------- word_math matrices (research 07)
' eigenval( ), eigenvec( ) and principalangle( ) are defined in Shared\Maxima-files\WMMatrix.mac (loaded by
' WMUnits.mac). CMaxima.beregn sends commands that use them through the unit branch (WMUnitsBeregn), so their
' errors (wm_err) and warnings (WMMatrixWarning) reach the user; a result without units falls back to the
' normal evaluation, which handles exact and numeric output as usual.
Public Function WMM_Uses(ByVal code As String) As Boolean
    WMM_Uses = WMM_HasName(code, "eigenval") Or WMM_HasName(code, "eigenvec") Or WMM_HasName(code, "principalangle")
End Function
Private Function WMM_HasName(ByVal code As String, ByVal nm As String) As Boolean
' nm as a whole name (so Maxima's eigenvalues( ) is not matched)
    Dim p As Long, c As String
    p = InStr(code, nm)
    Do While p > 0
        c = Mid$(code, p + Len(nm), 1)
        If Not (c Like "[A-Za-z0-9_]") Then
            If p = 1 Then WMM_HasName = True: Exit Function
            If Not (Mid$(code, p - 1, 1) Like "[A-Za-z0-9_]") Then WMM_HasName = True: Exit Function
        End If
        p = InStr(p + 1, code, nm)
    Loop
End Function
Public Function WMM_IsAngleCall(ByVal s As String) As Boolean
' True when s (the command as typed) is one call principalangle(...) and nothing else.
    Dim i As Long, d As Long
    s = Replace(Replace(Replace(s, ChrW$(8289), ""), " ", ""), ChrW$(160), "")
    If Left$(s, 15) <> "principalangle(" Or Right$(s, 1) <> ")" Then Exit Function
    For i = 15 To Len(s)
        Select Case Mid$(s, i, 1)
        Case "(": d = d + 1
        Case ")": d = d - 1
            If d = 0 And i < Len(s) Then Exit Function
        End Select
    Next
    WMM_IsAngleCall = (d = 0)
End Function
Public Function WMM_WarningFromOutput(ByVal raw As String) As String
' The warning from WMMatrixWarning() ("WMMWARN@@<message>") in Maxima's output, or "".
    Dim p As Long, q As Long
    p = InStrRev(raw, "WMMWARN@@")
    If p = 0 Then Exit Function
    p = p + 9
    q = InStr(p, raw, """")
    If q = 0 Then q = InStr(p, raw, vbLf)
    If q = 0 Then q = Len(raw) + 1
    WMM_WarningFromOutput = Trim$(Replace(Mid$(raw, p, q - p), vbCr, ""))
End Function
