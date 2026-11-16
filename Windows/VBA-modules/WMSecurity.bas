Attribute VB_Name = "WMSecurity"
' word_math security (research 08_security_audit.md, 2026-11-16)
' Document text (Word equations, free-typed ~~ ~~ equations, definitions, code blocks) becomes Maxima input, and
' Maxima can run operating-system commands, run Lisp and read or write files. WMSec_Check is called on every
' command just before it is sent to Maxima (CMaxima.RunMaxima, ConvertMaximaUnits, ResetMaxima) and refuses
' commands that name such a function. The Maxima side has a second line of defence: Shared\Maxima-files\
' WMSecurity.lisp (loaded by WMUnits.mac), which also covers names built at run time from strings.
Option Explicit

Public WMSec_LastRefusal As String

' Maxima names refused anywhere outside a string (Maxima is case-sensitive, so only these exact spellings
' are the dangerous functions; "Load" or "load1" are ordinary names).
Private Const DENY_NAMES As String = "|system|to_lisp|eval_string_lisp|eval_string|parse_string|" & _
    "load|loadfile|batch|batchload|demo|run_testsuite|setup_autoload|" & _
    "writefile|appendfile|closefile|save|stringout|with_stdout|printfile|" & _
    "opena|openw|openr|opena_binary|openw_binary|openr_binary|" & _
    "write_data|write_binary_data|read_matrix|read_array|read_hashed_array|read_list|read_nested_list|" & _
    "read_binary_matrix|read_binary_array|read_binary_list|" & _
    "delete_file|rename_file|copy_file|mkdir|rmdir|chdir|" & _
    "compile_file|translate_file|compfile|" & _
    "file_name|user_preamble|png_file|svg_file|pdf_file|ps_file|eps_file|" & _
    "maxima_tempdir|maxima_userdir|geomview_command|xmaxima_plot_command|mgnuplot_command|"
' Name prefixes refused (plot settings that name programs or files, and Maxima's file search lists)
Private Const DENY_PREFIXES As String = "|gnuplot_|file_search|"

Private Function WMSec_IsNameChar(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    WMSec_IsNameChar = (c Like "[A-Za-z0-9_%]") Or (AscW(c) > 127 Or AscW(c) < 0)
End Function

Private Function WMSec_NameRefused(ByVal nm As String) As Boolean
    Dim p As Variant
    If InStr(DENY_NAMES, "|" & nm & "|") > 0 Then WMSec_NameRefused = True: Exit Function
    For Each p In Split(Mid$(DENY_PREFIXES, 2, Len(DENY_PREFIXES) - 2), "|")
        If Left$(nm, Len(p)) = p Then WMSec_NameRefused = True: Exit Function
    Next
End Function

Private Function WMSec_StripOwn(ByVal s As String) As String
' Removes the exact statements WordMathEng itself adds that use load( ): numeric solving, direction fields
' and the user's own code file (always %AppData%\WordMat\WordMatCodeFile.mac).
    s = Replace(s, "load(""newton1"")", "")
    s = Replace(s, "load(""mnewton"")", "")
    s = Replace(s, "load(drawdf)", "")
    s = Replace(s, "load(""" & Replace(GetCodeFileName, "\", "/") & """)", "")
    WMSec_StripOwn = s
End Function

Public Function WMSec_Check(ByVal cmd As String) As String
' Returns "" when cmd may be sent to Maxima, otherwise the reason (the refused name or construct).
    Dim i As Long, n As Long, c As String, inStr_ As Boolean, j As Long, nm As String, k As Long
    cmd = WMSec_StripOwn(cmd)
    n = Len(cmd)
    i = 1
    Do While i <= n
        c = Mid$(cmd, i, 1)
        If inStr_ Then
            If c = "\" Then
                i = i + 1                              ' escaped character inside a string
            ElseIf c = """" Then
                inStr_ = False
            End If
            i = i + 1
        ElseIf c = """" Then
            inStr_ = True
            i = i + 1
        ElseIf c = "?" Then
            WMSec_Check = "?": Exit Function           ' ?name reaches Lisp directly
        ElseIf c = "\" Then
            WMSec_Check = "\": Exit Function           ' a backslash in a name hides it (sys\tem = system)
        ElseIf c = ":" Then
            k = i + 1
            Do While k <= n
                If Mid$(cmd, k, 1) <> " " Then Exit Do
                k = k + 1
            Loop
            If LCase$(Mid$(cmd, k, 4)) = "lisp" And Not WMSec_IsNameChar(Mid$(cmd, k + 4, 1)) Then
                WMSec_Check = ":lisp": Exit Function
            End If
            i = i + 1
        ElseIf WMSec_IsNameChar(c) Then
            j = i
            Do While j <= n
                If Not WMSec_IsNameChar(Mid$(cmd, j, 1)) Then Exit Do
                j = j + 1
            Loop
            nm = Mid$(cmd, i, j - i)
            If WMSec_NameRefused(nm) Then WMSec_Check = nm: Exit Function
            i = j
        Else
            i = i + 1
        End If
    Loop
End Function

Public Function WMSec_Allow(ByVal cmd As String, Optional ByVal ShowMessage As Boolean = True) As Boolean
' True when cmd may be sent to Maxima. Otherwise records the reason in WMSec_LastRefusal and, if asked,
' tells the user.
    Dim r As String
    r = WMSec_Check(cmd)
    If r = vbNullString Then WMSec_Allow = True: Exit Function
    WMSec_LastRefusal = r
    If ShowMessage Then
        MsgBox "WordMathEng security: this was not calculated, because it contains """ & r & """." & vbCrLf & vbCrLf & _
            "That could run commands, use Lisp or read or write files on this computer, so it is not allowed " & _
            "in a calculation. If it is meant as a variable name, choose another name (for example " & _
            "a capital letter or a digit at the end).", vbOKOnly + vbExclamation, "WordMathEng security"
    End If
End Function

Public Function WMSec_LoadStatement() As String
' "errcatch(load("".../WMSecurity.lisp""))$" for a Maxima process that doesn't load WMUnits.mac (WordMat's old
' units process, CMaxima.ConvertMaximaUnits). Loading it again is harmless.
    Dim p As String
    On Error Resume Next
    p = WMU_MacPath()
    If p = vbNullString Then Exit Function
    p = Left$(p, InStrRev(p, "\")) & "WMSecurity.lisp"
    If Dir(p) = vbNullString Then Exit Function
    WMSec_LoadStatement = "errcatch(load(""" & Replace(p, "\", "/") & """))$"
End Function
