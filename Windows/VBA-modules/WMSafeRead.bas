Attribute VB_Name = "WMSafeRead"
' word_math safe read
' ---------------------------------------------------------------------------------------------------------
' Non-destructive reading of Word equations (OMaths) as linear UnicodeMath text.
' WordMat used to linearise the equation in the user's document and Undo the change (ReadEquation2). Here the
' equation is copied into a private hidden scratch document and BuildUp/Linearize/ConvertToNormalText run
' there, so the user's document, its undo history, its Saved state and the selection never change, and the
' text returned is the same as before. Design and tests: word_math research\06_safe_equation_read.md.
' ---------------------------------------------------------------------------------------------------------
Option Explicit

Private mScratch As Document        ' hidden scratch document, created on demand, closed when Word is idle
Private mCloseQueued As Boolean     ' an OnTime close is pending
Public WMSR_TestOldReader As Boolean ' TEST ONLY (WMTest_ReadTiming): route reads to the old reader

Public Function WMSR_ReadOMath(ByVal ir As Range) As String
' Linear text of the first OMath in ir (as ReadEquation2 returned it), read from a copy.
' OMath.BuildUp in the hidden scratch document makes it the ActiveDocument (ActiveWindow stays the user's, and
' Document.Activate does not switch back). Reselecting the user's selection does, so that is done afterwards;
' the selection itself is unchanged (upstream's ReadEquation2 reselected it the same way). Errors are passed on.
    Dim sd As Document, om As OMath, r As Range, ad As Document, sel As Range
    Dim en As Long, es As String, ed As String
    If WMSR_TestOldReader Then WMSR_ReadOMath = WMSR_ReadOMathUndo(ir): Exit Function ' TEST ONLY
    On Error Resume Next
    Set ad = ActiveDocument
    Set sel = Selection.Range
    On Error GoTo fejl
    Set sd = ScratchDoc()
    sd.Content.Text = vbCr & "x"            ' the equation never ends the scratch document (Word end-of-document bug)
    Set r = sd.Range(0, 0)
    r.FormattedText = ir.OMaths(1).Range.FormattedText
    If sd.OMaths.Count > 0 Then
        Set om = sd.OMaths(1)
        om.Range.Font.Bold = False          ' bold gives BuildUp problems (as before, but on the copy)
        om.BuildUp
        om.Linearize
        om.ConvertToNormalText
        WMSR_ReadOMath = om.Range.Text
    End If
fejl:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not sd Is Nothing Then
        sd.Content.Delete
        sd.UndoClear
        sd.Saved = True
    End If
    RestoreActive ad, sel
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, es, ed
End Function

Private Sub RestoreActive(ad As Document, sel As Range)
' Makes ad the ActiveDocument again (see WMSR_ReadOMath).
    On Error Resume Next
    If ad Is Nothing Then Exit Sub
    If ActiveDocument.FullName = ad.FullName Then Exit Sub
    If Not sel Is Nothing Then sel.Select
    If ActiveDocument.FullName <> ad.FullName Then ad.Activate
End Sub

Public Function WMSR_ScratchIsOpen() As Boolean
' For the tests: True while the scratch document exists.
    Dim n As String
    On Error Resume Next
    If Not mScratch Is Nothing Then n = mScratch.Name
    WMSR_ScratchIsOpen = (Len(n) > 0)
End Function

Private Function ScratchDoc() As Document
' The scratch document; created hidden if needed. The user's active document stays active.
    Dim ok As Boolean
    On Error Resume Next
    If Not mScratch Is Nothing Then
        ok = (Len(mScratch.Name) > 0)       ' fails if it was closed behind our back
        If Not ok Then Set mScratch = Nothing
    End If
    Err.Clear
    If mScratch Is Nothing Then
        On Error GoTo 0
        Set mScratch = Documents.Add(Visible:=False)
        mScratch.Saved = True
        QueueClose
    End If
    Set ScratchDoc = mScratch
End Function

Private Sub QueueClose()
' Closes the scratch document as soon as Word is idle again (after the current command).
    If mCloseQueued Then Exit Sub
    On Error Resume Next
    Application.OnTime When:=Now, Name:="WMSR_CloseScratch"
    mCloseQueued = (Err.Number = 0)
End Sub

Public Sub WMSR_CloseScratch()
' Closes the scratch document without saving. Also called by OnTime.
    mCloseQueued = False
    On Error Resume Next
    If Not mScratch Is Nothing Then
        mScratch.Saved = True
        mScratch.Close SaveChanges:=wdDoNotSaveChanges
    End If
    Set mScratch = Nothing
End Sub

' ---------------------------------------------------------------------------------------------------------
' TEST ONLY: WordMat's old reader (CMaxima.ReadEquation2 up to fork main b9a6d02), kept unchanged so the
' equality test (WMTest_ReadCompare) can compare it with WMSR_ReadOMath. It changes the document and undoes
' the change; never call it from production code.
' ---------------------------------------------------------------------------------------------------------
Public Function WMSR_ReadOMathUndo(Optional ir As Range) As String
    Dim sr As Range
    Set sr = Selection.Range

    If ir Is Nothing Then
        Set ir = sr
    End If
    Dim OM As OMath
    Dim Oundo As UndoRecord
    Set Oundo = Application.UndoRecord
    Set OM = ir.OMaths(1)
    OM.Range.Font.Bold = False
    Oundo.StartCustomRecord
    If OM.Range.End + 100 >= ActiveDocument.Range.End Then
        OM.Range.Select
        Selection.Collapse wdCollapseEnd
        Selection.TypeParagraph
        Selection.TypeText "x"
    End If
    OM.BuildUp
    OM.Linearize
    OM.ConvertToNormalText
    WMSR_ReadOMathUndo = OM.Range.text
    Oundo.EndCustomRecord
    ActiveDocument.Undo
    sr.Select
End Function

Public Function WMSR_ReadOMathsUndo(ByVal r As Range) As String
' TEST ONLY: the old several-OMaths read of CMaxima.ReadSelection / DocData.ReadMathBoxes (on r instead of the
' selection): one change for all OMaths in r, then Undo. The texts are joined with " | ".
    Dim mathcol As OMaths, j As Long, n As Long, o As String
    Dim Oundo As UndoRecord
    n = r.OMaths.Count
    Set Oundo = Application.UndoRecord
    Oundo.StartCustomRecord
    Set mathcol = r.OMaths
    With r.OMaths
        .BuildUp
        .Linearize
    End With
    For j = 0 To n - 1
        mathcol(j + 1).ConvertToNormalText
        o = o & IIf(j > 0, " | ", "") & Replace(mathcol(j + 1).Range.text, ChrW$(11), "")
    Next
    Oundo.EndCustomRecord
    ActiveDocument.Undo
    WMSR_ReadOMathsUndo = o
End Function

