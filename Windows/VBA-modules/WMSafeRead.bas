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

Public Function WMSR_ReadOMath(ByVal ir As Range) As String
' Linear text of the first OMath in ir (as ReadEquation2 returned it), read from a copy.
    Dim sd As Document, om As OMath, r As Range
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
    sd.Content.Delete
    sd.UndoClear
    sd.Saved = True
End Function

Private Function ScratchDoc() As Document
' The scratch document; created hidden if needed. The user's active document stays active.
    Dim ad As Document, ok As Boolean
    On Error Resume Next
    If Not mScratch Is Nothing Then
        ok = (Len(mScratch.Name) > 0)       ' fails if it was closed behind our back
        If Not ok Then Set mScratch = Nothing
    End If
    Err.Clear
    If mScratch Is Nothing Then
        Set ad = Nothing
        Set ad = ActiveDocument
        On Error GoTo 0
        Set mScratch = Documents.Add(Visible:=False)
        mScratch.Saved = True
        If Not ad Is Nothing Then ad.Activate
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
