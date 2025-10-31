Attribute VB_Name = "InsertReference_M"
Option Explicit

' ======= CONFIG =======
Private Const PARA_STYLE_NAME As String = "ParaNum"   ' <-- change here if your paragraph-numbered style name changes
Private Const USE_QUOTES As Boolean = False           ' set False to disable quoting, True to use quotes

' ======= TYPES =======
Private Type TNumInfo
    ParaStart As Long
    IsParaNumStyle As Boolean
    ListString As String
    PlainText As String
End Type

Dim aPreviousPrefix As String

' ======= ENTRY POINT =======
Public Sub AutoCrossReference()
    Dim styleKey As String, prefix As String
    
    styleKey = InputBox("Enter the style (h:Heading, p:Paragraph, f:Figure, t:Table)", "Style")
    If styleKey = "" Then Exit Sub
    styleKey = LCase$(Trim$(styleKey))   ' use full string, not Left(…,1)
    If (styleKey <> "h" And styleKey <> "p" And styleKey <> "f" And styleKey <> "t") Then
        MsgBox "Invalid style. Use h, p, f, or t.", vbExclamation
        Exit Sub
    End If
    
    prefix = InputBox("Enter the numeric prefix (e.g., '2.3.' or 'Process 1.2.')", "Prefix")
    If prefix = "" Then Exit Sub
    
    EnsureMainStory
    
    Select Case styleKey
        Case "h": InsertForHeading prefix
        Case "f": InsertForCaption prefix, "Figure"
        Case "t": InsertForCaption prefix, "Table"
        Case "p": InsertForParaNum_Efficient prefix
    End Select
End Sub

Public Sub HeadingCrossReference()
    Dim styleKey As String, prefix As String
    styleKey = "h"
    
    prefix = InputBox("Enter the numeric prefix (e.g., '2.3.' or 'Process 1.2.')", "Prefix", aPreviousPrefix)
    If prefix = "" Then Exit Sub
    aPreviousPrefix = prefix
    
    EnsureMainStory
    InsertForHeading prefix
    
End Sub


' ===================== HEADINGS =====================
Private Sub InsertForHeading(ByVal userPrefix As String)
    Dim items As Variant, idx As Long, normPrefix As String
    
    On Error Resume Next
    items = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    On Error GoTo 0
    If IsEmpty(items) Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    normPrefix = Normalize(userPrefix)
    idx = FindFirstPrefixIndex(items, normPrefix)
    If idx = 0 Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    InsertComposite_H idx
End Sub

Private Sub InsertComposite_H(ByVal idx As Long)
    Dim isParagraphBlank As Boolean
    Dim insertStart As Long, insertEnd As Long
    Dim hadNumber As Boolean
    
    isParagraphBlank = IsCurrentLineBlank()
    insertStart = Selection.Start
    
    ' Number (no context) with local accept-revisions fallback
    hadNumber = TryInsertHeadingNumber(idx)
    If hadNumber Then Selection.TypeText ", "
    
    ' Text
    If Not TryInsertCrossRef(wdRefTypeHeading, wdContentText, idx) Then
        MsgBox "could not insert text for the reference", vbExclamation: Exit Sub
    End If
    
    ' ", page " + Page
    Selection.TypeText ", page "
    If Not TryInsertCrossRef(wdRefTypeHeading, wdPageNumber, idx) Then
        MsgBox "could not insert page number for the reference", vbExclamation: Exit Sub
    End If
    
    insertEnd = Selection.Start
    If Not isParagraphBlank Then WrapInsertedWithQuotes insertStart, insertEnd
End Sub

Private Function TryInsertHeadingNumber(ByVal idx As Long) As Boolean
    Dim startPos As Long: startPos = Selection.Start
    Dim ok As Boolean
    
    ok = TryInsertCrossRef(wdRefTypeHeading, wdNumberNoContext, idx)
    If Not ok Then
        If AcceptRevisionsForHeadingIndex(idx) Then
            ok = TryInsertCrossRef(wdRefTypeHeading, wdNumberNoContext, idx)
        End If
    End If
    
    If ok Then
        Dim rngNum As Range, txt As String
        Set rngNum = ActiveDocument.Range(Start:=startPos, End:=Selection.Start)
        txt = Trim$(RemoveCrAndNonPrint(rngNum.Text))
        If Len(txt) = 0 Then rngNum.Delete: ok = False
    End If
    
    TryInsertHeadingNumber = ok
End Function

' ===================== CAPTIONS (FIGURE/TABLE) =====================
Private Sub InsertForCaption(ByVal userPrefix As String, ByVal labelName As String)
    Dim items As Variant, idx As Long, normPrefix As String
    
    On Error Resume Next
    items = ActiveDocument.GetCrossReferenceItems(labelName)  ' "Figure" or "Table"
    On Error GoTo 0
    If IsEmpty(items) Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    normPrefix = Normalize(userPrefix)
    idx = FindFirstPrefixIndex(items, normPrefix)
    If idx = 0 Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    InsertComposite_Caption labelName, idx
End Sub

Private Sub InsertComposite_Caption(ByVal labelName As String, ByVal idx As Long)
    Dim isParagraphBlank As Boolean
    Dim insertStart As Long, insertEnd As Long
    Dim hadLabelNum As Boolean
    
    isParagraphBlank = IsCurrentLineBlank()
    insertStart = Selection.Start
    
    ' Label + Number
    hadLabelNum = TryInsertCrossRef(labelName, wdOnlyLabelAndNumber, idx)
    If hadLabelNum Then Selection.TypeText ", "
    
    ' Caption text
    If Not TryInsertCrossRef(labelName, wdOnlyCaptionText, idx) Then
        MsgBox "could not insert caption text for the reference", vbExclamation: Exit Sub
    End If
    
    ' ", page " + Page
    Selection.TypeText ", page "
    If Not TryInsertCrossRef(labelName, wdPageNumber, idx) Then
        MsgBox "could not insert page number for the reference", vbExclamation: Exit Sub
    End If
    
    insertEnd = Selection.Start
    If Not isParagraphBlank Then WrapInsertedWithQuotes insertStart, insertEnd
End Sub

' ===================== PARANUM (efficient) =====================
' Build a numbered-item index map once (i -> paragraph info), then iterate
' GetCrossReferenceItems("Numbered item") and test only ParaNum paragraphs.
Private Sub InsertForParaNum_Efficient(ByVal userPrefix As String)
    Dim infos() As TNumInfo, k As Long
    Dim items As Variant, i As Long, normPrefix As String
    Dim foundIdx As Long, cand As String
    
    k = BuildNumberedIndexMap(infos)                 ' O(N) once
    On Error Resume Next
    items = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
    On Error GoTo 0
    If IsEmpty(items) Or k = 0 Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    normPrefix = Normalize(userPrefix)
    If Len(normPrefix) = 0 Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    ' Loop numbered items only, filter to ParaNum style, then prefix-match
    For i = LBound(items) To UBound(items)
        If i >= 1 And i <= k Then
            If infos(i).IsParaNumStyle Then
                ' Try "ListString + text"
                If Len(infos(i).ListString) > 0 Then
                    cand = Normalize(infos(i).ListString & " " & infos(i).PlainText)
                    If Left$(cand, Len(normPrefix)) = normPrefix Then
                        foundIdx = i: Exit For
                    End If
                End If
                ' Try "text" only
                cand = Normalize(infos(i).PlainText)
                If Left$(cand, Len(normPrefix)) = normPrefix Then
                    foundIdx = i: Exit For
                End If
            End If
        End If
    Next i
    
    If foundIdx = 0 Then
        MsgBox "could not find '" & userPrefix & "' reference", vbInformation: Exit Sub
    End If
    
    ' Compose: "Paragraph " + Number + ", page " + Page
    InsertComposite_P foundIdx
End Sub

Private Sub InsertComposite_P(ByVal numberedIdx As Long)
    Dim isParagraphBlank As Boolean
    Dim insertStart As Long, insertEnd As Long
    
    isParagraphBlank = IsCurrentLineBlank()
    insertStart = Selection.Start
    
    Selection.TypeText "Paragraph "
    
    If Not TryInsertCrossRef(wdRefTypeNumberedItem, wdNumberNoContext, numberedIdx) Then
        MsgBox "could not insert number for the paragraph", vbExclamation: Exit Sub
    End If
    
    Selection.TypeText ", page "
    
    If Not TryInsertCrossRef(wdRefTypeNumberedItem, wdPageNumber, numberedIdx) Then
        MsgBox "could not insert page number for the paragraph", vbExclamation: Exit Sub
    End If
    
    insertEnd = Selection.Start
    If Not isParagraphBlank Then WrapInsertedWithQuotes insertStart, insertEnd
End Sub

' ===================== NUMBERED INDEX MAP =====================
' Returns K = count of numbered items (matching Word’s “Numbered item” list).
' Fills infos(1..K) only for numbered, non-bullet paragraphs — same indexing as the dialog.
Private Function BuildNumberedIndexMap(ByRef infos() As TNumInfo) As Long
    Dim sr As Range, p As Paragraph
    Dim k As Long, lt As WdListType
    Dim ls As String, pt As String

    Set sr = ActiveDocument.StoryRanges(wdMainTextStory)
    k = 0
    For Each p In sr.Paragraphs
        lt = p.Range.ListFormat.ListType
        If IsNumberedNotBullet(lt) Then
            k = k + 1
            ReDim Preserve infos(1 To k)
            infos(k).ParaStart = p.Range.Start
            infos(k).IsParaNumStyle = (LCase$(CStr(p.Range.Style)) = LCase$(PARA_STYLE_NAME))
            On Error Resume Next
            ls = p.Range.ListFormat.ListString
            On Error GoTo 0
            infos(k).ListString = ls
            pt = CleanParaText(p.Range.Text)
            infos(k).PlainText = pt
        End If
    Next p
    BuildNumberedIndexMap = k
End Function

' Treat as “numbered item” only if it’s a list AND not a bullet (nor picture-bullet)
Private Function IsNumberedNotBullet(ByVal lt As WdListType) As Boolean
    IsNumberedNotBullet = (lt <> wdListNoNumbering And lt <> wdListBullet And lt <> wdListPictureBullet)
End Function

' ===================== LOW-LEVEL HELPERS =====================
Private Function TryInsertCrossRef(ByVal refType As Variant, ByVal kind As WdReferenceKind, ByVal refItem As Variant) As Boolean
    On Error GoTo Fail
    Selection.InsertCrossReference _
        ReferenceType:=refType, _
        ReferenceKind:=kind, _
        ReferenceItem:=refItem, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False, _
        SeparateNumbers:=False, _
        SeparatorString:=" "
    TryInsertCrossRef = True
    Exit Function
Fail:
    TryInsertCrossRef = False
End Function

' Accept revisions only for the targeted heading, then return to original caret
Private Function AcceptRevisionsForHeadingIndex(ByVal idx As Long) As Boolean
    Dim saveStart As Long, saveEnd As Long, saveStory As WdStoryType
    Dim ok As Boolean
    saveStart = Selection.Start: saveEnd = Selection.End: saveStory = Selection.StoryType
    On Error GoTo Clean
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=idx
    Selection.Paragraphs(1).Range.Revisions.AcceptAll
    ok = True
Clean:
    If saveStory <> wdMainTextStory Then ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.SetRange saveStart, saveEnd
    AcceptRevisionsForHeadingIndex = ok
End Function

Private Sub WrapInsertedWithQuotes(ByVal startPos As Long, ByVal endPos As Long)
    Dim rng As Range
    Set rng = ActiveDocument.Range(Start:=startPos, End:=endPos)
    
    If USE_QUOTES Then
        ' Typographic single quotes: opening ‘ (U+2018), closing ’ (U+2019)
        rng.InsertBefore ChrW$(&H2018)
        rng.InsertAfter ChrW$(&H2019)
        ' Place caret exactly after the closing quote
        Selection.SetRange Start:=rng.End, End:=rng.End
    Else
        ' No quotes; place caret exactly at the end of the inserted content
        Selection.SetRange Start:=endPos, End:=endPos
    End If
End Sub

Private Function IsCurrentLineBlank() As Boolean
    Dim t As String
    t = Selection.Paragraphs(1).Range.Text
    IsCurrentLineBlank = (Len(Trim$(RemoveCrAndNonPrint(t))) = 0)
End Function

Private Function Normalize(ByVal s As String) As String
    Dim txt As String
    txt = Replace$(s, vbTab, " ")
    Do While InStr(txt, "  ") > 0
        txt = Replace$(txt, "  ", " ")
    Loop
    Normalize = LCase$(Trim$(txt))
End Function

Private Function CleanParaText(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, vbCr, "")
    t = Replace$(t, Chr$(7), "")
    t = Replace$(t, vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    CleanParaText = Trim$(t)
End Function

Private Function RemoveCrAndNonPrint(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, vbCr, "")
    t = Replace$(t, Chr$(7), "")
    RemoveCrAndNonPrint = t
End Function

Private Sub EnsureMainStory()
    If Selection.StoryType <> wdMainTextStory Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If
End Sub

Private Function FindFirstPrefixIndex(ByVal items As Variant, ByVal normPrefix As String) As Long
    Dim i As Long, cand As String
    On Error GoTo Done
    For i = LBound(items) To UBound(items)
        cand = Normalize(CStr(items(i)))
        If Left$(cand, Len(normPrefix)) = normPrefix Then
            FindFirstPrefixIndex = i
            Exit Function
        End If
    Next i
Done:
End Function


