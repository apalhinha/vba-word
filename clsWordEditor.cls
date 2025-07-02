VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWordEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === clsWordEditor ===
Option Explicit

' 2025-07-02, h, p, t (table) and smart quotes
' 2025-07-01, InsertReference

Private mRange As Range

Public Sub InsertReference(styleName As String, prefix As Variant)
    Dim refList As Variant
    Dim i As Long
    Dim foundIndex As Long
    Dim found As Boolean
    Dim itemText As String
    Dim refType As String
    Dim shouldCloseQuote As Boolean

    found = False

    If LCase(styleName) = "p" Then
        refType = "Picture"
        refList = ActiveDocument.GetCrossReferenceItems(refType)

        For i = 1 To UBound(refList)
            itemText = Trim(CStr(refList(i)))
            If itemText Like prefix & "*" Then
                foundIndex = i
                found = True
                Exit For
            End If
        Next i

    ElseIf LCase(styleName) = "t" Then
        refType = "Table"
        refList = ActiveDocument.GetCrossReferenceItems(refType)

        For i = 1 To UBound(refList)
            itemText = Trim(CStr(refList(i)))
            If itemText Like prefix & "*" Then
                foundIndex = i
                found = True
                Exit For
            End If
        Next i

    ElseIf LCase(styleName) = "h" Then
        refList = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)

        For i = 1 To UBound(refList)
            itemText = Trim(CStr(refList(i)))
            If itemText Like prefix & "*" Then
                foundIndex = i
                found = True
                Exit For
            End If
        Next i

    Else
        MsgBox "Unsupported style type: " & styleName, vbExclamation
        Exit Sub
    End If

    If Not found Then
        MsgBox "No Match found for prefix: " & prefix, vbExclamation
        Exit Sub
    End If

    ' Move selection to our insertion point
    mRange.Select

    ' Determine if the current line is empty
    Dim currentLineRange As Range
    Set currentLineRange = Selection.Paragraphs(1).Range
    If Len(Trim(Replace(currentLineRange.Text, Chr(13), ""))) > 0 And Trim(Replace(currentLineRange.Text, Chr(13), "")) <> Chr(9) & "Next process:" & Chr(9) Then
        ' Line has text, insert opening quote
        Selection.TypeText ChrW(CLng(8216))
        shouldCloseQuote = True
    End If

    ' Parse and build the reference according to REFERENCE_FORMAT
    Dim pos As Long
    Dim nextPos As Long
    Dim tag As String

    pos = 1
    Do While pos <= Len(REFERENCE_FORMAT)
        nextPos = InStr(pos, REFERENCE_FORMAT, "{")
        If nextPos = 0 Then
            Selection.TypeText Mid(REFERENCE_FORMAT, pos)
            Exit Do
        Else
            If nextPos > pos Then
                Selection.TypeText Mid(REFERENCE_FORMAT, pos, nextPos - pos)
            End If
            Dim endBrace As Long
            endBrace = InStr(nextPos, REFERENCE_FORMAT, "}")
            If endBrace > 0 Then
                tag = Mid(REFERENCE_FORMAT, nextPos + 1, endBrace - nextPos - 1)

                Select Case tag
                    Case "NUMBER"
                        If LCase(styleName) = "h" Then
                            Selection.InsertCrossReference ReferenceType:=wdRefTypeHeading, _
                                ReferenceKind:=wdNumberNoContext, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        Else
                            Selection.InsertCrossReference ReferenceType:=refType, _
                                ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        End If

                    Case "TEXT"
                        If LCase(styleName) = "h" Then
                            Selection.InsertCrossReference ReferenceType:=wdRefTypeHeading, _
                                ReferenceKind:=wdContentText, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        Else
                            Selection.InsertCrossReference ReferenceType:=refType, _
                                ReferenceKind:=wdOnlyCaptionText, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        End If

                    Case "PAGENUMBER"
                        If LCase(styleName) = "h" Then
                            Selection.InsertCrossReference ReferenceType:=wdRefTypeHeading, _
                                ReferenceKind:=wdPageNumber, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        Else
                            Selection.InsertCrossReference ReferenceType:=refType, _
                                ReferenceKind:=wdPageNumber, ReferenceItem:=foundIndex, _
                                InsertAsHyperlink:=True, IncludePosition:=False, _
                                SeparateNumbers:=False, SeparatorString:=" "
                        End If
                End Select

                pos = endBrace + 1
            Else
                Selection.TypeText Mid(REFERENCE_FORMAT, nextPos)
                Exit Do
            End If
        End If
    Loop

    ' Close quote if line originally had text
    If shouldCloseQuote Then
        Selection.TypeText ChrW(CLng(8217))
    End If

    ' Set mRange to after inserted content
    Set mRange = Selection.Range.Duplicate
    mRange.Collapse wdCollapseEnd
End Sub



' ===== Initialization methods =====
Public Sub InitializeAtCursor()
    Set mRange = Selection.Range.Duplicate
End Sub

Public Sub InitializeAtStartOfCurrentParagraph()
    Set mRange = Selection.Paragraphs(1).Range.Duplicate
    mRange.Collapse Direction:=wdCollapseStart
End Sub

Public Sub InitializeAtStartOfNextParagraph()
    Dim nextPara As Paragraph
    On Error Resume Next
    Set nextPara = Selection.Paragraphs(1).Next
    On Error GoTo 0
    
    If Not nextPara Is Nothing Then
        Set mRange = nextPara.Range.Duplicate
        mRange.Collapse Direction:=wdCollapseStart
    Else
        Set mRange = Selection.Paragraphs(1).Range.Duplicate
        mRange.Collapse Direction:=wdCollapseEnd
    End If
End Sub

' ===== Insert methods =====
Public Sub InsertText(ByVal txt As String)
    mRange.InsertAfter txt
    mRange.Collapse Direction:=wdCollapseEnd
    mRange.Select
End Sub

Public Sub InsertNewLine()
    mRange.InsertParagraphAfter
    mRange.Collapse Direction:=wdCollapseEnd
    mRange.Select
End Sub

