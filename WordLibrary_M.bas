Attribute VB_Name = "WordLibrary_M"
Option Explicit

' 2025-07-01 1st version, InsertCrossReference

Public Const REFERENCE_FORMAT As String = "{NUMBER}. {TEXT}, page {PAGENUMBER}"

Private Sub TestCrossReferencePicture()
  Dim refType As String, refNum As String
  Dim myRange As Word.Range
  ' Get the reference type (e.g., "Table" or "Figure")
  refType = InputBox("Enter reference type (e.g., Table, Figure):", "Cross-reference type")
  ' Get the reference number or name
  refNum = InputBox("Enter the reference number or name:", "Reference")
  ' Check if the user entered something
  If refNum <> "" Then
    ' Create a range object (e.g., at the current cursor position)
    Set myRange = Selection.Range

    ' Insert the cross-reference
    myRange.InsertCrossReference ReferenceType:=refType, _
                               ReferenceKind:=wdOnlyLabelAndNumber, _
                               ReferenceItem:=refNum, _
                               InsertAsHyperlink:=True
  End If
End Sub

Private Sub TestCrossReferenceHeading()
  Dim refList As Variant
  Dim headingNumber As String
  Dim i As Long
  ' Assuming the heading number is "1.2"
  headingNumber = "1.2"
  ' Get the list of available headings
  refList = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
  ' Find the index of the heading with the desired number
  For i = 1 To UBound(refList)
    If InStr(1, refList(i), headingNumber, vbTextCompare) > 0 Then
      ' Insert the cross-reference to the heading
      Selection.InsertCrossReference ReferenceType:=wdRefTypeHeading, _
                                   ReferenceKind:=wdNumberNoContext, _
                                   ReferenceItem:=i, _
                                   InsertAsHyperlink:=True
      Exit For
    End If
  Next i
  If i > UBound(refList) Then
    MsgBox "Heading not found: " & headingNumber
  End If
End Sub

Private Sub TestWordEditor()
    Dim editor As New clsWordEditor
    
    ' Examples of initialization
    editor.InitializeAtCursor
    editor.InsertText "Text and newline at cursor." & vbCrLf
    Exit Sub
    
    'editor.InitializeAtCursor
    'editor.InsertText "(inserted at cursor, no newline"
    'editor.InsertText ")"
    'Exit Sub
    
    'editor.InitializeAtStartOfCurrentParagraph
    'editor.InsertText "Start of current paragraph." & vbCrLf
    'Exit Sub
    
    'editor.InitializeAtStartOfCurrentParagraph
    'editor.InsertText "(Start of current paragraph)"
    'Exit Sub
    
    'editor.InitializeAtStartOfCurrentParagraph
    'editor.InsertText "(Text)"
    'editor.InsertNewLine
    'Exit Sub
    
    editor.InitializeAtStartOfNextParagraph
    editor.InsertText "Start of next paragraph."
    editor.InsertNewLine
    editor.InsertText "(After a new line)"
    Exit Sub
End Sub


