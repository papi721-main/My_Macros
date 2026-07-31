Public Sub InsertCaption()
    ' Launch Word's standard Insert Caption dialog window
    On Error Resume Next
    Dialogs(wdDialogInsertCaption).Show
    On Error GoTo 0
    
    ' Immediately inspect the paragraph where the caption was inserted
    Dim currentPara As Paragraph
    Set currentPara = Selection.Paragraphs(1)
    
    ' Check if the text starts with "Figure" or "Fig"
    Dim cleanText As String
    cleanText = Trim(currentPara.Range.text)
    
    If LCase(cleanText) Like "figure*" Or LCase(cleanText) Like "fig*" Then
        ' Uncheck "Keep with next" so the caption isn't pulled away from its image above
        currentPara.KeepWithNext = False
    End If
End Sub
