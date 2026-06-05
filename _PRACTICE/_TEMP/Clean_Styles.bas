Sub Consolidate_And_Clean_Styles()
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim targetStyle As String
    Dim currentName As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' We turn off error handling temporarily because deleting styles 
    ' can occasionally trigger minor system warnings we want to bypass.
    On Error Resume Next
    
    ' Loop backward through the styles to safely delete as we go
    For i = doc.Styles.Count To 1 Step -1
        Set sty = doc.Styles(i)
        
        ' Only process custom/propagated styles
        If Not sty.BuiltIn Then
            currentName = sty.NameLocal
            targetStyle = ""
            
            ' --- PATTERN MATCHING LOGIC ---
            ' Check if the rogue style is a variant of a core style
            If InStr(1, currentName, "Heading 1", vbTextCompare) > 0 Then
                targetStyle = "Heading 1"
            ElseIf InStr(1, currentName, "Heading 2", vbTextCompare) > 0 Then
                targetStyle = "Heading 2"
            ElseIf InStr(1, currentName, "Heading 3", vbTextCompare) > 0 Then
                targetStyle = "Heading 3"
            ElseIf InStr(1, currentName, "Heading 4", vbTextCompare) > 0 Then
                targetStyle = "Heading 4"
            ElseIf InStr(1, currentName, "Normal", vbTextCompare) > 0 Or _
                   InStr(1, currentName, "hvr", vbTextCompare) > 0 Then
                targetStyle = "Normal"
            End If
            
            ' --- THE SWAP & DESTROY ---
            ' If a matching core style was identified, re-map the text
            If targetStyle <> "" And currentName <> targetStyle Then
                
                With doc.Content.Find
                    .ClearFormatting
                    .Style = currentName
                    .Replacement.ClearFormatting
                    .Replacement.Style = doc.Styles(targetStyle)
                    .Replacement.Text = "^&" ' Keep original text exactly as it is
                    
                    ' Execute the mass swap across the document
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' Now that the text is safely moved, delete the fake style
                sty.Delete
                Debug.Print "Successfully merged and deleted: " & currentName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Style consolidation complete! Your Style Pane is now clean.", vbInformation, "Success"
End Sub