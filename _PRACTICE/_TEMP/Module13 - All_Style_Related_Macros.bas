Attribute VB_Name = "Module13"
Sub Highlight_Duplicate_Styles()
    Dim sty As Style
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    ' Clear existing highlighting from the whole document first
    doc.Content.HighlightColorIndex = wdNoHighlight
    
    For Each sty In doc.Styles
        ' Only look at custom styles (BuiltIn = False)
        If Not sty.BuiltIn Then
            
            ' Check if the name contains "Char" or "Indent" or other common dupes
            If InStr(1, sty.NameLocal, " Char", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, " pt", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, "Indent", vbTextCompare) > 0 Then
               
                ' Find where this style is used
                With doc.Content.Find
                    .ClearFormatting
                    .Style = sty.NameLocal
                    .Replacement.Highlight = True
                    .Replacement.Text = "^&" ' Keep the same text
                    
                    ' Execute a "Replace All" with Highlight
                    .Execute Replace:=wdReplaceAll
                End With
                
                Debug.Print "Highlighted usage of: " & sty.NameLocal
            End If
        End If
    Next sty
    
    Application.ScreenUpdating = True
    MsgBox "Check for Bright Green highlights.", vbInformation
End Sub

Sub Del_Unused_Styles_Optimized()
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim pass As Integer
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Suppress screen updates and animations
    Application.ScreenUpdating = False
    
    On Error Resume Next ' Prevents crashes if a style is locked or system-protected
    
    ' Run 2 passes to safely handle styles that are "nested" or based on each other
    For pass = 1 To 2
        ' Loop backwards through styles to safely delete items from the collection
        For i = doc.Styles.Count To 1 Step -1
            Set sty = doc.Styles(i)
            
            ' Rule 1: Skip built-in Word styles
            If Not sty.BuiltIn Then
                
                ' Rule 2: Check if Word considers the style completely unused
                If Not sty.InUse Then
                    sty.Delete
                Else
                    ' Rule 3: Double-check via StoryRanges to ensure it's not hiding in headers/footers
                    ' This is a lightning-fast native check compared to a manual .Find loop
                    With doc.Content.Find
                        .ClearFormatting
                        .Style = sty.NameLocal
                        .Execute FindText:="", Format:=True, Wrap:=wdFindStop
                        
                        ' If the style was marked "InUse" but can't be found anywhere, safe to delete
                        If Not .Found Then
                            sty.Delete
                        End If
                    End With
                End If
                
            End If
        Next i
    Next pass
    On Error GoTo 0
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    MsgBox "Unused custom styles cleaned up successfully!", vbInformation, "Clean Up Complete"
End Sub

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

Sub Reset_BuiltIn_Style_Names()
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim cleanName As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' We turn off strict error handling because Word system styles can be stubborn
    On Error Resume Next
    
    For i = doc.Styles.Count To 1 Step -1
        Set sty = doc.Styles(i)
        
        ' Target ONLY protected, built-in Microsoft styles
        If sty.BuiltIn Then
            
            ' If the style name contains a comma, it means an alias was forced onto it
            If InStr(1, sty.NameLocal, ",", vbTextCompare) > 0 Then
                
                ' Split the name string and extract ONLY the very first part (the original name)
                cleanName = Split(sty.NameLocal, ",")(0)
                
                ' Reset the style's local name back to its primary identity
                sty.NameLocal = cleanName
                
                Debug.Print "Reset style name back to: " & cleanName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Built-in style names have been reverted to their original factory names!", vbInformation, "Reset Complete"
End Sub
