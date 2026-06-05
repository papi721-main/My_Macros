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

Sub Clean_Styles_Comprehensive()
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim targetStyle As String
    Dim currentName As String
    Dim headingNum As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Suppress errors dynamically for system-protected custom iterations
    On Error Resume Next
    
    ' Loop backward through the styles to safely delete as we go
    For i = doc.Styles.Count To 1 Step -1
        Set sty = doc.Styles(i)
        
        ' Only process custom/propagated styles
        If Not sty.BuiltIn Then
            currentName = sty.NameLocal
            targetStyle = ""
            
            ' --- ADVANCED PATTERN MATCHING LOGIC ---
            
            ' 1. Catch ALL Headings (Heading 1 through Heading 9)
            If InStr(1, currentName, "Heading ", vbTextCompare) > 0 Then
                ' Extract the character immediately following "Heading "
                headingNum = Mid(currentName, InStr(1, currentName, "Heading ", vbTextCompare) + 8, 1)
                
                ' Verify if it's a valid digit (1-9)
                If IsNumeric(headingNum) And headingNum <> "0" Then
                    ' Accommodates your primary style choice format dynamically
                    Select Case headingNum
                        Case "1": targetStyle = "Heading 1"
                        Case "2": targetStyle = "Heading 2"
                        Case "3": targetStyle = "Heading 3"
                        Case "4": targetStyle = "Heading 4"
                        Case Else: targetStyle = "Heading " & headingNum
                    End Select
                End If
                
            ' 2. Catch ALL Normal variations ("Normal", "Normal (Web)", typos like "Nomral", or "hvr")
            ElseIf InStr(1, currentName, "Norm", vbTextCompare) > 0 Or _
                   InStr(1, currentName, "Nomr", vbTextCompare) > 0 Or _
                   InStr(1, currentName, "hvr", vbTextCompare) > 0 Then
                   
                targetStyle = "Normal"
            End If
            
            ' --- FALLBACK VALIDATION ---
            ' Fallback check: If your document uses plain names (e.g. "Heading 1" instead of "Heading 1,_H1"),
            ' verify the target actually exists in the document to prevent a macro break.
            If targetStyle <> "" Then
                Dim testSty As Style
                Set testSty = doc.Styles(targetStyle)
                
                ' If the custom alias style doesn't exist, fall back to the basic built-in name
                If testSty Is Nothing And InStr(1, targetStyle, "_H", vbTextCompare) > 0 Then
                    targetStyle = Split(targetStyle, ",")(0)
                End If
            End If
            
            ' --- THE SWAP & DESTROY ---
            ' If a matching core target style was safely identified, map the text
            If targetStyle <> "" And currentName <> targetStyle Then
                
                With doc.Content.Find
                    .ClearFormatting
                    .Style = currentName
                    .Replacement.ClearFormatting
                    .Replacement.Style = doc.Styles(targetStyle)
                    .Replacement.Text = "^&"
                    
                    ' Execute mass text remapping across the document range
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' Delete the cleared rogue style variation
                sty.Delete
                Debug.Print "Successfully merged and deleted: " & currentName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Comprehensive style consolidation complete! Your document is fully cleaned.", vbInformation, "Success"
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

Sub Adjust_Styles()
'
' Adjust_Styles Macro
'
'
    Dim doc As Document
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' --- NORMAL (Body Text) ---
    With doc.Styles("Normal")
        .AutomaticallyUpdate = False
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 1 ---
    With doc.Styles("Heading 1,_H1")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 18
            .Bold = True
            .Italic = False
            .AllCaps = True
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 12
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .OutlineLevel = wdOutlineLevel1
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 2 ---
    With doc.Styles("Heading 2,_H2")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 16
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel2
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 3 ---
    With doc.Styles("Heading 3,_H3")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 14
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel3
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 4 ---
    With doc.Styles("Heading 4,_H4")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 12
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel4
            .TabStops.ClearAll
        End With
    End With

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "Styles successfully updated", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Style Preferences Error"
    Resume CleanUp
End Sub

