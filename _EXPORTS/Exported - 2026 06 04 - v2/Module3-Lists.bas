Attribute VB_Name = "Module3"
Sub Lists_1_Build_Multi_Level_List_for_Headings()
' Optimized Multi-Level List linked explicitly to Heading Styles
    Dim doc As Document
    Dim LT As ListTemplate
    Dim lvl As Integer
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Create a clean, separate list template in this document
    Set LT = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' --- LEVEL 1 (Linked to Heading 1) ---
    With LT.ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.25)
        .TabPosition = InchesToPoints(0.25)
        .StartAt = 1
        .LinkedStyle = "Heading 1"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Italic = False
    End With
    
    ' --- LEVEL 2 (Linked to Heading 2) ---
    With LT.ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = InchesToPoints(0.5)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = "Heading 2"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Italic = False
    End With
    
    ' --- LEVEL 3 (Linked to Heading 3) ---
    With LT.ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.75)
        .TabPosition = InchesToPoints(0.75)
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = "Heading 3"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Italic = False
    End With
    
    ' --- LEVEL 4 (Linked to Heading 4) ---
    With LT.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1)
        .TabPosition = InchesToPoints(1)
        .ResetOnHigher = 3
        .StartAt = 1
        .LinkedStyle = "Heading 4"
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Italic = False
    End With
    
    ' --- LEVELS 5 to 9 (Unlinked deep levels processed cleanly via loop) ---
    Dim formats As Variant, styles As Variant
    Dim numPos As Variant, textPos As Variant
    
    formats = Array("(%5)", "(%6)", "%7.", "%8.", "%9.")
    styles = Array(wdListNumberStyleLowercaseLetter, wdListNumberStyleLowercaseRoman, _
                   wdListNumberStyleArabic, wdListNumberStyleLowercaseLetter, wdListNumberStyleLowercaseRoman)
    numPos = Array(1, 1.25, 1.5, 1.75, 2)
    textPos = Array(1.25, 1.5, 1.75, 2, 2.25)
    
    For lvl = 5 To 9
        With LT.ListLevels(lvl)
            .NumberFormat = formats(lvl - 5)
            .NumberStyle = styles(lvl - 5)
            .TrailingCharacter = wdTrailingSpace
            .Alignment = wdListLevelAlignLeft
            .NumberPosition = InchesToPoints(numPos(lvl - 5))
            .TextPosition = InchesToPoints(textPos(lvl - 5))
            .ResetOnHigher = lvl - 1
            .StartAt = 1
            .LinkedStyle = "" ' Intentionally blank for deep levels
            .Font.Name = "Calibri"
        End With
    Next lvl
    
    ' Refresh document styles to apply the numbering setup automatically
    Dim s As Style
    On Error Resume Next
    For lvl = 1 To 4
        Set s = doc.styles("Heading " & lvl)
        s.LinkToListTemplate ListTemplate:=LT, ListLevelNumber:=lvl
    Next lvl
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    MsgBox "Heading multi-level list generated and linked successfully!", vbInformation, "Success"
End Sub

Sub Lists_2_Build_Multi_Levels_List_for_List_Of_Volumes()
' Optimized List Creation for Volumes (Levels 4-9 set to Standard Arabic)
    Dim doc As Document
    Dim LT As ListTemplate
    Dim lvl As Integer
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Create an isolated, new list template inside this specific document
    Set LT = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' --- LEVEL 1 ---
    With LT.ListLevels(1)
        .NumberFormat = "Vol-%1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.6)
        .TabPosition = InchesToPoints(0.6)
        .StartAt = 1
        .LinkedStyle = ""
    End With
    
    ' --- LEVEL 2 ---
    With LT.ListLevels(2)
        .NumberFormat = "Vol-%1-%2:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseLetter
        .NumberPosition = InchesToPoints(0.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1.4)
        .TabPosition = InchesToPoints(1.4)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = ""
    End With
    
    ' --- LEVEL 3 ---
    With LT.ListLevels(3)
        .NumberFormat = "%3)"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(1.4)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1.7)
        .TabPosition = InchesToPoints(1.7)
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = ""
    End With
    
    ' --- LEVELS 4 to 9 (Standard Arabic Numbers: 4., 5., 6., etc.) ---
    For lvl = 4 To 9
        With LT.ListLevels(lvl)
            .NumberFormat = "%" & lvl & "."          ' Yields "4.", "5.", etc.
            .NumberStyle = wdListNumberStyleArabic   ' Enforces 1, 2, 3 format
            .TrailingCharacter = wdTrailingSpace
            .Alignment = wdListLevelAlignLeft
            
            ' Scale indents naturally as the levels go deeper
            .NumberPosition = InchesToPoints(0.25 * (lvl - 1))
            .TextPosition = InchesToPoints(0.25 * lvl)
            
            .ResetOnHigher = lvl - 1
            .StartAt = 1
            .LinkedStyle = ""
        End With
    Next lvl
    
    ' Apply the newly built template safely to the user's active selection
    Selection.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=LT, _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
        
    Application.ScreenUpdating = True
End Sub

Sub Lists_3_Reset_List_Numbering_Fonts()
'=============================================================================
' Name: Lists_3_Reset_List_Numbering_Fonts
' Purpose: Resets list numbering fonts back to default parent styles while
'          safely skipping bullet lists to prevent "missing character" symbols.
'=============================================================================
    Dim templ As ListTemplate
    Dim lev As ListLevel
    
    ' Speed optimization: Prevent screen flickering during deep object looping
    Application.ScreenUpdating = False
    
    ' Enable defensive error skipping to bypass protected or corrupted system list slots
    On Error Resume Next
    
    ' Loop through every list template container in the document
    For Each templ In ActiveDocument.ListTemplates
        
        ' Loop through all 9 possible levels inside the current template
        For Each lev In templ.ListLevels
            
            ' CRITICAL CHECK: Only reset if the list level is NOT using a bullet style.
            ' This prevents symbols from being wiped out or turned into broken squares.
            If lev.NumberStyle <> wdListNumberStyleBullet Then
                
                ' Strip out manual font overrides for numeric lists
                lev.Font.Reset
                
            End If
            
        Next lev
    Next templ
    
    ' Restore standard error tracking and turn screen updating back on
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "All numeric list fonts have been reset safely. Bullet symbols preserved!", vbInformation, "Reset Complete"
End Sub
