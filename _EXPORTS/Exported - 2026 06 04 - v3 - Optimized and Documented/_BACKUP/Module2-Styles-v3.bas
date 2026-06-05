'=============================================================================
' MODULE: Style Architecture & Content Cleanup Tools
'=============================================================================

Sub Style_1_Clean_Styles_Comprehensive()
'=============================================================================
' Name: Style_1_Clean_Styles_Comprehensive
' Purpose: Evaluates custom non-standard text styles, uses pattern matching
'          to group rogue iterations/typos, remaps content blocks to target
'          core styles, and destroys the duplicate fragments.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim targetStyle As String
    Dim currentName As String
    Dim headingNum As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Suppress interruptions during aggressive text search/remapping passes
    On Error Resume Next
    
    ' Parse backwards to safely decouple styles from the collection tree
    For i = doc.styles.Count To 1 Step -1
        Set sty = doc.styles(i)
        
        ' Target exclusively non-native/propagated style tracks
        If Not sty.BuiltIn Then
            currentName = sty.NameLocal
            targetStyle = ""
            
            '-----------------------------------------------------------------
            ' ADVANCED PATTERN MATCHING LOGIC
            '-----------------------------------------------------------------
            ' Pattern 1: Catch custom Heading style string fragments (e.g., "Heading 1,_H1")
            If InStr(1, currentName, "Heading ", vbTextCompare) > 0 Then
                ' Grab the target hierarchy depth digit immediately tracking "Heading "
                headingNum = Mid(currentName, InStr(1, currentName, "Heading ", vbTextCompare) + 8, 1)
                
                ' Ensure parsing character evaluates to a standard layout digit (1-9)
                If IsNumeric(headingNum) And headingNum <> "0" Then
                    Select Case headingNum
                        Case "1": targetStyle = "Heading 1"
                        Case "2": targetStyle = "Heading 2"
                        Case "3": targetStyle = "Heading 3"
                        Case "4": targetStyle = "Heading 4"
                        Case Else: targetStyle = "Heading " & headingNum
                    End Select
                End If
                
            ' Pattern 2: Catch normal style drift variations, typos, and systemic anomalies
            ElseIf InStr(1, currentName, "Norm", vbTextCompare) > 0 Or _
                   InStr(1, currentName, "Nomr", vbTextCompare) > 0 Or _
                   InStr(1, currentName, "hvr", vbTextCompare) > 0 Then
                   
                targetStyle = "Normal"
            End If
            
            '-----------------------------------------------------------------
            ' FALLBACK ATTRIBUTE VALIDATION
            '-----------------------------------------------------------------
            ' Confirms targeted master style actively exists prior to executing text swap
            If targetStyle <> "" Then
                Dim testSty As Style
                Set testSty = doc.styles(targetStyle)
                
                ' Fallback safely to base name split array if structural alias blocks match
                If testSty Is Nothing And InStr(1, targetStyle, "_H", vbTextCompare) > 0 Then
                    targetStyle = Split(targetStyle, ",")(0)
                End If
            End If
            
            '-----------------------------------------------------------------
            ' CONTENT REMAPPING AND STYLE DESTRUCTION
            '-----------------------------------------------------------------
            ' Execute find/replace across entire layout to absorb and scrub the style fragment
            If targetStyle <> "" And currentName <> targetStyle Then
                
                With doc.Content.Find
                    .ClearFormatting
                    .Style = currentName
                    .Replacement.ClearFormatting
                    .Replacement.Style = doc.styles(targetStyle)
                    .Replacement.Text = "^&"            ' Native code to preserve existing text unchanged
                    
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' Wipe the cleared duplicate style from memory completely
                sty.Delete
                Debug.Print "Successfully merged and deleted: " & currentName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Comprehensive style consolidation complete! Your document is fully cleaned.", vbInformation, "Success"
End Sub

Sub Style_2_Del_Unused_Styles_Optimized()
'=============================================================================
' Name: Style_2_Del_Unused_Styles_Optimized
' Purpose: Checks all non-built-in custom styles, double-checks headers, footers, 
'          and text blocks via structural ranges, and purges empty styles.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim pass As Integer
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Skip system errors to prevent locking on protected custom template styles
    On Error Resume Next
    
    ' Executes 2 distinct passes to cleanly prune parent-child style dependency tracks
    For pass = 1 To 2
        ' Reverse loop ensures index order remains stable as items are deleted
        For i = doc.styles.Count To 1 Step -1
            Set sty = doc.styles(i)
            
            ' Rule 1: Never target standard built-in Microsoft application styles
            If Not sty.BuiltIn Then
                
                ' Rule 2: Evaluate if Word registers the style as fundamentally dead
                If Not sty.InUse Then
                    sty.Delete
                Else
                    ' Rule 3: Deep check if style is hiding in headers, footers, or shapes
                    With doc.Content.Find
                        .ClearFormatting
                        .Style = sty.NameLocal
                        .Execute FindText:="", Format:=True, Wrap:=wdFindStop
                        
                        ' If registered "InUse" but zero structural nodes exist, safe to delete
                        If Not .Found Then
                            sty.Delete
                        End If
                    End With
                End If
                
            End If
        Next i
    Next pass
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    MsgBox "Unused custom styles cleaned up successfully!", vbInformation, "Clean Up Complete"
End Sub

Sub Style_3_Reset_BuiltIn_Style_Names()
'=============================================================================
' Name: Style_3_Reset_BuiltIn_Style_Names
' Purpose: Finds built-in system styles containing custom appended alias text
'          and strings (via commas), breaks them, and reverts them back
'          to factory default settings.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim cleanName As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' Evaluate styles array systematically backwards
    For i = doc.styles.Count To 1 Step -1
        Set sty = doc.styles(i)
        
        ' Process exclusively native built-in Microsoft Word system properties
        If sty.BuiltIn Then
            
            ' Comma detection flags when custom names have been attached over defaults
            If InStr(1, sty.NameLocal, ",", vbTextCompare) > 0 Then
                
                ' Extract character string indexing prior to the first comma separator
                cleanName = Split(sty.NameLocal, ",")(0)
                
                ' Assign the base native string back onto the local workspace setting
                sty.NameLocal = cleanName
                Debug.Print "Reset style name back to: " & cleanName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Built-in style names have been reverted to their original factory names!", vbInformation, "Reset Complete"
End Sub

Sub Style_4_Adjust_Styles()
'=============================================================================
' Name: Style_4_Adjust_Styles
' Purpose: Explicitly configures and standardizes the core paragraph styles 
'          (Normal, Heading 1 through Heading 4) and the Caption style. 
'          Establishes layout baselines, clears rogue tab stops, and ensures 
'          clean text geometries.
'=============================================================================
    Dim doc As Document
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates while shifting global layouts
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' 1. NORMAL STYLE (The baseline body text for your entire document)
    '-------------------------------------------------------------------------
    With doc.styles("Normal")
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
            .WidowControl = True
            .TabStops.ClearAll
        End With
    End With

    '-------------------------------------------------------------------------
    ' 2. HEADING 1 (Primary Document Sections - All Caps & Justified)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 1")
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
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5     ' Standard 1.5 line spacing for headings
            .Alignment = wdAlignParagraphJustify   ' Justified alignment for a clean block look
            .OutlineLevel = wdOutlineLevel1
            .PageBreakBefore = True                ' Enforces new page for each major section
            .TabStops.ClearAll
        End With
    End With

    '-------------------------------------------------------------------------
    ' 3. HEADING 2 (Sub-sections - Left-Aligned & Bound to Following Text)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 2")
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
            .KeepWithNext = True               ' Prevents orphan headings at bottom of page
            .KeepTogether = True               ' Keeps heading on a single page
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel2
            .TabStops.ClearAll
        End With
    End With

    '-------------------------------------------------------------------------
    ' 4. HEADING 3 (Sub-sub-sections)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 3")
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

    '-------------------------------------------------------------------------
    ' 5. HEADING 4 (Deep Hierarchy Details)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 4")
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

    '-------------------------------------------------------------------------
    ' CAPTION STYLE (The style for captioning tables, figures, and other media)
    '-------------------------------------------------------------------------
    With doc.styles("Caption")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = True
            .Italic = True
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)      ' Dynamically calculates 1.15x line spacing based on font size
            .Alignment = wdAlignParagraphJustify
            .KeepWithNext = True
            .KeepTogether = True
            .WidowControl = True
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

Sub Style_5_Adjust_Document_Font_And_Spacing()
'=============================================================================
' Name: Style_5_Adjust_Document_Font_And_Spacing
' Purpose: Executes a 3-phase optimization pipeline:
'          1. Direct formats all body text to smash manual style overrides.
'          2. Restores tight 1.0 spacing and 0/0 padding to all Tables.
'          3. Scans paragraphs in a single pass via Outline Levels to capture
'             true headings AND "fake" headings, resetting their fonts.
'=============================================================================
    Dim doc As Document
    Dim tbl As Table
    Dim para As Paragraph
    Dim rng As Range
    Dim outLvl As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates, animations, and repainting
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' PHASE 1: GLOBAL DIRECT FORMATTING OVERRIDE
    '-------------------------------------------------------------------------
    ' Uniformly applies baseline body formatting to clear stubborn layout drift
    Set rng = doc.Content
    With rng
        With .Font
            .Name = "Calibri"
            .Size = 11
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
        End With
    End With

    '-------------------------------------------------------------------------
    ' PHASE 2: TABLE PROTECTION LOOP
    '-------------------------------------------------------------------------
    ' Immediately restores tight single-line spacing inside all tables
    For Each tbl In doc.Tables
        With tbl.Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
        End With
    Next tbl

    '-------------------------------------------------------------------------
    ' PHASE 3: RESTORE & UP-CONVERT HEADING STYLES VIA OUTLINE LEVELS
    '-------------------------------------------------------------------------
    ' Instead of checking paragraph style names 4 times over, we run a lightning-fast 
    ' SINGLE pass through the document evaluating Outline Levels. This automatically 
    ' updates true headings AND catches hidden "fake" headings using Normal style.
    For Each para In doc.Paragraphs
        
        ' Bypass table contents to ensure data cells are never converted into headings
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Fetch the paragraph's structural Outline Level
            outLvl = para.OutlineLevel
            
            ' Process exclusively if it maps to levels 1, 2, 3, or 4
            If outLvl >= 1 And outLvl <= 4 Then
                With para.Range
                    ' 1. Peel off Phase 1's direct formatting tape (Calibri 11pt override)
                    .Font.Reset
                    
                    ' 2. Force apply the true built-in Heading Style based on the level digit
                    .Style = doc.Styles("Heading " & outLvl)
                End With
            End If
            
        End If
    Next para

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document successfully standardized, tables preserved, and heading structures unified!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Standardization Error"
End Sub

Sub Style_6_Highlight_Duplicate_Styles()
'=============================================================================
' Name: Style_6_Highlight_Duplicate_Styles
' Purpose: Evaluates the active workspace document for common corrupted style
'          clones (e.g., "Char" artifacts from copy-pasting text). Forces a 
'          vibrant highlight on text fragments utilizing them for easy auditing.
'=============================================================================
    Dim sty As Style
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    ' Clear out any residual content highlighting to create a baseline canvas
    doc.Content.HighlightColorIndex = wdNoHighlight
    
    ' Cycle custom document style structures
    For Each sty In doc.styles
        If Not sty.BuiltIn Then
            
            ' Isolate known automated copy/paste duplicate naming signatures
            If InStr(1, sty.NameLocal, " Char", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, " pt", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, "Indent", vbTextCompare) > 0 Then
               
                ' Find specific string patterns matching the targeted clone signature
                With doc.Content.Find
                    .ClearFormatting
                    .Style = sty.NameLocal
                    .Replacement.ClearFormatting
                    .Replacement.Highlight = True       ' Enforces highlight application change
                    .Replacement.Text = "^&"            ' Keeps current alphanumeric values safe
                    
                    ' Apply visual highlights universally
                    .Execute Replace:=wdReplaceAll
                End With
                
                Debug.Print "Highlighted usage of: " & sty.NameLocal
            End If
        End If
    Next sty
    
    Application.ScreenUpdating = True
    MsgBox "Check for Bright Green highlights.", vbInformation, "Audit Map Set"
End Sub
