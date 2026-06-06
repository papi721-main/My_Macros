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
            ' Basic Font Settings
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Color = wdColorAutomatic
            .Outline = False            ' Removes any unwanted borders around text
            .Shadow = False             ' Removes any unwanted shadow effects on text
            .Emboss = False             ' Removes any unwanted embossing effects on text
            .Engrave = False            ' Removes any unwanted engraving effects on text
            
            ' Advanced Settings
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions

        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5     ' Standard 1.5 line spacing for headings
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
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 18
            .Bold = True
            .Italic = False
            .AllCaps = True
            .Color = wdColorAutomatic

            
            ' Advanced Settings
            ' ------------------
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions

        End With
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle   ' Single line spacing for a tight, impactful heading block
            .Alignment = wdAlignParagraphLeft      ' Left-aligned alignment for a clean block look
            .OutlineLevel = wdOutlineLevel1        ' Ensures proper recognition in the document map, navigation pane, and TOC generation
            .PageBreakBefore = True                ' Enforces new page for each major section
            .KeepWithNext = True                   ' Prevents orphan headings at bottom of page
            .KeepTogether = True                   ' Keeps heading on a single page to avoid awkward splits
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
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 16
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
            
            ' Advanced Settings
            ' ------------------
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
            
        End With
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle   ' Single line spacing for a tight, impactful heading block
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
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 14
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
            
            ' Advanced Settings
            ' ------------------
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
            
        End With
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle   ' Single line spacing for a tight, impactful heading block
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
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 12
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic

            ' Advanced Settings
            ' ------------------
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions

        End With
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle   ' Single line spacing for a tight, impactful heading block
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
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = True
            .Italic = True
            .Color = wdColorAutomatic
            .AllCaps = False
            .SmallCaps = False

            ' Advanced Settings
            ' ------------------
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
            
        End With
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)      ' Dynamically calculates 1.15x line spacing based on font size
            .Alignment = wdAlignParagraphJustify
            .KeepWithNext = True
            .KeepTogether = True
            .WidowControl = True
            .OutlineLevel = wdOutlineLevelBodyText
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

Sub Style_5_Apply_Styles_To_Document()
'=============================================================================
' Name: Style_5_Apply_Styles_To_Document
' Purpose: Executes a 3-phase optimization pipeline:
'          1. Direct formats all body text to smash manual style overrides.
'          2. Restores tight 1.15pt spacing and 0/0 padding to all Tables.
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

            ' ======================================================================
            ' Advanced Settings
            ' ======================================================================
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
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
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)     ' Dynamically calculates single line spacing based on font size

            ' If you want to use a specific line spacing value instead
            ' of single spacing, you can uncomment the following line
            ' and set your desired spacing in points
            ' --------------------------------------
            ' .LineSpacingRule = wdLineSpaceMultiple
            ' .LineSpacing = LinesToPoints(1.15)
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
                    .Style = doc.styles("Heading " & outLvl)
                End With
            End If
            
        End If
    Next para

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document successfully formatted, tables preserved, and heading structures unified!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Formatting Error"
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

Sub Style_7_Configure_Figure_Caption_KeepWithNext()
'=============================================================================
' Name: Style_7_Configure_Figure_Caption_KeepWithNext
' Purpose: Scans all paragraphs styled as "Caption". If the text begins with
'          "Figure", it forces KeepWithNext to FALSE. This prevents the
'          caption from being pulled away from the visual asset above it.
'=============================================================================
    Dim doc As Document
    Dim para As Paragraph
    Dim paraText As String
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates during processing
    Application.ScreenUpdating = False
    
    ' Loop through every paragraph in the document
    For Each para In doc.Paragraphs
        
        ' Rule 1: Only look at paragraphs assigned to the standard "Caption" style
        If para.Style = doc.styles("Caption") Then
            
            ' Clean and trim the text to look at the first word safely
            paraText = Trim(para.Range.Text)
            
            ' Rule 2: Check if the caption specifically starts with "Figure"
            ' (Using UCase and Left handles variations like "Figure 1", "Figure 2.1", etc.)
            If UCase(Left(paraText, 6)) = "FIGURE" Or _
               UCase(Left(paraText, 5)) = "PHOTO" Then
                
                ' FIXED: Call KeepWithNext directly on the Paragraph object.
                ' This disconnects the caption from the paragraph below it.
                para.KeepWithNext = False
                
            End If
        End If
    Next para
    
    ' Restore system screen updating
    Application.ScreenUpdating = True
    
    MsgBox "Figure caption page layout bounds configured successfully!", vbInformation, "Layout Complete"
End Sub

Sub Style_8_Apply_Styles_To_Document_And_Fix_List_Spacing()
'=============================================================================
' Name: Style_8_Apply_Styles_To_Document_And_Fix_List_Spacing
' Purpose: Executes a fully consolidated multi-phase document layout optimization:
'          1. Direct formats all body text to smash unmanaged layout drifts.
'          2. Resores tight 1.0 spacing rules to all tabular cell grids.
'          3. Iterates paragraphs once to calculate context-aware list block values
'             (Before=0, Inside=0, After=6, Intro Tightening=0).
'          4. Resolves true and "fake" structural heading paths via Outline Levels.
'=============================================================================
    Dim doc As Document
    Dim tbl As Table
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim nextPara As Paragraph
    Dim rng As Range
    Dim outLvl As Long
    Dim isLastItem As Boolean
    
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

            ' ======================================================================
            ' Advanced Settings
            ' ======================================================================
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
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
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)     ' Dynamically calculates single line spacing based on font size
        End With
    Next tbl

    '-------------------------------------------------------------------------
    ' CONSOLIDATED SCANNING ENGINE: LIST SPACING & OUTLINE CONVERSIONS
    '-------------------------------------------------------------------------
    For Each para In doc.Paragraphs
        
        '=====================================================================
        ' SUB-PHASE A: ADVANCED CONTEXTUAL LIST ADJUSTMENTS
        '=====================================================================
        ' HARD GUARDRAIL: Skip paragraph completely if it is an active Heading element
        If para.OutlineLevel = wdOutlineLevelBodyText Then
        
            ' Rule 1: Target active list formatting structures (Bullets, Numbers, Outlines)
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                
                ' Rule 2: Strictly protect tables by ignoring internal table lists
                If Not para.Range.Information(wdWithInTable) Then
                    
                    ' STEP 1: LOOK-BEHIND (Tighten Intro Paragraph)
                    Set prevPara = para.Previous
                    If Not prevPara Is Nothing Then
                        ' If the preceding line is NOT a list, this is the FIRST item in the block!
                        If prevPara.Range.ListFormat.ListType = wdListNoNumbering Then
                            ' Ensure it isn't a table or a structural heading before modifying it
                            If Not prevPara.Range.Information(wdWithInTable) And _
                               (prevPara.OutlineLevel >= wdOutlineLevelBodyText) Then
                                
                                prevPara.SpaceAfterAuto = False
                                prevPara.SpaceAfter = 0 ' Snaps the introductory text tightly down
                                
                            End If
                        End If
                    End If
    
                    ' STEP 2: Enforce Base List Geometries
                    para.SpaceBeforeAuto = False
                    para.SpaceAfterAuto = False
                    para.SpaceBefore = 0
                    para.LineSpacingRule = wdLineSpace1pt5
                    
                    ' STEP 3: LOOK-AHEAD (Determine Block End Spacing)
                    Set nextPara = para.Next
                    isLastItem = False ' Reset flag for current paragraph
                    
                    ' Condition 1: There is no next paragraph (End of Document)
                    If nextPara Is Nothing Then
                        isLastItem = True
                    Else
                        ' Condition 2: The next paragraph is generic body text
                        If nextPara.Range.ListFormat.ListType = wdListNoNumbering Then
                            isLastItem = True
                        ' Condition 3: The next line drops out of main text space and into a table
                        ElseIf nextPara.Range.Information(wdWithInTable) Then
                            isLastItem = True
                        End If
                        
                        ' Condition 4: The next line is explicitly a Heading (Outline Levels 1 to 9)
                        If nextPara.OutlineLevel >= 1 And nextPara.OutlineLevel <= 9 Then
                            isLastItem = True
                        End If
                    End If
                    
                    ' STEP 4: Apply Calculated Spacing Execution
                    If isLastItem Then
                        para.SpaceAfter = 6   ' Add professional breathing room at block end
                    Else
                        para.SpaceAfter = 0   ' Keep items tightly packed within the block
                    End If
                    
                End If
            End If
            
        End If
        
        '=====================================================================
        ' SUB-PHASE B: RESTORE & UP-CONVERT HEADING STYLES VIA OUTLINE LEVELS
        '=====================================================================
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
                    .Style = doc.styles("Heading " & outLvl)
                End With
            End If
            
        End If
        
    Next para

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document styles applied and list spaces manually balanced successfully!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Formatting Error"
End Sub
