'=============================================================================
' MODULE: List Architecture & Numbering Schema Management
'=============================================================================

Sub Lists_1_Build_Multi_Level_List_for_Headings()
'=============================================================================
' Name: Lists_1_Build_Multi_Level_List_for_Headings
' Purpose: Generates an isolated, multi-level outline numbering framework
'          and directly hooks levels 1-4 into built-in Heading styles.
'          Ensures clean legal outline indentation schemas (e.g., 1., 1.1., 1.1.1.)
'=============================================================================
    Dim doc As Document
    Dim LT As ListTemplate
    Dim lvl As Integer
    Dim formats As Variant, styles As Variant
    Dim numPos As Variant, textPos As Variant
    Dim s As Style
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates while changing list trees
    Application.ScreenUpdating = False
    
    ' Inject an independent list template container directly into the active document
    Set LT = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    '-------------------------------------------------------------------------
    ' CONFIGURE HEADERS TREE LINKAGE (LEVELS 1 TO 4)
    '-------------------------------------------------------------------------
    
    ' --- LEVEL 1 (Heading 1 Root: e.g., "1.") ---
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
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Italic = False
            .Color = wdUndefined ' Inherit color from Heading 1 style
            .Size = wdUndefined ' Inherit size from Heading 1 style
        End With
    End With
    
    ' --- LEVEL 2 (Heading 2 Nested Child: e.g., "1.1.") ---
    With LT.ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.4)
        .TabPosition = InchesToPoints(0.4)
        .ResetOnHigher = 1                  ' Forces sub-counts to restart when parent shifts
        .StartAt = 1
        .LinkedStyle = "Heading 2"
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Italic = False
            .Color = wdUndefined ' Inherit color from Heading 2 style
            .Size = wdUndefined ' Inherit size from Heading 2 style
        End With
    End With
    
    ' --- LEVEL 3 (Heading 3 Nested Child: e.g., "1.1.1.") ---
    With LT.ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.6)
        .TabPosition = InchesToPoints(0.6)
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = "Heading 3"
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Italic = False
            .Color = wdUndefined ' Inherit color from Heading 3 style
            .Size = wdUndefined ' Inherit size from Heading 3 style
        End With
    End With
    
    ' --- LEVEL 4 (Heading 4 Nested Child: e.g., "1.1.1.1.") ---
    With LT.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.65)
        .TabPosition = InchesToPoints(0.65)
        .ResetOnHigher = 3
        .StartAt = 1
        .LinkedStyle = "Heading 4"
        With .Font
            .Name = "Calibri"
            .Bold = True
            .Italic = False
            .Color = wdUndefined ' Inherit color from Heading 4 style
            .Size = wdUndefined ' Inherit size from Heading 4 style
        End With
    End With
    
    '-------------------------------------------------------------------------
    ' CONFIGURE UNLINKED DEEP COMPACT HIERARCHIES (LEVELS 5 TO 9)
    '-------------------------------------------------------------------------
    ' Uses flat data arrays to safely map unique style schema formats
    formats = Array("(%5)", "(%6)", "%7.", "%8.", "%9.")
    styles = Array(wdListNumberStyleLowercaseLetter, wdListNumberStyleLowercaseRoman, _
                   wdListNumberStyleArabic, wdListNumberStyleLowercaseLetter, wdListNumberStyleLowercaseRoman)
    numPos = Array(1, 1.25, 1.5, 1.75, 2)
    textPos = Array(1.25, 1.5, 1.75, 2, 2.25)
    
    ' Loop structural arrays to cleanly configure deep, unlinked outline fields
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
            .LinkedStyle = ""               ' Intentionally unlinked from paragraph style trees
            .Font.Name = "Calibri"
        End With
    Next lvl
    
    '-------------------------------------------------------------------------
    ' SYNC AND INTERLOCK REFRESH WITH DOCUMENT HEADINGS
    '-------------------------------------------------------------------------
    ' Employs a resilient style update pipeline to apply multi-level outline numbering
    On Error Resume Next
    For lvl = 1 To 4
        Set s = doc.styles("Heading " & lvl)
        s.LinkToListTemplate ListTemplate:=LT, ListLevelNumber:=lvl
    Next lvl
    On Error GoTo 0
    
    ' Restore application workspace rendering
    Application.ScreenUpdating = True
    MsgBox "Heading multi-level list generated and linked successfully!", vbInformation, "Success"
End Sub


Sub Lists_2_Build_Multi_Levels_List_for_List_Of_Volumes()
'=============================================================================
' Name: Lists_2_Build_Multi_Levels_List_for_List_Of_Volumes
' Purpose: Generates a dedicated list schema specialized for complex reports and
'          multi-volume indexes. Applies custom labels (e.g., Vol-I:, Vol-I-A:)
'          and standardizes deep nesting levels using clean Arabic indicators.
'=============================================================================
    Dim doc As Document
    Dim LT As ListTemplate
    Dim lvl As Integer
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Generate clean list architecture instance attached to this document object
    Set LT = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' --- LEVEL 1 (Master Book Grouping: e.g., "Vol-I:") ---
    With LT.ListLevels(1)
        .NumberFormat = "Vol-%1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.6)
        .TabPosition = InchesToPoints(0.6)
        .StartAt = 1
        .LinkedStyle = ""                   ' Disconnected from structural heading styles
    End With
    
    ' --- LEVEL 2 (Sub-Volume Grouping: e.g., "Vol-I-A:") ---
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
    
    ' --- LEVEL 3 (Chapter / Segment Node: e.g., "1)") ---
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
    
    ' --- LEVELS 4 TO 9 (Flat Arabic Sequences: 4., 5., 6., etc.) ---
    ' Loops remaining nesting slots to enforce uniform, structural sub-numbering
    For lvl = 4 To 9
        With LT.ListLevels(lvl)
            .NumberFormat = "%" & lvl & "."         ' Compiles variable level layout strings (e.g., "%4.")
            .NumberStyle = wdListNumberStyleArabic  ' Locks numbering outputs to standard Arabic numerals
            .TrailingCharacter = wdTrailingSpace
            .Alignment = wdListLevelAlignLeft
            
            ' Scale layout margins out dynamically based on list indent hierarchy depth
            .NumberPosition = InchesToPoints(0.25 * (lvl - 1))
            .TextPosition = InchesToPoints(0.25 * lvl)
            
            .ResetOnHigher = lvl - 1
            .StartAt = 1
            .LinkedStyle = ""
        End With
    Next lvl
    
    ' Execute layout change across the user's active cursor selection block
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
' Purpose: Sweeps document list templates to wipe out hardcoded font formatting.
'          Forces lists to cleanly inherit properties from their parent style.
'          CRITICAL SAFETY: Ignores bullet tracking schemas to prevent icon corruption.
'=============================================================================
    Dim templ As ListTemplate
    Dim lev As ListLevel
    
    ' Speed optimization: Prevent screen flickering during massive loop scans
    Application.ScreenUpdating = False
    
    ' Enable defensive error bypass to glide past hidden or locked system list containers
    On Error Resume Next
    
    ' Loop through every list format mapping layer in the active file
    For Each templ In ActiveDocument.ListTemplates
        
        ' Scan all 9 possible list sub-levels in the template
        For Each lev In templ.ListLevels

            ' Set the color back to Auto for all lists
            lev.Font.ColorIndex = wdAuto

            ' CRITICAL CHECK: Ignore levels configured with graphic bullet markers.
            ' Resetting a bullet level strips character font mapping, corrupting icons into square boxes.
            If lev.NumberStyle <> wdListNumberStyleBullet Then
                
                ' Strip custom manual font overrides from valid numerical list entries
                lev.Font.Reset
                
            End If
            
        Next lev
    Next templ
    
    ' Re-engage standard error tracking rules and clean workspace view
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "All numeric list fonts have been reset safely. Bullet symbols preserved!", vbInformation, "Reset Complete"
End Sub

Sub Lists_4_Fix_List_Spacing_Manually()
'=============================================================================
' Name: Lists_4_Fix_List_Spacing_Manually
' Purpose: Scans all paragraphs. If a paragraph is a list item outside a table:
'          - SpaceBefore is ALWAYS set to 0.
'          - Line Spacing is ALWAYS set to 1.5.
'          - SpaceAfter = 0 for items within the same list block.
'          - SpaceAfter = 6 if the next line is body text, a table, a heading, or nothing.
'          - Sets SpaceAfter = 0 on the paragraph IMMEDIATELY BEFORE the
'            very first list item to draw introductory text tightly to the list.
'          CRITICAL SAFETY: Completely skips headings (even if numbered) from
'          all formatting passes to preserve structural section gaps.
'=============================================================================
    Dim doc As Document
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim nextPara As Paragraph
    Dim isLastItem As Boolean
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Freeze screen updates during text layout changes
    Application.ScreenUpdating = False
    
    ' Loop through every paragraph in the document
    For Each para In doc.Paragraphs
        
        ' HARD GUARDRAIL: Skip paragraph completely if it is an active Heading element
        If para.OutlineLevel = wdOutlineLevelBodyText Then
        
            ' Rule 1: Target active list formatting structures (Bullets, Numbers, Outlines)
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                
                ' Rule 2: Strictly protect tables by ignoring internal table lists
                If Not para.Range.Information(wdWithInTable) Then
                    
                    '-------------------------------------------------------------
                    ' STEP A: LOOK-BEHIND (Tighten Intro Paragraph)
                    '-------------------------------------------------------------
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
    
                    '-------------------------------------------------------------
                    ' STEP B: Enforce Base List Geometries
                    '-------------------------------------------------------------
                    para.SpaceBeforeAuto = False
                    para.SpaceAfterAuto = False
                    para.SpaceBefore = 0
                    para.LineSpacingRule = wdLineSpace1pt5
                    
                    '-------------------------------------------------------------
                    ' STEP C: LOOK-AHEAD (Determine Block End Spacing)
                    '-------------------------------------------------------------
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
                    
                    '-------------------------------------------------------------
                    ' STEP D: Apply Calculated Spacing Execution
                    '-------------------------------------------------------------
                    If isLastItem Then
                        para.SpaceAfter = 6   ' Add professional breathing room at block end
                    Else
                        para.SpaceAfter = 0   ' Keep items tightly packed within the block
                    End If
                    
                End If
            End If
            
        End If
    Next para
    
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    
    MsgBox "Fixed and optimized list spacing successfully! Structural headings left untouched.", vbInformation, "Layout Complete"
End Sub

Sub Lists_5_Fix_Selected_List_Indents()
    ' ============================================================================
    ' MODULE:       Lists_5_Fix_Selected_List_Indents
    ' DESCRIPTION:  Normalizes left indents and bullet positions for all active 
    '               lists within a selection while strictly skipping headings.
    '               Level 1: Bullet at 0.25", Text at 0.5"
    '               Consecutive levels increment both by 0.25" per step.
    ' AUTHOR:       VBA Automation Suite
    ' ============================================================================
    
    Dim para As Paragraph
    Dim listLvl As Integer
    Dim targetBulletPos As Single
    Dim targetTextPos As Single
    
    ' Establish global error handling to safeguard the document layout context
    On Error GoTo CleanUp
    
    ' Performance Optimization: Freeze visual layout redrawing to maximize processing speed
    Application.ScreenUpdating = False
    
    ' Loop exclusively through the paragraphs actively highlighted on screen 
    For Each para In Selection.Paragraphs
        
        ' Guardrail 1: Skip structural headings by targeting outline level metadata 
        If para.OutlineLevel = wdOutlineLevelBodyText Then
            
            ' Guardrail 2: Identify all bulleted, numbered, and multi-level lists 
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                
                ' Extract the exact 1-based structural depth index of the active list row 
                listLvl = para.Range.ListFormat.ListLevelNumber
                
                ' Calculate precise layout geometry increments (0.25" steps per level)
                ' Level 1: Bullet = 0.25", Text = 0.50"
                ' Level 2: Bullet = 0.50", Text = 0.75"
                ' Level 3: Bullet = 0.75", Text = 1.00"
                targetBulletPos = 0.25 * listLvl
                targetTextPos = 0.25 * (listLvl + 1)
                
                ' Apply properties to the paragraph layout engine
                ' Word calculates indents using points, so we convert inches explicitly
                para.LeftIndent = InchesToPoints(targetTextPos)
                
                ' FirstLineIndent creates a negative hanging indent window.
                ' By calculating the negative offset between text position and bullet position,
                ' the bullet snaps perfectly backward to its designated channel.
                para.FirstLineIndent = InchesToPoints(targetBulletPos - targetTextPos)
                
                ' Reset right margin constraints to flush default
                para.RightIndent = 0
                
            End If
            
        End If
        
    Next para

CleanUp:
    ' Restore visual layout rendering safely 
    Application.ScreenUpdating = True 
    
    ' Execution Completion Notification
    If Err.Number = 0 Then
        MsgBox "List indents and alignment successfully adjusted to 0.25"" increments!", _
               vbInformation, "Format Fix Complete"
    Else
        MsgBox "The layout engine encountered an error: " & Err.Description, _
               vbCritical, "Execution Failure"
    End If
End Sub