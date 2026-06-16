Sub Misc_1_Clear_All_Highlighting_Globally()
    ' =========================================================================
    ' MODULE NAME:  Misc_1_Clear_All_Highlighting_Globally
    ' PURPOSE:      Completely clears all background text highlighting across every
    '               structural layer of the active document, including stubborn 
    '               unlinked headers/footers.
    ' SCOPE:        Main document text, tables, headers, footers, textboxes, footnotes.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Employs a two-pass strategy combining background story loops
    '               with explicit section audits. Uses ScreenUpdating control
    '               to bypass the physical cursor, avoiding screen flicker.
    ' =========================================================================
    
    Dim doc As Document
    Dim sec As Section
    Dim compView As HeaderFooter
    Dim story As Range

    Set doc = ActiveDocument
    
    ' Disable screen updates to freeze visual page re-pagination.
    ' This suppresses application stuttering and drastically increases 
    ' background processing speed on long, multi-section documents[cite: 847, 848].
    Application.ScreenUpdating = False
    
    ' =========================================================================
    ' PASS 1: THE GLOBAL SWEEP (Catches Main Text, Tables, and Footnotes)
    ' =========================================================================
    ' This loop scans the background StoryRanges collection. It is the most 
    ' efficient way to bypass cursor selection and hit major text blocks[cite: 5, 17].
    For Each story In doc.StoryRanges
        Do
            ' Set the highlight index directly to wdNoHighlight.
            ' This strips the background formatting layer without touching the text text[cite: 4].
            story.HighlightColorIndex = wdNoHighlight
            
            ' Word chunks text into sub-ranges (e.g., linked text boxes or split footnotes).
            ' NextStoryRange ensures the pointer evaluates downstream links in this story[cite: 5].
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    For Each story In doc.StoryRanges
    
    ' =========================================================================
    ' PASS 2: DEEP SECTION PENETRATION (Forces dormant headers/footers awake)
    ' =========================================================================
    ' ARCHITECTURAL CULPRIT: Word does not automatically load every header and 
    ' footer layer into memory unless they are actively visible or opened by the user[cite: 26].
    ' If "Different First Page" or "Different Odd & Even Pages" layout flags are enabled, 
    ' unlinked header/footer canvases remain completely dormant in the background cache[cite: 27].
    ' Because Pass 1 loops right past dormant stories[cite: 28], Pass 2 explicitly declares a 
    ' nested section loop to force the layout processor to evaluate all layout sub-layers[cite: 29, 31].
    For Each sec In doc.Sections
        
        ' -----------------------------------------------------------------
        ' Sub-Phase A: Section Headers
        ' -----------------------------------------------------------------
        ' Step through the section's Header collection (First Page, Even Pages, Primary).
        For Each compView In sec.Headers
            ' The .Exists safety gate prevents Word from throwing a runtime error
            ' if the sub-layer is structurally unassigned or inactive[cite: 34].
            If compView.Exists Then
                ' Route through the sub-layer's .Range to expose its underlying text canvas.
                compView.Range.HighlightColorIndex = wdNoHighlight
            End If
        Next compView
        
        ' -----------------------------------------------------------------
        ' Sub-Phase B: Section Footers
        ' -----------------------------------------------------------------
        ' Step through the section's Footer collection (First Page, Even Pages, Primary).
        For Each compView In sec.Footers
            If compView.Exists Then
                ' Force a structural layout reset on any active footer graphics or text ranges.
                compView.Range.HighlightColorIndex = wdNoHighlight
            End If
        Next compView
        
    Next sec
    
    ' Re-enable screen rendering to display the finalized layout updates [cite: 1376]
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "All text highlighting has been forcefully stripped from all layers, including headers and footers.", _
           vbInformation, "Global Clear Successful"
End Sub

Sub Misc_2_Highlight_Target_Words()
    ' =========================================================================
    ' MODULE NAME:  Misc_2_Highlight_Target_Words
    ' PURPOSE:      Searches the main body text layer for an array of target keywords
    '               and applies a standardized green highlight overlay to them[cite: 4].
    ' SCOPE:        Main text body and embedded tables ONLY[cite: 4, 8]. Excludes headers/footers[cite: 4, 8].
    ' SETTINGS:     Case-insensitive match[cite: 13, 22]. Evaluates word fragments (e.g. finds "fig" inside "figure").
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Uses Word's native high-speed Find engine via background ranges[cite: 9, 17]. 
    '               By isolating the scope to the main canvas, it targets tables 
    '               seamlessly without needing intensive cell-by-cell loops.
    ' =========================================================================
    
    Dim docRange As Range
    Dim wordList As Variant
    Dim targetWord As Variant
    
    ' EXPLICIT CONFIGURATION: Define your target word list here
    ' Expand this array as needed to scale the list dynamically over time[cite: 9, 45].
    wordList = Array("tab", "fig", "annex", "plate") [cite: 3, 4]
    
    ' Freeze visual application repagination to eliminate macro lag.
    ' This suppresses display stuttering and drastically cuts processing overhead[cite: 65, 178].
    Application.ScreenUpdating = False [cite: 61]
    
    ' Establish the global highlight color index for the application session.
    ' This acts as the palette choice that the .Replacement engine will look to.
    Options.DefaultHighlightColorIndex = wdBrightGreen
    
    ' Iterate sequentially through each keyword defined in the configuration array [cite: 9]
    For Each targetWord In wordList
        
        ' -----------------------------------------------------------------
        ' SCOPE ISOLATION: TARGETING THE MAIN TEXT CANVAS
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL STRATEGY: By pulling from ActiveDocument.Content instead of 
        ' iterating all StoryRanges, the macro locks its execution context to the main text story[cite: 8, 12].
        ' Under Word's object hierarchy, tables embedded in the document body are 
        ' structurally part of the main text story[cite: 10, 21]. This allows the Find engine to 
        ' penetate tables natively while keeping headers/footers entirely untouched[cite: 8, 12].
        ' Note: The range object collapses as matches are found; it MUST be completely 
        ' re-instantiated on every word pass to reset boundaries from page 1 to the end.
        Set docRange = ActiveDocument.Content [cite: 72]
        
        With docRange.Find
            ' Clear out any residual search, font, or replacement criteria hanging in memory 
            ' from previous manual operations or macro runs to avoid lookup conflicts.
            .ClearFormatting
            .Replacement.ClearFormatting
            
            ' Assign the current target lookup text string from the loop tracker
            .Text = targetWord
            
            ' CASE-INSENSITIVITY ENGINE RULES:
            ' Setting .MatchCase to False ensures Word catches variations like "fig", "Fig", 
            ' "FIG", or "FiG" with identical precision across the range[cite: 13, 22].
            .MatchCase = False [cite: 13, 22]
            
            ' FRACTIONAL STRING PARSING OVERRIDES:
            ' Setting .MatchWholeWord to False allows the token engine to sweep up sub-strings.
            ' For example, searching for "fig" will successfully capture "figure" or "figures".
            ' If you require exact, standalone matches only, toggle this parameter to True[cite: 40, 42].
            .MatchWholeWord = False
            
            ' BOUNDARY DEFENSE:
            ' Setting .Wrap to wdFindStop instructs the engine to process the specific range block 
            ' from top to bottom exactly once. This eliminates the risk of Word hitting the end of 
            ' the document and wrapping back around into a continuous execution loop[cite: 405, 408].
            .Wrap = wdFindStop [cite: 408]
            
            ' Format instruction telling the backend layout processor that any string sequence 
            ' intercepted by the lookup pattern must have an active highlight attribute stamped over it.
            .Replacement.Highlight = True
            
            ' Fire the native Find engine to execute a global bulk replacement pass across the range[cite: 78].
            .Execute Replace:=wdReplaceAll [cite: 78]
        End With
    Next targetWord
    
    ' Re-enable screen rendering to display the finalized layout updates to the user [cite: 244]
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Target words successfully highlighted within the main document and tables.", _
           vbInformation, "Highlight Processing Complete"
End Sub

Sub Misc_3_Fix_Common_Misspellings()
    ' =========================================================================
    ' MODULE NAME:  Misc_3_Fix_Common_Misspellings
    ' PURPOSE:      Automatically identifies and replaces a pre-configured dictionary
    '               of common misspellings across every single layer of the document.
    ' SCOPE:        Main body text, tables, headers, footers, textboxes, and footnotes.
    ' RULES:        Case-insensitive matching, but strictly enforces whole-word checks 
    '               to avoid accidentally corrupting longer, correctly spelled words.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Utilizes an in-memory Scripting.Dictionary mapped against a
    '               two-tier background StoryRanges text traversal loop. This avoids 
    '               physical cursor selection, eliminating screen flicker and macro lag.
    ' =========================================================================
    
    Dim doc As Document
    Dim story As Range
    Dim errorMap As Object
    Dim incorrectWord As Variant
    Dim correctWord As String
    
    Set doc = ActiveDocument
    
    ' Freeze visual application repagination to completely eliminate macro lag.
    ' This suppresses display stuttering and prevents the layout processor from 
    ' attempting to constantly redraw the workspace line-by-line during the sweeps.
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' HIGH-SPEED MEMORY ALLOCATION
    ' -----------------------------------------------------------------
    ' Instantiate the dictionary object via Late Binding (ActiveX CreateObject).
    ' This allocates a hash-mapped data framework in your system memory, allowing the 
    ' macro to scale dynamically without any reference library path dependencies.
    Set errorMap = CreateObject("Scripting.Dictionary")
    
    ' =========================================================================
    ' CONFIGURATION DICTIONARY: Add your custom spelling mappings here
    ' Syntax: errorMap.Add "WRONG_WORD", "CORRECT_WORD"
    ' Feel free to stack hundreds of custom entries below as your needs scale.
    ' =========================================================================
    errorMap.Add "wereda", "woreda"
    errorMap.Add "weredas", "woredas"
    errorMap.Add "tabel", "table"
    errorMap.Add "programme", "program"
    errorMap.Add "labour", "labor"
    ' =========================================================================
    
    ' Loop sequentially through every key (incorrect word token) registered in the dictionary
    For Each incorrectWord In errorMap.Keys
        ' Extract the corresponding clean string value from the current key pointer
        correctWord = errorMap(incorrectWord)
        
        ' -----------------------------------------------------------------
        ' THE LAYER INTERCEPTION SWEEP (StoryRanges Traversal)
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL STRATEGY: Word isolates text blocks into independent structural
        ' canvases called Story Ranges. This outer loop targets those ranges directly.
        ' This chose guarantees that misspellings are wiped out in a single pass 
        ' not just from main body text, but also from inside tables, headers, footers, 
        ' textboxes, and footnotes.
        For Each story In doc.StoryRanges
            Do
                With story.Find
                    ' Clear out any residual search, font, or replacement criteria hanging 
                    ' in memory from previous manual operations or macro runs to avoid lookup conflicts.
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    
                    ' Configure replacement criteria parameters
                    .Text = CStr(incorrectWord)
                    .Replacement.Text = correctWord

Sub Misc_4_Trim_Headings()
    ' =========================================================================
    ' MODULE NAME:  Misc_4_Trim_Headings
    ' PURPOSE:      Sweeps the document to clean up margins around heading structures.
    '               Strips out leading/trailing spaces, rogue tab characters, and
    '               unwanted trailing periods, colons, or hyphens.
    ' SCOPE:        Document body paragraphs matching Heading styles (Levels 1 to 9).
    ' SAFETY:       Automatically protects tables and ignores generic body text layers.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Iterates paragraph-by-paragraph via background Range processing,
    '               bypassing the cursor Selection to eliminate screen stuttering[cite: 1204].
    ' =========================================================================
    
    Dim doc As Document
    Dim para As Paragraph
    Dim txtRange As Range
    Dim paraText As String
    Dim cleanText As String
    Dim originalText As String
    
    Set doc = ActiveDocument
    
    ' Disable application window rendering to completely freeze visual repagination.
    ' This prevents massive background layout recalculation lag on large documents[cite: 129].
    Application.ScreenUpdating = False
    
    ' Instantiate global runtime error trapping to protect the active workspace environment
    On Error GoTo ErrorHandler

    ' Iterate sequentially through every paragraph entry in the core text story [cite: 1205]
    For Each para In doc.Paragraphs
        
        ' -----------------------------------------------------------------
        ' GUARDRAIL PHASE: GRID ARCHITECTURE INSULATION
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL STRATEGY: Modifying text strings that reside inside data tables 
        ' can corrupt alignment matrices or cause unexpected cell overflow anomalies[cite: 99].
        ' Checking .Information(wdWithInTable) isolates and bypasses tabular grids[cite: 98].
        If Not para.Range.Information(wdWithInTable) Then
            
            ' FILTER MECHANISM: Evaluate the structural metadata level of the paragraph[cite: 145].
            ' This isolates built-in Headings (Levels 1 to 9) while ignoring standard 
            ' unnumbered body text ranges (wdOutlineLevelBodyText)[cite: 146, 505].
            If para.OutlineLevel >= 1 And para.OutlineLevel <= 9 Then
                
                ' Bind a pointer to the individual paragraph text span range [cite: 1203]
                Set txtRange = para.Range
                paraText = txtRange.text
                originalText = paraText
                
                ' STRING SANITIZATION NOTE: Word appends an internal carriage return character 
                ' (vbCr / ¶) to mark the end of every structural paragraph[cite: 1205, 1243]. 
                ' This must be peeled back temporarily to prevent string manipulation failures.
                If Right(paraText, 1) = vbCr Then paraText = Left(paraText, Len(paraText) - 1)
                
                ' Initialize our workspace text string variable
                cleanText = paraText
                
                ' -----------------------------------------------------------------
                ' THE LOOP ENGINE: CONTINUOUS MARGIN STRIPPING
                ' -----------------------------------------------------------------
                ' RATIONALE: Simple single-pass string operations fail if complex typos 
                ' exist (e.g., a heading starting with multiple spaces, followed by a dot).
                ' Wrapping character sweeps inside a Do...Loop instructs the engine to trim
                ' from the outside in until no targeted characters remain[cite: 100, 101].
                
                ' 1. Leading Boundary Character Sweeper
                Do
                    Dim initialLenAsLeading As Long
                    initialLenAsLeading = Len(cleanText)
                    
                    ' Clear default leading whitespace blocks
                    cleanText = Trim(cleanText)
                    
                    ' Squeeze out legacy horizontal tab formatting markers (vbTab) [cite: 102]
                    If Left(cleanText, 1) = vbTab Then cleanText = Mid(cleanText, 2)
                    
                    ' Erase unmanaged leading punctuation sequences [cite: 102]
                    If Left(cleanText, 1) = "." Or Left(cleanText, 1) = "-" Then cleanText = Mid(cleanText, 2)
                    
                Loop Until Len(cleanText) = initialLenAsLeading Or Len(cleanText) = 0
                
                ' 2. Trailing Boundary Character Sweeper
                Do
                    Dim initialLenAsTrailing As Long
                    initialLenAsTrailing = Len(cleanText)
                    
                    ' Clear default trailing whitespace blocks
                    cleanText = Trim(cleanText)
                    
                    ' Squeeze out legacy trailing horizontal tab markers (vbTab) [cite: 103]
                    If Right(cleanText, 1) = vbTab Then cleanText = Left(cleanText, Len(cleanText) - 1)
                    
                    ' Clear trailing periods, hyphens, and colons.
                    ' Converts "10.5. Title." to "10.5. Title" to comply with publishing guidelines[cite: 103].
                    If Right(cleanText, 1) = "." Or Right(cleanText, 1) = "-" Or Right(cleanText, 1) = ":" Then
                        cleanText = Left(cleanText, Len(cleanText) - 1)
                    End If
                    
                Loop Until Len(cleanText) = initialLenAsTrailing Or Len(cleanText) = 0
                
                ' -----------------------------------------------------------------
                ' RE-STAMPING ENGINE (IF EDITS OCCURRED)
                ' -----------------------------------------------------------------
                ' Performance Optimization: Only commit a write operation to the canvas 
                ' if the trimmed string differs from the original text block.
                If cleanText <> paraText Then
                    
                    ' DIRECT FORMATTING OVERRIDE TRAP: Replacing a Range's text via VBA 
                    ' (`txtRange.text = ...`) causes Word to drop structural style metadata 
                    ' or inherit random direct character formatting layer remnants[cite: 86, 104].
                    ' We cache the active Style pointer object before performing the mutation[cite: 1248].
                    Dim currentStyle As Variant
                    Set currentStyle = txtRange.Style
                    
                    ' Write the perfectly trimmed text string back to the canvas, re-attaching the carriage return
                    txtRange.text = cleanText & vbCr
                    
                    ' RE-STAMP & HARD LAYOUT RESET: Re-assigning the original Style rule combined 
                    ' with `.Font.Reset` forcefully strips away any residual manual formatting 
                    ' layer overrides and forces characters to instantly conform to your style sheet[cite: 87, 88].
                    txtRange.Style = currentStyle
                    txtRange.Font.Reset [cite: 1214]
                    
                End If
                
            End If
        End If
    Next para

CleanUp:
    ' Re-enable application layout rendering to display all finalized text improvements
    Application.ScreenUpdating = True
    MsgBox "Heading margins successfully trimmed and cleaned!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Structural Fallback: Ensure the display engine is safely unfrozen if a fatal error occurs
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Execution Fault"
End Sub

Sub Misc_5_Trim_Multiple_Spaces_In_Selection()
    ' =========================================================================
    ' MODULE NAME:  Misc_5_Trim_Multiple_Spaces_In_Selection
    ' PURPOSE:      Finds and replaces all double (and more than double) spaces within
    '               the user's highlighted selection (paragraphs and tables) and
    '               collapses them down into a single standard space.
    ' SCOPE:        Active user selection ONLY. Leaves unhighlighted text completely intact.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Utilizes Word's native, high-speed Find and Replace engine mapped
    '               against an in-memory Selection.Range[cite: 107]. Bypasses cell-by-cell loops 
    '               for rapid execution within tables[cite: 116].
    ' =========================================================================
    
    Dim selectRange As Range
    
    ' -----------------------------------------------------------------
    ' SELECTION GUARDRAIL PHASE
    ' -----------------------------------------------------------------
    ' In the Word object model, `wdSelectionIP` represents an Insertion Point 
    ' (a blinking cursor with zero highlighted text characters). 
    ' If the macro runs without a text block selected, executing a Find pass 
    ' can fail or exhibit undefined scope mutations. This conditional blocks execution.
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the paragraph(s) or table area you want to clean first.", _
               vbExclamation, "No Selection Detected"
        Exit Sub
    End If
    
    ' Assign the precise boundary limits of your current visual selection to a 
    ' background Range pointer[cite: 108]. This insulates the operational range 
    ' from changes if the user accidentally clicks on the screen while the macro runs.
    Set selectRange = Selection.Range
    
    ' Freeze visual application window rendering to completely eliminate macro lag[cite: 782].
    ' This suppresses screen flickering and accelerates execution rates when sweeping heavy blocks[cite: 783].
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' HIGH-SPEED FIND & REPLACE PROCESSING
    ' -----------------------------------------------------------------
    With selectRange.Find
        ' Erase any residual search, font, or replacement criteria hanging in memory 
        ' from previous manual operations or macro runs to avoid criteria mismatches.
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' Configure targeted string tokens
        .Text = "  "                 ' Target exactly two sequential spaces [cite: 109]
        .Replacement.Text = " "      ' Replace with one single standard space [cite: 109]
        
        .Forward = True
        
        ' THE SELECTION DEFENSE BOUNDARY RULE:
        ' Changing this wrapper rule from default (wdFindContinue) to wdFindStop is the 
        ' most critical guardrail in the script[cite: 110]. It explicitly commands the layout 
        ' processor that the moment it hits the outer perimeter of your text selection, 
        ' it must freeze execution[cite: 111]. This guarantees it never bleeds into the rest of the document[cite: 112].
        .Wrap = wdFindStop
        
        ' Clear structural layout and matching variables to prioritize speed
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' -------------------------------------------------------------
        ' THE CONTINUOUS COLLAPSE DO-WHILE ENGINE
        ' -------------------------------------------------------------
        ' EXPLANATION: If a user has mashed the spacebar five times ("     "), 
        ' a single standalone ReplaceAll execution pass only collapses it down to 
        ' three spaces, and a second pass drops it to two[cite: 113].
        ' Placing the .Execute method directly inside a rolling Do While loop forces 
        ' Word to continuously sweep through the range until it declares with absolute 
        ' certainty that zero instances of double-spaces remain on the canvas[cite: 114].
        '
        ' SEAMLESS TABLE ARCHIECTURE SUPPORT: 
        ' Because Word natively treats highlighted data cells as a continuous string 
        ' fragment within the selection object model, this Find operation handles table text 
        ' identically to standard text lines[cite: 115]. It clears out space clutter inside 
        ' your rows instantly without requiring slow, cell-by-cell nested loop routines[cite: 116].
        Do While .Execute(Replace:=wdReplaceAll)
            ' Loop body intentionally left blank; execution evaluation handles the tracking.
        Loop
    End With
    
    ' Re-enable application layout rendering to instantly display all finalized text improvements
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Successfully collapsed all multiple spaces down to single spaces within your selection!", _
           vbInformation, "Process Complete"
End Sub

Sub Misc_6_Correct_Selected_Paragraph_Indents()
    ' ============================================================================
    ' MODULE NAME:  Misc_6_Correct_Selected_Paragraph_Indents
    ' PURPOSE:      Resets left, right, and first-line/hanging indents to 0
    '               for selected paragraphs, explicitly skipping active lists[cite: 1, 2].
    ' SCOPE:        Applies exclusively to highlighted text blocks. Safely bypasses 
    '               automated bulleted, numbered, or multi-tier outline lists[cite: 2].
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Restricts processing boundaries to the active Selection object,
    '               eliminating the lag of global content scans and preventing
    '               screen flickering during batch formatting tasks[cite: 3, 4].
    ' ============================================================================
    
    Dim para As Paragraph
    
    ' Establish a global error handler trap to gracefully catch locked or corrupt layout blocks
    On Error GoTo CleanUp
    
    ' Performance Optimization: Freeze visual layout rendering to stop Word from redrawing 
    ' page layouts line-by-line, accelerating execution speeds on heavy contracts[cite: 3, 12, 13].
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' 1. SCOPE ISOLATION VIA SELECTION.PARAGRAPHS
    ' -----------------------------------------------------------------
    ' ARCHITECTURAL STRATEGY: Instead of querying the global text layer (ActiveDocument.Content), 
    ' which requires expensive processing loops, the macro isolates its boundaries to the 
    ' Selection object, evaluating only what is actively highlighted on the screen[cite: 4, 5].
    For Each para In Selection.Paragraphs
        
        ' -----------------------------------------------------------------
        ' 2. THE LIST IMMUNITY FILTER (wdListNoNumbering)
        ' -----------------------------------------------------------------
        ' In Microsoft Word's backend engine, the ListType property categorizes the underlying 
        ' layout schema of text elements[cite: 6]. By validating that this property matches 
        ' wdListNoNumbering, the macro forces a strict match for plain body text[cite: 7].
        ' This selectively shields automated bullet points (wdListBullet), basic numeric sequences 
        ' (wdListSimpleNumbering), and multi-tier legal hierarchies from modification[cite: 7].
        If para.Range.ListFormat.ListType = wdListNoNumbering Then
            
            ' -----------------------------------------------------------------
            ' 3. DIRECT PARAMETER RESETTING
            ' -----------------------------------------------------------------
            ' Following advanced Word VBA practices, layout operations are applied directly to 
            ' the Paragraph variable interface rather than routing through sub-objects[cite: 8].
            ' Setting these three parameters comprehensively cleans the margins:
            
            ' Aligns the baseline left-hand margin boundary perfectly against the page margins[cite: 9].
            para.LeftIndent = 0
            
            ' Normalizes the right-hand text span boundary layout[cite: 10].
            para.RightIndent = 0
            
            ' Clears both positive first-line drops and negative custom hanging indents[cite: 11],
            ' ensuring the paragraph text reads completely straight like a razor edge.
            para.FirstLineIndent = 0
            
        End If
        
    Next para

CleanUp:
    ' Essential Safety Pass: Forcefully restore visual application screen rendering[cite: 12],
    ' ensuring the user workspace updates cleanly even if a fatal error occurs.
    Application.ScreenUpdating = True
    
    ' Final Execution Report Badging
    If Err.Number = 0 Then
        MsgBox "Indentation successfully cleared for all selected non-list paragraphs!", _
               vbInformation, "Format Clean Complete"
    Else
        MsgBox "The layout engine encountered an error: " & Err.Description, _
               vbCritical, "Execution Failure"
    End If
End Sub

Sub Misc_7_Fix_Table_Row_Cell_Padding()
    ' =========================================================================
    ' MODULE NAME:  Misc_7_Fix_Table_Row_Cell_Padding
    ' PURPOSE:      Standardizes table formatting globally across the document.
    '               Strips inherited paragraph spacing, clears hidden cell padding,
    '               and sets rows to auto-fit text heights tightly without layout
    '               clipping or text lines wrapping unexpectedly.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Loops directly through the Document Collection, bypassing the 
    '               Selection object to prevent screen flicker and macro lag.
    ' =========================================================================
    
    Dim tbl As Table
    
    ' Disable screen updates to maximize background processing speed
    Application.ScreenUpdating = False
    
    ' Loop through every table asset globally in the active document range
    For Each tbl In ActiveDocument.Tables
        With tbl
            ' -----------------------------------------------------------------
            ' 1. LOCALIZED TEXT SPACING RESET
            ' -----------------------------------------------------------------
            ' Force paragraph spacing inside the grid to stay completely flush.
            ' This strips out inherited document-wide body text suffixes (e.g., 6pt after)
            ' that cause data cells to expand unevenly.
            .Range.ParagraphFormat.SpaceBefore = 0
            .Range.ParagraphFormat.SpaceAfter = 0
            
            ' Establish dynamic line spacing using a multiple multiplier.
            ' Setting this via LinesToPoints ensures that if font sizes are changed
            ' later, the line tracking scales dynamically behind the scenes.
            .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            .Range.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
            
            ' -----------------------------------------------------------------
            ' 2. CELL PADDING (MARGIN) STRIPPING
            ' -----------------------------------------------------------------
            ' Clear internal cell margins which trap white space inside a row.
            ' Top and Bottom are zeroed out to smash invisible layout walls.
            ' Left and Right keep a 0.05" micro-buffer so text does not collide
            ' directly into vertical table grid lines, maintaining readability.
            .TopPadding = InchesToPoints(0)
            .BottomPadding = InchesToPoints(0)
            .LeftPadding = InchesToPoints(0.05)
            .RightPadding = InchesToPoints(0.05)
            
            ' -----------------------------------------------------------------
            ' 3. DYNAMIC ROW HEIGHT CLAMPING
            ' -----------------------------------------------------------------
            ' Set the row sizing behavior to Auto and collapse explicit height limits.
            ' This hands layout rendering back to Word's engine, letting rows naturally 
            ' contract to match the exact baseline height of the characters inside them.
            .Rows.HeightRule = wdRowHeightAuto
            .Rows.Height = 0
        End With
    Next tbl
    
    ' Re-enable screen rendering to display the finalized layout updates
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Table cell padding cleared successfully!", vbInformation, "Layout Fixed"
End Sub