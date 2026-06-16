Sub Misc_1_Clear_All_Highlighting_Globally()
    ' =========================================================================
    ' MODULE NAME:  Misc_1_Clear_All_Highlighting_Globally
    ' PURPOSE:      Completely clears all background text highlighting across every
    '               structural layer of the active document, including stubborn 
    '               unlinked headers/footers.
    ' SCOPE:        Main document text, tables, headers, footers, text boxes, footnotes.
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
    '               and applies a standardized green highlight overlay to them.
    ' SCOPE:        Main text body and embedded tables ONLY.
    ' SETTINGS:     Case-insensitive match. Evaluates word fragments (e.g. finds "fig" inside "figure").
    ' =========================================================================
    
    Dim docRange As Range
    Dim wordList As Variant
    Dim targetWord As Variant
    
    ' EXPLICIT CONFIGURATION: Define your target word list here
    ' Expand this array as needed to scale the list dynamically over time.
    wordList = Array("tab", "fig", "annex", "plate") 
    
    ' Freeze visual application repagination to eliminate macro lag.
    ' This suppresses display stuttering and drastically cuts processing overhead.
    Application.ScreenUpdating = False 
    
    ' Establish the global highlight color index for the application session.
    ' This acts as the palette choice that the .Replacement engine will look to.
    Options.DefaultHighlightColorIndex = wdBrightGreen
    
    ' Iterate sequentially through each keyword defined in the configuration array 
    For Each targetWord In wordList
        
        ' -----------------------------------------------------------------
        ' SCOPE ISOLATION: TARGETING THE MAIN TEXT CANVAS
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL STRATEGY: By pulling from ActiveDocument.Content instead of 
        ' iterating all StoryRanges, the macro locks its execution context to the main text story.
        ' Under Word's object hierarchy, tables embedded in the document body are 
        ' structurally part of the main text story. This allows the Find engine to 
        ' penetate tables natively while keeping headers/footers entirely untouched.
        ' Note: The range object collapses as matches are found; it MUST be completely 
        ' re-instantiated on every word pass to reset boundaries from page 1 to the end.
        Set docRange = ActiveDocument.Content 
        
        With docRange.Find
            ' Clear out any residual search, font, or replacement criteria hanging in memory 
            ' from previous manual operations or macro runs to avoid lookup conflicts.
            .ClearFormatting
            .Replacement.ClearFormatting
            
            ' Assign the current target lookup text string from the loop tracker
            .Text = targetWord
            
            ' CASE-INSENSITIVITY ENGINE RULES:
            ' Setting .MatchCase to False ensures Word catches variations like "fig", "Fig", 
            ' "FIG", or "FiG" with identical precision across the range.
            .MatchCase = False 
            
            ' FRACTIONAL STRING PARSING OVERRIDES:
            ' Setting .MatchWholeWord to False allows the token engine to sweep up sub-strings.
            ' For example, searching for "fig" will successfully capture "figure" or "figures".
            ' If you require exact, standalone matches only, toggle this parameter to True.
            .MatchWholeWord = False
            
            ' BOUNDARY DEFENSE:
            ' Setting .Wrap to wdFindStop instructs the engine to process the specific range block 
            ' from top to bottom exactly once. This eliminates the risk of Word hitting the end of 
            ' the document and wrapping back around into a continuous execution loop.
            .Wrap = wdFindStop 
            
            ' Format instruction telling the backend layout processor that any string sequence 
            ' intercepted by the lookup pattern must have an active highlight attribute stamped over it.
            .Replacement.Highlight = True
            
            ' Fire the native Find engine to execute a global bulk replacement pass across the range.
            .Execute Replace:=wdReplaceAll 
        End With
    Next targetWord
    
    ' Re-enable screen rendering to display the finalized layout updates to the user 
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
    ' =========================================================================
    
    Dim doc As Document
    Dim para As Paragraph
    Dim txtRange As Range
    Dim paraText As String
    Dim cleanText As String
    Dim originalText As String
    
    Set doc = ActiveDocument
    
    ' Disable application window rendering to completely freeze visual repagination.
    ' This prevents massive background layout recalculation lag on large documents.
    Application.ScreenUpdating = False
    
    ' Instantiate global runtime error trapping to protect the active workspace environment
    On Error GoTo ErrorHandler

    ' Iterate sequentially through every paragraph entry in the core text story 
    For Each para In doc.Paragraphs
        
        ' -----------------------------------------------------------------
        ' GUARDRAIL PHASE: GRID ARCHITECTURE INSULATION
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL STRATEGY: Modifying text strings that reside inside data tables 
        ' can corrupt alignment matrices or cause unexpected cell overflow anomalies.
        ' Checking .Information(wdWithInTable) isolates and bypasses tabular grids.
        If Not para.Range.Information(wdWithInTable) Then
            
            ' FILTER MECHANISM: Evaluate the structural metadata level of the paragraph.
            ' This isolates built-in Headings (Levels 1 to 9) while ignoring standard 
            ' unnumbered body text ranges (wdOutlineLevelBodyText).
            If para.OutlineLevel >= 1 And para.OutlineLevel <= 9 Then
                
                ' Bind a pointer to the individual paragraph text span range 
                Set txtRange = para.Range
                paraText = txtRange.text
                originalText = paraText
                
                ' STRING SANITIZATION NOTE: Word appends an internal carriage return character 
                ' (vbCr / ¶) to mark the end of every structural paragraph. 
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
                ' from the outside in until no targeted characters remain.
                
                ' 1. Leading Boundary Character Sweeper
                Do
                    Dim initialLenAsLeading As Long
                    initialLenAsLeading = Len(cleanText)
                    
                    ' Clear default leading whitespace blocks
                    cleanText = Trim(cleanText)
                    
                    ' Squeeze out legacy horizontal tab formatting markers (vbTab) 
                    If Left(cleanText, 1) = vbTab Then cleanText = Mid(cleanText, 2)
                    
                    ' Erase unmanaged leading punctuation sequences 
                    If Left(cleanText, 1) = "." Or Left(cleanText, 1) = "-" Then cleanText = Mid(cleanText, 2)
                    
                Loop Until Len(cleanText) = initialLenAsLeading Or Len(cleanText) = 0
                
                ' 2. Trailing Boundary Character Sweeper
                Do
                    Dim initialLenAsTrailing As Long
                    initialLenAsTrailing = Len(cleanText)
                    
                    ' Clear default trailing whitespace blocks
                    cleanText = Trim(cleanText)
                    
                    ' Squeeze out legacy trailing horizontal tab markers (vbTab) 
                    If Right(cleanText, 1) = vbTab Then cleanText = Left(cleanText, Len(cleanText) - 1)
                    
                    ' Clear trailing periods, hyphens, and colons.
                    ' Converts "10.5. Title." to "10.5. Title" to comply with publishing guidelines.
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
                    ' or inherit random direct character formatting layer remnants.
                    ' We cache the active Style pointer object before performing the mutation.
                    Dim currentStyle As Variant
                    Set currentStyle = txtRange.Style
                    
                    ' Write the perfectly trimmed text string back to the canvas, re-attaching the carriage return
                    txtRange.text = cleanText & vbCr
                    
                    ' RE-STAMP & HARD LAYOUT RESET: Re-assigning the original Style rule combined 
                    ' with `.Font.Reset` forcefully strips away any residual manual formatting 
                    ' layer overrides and forces characters to instantly conform to your style sheet.
                    txtRange.Style = currentStyle
                    txtRange.Font.Reset 
                    
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
    ' background Range pointer. This insulates the operational range 
    ' from changes if the user accidentally clicks on the screen while the macro runs.
    Set selectRange = Selection.Range
    
    ' Freeze visual application window rendering to completely eliminate macro lag.
    ' This suppresses screen flickering and accelerates execution rates when sweeping heavy blocks.
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
        .Text = "  "                 ' Target exactly two sequential spaces 
        .Replacement.Text = " "      ' Replace with one single standard space 
        
        .Forward = True
        
        ' THE SELECTION DEFENSE BOUNDARY RULE:
        ' Changing this wrapper rule from default (wdFindContinue) to wdFindStop is the 
        ' most critical guardrail in the script. It explicitly commands the layout 
        ' processor that the moment it hits the outer perimeter of your text selection, 
        ' it must freeze execution.
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
        ' three spaces, and a second pass drops it to two.
        ' Placing the .Execute method directly inside a rolling Do While loop forces 
        ' Word to continuously sweep through the range until it declares with absolute 
        ' certainty that zero instances of double-spaces remain on the canvas.
        '
        ' SEAMLESS TABLE ARCHIECTURE SUPPORT: 
        ' Because Word natively treats highlighted data cells as a continuous string 
        ' fragment within the selection object model, this Find operation handles table text 
        ' identically to standard text lines. It clears out space clutter inside 
        ' your rows instantly without requiring slow, cell-by-cell nested loop routines.
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
    '               for selected paragraphs, explicitly skipping active lists.
    ' SCOPE:        Applies exclusively to highlighted text blocks. Safely bypasses 
    '               automated bulleted, numbered, or multi-tier outline lists.
    ' ============================================================================
    
    Dim para As Paragraph
    
    ' Establish a global error handler trap to gracefully catch locked or corrupt layout blocks
    On Error GoTo CleanUp
    
    ' Performance Optimization: Freeze visual layout rendering to stop Word from redrawing 
    ' page layouts line-by-line, accelerating execution speeds on heavy contracts.
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' 1. SCOPE ISOLATION VIA SELECTION.PARAGRAPHS
    ' -----------------------------------------------------------------
    ' ARCHITECTURAL STRATEGY: Instead of querying the global text layer (ActiveDocument.Content), 
    ' which requires expensive processing loops, the macro isolates its boundaries to the 
    ' Selection object, evaluating only what is actively highlighted on the screen.
    For Each para In Selection.Paragraphs
        
        ' -----------------------------------------------------------------
        ' 2. THE LIST IMMUNITY FILTER (wdListNoNumbering)
        ' -----------------------------------------------------------------
        ' In Microsoft Word's backend engine, the ListType property categorizes the underlying 
        ' layout schema of text elements. By validating that this property matches 
        ' wdListNoNumbering, the macro forces a strict match for plain body text.
        ' This selectively shields automated bullet points (wdListBullet), basic numeric sequences 
        ' (wdListSimpleNumbering), and multi-tier legal hierarchies from modification.
        If para.Range.ListFormat.ListType = wdListNoNumbering Then
            
            ' -----------------------------------------------------------------
            ' 3. DIRECT PARAMETER RESETTING
            ' -----------------------------------------------------------------
            ' Following advanced Word VBA practices, layout operations are applied directly to 
            ' the Paragraph variable interface rather than routing through sub-objects.
            ' Setting these three parameters comprehensively cleans the margins:
            
            ' Aligns the baseline left-hand margin boundary perfectly against the page margins.
            para.LeftIndent = 0
            
            ' Normalizes the right-hand text span boundary layout.
            para.RightIndent = 0
            
            ' Clears both positive first-line drops and negative custom hanging indents,
            ' ensuring the paragraph text reads completely straight like a razor edge.
            para.FirstLineIndent = 0
            
        End If
        
    Next para

CleanUp:
    ' Essential Safety Pass: Forcefully restore visual application screen rendering,
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

Sub Misc_8_Trim_Global_Table_Paragraph_Marks()
    ' =========================================================================
    ' MODULE NAME:  Misc_8_Trim_Global_Table_Cell_Paragraph_Marks
    ' PURPOSE:      Scans every table globally across the active document to identify
    '               and forcefully strip out manually inserted empty paragraph breaks
    '               (vbCr / ¶) hanging at the absolute top and bottom of table cells.
    ' SCOPE:        All document layers containing structured data grid tables.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Targets the ActiveDocument.Tables collection directly to bypass
    '               the cursor Selection object, optimizing background rendering speed.
    ' =========================================================================
    
    Dim tbl As Table
    Dim rowIdx As Row
    Dim cel As Cell
    Dim pga As Paragraph
    Dim pgaCount As Long
    Dim i As Long
    
    ' Disable screen re-pagination routines to maximize background layout processing speed
    Application.ScreenUpdating = False
    
    ' Globally trap unexpected errors (like split/merged cell index range errors)
    On Error GoTo ErrorHandler
    
    ' Iterate sequentially through every table asset residing in the document
    For Each tbl In ActiveDocument.Tables
        
        ' Safely traverse row-by-row through the target table grid layout
        For Each rowIdx In tbl.Rows
            
            ' Traverse cell-by-cell across the active row matrix
            For Each cel In rowIdx.Cells
                
                ' -------------------------------------------------------------
                ' PHASE 1: CLEANING LEADING EMPTY PARAGRAPHS (TOP OF CELL)
                ' -------------------------------------------------------------
                ' Continue evaluating the absolute first paragraph of the cell as long
                ' as there is more than 1 paragraph total in the cell.
                Do While cel.Range.Paragraphs.Count > 1
                    Set pga = cel.Range.Paragraphs(1)
                    
                    ' If the length of the first paragraph text is exactly 1,
                    ' it means the paragraph contains nothing but the invisible carriage return (¶).
                    If Len(pga.Range.Text) = 1 Then
                        pga.Range.Delete
                    Else
                        ' Exit the Do-Loop immediately the moment a text string is encountered
                        Exit Do
                    End If
                Loop
                
                ' -------------------------------------------------------------
                ' PHASE 2: CLEANING TRAILING EMPTY PARAGRAPHS (BOTTOM OF CELL)
                ' -------------------------------------------------------------
                ' Recalculate total paragraphs left inside this specific cell framework
                pgaCount = cel.Range.Paragraphs.Count
                
                ' If the cell only has 1 line left, skip to protect the baseline grid marker
                If pgaCount > 1 Then
                    
                    ' ARCHITECTURAL STRATEGY: Loop BACKWARDS from the second-to-last paragraph 
                    ' down to the first text line. We completely ignore the absolute final paragraph 
                    ' item index (pgaCount) because Word hitches its vital, un-deletable cell-end marker 
                    ' token directly to that slot. Deleting text around it from the bottom up avoids crashes.
                    For i = (pgaCount - 1) To 1 Step -1
                        Set pga = cel.Range.Paragraphs(i)
                        
                        ' Verify if this middle/trailing paragraph is a hollow line placeholder
                        If Len(pga.Range.Text) = 1 Then
                            pga.Range.Delete
                        Else
                            ' The moment valid textual content or data vectors are struck,
                            ' stop the backward sweep to prevent erasing interior content line breaks.
                            Exit For
                        End If
                    Next i
                    
                End If
                
            Next cel
        Next rowIdx
        
    Next tbl

CleanUp:
    ' Re-enable native application rendering to display the finalized tight layouts
    Application.ScreenUpdating = True
    MsgBox "All loose, empty paragraph breaks have been successfully trimmed from your tables!", _
           vbInformation, "Table Trim Complete"
    Exit Sub

ErrorHandler:
    ' Structural Fallback: Ensure the display engine is unfrozen if an un-deletable cell boundary breaks
    Application.ScreenUpdating = True
    MsgBox "An unexpected layout error occurred during cell trimming: " & Err.Description, _
           vbCritical, "Execution Fault"
End Sub

Sub Misc_9_Trim_Selected_Table_Paragraph_Marks()
    ' =========================================================================
    ' MODULE NAME:  Misc_9_Trim_Selected_Table_Paragraph_Marks
    ' PURPOSE:      Scans the user's actively selected table to identify and
    '               forcefully strip out manually inserted empty paragraph breaks
    '               (vbCr / ¶) hanging at the absolute top and bottom of cells.
    ' SCOPE:        Active user-selected table ONLY. Leaves the rest of the 
    '               document and other unselected tables completely untouched.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Isolates execution to the Selection framework, maximizing 
    '               background processing speed and preventing screen flicker.
    ' =========================================================================
    
    Dim tbl As Table
    Dim rowIdx As Row
    Dim cel As Cell
    Dim pga As Paragraph
    Dim pgaCount As Long
    Dim i As Long
    
    ' -----------------------------------------------------------------
    ' GUARDRAIL PHASE: SELECTION VALIDATION
    ' -----------------------------------------------------------------
    ' Before attempting to pull table properties, we must verify if the cursor 
    ' is actually inside a table layout block. Checking wdWithInTable prevents 
    ' the macro from crashing if executed over regular body paragraphs.
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please click inside or select the specific table you want to clean first.", _
               vbExclamation, "No Table Selected"
        Exit Sub
    End If
    
    ' Bind our table variable pointer to the first table object inside the active selection range
    Set tbl = Selection.Tables(1)
    
    ' Disable screen updating to freeze visual repagination routines for raw execution speed
    Application.ScreenUpdating = False
    
    ' Globally trap unexpected errors (such as structural collisions in complex split/merged cells)
    On Error GoTo ErrorHandler
    
    ' -----------------------------------------------------------------
    ' DATA GRID PROCESSING ENGINE
    ' -----------------------------------------------------------------
    ' Safely traverse row-by-row through the user's targeted table layout matrix 
    For Each rowIdx In tbl.Rows
        
        ' Traverse cell-by-cell across the active row framework 
        For Each cel In rowIdx.Cells
            
            ' -------------------------------------------------------------
            ' PHASE 1: CLEANING LEADING EMPTY PARAGRAPHS (TOP OF CELL)
            ' -------------------------------------------------------------
            ' Continue evaluating the absolute first paragraph of the cell as long
            ' as there is more than 1 paragraph total inside the cell.
            Do While cel.Range.Paragraphs.Count > 1
                Set pga = cel.Range.Paragraphs(1)
                
                ' If the text string length is exactly 1, it holds nothing but an empty carriage return (¶)[cite: 1230].
                If Len(pga.Range.Text) = 1 Then
                    pga.Range.Delete
                Else
                    ' Exit the loop immediately the moment a valid textual character asset is struck [cite: 1240]
                    Exit Do
                End If
            Loop
            
            ' -------------------------------------------------------------
            ' PHASE 2: CLEANING TRAILING EMPTY PARAGRAPHS (BOTTOM OF CELL)
            ' -------------------------------------------------------------
            ' Recalculate total paragraphs left inside this specific cell framework
            pgaCount = cel.Range.Paragraphs.Count
            
            ' If the cell only has 1 line left, skip to protect the baseline grid cell marker
            If pgaCount > 1 Then
                
                ' ARCHITECTURAL STRATEGY: Loop BACKWARDS from the second-to-last paragraph[cite: 1238].
                ' We completely ignore the absolute final index slot because Word links its structural, 
                ' un-deletable cell-end marker token to that position[cite: 1233, 1234]. Deleting from the bottom 
                ' up prevents runtime layout corruption[cite: 1234, 1238].
                For i = (pgaCount - 1) To 1 Step -1
                    Set pga = cel.Range.Paragraphs(i)
                    
                    ' Verify if this trailing paragraph is a hollow whitespace line placeholder
                    If Len(pga.Range.Text) = 1 Then
                        pga.Range.Delete
                    Else
                        ' The moment legitimate content is encountered, freeze the reverse sweep[cite: 1240].
                        ' This preserves intentional formatting line breaks between interior text blocks[cite: 1239, 1240].
                        Exit For
                    End If
                Next i
                
            End If
            
        Next cel
    Next rowIdx

CleanUp:
    ' Re-enable visual rendering to instantly present the newly tightened layout boundaries
    Application.ScreenUpdating = True
    MsgBox "Loose paragraph marks successfully trimmed from the selected table!", _
           vbInformation, "Selection Trim Complete"
    Exit Sub

ErrorHandler:
    ' Structural Fallback: Safely unfreeze the user workspace if an operation breaks down
    Application.ScreenUpdating = True
    MsgBox "An unexpected layout error occurred during cell trimming: " & Err.Description, _
           vbCritical, "Execution Fault"
End Sub