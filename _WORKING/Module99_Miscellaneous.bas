Sub Misc_1_Clear_All_Highlighting_Globally()
    ' =========================================================================
    ' MODULE NAME:  Misc_1_Clear_All_Highlighting_Globally
    ' PURPOSE:      Completely clears all background text highlighting across every
    '               structural layer of the active document, including stubborn
    '               unlinked headers/footers.
    ' SCOPE:        Main document text, tables, headers, footers, text boxes, footnotes.
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
    ' background processing speed on long, multi-section documents[cite: 1701, 1702].
    Application.ScreenUpdating = False
    
    ' =========================================================================
    ' PASS 1: THE GLOBAL SWEEP (Catches Main Text, Tables, and Footnotes)
    ' =========================================================================
    ' This loop scans the background StoryRanges collection. It is the most
    ' efficient way to bypass cursor selection and hit major text blocks[cite: 1648].
    For Each story In doc.StoryRanges
        Do
            ' Set the highlight index directly to wdNoHighlight.
            ' This strips the background formatting layer without touching the text text.
            story.HighlightColorIndex = wdNoHighlight
            
            ' Word chunks text into sub-ranges (e.g., linked text boxes or split footnotes).
            ' NextStoryRange ensures the pointer evaluates downstream links in this story.
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next story ' <-- ARCHITECTURAL FIX: Replaced duplicate For Each statement to close loop legally
    
    ' =========================================================================
    ' PASS 2: DEEP SECTION PENETRATION (Forces dormant headers/footers awake)
    ' =========================================================================
    ' ARCHITECTURAL CULPRIT: Word does not automatically load every header and
    ' footer layer into memory unless they are actively visible or opened by the user[cite: 1657].
    ' If "Different First Page" or "Different Odd & Even Pages" layout flags are enabled,
    ' unlinked header/footer canvases remain completely dormant in the background cache[cite: 1658].
    ' Because Pass 1 loops right past dormant stories, Pass 2 explicitly declares a
    ' nested section loop to force the layout processor to evaluate all layout sub-layers[cite: 1659, 1660].
    For Each sec In doc.Sections
        
        ' -----------------------------------------------------------------
        ' Sub-Phase A: Section Headers
        ' -----------------------------------------------------------------
        ' Step through the section's Header collection (First Page, Even Pages, Primary).
        For Each compView In sec.Headers
            ' The .Exists safety gate prevents Word from throwing a runtime error
            ' if the sub-layer is structurally unassigned or inactive[cite: 1665].
            If compView.Exists Then
                ' Route through the sub-layer's .Range to expose its underlying text canvas[cite: 1709].
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
    
    ' Re-enable screen rendering to display the finalized layout updates
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
    wordList = Array("tab", "fig", "annex", "plate", "photo")
    
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
            .replacement.ClearFormatting
            
            ' Assign the current target lookup text string from the loop tracker
            .text = targetWord
            
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
            .replacement.Highlight = True
            
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
    errorMap.Add "Gelaso", "Geleaso"
    errorMap.Add "Aba'ala", "Abala"
    errorMap.Add "Amahara", "Amhara"
    errorMap.Add "Awra", "Awura"
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
                    .replacement.ClearFormatting
                    
                    ' Configure replacement criteria parameters
                    .text = CStr(incorrectWord)
                    .replacement.text = correctWord
                    
                    ' -------------------------------------------------------------
                    ' CRITICAL SAFETY ENGINE REGISTRATION RULES
                    ' -------------------------------------------------------------
                    ' CASE FLEXIBILITY FILTER: Setting MatchCase to False allows Word
                    ' to safely auto-detect capitalization styles. If it encounters
                    ' "Dpcumentation" at the start of a sentence, it naturally replaces
                    ' it with "Documentation" while keeping the capitalized formatting intact.
                    .MatchCase = False
                    
                    ' ANTI-FRAGMENT COLLISION PROTECTION: This is the most critical safeguard.
                    ' Without .MatchWholeWord = True set, a command to fix a typo like "teh"
                    ' into "the" would accidentally warp a perfectly spelled word like
                    ' "technique" into "thechnique". Forcing whole-word matching eliminates this risk.
                    .MatchWholeWord = True
                    
                    ' BOUNDARY TRACKING DEFENSER: Setting .Wrap to wdFindStop instructs the engine
                    ' to process the isolated text story from top to bottom exactly once.
                    ' This prevents Word from getting trapped in an endless loop at the range margin.
                    .Wrap = wdFindStop
                    
                    ' Fire the native Find engine to execute a global bulk replacement pass across the story
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' LINKED RANGE ASSIGNMENT: Word frequently splits sub-layers (such as linked text
                ' boxes or detached footnotes) into sequential sub-story pointer sequences.
                ' NextStoryRange ensures the tracking pointer steps forward through downstream links.
                Set story = story.NextStoryRange
            Loop Until story Is Nothing
        Next story
    Next incorrectWord
    
    ' Re-enable screen rendering to display the finalized structural updates to the user
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Spelling correction sweep complete across all document layers!", _
           vbInformation, "Auto-Correction Successful"
End Sub


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
                ' (vbCr / �) to mark the end of every structural paragraph.
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
        .replacement.ClearFormatting
        
        ' Configure targeted string tokens
        .text = "  "                 ' Target exactly two sequential spaces
        .replacement.text = " "      ' Replace with one single standard space
        
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
            ' ---------------
            ' .Range.ParagraphFormat.SpaceBefore = 0
            ' .Range.ParagraphFormat.SpaceAfter = 0
            
            ' Establish dynamic line spacing using a multiple multiplier.
            ' Setting this via LinesToPoints ensures that if font sizes are changed
            ' later, the line tracking scales dynamically behind the scenes.
            ' ---------------
            ' .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            ' .Range.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
            
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
            
            ' Invoke Word's built-in Table Options dialog to ensure that the internal
            ' cell margin settings are fully flushed and applied to the layout engine.
            With Dialogs(wdDialogTableTableOptions)
                .allowspacing = False
                .Execute
            End With
            
            '.Spacing = 0
            
            ' -----------------------------------------------------------------
            ' 3. DYNAMIC ROW HEIGHT CLAMPING
            ' -----------------------------------------------------------------
            ' Set the row sizing behavior to Auto and collapse explicit height limits.
            ' This hands layout rendering back to Word's engine, letting rows naturally
            ' contract to match the exact baseline height of the characters inside them.
            ' ---------------
            ' .Rows.HeightRule = wdRowHeightAuto
            ' .Rows.Height = 0
        End With
    Next tbl
    
    ' Re-enable screen rendering to display the finalized layout updates
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Table cell padding cleared successfully!", vbInformation, "Layout Fixed"
End Sub

Sub Misc_8_Trim_Global_Table_Paragraph_Marks()
    ' =========================================================================
    ' MODULE NAME:  Misc_8_Trim_Global_Table_Paragraph_Marks
    ' PURPOSE:      Scans every table globally across the active document to identify
    '               and forcefully strip out manually inserted empty paragraph breaks
    '               (vbCr / �) hanging at the absolute top and bottom of table cells.
    ' SCOPE:        All document layers containing structured data grid tables.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Targets the ActiveDocument.Tables collection directly and utilizes
    '               a flat .Range.Cells loop pass to guarantee absolute immunity against
    '               vertically or horizontally merged layout structures.
    ' =========================================================================
    
    Dim tbl As Table
    Dim cel As Cell
    Dim pga As Paragraph
    Dim pgaCount As Long
    Dim i As Long
    
    ' Disable screen re-pagination routines to maximize background layout processing speed
    Application.ScreenUpdating = False
    
    ' Globally trap unexpected errors (safeguards system environment state)
    On Error GoTo ErrorHandler
    
    ' Iterate sequentially through every table asset residing in the document
    For Each tbl In ActiveDocument.Tables
        
        ' -----------------------------------------------------------------
        ' DATA GRID PROCESSING ENGINE (COMPILE-SAFE & MERGED-CELL SAFE)
        ' -----------------------------------------------------------------
        ' ARCHITECTURAL FIX: Calling tbl.Cells throws a compile error because the Table
        ' object lacks a direct cells member. To target the flat, linear array of cells,
        ' we must explicitly route through tbl.Range.Cells. This completely bypasses
        ' row-by-row coordinate grids, ensuring full stability even if your document
        ' contains complex split or vertically merged table elements.
        For Each cel In tbl.Range.Cells
            
            ' -------------------------------------------------------------
            ' PHASE 1: CLEANING LEADING EMPTY PARAGRAPHS (TOP OF CELL)
            ' -------------------------------------------------------------
            ' Continue evaluating the absolute first paragraph of the cell as long
            ' as there is more than 1 paragraph total inside the cell.
            Do While cel.Range.Paragraphs.Count > 1
                Set pga = cel.Range.Paragraphs(1)
                
                ' If the text string length is exactly 1, it holds nothing but an empty carriage return (�).
                If Len(pga.Range.text) = 1 Then
                    pga.Range.Delete
                Else
                    ' Exit the loop immediately the moment a valid textual character asset is struck
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
                
                ' ARCHITECTURAL STRATEGY: Loop BACKWARDS from the second-to-last paragraph.
                ' We completely ignore the absolute final index slot because Word links its structural,
                ' un-deletable cell-end marker token to that position. Deleting from the bottom
                ' up prevents runtime layout corruption.
                For i = (pgaCount - 1) To 1 Step -1
                    Set pga = cel.Range.Paragraphs(i)
                    
                    ' Verify if this trailing paragraph is a hollow whitespace line placeholder
                    If Len(pga.Range.text) = 1 Then
                        pga.Range.Delete
                    Else
                        ' The moment legitimate content is encountered, freeze the reverse sweep.
                        ' This preserves intentional formatting line breaks between interior text blocks.
                        Exit For
                    End If
                Next i
                
            End If
            
        Next cel
        
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
    ' PURPOSE:      Scans the user's actively selected table cells to identify and
    '               forcefully strip out manually inserted empty paragraph breaks
    '               (vbCr / �) hanging at the absolute top and bottom of cells.
    ' SCOPE:        Active user-selected cells/table ONLY. Leaves the rest of the
    '               document and other unselected tables completely untouched.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' PERFORMANCE:  Isolates execution to Selection.Range.Cells to completely bypass
    '               VBE "Member not found" limits while remaining 100% immune
    '               to vertically or horizontally merged table layouts.
    ' =========================================================================
    
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
    
    ' Disable screen updating to freeze visual repagination routines for raw execution speed
    Application.ScreenUpdating = False
    
    ' Globally trap unexpected errors (safeguards system environment state)
    On Error GoTo ErrorHandler
    
    ' -----------------------------------------------------------------
    ' DATA GRID PROCESSING ENGINE (COMPILE-SAFE & MERGED-CELL SAFE)
    ' -----------------------------------------------------------------
    ' ARCHITECTURAL FIX: Selection.Cells throws a compile error. To target the flat,
    ' linear array of highlighted cells cleanly, we look through Selection.Range.Cells.
    ' This bypasses row-by-row coordinate grids completely, ensuring full stability
    ' even if your selected grid contains complex split or vertically merged items.
    For Each cel In Selection.Range.Cells
        
        ' -------------------------------------------------------------
        ' PHASE 1: CLEANING LEADING EMPTY PARAGRAPHS (TOP OF CELL)
        ' -------------------------------------------------------------
        ' Continue evaluating the absolute first paragraph of the cell as long
        ' as there is more than 1 paragraph total inside the cell.
        Do While cel.Range.Paragraphs.Count > 1
            Set pga = cel.Range.Paragraphs(1)
            
            ' If the text string length is exactly 1, it holds nothing but an empty carriage return (�).
            If Len(pga.Range.text) = 1 Then
                pga.Range.Delete
            Else
                ' Exit the loop immediately the moment a valid textual character asset is struck
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
            
            ' ARCHITECTURAL STRATEGY: Loop BACKWARDS from the second-to-last paragraph.
            ' We completely ignore the absolute final index slot because Word links its structural,
            ' un-deletable cell-end marker token to that position. Deleting from the bottom
            ' up prevents runtime layout corruption.
            For i = (pgaCount - 1) To 1 Step -1
                Set pga = cel.Range.Paragraphs(i)
                
                ' Verify if this trailing paragraph is a hollow whitespace line placeholder
                If Len(pga.Range.text) = 1 Then
                    pga.Range.Delete
                Else
                    ' The moment legitimate content is encountered, freeze the reverse sweep.
                    ' This preserves intentional formatting line breaks between interior text blocks.
                    Exit For
                End If
            Next i
            
        End If
        
    Next cel

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


Sub Misc_10_Purge_Ghost_TOC_Levels()
    ' =========================================================================
    ' MODULE NAME:  Misc_10_Purge_Ghost_TOC_Levels
    ' PURPOSE:      Identifies and fixes "Fake Headings" or Outline Level Corruption
    '               where a blank line or a standard paragraph (like Normal style)
    '               possesses an explicit outline level tracking value.
    '               This prevents phantom blank rows, rogue numbering strings, and
    '               empty dot leaders from penetrating your Table of Contents.
    ' SCOPE:        All standard document body paragraphs. Explicitly ignores tables.
    ' =========================================================================
    
    Dim para As Paragraph
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Freeze visual window rendering to stop Word from constantly attempting to
    ' re-paginate the page layout layout line-by-line during the paragraph scans.
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' 1. THE LAYER INTERCEPTION SWEEP (Paragraph Traversal)
    ' -----------------------------------------------------------------
    ' Loop paragraph-by-paragraph through the active text story to catch hidden
    ' layout overrides stamped directly onto paragraph properties.
    For Each para In doc.Paragraphs
        
        ' GRID ARCHITECTURE GUARDRAIL: Modifying text properties inside data tables
        ' can corrupt multi-tier alignments or break cell layout constraints.
        ' Checking .Information(wdWithInTable) completely insulates table content cells.
        If Not para.Range.Information(wdWithInTable) Then
            
            ' OUTLINE LEVEL OVERLAY TRAP: Word's TOC engine scrapes the document for
            ' anything flagged with structural levels 1, 2, or 3.
            ' This conditional locks onto paragraphs carrying those specific flags.
            If para.OutlineLevel >= wdOutlineLevel1 And para.OutlineLevel <= wdOutlineLevel3 Then
                
                ' THE STYLE ASSIGNMENT VERIFICATION FILTER:
                ' Real headings should be explicitly bound to Word's built-in Heading styles.
                ' We use Left(..., 7) to evaluate the first 7 characters of the localized style name.
                ' If it does NOT start with "Heading", it is a ghost element masquerading as a layout landmark.
                If Left(para.Style.NameLocal, 7) <> "Heading" Then
                    
                    ' THE HARD DEMOTION RESET: Forcefully strip the rogue structural level property,
                    ' pushing the paragraph back down to standard unranked text (wdOutlineLevelBodyText).
                    ' This cleanly removes it from the Navigation Pane and your index hierarchy.
                    para.OutlineLevel = wdOutlineLevelBodyText
                    
                End If
            End If
        End If
    Next para
    
    ' -----------------------------------------------------------------
    ' 2. PROGRAMMATIC FIELD CODE COMPILATION
    ' -----------------------------------------------------------------
    Dim toc As TableOfContents
    
    ' Iterate sequentially through every Table of Contents field block embedded in the file.
    ' Instead of forcing a slow manual document selection pass, looping through the collection
    ' targets the field engines directly behind the scenes.
    For Each toc In doc.TablesOfContents
        ' Force an immediate layout recalculation update on the active TOC.
        ' This pulls the newly cleaned paragraph matrices into your index, ensuring those
        ' phantom empty lines and broken dot leaders instantly vanish from your layout page.
        toc.Update
    Next toc
    
    ' Re-enable visual application updates to present the finalized clean layout metrics
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Ghost TOC levels removed and table updated successfully!", vbInformation, "Layout Cleaned"
End Sub

Sub Misc_11_Reset_All_List_Style_Links()
    ' =========================================================================
    ' MODULE NAME:  Reset_All_List_Style_Links
    ' PURPOSE:      Iterates through every list template initialized in the document's
    '               background cache and forcefully strips away any linked style anchors.
    '               This breaks old or corrupted multi-level list linkages, preparing
    '               the template before clean master styles are reapplied.
    ' SCOPE:        Global background document ListTemplates collection cache.
    ' =========================================================================
    
    Dim lt As ListTemplate
    Dim lvl As ListLevel
    Dim doc As Document
    Dim i As Integer
    
    Set doc = ActiveDocument
    
    ' Disable visual window updates to completely eliminate macro lag.
    ' This prevents Word from attempting to visually redraw and re-paginate the
    ' workspace layout for every single list level processed behind the scenes.
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' CRITICAL CRASH PROTECTION SAFENET
    ' -----------------------------------------------------------------
    ' ARCHITECTURAL VULNERABILITY: ActiveDocument.ListTemplates frequently contains
    ' hidden, read-only system structures or corrupted artifact slots left behind
    ' by Word's tracking engine. When a loop hits one of these protected templates,
    ' reading or writing properties will throw a runtime error or lock the app.
    ' Enforcing On Error Resume Next ensures the script skips locked rows or copy-paste
    ' artifacts smoothly instead of crashing.
    On Error Resume Next
    
    ' -----------------------------------------------------------------
    ' THE TEMPLATE DECOUPLING LOOP
    ' -----------------------------------------------------------------
    ' Loop sequentially through every list template definition stored in the document range
    For Each lt In doc.ListTemplates
        ' Deep-scan all 9 available levels in the native multi-level list hierarchy
        For i = 1 To 9
            ' Bind our level tracking pointer variable
            Set lvl = lt.ListLevels(i)
            
            ' THE DECOUPLING CRITERIA RESET:
            ' Setting the .LinkedStyle property to a blank string ("") forcefully strips
            ' away any linked paragraph style mapping anchors. This ensures
            ' no rogue background formatting or fuzzy outline parameters continue to pollute
            ' the Navigation Pane or Table of Contents fields.
            lvl.LinkedStyle = ""
            
        Next i
    Next lt
    
    ' -----------------------------------------------------------------
    ' INDEX RE-COMPILATION PHASE
    ' -----------------------------------------------------------------
    Dim toc As TableOfContents
    
    ' Loop sequentially through any Table of Contents fields embedded in the text layers.
    ' Since decoupling list templates alters the structural metadata broadcast to the
    ' index engine, we must force a calculation check to update the table layouts immediately.
    For Each toc In doc.TablesOfContents
        ' Force an immediate layout update on the active TOC.
        ' This pulls the newly cleaned paragraph matrices into your index, ensuring those
        ' phantom empty lines and broken dot leaders instantly vanish from your layout page.
        toc.Update
    Next toc
    
    ' Re-enable visual application updates to present the finalized clean layout metrics
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "All list template style links have been successfully cleared!", vbInformation, "Links Reset"
End Sub

Sub Misc_12_Turn_On_Outline_Level_Highlighting()
    ' =========================================================================
    ' MODULE NAME:  Misc_12_Turn_On_Outline_Level_Highlighting
    ' PURPOSE:      Scans the document body to identify paragraphs assigned a
    '               structural outline level (1 to 9), mapping them to an RGB grid.
    '               ADVANCED FIX: Intercepts manually typed OR automated list numbers
    '               (e.g., 1.1, 1.1.1) that have zero structural outline level assigned,
    '               using Word's .ListString buffer to catch invisible digit streams.
    ' SCOPE:        Main document text paragraphs. Automatically protects tables.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' =========================================================================
    
    Dim doc As Document
    Dim para As Paragraph
    Dim counter As Long
    Dim currentLevel As Long
    Dim rgbPalette(1 To 9) As Long
    
    Dim regEx As Object
    Dim paraTxt As String
    Dim cleanTxt As String
    Dim displayDigits As String
    Dim alarmColor As Long
    Dim chCode As Long
    Dim i As Long
    
    Set doc = ActiveDocument
    counter = 0
    
    ' Freeze page layout re-pagination to maximize raw background processing speed
    Application.ScreenUpdating = False
    
    ' Initialize the late-bound VBScript Regular Expressions engine
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        ' UNIVERSAL BOUNDARY PATTERN:
        ' (?:^|[\r\n\x0B]) -> Matches start of paragraph or soft return
        ' [ \t\xA0]* -> Absorbs spaces, tabs, and non-breaking spaces
        ' \d+              -> Matches the first digit cluster
        ' (\.\d+)* -> Loops through nested sub-levels (.1.1)
        ' [.\)-:]* -> Catches trailing special punctuation styles
        ' ([ \t\xA0]+|$)   -> Safely locks boundary to whitespace or end of line
        .Pattern = "(?:^|[\r\n\x0B])[ \t\xA0]*\d+(\.\d+)*[.\)-:]*([ \t\xA0]+|$)"
        .IgnoreCase = True
        .Global = False
        .multiline = True
    End With
    
    ' Define our structural alert color (Alarmed Violet / Magenta)
    alarmColor = RGB(236, 72, 153)
    
    ' -----------------------------------------------------------------
    ' HEX PALETTE REGISTRATION MATRIX
    ' -----------------------------------------------------------------
    rgbPalette(1) = RGB(&HD6, &H4A, &H4A) ' Level 1: Crimson
    rgbPalette(2) = RGB(&HFF, &H8F, &H3F) ' Level 2: Orange
    rgbPalette(3) = RGB(&HE3, &HB4, &H41) ' Level 3: Gold
    rgbPalette(4) = RGB(&H2D, &HD4, &HBF) ' Level 4: Teal
    rgbPalette(5) = RGB(&H3B, &H82, &HF6) ' Level 5: Sapphire Blue
    rgbPalette(6) = RGB(&H63, &H66, &HF1) ' Level 6: Indigo
    rgbPalette(7) = RGB(&HA8, &H55, &HF7) ' Level 7: Amethyst Purple
    rgbPalette(8) = RGB(&HEC, &H48, &H99) ' Level 8: Magenta
    rgbPalette(9) = RGB(&H64, &H74, &H8B) ' Level 9: Slate Grey
    
    ' -----------------------------------------------------------------
    ' THE LANDMARK INTERCEPTION SWEEP
    ' -----------------------------------------------------------------
    For Each para In doc.Paragraphs
        
        ' HARD GUARDRAIL: Skip table components entirely to protect data cells
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Extract the layout tier property integer directly from the paragraph block
            currentLevel = para.OutlineLevel
            
            ' Case A: Valid structural outline landmark detected (Levels 1 through 9)
            If currentLevel >= 1 And currentLevel <= 9 Then
                para.Range.Shading.BackgroundPatternColor = rgbPalette(currentLevel)
                counter = counter + 1
                
            ' Case B: Paragraph behaves like body text structurally, look for hidden list/number flaws
            ElseIf currentLevel = wdOutlineLevelBodyText Then
                
                ' Initialize validation flag for this paragraph pass
                Dim isFlawedHeading As Boolean
                isFlawedHeading = False
                
                ' --- CRITICAL TEST 1: AUTOMATED WORD LIST ENGINES ---
                ' Extract the string representation of Word's automated list field number
                displayDigits = Trim(para.Range.ListFormat.ListString)
                
                ' If displayDigits holds content, Word is dynamically projecting numbers on screen
                If Len(displayDigits) > 0 Then
                    ' Test if those projected numbers match our heading pattern (e.g. 1.1 or 1.)
                    If regEx.Test(displayDigits & " ") Then
                        isFlawedHeading = True
                    End If
                End If
                
                ' --- CRITICAL TEST 2: RAW TEXT TYPED STRINGS ---
                ' Fall back to testing raw typed text if the list engine pass is clear
                If Not isFlawedHeading Then
                    paraTxt = Trim(para.Range.text)
                    
                    ' Sanitize unprintable unicode control variables from the text string
                    cleanTxt = ""
                    For i = 1 To Len(paraTxt)
                        chCode = AscW(Mid(paraTxt, i, 1))
                        If Not (chCode < 32 And chCode <> 9 And chCode <> 11 And chCode <> 13) And _
                           Not (chCode >= &H200B And chCode <= &H200F) Then
                            cleanTxt = cleanTxt & Mid(paraTxt, i, 1)
                        End If
                    Next i
                    
                    ' Run the regex check against our sanitized string text buffer
                    If regEx.Test(cleanTxt) Then
                        isFlawedHeading = True
                    End If
                End If
                
                ' ---------------------------------------------------------
                ' ALARM STAMP DEPLOYMENT
                ' ---------------------------------------------------------
                If isFlawedHeading Then
                    ' Forcefully tint the paragraph with our alert palette selection
                    para.Range.Shading.BackgroundPatternColor = alarmColor
                    counter = counter + 1
                End If
                
            End If
        End If
    Next para
    
    ' Re-enable application rendering to instantly push updates to the screen workspace
    Application.ScreenUpdating = True
    
    ' Signal audit metrics to the operator
    If counter > 0 Then
        MsgBox "Multi-tier diagnostic sweep complete! Colored " & counter & _
               " structural line elements and hidden flaws.", vbInformation, "Diagnostic View Active"
    Else
        MsgBox "Audit complete. No structural outline landmarks or typed number anomalies discovered.", vbInformation, "Sweep Finished"
    End If
End Sub

Sub Misc_13_Turn_Off_Outline_Level_Highlighting()
    ' =========================================================================
    ' MODULE NAME:  Misc_13_Turn_Off_Outline_Level_Highlighting
    ' PURPOSE:      Strips away ONLY the custom multi-color true-color diagnostic
    '               shading layer stamped behind paragraphs by the diagnostic routine.
    '               CRITICAL FEATURE: Preserves all pre-existing, manual, or style-based
    '               document background shading color markers.
    ' SCOPE:        Main document paragraphs layer. Automatically sweeps table matrices.
    ' PERFORMANCE:  Iterates text ranges inside background object streams,
    '               bypassing cursor movement logic to optimize execution speed.
    ' =========================================================================
    
    Dim para As Paragraph
    Dim rgbPalette(1 To 9) As Long
    Dim alarmColor As Long
    Dim currentShading As Long
    Dim isDiagnosticColor As Boolean
    Dim i As Integer
    
    ' Freeze visual application window rendering to prevent layout redraw lag
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' SIGNATURE MATRIX REGISTRATION
    ' -----------------------------------------------------------------
    ' We rebuild the exact same color palette list used in the turning on macro.
    ' This gives the clearing sweeper a lookup signature to compare against.
    alarmColor = RGB(236, 72, 153)
    
    rgbPalette(1) = RGB(&HD6, &H4A, &H4A)
    rgbPalette(2) = RGB(&HFF, &H8F, &H3F)
    rgbPalette(3) = RGB(&HE3, &HB4, &H41)
    rgbPalette(4) = RGB(&H2D, &HD4, &HBF)
    rgbPalette(5) = RGB(&H3B, &H82, &HF6)
    rgbPalette(6) = RGB(&H63, &H66, &HF1)
    rgbPalette(7) = RGB(&HA8, &H55, &HF7)
    rgbPalette(8) = RGB(&HEC, &H48, &H99)
    rgbPalette(9) = RGB(&H64, &H74, &H8B)
    
    ' -----------------------------------------------------------------
    ' TARGETED RESET SWEEP
    ' -----------------------------------------------------------------
    For Each para In ActiveDocument.Paragraphs
        
        ' Safety Guardrail: Insulate table data boundaries from formatting modifications
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Capture the active background pattern color value of the paragraph line
            currentShading = para.Range.Shading.BackgroundPatternColor
            
            ' If the line has some form of color treatment, audit the color signature
            If currentShading <> wdColorAutomatic Then
                isDiagnosticColor = False
                
                ' Check 1: Does it match our rogue manual list heading alarm color?
                If currentShading = alarmColor Then
                    isDiagnosticColor = True
                Else
                    ' Check 2: Does it match any of our 9 structural tier color signatures?
                    For i = 1 To 9
                        If currentShading = rgbPalette(i) Then
                            isDiagnosticColor = True
                            Exit For
                        End If
                    Next i
                End If
                
                ' CRITICAL INSULATION OVERRIDE:
                ' Only reset the text line back to transparent if the color matches
                ' our signature matrix perfectly. Otherwise, skip it entirely.
                If isDiagnosticColor Then
                    para.Range.Shading.BackgroundPatternColor = wdColorAutomatic
                End If
                
            End If
        End If
    Next para
    
    ' Re-enable application layout updates to instantly present the clean reporting canvas
    Application.ScreenUpdating = True
    
    ' Notify user upon successful completion
    MsgBox "Diagnostic multi-level shading cleared successfully! Intentional document highlights preserved.", vbInformation, "Reset Complete"
End Sub


Sub Misc_14_Remove_Highlight_Target_Words()
'=============================================================================
' Module:      Misc_14_Remove_Highlight_Target_Words
' Purpose:     Sweeps the main text body and embedded tables to strip background
'              highlight overlays from an array of target keywords[cite: 2351, 2352].
' Settings:    Case-insensitive fragment parsing (e.g., handles "fig" inside "figure")[cite: 2353, 2372].
' Scope:       Main document body and embedded cell layers ONLY[cite: 2352, 2360].
'=============================================================================
    Dim docRange As Range
    Dim wordList As Variant
    Dim targetWord As Variant
    
    ' Target array configuration matrix
    wordList = Array("tab", "fig", "annex", "plate", "photo")
    
    ' Speed Optimization: Suppress display updates to cut processing lag
    Application.ScreenUpdating = False
    
    ' Iterate sequentially through each keyword defined in the configuration array
    For Each targetWord In wordList
        
        ' RE-INSTANTIATION GATE: Pointers shift during .Execute passes; the range
        ' must be reset on every word iteration to scan from page 1 to EOF.
        Set docRange = ActiveDocument.Content
        
        With docRange.Find
            ' Clear residual manual settings or search criteria from the layout memory buffer
            .ClearFormatting
            .replacement.ClearFormatting
            
            .text = targetWord
            
            ' PERFORMANCE FILTER: Targets text strings that ALREADY carry a highlight flag,
            ' accelerating execution across multi-page files.
            '.Highlight = True
            .MatchCase = False
            .MatchWholeWord = False
            
            ' BOUNDARY GUARDRAIL: Halts search engine at the text endpoint,
            ' preventing infinite wrap-around duplication loops.
            .Wrap = wdFindStop
            
            ' THE CRITICAL INVERSION FIX: Instructs the layout engine to completely
            ' erase the background color index parameter from the text match[cite: 2303].
            .replacement.Highlight = False
            
            ' Execute global search-and-replace sweep over the active text layer
            .Execute Replace:=wdReplaceAll
        End With
    Next targetWord
    
    ' Restore standard application window rendering metrics
    Application.ScreenUpdating = True
    
    ' Signal execution completion to the operator
    MsgBox "Target words successfully stripped of highlighting within the main document and tables.", _
           vbInformation, "Highlight Removal Complete"
End Sub

Sub Misc_15_Delete_Image_Globally_By_Alt_Text()
' ==============================================================================
' MODULE NAME    : Misc_15_Delete_Image_Globally_By_Alt_Text
' PURPOSE        : Searches the main body text, embedded tables, and all section
'                  headers/footers for images matching ANY of the hardcoded Alt Text
'                  or Title tags, and permanently deletes them globally.
' COMPATIBILITY  : Microsoft Word (All Versions)
' ==============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim targetTags() As Variant
    Dim i As Long
    Dim imgDeletedCount As Long
    
    ' --------------------------------------------------------------------------
    ' CONFIGURATION: Define your array of target Alt Text / Title tags here.
    ' You can list one or multiple tags separated by commas.
    ' Examples: Array("DeleteMe") or Array("DeleteMe", "OldLogo", "Draft_Pic")
    ' --------------------------------------------------------------------------
    targetTags = Array("Melkamu Signature", "ABCE Stamp")
    
    ' Initialize references and counter
    Set doc = ActiveDocument
    imgDeletedCount = 0
    
    ' Freeze live screen redraws to suppress visual jitter and maximize performance
    Application.ScreenUpdating = False
    
    ' Inline error bypass to handle locked or uneditable shape properties smoothly
    On Error Resume Next
    
    ' ==========================================================================
    ' PASS 1: Main Text Story (Includes Document Body Text and Embedded Tables)
    ' ==========================================================================
    
    ' 1A. Floating Shapes (Main Body & Tables)
    ' Traversal Rule: Count backward (Step -1) to prevent array index shifting
    ' when shapes are deleted from the collection.
    For i = doc.Shapes.Count To 1 Step -1
        With doc.Shapes(i)
            If IsMatchingTag(.Title, .AlternativeText, targetTags) Then
                .Delete
                imgDeletedCount = imgDeletedCount + 1
            End If
        End With
    Next i
    
    ' 1B. Inline Shapes (Main Body & Tables)
    For i = doc.InlineShapes.Count To 1 Step -1
        With doc.InlineShapes(i)
            If IsMatchingTag(.Title, .AlternativeText, targetTags) Then
                .Delete
                imgDeletedCount = imgDeletedCount + 1
            End If
        End With
    Next i
    
    ' ==========================================================================
    ' PASS 2: Headers and Footers Across All Document Sections
    ' Multi-Layer Protection: Explicitly sweeps primary, first-page, and odd/even
    ' header/footer sub-layers to wake up unlinked or dormant layout sections.
    ' ==========================================================================
    For Each sec In doc.Sections
        
        ' Process Section Headers
        For Each hf In sec.Headers
            If hf.Exists Then
                Call DeleteImagesInHeaderFooterRange(hf, targetTags, imgDeletedCount)
            End If
        Next hf
        
        ' Process Section Footers
        For Each hf In sec.Footers
            If hf.Exists Then
                Call DeleteImagesInHeaderFooterRange(hf, targetTags, imgDeletedCount)
            End If
        Next hf
        
    Next sec
    
    ' Re-enable application UI screen updating
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    ' Display execution completion summary
    MsgBox "Cleanup Complete!" & vbCrLf & _
           "Target Tags Searched: " & Join(targetTags, ", ") & vbCrLf & _
           "Total instances of target image(s) deleted: " & imgDeletedCount, _
           vbInformation, "Delete Image Summary"
End Sub

Private Sub DeleteImagesInHeaderFooterRange(hf As HeaderFooter, tags() As Variant, ByRef deleteCount As Long)
' ==============================================================================
' PARENT MODULE     : Misc_15_Delete_Image_Globally_By_Alt_Text
' HELPER SUBROUTINE : DeleteImagesInHeaderFooterRange
' PURPOSE           : Safely sweeps floating (Shapes) and inline (InlineShapes)
'                     image collections inside a specific HeaderFooter layer.
' ==============================================================================
    Dim i As Long
    
    ' 1. Delete Floating Shapes in Header/Footer
    For i = hf.Shapes.Count To 1 Step -1
        With hf.Shapes(i)
            If IsMatchingTag(.Title, .AlternativeText, tags) Then
                .Delete
                deleteCount = deleteCount + 1
            End If
        End With
    Next i
    
    ' 2. Delete Inline Shapes in Header/Footer Range
    ' Routing through hf.Range.InlineShapes explicitly exposes inline images
    ' embedded inside header/footer text containers.
    For i = hf.Range.InlineShapes.Count To 1 Step -1
        With hf.Range.InlineShapes(i)
            If IsMatchingTag(.Title, .AlternativeText, tags) Then
                .Delete
                deleteCount = deleteCount + 1
            End If
        End With
    Next i
End Sub


Private Function IsMatchingTag(imgTitle As String, imgAltText As String, tags() As Variant) As Boolean
' ==============================================================================
' PARENT MODULE     : Misc_15_Delete_Image_Globally_By_Alt_Text
' HELPER FUNCTION   : IsMatchingTag
' PURPOSE           : Checks if either the Title or AlternativeText contains
'                     ANY of the tags specified in the targetTags array.
' ==============================================================================
    Dim vTag As Variant
    
    IsMatchingTag = False
    
    For Each vTag In tags
        If Trim(CStr(vTag)) <> "" Then
            If InStr(1, imgTitle, CStr(vTag), vbTextCompare) > 0 Or _
               InStr(1, imgAltText, CStr(vTag), vbTextCompare) > 0 Then
                IsMatchingTag = True
                Exit Function
            End If
        End If
    Next vTag
End Function

Sub Misc_16_Delete_Selected_Images_Globally_By_Size()
' ==============================================================================
' MODULE NAME    : Misc_16_Delete_Selected_Images_Globally_By_Size
' PURPOSE        : Reads all selected images (or floating shapes), captures their
'                  width, height, and type properties, and deletes all matching
'                  instances across the main document, tables, headers, and footers.
' COMPATIBILITY  : Microsoft Word (All Versions)
' ==============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim hf As HeaderFooter
    
    ' Dynamic storage for target shape profiles
    Dim targetWidths() As Single
    Dim targetHeights() As Single
    Dim targetTypes() As Long
    Dim targetCount As Long
    
    Dim i As Long
    Dim imgDeletedCount As Long
    Dim tolerance As Single
    
    Set doc = ActiveDocument
    imgDeletedCount = 0
    targetCount = 0
    tolerance = 0.5 ' Precision tolerance in points for Word layout rounding
    
    ' ==========================================================================
    ' PHASE 1: Capture Target Image Metrics from User Selection
    ' ==========================================================================
    If Selection.Type = wdSelectionInlineShape Then
        ' Single Inline Image Selected
        targetCount = Selection.InlineShapes.Count
        ReDim targetWidths(1 To targetCount)
        ReDim targetHeights(1 To targetCount)
        ReDim targetTypes(1 To targetCount)
        
        For i = 1 To targetCount
            targetWidths(i) = Selection.InlineShapes(i).Width
            targetHeights(i) = Selection.InlineShapes(i).Height
            targetTypes(i) = Selection.InlineShapes(i).Type
        Next i
        
    ElseIf Selection.ShapeRange.Count > 0 Then
        ' One or More Floating Shapes Selected
        targetCount = Selection.ShapeRange.Count
        ReDim targetWidths(1 To targetCount)
        ReDim targetHeights(1 To targetCount)
        ReDim targetTypes(1 To targetCount)
        
        For i = 1 To targetCount
            targetWidths(i) = Selection.ShapeRange(i).Width
            targetHeights(i) = Selection.ShapeRange(i).Height
            targetTypes(i) = Selection.ShapeRange(i).Type
        Next i
        
    Else
        MsgBox "Please select one or more target images first!", _
               vbExclamation, "No Selection Detected"
        Exit Sub
    End If
    
    ' Freeze live screen redraws to suppress visual jitter and maximize performance
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' ==========================================================================
    ' PHASE 2: Sweep Main Document Text and Embedded Tables
    ' Reverse Traversal Rule (Step -1) prevents index skipping upon deletion.
    ' ==========================================================================
    
    ' 2A. Sweep Floating Shapes (Main Body & Tables)
    For i = doc.Shapes.Count To 1 Step -1
        With doc.Shapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, targetWidths, targetHeights, targetTypes, targetCount, tolerance) Then
                .Delete
                imgDeletedCount = imgDeletedCount + 1
            End If
        End With
    Next i
    
    ' 2B. Sweep Inline Shapes (Main Body & Tables)
    For i = doc.InlineShapes.Count To 1 Step -1
        With doc.InlineShapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, targetWidths, targetHeights, targetTypes, targetCount, tolerance) Then
                .Delete
                imgDeletedCount = imgDeletedCount + 1
            End If
        End With
    Next i
    
    ' ==========================================================================
    ' PHASE 3: Sweep Headers and Footers Across All Document Sections
    ' Multi-Layer Protection: Explicitly checks primary, first-page, and odd/even
    ' sub-layers to wake up unlinked or dormant layout sections.
    ' ==========================================================================
    For Each sec In doc.Sections
        
        ' Process Section Headers
        For Each hf In sec.Headers
            If hf.Exists Then
                Call DeleteMatchingImagesInHeaderFooter(hf, targetWidths, targetHeights, targetTypes, targetCount, tolerance, imgDeletedCount)
            End If
        Next hf
        
        ' Process Section Footers
        For Each hf In sec.Footers
            If hf.Exists Then
                Call DeleteMatchingImagesInHeaderFooter(hf, targetWidths, targetHeights, targetTypes, targetCount, tolerance, imgDeletedCount)
            End If
        Next hf
        
    Next sec
    
    ' Re-enable application UI screen updating
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    ' Display execution completion summary
    MsgBox "Cleanup Complete!" & vbCrLf & _
           "Target Profile(s) Tracked: " & targetCount & vbCrLf & _
           "Total matching image instances deleted: " & imgDeletedCount, _
           vbInformation, "Delete Image Summary"
End Sub

Private Sub DeleteMatchingImagesInHeaderFooter(hf As HeaderFooter, _
                                              widths() As Single, _
                                              heights() As Single, _
                                              types() As Long, _
                                              tCount As Long, _
                                              tol As Single, _
                                              ByRef deleteCount As Long)
' ==============================================================================
' PARENT MODULE     : Misc_16_Delete_Selected_Images_Globally_By_Size
' HELPER SUBROUTINE : DeleteMatchingImagesInHeaderFooter
' PURPOSE           : Safely sweeps floating (Shapes) and inline (InlineShapes)
'                     image collections inside a specific HeaderFooter layer.
' ==============================================================================
    Dim i As Long
    
    ' 1. Floating Shapes in Header/Footer
    For i = hf.Shapes.Count To 1 Step -1
        With hf.Shapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, widths, heights, types, tCount, tol) Then
                .Delete
                deleteCount = deleteCount + 1
            End If
        End With
    Next i
    
    ' 2. Inline Shapes in Header/Footer Range
    ' Calling hf.Range.InlineShapes explicitly exposes inline images embedded in header/footer text.
    For i = hf.Range.InlineShapes.Count To 1 Step -1
        With hf.Range.InlineShapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, widths, heights, types, tCount, tol) Then
                .Delete
                deleteCount = deleteCount + 1
            End If
        End With
    Next i
End Sub


Sub Misc_17_Apply_AltText_To_Images_Globally_By_Size()
' ==============================================================================
' MODULE NAME    : Misc_17_Apply_AltText_To_Images_Globally_By_Size
' PURPOSE        : Reads selected image(s) or shape(s) to capture width, height,
'                  and type profiles. Prompts the user for Title and Alt Text,
'                  then updates all matching image instances across the main
'                  document body, embedded tables, headers, and footers.
' COMPATIBILITY  : Microsoft Word (All Versions)
' ==============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim hf As HeaderFooter
    
    ' Dynamic storage for target shape profiles
    Dim targetWidths() As Single
    Dim targetHeights() As Single
    Dim targetTypes() As Long
    Dim targetCount As Long
    
    Dim i As Long
    Dim imgUpdatedCount As Long
    Dim tolerance As Single
    Dim newTitle As String
    Dim newAltText As String
    
    Set doc = ActiveDocument
    imgUpdatedCount = 0
    targetCount = 0
    tolerance = 0.5 ' Precision tolerance in points for Word layout rounding
    
    ' ==========================================================================
    ' PHASE 1: Capture Target Image Metrics from User Selection
    ' ==========================================================================
    If Selection.Type = wdSelectionInlineShape Then
        ' Single Inline Image Selected
        targetCount = Selection.InlineShapes.Count
        ReDim targetWidths(1 To targetCount)
        ReDim targetHeights(1 To targetCount)
        ReDim targetTypes(1 To targetCount)
        
        For i = 1 To targetCount
            targetWidths(i) = Selection.InlineShapes(i).Width
            targetHeights(i) = Selection.InlineShapes(i).Height
            targetTypes(i) = Selection.InlineShapes(i).Type
        Next i
        
    ElseIf Selection.ShapeRange.Count > 0 Then
        ' One or More Floating Shapes Selected
        targetCount = Selection.ShapeRange.Count
        ReDim targetWidths(1 To targetCount)
        ReDim targetHeights(1 To targetCount)
        ReDim targetTypes(1 To targetCount)
        
        For i = 1 To targetCount
            targetWidths(i) = Selection.ShapeRange(i).Width
            targetHeights(i) = Selection.ShapeRange(i).Height
            targetTypes(i) = Selection.ShapeRange(i).Type
        Next i
        
    Else
        MsgBox "Please select one or more target images first!", _
               vbExclamation, "No Selection Detected"
        Exit Sub
    End If

    ' ==========================================================================
    ' PHASE 2: Prompt User for Title & Alt Text Description
    ' ==========================================================================
    newTitle = InputBox("Enter the TITLE for matching same-sized images:", "Set Image Title")
    If StrPtr(newTitle) = 0 Then Exit Sub ' User clicked Cancel
    
    newAltText = InputBox("Enter the ALT TEXT (Description) for matching same-sized images:", "Set Alt Text Description")
    If StrPtr(newAltText) = 0 Then Exit Sub ' User clicked Cancel
    
    ' Freeze live screen redraws to suppress visual jitter and maximize performance
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' ==========================================================================
    ' PHASE 3: Sweep Main Document Text and Embedded Tables
    ' ==========================================================================
    
    ' 3A. Sweep Floating Shapes (Main Body & Tables)
    For i = 1 To doc.Shapes.Count
        With doc.Shapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, targetWidths, targetHeights, targetTypes, targetCount, tolerance) Then
                .Title = newTitle
                .AlternativeText = newAltText
                imgUpdatedCount = imgUpdatedCount + 1
            End If
        End With
    Next i
    
    ' 3B. Sweep Inline Shapes (Main Body & Tables)
    For i = 1 To doc.InlineShapes.Count
        With doc.InlineShapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, targetWidths, targetHeights, targetTypes, targetCount, tolerance) Then
                .Title = newTitle
                .AlternativeText = newAltText
                imgUpdatedCount = imgUpdatedCount + 1
            End If
        End With
    Next i
    
    ' ==========================================================================
    ' PHASE 4: Sweep Headers and Footers Across All Document Sections
    ' Multi-Layer Protection: Explicitly checks primary, first-page, and odd/even
    ' sub-layers to wake up unlinked or dormant layout sections.
    ' ==========================================================================
    For Each sec In doc.Sections
        
        ' Process Section Headers
        For Each hf In sec.Headers
            If hf.Exists Then
                Call ApplyAltTextToHeaderFooter(hf, newTitle, newAltText, targetWidths, targetHeights, targetTypes, targetCount, tolerance, imgUpdatedCount)
            End If
        Next hf
        
        ' Process Section Footers
        For Each hf In sec.Footers
            If hf.Exists Then
                Call ApplyAltTextToHeaderFooter(hf, newTitle, newAltText, targetWidths, targetHeights, targetTypes, targetCount, tolerance, imgUpdatedCount)
            End If
        Next hf
        
    Next sec
    
    ' Re-enable application UI screen updating
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    ' Display execution completion summary
    MsgBox "Alt Text Update Complete!" & vbCrLf & _
           "Target Profile(s) Tracked: " & targetCount & vbCrLf & _
           "Total matching image instances updated: " & imgUpdatedCount, _
           vbInformation, "Update Alt Text Summary"
End Sub

Private Sub ApplyAltTextToHeaderFooter(hf As HeaderFooter, _
                                       titleText As String, _
                                       altText As String, _
                                       widths() As Single, _
                                       heights() As Single, _
                                       types() As Long, _
                                       tCount As Long, _
                                       tol As Single, _
                                       ByRef updateCount As Long)
' ==============================================================================
' PARENT MODULE     : Misc_17_Apply_AltText_To_Images_Globally_By_Size
' HELPER SUBROUTINE : ApplyAltTextToHeaderFooter
' PURPOSE           : Safely sweeps floating (Shapes) and inline (InlineShapes)
'                     image collections inside a specific HeaderFooter layer
'                     and updates Title and AlternativeText properties.
' ==============================================================================
    Dim i As Long
    
    ' 1. Floating Shapes in Header/Footer
    For i = 1 To hf.Shapes.Count
        With hf.Shapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, widths, heights, types, tCount, tol) Then
                .Title = titleText
                .AlternativeText = altText
                updateCount = updateCount + 1
            End If
        End With
    Next i
    
    ' 2. Inline Shapes in Header/Footer Range
    ' Calling hf.Range.InlineShapes explicitly exposes inline images embedded in header/footer text.
    For i = 1 To hf.Range.InlineShapes.Count
        With hf.Range.InlineShapes(i)
            If IsMatchingSizeAndType(.Width, .Height, .Type, widths, heights, types, tCount, tol) Then
                .Title = titleText
                .AlternativeText = altText
                updateCount = updateCount + 1
            End If
        End With
    Next i
End Sub

Private Function IsMatchingSizeAndType(imgWidth As Single, _
                                       imgHeight As Single, _
                                       imgType As Long, _
                                       widths() As Single, _
                                       heights() As Single, _
                                       types() As Long, _
                                       tCount As Long, _
                                       tol As Single) As Boolean
' ==============================================================================
' PARENT MODULES     : Misc_16_Delete_Selected_Images_Globally_By_Size
'                    : Misc_17_Apply_AltText_To_Images_Globally_By_Size
' HELPER FUNCTION    : IsMatchingSizeAndType
' PURPOSE            : Evaluates an image's dimensions against the captured target
'                      arrays using point precision tolerance.
' ==============================================================================
    Dim k As Long
    IsMatchingSizeAndType = False
    
    For k = 1 To tCount
        If imgType = types(k) Then
            If Abs(imgWidth - widths(k)) <= tol And _
               Abs(imgHeight - heights(k)) <= tol Then
                IsMatchingSizeAndType = True
                Exit Function
            End If
        End If
    Next k
End Function