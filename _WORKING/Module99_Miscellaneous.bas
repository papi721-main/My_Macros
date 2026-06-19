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
    ' background processing speed on long, multi-section documents.
    Application.ScreenUpdating = False
    
    ' =========================================================================
    ' PASS 1: THE GLOBAL SWEEP (Catches Main Text, Tables, and Footnotes)
    ' =========================================================================
    ' This loop scans the background StoryRanges collection. It is the most 
    ' efficient way to bypass cursor selection and hit major text blocks.
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
    ' footer layer into memory unless they are actively visible or opened by the user.
    ' If "Different First Page" or "Different Odd & Even Pages" layout flags are enabled, 
    ' unlinked header/footer canvases remain completely dormant in the background cache.
    ' Because Pass 1 loops right past dormant stories, Pass 2 explicitly declares a 
    ' nested section loop to force the layout processor to evaluate all layout sub-layers.
    For Each sec In doc.Sections
        
        ' -----------------------------------------------------------------
        ' Sub-Phase A: Section Headers
        ' -----------------------------------------------------------------
        ' Step through the section's Header collection (First Page, Even Pages, Primary).
        For Each compView In sec.Headers
            ' The .Exists safety gate prevents Word from throwing a runtime error
            ' if the sub-layer is structurally unassigned or inactive.
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
    ' MODULE NAME:  Misc_8_Trim_Global_Table_Paragraph_Marks
    ' PURPOSE:      Scans every table globally across the active document to identify
    '               and forcefully strip out manually inserted empty paragraph breaks
    '               (vbCr / ¶) hanging at the absolute top and bottom of table cells.
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
                
                ' If the text string length is exactly 1, it holds nothing but an empty carriage return (¶).
                If Len(pga.Range.Text) = 1 Then
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
                    If Len(pga.Range.Text) = 1 Then
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
    '               (vbCr / ¶) hanging at the absolute top and bottom of cells.
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
            
            ' If the text string length is exactly 1, it holds nothing but an empty carriage return (¶).
            If Len(pga.Range.Text) = 1 Then
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
                If Len(pga.Range.Text) = 1 Then
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
    '               CRITICAL ADDITION: Also intercepts paragraphs containing manually
    '               typed heading numbers (e.g., 1.1, 1.1.1) that have zero structural
    '               outline level assigned, stamping them with an Alarmed Violet tone.
    ' SCOPE:        Main document text paragraphs. Automatically protects tables.
    ' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
    ' =========================================================================
    
    Dim doc As Document
    Dim para As Paragraph
    Dim counter As Long
    Dim currentLevel As Long
    Dim rgbPalette(1 To 9) As Long
    
    ' Regular Expression variables for capturing manually typed heading markers
    Dim regEx As Object
    Dim paraTxt As String
    Dim alarmColor As Long
    
    Set doc = ActiveDocument
    counter = 0
    
    ' Freeze page layout re-pagination to maximize raw background processing speed
    Application.ScreenUpdating = False
    
    ' Initialize the late-bound VBScript Regular Expressions engine
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        ' Pattern looks for starts of lines matching digits separated by periods (e.g. 1., 1.1, 1.1.1.)
        ' followed immediately by a space or a tab marker.
        .Pattern = "^[ \t]*\d+(\.\d+)*[.\)-:]*[ \t]+"
        .IgnoreCase = True
        .Global = False
    End With
    
    ' Define our structural alert color (Alarmed Violet / Magenta) for manual numbers lacking an outline tier
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
                
                ' Route directly through the paragraph's Shading interface to apply 24-bit color huing
                para.Range.Shading.BackgroundPatternColor = rgbPalette(currentLevel)
                counter = counter + 1
                
            ' Case B: Paragraph behaves like body text structurally, but might contain an un-styled number string
            ElseIf currentLevel = wdOutlineLevelBodyText Then
                paraTxt = para.Range.text
                
                ' Execute the regex validator evaluation pass on the raw string line
                If regEx.Test(paraTxt) Then
                    
                    ' Verify that this isn't an automated Word List field layout block to prevent false positives
                    If para.Range.ListFormat.ListType = wdListNoNumbering Then
                        ' Apply the structural flaw alert highlight pattern
                        para.Range.Shading.BackgroundPatternColor = alarmColor
                        counter = counter + 1
                    End If
                    
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
    ' PURPOSE:      Strips away the multi-color true-color diagnostic shading layer
    '               stamped behind paragraphs, resetting text lines back to transparent.
    ' SCOPE:        Main document paragraphs layer. Automatically sweeps table matrices.
    ' PERFORMANCE:  Iterates text ranges inside background object streams,
    '               bypassing cursor movement logic to optimize execution speed.
    ' =========================================================================
    
    Dim para As Paragraph
    
    ' Freeze visual application window rendering to prevent layout redraw lag
    Application.ScreenUpdating = False
    
    ' -----------------------------------------------------------------
    ' THE TRANSMUTATION RESET SWEEP
    ' -----------------------------------------------------------------
    For Each para In ActiveDocument.Paragraphs
        
        ' Safety Guardrail: Insulate table data boundaries from formatting modifications
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Verify if the paragraph range holds an active background shading assignment.
            ' wdColorAutomatic represents default structural transparency in Word's layout engine.
            If para.Range.Shading.BackgroundPatternColor <> wdColorAutomatic Then
                
                ' Erase the background color matrix, stripping the tint cleanly from the text span
                para.Range.Shading.BackgroundPatternColor = wdColorAutomatic
                
            End If
        End If
    Next para
    
    ' Re-enable application layout updates to instantly present the clean reporting canvas
    Application.ScreenUpdating = True
    
    ' Notify user upon successful completion
    MsgBox "Diagnostic multi-level shading cleared successfully!", vbInformation, "Reset Complete"
End Sub

