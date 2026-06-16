Sub Clear_All_Highlighting_Globally()
'=============================================================================
' PURPOSES: Completely clears all background text highlighting across every
'           structural layer of the active document, including stubborn 
'           unlinked headers/footers.
' SCOPE:    Main document text, tables, headers, footers, textboxes, footnotes.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim compView As HeaderFooter
    Dim story As Range

    Set doc = ActiveDocument
    
    ' Optimize performance and prevent layout screen flicker
    Application.ScreenUpdating = False
    
    ' PASS 1: The Global Sweep (Catches Main Text, Tables, and Footnotes)
    For Each story In doc.StoryRanges
        Do
            story.HighlightColorIndex = wdNoHighlight
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next story
    
    ' PASS 2: Deep Section Penetration (Forces dormant headers/footers awake)
    For Each sec In doc.Sections
        
        ' Process all Headers across this section
        For Each compView In sec.Headers
            If compView.Exists Then
                compView.Range.HighlightColorIndex = wdNoHighlight
            End If
        Next compView
        
        ' Process all Footers across this section
        For Each compView In sec.Footers
            If compView.Exists Then
                compView.Range.HighlightColorIndex = wdNoHighlight
            End If
        Next compView
        
    Next sec
    
    ' Restore native layout rendering rules
    Application.ScreenUpdating = True
    
    MsgBox "All text highlighting has been forcefully stripped from all layers, including headers and footers.", _
           vbInformation, "Global Clear Successful"
End Sub

Sub Highlight_Target_Words()
'=============================================================================
' PURPOSES: Searches the main body text layer for an array of target keywords
'           and applies a standardized green highlight overlay to them.
' SCOPE:    Main text body and embedded tables ONLY. Excludes headers/footers.
' SETTINGS: Case-insensitive match. Evaluates word fragments (e.g. finds "fig" inside "figure").
'=============================================================================
    Dim docRange As Range
    Dim wordList As Variant
    Dim targetWord As Variant
    
    ' EXPLICIT CONFIGURATION: Define your target word list here
    wordList = Array("tab", "fig", "annex", "plate")
    
    ' Optimize performance by minimizing layout processor overhead
    Application.ScreenUpdating = False
    
    ' Define the default color index used for replacement formatting
    Options.DefaultHighlightColorIndex = wdBrightGreen
    
    ' Iterate sequentially through each keyword in the configuration array
    For Each targetWord In wordList
        
        ' Always re-instantiate the range to span the complete main content area
        ' Note: This encompasses all main text and body tables, ignoring headers/footers
        Set docRange = ActiveDocument.Content
        
        With docRange.Find
            ' Wipe any lingering search or replacement parameters from memory
            .ClearFormatting
            .Replacement.ClearFormatting
            
            ' Search and execution parameters
            .Text = targetWord
            .MatchCase = False       ' FALSE: Enforces case-insensitive parsing
            .MatchWholeWord = False  ' FALSE: Matches partial fragments (e.g., "fig" catches "figure")
                                     ' Change to True if you want EXACT complete word matches only.
            .Wrap = wdFindStop       ' Ensure the range processes top-to-bottom exactly once
            
            ' Tell the execution block to replace instances with highlighting
            .Replacement.Highlight = True
            
            ' Fire the find engine to execute a global bulk replacement pass
            .Execute Replace:=wdReplaceAll
        End With
    Next targetWord
    
    ' Restore standard application layout rendering
    Application.ScreenUpdating = True
    
    ' Notify user upon successful completion
    MsgBox "Target words successfully highlighted within the main document and tables.", _
           vbInformation, "Highlight Processing Complete"
End Sub

Sub Fix_Common_Misspellings()
'=============================================================================
' PURPOSE: Automatically identifies and replaces a pre-configured dictionary
'          of common misspellings across every single layer of the document.
' SCOPE:   Main body text, tables, headers, footers, textboxes, and footnotes.
' RULES:   Case-insensitive matching, but strictly enforces whole-word checks 
'          to avoid accidentally corrupting longer, correctly spelled words.
'=============================================================================
    Dim doc As Document
    Dim story As Range
    Dim errorMap As Object
    Dim incorrectWord As Variant
    Dim correctWord As String
    
    Set doc = ActiveDocument
    
    ' Optimize engine performance by turning off screen rendering/flicker
    Application.ScreenUpdating = False
    
    ' Instantiate a high-speed Dictionary object to store spelling pairs
    Set errorMap = CreateObject("Scripting.Dictionary")
    
    '=========================================================================
    ' CONFIGURATION DICTIONARY: Add your custom spelling mappings here
    ' Syntax: errorMap.Add "WRONG_WORD", "CORRECT_WORD"
    '=========================================================================
    errorMap.Add "wereda", "woreda"
    errorMap.Add "weredas", "woredas"
    errorMap.Add "tabel", "table"
    errorMap.Add "programme", "program"
    errorMap.Add "labour", "labor"
    '=========================================================================
    
    ' Loop sequentially through every key (incorrect word) in our dictionary
    For Each incorrectWord In errorMap.Keys
        correctWord = errorMap(incorrectWord)
        
        ' Deep-sweep through all layout story layers (Main text, headers, etc.)
        For Each story In doc.StoryRanges
            Do
                With story.Find
                    ' Wipe any lingering search parameters from memory
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    
                    ' Configure replacement criteria
                    .Text = CStr(incorrectWord)
                    .Replacement.Text = correctWord
                    
                    ' Crucial Safety Controls
                    .MatchCase = False       ' FALSE: Case-insensitive (catches "Teh", "TEH", "teh")
                    .MatchWholeWord = True   ' TRUE: Prevents ruining words like "teh" inside "tech"
                    .Wrap = wdFindStop       ' Clear the active range block exactly once
                    
                    ' Fire the global bulk replacement execution pass
                    .Execute Replace:=wdReplaceAll
                End With
                
                ' Traverse linked story shapes (e.g., text boxes inside footers)
                Set story = story.NextStoryRange
            Loop Until story Is Nothing
        Next story
    Next incorrectWord
    
    ' Restore standard application layout rendering
    Application.ScreenUpdating = True
    
    ' Notify user upon successful completion
    MsgBox "Spelling correction sweep complete across all document layers!", _
           vbInformation, "Auto-Correction Successful"
End Sub

Sub Trim_Headings()
'=============================================================================
' Name: Trim_Headings
' Purpose: Sweeps the document to clean up margins around heading structures.
'          Strips out leading/trailing spaces, rogue tab characters, and
'          unwanted trailing periods or special formatting artifacts.
' SAFETY: Automatically protects tables and ignores generic body text layers.
'=============================================================================
    Dim doc As Document
    Dim para As Paragraph
    Dim txtRange As Range
    Dim paraText As String
    Dim cleanText As String
    Dim originalText As String
    
    Set doc = ActiveDocument
    
    ' Speed Optimization: Turn off window rendering to process instantly in background
    Application.ScreenUpdating = False
    
    ' Enable error handling trap to safeguard application environment
    On Error GoTo ErrorHandler

    ' Iterate paragraph-by-paragraph through the active document text layer
    For Each para In doc.Paragraphs
        
        ' HARD GUARDRAIL: Skip table contents completely to protect data cells
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Process execution ONLY if the paragraph is an active structural heading (Levels 1 to 9)
            If para.OutlineLevel >= 1 And para.OutlineLevel <= 9 Then
                
                Set txtRange = para.Range
                paraText = txtRange.text
                originalText = paraText
                
                ' Strip out Word's native paragraph mark character (^p) for text string analysis
                If Right(paraText, 1) = vbCr Then paraText = Left(paraText, Len(paraText) - 1)
                
                ' ---------------------------------------------------------------------
                ' EDGE CLEANING LAYER: LEADING & TRAILING CHARACTERS
                ' ---------------------------------------------------------------------
                cleanText = paraText
                
                ' 1. Loop to continuously strip leading whitespaces, tabs, or rogue punctuation
                Do
                    Dim initialLenAsLeading As Long
                    initialLenAsLeading = Len(cleanText)
                    
                    cleanText = Trim(cleanText)
                    
                    ' Strip rogue leading tabs
                    If Left(cleanText, 1) = vbTab Then cleanText = Mid(cleanText, 2)
                    
                    ' Optional: Add characters here if you find other leading typos (e.g., rogue dots)
                    If Left(cleanText, 1) = "." Or Left(cleanText, 1) = "-" Then cleanText = Mid(cleanText, 2)
                    
                Loop Until Len(cleanText) = initialLenAsLeading Or Len(cleanText) = 0
                
                ' 2. Loop to continuously strip trailing spaces, tabs, trailing dots, or hyphens
                Do
                    Dim initialLenAsTrailing As Long
                    initialLenAsTrailing = Len(cleanText)
                    
                    cleanText = Trim(cleanText)
                    
                    ' Strip rogue trailing tabs
                    If Right(cleanText, 1) = vbTab Then cleanText = Left(cleanText, Len(cleanText) - 1)
                    
                    ' Clear trailing periods (e.g., converts "10.5. Title." to "10.5. Title")
                    If Right(cleanText, 1) = "." Or Right(cleanText, 1) = "-" Or Right(cleanText, 1) = ":" Then
                        cleanText = Left(cleanText, Len(cleanText) - 1)
                    End If
                    
                Loop Until Len(cleanText) = initialLenAsTrailing Or Len(cleanText) = 0
                
                ' ---------------------------------------------------------------------
                ' RE-STAMPING ENGINE (IF EDITS OCCURRED)
                ' ---------------------------------------------------------------------
                ' Rewrite the text range back into the document layer only if changes were made
                If cleanText <> paraText Then
                    
                    ' Cache active style reference name before rewriting the range string
                    Dim currentStyle As Variant
                    Set currentStyle = txtRange.Style
                    
                    ' Write the perfectly trimmed text string back to the canvas
                    txtRange.text = cleanText & vbCr
                    
                    ' Reassert style boundaries and strip manual font modifications
                    txtRange.Style = currentStyle
                    txtRange.Font.Reset
                    
                End If
                
            End If
        End If
    Next para

CleanUp:
    ' Restore standard application window rendering metrics
    Application.ScreenUpdating = True
    MsgBox "Heading margins successfully trimmed and cleaned!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen updates before throwing structural execution errors
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Execution Fault"
End Sub

Sub Trim_Multiple_Spaces_In_Selection()
'=============================================================================
' Name: Trim_Multiple_Spaces_In_Selection
' Purpose: Finds and replaces all double (and more than double) spaces within
'          the user's highlighted selection (paragraphs and tables) and
'          collapses them down into a single standard space.
' SAFETY: Confines operations strictly to the highlighted Selection Range,
'         ensuring the unselected rest of the document remains untouched.
'=============================================================================
    Dim selectRange As Range
    
    ' Guardrail: Check if there is a valid, active selection before proceeding
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the paragraph(s) or table area you want to clean first.", _
               vbExclamation, "No Selection Detected"
        Exit Sub
    End If
    
    ' Assign the precise boundary limits of your current selection
    Set selectRange = Selection.Range
    
    ' Speed Optimization: Silence screen rendering changes to execute instantly
    Application.ScreenUpdating = False
    
    ' Configure Word's high-speed Find and Replace engine
    With selectRange.Find
        .ClearFormatting
        .replacement.ClearFormatting
        
        .text = "  "                ' Target two spaces
        .replacement.text = " "     ' Replace with one space
        
        .Forward = True
        .Wrap = wdFindStop          ' CRITICAL: Halts the engine at selection borders
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Execute an initial loop to aggressively clear stacked spaces down to single spaces
        Do While .Execute(Replace:=wdReplaceAll)
            ' Loop cycles continuously if 3+ spaces fold down into 2 spaces,
            ' guaranteeing all multi-space gaps are entirely compressed.
        Loop
    End With
    
    ' Restore standard system display updates
    Application.ScreenUpdating = True
    
    MsgBox "Successfully collapsed all multiple spaces down to single spaces within your selection!", _
           vbInformation, "Process Complete"
End Sub

Sub Correct_Selected_Paragraph_Indents()
    ' ============================================================================
    ' MODULE:       Correct_Selected_Paragraph_Indents
    ' DESCRIPTION:  Resets left, right, and first-line/hanging indents to 0
    '               for selected paragraphs, explicitly skipping active lists.
    ' AUTHOR:       VBA Automation Suite
    ' ============================================================================
    
    Dim para As Paragraph
    
    ' Establish a global error handler to gracefully manage unexpected layout blocks
    On Error GoTo CleanUp
    
    ' Performance Optimization: Freeze visual layout rendering to maximize speed
    Application.ScreenUpdating = False
    
    ' Loop exclusively through the paragraphs actively highlighted on screen
    For Each para In Selection.Paragraphs
        
        ' Core Guardrail: Evaluate the paragraph's list layout metadata
        ' wdListNoNumbering represents plain, standard body text layers
        If para.Range.ListFormat.ListType = wdListNoNumbering Then
            
            ' Reset indentation metrics directly on the paragraph's top-level interface
            para.LeftIndent = 0
            para.RightIndent = 0
            para.FirstLineIndent = 0
            
        End If
        
    Next para

CleanUp:
    ' Essential Safety Pass: Forcefully restore screen updates under all conditions
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

Sub Fix_Table_Row_Cell_Padding()
    Dim tbl As Table
    
    ' Loop through every table asset globally in the document
    For Each tbl In ActiveDocument.Tables
        With tbl
            ' Force paragraph spacing limits inside the grid to stay tight
            .Range.ParagraphFormat.SpaceBefore = 0
            .Range.ParagraphFormat.SpaceAfter = 0
            .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            
            ' Strip out the invisible cell padding limits
            .TopPadding = InchesToPoints(0)
            .BottomPadding = InchesToPoints(0)
            .LeftPadding = InchesToPoints(0.05)
            .RightPadding = InchesToPoints(0.05)
            
            ' Allow rows to naturally auto-fit the font height tightly
            .Rows.HeightRule = wdRowHeightAuto
            .Rows.Height = 0
        End With
    Next tbl
    
    MsgBox "Table cell padding cleared successfully!", vbInformation, "Layout Fixed"
End Sub