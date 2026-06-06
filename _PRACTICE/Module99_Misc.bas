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