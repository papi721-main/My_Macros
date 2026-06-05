'=============================================================================
' MODULE: Reference Tables & Automated Document Indexes
'=============================================================================

Sub References_1_Build_TOC()
'=============================================================================
' Name: References_1_Build_TOC
' Purpose: Automatically loops through and formats document TOC styles (1-5) 
'          with step-ladder indentation. Drops an automated Table of Contents
'          at the active cursor, configured to display only levels 1 to 3, 
'          while keeping hidden styling pre-cached for levels 4 and 5.
'=============================================================================
    Dim doc As Document
    Dim lvl As Integer
    Dim tocStyleName As String
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates while drawing reference matrices
    Application.ScreenUpdating = False
    
    '-------------------------------------------------------------------------
    ' 1. OPTIMIZE AND BUILD TOC STYLES (LEVELS 1 TO 5)
    '-------------------------------------------------------------------------
    ' Uniformly cycles through all five styles using a loop to bypass recording clutter
    For lvl = 1 To 5
        tocStyleName = "TOC " & lvl
        
        With doc.styles(tocStyleName)
            ' Automation Settings
            .AutomaticallyUpdate = True
            .BaseStyle = "Normal"
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            
            ' Configure Typography Tree Universally
            With .Font
                .Name = "Calibri"
                .Size = 11
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
                .Color = wdColorAutomatic
                .AllCaps = False
                .SmallCaps = False
                
                ' Level 1 gets strong visual emphasis (Bold + All Caps); 
                ' Deeper child levels (2-5) remain low-profile and standard.
                If lvl = 1 Then
                    .Bold = True
                    .AllCaps = True
                End If
            End With
            
            ' Configure Layout and Indent Spacing
            With .ParagraphFormat
                ' Establishes a step-ladder indentation framework (Level 1 = 0", Level 2 = 0.2", Level 3 = 0.4", etc.)
                .LeftIndent = InchesToPoints(0.2 * (lvl - 1))
                .RightIndent = InchesToPoints(0.5)
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpace1pt5
                .Alignment = wdAlignParagraphLeft
                .OutlineLevel = wdOutlineLevelBodyText
                
                ' Clear unmanaged/rogue manual tab stops to safeguard right-aligned numbers
                .TabStops.ClearAll
            End With
        End With
    Next lvl

    '-------------------------------------------------------------------------
    ' 2. GENERATE THE TABLE OF CONTENTS (SHOWS LEVELS 1-3 ONLY)
    '-------------------------------------------------------------------------
    With doc
        ' Generate a fresh Table of Contents instance over the active cursor.
        ' Note: LowerHeadingLevel is explicitly locked to 3 so levels 4 & 5 do not render.
        .TablesOfContents.Add Range:=Selection.Range, _
                              RightAlignPageNumbers:=True, _
                              UseHeadingStyles:=True, _
                              UpperHeadingLevel:=1, _
                              LowerHeadingLevel:=3, _
                              IncludePageNumbers:=True, _
                              AddedStyles:="", _
                              UseHyperlinks:=True, _
                              HidePageNumbersInWeb:=True, _
                              UseOutlineLevels:=True
        
        ' Enforce standard dot leaders on the freshly built index field
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        
        ' Setting format container to wdTOCNormal lets elements dynamically target 
        ' your custom style properties seamlessly without breaking the margin tab tracks.
        .TablesOfContents.Format = wdTOCNormal
    End With
    
    ' Re-enable system workspace rendering
    Application.ScreenUpdating = True
    
    ' Completion message deactivated per request
    ' MsgBox "Custom Table of Contents generated successfully!", vbInformation, "TOC Complete"
End Sub


Sub References_2_Adjust_Table_of_Figures()
'=============================================================================
' Name: References_2_Adjust_Table_of_Figures
' Purpose: Accesses and standardizes typography, tab alignments, and row 
'          line spacing for the native, built-in Table of Figures document style.
'=============================================================================
    Dim doc As Document
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Target the built-in "Table of Figures" system style structure directly
    With doc.styles("Table of Figures")
        ' Baseline Setup Properties
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .NoSpaceBetweenParagraphsOfSameStyle = False
        
        ' Enforce Clear, Uncluttered Font Styling
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .Color = wdColorAutomatic
            .AllCaps = False
        End With
        
        ' Set Paragraph Layout and Spacing Geometries
        With .ParagraphFormat
            .LeftIndent = InchesToPoints(0)
            .RightIndent = InchesToPoints(0)
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .OutlineLevel = wdOutlineLevelBodyText
            
            ' Clear unmanaged manual tab stops to ensure clean right-flushed page numbers
            .TabStops.ClearAll
        End With
    End With
    
    ' Re-enable system workspace rendering
    Application.ScreenUpdating = True
End Sub