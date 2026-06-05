Attribute VB_Name = "Module4"
Sub References_1_Build_TOC()
'=============================================================================
' Name: References_1_Build_TOC
' Purpose: Configures Word TOC styles (1-5), enforces proper indentation, clears
'          rogue tab stops, inserts clean dot leaders, and builds a custom TOC
'          that explicitly displays only levels 1 through 3.
'=============================================================================
    Dim doc As Document
    Dim lvl As Integer
    Dim tocStyleName As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    '-------------------------------------------------------------------------
    ' 1. OPTIMIZE AND BUILD TOC STYLES (LEVELS 1 TO 5)
    '-------------------------------------------------------------------------
    For lvl = 1 To 5
        tocStyleName = "TOC " & lvl
        
        With doc.styles(tocStyleName)
            .AutomaticallyUpdate = True
            .BaseStyle = "Normal"
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            
            ' Configure Font Properties universally
            With .Font
                .Name = "Calibri"
                .Size = 11
                .Italic = False
                .Underline = wdUnderlineNone
                .Color = wdColorAutomatic
                
                ' Level 1 gets Bold + AllCaps; Levels 2 through 5 remain standard
                If lvl = 1 Then
                    .Bold = True
                    .AllCaps = True
                Else
                    .Bold = False
                    .AllCaps = False
                End If
            End With
            
            ' Configure Layout and Indent Spacing
            With .ParagraphFormat
                ' Establish a clean, step-ladder indentation (Level 1=0", Level 2=0.2", Level 3=0.4", etc.)
                .LeftIndent = InchesToPoints(0.2 * (lvl - 1))
                .RightIndent = InchesToPoints(0.5)
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpace1pt5
                .Alignment = wdAlignParagraphLeft
                .OutlineLevel = wdOutlineLevelBodyText
                
                ' Clear unmanaged tab stops to ensure clean document geometry
                .TabStops.ClearAll
            End With
        End With
    Next lvl

    '-------------------------------------------------------------------------
    ' 2. GENERATE THE TABLE OF CONTENTS (SHOWS LEVELS 1-3 ONLY)
    '-------------------------------------------------------------------------
    With doc
        ' Generate a fresh Table of Contents instance over the active cursor.
        ' Note: LowerHeadingLevel is bound to 3 so Heading 4 & 5 don't render.
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
        
        ' Enforce dot leaders on the freshly built table object
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        
        ' Setting format to wdTOCNormal lets Word style elements follow your
        ' custom style properties seamlessly without breaking the margin tabs.
        .TablesOfContents.Format = wdTOCNormal
    End With
    
    Application.ScreenUpdating = True
    
    ' Completion message, deactivated via commenting for now
    ' MsgBox "Custom Table of Contents generated successfully!", vbInformation, "TOC Complete"
End Sub

Sub References_2_Adjust_Table_of_Figures()
'=============================================================================
' Name: References_2_Adjust_Table_of_Figures
' Purpose: Streamlines and standardizes typography and layout geometries for
'          the built-in Table of Figures document style.
'=============================================================================
    Dim doc As Document
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    ' Target the built-in "Table of Figures" style structure directly
    With doc.styles("Table of Figures")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .NoSpaceBetweenParagraphsOfSameStyle = False
        
        ' Set Core Font Properties
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
            
            ' Clear unmanaged tab stops to ensure clean right-flushed page numbers
            .TabStops.ClearAll
        End With
    End With
    
    Application.ScreenUpdating = True
End Sub
