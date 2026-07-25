'=============================================================================
' MODULE: Bid Document Formatting and Layout Tools
'=============================================================================

Sub Bids_Create_TECH_2_Headings()
'=============================================================================
' Name: Bids_Create_TECH_2_Headings
' Purpose: Automatically creates or updates the custom "T2-Heading" style 
'          hierarchy (T2-Heading 2 through T2-Heading 6) with standardized 
'          typography, spacing rules, and outline levels.
'          This is specifically tailored for Technical Proposal Documents.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    
    ' Array configurations for changing parameters
    Dim styleNames As Variant
    Dim fontSizes As Variant
    Dim isAllCaps As Variant
    Dim isSmallCaps As Variant
    Dim outlineLevels As Variant
    
    Set doc = ActiveDocument
    
    ' Define specifications for each target style
    styleNames = Array("T2-Heading 2", "T2-Heading 3", "T2-Heading 4", "T2-Heading 5", "T2-Heading 6")
    fontSizes = Array(20, 18, 16, 14, 12)
    isAllCaps = Array(True, False, False, False, False)
    isSmallCaps = Array(False, False, False, False, False)
    outlineLevels = Array(wdOutlineLevel2, wdOutlineLevel3, wdOutlineLevel4, wdOutlineLevel5, wdOutlineLevel6)
    
    ' Speed optimization
    Application.ScreenUpdating = False
    On Error GoTo CleanUp
    
    ' Loop through and configure all 5 custom styles
    For i = LBound(styleNames) To UBound(styleNames)
        
        ' ---------------------------------------------------------------------
        ' SAFE STYLE CREATION / RETRIEVAL
        ' ---------------------------------------------------------------------
        On Error Resume Next
        Set sty = doc.Styles(CStr(styleNames(i)))
        If sty Is Nothing Then
            Set sty = doc.Styles.Add(Name:=CStr(styleNames(i)), Type:=wdStyleTypeParagraph)
        End If
        On Error GoTo CleanUp
        
        ' Apply core style relations
        With sty
            .BaseStyle = doc.Styles(wdStyleNormal)
            .NextParagraphStyle = doc.Styles(wdStyleNormal)
            .AutomaticallyUpdate = False
            
            ' -----------------------------------------------------------------
            ' FONT & TYPOGRAPHY CONFIGURATION
            ' -----------------------------------------------------------------
            With .Font
                .Name = "Calibri"
                .Size = fontSizes(i)
                .Bold = True
                .Italic = False
                .AllCaps = isAllCaps(i)
                .SmallCaps = isSmallCaps(i)
                
                ' ACTIVE CONFIGURATION: Apply custom hex color #182C52 natively
                .Color = RGB(24, 44, 82)
                
                ' ROLLBACK TOGGLE: Uncomment the line below to easily reset everything back to Automatic
                '.Color = wdColorAutomatic
                
                ' Advanced Typography Rules
                .Spacing = 0
                .Scaling = 100
                .Kerning = 0
                .Ligatures = wdLigaturesNone
                .NumberSpacing = wdNumberSpacingDefault
                .NumberForm = wdNumberFormDefault
                .StylisticSet = wdStylisticSetDefault
                .ContextualAlternates = 0
            End With
            
            ' -----------------------------------------------------------------
            ' PARAGRAPH GEOMETRY & LAYOUT METRICS
            ' -----------------------------------------------------------------
            With .ParagraphFormat
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .FirstLineIndent = InchesToPoints(0)
                .LeftIndent = InchesToPoints(0)
                .RightIndent = InchesToPoints(0)
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 6
                .SpaceAfter = 12
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphLeft
                .OutlineLevel = outlineLevels(i)
                .PageBreakBefore = False
                .KeepWithNext = True
                .KeepTogether = True
                .TabStops.ClearAll
                .Borders.Enable = False
            End With
        End With
        
        Set sty = Nothing
    Next i

    ' Restore screen rendering and notify user
    Application.ScreenUpdating = True
    MsgBox "T2-Heading styles suite (Levels 2 to 6) successfully created/updated!", _
           vbInformation, "Styles Suite Complete"
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while creating styles: " & Err.Description, _
           vbCritical, "Execution Fault"
End Sub