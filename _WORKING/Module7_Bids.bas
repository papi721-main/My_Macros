'=============================================================================
' MODULE: Bid Document Formatting and Layout Tools
'=============================================================================

Sub Bids_1_Create_TECH_2_Headings()
'=============================================================================
' Name: Bids_1_Create_TECH_2_Headings
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
        Set sty = doc.styles(CStr(styleNames(i)))
        If sty Is Nothing Then
            Set sty = doc.styles.Add(Name:=CStr(styleNames(i)), Type:=wdStyleTypeParagraph)
        End If
        On Error GoTo CleanUp
        
        ' Apply core style relations
        With sty
            .BaseStyle = doc.styles(wdStyleNormal)
            .NextParagraphStyle = doc.styles(wdStyleNormal)
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

Sub Bids_2_Remap_Headings_In_Selection_To_TECH_2_Headings()
'=============================================================================
' Name: Bids_2_Remap_Headings_In_Selection_To_TECH_2_Headings
' Purpose: Sweeps through highlighted/selected paragraphs and remaps headings
'          to custom "T2-Heading" styles based on their structural Outline Level.
'
' Target Mapping:
'   - Outline Level 2 (Heading 2) -> T2-Heading 3
'   - Outline Level 3 (Heading 3) -> T2-Heading 4
'   - Outline Level 4 (Heading 4) -> T2-Heading 5
'   - Outline Level 5 (Heading 5) -> T2-Heading 6
'=============================================================================
    Dim doc As Document
    Dim para As Paragraph
    Dim targetStyleName As String
    Dim successCount As Long
    
    ' Pre-flight check: Ensure user has a selection active
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select/highlight the portion of text containing the headings you want to remap.", _
               vbExclamation, "No Text Selected"
        Exit Sub
    End If
    
    Set doc = ActiveDocument
    successCount = 0
    
    ' Speed optimization
    Application.ScreenUpdating = False
    On Error GoTo CleanUp
    
    ' Loop strictly through paragraphs in the user's highlighted selection
    For Each para In Selection.Paragraphs
        
        ' Ignore content inside tables
        If Not para.Range.Information(wdWithInTable) Then
            
            targetStyleName = ""
            
            ' Determine target T2 style based on structural Outline Level
            Select Case para.OutlineLevel
                Case wdOutlineLevel2
                    targetStyleName = "T2-Heading 3"
                    
                Case wdOutlineLevel3
                    targetStyleName = "T2-Heading 4"
                    
                Case wdOutlineLevel4
                    targetStyleName = "T2-Heading 5"
                    
                Case wdOutlineLevel5
                    targetStyleName = "T2-Heading 6"
            End Select
            
            ' If a matching outline level was identified, re-assign the style
            If targetStyleName <> "" Then
                On Error Resume Next
                
                ' Apply the custom T2 style
                para.Style = doc.styles(targetStyleName)
                
                ' If style assignment succeeded, clear character-level direct formatting overrides
                If Err.Number = 0 Then
                    para.Range.Font.Reset
                    successCount = successCount + 1
                End If
                
                On Error GoTo CleanUp
            End If
            
        End If
    Next para
    
    ' Restore screen rendering and notify user
    Application.ScreenUpdating = True
    
    If successCount > 0 Then
        MsgBox successCount & " heading paragraph(s) in the selection were successfully remapped to T2-Heading styles!", _
               vbInformation, "Remap Complete"
    Else
        MsgBox "No Level 2\u20135 headings were found in the selected text.", _
               vbExclamation, "No Headings Remapped"
    End If
    
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while remapping styles: " & Err.Description, _
           vbCritical, "Execution Fault"
End Sub


Sub Bids_3_Adjust_Bid_Document_Styles()
'=============================================================================
' Name: Bids_3_Adjust_Bid_Document_Styles
' Purpose: Explicitly configures and standardizes core body styles (Normal,
'          Normal (Web), Body Text, etc.), Heading styles (1 through 4),
'          and the Caption style.
'          Establishes layout baselines, clears rogue tab stops, strips out
'          any legacy/manual paragraph borders, and ensures clean text geometries.
' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
' PERFORMANCE:  Modifies named stylesheet assets directly in memory, bypassing
'               the need to loop paragraph-by-paragraph or move the cursor.
'=============================================================================
    Dim doc As Document
    Dim normalStyleNames As Variant
    Dim headingNames As Variant
    Dim i As Long
    Dim stName As Variant
    
    Set doc = ActiveDocument
    
    ' Freeze visual application window rendering to prevent layout redraw lag
    Application.ScreenUpdating = False
    
    ' Establish global runtime error trapping to protect the active workspace environment
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' 1. NORMAL & BODY TEXT STYLES LOOP (Standardizes baseline body text styles)
    '-------------------------------------------------------------------------
    normalStyleNames = Array( _
        "Normal", _
        "Normal (Web)", _
        "Body Text", _
        "Body Text 2", _
        "Body Text 3", _
        "Normal Indent", _
        "Body Text Indent", _
        "Body Text Indent 2", _
        "Body Text Indent 3", _
        "Table Paragraph")

    For Each stName In normalStyleNames
        ' Temporary error bypass in case a specific variant style does not exist in the document
        On Error Resume Next
        With doc.styles(stName)
            .AutomaticallyUpdate = False
            With .Font
                ' Basic Font Properties
                .Name = "Calibri"
                .Size = 11
                .Bold = False
                .Italic = False
                .Color = wdColorAutomatic
                .Outline = False            ' Removes any unwanted borders around text characters
                .Shadow = False             ' Removes any legacy shadow tracking text effects
                .Emboss = False             ' Clears any manual embossing text effects
                .Engrave = False            ' Clears any manual engraving text effects
                
                ' Advanced Typography Rules
                .Spacing = 0                                ' Resets manual character spacing adjustments
                .Scaling = 100                              ' Normalizes font width scaling back to default
                .Kerning = 0                                ' Disables explicit font kerning limits
                .Ligatures = wdLigaturesNone                ' Prevents automatic ligature glyph combinations
                .NumberSpacing = wdNumberSpacingDefault     ' Standardizes numeric layout spacing
                .NumberForm = wdNumberFormDefault           ' Resets lining vs. old-style number overrides
                .StylisticSet = wdStylisticSetDefault       ' Disables advanced font stylistic glyph sets
                .ContextualAlternates = 0                    ' Shuts off contextual character alternates
            End With
            With .ParagraphFormat
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .FirstLineIndent = InchesToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .LeftIndent = InchesToPoints(0)
                .RightIndent = InchesToPoints(0)
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacingRule = wdLineSpace1pt5     ' Enforces consistent 1.5 line heights
                .Alignment = wdAlignParagraphJustify    ' Justified text layout for reporting blocks
                .WidowControl = True                    ' Prevents orphan sentences at page boundaries
                .TabStops.ClearAll                      ' Squeezes out rogue manual tab stop intervals
                
                ' ARCHITECTURAL STRATEGY: .Borders.Enable = False acts as a safe global clear pass.
                .Borders.Enable = False
            End With
        End With
        On Error GoTo ErrorHandler
    Next stName

    '-------------------------------------------------------------------------
    ' 2. HEADING 1 (Primary Document Sections - All Caps & Left-Aligned)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 1")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            '.Size = 18      ' Set to 18 for Reports
            .Size = 24       ' Set to 24 for Bids
            .Bold = True
            .Italic = False
            .AllCaps = True

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
            .LineSpacingRule = wdLineSpaceSingle   ' Single line spacing for a tight heading layout

            '.Alignment = wdAlignParagraphLeft      ' Set to Left-aligned for Reports
            .Alignment = wdAlignParagraphCenter    ' Set to Center-aligned for Bids

            .OutlineLevel = wdOutlineLevel1        ' Mandatory tier assignment for core TOC extraction

            '.PageBreakBefore = True                ' Set to True for Reports to force headings onto a new page
            .PageBreakBefore = False              ' Set to False for Bids to allow headings to flow with content

            .KeepWithNext = True                   ' Locks heading onto the same page as the following body text
            .KeepTogether = True                   ' Prevents heading text lines from splitting across pages
            .TabStops.ClearAll
            .Borders.Enable = False                ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 3. HEADING 2 (Sub-sections - Left-Aligned & Bound to Following Text)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 2")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 16
            .Bold = True
            .Italic = False
            
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
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel2        ' TOC Tier 2 registration anchor
            .TabStops.ClearAll
            .Borders.Enable = False                ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 4. HEADING 3 (Sub-sub-sections)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 3")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 14
            .Bold = True
            .Italic = False
            
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
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel3        ' TOC Tier 3 registration anchor
            .TabStops.ClearAll
            .Borders.Enable = False                ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 5. HEADING 4 (Deep Hierarchy Details)
    '-------------------------------------------------------------------------
    With doc.styles("Heading 4")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 12
            .Bold = True
            .Italic = False

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
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel4        ' TOC Tier 4 registration anchor
            .TabStops.ClearAll
            .Borders.Enable = False                ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 6. CAPTION STYLE (The style for captioning tables, figures, and other media)
    '-------------------------------------------------------------------------
    With doc.styles("Caption")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = True
            .Italic = True
            .Color = wdColorAutomatic
            .AllCaps = False
            .SmallCaps = False

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
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = InchesToPoints(0)
            .LeftIndent = InchesToPoints(0)
            .RightIndent = InchesToPoints(0)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)     ' Dynamically maps single-spaced multiple multipliers
            .Alignment = wdAlignParagraphJustify
            .KeepWithNext = True                    ' Keeps caption tethered onto the same page as its media asset
            .KeepTogether = True                    ' Prevents caption lines from wrapping awkwardly across page breaks
            .WidowControl = True                    ' Prevents orphan sentences at page boundaries
            .OutlineLevel = wdOutlineLevelBodyText  ' Keeps captions from accidentally bleeding into your TOC index
            .TabStops.ClearAll
            .Borders.Enable = False                 ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 7. CENTRALIZED HEADING COLOR PASS (Applies custom hex #182C52 via Loop)
    '-------------------------------------------------------------------------
    headingNames = Array("Heading 1", "Heading 2", "Heading 3", "Heading 4")
    
    For i = LBound(headingNames) To UBound(headingNames)
        With doc.styles(headingNames(i)).Font
            ' ACTIVE CONFIGURATION: Apply custom hex color #182C52 natively
            '.Color = RGB(24, 44, 82)
            
            ' ROLLBACK TOGGLE: Uncomment the line below to easily reset everything back to Automatic
            .Color = wdColorAutomatic
        End With
    Next i

CleanUp:
    ' Re-enable visual environment screen updates
    Application.ScreenUpdating = True
    MsgBox "Styles successfully updated with paragraph borders safely cleared!", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    ' Structural Fallback: Ensure system state unfreezes cleanly if a style lookup fails
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Style Preferences Error"
    Resume CleanUp
End Sub

Sub Bids_4_Adjust_Bid_Documents_Layout()
'=============================================================================
' Name: Bids_4_Adjust_Bid_Documents_Layout
' Purpose: Loops through document sections and configures page setups to A4.
'          - Restricts modifications strictly to A4 or Letter size pages.
'          - Section 1 (Cover): Preserves single header/footer mode.
'          - Section 2+ (TOC & Body): Forces OddAndEvenPagesHeaderFooter = True.
'          - If Odd & Even is ALREADY active: Adjusts margins, header/footer 
'            distances, and page dimensions only.
'          - Automatically bypasses sections locked by framed paragraphs.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim targetOrient As Long
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim i As Long
    
    ' Page dimension tracking variables (in points)
    Dim pWidth As Double
    Dim pHeight As Double
    Dim isA4OrLetter As Boolean
    Dim hasOddEvenAlready As Boolean
    
    ' Tracking variables for the final report
    Dim skippedSections As String
    Dim successCount As Long
    
    Set doc = ActiveDocument
    skippedSections = ""
    successCount = 0
    
    ' Speed optimization: Prevent screen flickering during deep object changes
    Application.ScreenUpdating = False
    
    ' Enable global error handling trap for unexpected core execution errors
    On Error GoTo CleanUp
    
    ' Loop through every isolated section block using a counter to remain stable
    For i = 1 To doc.Sections.Count
        Set sec = doc.Sections(i)
        
        ' ---------------------------------------------------------------------
        ' A4 / LETTER SIZE VALIDATION GUARDRAIL
        ' ---------------------------------------------------------------------
        pWidth = sec.PageSetup.PageWidth
        pHeight = sec.PageSetup.PageHeight
        isA4OrLetter = False
        
        ' Check native Word Enum constants first
        If sec.PageSetup.PaperSize = wdPaperA4 Or sec.PageSetup.PaperSize = wdPaperLetter Then
            isA4OrLetter = True
        Else
            ' Check explicit point bounds to catch custom-tagged A4 or Letter sizes
            ' Portrait check (Width 8.0" - 8.7", Height 10.5" - 12.0")
            If (pWidth >= InchesToPoints(8) And pWidth <= InchesToPoints(8.7)) And _
               (pHeight >= InchesToPoints(10.5) And pHeight <= InchesToPoints(12)) Then
                isA4OrLetter = True
            ' Landscape check (Width 10.5" - 12.0", Height 8.0" - 8.7")
            ElseIf (pWidth >= InchesToPoints(10.5) And pWidth <= InchesToPoints(12)) And _
                   (pHeight >= InchesToPoints(8) And pHeight <= InchesToPoints(8.7)) Then
                isA4OrLetter = True
            End If
        End If
        
        ' If the page is NOT A4 or Letter (e.g., A3, Legal), skip it
        If Not isA4OrLetter Then
            If skippedSections = "" Then
                skippedSections = CStr(i) & " (Not A4/Letter)"
            Else
                skippedSections = skippedSections & ", " & i & " (Not A4/Letter)"
            End If
            GoTo NextSection
        End If
        
        ' Determine targeted A4 dimensions based on existing page orientation
        Select Case sec.PageSetup.Orientation
            Case wdOrientPortrait
                targetOrient = wdOrientPortrait
                targetWidth = 8.27   ' A4 Width in inches
                targetHeight = 11.69 ' A4 Height in inches
                
            Case wdOrientLandscape
                targetOrient = wdOrientLandscape
                targetWidth = 11.69  ' A4 Landscape Width in inches
                targetHeight = 8.27  ' A4 Landscape Height in inches
                
            Case Else
                If skippedSections = "" Then
                    skippedSections = CStr(i)
                Else
                    skippedSections = skippedSections & ", " & i
                End If
                GoTo NextSection
        End Select
        
        ' ---------------------------------------------------------------------
        ' LOCAL INLINE SAFETY GUARDRAIL FOR FRAMED PARAGRAPHS
        ' ---------------------------------------------------------------------
        On Error Resume Next
        
        ' Check if Odd and Even setting is already enabled on this section
        hasOddEvenAlready = sec.PageSetup.OddAndEvenPagesHeaderFooter
        
        With sec.PageSetup
            ' --- ALWAYS APPLY: Core Dimensions & Margins ---
            .Orientation = targetOrient
            .PageWidth = InchesToPoints(targetWidth)
            .PageHeight = InchesToPoints(targetHeight)
            
            .TopMargin = InchesToPoints(0.25)
            .BottomMargin = InchesToPoints(0.25)
            .LeftMargin = InchesToPoints(0.75)
            .RightMargin = InchesToPoints(0.75)
            .Gutter = InchesToPoints(0)
            
            .HeaderDistance = InchesToPoints(0.25)
            .FooterDistance = InchesToPoints(0)
            
            ' --- CONDITIONAL SETUP BASED ON EXISTING ODD/EVEN STATE & SECTION INDEX ---
            If hasOddEvenAlready Then
                ' Section ALREADY has Odd & Even enabled: Only update metrics (above)
                ' and leave trays, breaks, and other flags untouched.
            Else
                ' Section does NOT have Odd & Even enabled yet: Apply full setup block
                .LineNumbering.Active = False
                .FirstPageTray = wdPrinterDefaultBin
                .OtherPagesTray = wdPrinterDefaultBin
                .SectionStart = wdSectionNewPage
                .DifferentFirstPageHeaderFooter = False
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .BookFoldPrinting = False
                .BookFoldRevPrinting = False
                .BookFoldPrintingSheets = 1
                .GutterPos = wdGutterPosLeft
                
                ' Enable Odd and Even headers/footers for Section 2 and onwards
                If i > 1 Then
                    .OddAndEvenPagesHeaderFooter = True
                Else
                    .OddAndEvenPagesHeaderFooter = False
                End If
            End If
        End With
        
        ' Check if the PageSetup properties threw a frame-lock error
        If Err.Number <> 0 Then
            Err.Clear
            If skippedSections = "" Then
                skippedSections = CStr(i) & " (Frame Locked)"
            Else
                skippedSections = skippedSections & ", " & i & " (Frame Locked)"
            End If
        Else
            successCount = successCount + 1
        End If
        
        ' Reset global error trapping rules for the next iteration step
        On Error GoTo CleanUp

NextSection:
    Next i

    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    
    ' Format and present the final completion report message box
    Dim reportMessage As String
    reportMessage = "Layout processing complete." & vbCrLf & vbCrLf & _
                    "Sections adjusted successfully: " & successCount
                    
    If skippedSections <> "" Then
        reportMessage = reportMessage & vbCrLf & vbCrLf & _
                        "ATTENTION: The following sections were SKIPPED " & _
                        "(Not A4/Letter size or locked by framed paragraphs):" & vbCrLf & _
                        "Section(s): " & skippedSections & vbCrLf & vbCrLf & _
                        "Please inspect and adjust these sections manually if needed."
        MsgBox reportMessage, vbWarning, "Process Complete with Bypasses"
    Else
        MsgBox reportMessage, vbInformation, "Success"
    End If
    
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Execution Fault"
End Sub
