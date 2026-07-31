Attribute VB_Name = "Module7_Bids"
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
        MsgBox "No Level 2–5 headings were found in the selected text.", _
               vbExclamation, "No Headings Remapped"
    End If
    
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "An error occurred while remapping styles: " & Err.Description, _
           vbCritical, "Execution Fault"
End Sub
