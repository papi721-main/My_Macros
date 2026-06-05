'=============================================================================
' MODULE: Document Layout Management Tools
'=============================================================================

Sub Layout_1_Adjust_Layout_On_All_Sections()
'=============================================================================
' Name: Layout_1_Adjust_Layout_On_All_Sections
' Purpose: Loops through all document sections and configures page setups.
'          Detects portrait vs. landscape orientations and assigns clean,
'          standardized A4 dimensions and precise 0.25"/0.75" margin bounds.
'=============================================================================
    Dim sec As Section
    Dim targetOrient As Long
    Dim targetWidth As Double
    Dim targetHeight As Double
    
    ' Speed optimization: Prevent screen flickering during deep object changes
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo CleanUp
    
    ' Loop through every isolated section block in the active document
    For Each sec In ActiveDocument.Sections
        
        ' Determine targeted A4 dimensions based on existing page orientation
        Select Case sec.PageSetup.Orientation
            Case wdOrientPortrait
                targetOrient = wdOrientPortrait
                targetWidth = 8.27   ' A4 Width in inches
                targetHeight = 11.69 ' A4 Height in inches
                
            Case wdOrientLandscape
                targetOrient = wdOrientLandscape
                targetWidth = 11.69  ' A4 Landscape Width in inches
                targetHeight = 8.27   ' A4 Landscape Height in inches
                
            Case Else
                ' Bypass unexpected or unreadable core layouts safely
                GoTo NextSection
        End Select
        
        ' Mass-apply explicit geometry specifications to the targeted section
        With sec.PageSetup
            ' Orientation and Physical Sizing
            .Orientation = targetOrient
            .PageWidth = InchesToPoints(targetWidth)
            .PageHeight = InchesToPoints(targetHeight)
            
            ' Margin Layout Boundaries
            .LineNumbering.Active = False
            .TopMargin = InchesToPoints(0.25)
            .BottomMargin = InchesToPoints(0.25)
            .LeftMargin = InchesToPoints(0.75)
            .RightMargin = InchesToPoints(0.75)
            .Gutter = InchesToPoints(0)
            
            ' Header and Footer Layout Placement
            .HeaderDistance = InchesToPoints(0.25)
            .FooterDistance = InchesToPoints(0.25)
            
            ' Printer Tray and Break Rules
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            
            ' Header/Footer Visibility Flags
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            
            ' Advanced Page Content Processing Properties
            .VerticalAlignment = wdAlignVerticalTop
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .BookFoldPrinting = False
            .BookFoldRevPrinting = False
            .BookFoldPrintingSheets = 1
            .GutterPos = wdGutterPosLeft
        End With

NextSection:
    Next sec

    ' Standard Completion Execution Path
    Application.ScreenUpdating = True
    MsgBox "Layout applied to all sections.", vbInformation, "Success"
    Exit Sub ' Secure break preventing fall-through into error sequence

CleanUp:
    ' Emergency Handling Sequence (Triggers only upon structural execution faults)
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
End Sub


Sub Layout_2_Change_Continuous_Break_To_NextPage_Break()
'=============================================================================
' Name: Layout_2_Change_Continuous_Break_To_NextPage_Break
' Purpose: Scans document sections backwards to find continuous section breaks
'          and up-converts them into standard "Next Page" section breaks.
'=============================================================================
    Dim nSectionNum As Integer
    Dim objDoc As Document
   
    ' Speed optimization: Freeze screen rendering during iteration pass
    Application.ScreenUpdating = False
   
    Set objDoc = ActiveDocument
    nSectionNum = objDoc.Sections.Count

    ' Reverse-loop through sections to prevent indexing shift errors upon modification
    For nSectionNum = nSectionNum To 1 Step -1
        With objDoc.Sections(nSectionNum).Range.PageSetup
            ' If the section is configured to start mid-page, force it to a new page
            If .SectionStart = wdSectionContinuous Then
                .SectionStart = wdSectionNewPage
            End If
        End With
    Next nSectionNum
   
    ' Restore standard application display
    Application.ScreenUpdating = True
End Sub


Sub Layout_3_Delete_Section_Breaks()
'=============================================================================
' Name: Layout_3_Delete_Section_Breaks
' Purpose: Instantly purges all Section Breaks (^b) from the document using
'          an optimized, single-pass Find and Replace operation.
'=============================================================================
    Dim objDoc As Document
    Dim rng As Range

    ' Speed optimization: Turn off screen rendering during deletion pass
    Application.ScreenUpdating = False
    
    Set objDoc = ActiveDocument
    Set rng = objDoc.Content

    ' Execute text parsing directly across the core body range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "^b"                ' Word's internal token pattern for Section Breaks
        .Replacement.Text = ""      ' Replace text with empty string (force deletion)
        .Forward = True
        
        ' Enforces strict single-pass execution over the scope layout range
        .Wrap = wdFindStop
        
        ' Process structural updates globally
        .Execute Replace:=wdReplaceAll
    End With

    ' Restore screen tracking
    Application.ScreenUpdating = True

    MsgBox "All section breaks have been successfully deleted.", vbInformation, "Purge Complete"
End Sub


Sub Layout_4_Convert_Sections_To_Page_Breaks()
'=============================================================================
' Name: Layout_4_Convert_Sections_To_Page_Breaks
' Purpose: Instantly converts all Section Breaks (^b) into standard Page Breaks (^m)
'          across the entire document using a single-pass Find and Replace.
'=============================================================================
    Dim objDoc As Document
    Dim rng As Range

    ' Speed optimization: Turn off screen rendering during conversion pass
    Application.ScreenUpdating = False
    
    Set objDoc = ActiveDocument
    Set rng = objDoc.Content

    ' Execute text parsing directly across the core body range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "^b"                ' Word's internal token pattern for Section Breaks
        .Replacement.Text = "^m"    ' Word's internal token pattern for Page Breaks
        .Forward = True
        
        ' Enforces strict single-pass execution over the scope layout range
        .Wrap = wdFindStop
        
        ' Process structural updates globally
        .Execute Replace:=wdReplaceAll
    End With

    ' Restore screen tracking
    Application.ScreenUpdating = True

    MsgBox "All section breaks have been converted to page breaks.", vbInformation, "Conversion Complete"
End Sub