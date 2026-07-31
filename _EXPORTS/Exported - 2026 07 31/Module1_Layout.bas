Attribute VB_Name = "Module1_Layout"
'=============================================================================
' MODULE: Document Layout Management Tools
'=============================================================================

Sub Layout_1_Adjust_Layout_On_All_Sections()
'=============================================================================
' Name: Layout_1_Adjust_Layout_On_All_Sections
' Purpose: Loops through all document sections and configures page setups.
'          Detects portrait vs. landscape orientations and assigns clean,
'          standardized A4 dimensions and precise 0.25"/0.75" margin bounds.
'          SAFE GUARDRAILS:
'            1. Restricts modifications STRICTLY to A4 or Letter size pages;
'               skips all other paper sizes (A3, Legal, Executive, etc.).
'            2. Automatically bypasses sections locked by framed paragraphs.
'          REPORTER: Tracks and displays skipped sections at completion.
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
        ' Measure section's current dimensions
        pWidth = sec.PageSetup.PageWidth
        pHeight = sec.PageSetup.PageHeight
        isA4OrLetter = False
        
        ' Check native Word Enum constants first
        If sec.PageSetup.PaperSize = wdPaperA4 Or sec.PageSetup.PaperSize = wdPaperLetter Then
            isA4OrLetter = True
        Else
            ' Check explicit point bounds to catch custom-tagged A4 or Letter sizes
            ' Letter: ~8.5" x 11" (612 pt x 792 pt) | A4: ~8.27" x 11.69" (595 pt x 842 pt)
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
        
        ' If the page is NOT A4 or Letter (e.g. A3, Legal, etc.), skip it
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
                targetHeight = 8.27   ' A4 Landscape Height in inches
                
            Case Else
                ' Log as skipped if orientation is unreadable
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
        ' If a section contains a framed paragraph, accessing its PageSetup properties
        ' throws an error. We suppress the error, attempt the block, and verify success.
        On Error Resume Next
        
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
            .FooterDistance = InchesToPoints(0)
            
            ' Printer Tray and Break Rules
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            
            ' Header/Footer Visibility Flags
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            
            ' Advanced Page Content Processing Properties
            '.VerticalAlignment = wdAlignVerticalTop
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .BookFoldPrinting = False
            .BookFoldRevPrinting = False
            .BookFoldPrintingSheets = 1
            .GutterPos = wdGutterPosLeft
        End With
        
        ' Check if the PageSetup properties threw a frame-lock error
        If Err.Number <> 0 Then
            ' Clear the error and log this specific section number as skipped
            Err.Clear
            If skippedSections = "" Then
                skippedSections = CStr(i) & " (Frame Locked)"
            Else
                skippedSections = skippedSections & ", " & i & " (Frame Locked)"
            End If
        Else
            ' No error occurred; increment our success tracker
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
    
    Exit Sub ' Secure break preventing fall-through into error sequence

CleanUp:
    ' Emergency Handling Sequence (Triggers only upon structural execution faults)
    Application.ScreenUpdating = True
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Execution Fault"
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
        .replacement.ClearFormatting
        
        .text = "^b"                ' Word's internal token pattern for Section Breaks
        .replacement.text = ""      ' Replace text with empty string (force deletion)
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
        .replacement.ClearFormatting
        
        .text = "^b"                ' Word's internal token pattern for Section Breaks
        .replacement.text = "^m"    ' Word's internal token pattern for Page Breaks
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
