Attribute VB_Name = "Module12"
Sub Apply_Report_Layout_To_All_Sections()
'
' Apply_Report_Layout_To_All_Sections Macro
'
    Dim sec As Section
    Dim targetOrient As Long
    Dim targetWidth As Double
    Dim targetHeight As Double
    
    ' Speed optimization: Turn off screen updating while running
    Application.ScreenUpdating = False
    
    ' Enable error handling
    On Error GoTo CleanUp
    
    ' Loop through every section in the document
    For Each sec In ActiveDocument.Sections
        
        ' Determine dimensions based on current orientation
        Select Case sec.PageSetup.Orientation
            Case wdOrientPortrait
                targetOrient = wdOrientPortrait
                targetWidth = 8.27
                targetHeight = 11.69
                
            Case wdOrientLandscape
                targetOrient = wdOrientLandscape
                targetWidth = 11.69
                targetHeight = 8.27
                
            Case Else
                ' Skip if it encounters an unexpected orientation type and move to next section
                GoTo NextSection
        End Select
        
        ' Apply all settings efficiently in a single block
        With sec.PageSetup
            .Orientation = targetOrient
            .PageWidth = InchesToPoints(targetWidth)
            .PageHeight = InchesToPoints(targetHeight)
            
            ' Shared Layout Settings
            .LineNumbering.Active = False
            .TopMargin = InchesToPoints(0.25)
            .BottomMargin = InchesToPoints(0.25)
            .LeftMargin = InchesToPoints(0.75)
            .RightMargin = InchesToPoints(0.75)
            .Gutter = InchesToPoints(0)
            .HeaderDistance = InchesToPoints(0.25)
            .FooterDistance = InchesToPoints(0.25)
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
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

    ' Standard completion message block (Only hits if NO errors occurred)
    Application.ScreenUpdating = True
    MsgBox "Layout applied to all sections.", vbInformation, "Success"
    Exit Sub ' Prevents the code from accidentally falling into the error handler

CleanUp:
    ' This section ONLY runs if a VBA runtime error is thrown
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
End Sub
