Attribute VB_Name = "Module1"
Sub Layout_1_Adjust_Layout_On_All_Sections()
'
' Adjust the layout on all sections, works on both
' Portrait and Landscape pages
' Paper = A4
' Top and Bottom Margin = 0.25
' Left and Right Margin = 0.75
' Headers and Footers = 0.25
'
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

Sub Layout_2_Change_Continuous_Break_To_NextPage_Break()
'
' Changes all Continuous Section Breaks to Next Page Section Breaks
'

  Dim nSectionNum As Integer
  Dim objDoc As Document
 
  Application.ScreenUpdating = False
 
  Set objDoc = ActiveDocument
  nSectionNum = objDoc.Sections.Count

  For nSectionNum = nSectionNum To 1 Step -1
    With objDoc.Sections(nSectionNum).Range.PageSetup
      If .SectionStart = wdSectionContinuous Then
        .SectionStart = wdSectionNewPage
      End If
    End With
  Next nSectionNum
 
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

    ' Execute search explicitly across the document body range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "^b"                ' Word's internal token for Section Breaks
        .Replacement.Text = ""      ' Overwrite with nothing (deletes them)
        .Forward = True
        
        ' FIXED: Changed to wdFindStop so Word executes across the range exactly
        ' once without looping back to the top of the file.
        .Wrap = wdFindStop
        
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

    ' Execute search explicitly across the document body range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = "^b"                ' Word's internal token for Section Breaks
        .Replacement.Text = "^m"    ' Word's internal token for Page Breaks
        .Forward = True
        .Wrap = wdFindStop          ' Process the range exactly once top-to-bottom
        
        .Execute Replace:=wdReplaceAll
    End With

    ' Restore screen tracking
    Application.ScreenUpdating = True

    MsgBox "All section breaks have been converted to page breaks.", vbInformation, "Conversion Complete"
End Sub
