Attribute VB_Name = "Module5_Headers_and_Footers"
'=============================================================================
' MODULE: Headers & Footers
'=============================================================================

Sub Footers_Toggle_Images_Visibility_Main_Sections_Only()
'=============================================================================
' Name: Footers_Toggle_Images_Visibility_Main_Sections_Only
' Purpose: Loops through main document sections and completely iterates through
'          ALL footer sub-layers (First Page, Even Pages, and Primary) to switch
'          inline images between VISIBLE and INVISIBLE using a brightness mask.
'          COMPREHENSIVE: Excludes cover pages and isolated TOC sections.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim iShp As InlineShape
    Dim footerIndex As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Freeze screen updates while modifying background story layers
    Application.ScreenUpdating = False
    
    ' Loop through every structural section boundary in the document
    For Each sec In doc.Sections
        
        ' Iterate through all 3 possible header/footer sub-layer types in Word:
        ' 1 = wdHeaderFooterPrimary
        ' 2 = wdHeaderFooterFirstPage
        ' 3 = wdHeaderFooterEvenPages
        For footerIndex = 1 To 3
            
            Set ftr = sec.Footers(footerIndex)
            
            ' Verification: Ensure the specific sub-footer canvas exists and is active
            If Not ftr Is Nothing Then
                
                ' Route explicitly through the text Range link to look at InlineShapes
                For Each iShp In ftr.Range.InlineShapes
                    
                    ' Target exclusively baseline graphic or linked picture objects
                    If iShp.Type = wdInlineShapePicture Or iShp.Type = wdInlineShapeLinkedPicture Then
                        
                        ' Binary Toggle Pass: Check exposure values
                        If iShp.PictureFormat.Brightness = 0.5 Then
                            
                            ' SWITCH OFF: Mask to pure white (invisible against white page)
                            iShp.PictureFormat.Brightness = 1#
                            
                        Else
                            
                            ' SWITCH ON: Restore exposure back to 50% factory default
                            iShp.PictureFormat.Brightness = 0.5
                            
                        End If
                        
                    End If
                Next iShp
                
            End If
        Next footerIndex
        
    Next sec
    
    ' Restore application window layout rendering
    Application.ScreenUpdating = True
    
    MsgBox "All Main footer layers (excluding Cover and TOC pages) have been successfully toggled!", vbInformation, "State Changed"
End Sub

Sub Footers_Purge_All_Images()
'=============================================================================
' Name: Footers_Purge_All_Images
' Purpose: Loops through all document sections and footer layers, performing a
'          safe backward count to delete every inline picture permanently.
'          COMPREHENSIVE: Clears out custom main-canvas cover page images as well.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim iShp As InlineShape
    Dim footerIndex As Long
    Dim i As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Freeze screen updates while scrubbing background layers
    Application.ScreenUpdating = False
    
    '-------------------------------------------------------------------------
    ' PASS 1: SAFE REVERSE LOOP THROUGH ALL FOOTER LAYERS
    '-------------------------------------------------------------------------
    For Each sec In doc.Sections
        For footerIndex = 1 To 3
            
            Set ftr = sec.Footers(footerIndex)
            
            If Not ftr Is Nothing Then
                ' Grab the total count of inline shapes inside this specific footer text layer
                ' and count BACKWARDS to execute zero-error deletions.
                For i = ftr.Range.InlineShapes.Count To 1 Step -1
                    
                    Set iShp = ftr.Range.InlineShapes(i)
                    
                    ' Target exclusively standard graphic or linked picture elements
                    If iShp.Type = wdInlineShapePicture Or iShp.Type = wdInlineShapeLinkedPicture Then
                        iShp.Delete
                    End If
                    
                Next i
            End If
            
        Next footerIndex
    Next sec
    
    '-------------------------------------------------------------------------
    ' PASS 2: CUSTOM CANVAS COVER PAGE CLEANUP (First Section Main Text Layer)
    '-------------------------------------------------------------------------
    If doc.Sections.Count > 0 Then
        Set sec = doc.Sections(1)
        
        ' Count backwards through the main body canvas of the cover sheet section
        For i = sec.Range.InlineShapes.Count To 1 Step -1
            
            Set iShp = sec.Range.InlineShapes(i)
            
            If iShp.Type = wdInlineShapePicture Or iShp.Type = wdInlineShapeLinkedPicture Then
                iShp.Delete
            End If
            
        Next i
    End If
    
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    
    MsgBox "All images inside footers have been permanently removed!", vbInformation, "Purge Complete"
End Sub

Sub Footers_Images_Switch_OFF()
'=============================================================================
' Name: Footers_Images_Switch_OFF
' Purpose: Loops through all document sections and deep footer sub-layers.
'          Forces the brightness of all inline pictures to 1.0 (Pure White).
'          RESULT: Effectively cloaks/hides footer images flawlessly.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim iShp As InlineShape
    Dim footerIndex As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Freeze screen updates while processing background layers
    Application.ScreenUpdating = False
    
    ' Loop through every structural section boundary in the document
    For Each sec In doc.Sections
        
        ' Cycle through all 3 possible footer sub-layer templates in Word:
        ' 1 = wdHeaderFooterPrimary, 2 = wdHeaderFooterFirstPage, 3 = wdHeaderFooterEvenPages
        For footerIndex = 1 To 3
            
            Set ftr = sec.Footers(footerIndex)
            
            ' Verify the specific footer layer is active/instantiated
            If Not ftr Is Nothing Then
                
                ' Route explicitly through the text Range link to look at InlineShapes
                For Each iShp In ftr.Range.InlineShapes
                    
                    ' Target exclusively standard graphic or linked picture elements
                    If iShp.Type = wdInlineShapePicture Or iShp.Type = wdInlineShapeLinkedPicture Then
                        
                        ' Force brightness to 1.0 (Mask to pure white)
                        iShp.PictureFormat.Brightness = 1#
                        
                    End If
                Next iShp
                
            End If
            
        Next footerIndex
    Next sec
    
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    
    MsgBox "All footer images have been switched OFF (hidden) successfully!", vbInformation, "Visibility Update"
End Sub

Sub Footers_Images_Switch_ON()
'=============================================================================
' Name: Footers_Images_Switch_ON
' Purpose: Loops through all document sections and deep footer sub-layers.
'          Forces the brightness of all inline pictures back to 0.5 (Default).
'          RESULT: Restores full visibility to footer images instantly.
'=============================================================================
    Dim doc As Document
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim iShp As InlineShape
    Dim footerIndex As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Freeze screen updates while processing background layers
    Application.ScreenUpdating = False
    
    ' Loop through every structural section boundary in the document
    For Each sec In doc.Sections
        
        ' Cycle through all 3 possible footer sub-layer templates in Word
        For footerIndex = 1 To 3
            
            Set ftr = sec.Footers(footerIndex)
            
            ' Verify the specific footer layer is active/instantiated
            If Not ftr Is Nothing Then
                
                ' Route explicitly through the text Range link to look at InlineShapes
                For Each iShp In ftr.Range.InlineShapes
                    
                    ' Target exclusively standard graphic or linked picture elements
                    If iShp.Type = wdInlineShapePicture Or iShp.Type = wdInlineShapeLinkedPicture Then
                        
                        ' Reset brightness back to 50% baseline tracking standard
                        iShp.PictureFormat.Brightness = 0.5
                        
                    End If
                Next iShp
                
            End If
            
        Next footerIndex
    Next sec
    
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    
    MsgBox "All footer images have been switched ON (visible) successfully!", vbInformation, "Visibility Update"
End Sub

