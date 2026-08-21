'=============================================================================
' MODULE: Style Architecture & Content Cleanup Tools
'=============================================================================

Sub Style_1_Clean_Styles_Comprehensive()
'=============================================================================
' Name: Style_1_Clean_Styles_Comprehensive()
' Purpose: Removes custom styles from a general Word document while preserving
'          their visible formatting, list structures, tables, and equations.
'
'-----------------------------------------------------------------------------
' CLEANING STRATEGY
'-----------------------------------------------------------------------------
'
' Custom Style Type     Replacement             Preserved
' ---------------------------------------------------------------------------
' Heading-like Para.    Heading 1-9             Formatting + hierarchy
' Other Paragraph       Normal                  Font + paragraph formatting
' List Paragraph        Normal / Heading        List + formatting
' Character             Default Paragraph Font  Character formatting
' Table                 Table Normal            Table/cell formatting
' Equations             Unchanged               Completely skipped
'
' Built-in Word styles are never deleted.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim tbl As Table
    Dim cel As Cell
    Dim cellPara As Paragraph
    Dim story As Range
    Dim storyRng As Range
    Dim searchRng As Range
    
    Dim i As Long
    Dim currentName As String
    Dim currentStyleName As String
    Dim targetStyle As Long
    Dim headingNum As String
    Dim headingPos As Long
    Dim storyEnd As Long
    
    Dim stylesRemoved As Long
    Dim stylesRemaining As Long
    
    '-------------------------------------------------------------------------
    ' LIST PRESERVATION STORAGE
    '-------------------------------------------------------------------------
    Dim hasList As Boolean
    Dim savedListTemplate As ListTemplate
    Dim savedListLevel As Long
    Dim continuePreviousList As Boolean
    
    '-------------------------------------------------------------------------
    ' FONT FORMATTING STORAGE
    '-------------------------------------------------------------------------
    Dim fName As Variant
    Dim fSize As Variant
    Dim fBold As Variant
    Dim fItalic As Variant
    Dim fUnderline As Variant
    Dim fColor As Variant
    Dim fAllCaps As Variant
    Dim fSmallCaps As Variant
    Dim fStrike As Variant
    Dim fDoubleStrike As Variant
    Dim fSuper As Variant
    Dim fSub As Variant
    Dim fHidden As Variant
    Dim fSpacing As Variant
    Dim fScaling As Variant
    
    '-------------------------------------------------------------------------
    ' PARAGRAPH FORMATTING STORAGE
    '-------------------------------------------------------------------------
    Dim pBefore As Single
    Dim pAfter As Single
    Dim pLineRule As Long
    Dim pLineSpacing As Single
    Dim pAlignment As Long
    Dim pFirstIndent As Single
    Dim pLeftIndent As Single
    Dim pRightIndent As Single
    Dim pKeepNext As Long
    Dim pKeepTogether As Long
    Dim pWidow As Long
    Dim pPageBreak As Long
    
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '=========================================================================
    ' PROCESS CUSTOM STYLES BACKWARDS THROUGH THE STYLE COLLECTION
    '=========================================================================
    ' Backward processing allows custom styles to be deleted safely without
    ' disturbing the remaining collection indexes.
    For i = doc.Styles.Count To 1 Step -1
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(i)
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
            ' Built-in Word styles are always protected
            If Not sty.BuiltIn Then
                
                currentName = sty.NameLocal
                
                '=================================================================
                ' PHASE 1: CUSTOM PARAGRAPH / LINKED STYLES
                '=================================================================
                If sty.Type = wdStyleTypeParagraph Or _
                   sty.Type = wdStyleTypeLinked Then
                    
                    '-------------------------------------------------------------
                    ' Determine destination style.
                    '
                    ' Any custom style containing "Heading 1" through "Heading 9"
                    ' is mapped to the corresponding built-in Word heading.
                    ' Everything else maps to Normal.
                    '-------------------------------------------------------------
                    targetStyle = wdStyleNormal
                    
                    headingPos = InStr(1, currentName, "Heading ", vbTextCompare)
                    
                    If headingPos > 0 Then
                        
                        headingNum = Mid$(currentName, headingPos + 8, 1)
                        
                        If IsNumeric(headingNum) Then
                            
                            Select Case CLng(headingNum)
                                Case 1: targetStyle = wdStyleHeading1
                                Case 2: targetStyle = wdStyleHeading2
                                Case 3: targetStyle = wdStyleHeading3
                                Case 4: targetStyle = wdStyleHeading4
                                Case 5: targetStyle = wdStyleHeading5
                                Case 6: targetStyle = wdStyleHeading6
                                Case 7: targetStyle = wdStyleHeading7
                                Case 8: targetStyle = wdStyleHeading8
                                Case 9: targetStyle = wdStyleHeading9
                            End Select
                            
                        End If
                        
                    End If
                    
                    '-------------------------------------------------------------
                    ' Search paragraphs throughout all Word story ranges.
                    '-------------------------------------------------------------
                    For Each story In doc.StoryRanges
                        
                        Set storyRng = story
                        
                        Do While Not storyRng Is Nothing
                            
                            For Each para In storyRng.Paragraphs
                                
                                currentStyleName = ""
                                
                                On Error Resume Next
                                currentStyleName = para.Style.NameLocal
                                Err.Clear
                                On Error GoTo ErrorHandler
                                
                                If StrComp(currentStyleName, currentName, _
                                           vbTextCompare) = 0 Then
                                    
                                    '-----------------------------------------
                                    ' EQUATION GUARDRAIL
                                    '-----------------------------------------
                                    If para.Range.OMaths.Count > 0 Then
                                        GoTo SkipParagraph
                                    End If
                                    
                                    '-----------------------------------------
                                    ' Preserve list structure.
                                    '-----------------------------------------
                                    hasList = False
                                    Set savedListTemplate = Nothing
                                    continuePreviousList = False
                                    
                                    On Error Resume Next
                                    
                                    If para.Range.ListFormat.ListType <> _
                                       wdListNoNumbering Then
                                        
                                        hasList = True
                                        
                                        savedListLevel = _
                                            para.Range.ListFormat.ListLevelNumber
                                        
                                        Set savedListTemplate = _
                                            para.Range.ListFormat.ListTemplate
                                        
                                        Set prevPara = para.Previous
                                        
                                        If Not prevPara Is Nothing Then
                                            
                                            If prevPara.Range.ListFormat.ListType <> _
                                               wdListNoNumbering Then
                                                
                                                continuePreviousList = True
                                                
                                            End If
                                            
                                        End If
                                        
                                    End If
                                    
                                    Err.Clear
                                    On Error GoTo ErrorHandler
                                    
                                    '-----------------------------------------
                                    ' Capture effective font formatting.
                                    '-----------------------------------------
                                    With para.Range.Font
                                        fName = .Name
                                        fSize = .Size
                                        fBold = .Bold
                                        fItalic = .Italic
                                        fUnderline = .Underline
                                        fColor = .Color
                                        fAllCaps = .AllCaps
                                        fSmallCaps = .SmallCaps
                                        fStrike = .StrikeThrough
                                        fDoubleStrike = .DoubleStrikeThrough
                                        fSuper = .Superscript
                                        fSub = .Subscript
                                        fHidden = .Hidden
                                        fSpacing = .Spacing
                                        fScaling = .Scaling
                                    End With
                                    
                                    '-----------------------------------------
                                    ' Capture paragraph formatting.
                                    '-----------------------------------------
                                    With para.Format
                                        pBefore = .SpaceBefore
                                        pAfter = .SpaceAfter
                                        pLineRule = .LineSpacingRule
                                        pLineSpacing = .LineSpacing
                                        pAlignment = .Alignment
                                        pFirstIndent = .FirstLineIndent
                                        pLeftIndent = .LeftIndent
                                        pRightIndent = .RightIndent
                                        pKeepNext = .KeepWithNext
                                        pKeepTogether = .KeepTogether
                                        pWidow = .WidowControl
                                        pPageBreak = .PageBreakBefore
                                    End With
                                    
                                    '-----------------------------------------
                                    ' Replace custom paragraph style.
                                    '-----------------------------------------
                                    para.Style = targetStyle
                                    
                                    '-----------------------------------------
                                    ' Restore list if style conversion removed
                                    ' bullets or numbering.
                                    '-----------------------------------------
                                    If hasList Then
                                        
                                        If para.Range.ListFormat.ListType = _
                                           wdListNoNumbering Then
                                            
                                            If Not savedListTemplate Is Nothing Then
                                                
                                                On Error Resume Next
                                                
                                                para.Range.ListFormat. _
                                                    ApplyListTemplateWithLevel _
                                                    ListTemplate:=savedListTemplate, _
                                                    ContinuePreviousList:=continuePreviousList, _
                                                    ApplyTo:=wdListApplyToSelection, _
                                                    DefaultListBehavior:=wdWord10ListBehavior, _
                                                    ApplyLevel:=savedListLevel
                                                
                                                Err.Clear
                                                On Error GoTo ErrorHandler
                                                
                                            End If
                                            
                                        End If
                                        
                                    End If
                                    
                                    '-----------------------------------------
                                    ' Restore paragraph formatting.
                                    '-----------------------------------------
                                    With para.Format
                                        .SpaceBefore = pBefore
                                        .SpaceAfter = pAfter
                                        .LineSpacingRule = pLineRule
                                        .LineSpacing = pLineSpacing
                                        .Alignment = pAlignment
                                        .KeepWithNext = pKeepNext
                                        .KeepTogether = pKeepTogether
                                        .WidowControl = pWidow
                                        .PageBreakBefore = pPageBreak
                                        
                                        ' List indentation remains controlled
                                        ' by the original list definition.
                                        If Not hasList Then
                                            .FirstLineIndent = pFirstIndent
                                            .LeftIndent = pLeftIndent
                                            .RightIndent = pRightIndent
                                        End If
                                        
                                    End With
                                    
                                    '-----------------------------------------
                                    ' Restore effective font formatting.
                                    '-----------------------------------------
                                    With para.Range.Font
                                        
                                        If fName <> "" Then .Name = fName
                                        If fSize <> wdUndefined Then .Size = fSize
                                        If fBold <> wdUndefined Then .Bold = fBold
                                        If fItalic <> wdUndefined Then .Italic = fItalic
                                        
                                        If fUnderline <> wdUndefined Then _
                                            .Underline = fUnderline
                                        
                                        If fColor <> wdUndefined Then .Color = fColor
                                        
                                        If fAllCaps <> wdUndefined Then _
                                            .AllCaps = fAllCaps
                                        
                                        If fSmallCaps <> wdUndefined Then _
                                            .SmallCaps = fSmallCaps
                                        
                                        If fStrike <> wdUndefined Then _
                                            .StrikeThrough = fStrike
                                        
                                        If fDoubleStrike <> wdUndefined Then _
                                            .DoubleStrikeThrough = fDoubleStrike
                                        
                                        If fSuper <> wdUndefined Then _
                                            .Superscript = fSuper
                                        
                                        If fSub <> wdUndefined Then _
                                            .Subscript = fSub
                                        
                                        If fHidden <> wdUndefined Then _
                                            .Hidden = fHidden
                                        
                                        If fSpacing <> wdUndefined Then _
                                            .Spacing = fSpacing
                                        
                                        If fScaling <> wdUndefined Then _
                                            .Scaling = fScaling
                                        
                                    End With
                                    
                                End If
                                
SkipParagraph:
                                
                            Next para
                            
                            Set storyRng = storyRng.NextStoryRange
                            
                        Loop
                        
                    Next story
                    
                '=================================================================
                ' PHASE 2: CUSTOM CHARACTER STYLES
                '=================================================================
                ElseIf sty.Type = wdStyleTypeCharacter Then
                    
                    For Each story In doc.StoryRanges
                        
                        Set storyRng = story
                        
                        Do While Not storyRng Is Nothing
                            
                            Set searchRng = storyRng.Duplicate
                            storyEnd = searchRng.End
                            
                            With searchRng.Find
                                .ClearFormatting
                                .Replacement.ClearFormatting
                                .Text = ""
                                .Style = sty
                                .Forward = True
                                .Wrap = wdFindStop
                                .Format = True
                            End With
                            
                            Do While searchRng.Find.Execute
                                
                                '---------------------------------------------
                                ' EQUATION GUARDRAIL
                                '---------------------------------------------
                                If searchRng.OMaths.Count > 0 Then
                                    
                                    searchRng.Collapse wdCollapseEnd
                                    searchRng.End = storyEnd
                                    GoTo ContinueCharacterSearch
                                    
                                End If
                                
                                '---------------------------------------------
                                ' Capture current character appearance.
                                '---------------------------------------------
                                With searchRng.Font
                                    fName = .Name
                                    fSize = .Size
                                    fBold = .Bold
                                    fItalic = .Italic
                                    fUnderline = .Underline
                                    fColor = .Color
                                    fAllCaps = .AllCaps
                                    fSmallCaps = .SmallCaps
                                    fStrike = .StrikeThrough
                                    fDoubleStrike = .DoubleStrikeThrough
                                    fSuper = .Superscript
                                    fSub = .Subscript
                                    fHidden = .Hidden
                                    fSpacing = .Spacing
                                    fScaling = .Scaling
                                End With
                                
                                '---------------------------------------------
                                ' Remove custom character style.
                                '---------------------------------------------
                                searchRng.Style = wdStyleDefaultParagraphFont
                                
                                '---------------------------------------------
                                ' Restore appearance as direct formatting.
                                '---------------------------------------------
                                With searchRng.Font
                                    
                                    If fName <> "" Then .Name = fName
                                    If fSize <> wdUndefined Then .Size = fSize
                                    If fBold <> wdUndefined Then .Bold = fBold
                                    If fItalic <> wdUndefined Then .Italic = fItalic
                                    
                                    If fUnderline <> wdUndefined Then _
                                        .Underline = fUnderline
                                    
                                    If fColor <> wdUndefined Then .Color = fColor
                                    
                                    If fAllCaps <> wdUndefined Then _
                                        .AllCaps = fAllCaps
                                    
                                    If fSmallCaps <> wdUndefined Then _
                                        .SmallCaps = fSmallCaps
                                    
                                    If fStrike <> wdUndefined Then _
                                        .StrikeThrough = fStrike
                                    
                                    If fDoubleStrike <> wdUndefined Then _
                                        .DoubleStrikeThrough = fDoubleStrike
                                    
                                    If fSuper <> wdUndefined Then _
                                        .Superscript = fSuper
                                    
                                    If fSub <> wdUndefined Then _
                                        .Subscript = fSub
                                    
                                    If fHidden <> wdUndefined Then _
                                        .Hidden = fHidden
                                    
                                    If fSpacing <> wdUndefined Then _
                                        .Spacing = fSpacing
                                    
                                    If fScaling <> wdUndefined Then _
                                        .Scaling = fScaling
                                    
                                End With
                                
                                searchRng.Collapse wdCollapseEnd
                                searchRng.End = storyEnd
                                
ContinueCharacterSearch:
                                
                            Loop
                            
                            Set storyRng = storyRng.NextStoryRange
                            
                        Loop
                        
                    Next story
                    
                '=================================================================
                ' PHASE 3: CUSTOM TABLE STYLES
                '=================================================================
                ElseIf sty.Type = wdStyleTypeTable Then
                    
                    For Each tbl In doc.Tables
                        
                        currentStyleName = ""
                        
                        On Error Resume Next
                        currentStyleName = tbl.Style.NameLocal
                        Err.Clear
                        On Error GoTo ErrorHandler
                        
                        If StrComp(currentStyleName, currentName, _
                                   vbTextCompare) = 0 Then
                            
                            '---------------------------------------------
                            ' Preserve cell paragraph/text formatting.
                            '---------------------------------------------
                            For Each cel In tbl.Range.Cells
                                
                                For Each cellPara In cel.Range.Paragraphs
                                    
                                    If cellPara.Range.OMaths.Count > 0 Then
                                        GoTo SkipCellParagraph
                                    End If
                                    
                                    With cellPara.Format
                                        pBefore = .SpaceBefore
                                        pAfter = .SpaceAfter
                                        pLineRule = .LineSpacingRule
                                        pLineSpacing = .LineSpacing
                                        pAlignment = .Alignment
                                        pFirstIndent = .FirstLineIndent
                                        pLeftIndent = .LeftIndent
                                        pRightIndent = .RightIndent
                                    End With
                                    
                                    With cellPara.Range.Font
                                        fName = .Name
                                        fSize = .Size
                                        fBold = .Bold
                                        fItalic = .Italic
                                        fUnderline = .Underline
                                        fColor = .Color
                                    End With
                                    
                                    With cellPara.Format
                                        .SpaceBefore = pBefore
                                        .SpaceAfter = pAfter
                                        .LineSpacingRule = pLineRule
                                        .LineSpacing = pLineSpacing
                                        .Alignment = pAlignment
                                        .FirstLineIndent = pFirstIndent
                                        .LeftIndent = pLeftIndent
                                        .RightIndent = pRightIndent
                                    End With
                                    
                                    With cellPara.Range.Font
                                        
                                        If fName <> "" Then .Name = fName
                                        If fSize <> wdUndefined Then .Size = fSize
                                        If fBold <> wdUndefined Then .Bold = fBold
                                        If fItalic <> wdUndefined Then .Italic = fItalic
                                        
                                        If fUnderline <> wdUndefined Then _
                                            .Underline = fUnderline
                                        
                                        If fColor <> wdUndefined Then .Color = fColor
                                        
                                    End With
                                    
SkipCellParagraph:
                                    
                                Next cellPara
                                
                            Next cel
                            
                            ' Replace custom table style
                            On Error Resume Next
                            tbl.Style = wdStyleTableNormal
                            Err.Clear
                            On Error GoTo ErrorHandler
                            
                        End If
                        
                    Next tbl
                    
                '=================================================================
                ' PHASE 4: CUSTOM LIST STYLES
                '=================================================================
                ElseIf sty.Type = wdStyleTypeList Then
                    
                    ' List styles are deliberately not force-deleted here.
                    ' Their numbering may still be referenced by paragraphs.
                    '
                    ' Paragraph/list cleanup above removes custom paragraph
                    ' styling while retaining active ListFormat definitions.
                    GoTo KeepCustomStyle
                    
                End If
                
                '=================================================================
                ' DELETE CUSTOM STYLE IF IT IS NOW UNUSED
                '=================================================================
                On Error Resume Next
                
                If Not sty.InUse Then
                    
                    sty.Delete
                    
                    If Err.Number = 0 Then
                        stylesRemoved = stylesRemoved + 1
                    Else
                        stylesRemaining = stylesRemaining + 1
                        Err.Clear
                    End If
                    
                Else
                    
                    stylesRemaining = stylesRemaining + 1
                    
                End If
                
                On Error GoTo ErrorHandler
                
KeepCustomStyle:
                
            End If
            
        End If
        
    Next i

CleanUp:
    Application.ScreenUpdating = True
    
    If stylesRemaining = 0 Then
        
        MsgBox "Comprehensive style cleanup completed successfully." & _
               vbCrLf & vbCrLf & _
               stylesRemoved & " custom styles removed." & vbCrLf & _
               "Formatting, lists, tables, and equations were preserved.", _
               vbInformation, "Style Cleanup Complete"
        
    Else
        
        MsgBox "Comprehensive style cleanup completed." & _
               vbCrLf & vbCrLf & _
               stylesRemoved & " custom styles removed." & vbCrLf & _
               stylesRemaining & " custom style(s) remain in use." & _
               vbCrLf & vbCrLf & _
               "Equation-dependent and active list styles are intentionally retained.", _
               vbInformation, "Style Cleanup Complete"
        
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Comprehensive Style Cleanup Error"
End Sub

Sub Style_2_Del_Unused_Styles_Optimized()
'=============================================================================
' Name: Style_2_Del_Unused_Styles_Optimized
' Purpose: Quickly finds and removes unused custom styles across all document
'          story ranges while preserving all built-in Microsoft Word styles.
'
'-----------------------------------------------------------------------------
' SETTINGS
'-----------------------------------------------------------------------------
' dryRun = False   -> Actually deletes unused custom styles
' dryRun = True    -> Reports what WOULD be deleted without modifying document
'
' SEARCH SCOPE:
' Main document, headers, footers, footnotes, endnotes, text frames, etc.
'
' FINAL SUMMARY:
'   - Custom styles deleted / would be deleted
'   - Built-in styles deleted
'   - Custom styles remaining
'   - Built-in styles remaining
'
' NOTE:
' Built-in styles are always protected, so Built-in Styles Deleted = 0.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim story As Range
    Dim storyRng As Range
    Dim findRng As Range
    
    Dim i As Long
    Dim pass As Integer
    
    Dim deletedCustomCount As Long
    Dim deletedBuiltInCount As Long
    Dim remainingCustomCount As Long
    Dim remainingBuiltInCount As Long
    
    Dim styleUsed As Boolean
    Dim reportText As String
    
    Dim dryRun As Boolean
    
    Set doc = ActiveDocument
    
    '-------------------------------------------------------------------------
    ' USER SETTING
    '-------------------------------------------------------------------------
    dryRun = False
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' TWO-PASS CLEANUP
    '-------------------------------------------------------------------------
    ' A second pass can remove styles that were previously retained because
    ' another custom style depended on them.
    For pass = 1 To 2
        
        For i = doc.Styles.Count To 1 Step -1
            
            Set sty = Nothing
            
            On Error Resume Next
            Set sty = doc.Styles(i)
            Err.Clear
            On Error GoTo ErrorHandler
            
            If Not sty Is Nothing Then
                
                '-------------------------------------------------------------
                ' BUILT-IN STYLES ARE ALWAYS PROTECTED
                '-------------------------------------------------------------
                If Not sty.BuiltIn Then
                    
                    styleUsed = False
                    
                    '---------------------------------------------------------
                    ' RULE 1: Trust Word when it reports the style as unused.
                    '---------------------------------------------------------
                    If Not sty.InUse Then
                        
                        styleUsed = False
                        
                    Else
                        
                        '-----------------------------------------------------
                        ' RULE 2: Deep search all story ranges.
                        '-----------------------------------------------------
                        For Each story In doc.StoryRanges
                            
                            Set storyRng = story
                            
                            Do While Not storyRng Is Nothing
                                
                                Set findRng = storyRng.Duplicate
                                
                                On Error Resume Next
                                
                                With findRng.Find
                                    .ClearFormatting
                                    .Replacement.ClearFormatting
                                    .Text = ""
                                    .Style = sty
                                    .Forward = True
                                    .Wrap = wdFindStop
                                    .Format = True
                                End With
                                
                                If findRng.Find.Execute Then
                                    styleUsed = True
                                End If
                                
                                Err.Clear
                                On Error GoTo ErrorHandler
                                
                                If styleUsed Then Exit Do
                                
                                Set storyRng = storyRng.NextStoryRange
                                
                            Loop
                            
                            If styleUsed Then Exit For
                            
                        Next story
                        
                    End If
                    
                    '---------------------------------------------------------
                    ' DELETE OR COUNT UNUSED CUSTOM STYLE
                    '---------------------------------------------------------
                    If Not styleUsed Then
                        
                        If dryRun Then
                            
                            deletedCustomCount = deletedCustomCount + 1
                            
                        Else
                            
                            On Error Resume Next
                            Err.Clear
                            
                            sty.Delete
                            
                            If Err.Number = 0 Then
                                deletedCustomCount = deletedCustomCount + 1
                            End If
                            
                            Err.Clear
                            On Error GoTo ErrorHandler
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        Next i
        
    Next pass
    
    '-------------------------------------------------------------------------
    ' COUNT ALL STYLES REMAINING
    '-------------------------------------------------------------------------
    For Each sty In doc.Styles
        
        If sty.BuiltIn Then
            remainingBuiltInCount = remainingBuiltInCount + 1
        Else
            remainingCustomCount = remainingCustomCount + 1
        End If
        
    Next sty
    
    ' Built-in styles are never targeted by this macro
    deletedBuiltInCount = 0
    
    Application.ScreenUpdating = True
    
    '-------------------------------------------------------------------------
    ' BUILD FINAL SUMMARY
    '-------------------------------------------------------------------------
    If dryRun Then
        
        reportText = _
            "Style cleanup dry run completed." & vbCrLf & vbCrLf & _
            "Custom styles to be deleted: " & deletedCustomCount & vbCrLf & _
            "Built-in styles to be deleted: " & deletedBuiltInCount & vbCrLf & _
            "Custom styles remaining: " & remainingCustomCount & vbCrLf & _
            "Built-in styles remaining: " & remainingBuiltInCount
        
        MsgBox reportText, vbInformation, "Dry Run Complete"
        
    Else
        
        reportText = _
            "Style cleanup completed successfully." & vbCrLf & vbCrLf & _
            "Custom styles deleted: " & deletedCustomCount & vbCrLf & _
            "Built-in styles deleted: " & deletedBuiltInCount & vbCrLf & _
            "Custom styles remaining: " & remainingCustomCount & vbCrLf & _
            "Built-in styles remaining: " & remainingBuiltInCount
        
        MsgBox reportText, vbInformation, "Clean Up Complete"
        
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Style Cleanup Error"
End Sub

Sub Style_3_Reset_BuiltIn_Style_Names()
'=============================================================================
' Name: Style_3_Reset_BuiltIn_Style_Names
' Purpose: Finds built-in system styles containing custom appended alias text
'          and strings (via commas), breaks them, and reverts them back
'          to factory default settings.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim cleanName As String
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' Evaluate styles array systematically backwards
    For i = doc.styles.Count To 1 Step -1
        Set sty = doc.styles(i)
        
        ' Process exclusively native built-in Microsoft Word system properties
        If sty.BuiltIn Then
            
            ' Comma detection flags when custom names have been attached over defaults
            If InStr(1, sty.NameLocal, ",", vbTextCompare) > 0 Then
                
                ' Extract character string indexing prior to the first comma separator
                cleanName = Split(sty.NameLocal, ",")(0)
                
                ' Assign the base native string back onto the local workspace setting
                sty.NameLocal = cleanName
                Debug.Print "Reset style name back to: " & cleanName
            End If
            
        End If
    Next i
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Built-in style names have been reverted to their original factory names!", vbInformation, "Reset Complete"
End Sub

Sub Style_4_Adjust_Styles()
'=============================================================================
' Name: Style_4_Adjust_Styles
' Purpose: Explicitly configures and standardizes core body styles (Normal,
'          Normal (Web), Body Text, etc.), Heading styles (1 through 4),
'          Caption style, and Hyperlinks (idle and visited).
'          Establishes layout baselines, clears rogue tab stops, strips out
'          any legacy/manual paragraph borders, and ensures clean text geometries.
' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
' PERFORMANCE:  Modifies named stylesheet assets directly in memory, bypassing
'               the need to loop paragraph-by-paragraph or move the cursor.
'=============================================================================
    Dim doc As Document
    Dim normalStyleNames As Variant
    Dim headingNames As Variant
    Dim linkStyleNames As Variant
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
        With doc.Styles(stName)
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
    With doc.Styles("Heading 1")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        With .Font
            .Name = "Calibri"
            .Size = 18
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
            .Alignment = wdAlignParagraphLeft      ' Left-aligned for a clean layout break
            .OutlineLevel = wdOutlineLevel1        ' Mandatory tier assignment for core TOC extraction
            .PageBreakBefore = True                ' Automatically pushes major sections to a new page
            .KeepWithNext = True                   ' Locks heading onto the same page as the following body text
            .KeepTogether = True                   ' Prevents heading text lines from splitting across pages
            .TabStops.ClearAll
            .Borders.Enable = False                ' Safe structural border clear
        End With
    End With

    '-------------------------------------------------------------------------
    ' 3. HEADING 2 (Sub-sections - Left-Aligned & Bound to Following Text)
    '-------------------------------------------------------------------------
    With doc.Styles("Heading 2")
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
    With doc.Styles("Heading 3")
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
    With doc.Styles("Heading 4")
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
    With doc.Styles("Caption")
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
        With doc.Styles(headingNames(i)).Font
            ' ACTIVE CONFIGURATION: Apply custom hex color #182C52 natively
            '.Color = RGB(24, 44, 82)
            
            ' ROLLBACK TOGGLE: Uncomment the line below to easily reset everything back to Automatic
            .Color = wdColorAutomatic
        End With
    Next i

    '-------------------------------------------------------------------------
    ' 8. HYPERLINK STYLES PASS (Standardizes Idle & Visited/Pressed Links)
    '-------------------------------------------------------------------------
    ' "Hyperlink" = Unvisited / Idle link style
    ' "FollowedHyperlink" = Visited / Pressed link style
    linkStyleNames = Array("Hyperlink", "FollowedHyperlink")
    
    For Each stName In linkStyleNames
        On Error Resume Next
        With doc.Styles(stName).Font
            .Color = RGB(5, 99, 193)    ' Sets custom hex color #0563C1
            .Underline = wdUnderlineSingle
        End With
        On Error GoTo ErrorHandler
    Next stName

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

Sub Style_5_Apply_Styles_To_Document()
'=============================================================================
' Name: Style_5_Apply_Styles_To_Document
' Purpose: Executes a 3-phase optimization pipeline:
'          1. Direct formats all body text to smash manual style overrides.
'          2. Restores tight 1.15pt spacing and 0/0 padding to all Tables.
'          3. Scans paragraphs in a single pass via Outline Levels to capture
'             true headings AND "fake" headings, resetting their fonts.
'=============================================================================
    Dim doc As Document
    Dim tbl As Table
    Dim para As Paragraph
    Dim rng As Range
    Dim outLvl As Long
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates, animations, and repainting
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' PHASE 1: GLOBAL DIRECT FORMATTING OVERRIDE
    '-------------------------------------------------------------------------
    ' Uniformly applies baseline body formatting to clear stubborn layout drift
    Set rng = doc.Content
    With rng
        With .Font
            .Name = "Calibri"
            .Size = 11

            ' ======================================================================
            ' Advanced Settings
            ' ======================================================================
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .Ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
        End With
        With .ParagraphFormat
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
        End With
    End With

    '-------------------------------------------------------------------------
    ' PHASE 2: TABLE PROTECTION LOOP
    '-------------------------------------------------------------------------
    ' Immediately restores tight single-line spacing inside all tables
    For i = 3 To doc.Sections.Count
        For Each tbl In doc.Sections(i).Range.Tables
            ' PHASE 1: Localized Text Spacing Reset
            With tbl.Range.ParagraphFormat
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)     ' Dynamically calculates spacing based on font size

            ' If you want to use a specific line spacing value instead
            ' of single spacing, you can uncomment the following line
            ' and set your desired spacing in points
            ' --------------------------------------
            ' .LineSpacingRule = wdLineSpaceMultiple
            ' .LineSpacing = LinesToPoints(1.15)
            End With
            
            ' PHASE 2: Dynamic Row Height Clamping
            ' Inline local error bypass to prevent vertically merged cells from crashing the script
            On Error Resume Next 
            tbl.Rows.Height = 0
            tbl.Rows.HeightRule = wdRowHeightAuto
            On Error GoTo ErrorHandler
            
        Next tbl
    Next i

    '-------------------------------------------------------------------------
    ' PHASE 3: RESTORE & UP-CONVERT HEADING STYLES VIA OUTLINE LEVELS
    '-------------------------------------------------------------------------
    ' Instead of checking paragraph style names 4 times over, we run a lightning-fast
    ' SINGLE pass through the document evaluating Outline Levels. This automatically
    ' updates true headings AND catches hidden "fake" headings using Normal style.
    For Each para In doc.Paragraphs
        
        ' Bypass table contents to ensure data cells are never converted into headings
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Fetch the paragraph's structural Outline Level
            outLvl = para.OutlineLevel
            
            ' Process exclusively if it maps to levels 1, 2, 3, or 4
            If outLvl >= 1 And outLvl <= 4 Then
                With para.Range
                    ' 1. Peel off Phase 1's direct formatting tape (Calibri 11pt override)
                    .Font.Reset
                    
                    ' 2. Force apply the true built-in Heading Style based on the level digit
                    .Style = doc.styles("Heading " & outLvl)
                End With
            End If
            
        End If
    Next para

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document successfully formatted, tables preserved, and heading structures unified!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Formatting Error"
End Sub


Sub Style_6_Highlight_Duplicate_Styles()
'=============================================================================
' Name: Style_6_Highlight_Duplicate_Styles
' Purpose: Evaluates the active workspace document for common corrupted style
'          clones (e.g., "Char" artifacts from copy-pasting text). Forces a
'          vibrant highlight on text fragments utilizing them for easy auditing.
'=============================================================================
    Dim sty As Style
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    ' Clear out any residual content highlighting to create a baseline canvas
    doc.Content.HighlightColorIndex = wdNoHighlight
    
    ' Cycle custom document style structures
    For Each sty In doc.styles
        If Not sty.BuiltIn Then
            
            ' Isolate known automated copy/paste duplicate naming signatures
            If InStr(1, sty.NameLocal, " Char", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, " pt", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, "Indent", vbTextCompare) > 0 Then
               
                ' Find specific string patterns matching the targeted clone signature
                With doc.Content.Find
                    .ClearFormatting
                    .Style = sty.NameLocal
                    .replacement.ClearFormatting
                    .replacement.Highlight = True       ' Enforces highlight application change
                    .replacement.text = "^&"            ' Keeps current alphanumeric values safe
                    
                    ' Apply visual highlights universally
                    .Execute Replace:=wdReplaceAll
                End With
                
                Debug.Print "Highlighted usage of: " & sty.NameLocal
            End If
        End If
    Next sty
    
    Application.ScreenUpdating = True
    MsgBox "Check for Bright Green highlights.", vbInformation, "Audit Map Set"
End Sub

Sub Style_7_Configure_Figure_Caption_KeepWithNext()
'=============================================================================
' Name: Style_7_Configure_Figure_Caption_KeepWithNext
' Purpose: Scans all paragraphs styled as "Caption". If the text begins with
'          "Figure", it forces KeepWithNext to FALSE. This prevents the
'          caption from being pulled away from the visual asset above it.
'=============================================================================
    Dim doc As Document
    Dim para As Paragraph
    Dim paraText As String
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates during processing
    Application.ScreenUpdating = False
    
    ' Loop through every paragraph in the document
    For Each para In doc.Paragraphs
        
        ' Rule 1: Only look at paragraphs assigned to the standard "Caption" style
        If para.Style = doc.styles("Caption") Then
            
            ' Clean and trim the text to look at the first word safely
            paraText = Trim(para.Range.text)
            
            ' Rule 2: Check if the caption specifically starts with "Figure"
            ' (Using UCase and Left handles variations like "Figure 1", "Figure 2.1", etc.)
            If UCase(Left(paraText, 6)) = "FIGURE" Or _
               UCase(Left(paraText, 5)) = "PHOTO" Then
                
                ' FIXED: Call KeepWithNext directly on the Paragraph object.
                ' This disconnects the caption from the paragraph below it.
                para.KeepWithNext = False
                
            End If
        End If
    Next para
    
    ' Restore system screen updating
    Application.ScreenUpdating = True
    
    MsgBox "Figure caption page layout bounds configured successfully!", vbInformation, "Layout Complete"
End Sub

Sub Style_8_Apply_Styles_To_Document_V2()
'=============================================================================
' Name: Style_8_Apply_Styles_To_Document_V2()
' Purpose: Executes a fully consolidated multi-phase document layout optimization:
'          1. Direct formats all body text to smash unmanaged layout drifts.
'          2. Resores tight 1.0 spacing rules to all tabular cell grids.
'          3. Iterates paragraphs once to calculate context-aware list block values
'             (Before=0, Inside=0, After=6, Intro Tightening=0).
'          4. Resolves true and "fake" structural heading paths via Outline Levels.
'=============================================================================
    Dim doc As Document
    Dim tbl As Table
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim nextPara As Paragraph
    Dim rng As Range
    Dim outLvl As Long
    Dim isLastItem As Boolean
    
    Set doc = ActiveDocument
    
    ' Speed optimization: Turn off screen updates, animations, and repainting
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' PHASE 1: GLOBAL DIRECT FORMATTING OVERRIDE
    '-------------------------------------------------------------------------
    ' Uniformly applies baseline body formatting to clear stubborn layout drift
    Set rng = doc.Content
    With rng
        With .Font
            .Name = "Calibri"
            .Size = 11

            ' ======================================================================
            ' Advanced Settings
            ' ======================================================================
            .Spacing = 0                                ' Resets any manual character spacing adjustments
            .Scaling = 100                              ' Resets any manual font scaling adjustments
            .Kerning = 0                                ' Resets any manual kerning adjustments
            .Ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
            .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
            .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments (e.g., old-style vs. lining)
            .StylisticSet = wdStylisticSetDefault       ' Resets any manual stylistic set selections
            .ContextualAlternates = 0                   ' Disables any unwanted contextual alternate glyph substitutions
        End With
        With .ParagraphFormat
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
        End With
    End With

    '-------------------------------------------------------------------------
    ' PHASE 2: TABLE PROTECTION LOOP
    '-------------------------------------------------------------------------
    ' Immediately restores tight single-line spacing inside all tables
    For i = 3 To doc.Sections.Count
        For Each tbl In doc.Sections(i).Range.Tables
            ' PHASE 1: Localized Text Spacing Reset
            With tbl.Range.ParagraphFormat
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)     ' Dynamically calculates spacing based on font size
            End With
            
            ' PHASE 2: Dynamic Row Height Clamping
            ' Inline local error bypass to prevent vertically merged cells from crashing the script
            On Error Resume Next 
            tbl.Rows.Height = 0
            tbl.Rows.HeightRule = wdRowHeightAuto
            On Error GoTo ErrorHandler
            
        Next tbl
    Next i

    '-------------------------------------------------------------------------
    ' CONSOLIDATED SCANNING ENGINE: LIST SPACING & OUTLINE CONVERSIONS
    '-------------------------------------------------------------------------
    For Each para In doc.Paragraphs
        
        '=====================================================================
        ' SUB-PHASE A: ADVANCED CONTEXTUAL LIST ADJUSTMENTS
        '=====================================================================
        ' HARD GUARDRAIL: Skip paragraph completely if it is an active Heading element
        If para.OutlineLevel = wdOutlineLevelBodyText Then
        
            ' Rule 1: Target active list formatting structures (Bullets, Numbers, Outlines)
            If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                
                ' Rule 2: Strictly protect tables by ignoring internal table lists
                If Not para.Range.Information(wdWithInTable) Then
                    
                    ' STEP 1: LOOK-BEHIND (Tighten Intro Paragraph)
                    Set prevPara = para.Previous
                    If Not prevPara Is Nothing Then
                        ' If the preceding line is NOT a list, this is the FIRST item in the block!
                        If prevPara.Range.ListFormat.ListType = wdListNoNumbering Then
                            ' Ensure it isn't a table or a structural heading before modifying it
                            If Not prevPara.Range.Information(wdWithInTable) And _
                               (prevPara.OutlineLevel >= wdOutlineLevelBodyText) Then
                                
                                prevPara.SpaceAfterAuto = False
                                prevPara.SpaceAfter = 0 ' Snaps the introductory text tightly down
                                
                            End If
                        End If
                    End If
    
                    ' STEP 2: Enforce Base List Geometries
                    para.SpaceBeforeAuto = False
                    para.SpaceAfterAuto = False
                    para.SpaceBefore = 0
                    para.LineSpacingRule = wdLineSpace1pt5
                    
                    ' STEP 3: LOOK-AHEAD (Determine Block End Spacing)
                    Set nextPara = para.Next
                    isLastItem = False ' Reset flag for current paragraph
                    
                    ' Condition 1: There is no next paragraph (End of Document)
                    If nextPara Is Nothing Then
                        isLastItem = True
                    Else
                        ' Condition 2: The next paragraph is generic body text
                        If nextPara.Range.ListFormat.ListType = wdListNoNumbering Then
                            isLastItem = True
                        ' Condition 3: The next line drops out of main text space and into a table
                        ElseIf nextPara.Range.Information(wdWithInTable) Then
                            isLastItem = True
                        End If
                        
                        ' Condition 4: The next line is explicitly a Heading (Outline Levels 1 to 9)
                        If nextPara.OutlineLevel >= 1 And nextPara.OutlineLevel <= 9 Then
                            isLastItem = True
                        End If
                    End If
                    
                    ' STEP 4: Apply Calculated Spacing Execution
                    If isLastItem Then
                        para.SpaceAfter = 6   ' Add professional breathing room at block end
                    Else
                        para.SpaceAfter = 0   ' Keep items tightly packed within the block
                    End If
                    
                End If
            End If
            
        End If
        
        '=====================================================================
        ' SUB-PHASE B: RESTORE & UP-CONVERT HEADING STYLES VIA OUTLINE LEVELS
        '=====================================================================
        ' Bypass table contents to ensure data cells are never converted into headings
        If Not para.Range.Information(wdWithInTable) Then
            
            ' Fetch the paragraph's structural Outline Level
            outLvl = para.OutlineLevel
            
            ' Process exclusively if it maps to levels 1, 2, 3, or 4
            If outLvl >= 1 And outLvl <= 4 Then
                With para.Range
                    ' 1. Peel off Phase 1's direct formatting tape (Calibri 11pt override)
                    .Font.Reset
                    
                    ' 2. Force apply the true built-in Heading Style based on the level digit
                    .Style = doc.styles("Heading " & outLvl)
                End With
            End If
            
        End If
        
    Next para

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document styles applied and list spaces manually balanced successfully!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Formatting Error"
End Sub


Sub Style_9_Apply_Styles_To_Document_V3()
'=============================================================================
' Name: Style_9_Apply_Styles_To_Document_V3()
' Purpose: Executes a fully consolidated multi-phase document layout optimization:
'          1. Direct formats all body text to smash unmanaged layout drifts
'             (SKIPS Sections 1 and 2 entirely).
'          2. Restores tight 1.0 spacing rules to all tabular cell grids.
'          3. Iterates paragraphs sequentially to calculate context-aware list block values
'             (Before=0, Inside=0, After=6, Intro Tightening=0).
'          4. Resolves true and "fake" structural heading paths via Outline Levels.
'          5. Detects Figure/Fig captions and turns off KeepWithNext/KeepTogether.
'=============================================================================
    Dim doc As Document
    Dim tbl As Table
    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim nextPara As Paragraph
    Dim rng As Range
    Dim outLvl As Long
    Dim isLastItem As Boolean
    Dim i As Long
    Dim sec As Section
    Dim cleanText As String
    
    Set doc = ActiveDocument
    
    ' Hard Guardrail: If the document doesn't have at least 3 sections, don't run.
    If doc.Sections.Count < 3 Then
        MsgBox "The document must contain at least 3 sections to format the content areas safely.", vbExclamation, "Execution Halted"
        Exit Sub
    End If
    
    ' Speed optimization: Turn off screen updates, animations, and repainting
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' PHASE 1: GLOBAL DIRECT FORMATTING OVERRIDE (TARGETED)
    '-------------------------------------------------------------------------
    ' Uniformly applies baseline body formatting starting precisely from Section 3
    ' to protect the front matter (Cover page & TOC) from direct overrides.
    For i = 3 To doc.Sections.Count
        Set rng = doc.Sections(i).Range
        With rng
            With .Font
                .Name = "Calibri"
                .Size = 11

                ' ======================================================================
                ' Advanced Settings
                ' ======================================================================
                .Spacing = 0                                ' Resets any manual character spacing adjustments
                .Scaling = 100                              ' Resets any manual font scaling adjustments
                .Kerning = 0                                ' Resets any manual kerning adjustments
                .Ligatures = wdLigaturesNone                ' Disables any unwanted ligature formations
                .NumberSpacing = wdNumberSpacingDefault     ' Resets any manual number spacing adjustments
                .NumberForm = wdNumberFormDefault           ' Resets any manual number form adjustments
                .StylisticSet = wdStylisticSetDefault         ' Resets any manual stylistic set selections
                .ContextualAlternates = 0                    ' Disables contextual alternate glyph substitutions
            End With
            With .ParagraphFormat
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacingRule = wdLineSpace1pt5
            End With
        End With
    Next i

    '-------------------------------------------------------------------------
    ' PHASE 2: TABLE PROTECTION LOOP
    '-------------------------------------------------------------------------
    ' Restores tight single-line spacing inside tables (only in Section 3 and beyond)
    For i = 3 To doc.Sections.Count
        For Each tbl In doc.Sections(i).Range.Tables
            ' PHASE 1: Localized Text Spacing Reset
            With tbl.Range.ParagraphFormat
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)     ' Dynamically calculates spacing based on font size
            End With
            
            ' PHASE 2: Dynamic Row Height Clamping
            ' Inline local error bypass to prevent vertically merged cells from crashing the script
            On Error Resume Next 
            tbl.Rows.Height = 0
            tbl.Rows.HeightRule = wdRowHeightAuto
            On Error GoTo ErrorHandler
            
        Next tbl
    Next i

    '-------------------------------------------------------------------------
    ' CONSOLIDATED SCANNING ENGINE: LIST SPACING, OUTLINE CONVERSIONS & CAPTIONS
    '-------------------------------------------------------------------------
    ' Sweeps paragraphs section-by-section starting cleanly at Section 3.
    For i = 3 To doc.Sections.Count
        For Each para In doc.Sections(i).Range.Paragraphs
            
            '=====================================================================
            ' SUB-PHASE A: ADVANCED CONTEXTUAL LIST ADJUSTMENTS
            '=====================================================================
            ' HARD GUARDRAIL: Skip paragraph completely if it is an active Heading element
            If para.OutlineLevel = wdOutlineLevelBodyText Then
            
                ' Rule 1: Target active list formatting structures (Bullets, Numbers, Outlines)
                If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                    
                    ' Rule 2: Strictly protect tables by ignoring internal table lists
                    If Not para.Range.Information(wdWithInTable) Then
                        
                        ' STEP 1: LOOK-BEHIND (Tighten Intro Paragraph)
                        Set prevPara = para.Previous
                        If Not prevPara Is Nothing Then
                            ' If the preceding line is NOT a list, this is the FIRST item in the block!
                            If prevPara.Range.ListFormat.ListType = wdListNoNumbering Then
                                ' Ensure it isn't a table or a structural heading before modifying it
                                If Not prevPara.Range.Information(wdWithInTable) And _
                                   (prevPara.OutlineLevel >= wdOutlineLevelBodyText) Then
                                    
                                    prevPara.SpaceAfterAuto = False
                                    prevPara.SpaceAfter = 0 ' Snaps the introductory text tightly down
                                    
                                End If
                            End If
                        End If
        
                        ' STEP 2: Enforce Base List Geometries
                        para.SpaceBeforeAuto = False
                        para.SpaceAfterAuto = False
                        para.SpaceBefore = 0
                        para.LineSpacingRule = wdLineSpace1pt5
                        
                        ' STEP 3: LOOK-AHEAD (Determine Block End Spacing)
                        Set nextPara = para.Next
                        isLastItem = False ' Reset flag for current paragraph
                        
                        ' Condition 1: There is no next paragraph (End of Document)
                        If nextPara Is Nothing Then
                            isLastItem = True
                        Else
                            ' Condition 2: The next paragraph is generic body text
                            If nextPara.Range.ListFormat.ListType = wdListNoNumbering Then
                                isLastItem = True
                            ' Condition 3: The next line drops out of main text space and into a table
                            ElseIf nextPara.Range.Information(wdWithInTable) Then
                                isLastItem = True
                            End If
                            
                            ' Condition 4: The next line is explicitly a Heading (Outline Levels 1 to 9)
                            If nextPara.OutlineLevel >= 1 And nextPara.OutlineLevel <= 9 Then
                                isLastItem = True
                            End If
                        End If
                        
                        ' STEP 4: Apply Calculated Spacing Execution
                        If isLastItem Then
                            para.SpaceAfter = 6   ' Add professional breathing room at block end
                        Else
                            para.SpaceAfter = 0   ' Keep items tightly packed within the block
                        End If
                        
                    End If
                End If
                
            End If
            
            '=====================================================================
            ' SUB-PHASE B: RESTORE & UP-CONVERT HEADING STYLES VIA OUTLINE LEVELS
            '=====================================================================
            ' Bypass table contents to ensure data cells are never converted into headings
            If Not para.Range.Information(wdWithInTable) Then
                
                ' Fetch the paragraph's structural Outline Level
                outLvl = para.OutlineLevel
                
                ' Process exclusively if it maps to levels 1, 2, 3, or 4
                If outLvl >= 1 And outLvl <= 4 Then
                    With para.Range
                        ' 1. Peel off Phase 1's direct formatting tape (Calibri 11pt override)
                        .Font.Reset
                        
                        ' 2. Force apply the true built-in Heading Style based on the level digit
                        .Style = doc.Styles("Heading " & outLvl)
                    End With
                End If
                
            End If

            '=====================================================================
            ' SUB-PHASE C: FIGURE CAPTION LAYOUT CONTROL
            '=====================================================================
            ' Detects figure captions and prevents them from pulling away from images above
            cleanText = Trim(para.Range.Text)
            ' Strip trailing paragraph mark / line feed break characters
            cleanText = Replace(cleanText, vbCr, "")
            cleanText = Replace(cleanText, vbLf, "")
            cleanText = Trim(cleanText)

            If LCase(cleanText) Like "figure*" Or LCase(cleanText) Like "fig*" Then
                ' Uncheck "Keep with next" and "Keep lines together" so the caption isn't pulled away from its image above
                para.KeepWithNext = False
                para.KeepTogether = False
            End If
            
        Next para
    Next i

CleanUp:
    ' Restore standard application window rendering
    Application.ScreenUpdating = True
    MsgBox "Document styles applied, list spaces balanced, and figure captions updated successfully (Sections 1 & 2 skipped)!", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    ' Gracefully restore screen rendering before throwing the runtime message box
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Formatting Error"
End Sub


Sub Style_10_Adjust_Headers_Footers_Styles()
'=============================================================================
' Name: Style_10_Adjust_Headers_Footers_Styles
' Purpose: Standardizes header and footer styles across the active document.
'          Resets font and paragraph formatting for built-in Header/Footer
'          styles so they align with the document's base typography.
'          Also clears layout noise such as rogue tab stops and legacy manual
'          paragraph borders to keep header/footer text geometry clean.
' COMPATIBILITY: Microsoft Word 2007 and newer (Word Layout Engine)
' PERFORMANCE:  Modifies named stylesheet assets directly in memory, bypassing
'               the need to loop paragraph-by-paragraph or move the cursor.
'=============================================================================
    Dim doc As Document
    Dim i As Long
    Dim stName As Variant
    Dim headerFooterStyleNames As Variant
    
    Set doc = ActiveDocument
    
    ' Freeze visual application window rendering to prevent layout redraw lag
    Application.ScreenUpdating = False
    
    ' Establish global runtime error trapping to protect the active workspace environment
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' HEADERS & FOOTERS (Global Header/Footer Style Reset)
    '-------------------------------------------------------------------------
    headerFooterStyleNames = Array( _
        "Header", _
        "Footer")
    
    For Each stName In headerFooterStyleNames
        ' Temporary error bypass in case a specific variant style does not exist in the document
        On Error Resume Next
        With doc.Styles(stName)
            .AutomaticallyUpdate = False
            With .Font
                ' Basic Font Properties
                .Name = "Calibri"
                ' .Size = 11
                ' .Bold = False
                ' .Italic = False
                ' .Color = wdColorAutomatic
                ' .Outline = False            ' Removes any unwanted borders around text characters
                ' .Shadow = False             ' Removes any legacy shadow tracking text effects
                ' .Emboss = False             ' Clears any manual embossing text effects
                ' .Engrave = False            ' Clears any manual engraving text effects
                
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
            ' With .ParagraphFormat
            '     .LineUnitBefore = 0
            '     .LineUnitAfter = 0
            '     .FirstLineIndent = InchesToPoints(0)
            '     .OutlineLevel = wdOutlineLevelBodyText
            '     .LeftIndent = InchesToPoints(0)
            '     .RightIndent = InchesToPoints(0)
            '     .SpaceBeforeAuto = False
            '     .SpaceAfterAuto = False
            '     .SpaceBefore = 0
            '     .SpaceAfter = 0            ' Commented out to allow for a small buffer between header/footer and body text
            '     .LineSpacingRule = wdLineSpaceSingle     ' Enforces single line heights for headers and footers
            '     .Alignment = wdAlignParagraphJustify    ' Justified text layout for reporting blocks
            '     .WidowControl = True                    ' Prevents orphan sentences at page boundaries
                
            '     ' ARCHITECTURAL STRATEGY: .Borders.Enable = False acts as a safe global clear pass.
            '     .Borders.Enable = False
            ' End With
        End With
        On Error GoTo ErrorHandler
    Next stName

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

Sub Style_11_Del_Unused_Styles_Improved()
'=============================================================================
' Name: Style_11_Del_Unused_Styles_Improved()
' Purpose: Efficiently removes unused non-built-in custom styles using a
'          staged cleanup strategy, then optionally exports deleted-style
'          information to Microsoft Excel.
'
'-----------------------------------------------------------------------------
' OPTIMIZATION STRATEGY
'-----------------------------------------------------------------------------
'
' Stage                         Action
' ---------------------------------------------------------------------------
' Pass 1 - Fast Cleanup         Delete styles where .InUse = False
' Pass 2 - Deep Validation      Check only remaining "InUse" custom styles
'                               and stop searching after the first real use
' Pass 3 - Dependency Cleanup   Quickly retry styles made unused by deletion
'
' This avoids counting every occurrence of every style and avoids performing
' expensive story-range searches unless Word claims the style is still active.
'
'-----------------------------------------------------------------------------
' FINAL SUMMARY
'-----------------------------------------------------------------------------
'
' Reports:
'   - Custom styles deleted
'   - Custom styles remaining
'   - Built-in styles remaining
'
' Optional Excel report:
'   - Style Name
'   - Style Type
'   - Deleted On Pass
'
'-----------------------------------------------------------------------------
' PRIVATE HELPERS USED BY THIS PARENT SUB
'-----------------------------------------------------------------------------
'
' Helper_Style_11_IsStyleActuallyUsed
'   Returns True immediately upon finding the first actual use of a style.
'
' Helper_Style_11_GetReadableStyleType
'   Converts WdStyleType into a readable description.
'
' Helper_Style_11_ExportDeletedStylesToExcel
'   Creates the optional Excel deletion report.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    
    Dim deletedNames As Collection
    Dim deletedTypes As Collection
    Dim deletedPasses As Collection
    
    Dim i As Long
    Dim deletedCount As Long
    Dim remainingCustomCount As Long
    Dim builtInCount As Long
    
    Dim styleName As String
    Dim styleTypeName As String
    
    Dim response As VbMsgBoxResult
    Dim reportText As String
    
    Set doc = ActiveDocument
    
    Set deletedNames = New Collection
    Set deletedTypes = New Collection
    Set deletedPasses = New Collection
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '=========================================================================
    ' PASS 1: FAST CLEANUP USING WORD'S InUse FLAG
    '=========================================================================
    ' This pass performs no Find operations. Styles Word already identifies
    ' as unused are deleted immediately.
    For i = doc.Styles.Count To 1 Step -1
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(i)
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
            If Not sty.BuiltIn Then
                
                If Not sty.InUse Then
                    
                    styleName = sty.NameLocal
                    styleTypeName = _
                        Helper_Style_11_GetReadableStyleType(sty)
                    
                    On Error Resume Next
                    Err.Clear
                    
                    sty.Delete
                    
                    If Err.Number = 0 Then
                        
                        deletedCount = deletedCount + 1
                        deletedNames.Add styleName
                        deletedTypes.Add styleTypeName
                        deletedPasses.Add 1
                        
                    End If
                    
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                End If
                
            End If
            
        End If
        
    Next i
    
    '=========================================================================
    ' PASS 2: DEEP VALIDATION OF REMAINING CUSTOM STYLES
    '=========================================================================
    ' Word's .InUse property may remain True because of template/dependency
    ' references even when no document content actually uses the style.
    '
    ' Only those styles require the more expensive structural search.
    For i = doc.Styles.Count To 1 Step -1
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(i)
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
            If Not sty.BuiltIn Then
                
                If sty.InUse Then
                    
                    If Not Helper_Style_11_IsStyleActuallyUsed(doc, sty) Then
                        
                        styleName = sty.NameLocal
                        styleTypeName = _
                            Helper_Style_11_GetReadableStyleType(sty)
                        
                        On Error Resume Next
                        Err.Clear
                        
                        sty.Delete
                        
                        If Err.Number = 0 Then
                            
                            deletedCount = deletedCount + 1
                            deletedNames.Add styleName
                            deletedTypes.Add styleTypeName
                            deletedPasses.Add 2
                            
                        End If
                        
                        Err.Clear
                        On Error GoTo ErrorHandler
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next i
    
    '=========================================================================
    ' PASS 3: FAST DEPENDENCY CLEANUP
    '=========================================================================
    ' Removing custom child styles can make parent custom styles newly unused.
    ' This final pass uses only the fast .InUse check.
    For i = doc.Styles.Count To 1 Step -1
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(i)
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
            If Not sty.BuiltIn Then
                
                If Not sty.InUse Then
                    
                    styleName = sty.NameLocal
                    styleTypeName = _
                        Helper_Style_11_GetReadableStyleType(sty)
                    
                    On Error Resume Next
                    Err.Clear
                    
                    sty.Delete
                    
                    If Err.Number = 0 Then
                        
                        deletedCount = deletedCount + 1
                        deletedNames.Add styleName
                        deletedTypes.Add styleTypeName
                        deletedPasses.Add 3
                        
                    End If
                    
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                End If
                
            End If
            
        End If
        
    Next i
    
    '=========================================================================
    ' COUNT REMAINING STYLES
    '=========================================================================
    For Each sty In doc.Styles
        
        If sty.BuiltIn Then
            builtInCount = builtInCount + 1
        Else
            remainingCustomCount = remainingCustomCount + 1
        End If
        
    Next sty
    
    Application.ScreenUpdating = True
    
    '=========================================================================
    ' FINAL SUMMARY
    '=========================================================================
    reportText = _
        "Style cleanup completed successfully." & vbCrLf & vbCrLf & _
        "Custom styles deleted: " & deletedCount & vbCrLf & _
        "Custom styles remaining: " & remainingCustomCount & vbCrLf & _
        "Built-in styles: " & builtInCount
    
    If deletedCount > 0 Then
        
        reportText = reportText & vbCrLf & vbCrLf & _
                     "Would you like to view the deleted-style details in Excel?"
        
        response = MsgBox( _
            reportText, _
            vbYesNo + vbInformation, _
            "Style Cleanup Summary")
        
        If response = vbYes Then
            
            Helper_Style_11_ExportDeletedStylesToExcel _
                deletedNames, _
                deletedTypes, _
                deletedPasses
            
        End If
        
    Else
        
        MsgBox reportText & vbCrLf & vbCrLf & _
               "No unused custom styles were found.", _
               vbInformation, _
               "Style Cleanup Summary"
        
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Style Cleanup Error"
End Sub


Private Function Helper_Style_11_IsStyleActuallyUsed( _
    ByVal doc As Document, _
    ByVal sty As Style) As Boolean
'=============================================================================
' Parent Sub: Style_11_Del_Unused_Styles_Improved()
'
' Purpose: Determines whether a custom style has at least one actual use in
'          the document.
'
' OPTIMIZATION:
' Stops immediately after finding the first occurrence. It does not count
' every use of the style.
'
' Search Scope:
'   - Main document
'   - Headers / Footers
'   - Footnotes / Endnotes
'   - Accessible text frames and other Word story ranges
'   - Tables, when evaluating a table style
'
' Returns:
'   True  = At least one actual use was found
'   False = No actual document use was found
'
' Safety:
' If Word cannot reliably evaluate an unusual style type, the function returns
' True so the style is retained rather than risking deletion of an active style.
'=============================================================================
    Dim story As Range
    Dim storyRng As Range
    Dim findRng As Range
    Dim tbl As Table
    
    Dim tableStyleName As String
    
    On Error GoTo SafetyExit
    
    '-------------------------------------------------------------------------
    ' TABLE STYLES
    '-------------------------------------------------------------------------
    If sty.Type = wdStyleTypeTable Then
        
        For Each story In doc.StoryRanges
            
            Set storyRng = story
            
            Do While Not storyRng Is Nothing
                
                For Each tbl In storyRng.Tables
                    
                    tableStyleName = ""
                    
                    On Error Resume Next
                    tableStyleName = tbl.Style.NameLocal
                    Err.Clear
                    On Error GoTo SafetyExit
                    
                    If StrComp( _
                        tableStyleName, _
                        sty.NameLocal, _
                        vbTextCompare) = 0 Then
                        
                        Helper_Style_11_IsStyleActuallyUsed = True
                        Exit Function
                        
                    End If
                    
                Next tbl
                
                Set storyRng = storyRng.NextStoryRange
                
            Loop
            
        Next story
        
    '-------------------------------------------------------------------------
    ' LIST STYLES
    '-------------------------------------------------------------------------
    ' List styles are structurally different from ordinary paragraph and
    ' character styles. If Word already reports one as InUse, retain it rather
    ' than risking damage to active numbering structures.
    ElseIf sty.Type = wdStyleTypeList Then
        
        Helper_Style_11_IsStyleActuallyUsed = True
        Exit Function
        
    '-------------------------------------------------------------------------
    ' PARAGRAPH / CHARACTER / LINKED STYLES
    '-------------------------------------------------------------------------
    Else
        
        For Each story In doc.StoryRanges
            
            Set storyRng = story
            
            Do While Not storyRng Is Nothing
                
                Set findRng = storyRng.Duplicate
                
                On Error Resume Next
                Err.Clear
                
                With findRng.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Style = sty
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = True
                End With
                
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo SafetyExit
                    GoTo NextStoryRange
                End If
                
                If findRng.Find.Execute Then
                    
                    Helper_Style_11_IsStyleActuallyUsed = True
                    Exit Function
                    
                End If
                
                On Error GoTo SafetyExit
                
NextStoryRange:
                Set storyRng = storyRng.NextStoryRange
                
            Loop
            
        Next story
        
    End If
    
    Helper_Style_11_IsStyleActuallyUsed = False
    Exit Function

SafetyExit:
    ' Conservative fallback: retain styles Word cannot evaluate safely.
    Helper_Style_11_IsStyleActuallyUsed = True
End Function


Private Function Helper_Style_11_GetReadableStyleType( _
    ByVal sty As Style) As String
'=============================================================================
' Parent Sub: Style_11_Del_Unused_Styles_Improved()
'
' Purpose: Converts Word's WdStyleType value into a readable description for
'          the optional Excel deletion report.
'=============================================================================
    Select Case sty.Type
        
        Case wdStyleTypeParagraph
            Helper_Style_11_GetReadableStyleType = "Paragraph"
            
        Case wdStyleTypeCharacter
            Helper_Style_11_GetReadableStyleType = "Character"
            
        Case wdStyleTypeTable
            Helper_Style_11_GetReadableStyleType = "Table"
            
        Case wdStyleTypeList
            Helper_Style_11_GetReadableStyleType = "List"
            
        Case wdStyleTypeLinked
            Helper_Style_11_GetReadableStyleType = "Linked"
            
        Case Else
            Helper_Style_11_GetReadableStyleType = "Other"
            
    End Select
End Function


Private Sub Helper_Style_11_ExportDeletedStylesToExcel( _
    ByVal deletedNames As Collection, _
    ByVal deletedTypes As Collection, _
    ByVal deletedPasses As Collection)
'=============================================================================
' Parent Sub: Style_11_Del_Unused_Styles_Improved()
'
' Purpose: Creates an Excel workbook containing details of every custom style
'          successfully removed by the parent cleanup macro.
'
'-----------------------------------------------------------------------------
' EXCEL REPORT
'-----------------------------------------------------------------------------
'
' Column              Description
' ---------------------------------------------------------------------------
' Style Name          Deleted custom style name
' Style Type          Paragraph, Character, Table, List, etc.
' Deleted On Pass     Cleanup stage on which deletion succeeded
'
' Pass 1 = Word directly reported the style unused.
' Pass 2 = Word reported InUse, but no actual content use was found.
' Pass 3 = Style became unused after dependency cleanup.
'
' Uses late binding, so no Excel reference must be enabled manually.
'=============================================================================
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    
    Dim i As Long
    Dim lastRow As Long
    
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' CREATE EXCEL REPORT
    '-------------------------------------------------------------------------
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Worksheets(1)
    
    xlApp.Visible = True
    xlWS.Name = "Deleted Styles"
    
    '-------------------------------------------------------------------------
    ' REPORT HEADINGS
    '-------------------------------------------------------------------------
    xlWS.Cells(1, 1).Value = "Style Name"
    xlWS.Cells(1, 2).Value = "Style Type"
    xlWS.Cells(1, 3).Value = "Deleted On Pass"
    
    With xlWS.Range("A1:C1")
        .Font.Bold = True
        .HorizontalAlignment = -4108       ' xlCenter
    End With
    
    '-------------------------------------------------------------------------
    ' WRITE DELETED STYLE INFORMATION
    '-------------------------------------------------------------------------
    For i = 1 To deletedNames.Count
        
        xlWS.Cells(i + 1, 1).Value = deletedNames(i)
        xlWS.Cells(i + 1, 2).Value = deletedTypes(i)
        xlWS.Cells(i + 1, 3).Value = deletedPasses(i)
        
    Next i
    
    lastRow = deletedNames.Count + 1
    
    '-------------------------------------------------------------------------
    ' FORMAT EXCEL REPORT
    '-------------------------------------------------------------------------
    With xlWS
        
        .Columns("A:C").AutoFit
        .Range("A1:C" & lastRow).AutoFilter
        
        With .Range("A1:C" & lastRow).Borders
            .LineStyle = 1
            .Weight = 2
        End With
        
    End With
    
    ' Freeze report heading
    xlWS.Activate
    xlApp.ActiveWindow.SplitRow = 1
    xlApp.ActiveWindow.FreezePanes = True
    
    Exit Sub

ErrorHandler:
    MsgBox "The deleted-style Excel report could not be created." & _
           vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Excel Report Error"
End Sub