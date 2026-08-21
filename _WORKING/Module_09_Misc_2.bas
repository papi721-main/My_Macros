Sub Misc_21_List_Document_Styles_And_Usage()
'=============================================================================
' Name: Misc_21_List_Document_Styles_And_Usage()
' Purpose: Creates an Excel workbook listing all styles available in the active
'          Word document, including style type, usage status, and whether each
'          style is built-in or custom.
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim rowNum As Long
    Dim styleType As String
    
    Set doc = ActiveDocument
    
    ' Speed optimization
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' CREATE EXCEL WORKBOOK
    '-------------------------------------------------------------------------
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Worksheets(1)
    
    xlApp.Visible = True
    xlWS.Name = "Style Report"
    
    ' Set report headings
    xlWS.Cells(1, 1).Value = "Style Name"
    xlWS.Cells(1, 2).Value = "Style Type"
    xlWS.Cells(1, 3).Value = "Used?"
    xlWS.Cells(1, 4).Value = "Built-in / Custom"
    
    With xlWS.Range("A1:D1")
        .Font.Bold = True
        .HorizontalAlignment = -4108   ' xlCenter
    End With
    
    '-------------------------------------------------------------------------
    ' SCAN ALL DOCUMENT STYLES
    '-------------------------------------------------------------------------
    rowNum = 2
    
    For Each sty In doc.styles
        
        ' Identify style type
        Select Case sty.Type
            Case wdStyleTypeParagraph
                styleType = "Paragraph"
                
            Case wdStyleTypeCharacter
                styleType = "Character"
                
            Case wdStyleTypeTable
                styleType = "Table"
                
            Case wdStyleTypeList
                styleType = "List"
                
            Case wdStyleTypeLinked
                styleType = "Linked"
                
            Case Else
                styleType = "Other"
        End Select
        
        ' Write style information to Excel
        xlWS.Cells(rowNum, 1).Value = sty.NameLocal
        xlWS.Cells(rowNum, 2).Value = styleType
        
        If sty.InUse Then
            xlWS.Cells(rowNum, 3).Value = "YES"
        Else
            xlWS.Cells(rowNum, 3).Value = "NO"
        End If
        
        If sty.BuiltIn Then
            xlWS.Cells(rowNum, 4).Value = "Built-in"
        Else
            xlWS.Cells(rowNum, 4).Value = "Custom"
        End If
        
        rowNum = rowNum + 1
        
    Next sty
    
    '-------------------------------------------------------------------------
    ' FORMAT EXCEL REPORT
    '-------------------------------------------------------------------------
    With xlWS
        
        ' Auto-fit all report columns
        .Columns("A:D").AutoFit
        
        ' Apply filter to headings
        .Range("A1:D" & rowNum - 1).AutoFilter
        
        ' Add borders around the report
        With .Range("A1:D" & rowNum - 1).Borders
            .LineStyle = 1
            .Weight = 2
        End With
        
    End With
    
    ' Freeze heading row
    xlWS.Activate
    xlApp.ActiveWindow.SplitRow = 1
    xlApp.ActiveWindow.FreezePanes = True
    
CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox doc.styles.count & " styles were exported to Excel successfully.", _
           vbInformation, "Style Report Complete"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Style Report Error"
    
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub

Sub Misc_22_Add_Thousand_Separators_To_Numbers_In_Selection()
'=============================================================================
' Name: Misc_22_Add_Thousand_Separators_To_Numbers_In_Selection()
' Purpose: Adds thousand separators to numbers in the current selection,
'          including selected table cells, while preserving decimal places.
'=============================================================================
    Dim rng As Range
    Dim numRng As Range
    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim txt As String
    Dim ch As String
    Dim numText As String
    Dim cleanNum As String
    Dim formattedNum As String
    Dim decimalPos As Long
    Dim decimalPlaces As Long
    Dim originalEnd As Long
    
    ' Guardrail: Require an actual selection
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "Please select the text or table cells you want to format.", _
               vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Set rng = Selection.Range.Duplicate
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' SCAN SELECTION FROM END TO START
    '-------------------------------------------------------------------------
    ' Working backwards prevents inserted commas from shifting the positions
    ' of numbers that have not yet been processed.
    i = rng.End - 1
    
    Do While i >= rng.Start
        
        Set numRng = ActiveDocument.Range(i, i + 1)
        ch = numRng.text
        
        ' Detect the end of a numeric sequence
        If ch Like "[0-9]" Then
            
            endPos = i + 1
            startPos = i
            
            ' Move backward through digits, commas, and decimal points
            Do While startPos > rng.Start
                
                Set numRng = ActiveDocument.Range(startPos - 1, startPos)
                ch = numRng.text
                
                If ch Like "[0-9]" Or ch = "," Or ch = "." Then
                    startPos = startPos - 1
                Else
                    Exit Do
                End If
                
            Loop
            
            Set numRng = ActiveDocument.Range(startPos, endPos)
            numText = numRng.text
            
            ' Remove existing thousand separators
            cleanNum = Replace(numText, ",", "")
            
            If IsNumeric(cleanNum) Then
                
                ' Determine number of decimal places
                decimalPos = InStr(cleanNum, ".")
                
                If decimalPos > 0 Then
                    
                    decimalPlaces = Len(cleanNum) - decimalPos
                    
                    formattedNum = Format( _
                        CDbl(cleanNum), _
                        "#,##0." & String(decimalPlaces, "0") _
                    )
                    
                Else
                    
                    formattedNum = Format(CDbl(cleanNum), "#,##0")
                    
                End If
                
                ' Replace only when a change is required
                If formattedNum <> numText Then
                    numRng.text = formattedNum
                End If
                
            End If
            
            ' Continue scanning before the number just processed
            i = startPos - 1
            
        Else
            
            i = i - 1
            
        End If
        
    Loop
    
CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox "Thousand separators applied successfully.", _
           vbInformation, "Formatting Complete"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Number Formatting Error"
End Sub

Sub Misc_23_Number_To_Words_Converter()
'=============================================================================
' Name: Misc_23_Number_To_Words_Converter()
' Purpose: Converts the currently selected whole number into English words.
'          Example: 8,450 -> Eight Thousand Four Hundred Fifty
'=============================================================================
    Dim txt As String
    Dim num As Double
    Dim result As String
    Dim groups(0 To 4) As Long
    Dim groupNames(0 To 4) As String
    Dim i As Long
    Dim n As Long
    Dim part As String
    
    Dim ones As Variant
    Dim teens As Variant
    Dim tens As Variant
    
    ' Guardrail: Require selected text
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "Please select a number first.", _
               vbExclamation, "No Selection"
        Exit Sub
    End If
    
    txt = Trim(Selection.text)
    txt = Replace(txt, ",", "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    
    If Not IsNumeric(txt) Then
        MsgBox "The selected text is not a valid number.", _
               vbExclamation, "Invalid Number"
        Exit Sub
    End If
    
    num = CDbl(txt)
    
    ' This version handles whole numbers only
    If num <> Fix(num) Then
        MsgBox "This macro currently supports whole numbers only.", _
               vbExclamation, "Decimal Number"
        Exit Sub
    End If
    
    If num < 0 Or num > 999999999999999# Then
        MsgBox "Please select a number from 0 to 999,999,999,999,999.", _
               vbExclamation, "Number Out of Range"
        Exit Sub
    End If
    
    ' Word lookup arrays
    ones = Array("", "One", "Two", "Three", "Four", "Five", _
                 "Six", "Seven", "Eight", "Nine")
    
    teens = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", _
                  "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", _
                 "Sixty", "Seventy", "Eighty", "Ninety")
    
    groupNames(0) = ""
    groupNames(1) = "Thousand"
    groupNames(2) = "Million"
    groupNames(3) = "Billion"
    groupNames(4) = "Trillion"
    
    ' Special case for zero
    If num = 0 Then
        Selection.text = "Zero"
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------
    ' SPLIT NUMBER INTO GROUPS OF THREE DIGITS
    '-------------------------------------------------------------------------
    For i = 0 To 4
        groups(i) = CLng(num - Int(num / 1000) * 1000)
        num = Int(num / 1000)
    Next i
    
    '-------------------------------------------------------------------------
    ' CONVERT EACH THREE-DIGIT GROUP
    '-------------------------------------------------------------------------
    For i = 4 To 0 Step -1
        
        n = groups(i)
        part = ""
        
        If n > 0 Then
            
            ' Hundreds
            If n >= 100 Then
                part = ones(Int(n / 100)) & " Hundred"
                n = n Mod 100
            End If
            
            ' Tens and ones
            If n >= 20 Then
                
                If part <> "" Then part = part & " "
                
                part = part & tens(Int(n / 10))
                
                If n Mod 10 > 0 Then
                    part = part & " " & ones(n Mod 10)
                End If
                
            ElseIf n >= 10 Then
                
                If part <> "" Then part = part & " "
                
                part = part & teens(n - 10)
                
            ElseIf n > 0 Then
                
                If part <> "" Then part = part & " "
                
                part = part & ones(n)
                
            End If
            
            ' Add Thousand / Million / Billion / Trillion
            If groupNames(i) <> "" Then
                part = part & " " & groupNames(i)
            End If
            
            If result <> "" Then result = result & " "
            result = result & part
            
        End If
        
    Next i
    
    ' Replace selected number with written form
    Selection.text = result
    
    MsgBox "Number converted to words successfully.", _
           vbInformation, "Conversion Complete"
End Sub


Sub Misc_24_Financial_Number_To_Words_Converter()
'=============================================================================
' Name: Misc_24_Financial_Number_To_Words_Converter()
' Purpose: Converts the currently selected numeric amount into formal financial
'          wording using Birr and Cents.
'          Example: 8,450.75 -> Eight Thousand Four Hundred Fifty Birr and
'          Seventy-Five Cents Only
'=============================================================================
    Dim txt As String
    Dim amount As Double
    Dim wholePart As Double
    Dim centsPart As Long
    Dim result As String
    Dim wholeWords As String
    Dim centWords As String
    
    Dim groups(0 To 4) As Long
    Dim groupNames(0 To 4) As String
    Dim ones As Variant
    Dim teens As Variant
    Dim tens As Variant
    
    Dim i As Long
    Dim n As Long
    Dim part As String
    Dim workingNum As Double
    
    ' Guardrail: Require selected text
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "Please select a numeric amount first.", _
               vbExclamation, "No Selection"
        Exit Sub
    End If
    
    txt = Trim(Selection.text)
    
    ' Remove common formatting characters
    txt = Replace(txt, ",", "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    txt = Replace(txt, "ETB", "", , , vbTextCompare)
    txt = Replace(txt, "Birr", "", , , vbTextCompare)
    txt = Trim(txt)
    
    If Not IsNumeric(txt) Then
        MsgBox "The selected text is not a valid numeric amount.", _
               vbExclamation, "Invalid Amount"
        Exit Sub
    End If
    
    amount = CDbl(txt)
    
    If amount < 0 Or amount > 1E+15 Then
        MsgBox "Please select an amount between 0 and 999,999,999,999,999.99.", _
               vbExclamation, "Amount Out of Range"
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------
    ' PREPARE NUMBER LOOKUP VALUES
    '-------------------------------------------------------------------------
    ones = Array("", "One", "Two", "Three", "Four", "Five", _
                 "Six", "Seven", "Eight", "Nine")
    
    teens = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", _
                  "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", _
                 "Sixty", "Seventy", "Eighty", "Ninety")
    
    groupNames(0) = ""
    groupNames(1) = "Thousand"
    groupNames(2) = "Million"
    groupNames(3) = "Billion"
    groupNames(4) = "Trillion"
    
    '-------------------------------------------------------------------------
    ' SEPARATE BIRR AND CENTS
    '-------------------------------------------------------------------------
    wholePart = Fix(amount)
    centsPart = CLng(Round((amount - wholePart) * 100, 0))
    
    ' Correct rounding such as 99.999 -> 100 cents
    If centsPart = 100 Then
        wholePart = wholePart + 1
        centsPart = 0
    End If
    
    '-------------------------------------------------------------------------
    ' CONVERT WHOLE NUMBER TO WORDS
    '-------------------------------------------------------------------------
    If wholePart = 0 Then
        
        wholeWords = "Zero"
        
    Else
        
        workingNum = wholePart
        
        For i = 0 To 4
            groups(i) = CLng(workingNum - Int(workingNum / 1000) * 1000)
            workingNum = Int(workingNum / 1000)
        Next i
        
        For i = 4 To 0 Step -1
            
            n = groups(i)
            part = ""
            
            If n > 0 Then
                
                ' Hundreds
                If n >= 100 Then
                    part = ones(Int(n / 100)) & " Hundred"
                    n = n Mod 100
                End If
                
                ' Tens and ones
                If n >= 20 Then
                    
                    If part <> "" Then part = part & " "
                    
                    part = part & tens(Int(n / 10))
                    
                    If n Mod 10 > 0 Then
                        part = part & "-" & ones(n Mod 10)
                    End If
                    
                ElseIf n >= 10 Then
                    
                    If part <> "" Then part = part & " "
                    part = part & teens(n - 10)
                    
                ElseIf n > 0 Then
                    
                    If part <> "" Then part = part & " "
                    part = part & ones(n)
                    
                End If
                
                If groupNames(i) <> "" Then
                    part = part & " " & groupNames(i)
                End If
                
                If wholeWords <> "" Then wholeWords = wholeWords & " "
                wholeWords = wholeWords & part
                
            End If
            
        Next i
        
    End If
    
    '-------------------------------------------------------------------------
    ' CONVERT CENTS TO WORDS
    '-------------------------------------------------------------------------
    If centsPart > 0 Then
        
        n = centsPart
        
        If n >= 20 Then
            
            centWords = tens(Int(n / 10))
            
            If n Mod 10 > 0 Then
                centWords = centWords & "-" & ones(n Mod 10)
            End If
            
        ElseIf n >= 10 Then
            
            centWords = teens(n - 10)
            
        Else
            
            centWords = ones(n)
            
        End If
        
    End If
    
    '-------------------------------------------------------------------------
    ' BUILD FINAL FINANCIAL WORDING
    '-------------------------------------------------------------------------
    If wholePart = 1 Then
        result = wholeWords & " Birr"
    Else
        result = wholeWords & " Birr"
    End If
    
    If centsPart > 0 Then
        
        If centsPart = 1 Then
            result = result & " and " & centWords & " Cent"
        Else
            result = result & " and " & centWords & " Cents"
        End If
        
    End If
    
    result = result & " Only"
    
    ' Replace selected amount
    Selection.text = result
    
    MsgBox "Amount converted to financial wording successfully.", _
           vbInformation, "Conversion Complete"
End Sub

Sub Misc_25_Highlight_All_Numbers()
'=============================================================================
' Name: Misc_25_Highlight_All_Numbers()
' Purpose: Highlights all numbers in the document using custom violet shading.
'          Highlight Color: #D8B4FE
'=============================================================================
    Dim doc As Document
    Dim story As Range
    Dim rng As Range
    Dim numberColor As Long
    
    Set doc = ActiveDocument
    numberColor = RGB(216, 180, 254)     ' #D8B4FE - Violet
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' FIND AND HIGHLIGHT ALL NUMBERS
    '-------------------------------------------------------------------------
    For Each story In doc.StoryRanges
        
        Set rng = story.Duplicate
        
        Do
            
            With rng.Find
                .ClearFormatting
                .text = "[0-9]@"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
            End With
            
            Do While rng.Find.Execute
                
                With rng.Shading
                    .Texture = wdTextureNone
                    .ForegroundPatternColor = wdColorAutomatic
                    .BackgroundPatternColor = numberColor
                End With
                
                rng.Collapse wdCollapseEnd
                rng.End = story.End
                
            Loop
            
            Set rng = rng.NextStoryRange
            
        Loop While Not rng Is Nothing
        
    Next story
    
CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox "All numbers highlighted with #D8B4FE.", _
           vbInformation, "Number Review Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Number Highlighting Error"
End Sub


Sub Misc_26_Clear_Highlight_From_Numbers()
'=============================================================================
' Name: Misc_26_Clear_Highlight_From_Numbers()
' Purpose: Clears #D8B4FE shading from numbers only, leaving all other
'          document highlighting and shading unchanged.
'=============================================================================
    Dim doc As Document
    Dim story As Range
    Dim rng As Range
    Dim numberColor As Long
    
    Set doc = ActiveDocument
    numberColor = RGB(216, 180, 254)     ' #D8B4FE - Violet
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' FIND NUMBERS AND CLEAR ONLY THE ASSIGNED COLOR
    '-------------------------------------------------------------------------
    For Each story In doc.StoryRanges
        
        Set rng = story.Duplicate
        
        Do
            
            With rng.Find
                .ClearFormatting
                .text = "[0-9]@"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchWildcards = True
            End With
            
            Do While rng.Find.Execute
                
                If rng.Shading.BackgroundPatternColor = numberColor Then
                    
                    With rng.Shading
                        .Texture = wdTextureNone
                        .ForegroundPatternColor = wdColorAutomatic
                        .BackgroundPatternColor = wdColorAutomatic
                    End With
                    
                End If
                
                rng.Collapse wdCollapseEnd
                rng.End = story.End
                
            Loop
            
            Set rng = rng.NextStoryRange
            
        Loop While Not rng Is Nothing
        
    Next story
    
CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox "Number highlighting cleared successfully.", _
           vbInformation, "Number Review Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Number Highlighting Error"
End Sub

Sub Misc_27_Adjust_Styles_To_Obsidian()
'=============================================================================
' Name: Misc_27_Adjust_Styles_To_Obsidian()
' Purpose: Configures Word styles to resemble the Obsidian light-mode
'          typography system while retaining Word-friendly document geometry.
'
'-----------------------------------------------------------------------------
' OBSIDIAN -> MICROSOFT WORD TYPOGRAPHY MAPPING
'-----------------------------------------------------------------------------
'
' Obsidian Element       Word Style          Font / Size           Color
' -------------------------------------------------------------------------------
' Normal Text            Normal / Body*      Calibri / 12 pt       #242424
' Inline Title           Title               Heading Font / 22 pt  #1F2937
' H1                     Heading 1           Heading Font / 20 pt  #B3261E
' H2                     Heading 2           Heading Font / 18 pt  #C2410C
' H3                     Heading 3           Heading Font / 16 pt  #8A5B00
' H4                     Heading 4           Heading Font / 14 pt  #0F766E
' H5                     Heading 5           Heading Font / 13 pt  #6D28D9
' H6                     Heading 6           Heading Font / 12 pt  #4B5563
' Bold                    Strong              Calibri / Bold        #005FCC
' Italic                  Emphasis            Calibri / Italic      #1F7A4D
' Bold + Italic           Intense Emphasis    Calibri / BoldItalic  #7C3AED
' Caption                 Caption             Calibri / 11 pt       #4B5563
' Hyperlink               Hyperlink           Underlined            #005FCC
'
'-----------------------------------------------------------------------------
' CENTRAL TYPOGRAPHY SETTINGS
'-----------------------------------------------------------------------------
' Body Font:        Controlled by bodyFontName
' Heading Font:     Controlled by headingFontName
' Body Size:        12 pt
' Body Spacing:     1.15 lines / 0 pt before / 6 pt after
' Heading Spacing:  Single / 6 pt before / 12 pt after
'=============================================================================
    Dim doc As Document
    Dim sty As Style
    Dim stName As String
    Dim linkStyleNames As Variant
    Dim i As Long
    
    Dim bodyFontName As String
    Dim headingFontName As String
    
    Set doc = ActiveDocument
    
    '-------------------------------------------------------------------------
    ' CENTRAL FONT SETTINGS
    '-------------------------------------------------------------------------
    bodyFontName = "Calibri"
    headingFontName = "Faculty Glyphic"
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' 1. NORMAL & BODY TEXT STYLES
    '-------------------------------------------------------------------------
    ' Dynamically captures Normal, Normal (Web), Normal Indent, Body Text,
    ' Body Text 2, Body Text Indent, and similarly named paragraph styles.
    For Each sty In doc.styles
        
        If sty.Type = wdStyleTypeParagraph Or sty.Type = wdStyleTypeLinked Then
            
            stName = LCase(Trim(sty.NameLocal))
            
            If stName Like "normal*" Or stName Like "body text*" Then
                
                With sty
                    .AutomaticallyUpdate = False
                    
                    With .Font
                        .Name = bodyFontName
                        .Size = 12
                        .Bold = False
                        .Italic = False
                        .Color = RGB(36, 36, 36)           ' #242424
                        .Outline = False
                        .Shadow = False
                        .Emboss = False
                        .Engrave = False
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
                        .FirstLineIndent = 0
                        .OutlineLevel = wdOutlineLevelBodyText
                        .LeftIndent = 0
                        .RightIndent = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfterAuto = False
                        .SpaceBefore = 0
                        .SpaceAfter = 6
                        .LineSpacingRule = wdLineSpaceMultiple
                        .LineSpacing = LinesToPoints(1.15)
                        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
                        .TabStops.ClearAll
                        .Borders.Enable = False
                    End With
                    
                End With
                
            End If
            
        End If
        
    Next sty

    '-------------------------------------------------------------------------
    ' 2. TITLE STYLE
    '-------------------------------------------------------------------------
    On Error Resume Next
    
    With doc.styles("Title")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        
        With .Font
            .Name = headingFontName
            .Size = 22
            .Bold = False
            .Italic = False
            .Color = RGB(31, 41, 55)                    ' #1F2937
            .Spacing = 0
            .Scaling = 100
            .AllCaps = False
        End With
        
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 12
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .Borders.Enable = False
        End With
    End With
    
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' 3. HEADING 1
    '-------------------------------------------------------------------------
    With doc.styles("Heading 1")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 20
            .Bold = True
            .Italic = False
            .AllCaps = False
            .Color = RGB(179, 38, 30)                   ' #B3261E
            .Spacing = 0
            .Scaling = 100
            .Kerning = 0
            .Ligatures = wdLigaturesNone
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel1
            .PageBreakBefore = True
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 4. HEADING 2
    '-------------------------------------------------------------------------
    With doc.styles("Heading 2")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 18
            .Bold = True
            .Italic = False
            .Color = RGB(194, 65, 12)                   ' #C2410C
            .Spacing = 0
            .Scaling = 100
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel2
            .PageBreakBefore = False
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 5. HEADING 3
    '-------------------------------------------------------------------------
    With doc.styles("Heading 3")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 16
            .Bold = True
            .Italic = False
            .Color = RGB(138, 91, 0)                    ' #8A5B00
            .Spacing = 0
            .Scaling = 100
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel3
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 6. HEADING 4
    '-------------------------------------------------------------------------
    With doc.styles("Heading 4")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 14
            .Bold = True
            .Italic = False
            .Color = RGB(15, 118, 110)                  ' #0F766E
            .Spacing = 0
            .Scaling = 100
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel4
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 7. HEADING 5
    '-------------------------------------------------------------------------
    With doc.styles("Heading 5")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 13
            .Bold = True
            .Italic = False
            .Color = RGB(109, 40, 217)                  ' #6D28D9
            .Spacing = 0
            .Scaling = 100
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel5
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 8. HEADING 6
    '-------------------------------------------------------------------------
    With doc.styles("Heading 6")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = True
        
        With .Font
            .Name = headingFontName
            .Size = 12
            .Bold = True
            .Italic = False
            .Color = RGB(75, 85, 99)                    ' #4B5563
            .Spacing = 0
            .Scaling = 100
        End With
        
        With .ParagraphFormat
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 6
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .OutlineLevel = wdOutlineLevel6
            .KeepWithNext = True
            .KeepTogether = True
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 9. EMPHASIS STYLES
    '-------------------------------------------------------------------------
    On Error Resume Next
    
    With doc.styles("Strong").Font
        .Name = bodyFontName
        .Size = 12
        .Bold = True
        .Italic = False
        .Color = RGB(0, 95, 204)                        ' #005FCC
    End With
    
    With doc.styles("Emphasis").Font
        .Name = bodyFontName
        .Size = 12
        .Bold = False
        .Italic = True
        .Color = RGB(31, 122, 77)                       ' #1F7A4D
    End With
    
    With doc.styles("Intense Emphasis").Font
        .Name = bodyFontName
        .Size = 12
        .Bold = True
        .Italic = True
        .Color = RGB(124, 58, 237)                      ' #7C3AED
    End With
    
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' 10. CAPTION STYLE
    '-------------------------------------------------------------------------
    With doc.styles("Caption")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        
        With .Font
            .Name = bodyFontName
            .Size = 11
            .Bold = True
            .Italic = True
            .Color = RGB(75, 85, 99)                    ' #4B5563
            .Spacing = 0
        End With
        
        With .ParagraphFormat
            .SpaceBefore = 6
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(1.15)
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .WidowControl = True
            .OutlineLevel = wdOutlineLevelBodyText
            .TabStops.ClearAll
            .Borders.Enable = False
        End With
    End With

    '-------------------------------------------------------------------------
    ' 11. HYPERLINK STYLES
    '-------------------------------------------------------------------------
    linkStyleNames = Array("Hyperlink", "FollowedHyperlink")
    
    For i = LBound(linkStyleNames) To UBound(linkStyleNames)
        On Error Resume Next
        
        With doc.styles(linkStyleNames(i)).Font
            .Name = bodyFontName
            .Color = RGB(0, 95, 204)                    ' #005FCC
            .Underline = wdUnderlineSingle
        End With
        
        On Error GoTo ErrorHandler
    Next i

CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox "Word styles updated to match the Obsidian typography system.", _
           vbInformation, "Typography Updated"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Style Preferences Error"
    Resume CleanUp
End Sub

Sub Misc_28_Clean_Pandoc_Styles()
'=============================================================================
' Name: Misc_28_Clean_Pandoc_Styles()
' Purpose: Removes Pandoc-generated custom styles while preserving visible
'          formatting, list structure, tables, and mathematical equations.
'
'-----------------------------------------------------------------------------
' CLEANING STRATEGY
'-----------------------------------------------------------------------------
'
' Pandoc Type       Replacement              Preserved
' ---------------------------------------------------------------------------
' Paragraph         Normal                   Font + paragraph formatting
' List Paragraph    Normal                   Formatting + list structure
' Character         Default Paragraph Font   Character formatting
' Math Content      Unchanged                Entire equation is skipped
' Table             Table Normal             Existing cell/paragraph formatting
'
' NOTE:
' Equations are deliberately excluded from style conversion.
' Styles still required by equation content may remain in the document.
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
    Dim stName As Variant
    
    Dim paragraphStyles As Variant
    Dim characterStyles As Variant
    
    Dim currentStyleName As String
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
    
    '-------------------------------------------------------------------------
    ' PANDOC STYLE DEFINITIONS
    '-------------------------------------------------------------------------
    paragraphStyles = Split( _
        "Abstract|Abstract Title|Author|Captioned Figure|Compact|" & _
        "Definition|Definition Term|Figure|First Paragraph|" & _
        "Footnote Block Text|Image Caption|Source Code|Table Caption", "|")
    
    characterStyles = Split( _
        "AlertTok|AnnotationTok|AttributeTok|BaseNTok|BuiltInTok|CharTok|" & _
        "CommentTok|CommentVarTok|ConstantTok|ControlFlowTok|DataTypeTok|" & _
        "DecValTok|DocumentationTok|ErrorTok|ExtensionTok|FloatTok|" & _
        "FunctionTok|ImportTok|InformationTok|KeywordTok|NormalTok|" & _
        "OperatorTok|OtherTok|PreprocessorTok|RegionMarkerTok|" & _
        "Section Number|SpecialCharTok|SpecialStringTok|StringTok|" & _
        "VariableTok|VerbatimStringTok|WarningTok", "|")

    '=========================================================================
    ' PHASE 1: PANDOC PARAGRAPH STYLES -> NORMAL
    '=========================================================================
    For Each stName In paragraphStyles
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(CStr(stName))
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
            For Each story In doc.StoryRanges
                
                Set storyRng = story
                
                Do While Not storyRng Is Nothing
                    
                    For Each para In storyRng.Paragraphs
                        
                        currentStyleName = ""
                        
                        On Error Resume Next
                        currentStyleName = para.Style.NameLocal
                        Err.Clear
                        On Error GoTo ErrorHandler
                        
                        If StrComp(currentStyleName, CStr(stName), _
                                   vbTextCompare) = 0 Then
                            
                            '-------------------------------------------------
                            ' EQUATION GUARDRAIL
                            '-------------------------------------------------
                            If para.Range.OMaths.Count > 0 Then
                                GoTo SkipParagraph
                            End If
                            
                            '-------------------------------------------------
                            ' Preserve list structure.
                            '-------------------------------------------------
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
                            
                            '-------------------------------------------------
                            ' Capture effective font formatting.
                            '-------------------------------------------------
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
                            
                            '-------------------------------------------------
                            ' Capture paragraph formatting.
                            '-------------------------------------------------
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
                            
                            '-------------------------------------------------
                            ' Replace Pandoc style with built-in Normal.
                            ' Using the built-in constant avoids Style
                            ' collection lookup Error 5941.
                            '-------------------------------------------------
                            para.Style = wdStyleNormal
                            
                            '-------------------------------------------------
                            ' Restore list only if Word removed it.
                            '-------------------------------------------------
                            If hasList Then
                                
                                If para.Range.ListFormat.ListType = _
                                   wdListNoNumbering Then
                                    
                                    If Not savedListTemplate Is Nothing Then
                                        
                                        ' A damaged or unusual Pandoc list
                                        ' template must not halt the cleanup.
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
                            
                            '-------------------------------------------------
                            ' Restore paragraph formatting.
                            '-------------------------------------------------
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
                                
                                ' List indentation remains controlled by the
                                ' list definition rather than direct indents.
                                If Not hasList Then
                                    .FirstLineIndent = pFirstIndent
                                    .LeftIndent = pLeftIndent
                                    .RightIndent = pRightIndent
                                End If
                                
                            End With
                            
                            '-------------------------------------------------
                            ' Restore effective character formatting.
                            '-------------------------------------------------
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
            
            '-----------------------------------------------------------------
            ' Attempt deletion. Equation-dependent styles are allowed to stay.
            '-----------------------------------------------------------------
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
            
        End If
        
    Next stName

    '=========================================================================
    ' PHASE 2: PANDOC CHARACTER STYLES
    '=========================================================================
    For Each stName In characterStyles
        
        Set sty = Nothing
        
        On Error Resume Next
        Set sty = doc.Styles(CStr(stName))
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not sty Is Nothing Then
            
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
                        
                        '-----------------------------------------------------
                        ' EQUATION GUARDRAIL
                        '-----------------------------------------------------
                        If searchRng.OMaths.Count > 0 Then
                            
                            searchRng.Collapse wdCollapseEnd
                            searchRng.End = storyEnd
                            GoTo ContinueCharacterSearch
                            
                        End If
                        
                        '-----------------------------------------------------
                        ' Capture character appearance.
                        '-----------------------------------------------------
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
                        
                        '-----------------------------------------------------
                        ' Remove custom character style.
                        '
                        ' Assign the built-in constant directly instead of
                        ' retrieving it from doc.Styles().
                        '-----------------------------------------------------
                        searchRng.Style = wdStyleDefaultParagraphFont
                        
                        '-----------------------------------------------------
                        ' Restore character appearance directly.
                        '-----------------------------------------------------
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
            
            '-----------------------------------------------------------------
            ' Delete if no longer used.
            '-----------------------------------------------------------------
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
            
        End If
        
    Next stName

    '=========================================================================
    ' PHASE 3: PANDOC TABLE STYLE -> TABLE NORMAL
    '=========================================================================
    Set sty = Nothing
    
    On Error Resume Next
    Set sty = doc.Styles("Table")
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not sty Is Nothing Then
        
        For Each tbl In doc.Tables
            
            currentStyleName = ""
            
            On Error Resume Next
            currentStyleName = tbl.Style.NameLocal
            Err.Clear
            On Error GoTo ErrorHandler
            
            If StrComp(currentStyleName, "Table", vbTextCompare) = 0 Then
                
                '-------------------------------------------------------------
                ' Preserve cell paragraph/text formatting.
                '-------------------------------------------------------------
                For Each cel In tbl.Range.Cells
                    
                    For Each cellPara In cel.Range.Paragraphs
                        
                        ' Skip equation-containing paragraphs entirely
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
                
                '-------------------------------------------------------------
                ' Replace custom Pandoc table style using the built-in
                ' constant directly.
                '-------------------------------------------------------------
                On Error Resume Next
                tbl.Style = wdStyleTableNormal
                Err.Clear
                On Error GoTo ErrorHandler
                
            End If
            
        Next tbl
        
        '---------------------------------------------------------------------
        ' Delete Pandoc Table style if it is no longer used.
        '---------------------------------------------------------------------
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
        
    End If

CleanUp:
    Application.ScreenUpdating = True
    
    If stylesRemaining = 0 Then
        
        MsgBox "Pandoc styles cleaned successfully." & vbCrLf & vbCrLf & _
               stylesRemoved & " custom styles removed." & vbCrLf & _
               "Lists, tables, and equations were preserved.", _
               vbInformation, "Pandoc Style Cleanup"
        
    Else
        
        MsgBox "Pandoc style cleanup completed." & vbCrLf & vbCrLf & _
               stylesRemoved & " custom styles removed." & vbCrLf & _
               stylesRemaining & " style(s) remain in use." & vbCrLf & vbCrLf & _
               "Equation-dependent styles are intentionally retained.", _
               vbInformation, "Pandoc Style Cleanup"
        
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Pandoc Style Cleanup Error"
End Sub

Sub Misc_29_Apply_Obsidian_Styles_To_Document()
'=============================================================================
' Name: Misc_29_Apply_Obsidian_Styles_To_Document()
' Purpose: Applies the configured Obsidian-style typography throughout the
'          document while normalizing body text, tables, lists, headings,
'          and figure-caption layout behavior.
'
'-----------------------------------------------------------------------------
' FORMATTING SUMMARY
'-----------------------------------------------------------------------------
'
' Element             Font / Size        Spacing                Other
' -------------------------------------------------------------------------------
' Body Text           Calibri / 12 pt    1.15 / 0 Before / 6 After
' Lists               Calibri / 12 pt    1.15 / Contextual After
' Tables              Calibri / 12 pt    1.15 / 0 Before / 0 After
' Headings 1-6        Defined by         Single / Style-defined
'                     Misc_26
' Figure Captions     Style-defined       Existing spacing       KeepNext=False
'                                                               KeepTogether=False
'
' NOTE:
' Run Misc_26_Adjust_Styles_To_Obsidian() first so Heading 1-6, Normal,
' emphasis, captions, and related styles contain the desired typography.
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
    Dim cleanText As String
    
    Dim bodyFontName As String
    Dim bodyFontSize As Single
    
    Set doc = ActiveDocument
    
    '-------------------------------------------------------------------------
    ' CENTRAL BODY TYPOGRAPHY SETTINGS
    '-------------------------------------------------------------------------
    bodyFontName = "Calibri"
    bodyFontSize = 12
    
    ' Speed optimization
    Application.ScreenUpdating = False
    
    ' Enable error handling trap
    On Error GoTo ErrorHandler

    '-------------------------------------------------------------------------
    ' PHASE 1: GLOBAL BODY FORMATTING OVERRIDE
    '-------------------------------------------------------------------------
    ' Applies the baseline Obsidian-style body typography throughout ALL
    ' sections. Structural heading styles are restored later by outline level.
    For i = 1 To doc.Sections.count
        
        Set rng = doc.Sections(i).Range
        
        With rng
            
            With .Font
                .Name = bodyFontName
                .Size = bodyFontSize
                
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
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 6
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)
            End With
            
        End With
        
    Next i

    '-------------------------------------------------------------------------
    ' PHASE 2: TABLE PROTECTION LOOP
    '-------------------------------------------------------------------------
    ' Restores compact 1.15-line spacing inside every table while retaining
    ' automatic row heights.
    For i = 1 To doc.Sections.count
        
        For Each tbl In doc.Sections(i).Range.Tables
            
            With tbl.Range
                
                With .Font
                    .Name = bodyFontName
                    .Size = bodyFontSize
                End With
                
                With .ParagraphFormat
                    .SpaceBeforeAuto = False
                    .SpaceAfterAuto = False
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = LinesToPoints(1.15)
                End With
                
            End With
            
            ' Protect vertically merged cells from row-height errors
            On Error Resume Next
            tbl.Rows.Height = 0
            tbl.Rows.HeightRule = wdRowHeightAuto
            On Error GoTo ErrorHandler
            
        Next tbl
        
    Next i

    '-------------------------------------------------------------------------
    ' PHASE 3: LISTS, HEADINGS & FIGURE CAPTIONS
    '-------------------------------------------------------------------------
    ' Scans every paragraph in every section.
    For i = 1 To doc.Sections.count
        
        For Each para In doc.Sections(i).Range.Paragraphs
            
            '=================================================================
            ' SUB-PHASE A: CONTEXTUAL LIST SPACING
            '=================================================================
            If para.OutlineLevel = wdOutlineLevelBodyText Then
                
                If para.Range.ListFormat.ListType <> wdListNoNumbering Then
                    
                    ' Keep table lists under the dedicated table formatting
                    If Not para.Range.Information(wdWithInTable) Then
                        
                        '-----------------------------------------------------
                        ' STEP 1: Tighten introductory paragraph above list
                        '-----------------------------------------------------
                        Set prevPara = para.Previous
                        
                        If Not prevPara Is Nothing Then
                            
                            If prevPara.Range.ListFormat.ListType = _
                               wdListNoNumbering Then
                                
                                If Not prevPara.Range.Information(wdWithInTable) _
                                   And prevPara.OutlineLevel = _
                                       wdOutlineLevelBodyText Then
                                    
                                    prevPara.SpaceAfterAuto = False
                                    prevPara.SpaceAfter = 0
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        '-----------------------------------------------------
                        ' STEP 2: Apply base list geometry
                        '-----------------------------------------------------
                        With para
                            .SpaceBeforeAuto = False
                            .SpaceAfterAuto = False
                            .SpaceBefore = 0
                            .LineSpacingRule = wdLineSpaceMultiple
                            .LineSpacing = LinesToPoints(1.15)
                        End With
                        
                        '-----------------------------------------------------
                        ' STEP 3: Determine whether this is final list item
                        '-----------------------------------------------------
                        Set nextPara = para.Next
                        isLastItem = False
                        
                        If nextPara Is Nothing Then
                            
                            isLastItem = True
                            
                        Else
                            
                            ' Next paragraph exits the list
                            If nextPara.Range.ListFormat.ListType = _
                               wdListNoNumbering Then
                                
                                isLastItem = True
                                
                            ' Next paragraph enters a table
                            ElseIf nextPara.Range.Information(wdWithInTable) Then
                                
                                isLastItem = True
                                
                            End If
                            
                            ' Next paragraph is a structural heading
                            If nextPara.OutlineLevel >= 1 And _
                               nextPara.OutlineLevel <= 9 Then
                                
                                isLastItem = True
                                
                            End If
                            
                        End If
                        
                        '-----------------------------------------------------
                        ' STEP 4: Apply calculated after-spacing
                        '-----------------------------------------------------
                        If isLastItem Then
                            para.SpaceAfter = 6
                        Else
                            para.SpaceAfter = 0
                        End If
                        
                    End If
                    
                End If
                
            End If
            
            '=================================================================
            ' SUB-PHASE B: RESTORE HEADING 1-6 STYLES VIA OUTLINE LEVEL
            '=================================================================
            ' Structural headings receive the typography already configured by
            ' Misc_26_Adjust_Styles_To_Obsidian().
            If Not para.Range.Information(wdWithInTable) Then
                
                outLvl = para.OutlineLevel
                
                If outLvl >= 1 And outLvl <= 6 Then
                    
                    With para.Range
                        
                        ' Remove Phase 1's direct Calibri body override
                        .Font.Reset
                        
                        ' Restore the corresponding true Word heading style
                        .Style = doc.styles("Heading " & outLvl)
                        
                    End With
                    
                End If
                
            End If

            '=================================================================
            ' SUB-PHASE C: FIGURE CAPTION LAYOUT CONTROL
            '=================================================================
            ' Retains the original Figure/Fig caption behavior:
            ' KeepWithNext = False and KeepTogether = False.
            cleanText = Trim(para.Range.text)
            
            cleanText = Replace(cleanText, vbCr, "")
            cleanText = Replace(cleanText, vbLf, "")
            cleanText = Trim(cleanText)

            If LCase(cleanText) Like "figure*" Or _
               LCase(cleanText) Like "fig*" Then
                
                para.KeepWithNext = False
                para.KeepTogether = False
                
            End If
            
        Next para
        
    Next i

CleanUp:
    Application.ScreenUpdating = True
    
    MsgBox "Obsidian typography applied successfully across all sections." & _
           vbCrLf & _
           "Tables, lists, headings, and figure captions were normalized.", _
           vbInformation, "Process Complete"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Formatting Error"
End Sub

Sub Misc_30_Format_Selected_Tables()
'=============================================================================
' Name: Misc_30_Format_Selected_Tables()
' Purpose: Formats all tables contained within the current selection.
'          Applies AutoFit to contents, enables full borders, formats and
'          shades the header row, repeats it across pages, prevents row splits,
'          and aligns all cell contents to the top.
'
'-----------------------------------------------------------------------------
' TABLE FORMATTING APPLIED
'-----------------------------------------------------------------------------
'
' Setting                    Applied Value
' ---------------------------------------------------------------------------
' Column Width               AutoFit to Contents
' Outer Borders              Single Line
' Inner Horizontal Borders   Single Line
' Inner Vertical Borders     Single Line
' First Row                  Bold
' First Row Shading          #DDD9C3
' First Row Repeat           Enabled
' Row Splitting              Disabled
' Vertical Alignment         Top (0)
'
' NOTE:
' The macro works when one or more tables are selected, or when the cursor is
' positioned inside a single table.
'=============================================================================
    Dim tbl As Table
    Dim cel As Cell
    Dim headerColor As Long
    
    '-------------------------------------------------------------------------
    ' HEADER COLOR
    '-------------------------------------------------------------------------
    headerColor = RGB(221, 217, 195)      ' #DDD9C3
    
    '-------------------------------------------------------------------------
    ' GUARDRAIL: REQUIRE AT LEAST ONE TABLE
    '-------------------------------------------------------------------------
    If Selection.Tables.Count = 0 Then
        MsgBox "No table found in the current selection. Please select text " & _
               "containing a table or click inside one.", _
               vbExclamation, "No Table Selected"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    '-------------------------------------------------------------------------
    ' FORMAT ALL TABLES WITHIN THE CURRENT SELECTION
    '-------------------------------------------------------------------------
    For Each tbl In Selection.Tables
        
        With tbl
            
            ' Auto-fit columns based on cell contents
            .AutoFitBehavior wdAutoFitContent
            
            ' Prevent rows from splitting across pages
            .Rows.AllowBreakAcrossPages = False
            
            ' Enable and standardize all table borders
            .Borders.Enable = True
            
            .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
            .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
            
            '-------------------------------------------------------------
            ' VERTICAL ALIGNMENT
            '-------------------------------------------------------------
            ' Align all cell contents to the top for a compact layout.
            For Each cel In .Range.Cells
                cel.VerticalAlignment = 0
            Next cel
            
            '-------------------------------------------------------------
            ' HEADER ROW
            '-------------------------------------------------------------
            With .Rows(1)
                
                ' Repeat header row on subsequent pages
                .HeadingFormat = True
                
                ' Bold header text
                .Range.Font.Bold = True
                
                ' Apply custom header shading
                .Shading.Texture = wdTextureNone
                .Shading.ForegroundPatternColor = wdColorAutomatic
                .Shading.BackgroundPatternColor = headerColor
                
            End With
            
        End With
        
    Next tbl

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Table Formatting Error"
End Sub