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
