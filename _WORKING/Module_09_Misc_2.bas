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

