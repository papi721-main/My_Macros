Sub Format_Selected_Tables()
'=============================================================================
' Name: Format_Selected_Tables
' Purpose: Formats all tables residing within the active text selection block.
'          Applies AutoFit to contents, side borders (left/right), bolds
'          the header row, and sets header row repeating across pages.
'=============================================================================
    Dim tbl As Table
    
    ' GUARDRAIL: Check if the current selection block contains any tables.
    ' Selection.Tables.Count works whether you highlighted text across tables
    ' or simply clicked your cursor inside a single table.
    If Selection.Tables.count = 0 Then
        MsgBox "No table found in the current selection. Please select text containing a table or click inside one.", _
               vbExclamation, "No Table Selected"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Loop strictly through the tables captured within the active selection scope
    For Each tbl In Selection.Tables
        With tbl
            ' 1. Set AutoFit behavior to adjust column widths based on cell contents
            .AutoFitBehavior wdAutoFitContent
            
            ' 2. Apply full side borders (Left and Right perimeter gridlines)
            .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
            .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .Borders(wdBorderInsideVertical).LineStyle = wdLineStyleSingle
            .Borders(wdBorderInsideHorizontal).LineStyle = wdLineStyleSingle
            
            ' 3. Configure Header Row formatting metrics
            With .Rows(1)
                .HeadingFormat = True          ' Repeat header row across page breaks
                .Range.Font.Bold = True        ' Bold header text
            End With
        End With
    Next tbl
    
    Application.ScreenUpdating = True
End Sub

