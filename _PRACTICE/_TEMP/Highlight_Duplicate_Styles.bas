Sub Highlight_Duplicate_Styles()
    Dim sty As Style
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    ' Clear existing highlighting from the whole document first
    doc.Content.HighlightColorIndex = wdNoHighlight
    
    For Each sty In doc.Styles
        ' Only look at custom styles (BuiltIn = False)
        ' And specifically target names like those in image_5bfe2b.png
        If Not sty.BuiltIn Then
            
            ' Check if the name contains "Char" or "Indent" or other common dupes
            If InStr(1, sty.NameLocal, " Char", vbTextCompare) > 0 Or _
               InStr(1, sty.NameLocal, "Indent", vbTextCompare) > 0 Then
               
                ' Find where this style is used
                With doc.Content.Find
                    .ClearFormatting
                    .Style = sty.NameLocal
                    .Replacement.Highlight = True
                    .Replacement.Text = "^&" ' Keep the same text
                    
                    ' Execute a "Replace All" with Highlight
                    .Execute Replace:=wdReplaceAll
                End With
                
                Debug.Print "Highlighted usage of: " & sty.NameLocal
            End If
        End If
    Next sty
    
    Application.ScreenUpdating = True
    MsgBox "Check for Bright Green highlights. These are your duplicate 'Char' styles.", vbInformation
End Sub