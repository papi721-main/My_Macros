Attribute VB_Name = "Module9"
Sub DeleteUnusedStyles()
    Dim oStyle As Style

    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
            With ActiveDocument.Content.Find
                .ClearFormatting
                .Style = oStyle.NameLocal
                .Execute FindText:="", Format:=True
                If .Found = False Then oStyle.Delete
            End With
        End If
    Next oStyle
End Sub
