Attribute VB_Name = "Module3"
Sub DelUnusedStyles()
'Delete all unused styles, except for built-in styles,
'in the current document.
'You can't delete built-in styles.
Dim s As Style
For Each s In ActiveDocument.Styles
    'Only execute With if current s isn't a built-in style.
    If s.BuiltIn = False Then
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Style = s.NameLocal
            .Execute FindText:="", Format:=True
            If .Found = False Then s.Delete
        End With
    End If
Next
End Sub
