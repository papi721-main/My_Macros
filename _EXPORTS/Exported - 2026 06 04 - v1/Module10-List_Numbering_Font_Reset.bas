Attribute VB_Name = "Module10"
Sub List_Numbering_Font_Reset()
    For Each templ In ActiveDocument.ListTemplates
        For Each lev In templ.ListLevels
            lev.Font.Reset
        Next lev
    Next templ
End Sub
