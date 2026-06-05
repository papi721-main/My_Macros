Attribute VB_Name = "Module11"
Sub Delete_Section_Breaks()
    Dim objDoc As Document
    Dim rng As Range

    Application.ScreenUpdating = False
    Set objDoc = ActiveDocument

    ' Create a range covering the entire document
    Set rng = objDoc.Content

    ' Use Find and Replace to delete section breaks
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^b"               ' Section break code
        .Replacement.Text = ""     ' Replace with nothing
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True

    MsgBox "All section breaks have been deleted.", vbInformation
End Sub

