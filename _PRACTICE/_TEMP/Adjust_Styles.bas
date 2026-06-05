Sub Adjust_Styles()
'
' Adjust_Styles Macro
'
'
    Dim doc As Document
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' --- NORMAL (Body Text) ---
    With doc.Styles("Normal")
        .AutomaticallyUpdate = False
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 1 ---
    With doc.Styles("Heading 1,_H1")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 18
            .Bold = True
            .Italic = False
            .AllCaps = True
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 12
            .SpaceAfter = 12
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .OutlineLevel = wdOutlineLevel1
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 2 ---
    With doc.Styles("Heading 2,_H2")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 16
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel2
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 3 ---
    With doc.Styles("Heading 3,_H3")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 14
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel3
            .TabStops.ClearAll
        End With
    End With

    ' --- HEADING 4 ---
    With doc.Styles("Heading 4,_H4")
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .AutomaticallyUpdate = False
        .NoSpaceBetweenParagraphsOfSameStyle = False
        With .Font
            .Name = "Calibri"
            .Size = 12
            .Bold = True
            .Italic = False
            .Color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 6
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphLeft
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel4
            .TabStops.ClearAll
        End With
    End With

CleanUp:
    Application.ScreenUpdating = True
    MsgBox "Styles successfully updated", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Style Preferences Error"
    Resume CleanUp
End Sub