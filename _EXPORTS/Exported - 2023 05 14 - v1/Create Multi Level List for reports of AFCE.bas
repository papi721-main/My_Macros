Attribute VB_Name = "NewMacros"
Sub Create_Mutli_Level_List()
Attribute Create_Mutli_Level_List.VB_Description = "A multi level list for reports in AFCE"
Attribute Create_Mutli_Level_List.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Create_Mutli_Level_List"
'
' Create_Mutli_Level_List Macro
' A multi level list for reports in AFCE
'
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.25)
        .TabPosition = InchesToPoints(0.25)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUnderlineNone
            .Color = wdColorAutomatic
            .Size = 14
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Lucida Sans"
        End With
        .LinkedStyle = "_H1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = InchesToPoints(0.5)
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUnderlineNone
            .Color = wdColorAutomatic
            .Size = 12
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Lucida Sans"
        End With
        .LinkedStyle = "_H2"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.75)
        .TabPosition = InchesToPoints(0.75)
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUnderlineNone
            .Color = wdColorAutomatic
            .Size = 12
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Lucida Sans"
        End With
        .LinkedStyle = "_H3"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1)
        .TabPosition = InchesToPoints(1)
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUnderlineNone
            .Color = wdColorAutomatic
            .Size = 12
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Lucida Sans"
        End With
        .LinkedStyle = "_H4"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "(%5)"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = InchesToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1.25)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "(%6)"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseRoman
        .NumberPosition = InchesToPoints(1.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%7."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(1.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(1.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%8."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = InchesToPoints(1.75)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(2)
        .TabPosition = InchesToPoints(1.75)
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%9."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleLowercaseRoman
        .NumberPosition = InchesToPoints(2)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(2.25)
        .TabPosition = wdUndefined
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
End Sub
Sub Create_Headings()
Attribute Create_Headings.VB_Description = "Create Headings for reports of AFCE"
Attribute Create_Headings.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Create_Headings"
'
' Create_Headings Macro
' Create Headings for reports of AFCE
'
    ActiveDocument.Styles.Add Name:="_H0", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("_H0").AutomaticallyUpdate = False
    With ActiveDocument.Styles("_H0").Font
        .Name = "Lucida Sans"
        .Size = 14
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("_H0").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 12
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = True
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H0").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("_H0").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("_H0").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("_H0").Frame.Delete
    ActiveDocument.Styles.Add Name:="_H1", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("_H1").AutomaticallyUpdate = False
    With ActiveDocument.Styles("_H1").Font
        .Name = "Lucida Sans"
        .Size = 14
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("_H1").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 12
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = True
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevel1
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H1").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("_H1").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("_H1").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("_H1").Frame.Delete
    With ActiveDocument.Styles("_H0").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 12
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = True
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevel1
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H0").NoSpaceBetweenParagraphsOfSameStyle = False
    With ActiveDocument.Styles("_H0")
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "Normal"
    End With
    ActiveDocument.Styles.Add Name:="_H2", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("_H2").AutomaticallyUpdate = False
    With ActiveDocument.Styles("_H2").Font
        .Name = "Lucida Sans"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("_H2").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 6
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevel2
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H2").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("_H2").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("_H2").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("_H2").Frame.Delete
    ActiveDocument.Styles.Add Name:="_H3", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("_H3").AutomaticallyUpdate = False
    With ActiveDocument.Styles("_H3").Font
        .Name = "Lucida Sans"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("_H3").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 3
        .SpaceBeforeAuto = False
        .SpaceAfter = 3
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevel3
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H3").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("_H3").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("_H3").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("_H3").Frame.Delete
    ActiveDocument.Styles.Add Name:="_H4", Type:=wdStyleTypeParagraph
    ActiveDocument.Styles("_H4").AutomaticallyUpdate = False
    With ActiveDocument.Styles("_H4").Font
        .Name = "Lucida Sans"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("_H4").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .OutlineLevel = wdOutlineLevel4
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
    ActiveDocument.Styles("_H4").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("_H4").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("_H4").ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    ActiveDocument.Styles("_H4").Frame.Delete
End Sub
