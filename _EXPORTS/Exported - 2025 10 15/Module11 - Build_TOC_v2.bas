Attribute VB_Name = "NewMacros"
Sub Build_TOC_v2()
Attribute Build_TOC_v2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Build_TOC_v2"
'
' Build_TOC_v2 Macro
'
'
    With ActiveDocument.Styles("TOC 1")
        .AutomaticallyUpdate = True
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
    With ActiveDocument.Styles("TOC 1").Font
        .Name = "Calibri"
        .Size = 11
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
    With ActiveDocument.Styles("TOC 1").ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0.5)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
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
    ActiveDocument.Styles("TOC 1").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("TOC 1").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("TOC 2")
        .AutomaticallyUpdate = True
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
    With ActiveDocument.Styles("TOC 2").Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
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
    With ActiveDocument.Styles("TOC 2").ParagraphFormat
        .LeftIndent = InchesToPoints(0.2)
        .RightIndent = InchesToPoints(0.5)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
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
    ActiveDocument.Styles("TOC 2").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("TOC 2").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument.Styles("TOC 3")
        .AutomaticallyUpdate = True
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
    With ActiveDocument.Styles("TOC 3").ParagraphFormat
        .LeftIndent = InchesToPoints(0.4)
        .RightIndent = InchesToPoints(0.5)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
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
    ActiveDocument.Styles("TOC 3").NoSpaceBetweenParagraphsOfSameStyle = False
    ActiveDocument.Styles("TOC 3").ParagraphFormat.TabStops.ClearAll
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            True
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
End Sub
