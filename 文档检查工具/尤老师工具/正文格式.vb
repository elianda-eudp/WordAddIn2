    With text_format
        With .Font
            .NameFarEast = "宋体"     ' 字体
            .NameAscii = "宋体"       ' 字体
            .NameOther = "宋体"       ' 字体
            .Name = "+中文正文"       ' 字体
            .Size = 10.5                ' 字号
            .Bold = False               '       加粗
            .Italic = False
            .Underline = wdUnderlineNone
            .UnderlineColor = wdColorAutomatic
            .Strikethrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .Smallcaps = False
            .Allcaps = False
            .Color = wdColorAutomatic
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Spacing = 0
            .Scaling = 100
            .Position = 0
            .Kerning = 五号     ' 字号
            .Animation = wdAnimationNone
            .DisableCharacterSpaceGrid = False
            .EmphasisMark = wdEmphasisMarkNone
            .Ligatures = wdLigaturesNone
            .NumberSpacing = wdNumberSpacingDefault
            .NumberForm = wdNumberFormDefault
            .StylisticSet = wdStylisticSetDefault
            .ContextualAlternates = 0
        End With
    End With
        
        
        
        With paragreph_format
            .LeftIndent = CentimetersToPoints(0)  ' 缩进
            .RightIndent = CentimetersToPoints(0)  '缩进
            .SpaceBefore = 0             ' 段前间距
            .SpaceBeforeAuto = False
            .SpaceAfter = 0              '段后间距
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle   ' 段落的行距
            .Alignment = wdAlignParagraphLeft         ' 段落的对齐方式
            .WidowControl = False
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)   ' 首行缩进的尺寸
            .OutlineLevel = wdOutLineLevelBodyText     ' 大纲级别
            .CharacterUnitLeftIndent = 0     ' 左缩进量
            .CharacterUnitRightIndent = 0   ' 右缩进量
            .CharacterUnitFirstLineIndent = 2   ' 首行缩进
            .LineUnitBefore = 0          ' 段前间距
            .LineUnitAfter = 0           ' 段前间距
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaselineAlignment = wdBaselineAlignAuto
        End With
        
        
