'    Option Explicit
  Public title_one_format As Word.style
  Public paragreph_format As Word.ParagraphFormat


 

'    Set paragreph_format = New Word.ParagraphFormat
    

    Public Function set_format(ByVal style As Word.style, ByVal paragreph As Word.ParagraphFormat)
        Set title_one_format = style
        Set paragreph_format = paragreph


    
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.25)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
        .NumberFormat = "%1.%2.%3.%4.%5.%6."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = wdUndefined
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.25)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.75)
        .TabPosition = wdUndefined
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .Strikethrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .Allcaps = wdUndefined
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



    
    With title_one_format
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
        .AutomaticallyUpdate = False
        .BaseStyle = ""                            ' 样式基于
        .NextParagraphStyle = "正文"             ' 后续段落样式
        .LinkToListTemplate ListTemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
        With .Font
            .NameFarEast = "宋体"     ' 字体
            .NameAscii = "宋体"       ' 字体
            .NameOther = "宋体"       ' 字体
            .Name = "+中文正文"       ' 字体
            .Size = 22                ' 字号
            .Bold = True               '       加粗
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
            .Kerning = 二号     ' 字号
'            MsgBox TypeName(.Kerning)
'            MsgBox val("二号")
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
            .OutlineLevel = wdOutlineLevel1     ' 大纲级别
            .CharacterUnitLeftIndent = 0     ' 左缩进量
            .CharacterUnitRightIndent = 0   ' 右缩进量
            .CharacterUnitFirstLineIndent = 0   ' 首行缩进
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
    End Function


