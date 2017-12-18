Public Class Title_three
    '    Option Explicit
    Public title_one_format As Word.Style
    Public paragreph_format As Word.ParagraphFormat




    '    Set paragreph_format = New Word.ParagraphFormat


    Public Function Set_format(ByVal style As Word.Style, ByVal paragreph As Word.ParagraphFormat, ByVal wd As Word.Application)
        title_one_format = style
        paragreph_format = paragreph





        With title_one_format
            'ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
            .AutomaticallyUpdate = False
            .BaseStyle = ""                            ' 样式基于
            .NextParagraphStyle = "正文"             ' 后续段落样式
            .LinkToListTemplate(ListTemplate:=lt)
            With .Font
                .NameFarEast = "宋体"     ' 字体
                .NameAscii = "宋体"       ' 字体
                .NameOther = "宋体"       ' 字体
                .Name = "宋体"       ' 字体
                .Size = 16                ' 字号
                .Bold = True               '       加粗
                .Italic = False
                .Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
                .UnderlineColor = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                .StrikeThrough = False
                .DoubleStrikeThrough = False
                .Outline = False
                .Emboss = False
                .Shadow = False
                .Hidden = False
                .SmallCaps = False
                .AllCaps = False
                .Color = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
                .Engrave = False
                .Superscript = False
                .Subscript = False
                .Spacing = 0
                .Scaling = 100
                .Position = 0
                .Kerning = 16    ' 字号
                .Animation = Microsoft.Office.Interop.Word.WdAnimation.wdAnimationNone
                .DisableCharacterSpaceGrid = False
                .EmphasisMark = Microsoft.Office.Interop.Word.WdEmphasisMark.wdEmphasisMarkNone
                .Ligatures = Microsoft.Office.Interop.Word.WdLigatures.wdLigaturesNone
                .NumberSpacing = Microsoft.Office.Interop.Word.WdNumberSpacing.wdNumberSpacingDefault
                .NumberForm = Microsoft.Office.Interop.Word.WdNumberForm.wdNumberFormDefault
                .StylisticSet = Microsoft.Office.Interop.Word.WdStylisticSet.wdStylisticSetDefault
                .ContextualAlternates = 0
            End With
        End With



        With paragreph_format
            .LeftIndent = wd.CentimetersToPoints(0)  ' 缩进
            .RightIndent = wd.CentimetersToPoints(0)  '缩进
            .SpaceBefore = 0             ' 段前间距
            .SpaceBeforeAuto = False
            .SpaceAfter = 0              '段后间距
            .SpaceAfterAuto = False
            .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle   ' 段落的行距
            .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft         ' 段落的对齐方式
            .WidowControl = False
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = wd.CentimetersToPoints(0)   ' 首行缩进的尺寸
            .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3     ' 大纲级别
            .CharacterUnitLeftIndent = 0     ' 左缩进量
            .CharacterUnitRightIndent = 0   ' 右缩进量
            .CharacterUnitFirstLineIndent = 0   ' 首行缩进
            .LineUnitBefore = 0          ' 段前间距
            .LineUnitAfter = 0           ' 段前间距
            .MirrorIndents = False
            .TextboxTightWrap = Microsoft.Office.Interop.Word.WdTextboxTightWrap.wdTightNone
            .CollapsedByDefault = False
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = Microsoft.Office.Interop.Word.WdBaselineAlignment.wdBaselineAlignAuto
        End With
        Return True

    End Function




End Class
