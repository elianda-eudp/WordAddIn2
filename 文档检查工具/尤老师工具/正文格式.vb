    With text_format
        With .Font
            .NameFarEast = "����"     ' ����
            .NameAscii = "����"       ' ����
            .NameOther = "����"       ' ����
            .Name = "+��������"       ' ����
            .Size = 10.5                ' �ֺ�
            .Bold = False               '       �Ӵ�
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
            .Kerning = ���     ' �ֺ�
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
            .LeftIndent = CentimetersToPoints(0)  ' ����
            .RightIndent = CentimetersToPoints(0)  '����
            .SpaceBefore = 0             ' ��ǰ���
            .SpaceBeforeAuto = False
            .SpaceAfter = 0              '�κ���
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle   ' ������о�
            .Alignment = wdAlignParagraphLeft         ' ����Ķ��뷽ʽ
            .WidowControl = False
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)   ' ���������ĳߴ�
            .OutlineLevel = wdOutLineLevelBodyText     ' ��ټ���
            .CharacterUnitLeftIndent = 0     ' ��������
            .CharacterUnitRightIndent = 0   ' ��������
            .CharacterUnitFirstLineIndent = 2   ' ��������
            .LineUnitBefore = 0          ' ��ǰ���
            .LineUnitAfter = 0           ' ��ǰ���
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
        
        
