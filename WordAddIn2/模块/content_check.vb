Module content_check
    Public Function Content_check(ByVal para_p As Word.Paragraph, ByVal wd As Word.Application)
        Dim para As Word.Paragraph
        para = para_p
        Dim err_str As String
        Dim return_val As Boolean
        Dim comment As Word.Comments
        Dim text_object As Object
        text_object = create_std_doc.text_object

        Dim t As Object
        t = para_p.Style
        err_str = ""


        '字体
        'If text_object.text_format.Font.NameFarEast <> para_p.Range.Font.NameFarEast Then
        '    err_str = err_str & "字体 设置错误:正确值为:" & text_object.text_format.Font.NameFarEast & ";当前值:" & para_p.Range.Font.NameFarEast & "." & vbCrLf
        'End If
        'If text_object.text_format.Font.NameAscii <> para_p.Range.Font.NameAscii Then
        '    err_str = err_str & "字体 设置错误:正确值为:" & text_object.text_format.Font.NameAscii & ";当前值:" & para_p.Range.Font.NameAscii & "." & vbCrLf
        'End If
        'If text_object.text_format.Font.NameOther <> para_p.Range.Font.NameOther Then
        '    err_str = err_str & "字体 设置错误:正确值为:" & text_object.text_format.Font.NameOther & ";当前值:" & para_p.Range.Font.NameOther & "." & vbCrLf
        'End If
        If text_object.text_format.Font.Name <> para_p.Range.Font.Name Then
            err_str = err_str & "字体 设置错误:正确值为:" & text_object.text_format.Font.Name & ";当前值:" & para_p.Range.Font.Name & "." & vbCrLf
        End If

        If text_object.text_format.Font.Size <> para_p.Range.Font.Size Then
            err_str = err_str & "字号 设置错误:正确值为:" & text_object.text_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
        End If



        '粗体
        If text_object.text_format.Font.Bold <> para_p.Range.Font.Bold Then
            err_str = err_str & "粗体 设置错误:正确值为:" & text_object.text_format.Font.Bold & ";当前值:" & para_p.Range.Font.Bold & "." & vbCrLf
        End If

        '字号
        'If text_object.text_format.Font.Kerning <> para_p.Range.Font.Kerning Then
        '    err_str = err_str & "字号 设置错误:正确值为:" & text_object.text_format.Font.Kerning & ";当前值:" & para_p.Range.Font.Kerning & "." & vbCrLf
        'End If

        REM 段落属性

        '段落左缩进
        If text_object.paragreph_format.LeftIndent <> para_p.Format.LeftIndent Then
            err_str = err_str & "段落左缩进 设置错误:正确值为:" & text_object.paragreph_format.LeftIndent & ";当前值:" & para_p.Format.LeftIndent & "." & vbCrLf
        End If

        '段落右缩进
        If text_object.paragreph_format.RightIndent <> para_p.Format.RightIndent Then
            err_str = err_str & "段落右缩进 设置错误:正确值为:" & text_object.paragreph_format.RightIndent & ";当前值:" & para_p.Format.RightIndent & "." & vbCrLf
        End If

        '段前间距
        If text_object.paragreph_format.SpaceBefore <> para_p.Format.SpaceBefore Then
            err_str = err_str & "段前间距 设置错误:正确值为:" & text_object.paragreph_format.SpaceBefore & ";当前值:" & para_p.Format.SpaceBefore & "." & vbCrLf
        End If
        If text_object.paragreph_format.SpaceBeforeAuto <> para_p.Format.SpaceBeforeAuto Then
            err_str = err_str & "段前间距 设置错误:正确值为:" & text_object.paragreph_format.SpaceBeforeAuto & ";当前值:" & para_p.Format.SpaceBeforeAuto & "." & vbCrLf
        End If

        '段后间距
        If text_object.paragreph_format.SpaceAfter <> para_p.Format.SpaceAfter Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & text_object.paragreph_format.SpaceAfter & ";当前值:" & para_p.Format.SpaceAfter & "." & vbCrLf
        End If
        If text_object.paragreph_format.SpaceAfterAuto <> para_p.Format.SpaceAfterAuto Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & text_object.paragreph_format.SpaceAfterAuto & ";当前值:" & para_p.Format.SpaceAfterAuto & "." & vbCrLf
        End If

        '段后间距
        If text_object.paragreph_format.SpaceAfter <> para_p.Format.SpaceAfter Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & text_object.paragreph_format.SpaceAfter & ";当前值:" & para_p.Format.SpaceAfter & "." & vbCrLf
        End If

        '段落的行距
        If text_object.paragreph_format.LineSpacingRule <> para_p.Format.LineSpacingRule Then
            err_str = err_str & "段落的行距 设置错误:正确值为:" & text_object.paragreph_format.LineSpacingRule & ";当前值:" & para_p.Format.LineSpacingRule & "." & vbCrLf
        End If

        '段落的对齐方式
        If text_object.paragreph_format.Alignment <> para_p.Format.Alignment Then
            err_str = err_str & "段落的对齐方式 设置错误:正确值为:" & text_object.paragreph_format.Alignment & ";当前值:" & para_p.Format.Alignment & "." & vbCrLf
        End If

        '首行缩进的尺寸
        If text_object.paragreph_format.FirstLineIndent <> para_p.Format.FirstLineIndent Then
            err_str = err_str & "首行缩进的尺寸 设置错误:正确值为:" & text_object.paragreph_format.FirstLineIndent & ";当前值:" & para_p.Format.FirstLineIndent & "." & vbCrLf
        End If

        '大纲级别
        'If text_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
        '    err_str = err_str & "大纲级别 设置错误:正确值为:" & text_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
        'End If


        '左缩进量
        If text_object.paragreph_format.CharacterUnitLeftIndent <> para_p.Format.CharacterUnitLeftIndent Then
            err_str = err_str & "左缩进量 设置错误:正确值为:" & text_object.paragreph_format.CharacterUnitLeftIndent & ";当前值:" & para_p.Format.CharacterUnitLeftIndent & "." & vbCrLf
        End If

        '右缩进量
        If text_object.paragreph_format.CharacterUnitRightIndent <> para_p.Format.CharacterUnitRightIndent Then
            err_str = err_str & "右缩进量 设置错误:正确值为:" & text_object.paragreph_format.CharacterUnitRightIndent & ";当前值:" & para_p.Format.CharacterUnitRightIndent & "." & vbCrLf
        End If

        '首行缩进
        If text_object.paragreph_format.CharacterUnitFirstLineIndent <> para_p.Format.CharacterUnitFirstLineIndent Then
            err_str = err_str & "特殊格式首行缩进 设置错误:正确值为:" & text_object.paragreph_format.CharacterUnitFirstLineIndent & ";当前值:" & para_p.Format.CharacterUnitFirstLineIndent & "." & vbCrLf
        End If

        '段前间距
        If text_object.paragreph_format.LineUnitBefore <> para_p.Format.LineUnitBefore Then
            err_str = err_str & "段前间距 设置错误:正确值为:" & text_object.paragreph_format.LineUnitBefore & ";当前值:" & para_p.Format.LineUnitBefore & "." & vbCrLf
        End If

        '段后间距
        If text_object.paragreph_format.LineUnitAfter <> para_p.Format.LineUnitAfter Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & text_object.paragreph_format.LineUnitAfter & ";当前值:" & para_p.Format.LineUnitAfter & "." & vbCrLf
        End If

        '编号特殊字符
        'If text_object.paragreph_format.LineUnitAfter <> para_p.Format.LineUnitAfter Then
        '    err_str = err_str & "段后间距 设置错误:正确值为:" & text_object.paragreph_format.LineUnitAfter & ";当前值:" & para_p.Format.LineUnitAfter & "." & vbCrLf
        'End If
        'para_p.Format.Style.LinkToListTemplate(ListTemplate:=lt)

        '
        'If text_object.paragreph_format.MirrorIndents <> para_p.Format.MirrorIndents Then
        '    err_str = err_str & "MirrorIndents 设置错误:正确值为:" & text_object.paragreph_format.MirrorIndents & ";当前值:" & para_p.Format.MirrorIndents & "." & vbCrLf
        'End If


        ''
        'If text_object.paragreph_format.TextboxTightWrap <> para_p.Format.TextboxTightWrap Then
        '    err_str = err_str & "TextboxTightWrap 设置错误:正确值为:" & text_object.paragreph_format.TextboxTightWrap & ";当前值:" & para_p.Format.TextboxTightWrap & "." & vbCrLf
        'End If


        ''
        'If text_object.paragreph_format.CollapsedByDefault <> para_p.Format.CollapsedByDefault Then
        '    err_str = err_str & "CollapsedByDefault 设置错误:正确值为:" & text_object.paragreph_format.CollapsedByDefault & ";当前值:" & para_p.Format.CollapsedByDefault & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.AutoAdjustRightIndent <> para_p.Format.AutoAdjustRightIndent Then
        '    err_str = err_str & "AutoAdjustRightIndent 设置错误:正确值为:" & text_object.paragreph_format.AutoAdjustRightIndent & ";当前值:" & para_p.Format.AutoAdjustRightIndent & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.DisableLineHeightGrid <> para_p.Format.DisableLineHeightGrid Then
        '    err_str = err_str & "DisableLineHeightGrid 设置错误:正确值为:" & text_object.paragreph_format.DisableLineHeightGrid & ";当前值:" & para_p.Format.DisableLineHeightGrid & "." & vbCrLf
        'End If


        ''
        'If text_object.paragreph_format.FarEastLineBreakControl <> para_p.Format.FarEastLineBreakControl Then
        '    err_str = err_str & "FarEastLineBreakControl 设置错误:正确值为:" & text_object.paragreph_format.FarEastLineBreakControl & ";当前值:" & para_p.Format.FarEastLineBreakControl & "." & vbCrLf
        'End If


        ''
        'If text_object.paragreph_format.WordWrap <> para_p.Format.WordWrap Then
        '    err_str = err_str & "WordWrap 设置错误:正确值为:" & text_object.paragreph_format.WordWrap & ";当前值:" & para_p.Format.WordWrap & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.HangingPunctuation <> para_p.Format.HangingPunctuation Then
        '    err_str = err_str & "HangingPunctuation 设置错误:正确值为:" & text_object.paragreph_format.HangingPunctuation & ";当前值:" & para_p.Format.HangingPunctuation & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.HalfWidthPunctuationOnTopOfLine <> para_p.Format.HalfWidthPunctuationOnTopOfLine Then
        '    err_str = err_str & "HalfWidthPunctuationOnTopOfLine 设置错误:正确值为:" & text_object.paragreph_format.HalfWidthPunctuationOnTopOfLine & ";当前值:" & para_p.Format.HalfWidthPunctuationOnTopOfLine & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.AddSpaceBetweenFarEastAndAlpha <> para_p.Format.AddSpaceBetweenFarEastAndAlpha Then
        '    err_str = err_str & "AddSpaceBetweenFarEastAndAlpha 设置错误:正确值为:" & text_object.paragreph_format.AddSpaceBetweenFarEastAndAlpha & ";当前值:" & para_p.Format.AddSpaceBetweenFarEastAndAlpha & "." & vbCrLf
        'End If


        ''
        'If text_object.paragreph_format.AddSpaceBetweenFarEastAndDigit <> para_p.Format.AddSpaceBetweenFarEastAndDigit Then
        '    err_str = err_str & "AddSpaceBetweenFarEastAndDigit 设置错误:正确值为:" & text_object.paragreph_format.AddSpaceBetweenFarEastAndDigit & ";当前值:" & para_p.Format.AddSpaceBetweenFarEastAndDigit & "." & vbCrLf
        'End If

        ''
        'If text_object.paragreph_format.BaselineAlignment <> para_p.Format.BaseLineAlignment Then
        '    err_str = err_str & "BaselineAlignment 设置错误:正确值为:" & text_object.paragreph_format.BaselineAlignment & ";当前值:" & para_p.Format.BaseLineAlignment & "." & vbCrLf
        'End If




        Dim err_comment As Object
        err_comment = New Err_comment
        comment = wd.ActiveDocument.Comments
        return_val = err_comment.set_comment(para, comment, err_str)

        Return True
    End Function

End Module
