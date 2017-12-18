Module title_check
    Public Function Title_check(ByVal para_p As Word.Paragraph, ByVal wd As Word.Application)
        Dim para As Word.Paragraph
        para = para_p
        Dim err_str As String
        Dim return_val As Boolean
        Dim comment As Word.Comments
        title_one_object = create_std_doc.title_one_object
        title_two_object = create_std_doc.title_two_object
        title_three_object = create_std_doc.title_three_object
        title_four_object = create_std_doc.title_four_object
        title_five_object = create_std_doc.title_five_object
        title_six_object = create_std_doc.title_six_object
        title_seven_object = create_std_doc.title_seven_object
        title_eight_object = create_std_doc.title_eight_object
        title_nine_object = create_std_doc.title_nine_object


        Dim t As Object
        t = para_p.Style.ToString
        err_str = ""
        'MsgBox para_p.style.Font.Size
        If t = "标题 1" Or para_p.Format.OutlineLevel = title_one_object.paragreph_format.OutlineLevel Then
            If title_one_object.title_one_format.Font.Size <> para_p.Range.Font.Size Then
                err_str = err_str & "字号 设置错误:正确值为:" & title_one_object.title_one_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
            End If

            '大纲级别
            '        If title_one_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
            '            err_str = err_str &   "大纲级别 设置错误:正确值为:" & title_one_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
            '        End If
        ElseIf t = "标题 2" Or title_two_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            If title_two_object.title_one_format.Font.Size <> para_p.Range.Font.Size Then
                err_str = err_str & "字号 设置错误:正确值为:" & title_two_object.title_one_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
            End If

            '大纲级别
            '        If title_two_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
            '            err_str = err_str &   "大纲级别 设置错误:正确值为:" & title_two_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
            '        End If
        ElseIf t = "标题 3" Or title_three_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            If title_three_object.title_one_format.Font.Size <> para_p.Range.Font.Size Then
                err_str = err_str & "字号 设置错误:正确值为:" & title_three_object.title_one_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
            End If

            '大纲级别
            '        If title_three_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
            '            err_str = err_str &   "大纲级别 设置错误:正确值为:" & title_three_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
            '        End If
        ElseIf t = "标题 4" Or title_four_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            If title_four_object.title_one_format.Font.Size <> para_p.Range.Font.Size Then
                err_str = err_str & "字号 设置错误:正确值为:" & title_four_object.title_one_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
            End If

            '大纲级别
            '        If title_four_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
            '            err_str = err_str &   "大纲级别 设置错误:正确值为:" & title_four_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
            '        End If
        ElseIf t = "标题 5" Or title_five_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            If title_five_object.title_one_format.Font.Size <> para_p.Range.Font.Size Then
                err_str = err_str & "字号 设置错误:正确值为:" & title_five_object.title_one_format.Font.Size & ";当前值:" & para_p.Range.Font.Size & "." & vbCrLf
            End If

            '大纲级别
            '        If title_five_object.paragreph_format.OutlineLevel <> para_p.Format.OutlineLevel Then
            '            err_str = err_str &   "大纲级别 设置错误:正确值为:" & title_five_object.paragreph_format.OutlineLevel & ";当前值:" & para_p.Format.OutlineLevel & "." & vbCrLf
            '        End If
        End If
        'MsgBox err_str & "," & para_p.Range.text
        'MsgBox para_p.Range.Font.Size

        'AutomaticallyUpdate
        If title_one_object.title_one_format.AutomaticallyUpdate <> para_p.Style.AutomaticallyUpdate Then
            err_str = err_str & "自动重新定义此样式 设置错误:正确值为:" & title_one_object.title_one_format.AutomaticallyUpdate & ";当前值:" & para_p.Style.AutomaticallyUpdate & "." & vbCrLf
        End If

        '样式基于
        If title_one_object.title_one_format.BaseStyle.ToString <> para_p.Style.BaseStyle.ToString Then
            err_str = err_str & "样式基于 设置错误:正确值为:" & title_one_object.title_one_format.BaseStyle.ToString & ";当前值:" & para_p.Style.BaseStyle.ToString & "." & vbCrLf
        End If

        'MsgBox WdBuiltinStyle.wdStyleHeading7
        '后续段落样式
        If title_one_object.title_one_format.NextParagraphStyle.ToString <> para_p.Style.NextParagraphStyle.ToString Then
            err_str = err_str & "后续段落样式 设置错误:正确值为:" & title_one_object.title_one_format.NextParagraphStyle.ToString & ";当前值:" & para_p.Style.NextParagraphStyle.ToString & "." & vbCrLf
        End If

        '字体
        'If title_one_object.title_one_format.Font.NameFarEast <> para_p.Range.Font.NameFarEast Then
        '    err_str = err_str & "NameFarEast字体 设置错误:正确值为:" & title_one_object.title_one_format.Font.NameFarEast & ";当前值:" & para_p.Range.Font.NameFarEast & "." & vbCrLf
        'End If
        'If title_one_object.title_one_format.Font.NameAscii <> para_p.Range.Font.NameAscii Then
        '    err_str = err_str & "NameAscii字体 设置错误:正确值为:" & title_one_object.title_one_format.Font.NameAscii & ";当前值:" & para_p.Range.Font.NameAscii & "." & vbCrLf
        'End If
        'If title_one_object.title_one_format.Font.NameOther <> para_p.Range.Font.NameOther Then
        '    err_str = err_str & "NameOther字体 设置错误:正确值为:" & title_one_object.title_one_format.Font.NameOther & ";当前值:" & para_p.Range.Font.NameOther & "." & vbCrLf
        'End If
        If title_one_object.title_one_format.Font.Name <> para_p.Range.Font.Name Then
            err_str = err_str & "字体 设置错误:正确值为:" & title_one_object.title_one_format.Font.Name & ";当前值:" & para_p.Range.Font.Name & "." & vbCrLf
        End If


        '粗体
        If title_one_object.title_one_format.Font.Bold <> para_p.Range.Font.Bold Then
            err_str = err_str & "粗体 设置错误:正确值为:" & title_one_object.title_one_format.Font.Bold & ";当前值:" & para_p.Range.Font.Bold & "." & vbCrLf
        End If

        '字号
        '    If title_one_object.title_one_format.Font.Kerning <> para_p.Range.Font.Kerning Then
        '        err_str = err_str & "Kerning字号 设置错误:正确值为:" & title_one_object.title_one_format.Font.Kerning & ";当前值:" & para_p.Range.Font.Kerning & "." & vbCrLf
        '    End If

        '段落左缩进
        If title_one_object.paragreph_format.LeftIndent <> para_p.Format.LeftIndent Then
            err_str = err_str & "段落左缩进 设置错误:正确值为:" & title_one_object.paragreph_format.LeftIndent & ";当前值:" & para_p.Format.LeftIndent & "." & vbCrLf
        End If

        '段落右缩进
        If title_one_object.paragreph_format.RightIndent <> para_p.Format.RightIndent Then
            err_str = err_str & "段落右缩进 设置错误:正确值为:" & title_one_object.paragreph_format.RightIndent & ";当前值:" & para_p.Format.RightIndent & "." & vbCrLf
        End If

        '段前间距
        If title_one_object.paragreph_format.SpaceBefore <> para_p.Format.SpaceBefore Then
            err_str = err_str & "段前间距 设置错误:正确值为:" & title_one_object.paragreph_format.SpaceBefore & ";当前值:" & para_p.Format.SpaceBefore & "." & vbCrLf
        End If
        If title_one_object.paragreph_format.SpaceBeforeAuto <> para_p.Format.SpaceBeforeAuto Then
            err_str = err_str & "自动设置指定段落的段前间距 设置错误:正确值为:" & title_one_object.paragreph_format.SpaceBeforeAuto & ";当前值:" & para_p.Format.SpaceBeforeAuto & "." & vbCrLf
        End If

        '段后间距
        If title_one_object.paragreph_format.SpaceAfter <> para_p.Format.SpaceAfter Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & title_one_object.paragreph_format.SpaceAfter & ";当前值:" & para_p.Format.SpaceAfter & "." & vbCrLf
        End If
        If title_one_object.paragreph_format.SpaceAfterAuto <> para_p.Format.SpaceAfterAuto Then
            err_str = err_str & "自动设置指定段落的段后间距 设置错误:正确值为:" & title_one_object.paragreph_format.SpaceAfterAuto & ";当前值:" & para_p.Format.SpaceAfterAuto & "." & vbCrLf
        End If


        '段落的行距
        'MsgBox("title_one_object.paragreph_format.LineSpacingRule" & title_one_object.paragreph_format.LineSpacingRule)
        'MsgBox("para_p.Format.LineSpacingRule" & para_p.Format.LineSpacingRule)
        If title_one_object.paragreph_format.LineSpacingRule <> para_p.Format.LineSpacingRule Then
            err_str = err_str & "段落的行距 设置错误:正确值为:" & title_one_object.paragreph_format.LineSpacingRule & ";当前值:" & para_p.Format.LineSpacingRule & "." & vbCrLf
        End If

        '段落的对齐方式
        If title_one_object.paragreph_format.Alignment <> para_p.Format.Alignment Then
            err_str = err_str & "段落的对齐方式 设置错误:正确值为:" & title_one_object.paragreph_format.Alignment & ";当前值:" & para_p.Format.Alignment & "." & vbCrLf
        End If

        '首行缩进的尺寸
        If title_one_object.paragreph_format.FirstLineIndent <> para_p.Format.FirstLineIndent Then
            err_str = err_str & "首行缩进的尺寸 设置错误:正确值为:" & title_one_object.paragreph_format.FirstLineIndent & ";当前值:" & para_p.Format.FirstLineIndent & "." & vbCrLf
        End If



        '左缩进量
        If title_one_object.paragreph_format.CharacterUnitLeftIndent <> para_p.Format.CharacterUnitLeftIndent Then
            err_str = err_str & "左缩进量 设置错误:正确值为:" & title_one_object.paragreph_format.CharacterUnitLeftIndent & ";当前值:" & para_p.Format.CharacterUnitLeftIndent & "." & vbCrLf
        End If

        '右缩进量
        If title_one_object.paragreph_format.CharacterUnitRightIndent <> para_p.Format.CharacterUnitRightIndent Then
            err_str = err_str & "右缩进量 设置错误:正确值为:" & title_one_object.paragreph_format.CharacterUnitRightIndent & ";当前值:" & para_p.Format.CharacterUnitRightIndent & "." & vbCrLf
        End If

        '首行缩进
        If title_one_object.paragreph_format.CharacterUnitFirstLineIndent <> para_p.Format.CharacterUnitFirstLineIndent Then
            err_str = err_str & "特殊格式首行缩进 设置错误:正确值为:" & title_one_object.paragreph_format.CharacterUnitFirstLineIndent & ";当前值:" & para_p.Format.CharacterUnitFirstLineIndent & "." & vbCrLf
        End If

        '段前间距
        If title_one_object.paragreph_format.LineUnitBefore <> para_p.Format.LineUnitBefore Then
            err_str = err_str & "段前间距 设置错误:正确值为:" & title_one_object.paragreph_format.LineUnitBefore & ";当前值:" & para_p.Format.LineUnitBefore & "." & vbCrLf
        End If

        '段后间距
        If title_one_object.paragreph_format.LineUnitAfter <> para_p.Format.LineUnitAfter Then
            err_str = err_str & "段后间距 设置错误:正确值为:" & title_one_object.paragreph_format.LineUnitAfter & ";当前值:" & para_p.Format.LineUnitAfter & "." & vbCrLf
        End If









        ''
        'If title_one_object.paragreph_format.MirrorIndents <> para_p.Format.MirrorIndents Then
        '    err_str = err_str & "MirrorIndents 设置错误:正确值为:" & title_one_object.paragreph_format.MirrorIndents & ";当前值:" & para_p.Format.MirrorIndents & "." & vbCrLf
        'End If


        ''
        'If title_one_object.paragreph_format.TextboxTightWrap <> para_p.Format.TextboxTightWrap Then
        '    err_str = err_str & "TextboxTightWrap 设置错误:正确值为:" & title_one_object.paragreph_format.TextboxTightWrap & ";当前值:" & para_p.Format.TextboxTightWrap & "." & vbCrLf
        'End If


        ''
        'If title_one_object.paragreph_format.CollapsedByDefault <> para_p.Format.CollapsedByDefault Then
        '    err_str = err_str & "CollapsedByDefault 设置错误:正确值为:" & title_one_object.paragreph_format.CollapsedByDefault & ";当前值:" & para_p.Format.CollapsedByDefault & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.AutoAdjustRightIndent <> para_p.Format.AutoAdjustRightIndent Then
        '    err_str = err_str & "AutoAdjustRightIndent 设置错误:正确值为:" & title_one_object.paragreph_format.AutoAdjustRightIndent & ";当前值:" & para_p.Format.AutoAdjustRightIndent & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.DisableLineHeightGrid <> para_p.Format.DisableLineHeightGrid Then
        '    err_str = err_str & "DisableLineHeightGrid 设置错误:正确值为:" & title_one_object.paragreph_format.DisableLineHeightGrid & ";当前值:" & para_p.Format.DisableLineHeightGrid & "." & vbCrLf
        'End If


        ''
        'If title_one_object.paragreph_format.FarEastLineBreakControl <> para_p.Format.FarEastLineBreakControl Then
        '    err_str = err_str & "FarEastLineBreakControl 设置错误:正确值为:" & title_one_object.paragreph_format.FarEastLineBreakControl & ";当前值:" & para_p.Format.FarEastLineBreakControl & "." & vbCrLf
        'End If


        ''
        'If title_one_object.paragreph_format.WordWrap <> para_p.Format.WordWrap Then
        '    err_str = err_str & "WordWrap 设置错误:正确值为:" & title_one_object.paragreph_format.WordWrap & ";当前值:" & para_p.Format.WordWrap & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.HangingPunctuation <> para_p.Format.HangingPunctuation Then
        '    err_str = err_str & "HangingPunctuation 设置错误:正确值为:" & title_one_object.paragreph_format.HangingPunctuation & ";当前值:" & para_p.Format.HangingPunctuation & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.HalfWidthPunctuationOnTopOfLine <> para_p.Format.HalfWidthPunctuationOnTopOfLine Then
        '    err_str = err_str & "HalfWidthPunctuationOnTopOfLine 设置错误:正确值为:" & title_one_object.paragreph_format.HalfWidthPunctuationOnTopOfLine & ";当前值:" & para_p.Format.HalfWidthPunctuationOnTopOfLine & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.AddSpaceBetweenFarEastAndAlpha <> para_p.Format.AddSpaceBetweenFarEastAndAlpha Then
        '    err_str = err_str & "AddSpaceBetweenFarEastAndAlpha 设置错误:正确值为:" & title_one_object.paragreph_format.AddSpaceBetweenFarEastAndAlpha & ";当前值:" & para_p.Format.AddSpaceBetweenFarEastAndAlpha & "." & vbCrLf
        'End If


        ''
        'If title_one_object.paragreph_format.AddSpaceBetweenFarEastAndDigit <> para_p.Format.AddSpaceBetweenFarEastAndDigit Then
        '    err_str = err_str & "AddSpaceBetweenFarEastAndDigit 设置错误:正确值为:" & title_one_object.paragreph_format.AddSpaceBetweenFarEastAndDigit & ";当前值:" & para_p.Format.AddSpaceBetweenFarEastAndDigit & "." & vbCrLf
        'End If

        ''
        'If title_one_object.paragreph_format.BaseLineAlignment <> para_p.Format.BaseLineAlignment Then
        '    err_str = err_str & "BaselineAlignment 设置错误:正确值为:" & title_one_object.paragreph_format.BaseLineAlignment & ";当前值:" & para_p.Format.BaseLineAlignment & "." & vbCrLf
        'End If




        Dim s_comment As Err_comment
        s_comment = New Err_comment
        comment = wd.ActiveDocument.Comments
        return_val = s_comment.Set_comment(para, comment, err_str)
        Return True
    End Function



    Public Function Title_serial_set(ByVal para_p As Word.Paragraph, ByVal wd As Word.Application)
        Dim para As Word.Paragraph
        para = para_p
        Dim err_str As String
        'Dim return_val As Boolean
        'Dim comment As Word.Comments
        title_one_object = create_std_doc.title_one_object
        title_two_object = create_std_doc.title_two_object
        title_three_object = create_std_doc.title_three_object
        title_four_object = create_std_doc.title_four_object
        title_five_object = create_std_doc.title_five_object
        title_six_object = create_std_doc.title_six_object
        title_seven_object = create_std_doc.title_seven_object
        title_eight_object = create_std_doc.title_eight_object
        title_nine_object = create_std_doc.title_nine_object


        Dim t As Object
        t = para_p.Style.ToString
        err_str = ""
        'MsgBox para_p.style.Font.Size
        'para_p.ResetAdvanceTo
        If t = "标题 1" Or para_p.Format.OutlineLevel = title_one_object.paragreph_format.OutlineLevel Then

            para_p.Style = title_one_object.title_one_format
            para_p.Range.Style = title_one_object.title_one_format

            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 2" Or title_two_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_two_object.title_one_format
            para_p.Range.Style = title_two_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 3" Or title_three_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_three_object.title_one_format
            para_p.Range.Style = title_three_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 4" Or title_four_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_four_object.title_one_format
            para_p.Range.Style = title_four_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 5" Or title_five_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_five_object.title_one_format
            para_p.Range.Style = title_five_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 6" Or title_six_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_six_object.title_one_format
            para_p.Range.Style = title_six_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 7" Or title_seven_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_seven_object.title_one_format
            para_p.Range.Style = title_seven_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 8" Or title_eight_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_eight_object.title_one_format
            para_p.Range.Style = title_eight_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        ElseIf t = "标题 9" Or title_nine_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then

            para_p.Style = title_nine_object.title_one_format
            para_p.Range.Style = title_nine_object.title_one_format
            err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        End If

        para_p.Format.LeftIndent = wd.CentimetersToPoints(0)
        para_p.Format.FirstLineIndent = wd.CentimetersToPoints(0)
        'para_p.Range.ListFormat.ApplyListTemplate listtemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(2)
        'para_p.Range.style.LinkToListTemplate listtemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(2)



        '            para_p.Range.style.LinkToListTemplate listtemplate:=lt
        ' err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf

        'Dim s_comment As Err_comment
        's_comment = New Err_comment
        'comment = wd.ActiveDocument.Comments
        'return_val = s_comment.Set_comment(para, comment, err_str)
        Return True
    End Function

    Public Function Title_serial_nothing_set(ByVal para_p As Word.Paragraph, ByVal wd As Word.Application)
        Dim para As Word.Paragraph
        para = para_p
        Dim err_str As String


        title_one_object = create_std_doc.title_one_object
        title_two_object = create_std_doc.title_two_object
        title_three_object = create_std_doc.title_three_object
        title_four_object = create_std_doc.title_four_object
        title_five_object = create_std_doc.title_five_object


        Dim t As Object
        t = para_p.Style.ToString

        err_str = ""
        'MsgBox para_p.style.Font.Size
        'para_p.ResetAdvanceTo
        If t = "标题 1" Or para_p.Format.OutlineLevel = title_one_object.paragreph_format.OutlineLevel Then
            'para_p.Range.Select
            para_p.Style = wd.ActiveDocument.Styles(1)
            'para_p.style = title_one_object.title_one_format

        ElseIf t = "标题 2" Or title_two_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            'para_p.Range.Select
            'para_p.style = title_two_object.title_one_format
            para_p.Style = wd.ActiveDocument.Styles(2)
            'para_p.style = title_two_object.title_one_format

        ElseIf t = "标题 3" Or title_three_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            'para_p.Range.Select
            'para_p.style = title_three_object.title_one_format
            para_p.Style = wd.ActiveDocument.Styles(3)
            'para_p.style = title_three_object.title_one_format
        ElseIf t = "标题 4" Or title_four_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            'para_p.Range.Select
            'para_p.style = title_four_object.title_one_format
            para_p.Style = wd.ActiveDocument.Styles(4)
            'para_p.style = title_four_object.title_one_format

        ElseIf t = "标题 5" Or title_five_object.paragreph_format.OutlineLevel = para_p.Format.OutlineLevel Then
            'para_p.Range.Select
            'para_p.style = title_five_object.title_one_format
            para_p.Style = wd.ActiveDocument.Styles(5)
            'para_p.style = title_five_object.title_one_format
        End If
        'para_p.Range.ListFormat.ApplyListTemplate listtemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(2)
        'para_p.Range.style.LinkToListTemplate listtemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(2)



        '            para_p.Range.style.LinkToListTemplate listtemplate:=lt
        ' err_str = err_str & "多级标题已重设，请注意人工复核!" & vbCrLf
        Return True
    End Function



End Module
