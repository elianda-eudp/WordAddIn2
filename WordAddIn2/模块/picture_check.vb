Module picture_check
    Public Function Picture_check(ByVal para_p As Word.Shape, ByVal para As Word.Paragraph, ByVal wd As Word.Application)
        Dim shape As Word.Shape
        Dim para_fomat As Word.Paragraph
        Dim err_str As String
        Dim return_val As Boolean
        Dim comment As Word.Comments
        picture_object = create_std_doc.picture_object

        shape = para_p
        para_fomat = para

        err_str = ""
        With shape
            '嵌入
            '        If .WrapFormat.Type <> wdWrapInline Then
            '            err_str = err_str & "图片" & "嵌入 设置错误:正确值为:" & wdWrapInline & ";当前值:" & .WrapFormat.Type & "." & vbCrLf
            '        End If
            '        If .Type = msoPicture Then
            .ConvertToInlineShape()
            .WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline
            .WrapFormat.AllowOverlap = False    '不允许重叠
            err_str = err_str & "图片" & "原嵌入方式为悬浮，现被直接修改为嵌入，请人工判断位置是否正确!" & vbCrLf
            '        End If
        End With

        Dim s_comment As Err_comment
        s_comment = New Err_comment
        comment = wd.ActiveDocument.Comments
        return_val = s_comment.Set_comment(para_fomat, comment, err_str)
        Return True
    End Function

    Public Function Inline_picture_check(ByVal para_p As Word.InlineShape, ByVal para As Word.Paragraph, ByVal wd As Word.Application)
        Dim shape As Word.InlineShape
        Dim para_fomat As Word.Paragraph
        Dim err_str As String
        Dim return_val As Boolean
        Dim comment As Word.Comments
        picture_object = create_std_doc.picture_object

        shape = para_p
        para_fomat = para


        With shape
            '嵌入
            '        If .WrapFormat.Type <> wdWrapInline Then
            '            err_str = err_str & "图片" & "嵌入 设置错误:正确值为:" & wdWrapInline & ";当前值:" & .WrapFormat.Type & "." & vbCrLf
            '        End If
            '        If .Type = msoPicture Then
            '            .ConvertToInlineShape
            '            .WrapFormat.Type = wdWrapInline
            '            .WrapFormat.AllowOverlap = False    '不允许重叠
            '        End If
        End With

        err_str = ""
        With picture_object.picture_parag_format

            '对齐
            If .Alignment <> para_fomat.Format.Alignment Then
                err_str = err_str & "图片" & "对齐 设置错误:正确值为:" & .Alignment & ";当前值:" & para_fomat.Format.Alignment & "." & vbCrLf
            End If
            '        para_fomat.Format.Alignment = wdAlignParagraphCenter
        End With
        Dim s_comment As Object = Nothing
        s_comment = New Err_comment
        comment = wd.ActiveDocument.Comments

        return_val = s_comment.set_comment(para_fomat, comment, err_str)
        Return True
    End Function


End Module
