Module check_title_up
    Public Function check_title_up(ByVal para_1 As Word.Paragraph, ByVal para_2 As Word.Paragraph, ByVal wd As Word.Application)

        If para_1.Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Or para_2.Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Then

        ElseIf para_1.Format.OutlineLevel >= para_2.Format.OutlineLevel And para_2.Format.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevelBodyText Then
            Dim err_str As String
            Dim return_val As Boolean
            Dim comment As Word.Comments
            table_object = create_std_doc.table_object
            ' Dim i As Long, j As Long
            'Dim select_con As Word.Section
            err_str = ""
            err_str = err_str & "标题下无正文，内容错误!" & vbCrLf
            Dim s_comment As Err_comment
            s_comment = New Err_comment
            comment = wd.ActiveDocument.Comments
            return_val = s_comment.Set_comment(para_1, comment, err_str)
        End If


        Return True

    End Function
End Module
