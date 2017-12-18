Module check_oper
    Public Function Check_main(ByVal para_p As Word.Paragraph, ByVal wd As Word.Application)
        Dim para As Word.Paragraph
        para = para_p

        Dim t As Object
        t = para_p.Style.ToString

        If InStr(t, "标题") > 0 Or para.Format.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevelBodyText Then
            Call title_check.Title_serial_set(para, wd)
            Call title_check.Title_check(para, wd)

        ElseIf InStr(t, "正文") Or para.Range.ParagraphFormat.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText Then
            Call content_check.Content_check(para, wd)
        End If


        Return True

    End Function

    Public Function Check_table(ByVal table_p As Word.Table, ByVal wd As Word.Application)
        Dim table As Word.Table
        table = table_p


        Call table_check.Table_check(table, wd)


        Return True
    End Function


    Public Function Check_shape(ByVal Shape_p As Word.Shape, ByVal para As Word.Paragraph, ByVal wd As Word.Application)
        Dim shape As Word.Shape
        shape = Shape_p

        Call picture_check.Picture_check(shape, para, wd)



        Return True
    End Function

End Module
