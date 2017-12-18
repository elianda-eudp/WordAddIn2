Module table_check
    Public Function Table_check(ByVal table_p As Word.Table, ByVal wd As Word.Application)

        Dim err_str As String
        Dim return_val As Boolean
        Dim comment As Word.Comments
        table_object = create_std_doc.table_object
        ' Dim i As Long, j As Long
        'Dim select_con As Word.Section

        err_str = ""

        '校验表头
        With table_object.table_format.Rows(1).Range
            '字体
            If .Font.Name <> table_p.Rows(1).Range.Font.Name Then
                err_str = err_str & "表头" & "字体 设置错误:正确值为:" & .Font.Name & ";当前值:" & table_p.Range.Font.Name & "." & vbCrLf
            End If

            '字号
            If .Font.Size <> table_p.Rows(1).Range.Font.Size Then
                err_str = err_str & "表头" & "字号 设置错误:正确值为:" & .Font.Size & ";当前值:" & table_p.Range.Font.Size & "." & vbCrLf
            End If

            '对齐方式
            If .ParagraphFormat.Alignment <> table_p.Rows(1).Range.ParagraphFormat.Alignment Then
                err_str = err_str & "表头" & "对齐方式 设置错误:正确值为:" & .ParagraphFormat.Alignment & ";当前值:" & table_p.Range.ParagraphFormat.Alignment & "." & vbCrLf
            End If

            '粗体
            If .Font.Bold <> table_p.Rows(1).Range.Font.Bold Then
                err_str = err_str & "表头" & "粗体 设置错误:正确值为:" & .Font.Bold & ";当前值:" & table_p.Range.Font.Bold & "." & vbCrLf
            End If

            '垂直对齐方式
            If .Cells.VerticalAlignment <> table_p.Rows(1).Range.Cells.VerticalAlignment Then
                err_str = err_str & "表头" & "垂直对齐方式 设置错误:正确值为:" & .Cells.VerticalAlignment & ";当前值:" & table_p.Range.Cells.VerticalAlignment & "." & vbCrLf
            End If

        End With

        Dim s_comment As Err_comment
        s_comment = New Err_comment
        comment = wd.ActiveDocument.Comments
        Dim myrange As Word.Range
        myrange = table_p.Rows(1).Range
        return_val = s_comment.Range_set_comment(myrange, comment, err_str)

        s_comment = Nothing
        comment = Nothing
        err_str = ""

        Dim pos As Integer
        '设置表体
        pos = table_p.Cell(2, 1).Range.Start
        Dim pos2 As Integer
        'MsgBox(table_p.Rows.Count)
        'MsgBox(table_p.Rows(table_p.Rows.Count).Cells.Count)
        pos2 = table_p.Cell(table_p.Rows.Count, table_p.Rows(table_p.Rows.Count).Cells.Count).Range.End
        'MsgBox(table_p.Cell(table_p.Rows.Count, table_p.Rows(table_p.Rows.Count).Cells.Count).Range.End)

        myrange = wd.ActiveDocument.Range(Start:=pos, End:=pos2)
        With myrange '文档表格第2行到最后一行
            .Font.Size = 10 '字号
            .Font.Name = "宋体" '字体
            .Font.Bold = False
            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft '(水平左对齐)
            .Cells.VerticalAlignment = 1 '垂直居中
        End With


        '校验表格式
        With table_object.table_format
            '网格型
            If .Style.ToString <> table_p.Style.ToString Then
                err_str = err_str & "表格" & "网格型 设置错误:正确值为:" & .Style.ToString & ";当前值:" & table_p.Style.ToString & "." & vbCrLf
            End If

            '对齐方式
            If .Rows.Alignment <> table_p.Rows.Alignment Then
                err_str = err_str & "表格" & "对齐方式 设置错误:正确值为:" & .Rows.Alignment & ";当前值:" & table_p.Rows.Alignment & "." & vbCrLf
            End If

            '行标题重复
            If .Rows.HeadingFormat <> table_p.Rows.HeadingFormat Then
                err_str = err_str & "表格" & "行标题重复 设置错误:正确值为:" & .Rows.HeadingFormat & ";当前值:" & table_p.Rows.HeadingFormat & "." & vbCrLf
            End If


        End With


        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft)

            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle Then
                err_str = err_str & "表格" & "左边框线型 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle & "." & vbCrLf
            End If


            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineWidth Then
                err_str = err_str & "表格" & "左边框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineWidth & "." & vbCrLf
            End If


            '        If .LineStyle <> table_p.Borders(wdBorderLeft).Color Then
            '            err_str = err_str & "表格" & "线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderLeft).Color & "." & vbCrLf
            '        End If


        End With


        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight)
            '
            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle Then
                err_str = err_str & "表格" & "右边框线型 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle & "." & vbCrLf
            End If

            '
            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineWidth Then
                err_str = err_str & "表格" & "右边框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineWidth & "." & vbCrLf
            End If

            '
            '        If .LineStyle <> table_p.Borders(wdBorderRight).Color Then
            '            err_str = err_str & "表格" & "线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderRight).Color & "." & vbCrLf
            '        End If


        End With

        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop)
            '
            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle Then
                err_str = err_str & "表格" & "上框线 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle & "." & vbCrLf
            End If

            '
            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineWidth Then
                err_str = err_str & "表格" & "上框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineWidth & "." & vbCrLf
            End If

            '
            '        If .LineStyle <> table_p.Borders(wdBorderTop).Color Then
            '            err_str = err_str & "表格" & "上框线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderTop).Color & "." & vbCrLf
            '        End If


        End With


        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
            '
            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle Then
                err_str = err_str & "表格" & "底边框线 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle & "." & vbCrLf
            End If

            '
            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineWidth Then
                err_str = err_str & "表格" & "底边框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineWidth & "." & vbCrLf
            End If

            '
            '        If .LineStyle <> table_p.Borders(wdBorderBottom).Color Then
            '            err_str = err_str & "表格" & "底边框线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderBottom).Color & "." & vbCrLf
            '        End If


        End With


        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal)
            '
            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle Then
                err_str = err_str & "表格" & "横向框线 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle & "." & vbCrLf
            End If

            '
            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineWidth Then
                err_str = err_str & "表格" & "横向框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineWidth & "." & vbCrLf
            End If

            '
            '        If .LineStyle <> table_p.Borders(wdBorderHorizontal).Color Then
            '            err_str = err_str & "表格" & "横向框线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderHorizontal).Color & "." & vbCrLf
            '        End If


        End With

        With table_object.table_format.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical)
            '
            If .LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle Then
                err_str = err_str & "表格" & "纵向框线 设置错误:正确值为:" & .LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle & "." & vbCrLf
            End If

            '
            If .LineWidth <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineWidth Then
                err_str = err_str & "表格" & "纵向框线宽度 设置错误:正确值为:" & .LineWidth & ";当前值:" & table_p.Borders(Word.WdBorderType.wdBorderVertical).LineWidth & "." & vbCrLf
            End If

            '
            '        If .LineStyle <> table_p.Borders(wdBorderVertical).Color Then
            '            err_str = err_str & "表格" & "纵向框线颜色 设置错误:正确值为:" & .Color & ";当前值:" & table_p.Borders(wdBorderVertical).Color & "." & vbCrLf
            '        End If


        End With

        With table_object.table_format
            '
            If .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle Then
                err_str = err_str & "表格" & "方向从左上角开始的斜向边框线 设置错误:正确值为:" & .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle & "." & vbCrLf
            End If

            If .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle <> table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle Then
                err_str = err_str & "表格" & "方向从左下角开始的斜向边框线 设置错误:正确值为:" & .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle & ";当前值:" & table_p.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle & "." & vbCrLf
            End If

            If .Borders.Shadow <> table_p.Borders.Shadow Then
                err_str = err_str & "表格" & "边框设置为阴影格式 设置错误:正确值为:" & .Borders.Shadow & ";当前值:" & table_p.Borders.Shadow & "." & vbCrLf
            End If



        End With

        'For i = 1 To table_p.Rows.Count
        '    For j = 1 To table_p.Rows(i).Cells.Count
        '        If Len(table_p.Cell(Row:=i, Column:=j).Range.Text) = 2 Then
        '            err_str = err_str & "表格坐标 行:" & i & "列:" & j & "内容为空" & vbCrLf
        '        End If
        '    Next j
        'Next i




        s_comment = New Err_comment
        comment = wd.ActiveDocument.Comments
        return_val = s_comment.Table_set_comment(table_p, comment, err_str)

        Return True
    End Function

End Module
