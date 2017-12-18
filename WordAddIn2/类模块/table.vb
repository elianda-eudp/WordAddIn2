Public Class Table
    '    Option Explicit
    Public table_format As Word.Table




    '    Set paragreph_format = New Word.ParagraphFormat


    Public Function Set_format(ByVal style As Word.Table, ByVal wd As Word.Application)
        table_format = style


        '设置表头
        With table_format.Rows(1).Range
            .Font.Name = "宋体"
            .Font.Size = 12
            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Font.Bold = Word.WdConstants.wdToggle
            .Cells.VerticalAlignment = 1 '垂直居中

        End With

        Dim pos As Integer
        '设置表体
        pos = table_format.Cell(2, 1).Range.Start
        Dim pos2 As Integer
        pos2 = table_format.Cell(table_format.Rows.Count, table_format.Columns.Count).Range.End

        Dim myrange As Object
        myrange = wd.ActiveDocument.Range(Start:=pos, End:=pos2)
        With myrange '文档表格第2行到最后一行
            .Font.Size = 10 '字号
            .Font.Name = "宋体" '字体
            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft '两端对齐(水平左对齐)
            .Cells.VerticalAlignment = 1 '垂直居中
        End With

        '设置表格式
        With table_format
            'If .Style <> "网格型" Then
            '.Style = "网格型"
            'End If
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = False
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = False
            .ApplyStyleRowBands = True
            .ApplyStyleColumnBands = False

            .TopPadding = wd.CentimetersToPoints(0)
            .BottomPadding = wd.CentimetersToPoints(0)
            .LeftPadding = wd.CentimetersToPoints(0.19)
            .RightPadding = wd.CentimetersToPoints(0.19)
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowAutoFit = True
            .Rows.Alignment = 1 '水平居中
            '.AutoFitBehavior (1) '根据内容自动调整列宽
            .Rows.HeadingFormat = True '行标题重复





            With .Borders(Word.WdBorderType.wdBorderLeft)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With

            With .Borders(Word.WdBorderType.wdBorderRight)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With
            With .Borders(Word.WdBorderType.wdBorderTop)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With
            With .Borders(Word.WdBorderType.wdBorderBottom)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With
            With .Borders(Word.WdBorderType.wdBorderHorizontal)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With
            With .Borders(Word.WdBorderType.wdBorderVertical)
                .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                .Color = Word.WdColor.wdColorAutomatic
            End With
            .Borders(Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Borders(Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Borders.Shadow = False
        End With

        With wd.Options
            .DefaultBorderLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            .DefaultBorderLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt
            .DefaultBorderColor = Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic
        End With


        '    With Selection.Cells(1)
        '        .WordWrap = True
        '        .FitText = False
        '    End With





        Return True
    End Function






End Class
