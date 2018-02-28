'    Option Explicit
  Public table_format As Word.table


 

'    Set paragreph_format = New Word.ParagraphFormat
    

Public Function set_format(ByVal style As Word.table)
        Set table_format = style
        
        
    '设置表头
    With table_format.Rows(1).Range
        .Font.Name = "宋体"
        .Font.Size = 12
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Bold = wdToggle
        .Cells.VerticalAlignment = 1 '垂直居中
        
    End With

    '设置表体
    pos = table_format.Cell(2, 1).Range.Start
    pos2 = table_format.Cell(table_format.Rows.Count, table_format.Columns.Count).Range.End

    Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
    With myRange '文档表格第2行到最后一行
        .Font.Size = 10 '字号
        .Font.Name = "宋体" '字体
        .ParagraphFormat.Alignment = wdAlignParagraphLeft '两端对齐(水平左对齐)
        .Cells.VerticalAlignment = 1 '垂直居中
    End With
    
    '设置表格式
    With table_format
        If .style <> "网格型" Then
            .style = "网格型"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        
        .TopPadding = CentimetersToPoints(0)
        .BottomPadding = CentimetersToPoints(0)
        .LeftPadding = CentimetersToPoints(0.19)
        .RightPadding = CentimetersToPoints(0.19)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = True
        .Rows.Alignment = 1 '水平居中
        '.AutoFitBehavior (1) '根据内容自动调整列宽
        .Rows.HeadingFormat = True '行标题重复
        


        
        
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderVertical)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth050pt
            .DefaultBorderColor = wdColorAutomatic
        End With


'    With Selection.Cells(1)
'        .WordWrap = True
'        .FitText = False
'    End With


        
        
        

End Function





