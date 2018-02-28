'    Option Explicit
  Public table_format As Word.table


 

'    Set paragreph_format = New Word.ParagraphFormat
    

Public Function set_format(ByVal style As Word.table)
        Set table_format = style
        
        
    '���ñ�ͷ
    With table_format.Rows(1).Range
        .Font.Name = "����"
        .Font.Size = 12
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Bold = wdToggle
        .Cells.VerticalAlignment = 1 '��ֱ����
        
    End With

    '���ñ���
    pos = table_format.Cell(2, 1).Range.Start
    pos2 = table_format.Cell(table_format.Rows.Count, table_format.Columns.Count).Range.End

    Set myRange = ActiveDocument.Range(Start:=pos, End:=pos2)
    With myRange '�ĵ�����2�е����һ��
        .Font.Size = 10 '�ֺ�
        .Font.Name = "����" '����
        .ParagraphFormat.Alignment = wdAlignParagraphLeft '���˶���(ˮƽ�����)
        .Cells.VerticalAlignment = 1 '��ֱ����
    End With
    
    '���ñ��ʽ
    With table_format
        If .style <> "������" Then
            .style = "������"
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
        .Rows.Alignment = 1 'ˮƽ����
        '.AutoFitBehavior (1) '���������Զ������п�
        .Rows.HeadingFormat = True '�б����ظ�
        


        
        
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





