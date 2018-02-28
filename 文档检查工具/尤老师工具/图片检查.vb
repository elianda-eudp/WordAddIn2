    '设置表格式
    With picture_format
        If .Type = msoPicture Then
            .ConvertToInlineShape
        End If
    End With
    
    para.Alignment = wdAlignParagraphCenter
    