

Sub SelectAllShapes() '选中文档中的所有图片，并设置格式为嵌入居中

  Dim ShapesNum As Long, i As Long
  'Application.ScreenUpdating = False
  Dim ThMax As Integer
  Dim t As Word.Paragraph
  Dim flag As Integer
  ThMax = ActiveDocument.Paragraphs.Count
  flag = 0
    For i = 1 To ThMax
    'MsgBox i
        Set t = ActiveDocument.Paragraphs(i)
        t.Range.Select
        
        If flag = 0 And Selection.Information(Word.WdInformation.wdWithInTable) Then
            Dim col As Integer
            col = Selection.Information(Word.WdInformation.wdEndOfRangeColumnNumber)
            Dim Row As Integer
            Row = Selection.Information(Word.WdInformation.wdEndOfRangeRowNumber)
            i = i + Row * (col + 1) - 2
            flag = 1
        ElseIf flag = 1 And Selection.Information(Word.WdInformation.wdWithInTable) Then
        
        Else
        
            flag = 0
            If t.Range.ShapeRange.Count > 0 Then
            '图
                Dim active_shape As Object
                'active_shape =
                With t.Range.ShapeRange(1)
                    .WrapFormat.Type = wdWrapInline
                    .WrapFormat.AllowOverlap = False    '不允许重叠
                    .ConvertToInlineShape

                End With
            
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
            ElseIf t.Range.InlineShapes.Count > 0 Then
                '图
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
            End If
        End If
        
    Next i
  Application.ScreenUpdating = True

End Sub
