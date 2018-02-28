    Sub check_title_up() '标题下无正文检查

        Dim ShapesNum As Long, i As Long, j As Long
        'Application.ScreenUpdating = False
        Dim t As Word.Paragraph
        Dim flag As Integer
        flag = 0
        For i = 1 To ActiveDocument.Paragraphs.Count
        
            If i > ActiveDocument.Paragraphs.Count Then
                Exit For
            End If
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
                If Len(Trim(ActiveDocument.Paragraphs(i).Range.Text)) = 1 Then
                    ActiveDocument.Paragraphs(i).Range.Delete
                    i = i - 1
                Else
                    If i + 1 <= ActiveDocument.Paragraphs.Count Then
                        If t.Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Or ActiveDocument.Paragraphs(i + 1).Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Then

                        ElseIf t.Format.OutlineLevel >= ActiveDocument.Paragraphs(i + 1).Format.OutlineLevel And ActiveDocument.Paragraphs(i + 1).Format.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevelBodyText Then
                            Dim err_str As String
                            Dim return_val As Boolean
                            Dim comments As Word.comments
                            Dim comment As Word.comment
                            comment_str = "标题下无正文，内容错误!" & vbCrLf

                            Set comments = ActiveDocument.comments
                            Set comment = comments.Add(t.Range, comment_str)
                            comment.Author = "coin wo-wo"
                            comment.Range.Text = comment_str
                        End If
                    End If
                End If
            End If
        Next i
        Application.ScreenUpdating = True

    End Sub


