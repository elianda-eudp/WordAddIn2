Imports WordAddIn2.FormMain
Imports System.Threading.Tasks
Module main_handle
    Public file_name2 As String
    'Public main_from As FormMain
    Sub Main_handle(ByVal wd_p As Word.Application, ByVal form_p As FormMain)

        'MsgBox("form.docpath" + form.docpath)
        Dim wd As Object
        Dim form As Object
        wd = wd_p
        form = form_p

        With wd
            Dim flag As Integer
            flag = 0
            file_name2 = "校验标准文档.doc"
            Dim docNew As Object
            docNew = wd.Documents.Add
            With docNew
                .SaveAs(Filename:=file_name2)
            End With
            'MsgBox(form.docpath)
            wd.Documents(file_name2).Activate()
            'wd.Visible = False

            Call list_template.Listtemplate(wd)
            Call create_std_doc.Create_std_doc_main(form, wd)
            wd.Documents(form.file_name).Activate()

            'wd.ActiveDocument.Styles(1) = title_one_object.title_one_format
            'wd.ActiveDocument.Styles(2) = title_two_object.title_one_format
            'wd.ActiveDocument.Styles(3) = title_three_object.title_one_format
            'wd.ActiveDocument.Styles(4) = title_four_object.title_one_format
            'wd.ActiveDocument.Styles(5) = title_five_object.title_one_format
            'wd.ActiveDocument.Styles(6) = title_six_object.title_one_format
            'wd.ActiveDocument.Styles(7) = title_seven_object.title_one_format
            'wd.ActiveDocument.Styles(8) = title_eight_object.title_one_format
            'wd.ActiveDocument.Styles(9) = title_nine_object.title_one_format
            Call list_template.Listtemplate(wd)

            Dim ThMax As Integer = wd.ActiveDocument.Paragraphs.Count

            For j = 1 To wd.ActiveDocument.Paragraphs.Count
                Dim t As Word.Paragraph
                t = wd.ActiveDocument.Paragraphs(j)
                t.Range.Select()

                'MsgBox .Selection.Information(wdActiveEndAdjustedPageNumber)
                'MsgBox .Selection.Information(wdActiveEndPageNumber)
                'MsgBox BeginPageNumber
                'Call said.mmm

                If j + 1 <= wd.ActiveDocument.Paragraphs.Count Then
                    Call check_title_up.check_title_up(t, wd.ActiveDocument.Paragraphs(j + 1), wd)
                End If

                If .Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber) < form.BeginPageNumber Then

                    'MsgBox .Selection.Information(wdActiveEndPageNumber) & "页不需要检查!"

                ElseIf .Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber) > form.EndPageNumber Then
                    'MsgBox .Selection.Information(wdActiveEndAdjustedPageNumber) & "页结束检查!"
                    Exit For
                Else
                    Dim active_shape As Word.Shape
                    'MsgBox .Selection.Information(wdActiveEndPageNumber) & "页正在检查!"
                    If .Selection.Information(Word.WdInformation.wdWithInTable) Then
                        '表
                        If flag = 0 And t.Range.Tables.Count > 0 Then
                            Dim active_table As Word.Table
                            active_table = t.Range.Tables(1)
                            Call check_oper.Check_table(active_table, wd)
                            t.Range.Tables(1).Select()
                            Dim col As Object
                            col = .Selection.Information(Word.WdInformation.wdEndOfRangeColumnNumber)
                            Dim Row As Object
                            Row = .Selection.Information(Word.WdInformation.wdEndOfRangeRowNumber)
                            j = j + Row * (col + 1) - 2
                            flag = 1
                            If .ActiveDocument.PageSetup.Orientation = 0 Then
                                active_table.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                                active_table.PreferredWidth = wd.CentimetersToPoints(18)
                            Else
                                active_table.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
                                active_table.PreferredWidth = wd.CentimetersToPoints(26)
                            End If

                        End If
                    ElseIf t.Range.ShapeRange.Count > 0 Then

                        '图
                        active_shape = t.Range.ShapeRange(1)
                        Call check_oper.Check_shape(active_shape, t, wd)
                        flag = 0
                    ElseIf t.Range.InlineShapes.Count > 0 Then

                        '图
                        Dim in_active_shape As Word.InlineShape
                        in_active_shape = t.Range.InlineShapes(1)
                        Call picture_check.Inline_picture_check(in_active_shape, t, wd)
                        flag = 0
                    Else
                        If Len(Trim(wd.ActiveDocument.Paragraphs(j).Range.Text)) = 1 Then
                            'MsgBox wd.ActiveDocument.Paragraphs(i - 1).Range.text
                            wd.ActiveDocument.Paragraphs(j).Range.Delete()
                            j = j - 1
                        Else
                            Call check_oper.Check_main(t, wd)
                        End If

                        flag = 0
                    End If
                End If

                If j >= wd.ActiveDocument.Paragraphs.Count Then
                    Exit For
                End If
            Next j

        End With
        wd.Documents(form.file_name).Activate()
        Call list_template.Listtemplate(wd)
        'With wd
        '    Dim i As Integer
        '    i = 1
        '    Dim flag As Integer
        '    flag = 0
        '    Do
        '        If .Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber) < form.BeginPageNumber Then

        '            'MsgBox .Selection.Information(wdActiveEndPageNumber) & "页不需要检查!"

        '        ElseIf .Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber) > form.EndPageNumber Then
        '            'MsgBox .Selection.Information(wdActiveEndAdjustedPageNumber) & "页结束检查!"
        '            Exit Do
        '        Else

        '            If Len(Trim(wd.ActiveDocument.Paragraphs(i).Range.Text)) = 1 Then
        '                wd.ActiveDocument.Paragraphs(i).Range.Delete()
        '            End If
        '        End If

        '        i = i + 1

        '        If i > wd.ActiveDocument.Paragraphs.Count Then
        '            Exit Do
        '        End If

        '    Loop
        'End With
        'wd.Documents(file_name2).Close(SaveChanges:=Word.WdSaveOptions.wdDoNotSaveChanges)
        wd.Documents(file_name2).Save()
        wd.Documents(form.file_name).Save()
        wd.Documents(file_name2).Close()

    End Sub

End Module
