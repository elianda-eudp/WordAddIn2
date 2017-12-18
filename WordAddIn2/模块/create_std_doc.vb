Imports System.Collections
Imports System.Resources

Module create_std_doc
    Public title_three_object As Title_three

    Public title_one_object As Title_one
    Public title_two_object As Title_two

    Public title_four_object As Title_four
    Public title_five_object As Title_five
    Public title_six_object As Title_six
    Public title_seven_object As Title_seven
    Public title_eight_object As Title_eight
    Public title_nine_object As Title_nine

    Public text_object As Text
    Public table_object As Table
    Public picture_object As Picture
    Sub Create_std_doc_main(ByVal form As FormMain, ByVal wd As Word.Application)

        'Dim docpath As String
        'Dim file_name As String

        'Dim i As Integer
        Dim return_val As Boolean
        Dim rngFormat As Object
        Dim t As Object
        Dim para As Object
        Dim std_wd As Word.Application
        std_wd = wd
        'Call word_files_check.ActiveDocumentName
        wd.Documents(main_handle.file_name2).Activate()
        'Call word_files_check.ActiveDocumentName




        title_one_object = New Title_one
        title_two_object = New Title_two
        title_three_object = New Title_three
        title_four_object = New Title_four
        title_five_object = New Title_five
        title_six_object = New Title_six
        title_seven_object = New Title_seven
        title_eight_object = New Title_eight
        title_nine_object = New Title_nine
        text_object = New Text
        table_object = New Table
        picture_object = New Picture

        '    MsgBox "二号" & 二号
        '    MsgBox "小二号" & 小二号
        '    MsgBox "三号" & 三号
        '    MsgBox "小三号" & 小三号


        rngFormat = std_wd.ActiveDocument.Range(Start:=0, End:=0)
        With rngFormat
            .InsertAfter(text:="我是标题一")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题二")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题三")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题四")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题五")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题六")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题七")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题八")
            .InsertParagraphAfter
            .InsertAfter(text:="我是标题九")
            .InsertParagraphAfter
            .InsertAfter(text:="我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文我是正文")
            .InsertParagraphAfter
            .InsertParagraphAfter

        End With
        std_wd.ActiveDocument.Paragraphs(1).Style = std_wd.ActiveDocument.Styles("标题 1")
        std_wd.ActiveDocument.Paragraphs(2).Style = std_wd.ActiveDocument.Styles("标题 2")
        std_wd.ActiveDocument.Paragraphs(3).Style = std_wd.ActiveDocument.Styles("标题 3")
        std_wd.ActiveDocument.Paragraphs(4).Style = std_wd.ActiveDocument.Styles("标题 4")
        std_wd.ActiveDocument.Paragraphs(5).Style = std_wd.ActiveDocument.Styles("标题 5")
        std_wd.ActiveDocument.Paragraphs(6).Style = std_wd.ActiveDocument.Styles("标题 6")
        std_wd.ActiveDocument.Paragraphs(7).Style = std_wd.ActiveDocument.Styles("标题 7")
        std_wd.ActiveDocument.Paragraphs(8).Style = std_wd.ActiveDocument.Styles("标题 8")
        std_wd.ActiveDocument.Paragraphs(9).Style = std_wd.ActiveDocument.Styles("标题 9")

        '    std_wd.ActiveDocument.Paragraphs(1).Range.SetListLevel (1)
        '    std_wd.ActiveDocument.Paragraphs(2).Range.SetListLevel (2)
        '    std_wd.ActiveDocument.Paragraphs(3).Range.SetListLevel (3)
        '    std_wd.ActiveDocument.Paragraphs(4).Range.SetListLevel (4)
        '    std_wd.ActiveDocument.Paragraphs(5).Range.SetListLevel (5)
        std_wd.ActiveDocument.Paragraphs(10).Style = std_wd.ActiveDocument.Styles("正文")


        std_wd.ActiveDocument.Tables.Add(Range:=std_wd.ActiveDocument.Paragraphs(11).Range, NumRows:=4, NumColumns:=7)
        Dim mypic As Microsoft.Office.Interop.Word.InlineShape

        std_wd.ActiveDocument.Range(std_wd.ActiveDocument.Content.End - 1, std_wd.ActiveDocument.Content.End - 1).Select()
        'Dim Picture_1 As New ResourceReader("Visual Studio 2017\Projects\WordAddIn2\WordAddIn2\Resource.resx")
        'Picture_1 = New ResXResourceReader("Resource.resx")
        'Picture_1.BasePath = "C:\Users\Administrator\Documents\Visual Studio 2017\Projects\WordAddIn2\WordAddIn2"

        'Dim id As IDictionaryEnumerator = Picture_1.GetEnumerator() '创建资源ID枚举
        'id.MoveNext() '开始枚举推进
        'mypic = id.Value '使用图像

        'Picture_1.Close()

        'mypic = std_wd.Selection.InlineShapes.AddPicture(form.docpath & "\字号对照表.gif")
        mypic = std_wd.Selection.InlineShapes.AddPicture("C:\Program Files (x86)\Microsoft\file_check\字号对照表.gif")

        'mypic = My.Resources.Image1





        mypic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
        mypic.ConvertToShape()

        'Call word_files_check.ActiveDocumentName



        return_val = title_one_object.Set_format(std_wd.ActiveDocument.Paragraphs(1).Style, std_wd.ActiveDocument.Paragraphs(1).Format, std_wd)
        'MsgBox("title_one_object.paragreph_format.LineSpacingRule" & title_one_object.paragreph_format.LineSpacingRule)

        return_val = title_two_object.Set_format(std_wd.ActiveDocument.Paragraphs(2).Style, std_wd.ActiveDocument.Paragraphs(2).Format, std_wd)

        return_val = title_three_object.Set_format(std_wd.ActiveDocument.Paragraphs(3).Style, std_wd.ActiveDocument.Paragraphs(3).Format, std_wd)

        return_val = title_four_object.Set_format(std_wd.ActiveDocument.Paragraphs(4).Style, std_wd.ActiveDocument.Paragraphs(4).Format, std_wd)

        return_val = title_five_object.Set_format(std_wd.ActiveDocument.Paragraphs(5).Style, std_wd.ActiveDocument.Paragraphs(5).Format, std_wd)

        return_val = title_six_object.Set_format(std_wd.ActiveDocument.Paragraphs(6).Style, std_wd.ActiveDocument.Paragraphs(6).Format, std_wd)

        return_val = title_seven_object.Set_format(std_wd.ActiveDocument.Paragraphs(7).Style, std_wd.ActiveDocument.Paragraphs(7).Format, std_wd)

        return_val = title_eight_object.Set_format(std_wd.ActiveDocument.Paragraphs(8).Style, std_wd.ActiveDocument.Paragraphs(8).Format, std_wd)

        return_val = title_nine_object.Set_format(std_wd.ActiveDocument.Paragraphs(9).Style, std_wd.ActiveDocument.Paragraphs(9).Format, std_wd)

        return_val = text_object.Set_format(std_wd.ActiveDocument.Paragraphs(10).Style, std_wd.ActiveDocument.Paragraphs(10).Format, std_wd)

        If std_wd.ActiveDocument.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait Then
            std_wd.ActiveDocument.Tables(1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
            std_wd.ActiveDocument.Tables(1).PreferredWidth = std_wd.CentimetersToPoints(18)
        Else
            std_wd.ActiveDocument.Tables(1).PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints
            std_wd.ActiveDocument.Tables(1).PreferredWidth = std_wd.CentimetersToPoints(26)
        End If

        return_val = table_object.Set_format(std_wd.ActiveDocument.Tables(1), std_wd)

        'std_wd.ActiveDocument.Shapes.SelectAll
        'MsgBox title_three_object.title_one_format.Font.NameFarEast

        'Call word_files_check.ActiveDocumentName
        t = std_wd.ActiveDocument.Shapes(1)

        para = std_wd.Selection.ParagraphFormat
        return_val = picture_object.Set_format(t, para)





    End Sub



End Module
