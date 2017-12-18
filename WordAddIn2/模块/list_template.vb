Imports Microsoft.Office.Interop.Word

Module list_template

    Public lt As Word.ListTemplate


    Public Function Clear_listtemplate(ByVal wd As Word.Application)
        For i = 1 To 9

            With wd.ActiveDocument.Styles("标题 " & i)
                .LinkToListTemplate（ListTemplate:=Nothing）
                .AutomaticallyUpdate = False
                .NoSpaceBetweenParagraphsOfSameStyle = False
                '.ParagraphFormat.TabStops.ClearAll
                .BaseStyle = ""
                .NextParagraphStyle = "正文"
            End With
        Next i
        With wd.ActiveDocument.Styles("标题 1")
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "正文"
        End With

        With wd.ActiveDocument.Styles("标题 2")
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "正文"
        End With

        With wd.ActiveDocument.Styles("标题 3")
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "正文"
        End With

        With wd.ActiveDocument.Styles("标题 4")
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "正文"
        End With

        With wd.ActiveDocument.Styles("标题 5")
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "正文"
        End With
        Return True
    End Function


    Public Function Listtemplate(ByVal wd As Word.Application)
        Dim str As String


        lt = wd.ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
        str = ""
        For i = 1 To 9

            lt.ListLevels(i).TrailingCharacter = Microsoft.Office.Interop.Word.WdTrailingCharacter.wdTrailingTab
            lt.ListLevels(i).NumberStyle = Microsoft.Office.Interop.Word.WdListNumberStyle.wdListNumberStyleArabic
            lt.ListLevels(i).NumberPosition = Fix(0)
            lt.ListLevels(i).Alignment = WdListLevelAlignment.wdListLevelAlignLeft
            lt.ListLevels(i).TextPosition = wd.CentimetersToPoints(0.75)
            lt.ListLevels(i).StartAt = 1
            lt.ListLevels(i).ResetOnHigher = i - 1
            str = str & "%" & i & "."
            lt.ListLevels(i).NumberFormat = str
            lt.ListLevels(i).LinkedStyle = "标题 " & i

            With wd.ActiveDocument.Styles("标题 " & i)
                .LinkToListTemplate(ListTemplate:=lt)
            End With
            'MsgBox("标题 " & i)
        Next i
        'ActiveDocument.Content.ListFormat.ApplyListTemplate listtemplate:=lt
        Return True
    End Function

End Module
