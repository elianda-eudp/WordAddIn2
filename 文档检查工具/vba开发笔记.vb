一、本示例将 myRange 设置为代表活动文档内容区域，折叠 myRange，然后在文档的最后插入一个 2ｘ2 表格。

Set myRange = ActiveDocument.Content
myRange.Collapse Direction:=wdCollapseEnd
ActiveDocument.Tables.Add Range:=myRange, NumRows:=2, NumColumns:=2


示例
本示例将所选内容折叠为选定部分的开头。

Selection.Collapse Direction:=wdCollapseStart


二、
Selection 对象有多种方法和属性，可用于折叠、扩展或以其他方式更改当前所选的内容。下列示例将插入点移动到到文档末尾并选择最后三行内容。

Selection.EndOf Unit:=wdStory, Extend:=wdMove
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
Selection 对象有多种方法和属性，可用于编辑文档中的所选内容。下列示例选择活动文档中的第一句，并用新的段落替换该句。

Options.ReplaceSelection = True
ActiveDocument.Sentences(1).Select
Selection.TypeText "Material below is confidential."
Selection.TypeParagraph


使用诸如 Flags、Information 和 Type 属性可返回当前所选内容的信息。您可以在过程中使用下列示例来确定活动文档中是否的确有选定的内容；如果没有，将略过该过程的其余部分。 

If Selection.Type = wdSelectionIP Then
    MsgBox Prompt:="You haven't selected any text! Exiting procedure..."
    Exit Sub
End If


即使将所选内容折叠为插入点，该对象也不一定为空。例如，Text属性仍将返回至插入点右侧的字符；该字符也将在 Selection对象的Characters集合中显示。但是，来自折叠的所选内容的诸如Cut 或Copy 等调用方法将导致出错。

用户可以选定代表不连续文本的文档区域（例如，使用鼠标和 Alt 键）。由于该行为不可预知，您可能希望在代码中包含一个步骤，用于在对所选内容的Type属性进行任何操作前对其进行检查（Selection.Type = wdSelectionBlock）。类似地，包含表格单元格的所选内容也会导致不可预知的行为。Information属性将告知您所选内容是否在表格中。（Selection.Information(wdWithinTable) = True）。下列示例确定所选内容是否正常（换言之，它不是表格中的一行或一列，不是文本中的一竖排区，等等）；您可以在执行任何操作前用它测试当前所选内容。


If Selection.Type <> wdSelectionNormal Then
    MsgBox Prompt:="Not a valid selection! Exiting procedure..."
    Exit Sub
End If


解决vba 解决“远程服务器不存在或不可用”
后来我在其他高手的帖子找到了另一个原因（彻底解决了问题），就是在整段代码倒数第二行（也就是end sub前）加上end，把word进程强退出，就不会再出现那个问题了。

https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.word.wdbuiltinstyle.aspx#Anchor_1


四、word遍历段落怎么判断段落是否在表格中呢

Sub 循环遍历段落()
'如果段落在表格中，则变为红色；否则蓝色
    Dim i As Paragraph
    For Each i In ActiveDocument.Paragraphs
        If i.Range.Information(wdWithInTable) = True Then i.Range.Font.Color = wdColorRed Else i.Range.Font.Color = wdColorBlue
    Next
End Sub


wdWrapInline 7 将形状嵌入到文字中。
wdWrapNone 3 将形状放在文字前面。请参阅  wdWrapFront 。
wdWrapSquare 0 使文字环绕形状。行在形状的另一侧延续。
wdWrapThrough 2 使文字环绕形状。
wdWrapTight 1 使文字紧密地环绕形状。
wdWrapTopBottom 4 将文字放在形状的上方和下方。
wdWrapBehind 5 将形状放在文字后面。
wdWrapFront 6 将形状放在文字前面。


悬浮的图片处理方式，直接转换，等待处理

    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior


