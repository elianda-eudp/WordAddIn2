һ����ʾ���� myRange ����Ϊ�����ĵ����������۵� myRange��Ȼ�����ĵ���������һ�� 2��2 ���

Set myRange = ActiveDocument.Content
myRange.Collapse Direction:=wdCollapseEnd
ActiveDocument.Tables.Add Range:=myRange, NumRows:=2, NumColumns:=2


ʾ��
��ʾ������ѡ�����۵�Ϊѡ�����ֵĿ�ͷ��

Selection.Collapse Direction:=wdCollapseStart


����
Selection �����ж��ַ��������ԣ��������۵�����չ����������ʽ���ĵ�ǰ��ѡ�����ݡ�����ʾ����������ƶ������ĵ�ĩβ��ѡ������������ݡ�

Selection.EndOf Unit:=wdStory, Extend:=wdMove
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
Selection �����ж��ַ��������ԣ������ڱ༭�ĵ��е���ѡ���ݡ�����ʾ��ѡ���ĵ��еĵ�һ�䣬�����µĶ����滻�þ䡣

Options.ReplaceSelection = True
ActiveDocument.Sentences(1).Select
Selection.TypeText "Material below is confidential."
Selection.TypeParagraph


ʹ������ Flags��Information �� Type ���Կɷ��ص�ǰ��ѡ���ݵ���Ϣ���������ڹ�����ʹ������ʾ����ȷ����ĵ����Ƿ��ȷ��ѡ�������ݣ����û�У����Թ��ù��̵����ಿ�֡� 

If Selection.Type = wdSelectionIP Then
    MsgBox Prompt:="You haven't selected any text! Exiting procedure..."
    Exit Sub
End If


��ʹ����ѡ�����۵�Ϊ����㣬�ö���Ҳ��һ��Ϊ�ա����磬Text�����Խ�������������Ҳ���ַ������ַ�Ҳ���� Selection�����Characters��������ʾ�����ǣ������۵�����ѡ���ݵ�����Cut ��Copy �ȵ��÷��������³���

�û�����ѡ�����������ı����ĵ��������磬ʹ������ Alt ���������ڸ���Ϊ����Ԥ֪��������ϣ���ڴ����а���һ�����裬�����ڶ���ѡ���ݵ�Type���Խ����κβ���ǰ������м�飨Selection.Type = wdSelectionBlock�������Ƶأ��������Ԫ�����ѡ����Ҳ�ᵼ�²���Ԥ֪����Ϊ��Information���Խ���֪����ѡ�����Ƿ��ڱ���С���Selection.Information(wdWithinTable) = True��������ʾ��ȷ����ѡ�����Ƿ�����������֮�������Ǳ���е�һ�л�һ�У������ı��е�һ���������ȵȣ�����������ִ���κβ���ǰ�������Ե�ǰ��ѡ���ݡ�


If Selection.Type <> wdSelectionNormal Then
    MsgBox Prompt:="Not a valid selection! Exiting procedure..."
    Exit Sub
End If


���vba �����Զ�̷����������ڻ򲻿��á�
���������������ֵ������ҵ�����һ��ԭ�򣨳��׽�������⣩�����������δ��뵹���ڶ��У�Ҳ����end subǰ������end����word����ǿ�˳����Ͳ����ٳ����Ǹ������ˡ�

https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.word.wdbuiltinstyle.aspx#Anchor_1


�ġ�word����������ô�ж϶����Ƿ��ڱ������

Sub ѭ����������()
'��������ڱ���У����Ϊ��ɫ��������ɫ
    Dim i As Paragraph
    For Each i In ActiveDocument.Paragraphs
        If i.Range.Information(wdWithInTable) = True Then i.Range.Font.Color = wdColorRed Else i.Range.Font.Color = wdColorBlue
    Next
End Sub


wdWrapInline 7 ����״Ƕ�뵽�����С�
wdWrapNone 3 ����״��������ǰ�档�����  wdWrapFront ��
wdWrapSquare 0 ʹ���ֻ�����״��������״����һ��������
wdWrapThrough 2 ʹ���ֻ�����״��
wdWrapTight 1 ʹ���ֽ��ܵػ�����״��
wdWrapTopBottom 4 �����ַ�����״���Ϸ����·���
wdWrapBehind 5 ����״�������ֺ��档
wdWrapFront 6 ����״��������ǰ�档


������ͼƬ����ʽ��ֱ��ת�����ȴ�����

    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior


