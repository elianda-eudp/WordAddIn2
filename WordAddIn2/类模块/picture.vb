Public Class Picture
    '    Option Explicit
    Public picture_format As Word.Shape
    Public picture_parag_format As Word.ParagraphFormat





    '    Set paragreph_format = New Word.ParagraphFormat


    Public Function Set_format(ByVal style As Word.Shape, ByVal para As Word.ParagraphFormat) As Boolean
        picture_format = style
        picture_parag_format = para


        '设置表格式
        With picture_format
            If .Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                .ConvertToInlineShape()
            End If
        End With

        picture_parag_format.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        Return True

    End Function







End Class
