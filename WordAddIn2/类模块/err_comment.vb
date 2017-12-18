Public Class Err_comment
    Public para As Word.Paragraph
    Public table As Word.Table

    Public comments As Word.Comments
    Public comment As Word.Comment

    Public comment_str As String

    Public Function Set_comment(ByVal para_p As Word.Paragraph, ByVal p_comment As Word.Comments, ByVal str As String) As Boolean
        para = para_p
        comments = p_comment
        comment_str = str

        With para

            If (comment_str <> "") Then
                comment = comments.Add(.Range, comment_str)
                comment.Author = "coin wo-wo"
                comment.Range.Text = comment_str
            End If
        End With

        Return True
    End Function

    Public Function Table_set_comment(ByVal para_p As Word.Table, ByVal p_comment As Word.Comments, ByVal str As String) As Boolean
        table = para_p
        comments = p_comment
        comment_str = str

        With table

            If (comment_str <> "") Then
                comment = comments.Add(.Range, comment_str)
                comment.Author = "coin wo-wo"
                comment.Range.Text = comment_str
            End If
        End With
        Return True

    End Function


    Public Function Range_set_comment(ByVal para_p As Word.Range, ByVal p_comment As Word.Comments, ByVal str As String) As Boolean
        comments = p_comment
        comment_str = str

        If (comment_str <> "") Then
            comment = comments.Add(para_p, comment_str)
            comment.Author = "coin wo-wo"
            comment.Range.Text = comment_str
        End If
        Return True
    End Function

End Class
