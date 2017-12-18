
Imports WordAddIn2

Public Class FormMain
    Public wd As Word.Application
    Public BeginPageNumber As String
    Public EndPageNumber As String
    Public docpath As String
    Public file_name As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form As New FormMain

        form.docpath = TextBox1.Text
        form.file_name = TextBox2.Text
        form.BeginPageNumber = TextBox3.Text
        form.EndPageNumber = TextBox4.Text
        Dim wd As Word.Application
        wd = Globals.ThisAddIn.Application
        'MsgBox（"被检查文件名:" & form.docpath & "\" & form.file_name）
        Call main_handle.Main_handle(wd, form)

    End Sub

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "E:\邮储理财项目组工作\新技术\文档检查工具"
        TextBox2.Text = "被检查文档示例.docx"
        TextBox3.Text = "1"
        TextBox4.Text = "9999"
    End Sub


End Class