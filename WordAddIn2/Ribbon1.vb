Imports Microsoft.Office.Tools.Ribbon
Imports WordAddIn2


Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim a As Object
        a = New FormMain

        a.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim wd As Word.Application
        wd = Globals.ThisAddIn.Application
        Call list_template.Listtemplate(wd)
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim a As Object
        a = New Form1

        a.Show()
    End Sub
End Class
