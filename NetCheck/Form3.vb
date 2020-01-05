Public Class Form3

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.globe
        Me.WebBrowser1.DocumentText = String.Copy(My.Resources.NetCheck)
    End Sub

End Class