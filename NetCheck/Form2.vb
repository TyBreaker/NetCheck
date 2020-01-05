Public Class Form2

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Cancel Settings:
        Me.Close()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Preload default settings
        Me.Icon = My.Resources.globe
        Me.TextBox1.Text = My.Forms.Form1.pingDatabase
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' choose database location:
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

End Class