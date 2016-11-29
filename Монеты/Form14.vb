Public Class Form14

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Выбор файла базы данных
        OpenFileDialog1.FileName = TextBox1.Text
        If (OpenFileDialog1.ShowDialog() = DialogResult.Cancel) Then
            Return
        End If
        Dim str As String = OpenFileDialog1.FileName
        TextBox1.Text = str
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FolderBrowserDialog1.SelectedPath = TextBox2.Text
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.Cancel) Then
            Return
        End If
        Dim str As String = FolderBrowserDialog1.SelectedPath
        TextBox2.Text = str
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        AboutBox1.ShowDialog()
    End Sub
End Class