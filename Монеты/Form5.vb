Public Class Form5
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Юридические_лица". При необходимости она может быть перемещена или удалена.
        Me.Юридические_лицаTableAdapter.Fill(Me.МонетыDataSet.Юридические_лица)

    End Sub

    Private Sub Form5_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Юридические_лицаTableAdapter.Update(Me.МонетыDataSet.Юридические_лица)
    End Sub
End Class