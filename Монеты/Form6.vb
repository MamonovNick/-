Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Список_валют". При необходимости она может быть перемещена или удалена.
        Me.Список_валютTableAdapter.Fill(Me.МонетыDataSet.Список_валют)

    End Sub

    Private Sub Form6_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Список_валютTableAdapter.Update(Me.МонетыDataSet.Список_валют)
    End Sub
End Class