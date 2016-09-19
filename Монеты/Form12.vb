Public Class Form12
    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Цены". При необходимости она может быть перемещена или удалена.
        Me.ЦеныTableAdapter.Fill(Me.МонетыDataSet.Цены)

    End Sub
End Class