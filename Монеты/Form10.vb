Public Class Form10
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Индивидуальные_ставки_условных". При необходимости она может быть перемещена или удалена.
        Me.Индивидуальные_ставки_условныхTableAdapter.Fill(Me.МонетыDataSet.Индивидуальные_ставки_условных)

    End Sub
End Class