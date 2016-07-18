Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Подразделения". При необходимости она может быть перемещена или удалена.
        Me.SecDA.Fill(Me.МонетыDataSet.Подразделения)
    End Sub

    Private Sub Form4_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.SecDA.Update(Me.МонетыDataSet.Подразделения)
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MsgBox("Неправильные данные", MsgBoxStyle.Critical, "Ошибка")
    End Sub
End Class