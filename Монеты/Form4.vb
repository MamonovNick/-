Public Class Form4
    Private SumC As Single = 0

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Подразделения". При необходимости она может быть перемещена или удалена.
        Me.SecDA.Fill(Me.МонетыDataSet.Подразделения)
    End Sub

    Private Sub Form4_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.SecDA.Update(Me.МонетыDataSet.Подразделения)
        For i As Integer = 0 To DataGridView1.RowCount() - 1 Step i + 1
            SumC += IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), 0, DataGridView1.Rows(i).Cells(3).Value)
        Next i
        If SumC <> 1 Then
            Dim res As DialogResult = MsgBox("Сумма коэффициентов не равна 1 и составляет " + CStr(SumC) + " Исправить?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Ошибка")
            If res = DialogResult.Yes Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MsgBox("Неправильные данные", MsgBoxStyle.Critical, "Ошибка")
    End Sub
End Class