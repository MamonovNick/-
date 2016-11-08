Imports System.Windows.Forms

Public Class Dialog3
    Private tbt As New DataTable() ' переменная таблица для вывода в грид

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Class1.setSelectedIndex(GridView1.FocusedRowHandle)
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GridView1.OptionsView.ColumnAutoWidth = True
        tbt = Module1.GetTablePrices(Class1.GetCat, Class1.getDate(1))
        GridControl1.DataSource = tbt
        GridView1.OptionsBehavior.Editable = False
    End Sub
End Class
