Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Dialog3
    Private tbt As New DataTable() ' переменная таблица для вывода в грид
    Private RealClose As Boolean = True

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim a As Double
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Select Case Class1.getPriceType
            Case 3
                If TextBox1.Text <> "" Then
                    Try
                        a = CDbl(TextBox1.Text)
                    Catch ex As Exception
                        a = Nothing
                    End Try

                    If a <> Nothing Then
                        Class1.setPrice(a)
                        Me.Close()
                    Else
                        MsgBox("Неверно введено число", MsgBoxStyle.Critical)
                        RealClose = False
                    End If
                Else
                    Class1.setPrice(Nothing)
                    Class1.setSelectedIndex(GridView1.FocusedRowHandle)
                    Me.Close()
                End If
            Case 4
                Class1.setPrice(Nothing)
                Class1.setSelectedIndex(GridView1.FocusedRowHandle)
                Me.Close()
        End Select
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbt.Reset()
        tbt.Columns.Clear()
        GridView1.Columns.Clear()
        GridView1.OptionsView.ColumnAutoWidth = True
        Select Case Class1.getPriceType
            Case 1
                tbt = Module1.GetTablePrices(Class1.GetCat(), Class1.getDate(1))
                Label1.Enabled = False
                TextBox1.Enabled = False
            Case 2
                tbt = Module1.GetTablePricesForCond(Class1.GetCat(), Class1.getDate(1), Class1.getStoragePlace())
                Label1.Enabled = False
                TextBox1.Enabled = False
            Case 3
                tbt = Module1.GetTablePricesForOperations(Class1.GetCat, Class1.getDate(1), Class1.getStoragePlace, Class1.getFullPrice(), Class1.getState())
                Label1.Enabled = True
                TextBox1.Enabled = True
            Case 4
                tbt = Module1.GetTablePricesForOperations(Class1.GetCat, Class1.getDate(1), Class1.getStoragePlace, Class1.getFullPrice(), Class1.getState())
                Label1.Enabled = False
                TextBox1.Enabled = False
        End Select
        GridControl1.DataSource = tbt
        GridView1.OptionsBehavior.Editable = False
    End Sub

    Private Sub Dialog3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not RealClose Then
            e.Cancel = True
            RealClose = Not RealClose
        End If
    End Sub
End Class
