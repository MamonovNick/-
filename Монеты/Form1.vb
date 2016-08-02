Public Class MainForm
    Private FirstOpen As Boolean = True
    Private OperationBool As Boolean = False 'Открытие подменю операций
    Private TabNum As Int16 = -1 'Номер текущей таблицы и пункта меню
    Private MenuPosNum As Int16 = 6 ' Количество позиций меню
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private delCommand As OleDb.OleDbCommand ' переменная для запроса удаления
    Private updCommand As OleDb.OleDbCommand ' переменная для запроса апдейта
    Private insCommand As OleDb.OleDbCommand ' переменная для запроса вставки

    Private DA As New OleDb.OleDbDataAdapter ' адаптер для перемещения монет

    Private bs1 As New BindingSource() 'Переменная bindingsourse
    Private tbt As New DataTable() ' переменная таблица для вывода в грид

    Private Sub Update_table()
        DA.Update(tbt)
    End Sub

    Private Sub Clear_Checked()
        ToolStripButton1.Checked = False
        ToolStripButton3.Checked = False
        ToolStripButton4.Checked = False
        ToolStripButton5.Checked = False
        ToolStripButton6.Checked = False
        ToolStripButton10.Checked = False
    End Sub

    Private Sub МонетыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МонетыToolStripMenuItem.Click
        G_coins.Show()
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ПодразделенияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПодразделенияToolStripMenuItem.Click
        Form4.Show()
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'Only for debugging
        'For tech only 
        'Need to be deleted in release ver
        'Form7.Show()
        'Form7.Activate()
    End Sub

    Private Sub КонтрагентыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КонтрагентыToolStripMenuItem.Click
        Form5.Show()
    End Sub

    Private Sub ВидыВалютToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВидыВалютToolStripMenuItem.Click
        Form6.Show()
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If OperationBool Then
            ToolStripButton3.Visible = False
            ToolStripButton4.Visible = False
        Else
            ToolStripButton3.Visible = True
            ToolStripButton4.Visible = True
        End If
        OperationBool = Not OperationBool
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        TabNum = 4
        Button1.Visible = True
        Button1.Text = "Сделать спецификацию"
        Clear_Checked()
        ToolStripButton5.Checked = True
        tbt.Clear()
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT Завершено, Дата, [Каталожный номер], Монета, Количество, Цена, ВхНДС as НДС, Состояние, Дефекты, МестоХраненияСтарое as [Старое место хранения], МестоХраненияНовое as [Новое место хранения], Спецификация 
FROM [Перемещение между хранилищами]", Con)

        DA.SelectCommand = SqlCom
        DA.Fill(tbt)
        bs1.DataSource = tbt
        DataGridView1.DataSource = bs1
        TableLayoutPanel1.SetRowSpan(DataGridView1, 2)
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).ReadOnly = True
        DataGridView1.Columns(9).ReadOnly = True

    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing


        'Update table!!!!
        'Update_table()
        'заполнить

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case TabNum
            Case 4
                MsgBox("Данное действие еще не реализовано!", MsgBoxStyle.Critical, "Ошибка")
        End Select
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        TabNum = 5
        Button1.Visible = False
        Clear_Checked()
        ToolStripButton6.Checked = True
        tbt.Clear()
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT * 
FROM [Изменение состояния]", Con)

        DA.SelectCommand = SqlCom
        DA.Fill(tbt)
        bs1.DataSource = tbt
        DataGridView1.DataSource = bs1
        TableLayoutPanel1.SetRowSpan(DataGridView1, 3)
        'DataGridView1.Columns(6).ReadOnly = True
        'DataGridView1.Columns(7).ReadOnly = True
        'DataGridView1.Columns(8).ReadOnly = True
        'DataGridView1.Columns(9).ReadOnly = True
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        TabNum = 6
        Button1.Visible = False
        Clear_Checked()
        ToolStripButton10.Checked = True
        tbt.Clear()
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT * 
FROM [Приобретение монет ТБ в ЦБ]", Con)

        DA.SelectCommand = SqlCom
        DA.Fill(tbt)
        bs1.DataSource = tbt
        DataGridView1.DataSource = bs1
        TableLayoutPanel1.SetRowSpan(DataGridView1, 3)
        'DataGridView1.Columns(6).ReadOnly = True
        'DataGridView1.Columns(7).ReadOnly = True
        'DataGridView1.Columns(8).ReadOnly = True
        'DataGridView1.Columns(9).ReadOnly = True
    End Sub
End Class