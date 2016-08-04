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
    'Private tbt2 As New DataTable() ' переменная таблица для вывода в грид
    'Private tbt3 As New DataTable() ' переменная таблица для вывода в грид
    'Private tbt4 As New DataTable() ' переменная таблица для вывода в грид
    'Private tbt5 As New DataTable() ' переменная таблица для вывода в грид
    'Private tbt6 As New DataTable() ' переменная таблица для вывода в грид


    Private Sub Update_table()
        'Select Case TabNum
        '    Case 1
        '        DA.Update(tbt)
        '    Case 2
        '        DA.Update(tbt2)
        '    Case 3
        '        DA.Update(tbt3)
        '    Case 4
        '        DA.Update(tbt4)
        '    Case 5
        '        DA.Update(tbt5)
        '    Case 6
        '        DA.Update(tbt6)
        'End Select
        DA.Update(tbt)
    End Sub

    Private Sub Clear_Checked()
        'Отмена выделения пунктов меню
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
        'раскрытие пункта меню для опреций
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
        'tbt.Dispose()
        'DA.Dispose()
        tbt.Reset()
        tbt = New DataTable
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
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
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
        'tbt5
        'tbt.Dispose()
        'DA.Dispose()
        tbt.Reset()
        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT * 
FROM [Изменение состояния]", Con)

        DA.SelectCommand = SqlCom
        DA.Fill(tbt)
        bs1.DataSource = tbt
        'tbt5.Reset()
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
        'tbt.Dispose()
        'DA.Dispose()
        'bs1.Dispose()
        tbt.Reset()
        tbt = New DataTable
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

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'Обработка пункта меню "Заявки"
        'Обновление предыдущей таблицы, в случае если она была
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        TabNum = 1 'номер текущей таблицы
        Clear_Checked() 'отмена выделения пункта меню
        ToolStripButton1.Checked = True 'выделяем текущий пункт меню
        tbt.Reset() 'очищаем таблицу
        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT Дата, Номер, [Вид операции], ВидУчастника, Подразделение, [Каталожный номер], Монета, Количество, Состояние, Исполнено, Закрыто FROM Заявки", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM Заявки WHERE (Дата = @Дата) AND (Номер = @Номер) AND ([Вид операции] = @Операция) AND (Подразделение = @Подразделение) AND ([Каталожный номер] = @КатНомер) AND (Номер = @Номер)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.DeleteCommand.Parameters.Add("@Операция", OleDb.OleDbType.VarChar, 17, "Вид операции")
        DA.DeleteCommand.Parameters.Add("@Подразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.DeleteCommand.Parameters.Add("@КатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")

        updCommand = New OleDb.OleDbCommand("UPDATE Заявки 
SET Дата = qДата, Номер = qНомер, [Вид операции] = qОперация, ВидУчастника = qУчастник, Подразделение = qПодразделение, [Каталожный номер] = qКатНомер, Монета = qМонета, Количество = qКолво, Состояние = qСостояние, Исполнено = qИсполнено, Закрыто = qЗакрыто 
WHERE ((Дата = Дата_Orig) AND (Номер = Номер_Orig) AND ([Вид операции] = Операция_Orig) AND (Подразделение = Подразделение_Orig) AND ([Каталожный номер] = КатНомер_Orig))", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("qДата", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Add("qНомер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.UpdateCommand.Parameters.Add("qОперация", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.UpdateCommand.Parameters.Add("qУчастник", OleDb.OleDbType.VarChar, 15, "ВидУчастника")
        DA.UpdateCommand.Parameters.Add("qПодразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.UpdateCommand.Parameters.Add("qКатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Add("qМонета", OleDb.OleDbType.VarChar, 60, "Монета")
        DA.UpdateCommand.Parameters.Add("qКолво", OleDb.OleDbType.Integer, -1, "Количество")
        DA.UpdateCommand.Parameters.Add("qСостояние", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.UpdateCommand.Parameters.Add("qИсполнено", OleDb.OleDbType.Integer, -1, "Исполнено")
        DA.UpdateCommand.Parameters.Add("qЗакрыто", OleDb.OleDbType.Boolean, -1, "Закрыто")
        DA.UpdateCommand.Parameters.Add("Дата_Orig", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Item(11).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Номер_Orig", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.UpdateCommand.Parameters.Item(12).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Операция_Orig", OleDb.OleDbType.VarChar, 17, "Вид операции")
        DA.UpdateCommand.Parameters.Item(13).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Подразделение_Orig", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.UpdateCommand.Parameters.Item(14).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("КатНомер_Orig", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(15).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO Заявки
                          (Дата, Номер, [Вид операции], ВидУчастника, Подразделение, [Каталожный номер], Монета, Количество, Состояние, Исполнено, Закрыто)
        VALUES (@Дата, @Номер, @Операция, @Участник, @Подразделение, @КатНомер, @Монета, @Колво, @Состояние, @Исполнено, @Закрыто)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.InsertCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.InsertCommand.Parameters.Add("@Операция", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.InsertCommand.Parameters.Add("@Участник", OleDb.OleDbType.VarChar, 15, "ВидУчастника")
        DA.InsertCommand.Parameters.Add("@Подразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.InsertCommand.Parameters.Add("@КатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.InsertCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета")
        DA.InsertCommand.Parameters.Add("@Колво", OleDb.OleDbType.Integer, -1, "Количество")
        DA.InsertCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.InsertCommand.Parameters.Add("@Исполнено", OleDb.OleDbType.Integer, -1, "Исполнено")
        DA.InsertCommand.Parameters.Add("@Закрыто", OleDb.OleDbType.Boolean, -1, "Закрыто")

        DA.Fill(tbt)
        bs1.DataSource = tbt
        DataGridView1.DataSource = bs1

        Dim DateColumn As New CalendarColumn()
        DateColumn.DataPropertyName = "Дата"
        DateColumn.Name = "Дата"

        Dim oldColIndex As Int32 = DataGridView1.Columns("Дата").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, DateColumn)

        TableLayoutPanel1.SetRowSpan(DataGridView1, 3)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'tbt2
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'tbt3
    End Sub

End Class