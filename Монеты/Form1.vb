Public Class MainForm
    Private FirstOpen As Boolean = True
    Private OperationBool As Boolean = False 'Открытие подменю операций
    Private TabNum As Int16 = -1 'Номер текущей таблицы и пункта меню
    Private PrevTabNum As Int16 = -1 'Номер предыдущей таблицы
    Private MenuPosNum As Int16 = 6 ' Количество позиций меню
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private delCommand As OleDb.OleDbCommand ' переменная для запроса удаления
    Private updCommand As OleDb.OleDbCommand ' переменная для запроса апдейта
    Private insCommand As OleDb.OleDbCommand ' переменная для запроса вставки

    Private comboColumn2 As New DataGridViewComboBoxColumn

    Private DA As New OleDb.OleDbDataAdapter ' адаптер
    Private DA2 As New OleDb.OleDbDataAdapter ' вспомогательный адаптер

    Private bs1 As New BindingSource() 'Переменная bindingsourse
    Private tbt As New DataTable() ' переменная таблица для вывода в грид
    Private tbt2 As New DataTable() ' переменная табоица для проверки задвоений
    'Private tbtForComboBox As New DataTable() ' переменная таблица для комбобокса
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
        Try
            DA.Update(tbt)
        Catch e As System.Data.DBConcurrencyException
            MsgBox("Изменения не были сохранены!", MsgBoxStyle.Critical, "Внимание")
        End Try
    End Sub

    Private Function GetTableForCmb() As DataTable
        Dim da As OleDb.OleDbDataAdapter
        Dim tbt As New DataTable
        Dim sqlcom As New OleDb.OleDbCommand("SELECT Подразделения.Наименование 
FROM Подразделения 
ORDER BY Подразделения.ВидУчастника DESC , Подразделения.Наименование", Con)
        da = New OleDb.OleDbDataAdapter(sqlcom)
        da.Fill(tbt)
        Return tbt
    End Function

    Private Sub Clear_Form()
        'Отмена выделения пунктов меню
        ToolStripButton1.Checked = False
        ToolStripButton3.Checked = False
        ToolStripButton4.Checked = False
        ToolStripButton5.Checked = False
        ToolStripButton6.Checked = False
        ToolStripButton10.Checked = False
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Panel1.Visible = False
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
        Clear_Form()
        ToolStripButton5.Checked = True
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
        If Not FirstOpen Then
            Update_table()
        End If
        Select Case TabNum
            Case 1
                'загрузить таблицу
                TabNum = 7
                tbt.Reset()
                tbt = New DataTable
                DA = New OleDb.OleDbDataAdapter
                SqlCom = New OleDb.OleDbCommand("TRANSFORM IIf(Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))=0,"""",Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))) AS Выражение2 
SELECT IIf(IsNull([Для свода заявок].Монета), NULL, IIf((Монеты.[Валюта] = ""российск. рубль""),""росс."",""иностр."")) AS Вид, Монеты.Металл, Монеты.Масса, [Для свода заявок].Монета, [Для свода заявок].Состояние
FROM     (Монеты RIGHT OUTER JOIN
                      (SELECT Подразделение, ВидУчастника, [Каталожный номер], Монета, Состояние, Количество, Исполнено
                       FROM      Заявки
                       WHERE   (Дата >= ?) AND (Закрыто = False) AND ([Вид операции] = 'Выдача')
                       UNION ALL
                       SELECT Наименование, ВидУчастника, NULL AS [Каталожный номер], NULL AS Монета, NULL AS Состояние, NULL AS Количество, NULL AS Исполнено
                       FROM     Подразделения
                       WHERE  (Активный = True)) [Для свода заявок] ON Монеты.[Каталожный номер] = [Для свода заявок].[Каталожный номер])
GROUP BY IIf(IsNull([Монета]),Null,IIf([Валюта]=""российск. рубль"",""росс."",""иностр."")), Монеты.Металл, Монеты.Масса, Монеты.Год, [Для свода заявок].[Каталожный номер], [Для свода заявок].Монета, [Для свода заявок].Состояние
ORDER BY IIf(IsNull([Монета]),Null,IIf([Валюта]=""российск. рубль"",""росс."",""иностр."")) DESC , Монеты.Металл DESC , Монеты.Масса, Монеты.Год, IIf([ВидУчастника]=""терр. банк"","" ""+[Подразделение],IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]))
PIVOT IIf([ВидУчастника]=""терр. банк"","" ""+[Подразделение],IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]));", Con)
                DA.SelectCommand = SqlCom
                DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
                DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

                DA.Fill(tbt)
                DataGridView1.DataSource = tbt
                DataGridView1.ReadOnly = True
                Panel1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button1.Text = "Назад"
            Case 4
                MsgBox("Данное действие еще не реализовано!", MsgBoxStyle.Critical, "Ошибка")
            Case 7, 8, 9
                ToolStripButton1_Click(sender, e)
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
        Clear_Form()
        ToolStripButton6.Checked = True
        tbt.Reset()
        tbt = New DataTable
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
        Clear_Form()
        ToolStripButton10.Checked = True
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

    Private Sub Table1_Make()
        If TabNum <> 1 Then
            Return
        End If
        tbt.Reset() 'очищаем таблицу
        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        SqlCom = New OleDb.OleDbCommand("SELECT Дата, Номер, [Вид операции], ВидУчастника, Подразделение, [Каталожный номер], Монета, Количество, Состояние, Исполнено, Закрыто 
FROM Заявки
WHERE ((Дата >= @Дата)" + IIf(CheckBox1.Checked, " AND (Закрыто = False)", "") + IIf(RadioButton2.Checked, " AND (Подразделение = """ + CStr(ComboBox1.Text) + """)", "") + ")", Con)
        DA.SelectCommand = SqlCom
        DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

        delCommand = New OleDb.OleDbCommand("DELETE FROM Заявки 
WHERE ((Дата = @Дата) OR Дата IS NULL) AND ((Номер = @Номер) OR Номер IS NULL) AND (([Вид операции] = @Операция) OR [Вид операции] IS NULL) AND ((Подразделение = @Подразделение) OR Подразделение IS NULL) AND (([Каталожный номер] = @КатНомер) OR [Каталожный номер] IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.DeleteCommand.Parameters.Add("@Операция", OleDb.OleDbType.VarChar, 17, "Вид операции")
        DA.DeleteCommand.Parameters.Add("@Подразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.DeleteCommand.Parameters.Add("@КатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")

        updCommand = New OleDb.OleDbCommand("UPDATE Заявки 
SET Дата = qДата, Номер = qНомер, [Вид операции] = qОперация, ВидУчастника = qУчастник, Подразделение = qПодразделение, [Каталожный номер] = qКатНомер, Монета = qМонета, Количество = qКолво, Состояние = qСостояние, Исполнено = qИсполнено, Закрыто = qЗакрыто 
WHERE ((Дата = Дата_Orig) AND ((Номер = Номер_Orig) OR Номер IS NULL) AND ([Вид операции] = Операция_Orig) AND (Подразделение = Подразделение_Orig) AND (([Каталожный номер] = КатНомер_Orig) OR [Каталожный номер] IS NULL))", Con)
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
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'Обработка пункта меню "Заявки"
        'Обновление предыдущей таблицы, в случае если она была
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        PrevTabNum = TabNum
        TabNum = 1 'номер текущей таблицы
        Clear_Form() 'отмена выделения пункта меню

        ComboBox1.DisplayMember = "Наименование"
        ComboBox1.DataSource = GetTableForCmb()
        ComboBox1.SelectedIndex = 35

        ToolStripButton1.Checked = True 'выделяем текущий пункт меню

        Table1_Make() ' Заполнение адаптера и таблицы и грида

        'колонка с датой
        Dim DateColumn As New CalendarColumn()
        DateColumn.DataPropertyName = "Дата"
        DateColumn.Name = "Дата"

        Dim oldColIndex As Int32 = DataGridView1.Columns("Дата").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, DateColumn)

        'колонка выдача прием
        Dim comboColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
        comboColumn.Items.AddRange("приём", "выдача")
        comboColumn.Name = "Вид операции"
        comboColumn.DataPropertyName = "Вид операции"
        comboColumn.SortMode = DataGridViewColumnSortMode.Automatic

        oldColIndex = DataGridView1.Columns("Вид операции").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, comboColumn)
        DataGridView1.ReadOnly = False

        'колонка вид участника
        'Dim comboColumn2 As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
        comboColumn2 = New DataGridViewComboBoxColumn()

        comboColumn2.Items.AddRange("Москва", "терр. банк", "управление")
        comboColumn2.Name = "ВидУчастника"
        comboColumn2.DataPropertyName = "ВидУчастника"
        comboColumn2.SortMode = DataGridViewColumnSortMode.Automatic

        oldColIndex = DataGridView1.Columns("ВидУчастника").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, comboColumn2)

        TableLayoutPanel1.SetRowSpan(DataGridView1, 2)

        Button1.Visible = True
        Button1.Enabled = True
        Button2.Visible = True
        Button2.Enabled = True
        Button3.Visible = True
        Button3.Enabled = True
        Button4.Visible = True
        Button4.Enabled = True
        Button1.Text = """Cвод для ""хвостов"
        Button2.Text = "Свод для новых"
        Button3.Text = "Закрыть заявки"
        Button4.Text = "Искать повторы"

        Panel1.Visible = True
        Panel1.Enabled = True

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'tbt2
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'tbt3
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        Select Case TabNum
            Case 1
                'If e.ColumnIndex = DataGridView1.Columns("ВидУчастника").Index Then
                '    comboColumn2.Items.Add(tbt.Rows(e.RowIndex)(e.ColumnIndex))
                'End If
            Case 2

        End Select
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Подразделения". При необходимости она может быть перемещена или удалена.
        Me.SecDA.Fill(Me.МонетыDataSet.Подразделения)
        TableLayoutPanel1.SetRowSpan(DataGridView1, 3)
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Not FirstOpen Then
            Update_table()
        End If
        Table1_Make()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Not FirstOpen Then
            Update_table()
        End If
        Table1_Make()
        If RadioButton1.Checked Then
            ComboBox1.Enabled = False
        Else
            ComboBox1.Enabled = True
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not FirstOpen Then
            Update_table()
        End If
        Table1_Make()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Not FirstOpen Then
            Update_table()
        End If
        Table1_Make()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not FirstOpen Then
            Update_table()
        End If
        Select Case TabNum
            Case 1
                PrevTabNum = TabNum
                TabNum = 8
                tbt.Reset()
                tbt = New DataTable
                DA = New OleDb.OleDbDataAdapter
                SqlCom = New OleDb.OleDbCommand("TRANSFORM IIf(Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))=0,"""",Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))) AS Выражение1
SELECT IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]) AS Филиал
FROM Монеты RIGHT JOIN (SELECT Подразделение, ВидУчастника, [Каталожный номер], Монета, Состояние, Количество, Исполнено
                       FROM      Заявки
                       WHERE   (Дата >= ?) AND (Закрыто = False) AND ([Вид операции] = 'Выдача')
                       UNION ALL
                       SELECT Наименование, ВидУчастника, NULL AS [Каталожный номер], NULL AS Монета, NULL AS Состояние, NULL AS Количество, NULL AS Исполнено
                       FROM     Подразделения
                       WHERE  (Активный = True)) [Для свода заявок] ON Монеты.[Каталожный номер] = [Для свода заявок].[Каталожный номер]
GROUP BY IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]), IIf([ВидУчастника]=""терр. банк"",1,IIf([ВидУчастника]=""Москва"" And [Подразделение]<>""Московский банк"",2,IIf([Подразделение]=""Московский банк"",3,4)))
ORDER BY IIf([ВидУчастника]=""терр. банк"",1,IIf([ВидУчастника]=""Москва"" And [Подразделение]<>""Московский банк"",2,IIf([Подразделение]=""Московский банк"",3,4)))
PIVOT [Монета]+"", ""+[Состояние];", Con)
                DA.SelectCommand = SqlCom
                DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
                DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

                DA.Fill(tbt)
                DataGridView1.DataSource = tbt
                DataGridView1.ReadOnly = True
                Panel1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button1.Text = "Назад"
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not FirstOpen Then
            Update_table()
        End If

        Select Case TabNum
            Case 1
                ' Закрыть заявки
                If Dialog2.ShowDialog() = DialogResult.OK Then
                    Dim updCommand As New OleDb.OleDbCommand("UPDATE Заявки 
SET Закрыто = True 
WHERE ((Закрыто = False) AND (Дата Between ? AND ?))", Con)
                    DA.UpdateCommand = updCommand
                    DA.UpdateCommand.Parameters.Clear()
                    DA.UpdateCommand.Parameters.Add("@Дата1", OleDb.OleDbType.Date, -1, "Дата")
                    DA.UpdateCommand.Parameters(0).Value = Class1.getDate(1)
                    DA.UpdateCommand.Parameters.Add("@Дата2", OleDb.OleDbType.Date, -1, "Дата")
                    DA.UpdateCommand.Parameters(1).Value = Class1.getDate(0)

                    Con.Close()
                    DA.UpdateCommand.Connection = Con
                    DA.UpdateCommand.Connection.Open()
                    DA.UpdateCommand.ExecuteNonQuery()

                    ToolStripButton1_Click(sender, e)
                End If
        End Select
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not FirstOpen Then
            Update_table()
        End If
        Select Case TabNum
            Case 1
                'Искать повторы
                tbt2.Reset()
                SqlCom = New OleDb.OleDbCommand("SELECT * FROM [Проверка повторов в заявках]", Con)
                DA2.SelectCommand = SqlCom

                DA2.Fill(tbt2)
                If tbt2.Rows.Count = 0 Then
                    MsgBox("Повторов не найдено",, "Поиск повторяющихся строк")
                Else
                    MsgBox("Повторы найдены!", MsgBoxStyle.Exclamation, "Поиск повторяющихся строк")
                    DataGridView1.DataSource = tbt2

                    TabNum = 9
                    DataGridView1.ReadOnly = True
                    Panel1.Enabled = False
                    Button2.Enabled = False
                    Button3.Enabled = False
                    Button4.Enabled = False
                    Button1.Text = "Назад"
                End If
        End Select
    End Sub
End Class