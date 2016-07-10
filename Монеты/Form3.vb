Public Class Form3
    Private Pic_location As String = "D:\Монеты-Access\Фотографии монет\" 'расположение фотографий монет
    Private FirstOpen As Boolean = True
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private SqlComForCmb As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private DAh As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private DAh2 As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private DAt As New OleDb.OleDbDataAdapter
    Private DApr As OleDb.OleDbDataAdapter
    Private DAforCmb As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private table As DataTable = New DataTable()
    Private tableForCmb As DataTable = New DataTable()
    Private tableForCmb2 As DataTable = New DataTable()
    Private tableForAlternateMonets As DataTable = New DataTable() ' Таблица для загрузки полей 
    Private bs1 As New BindingSource() 'Переменная bindingsourse
    Private tbt As New DataTable() ' переменная таблица для вывода в грид
    Private tableForCheck As New DataTable() ' таблица для проверки наборов
    Private rus_forein As String
    Private state As Boolean = False

    Private Sub Update_table()
        DAt.Update(tbt)
    End Sub

    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        G_coins.Show()
    End Sub

    Private Sub GetTable()
        SqlCom = New OleDb.OleDbCommand("SELECT Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2) AS [Кач-во], Монеты.Масса, Монеты.Качество, Монеты.Серия, Монеты.ФотоАверс, Монеты.ФотоРеверс 
FROM     (Монеты INNER JOIN
                  [Состав наборов] ON Монеты.[Каталожный номер] = [Состав наборов].КатНомерНабора)
GROUP BY Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2), Монеты.Масса, Монеты.Качество, Монеты.Серия, Монеты.Год, Монеты.ФотоАверс, Монеты.ФотоРеверс, 
                  Монеты.Счётчик
ORDER BY Монеты.Год, Монеты.Счётчик", Con) ' Указываем строку запроса и привязываем к соединению

        DAh = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh.Fill(table) ' Заполняем таблицу результатми
    End Sub

    Private Sub GetTableForCmb()
        SqlComForCmb = New OleDb.OleDbCommand("SELECT Монеты.[Каталожный номер] ++ ' ' ++ Монеты.Краткое ++ ' ' ++ Монеты.Металл as txt FROM (SELECT Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2) AS [Кач-во], Монеты.Масса, Монеты.Качество, Монеты.Серия
FROM     (Монеты INNER JOIN
                  [Состав наборов] ON Монеты.[Каталожный номер] = [Состав наборов].КатНомерНабора)
GROUP BY Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2), Монеты.Масса, Монеты.Качество, Монеты.Серия, Монеты.Год, 
                  Монеты.Счётчик
ORDER BY Монеты.Год, Монеты.Счётчик)", Con) ' Указываем строку запроса и привязываем к соединению
        DAforCmb = New OleDb.OleDbDataAdapter(SqlComForCmb) 'Через адаптер получаем результаты запроса
        DAforCmb.Fill(tableForCmb) ' Заполняем таблицу результатми
    End Sub

    Private Sub GetTableforcmb2()
        SqlCom = New OleDb.OleDbCommand("SELECT Монеты.[Каталожный номер] ++ ' ' ++  Монеты.Краткое ++ ' ' ++ Монеты.Металл as txt2, Монеты.[Каталожный номер], [Монеты].[Краткое], Монеты.Металл, Монеты.Номинал, Монеты.Валюта, [Монеты].[Качество], Монеты.Масса, Монеты.Год, Монеты.Проба, Монеты.Диаметр, Монеты.Тираж, Монеты.Статус, Монеты.Серия, Монеты.ФотоАверс, Монеты.ФотоРеверс 
FROM Монеты 
ORDER BY IIf([Монеты].[Валюта]='российск. рубль',2,1), Монеты.Металл DESC , Монеты.Масса, Монеты.Валюта, Монеты.Номинал, Монеты.Год, Монеты.Краткое;", Con) ' Указываем строку запроса и привязываем к соединению

        DAh2 = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh2.Fill(tableForCmb2) ' Заполняем таблицу результатми
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetTable()
        GetTableForCmb()
        GetTableforcmb2()
        ComboBox1.DisplayMember = "txt" 'Столбец для отображения в combobox
        ComboBox1.DataSource = tableForCmb ' Загрузка таблицы для отображения в combobox
        ComboBox1_SelectionChangeCommitted(sender, e) ' Фиксация изменения строки
        ComboBox2.DisplayMember = "txt2" 'Столбец для отображения в combobox
        ComboBox2.DataSource = tableForCmb2 ' Загрузка таблицы для отображения в combobox

        Dim delCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand("DELETE FROM [Состав наборов] WHERE (КатНомерНабора = @КатНомерНабора) AND (КатНомерМонеты = @КатНомерМонеты)", Con)
        DAt.DeleteCommand = delCommand
        DAt.DeleteCommand.Parameters.Add("@КатНомерНабора", OleDb.OleDbType.VarChar, 9, "Набор")
        DAt.DeleteCommand.Parameters.Add("@КатНомерМонеты", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        Dim updCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand("UPDATE [Состав наборов] SET КатНомерМонеты = @КатНомерМонеты WHERE (КатНомерНабора = @КатНомерНабора)", Con)
        DAt.UpdateCommand = updCommand
        DAt.UpdateCommand.Parameters.Add("@КатНомерНабора", OleDb.OleDbType.VarChar, 9, "Набор")
        DAt.UpdateCommand.Parameters.Add("@КатНомерМонеты", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        Dim insCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand("INSERT INTO [Состав наборов] (КатНомерНабора, КатНомерМонеты) VALUES (@КатНомерНабора, @КатНомерМонеты)", Con)
        DAt.InsertCommand = insCommand
        insCommand.Parameters.Add("@КатНомерНабора", OleDb.OleDbType.VarChar, 9, "Набор")
        insCommand.Parameters.Add("@КатНомерМонеты", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
    End Sub

    Private Function GetCol2_DataGridViewColumn() As DataGridViewColumn
        Dim comboColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
        'this.GetCol2Table() - возвращает второй источник данных для ComboBox (см. ниже)
        comboColumn.DataSource = GetCol2Table()
        comboColumn.DataPropertyName = "КатНомерМонеты"
        comboColumn.DisplayMember = "КатНомерМонеты"
        comboColumn.ValueMember = "КатНомерМонеты"
        comboColumn.Name = "КатНомерМонеты"
        Return comboColumn
    End Function

    Private Function GetCol2Table() As DataTable
        Dim DA As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
        Dim table As DataTable = New DataTable()
        SqlCom = New OleDb.OleDbCommand("SELECT [Каталожный номер] as КатНомерМонеты FROM Монеты ORDER BY [Каталожный номер]", Con) ' Указываем строку запроса и привязываем к соединению
        DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DA.Fill(table) ' Заполняем таблицу результатми
        Return table
    End Function

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'заполнение полей параметров набора
        Try
            Label7.Text = table.Rows(ComboBox1.SelectedIndex)(1)
            Label9.Text = table.Rows(ComboBox1.SelectedIndex)(2)
            Label10.Text = table.Rows(ComboBox1.SelectedIndex)(7)
            Label11.Text = table.Rows(ComboBox1.SelectedIndex)(6)
            Label12.Text = table.Rows(ComboBox1.SelectedIndex)(4)
            Label13.Text = table.Rows(ComboBox1.SelectedIndex)(3)
            Label14.Text = table.Rows(ComboBox1.SelectedIndex)(8)
        Catch ex As System.InvalidCastException
            Label14.Text = ""
        End Try

        Try
            rus_forein = IIf(table.Rows(ComboBox1.SelectedIndex)(4) = "российск. рубль", "Российские\", "Иностранные\")

            'проверка на существование значений в ячейках картинок и обновление картинок
            If table.Rows(ComboBox1.SelectedIndex)(10).ToString = "" Then
                PictureBox2.Visible = False
                If table.Rows(ComboBox1.SelectedIndex)(9).ToString = "" Then
                    PictureBox1.Visible = False
                Else
                    PictureBox1.Visible = True
                    PictureBox1.Load(Pic_location + rus_forein + "Аверсы\" + table.Rows(ComboBox1.SelectedIndex)(9))
                End If
            Else
                PictureBox1.Visible = True
                PictureBox1.Load(Pic_location + rus_forein + "Реверсы\" + table.Rows(ComboBox1.SelectedIndex)(10))
                If table.Rows(ComboBox1.SelectedIndex)(9).ToString = "" Then
                    PictureBox2.Visible = False
                Else
                    PictureBox2.Visible = True
                    PictureBox2.Load(Pic_location + rus_forein + "Аверсы\" + table.Rows(ComboBox1.SelectedIndex)(9))
                End If
            End If
        Catch ex As Exception
        End Try

        If FirstOpen Then
            FirstOpen = False
        Else
            Update_table()
        End If

        'заполнение таблицы
        SqlCom = New OleDb.OleDbCommand("SELECT [Состав наборов].КатНомерНабора as Набор, [Состав наборов].КатНомерМонеты as `Каталожный номер`, Монеты.Краткое as `Краткое наименование`, Монеты.Год, Монеты.Качество, Монеты.Металл, Монеты.Проба, Монеты.Номинал, Монеты.Валюта, Монеты.Масса, Монеты.Диаметр as d, Монеты.Тираж, Монеты.Статус, Монеты.Серия, Монеты.ФотоАверс, Монеты.ФотоРеверс 
        FROM [Состав наборов] LEFT JOIN Монеты On [Состав наборов].КатНомерМонеты = Монеты.[Каталожный номер] 
        WHERE((([Состав наборов].КатНомерНабора) = '" + table.Rows(ComboBox1.SelectedIndex)(0) + "'))
        ORDER BY Монеты.Металл, Монеты.Масса, Монеты.Валюта, Монеты.Номинал, Монеты.Год, Монеты.Краткое;", Con)
        DAt.SelectCommand = SqlCom
        tbt.Clear()
        DAt.Fill(tbt)
        bs1.DataSource = tbt
        DataGridView1.DataSource = bs1
        DataGridView1.Columns(14).Visible = False
        DataGridView1.Columns(15).Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        tbt.Rows.Add(table.Rows(ComboBox1.SelectedIndex)(0),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(1),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(2),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(8),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(6),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(3),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(9),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(4),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(5),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(7),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(10),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(11),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(12),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(13),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(14),
                     tableForCmb2.Rows(ComboBox2.SelectedIndex)(15))
        DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        Catch c As System.ArgumentNullException
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.Columns.Count < 14 Then
            PictureBox3.Load(Pic_location + "АверсФон.jpg")
            PictureBox4.Load(Pic_location + "АверсФон.jpg")
        Else
            Try
                rus_forein = IIf(DataGridView1.CurrentRow.Cells(8).Value = "российск. рубль", "Российские\", "Иностранные\")
                'проверка на существование значений в ячейках картинок и обновление картинок
                If DataGridView1.CurrentRow.Cells(15).Value.ToString = "" Then
                    PictureBox3.Load(Pic_location + "АверсФон.jpg")
                Else
                    PictureBox3.Load(Pic_location + rus_forein + "Реверсы\" + DataGridView1.CurrentRow.Cells(15).Value)
                End If
                If DataGridView1.CurrentRow.Cells(14).Value.ToString = "" Then
                    PictureBox4.Load(Pic_location + "АверсФон.jpg")
                Else
                    PictureBox4.Load(Pic_location + rus_forein + "Аверсы\" + DataGridView1.CurrentRow.Cells(14).Value)
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SqlCom = New OleDb.OleDbCommand("SELECT КатНомерНабора as `Каталожный номер`, [Наименование набора], Номинал as `Номинал набора` , РазницаН as `Разница с включенными монетами_1`, [Масса набора], РазницаМ as `Разница с включенными монетами_2` 
FROM [Проверка наполнения наборов]", Con)
        tableForCheck.Clear()
        DApr = New OleDb.OleDbDataAdapter(SqlCom)
        DApr.Fill(tableForCheck)

        If (tableForCheck.Rows.Count = 0) Then
            MsgBox("Расхождений не найдено",, "Проверка расхождения наборов")
        Else
            If MsgBox("Расхождения найдены", MsgBoxStyle.OkCancel, "Проверка расхождения наборов") = MsgBoxResult.Ok Then
                Me.Text = "Проверка наполнения наборов"
                Button2.Enabled = False
                FlowLayoutPanel1.Enabled = False
                PictureBox1.Visible = False
                PictureBox2.Visible = False
                Panel1.Enabled = False
                Button1.Text = "Выход"
                state = True
                DataGridView1.DataSource = tableForCheck
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If state Then
            ' случай проверки расхождений
            Button1.Text = "Собрать новый набор"
            Me.Text = "Связи между наборами и входящими в их состав монетами"
            Button2.Enabled = True
            FlowLayoutPanel1.Enabled = True
            PictureBox1.Visible = True
            PictureBox2.Visible = True
            Panel1.Enabled = True
            state = False
            ComboBox1_SelectionChangeCommitted(sender, e)
        Else
            ' создание нового набора
            Dialog1.Show()
        End If
    End Sub
End Class