Imports System.Data.OleDb
Imports System.Data
Imports System.Xml
Imports System.Data.DataException

Public Class G_coins
    Private Pic_location As String = "D:\Монеты-Access\Фотографии монет\" 'расположение фотографий монет
    Private ActiveTable As Int16 ' Номер активной таблицы
    Dim bs1 As New BindingSource() 'Переменная bindingsourse

    Private Sub Update_table()
        Select Case ActiveTable
            Case 0 ' Coins_guide
                DA.Update(МонетыDataSet1)
        End Select
    End Sub

    Private Sub AdjustColumnOrder()
        With DataGridView1
            'настройка порядка столбцов в таблице
            'индексы с нулевого, имена столбцов из файла .xsd
            .Columns("Каталожный номер").DisplayIndex = 0
            '.Columns("Проба").DisplayIndex = 2
        End With
    End Sub

    Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.Click
        Me.Close() 'закрытие текущей формы
        MainForm.Show() 'отображение основной формы
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        'загрузка изображений монет в picturebox
        Dim rus_forein As String
        Try
            If ActiveTable <> 0 Then
                PictureBox1.Load(Pic_location + "АверсФон.jpg")
                PictureBox2.Load(Pic_location + "АверсФон.jpg")
                Exit Sub
            End If
            rus_forein = IIf(DataGridView1.CurrentRow.Cells(6).Value = "российск. рубль", "Российские\", "Иностранные\")
            'проверка на существование значений в ячейках картинок и обновление картинок
            If DataGridView1.CurrentRow.Cells(14).Value.ToString = "Нет" Then
                PictureBox1.Load(Pic_location + "АверсФон.jpg")
            Else
                PictureBox1.Load(Pic_location + rus_forein + "Реверсы\" + DataGridView1.CurrentRow.Cells(14).Value)
            End If
            If DataGridView1.CurrentRow.Cells(15).Value.ToString = "Нет" Then
                PictureBox2.Load(Pic_location + "АверсФон.jpg")
            Else
                PictureBox2.Load(Pic_location + rus_forein + "Аверсы\" + DataGridView1.CurrentRow.Cells(15).Value)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function GetCol2_DataGridViewColumn() As DataGridViewColumn
        Dim comboColumn As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn()
        'this.GetCol2Table() - возвращает второй источник данных для ComboBox (см. ниже)
        Try
            comboColumn.DataSource = GetCol2Table()
        Catch ex As Exception

        End Try
        comboColumn.DataPropertyName = "Валюта"
        comboColumn.DisplayMember = "Валюта"
        comboColumn.ValueMember = "Валюта"
        comboColumn.Name = "Валюта"
        Return comboColumn
    End Function

    Private Function GetCol2Table() As DataTable
        Dim ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
        Dim Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DA As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
        Dim table As DataTable = New DataTable()
        SqlCom = New OleDb.OleDbCommand("SELECT [Список валют].[Валюта] FROM [Список валют] ORDER BY [Список валют].[Валюта]", Con) ' Указываем строку запроса и привязываем к соединению
        DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DA.Fill(table) ' Заполняем таблицу результатами
        'table.Rows.Add()
        Return table
    End Function

    Private Sub G_coins_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActiveTable = 0
        DA.Fill(МонетыDataSet1.Монеты) ' Заполняем таблицу результатми
        bs1.DataSource = МонетыDataSet1.Монеты 'Привязываем BindingSourse к источнику данных
        BindingNavigator1.BindingSource = bs1 'Привязываем навигатор к источнику данных
        DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
        Dim oldColIndex As Int32 = DataGridView1.Columns("Валюта").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, GetCol2_DataGridViewColumn())
        AdjustColumnOrder() 'Изменяем порядок строк
        DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1) 'Выделяем последнюю строку
        DataGridView1_SelectionChanged(sender, e)
    End Sub

    Private Sub Coins_guide_Click(sender As Object, e As EventArgs) Handles Coins_guide.Click
        Update_table()
        ActiveTable = 0
        'МонетыDataSet1.Clear()
        'DA.Fill(МонетыDataSet1.Монеты) ' Заполняем таблицу результатми
        bs1.DataSource = МонетыDataSet1.Монеты 'Привязываем BindingSourse к источнику данных
        BindingNavigator1.BindingSource = bs1 'Привязываем навигатор к источнику данных
        DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
        Dim oldColIndex As Int32 = DataGridView1.Columns("Валюта").Index
        DataGridView1.Columns.RemoveAt(oldColIndex)
        DataGridView1.Columns.Insert(oldColIndex, GetCol2_DataGridViewColumn())
        AdjustColumnOrder() 'Изменяем порядок строк
        'DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1) 'Выделяем последнюю строку
        DataGridView1_SelectionChanged(sender, e)
        DataGridView1.ReadOnly = False
        'BindingNavigator1.AddNewItem = BindingNavigatorAddNewItem
        'BindingNavigator1.DeleteItem = BindingNavigatorDeleteItem
    End Sub

    Private Sub Addition_Click(sender As Object, e As EventArgs) Handles Addition.Click
        Update_table()
        ActiveTable = 1
        DA2.Fill(МонетыDataSet.Последние_введённые_монеты)
        bs1.DataSource = МонетыDataSet.Последние_введённые_монеты 'Привязываем BindingSourse к источнику данных
        BindingNavigator1.BindingSource = bs1 'Привязываем навигатор к источнику данных
        DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
        DataGridView1.ReadOnly = True
        BindingNavigator1.AddNewItem = Nothing
        BindingNavigator1.DeleteItem = Nothing
    End Sub

    Private Sub G_coins_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Update_table()
    End Sub

    Private Sub Catalog_Click(sender As Object, e As EventArgs) Handles Catalog.Click
        Update_table()
        ActiveTable = 2
        Каталожные_номера_иностранныхTableAdapter.Fill(МонетыDataSet.Каталожные_номера_иностранных)
        bs1.DataSource = МонетыDataSet.Каталожные_номера_иностранных 'Привязываем BindingSourse к источнику данных
        BindingNavigator1.BindingSource = bs1 'Привязываем навигатор к источнику данных
        DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
        DataGridView1.ReadOnly = True
        BindingNavigator1.AddNewItem = Nothing
        BindingNavigator1.DeleteItem = Nothing
    End Sub

    Private Sub Tr_sets_Click(sender As Object, e As EventArgs) Handles Tr_sets.Click
        Me.Visible = False
        Form3.Show()
    End Sub
End Class