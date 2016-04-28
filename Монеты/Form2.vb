Imports System.Data.OleDb
Imports System.Data
Imports System.Xml
Imports System.Data.DataException

Public Class G_coins
    Private ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb" 'строка подключения
    Public CmdText As String = "SELECT * FROM [Монеты]"
    Private Pic_location As String = "D:\Монеты-Access\Фотографии монет\" 'расположение фотографий монет
    Dim bs1 As New BindingSource() 'Переменная bindingsourse
    Dim Con As New OleDb.OleDbConnection(ConnectionString) ' Переменная для подключения базы
    Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Dim DS As New Data.DataSet
    'Dim DT As New Data.DataTable ' Таблица для хранения результатов запроса
    'Dim DT2 As New Data.DataTable ' Таблица для хранения результатов запроса
    'Dim DT3 As New Data.DataTable ' Таблица для хранения результатов запроса
    ' Dim DA As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса

    'Private Sub Clear_tables()
    '    DT.Clear()
    '    DT2.Clear()
    '    DT3.Clear()
    'End Sub

    'Private Sub AdjustColumnOrder()
    '    With DataGridView1
    '        'настройка порядка столбцов в таблице
    '        'индексы с нулевого, имена столбцов из файла .xsd
    '        .Columns("Каталожный номер").DisplayIndex = 0
    '        .Columns("Проба").DisplayIndex = 2
    '    End With
    'End Sub

    'Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.Click
    '    Me.Close() 'закрытие текущей формы
    '    MainForm.Show() 'отображение основной формы
    'End Sub

    'Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
    '    'загрузка изображений монет в picturebox
    '    Dim rus_forein As String
    '    Try
    '        If DataGridView1.ColumnCount < 15 Or DataGridView1.CurrentRow.Cells(6).Value.ToString = "" Then
    '            PictureBox1.Load(Pic_location + "АверсФон.jpg")
    '            PictureBox2.Load(Pic_location + "АверсФон.jpg")
    '            Exit Sub
    '        End If
    '        rus_forein = IIf(DataGridView1.CurrentRow.Cells(6).Value = "российск. рубль", "Российские\", "Иностранные\")
    '        'проверка на существование значений в ячейках картинок и обновление картинок
    '        If DataGridView1.CurrentRow.Cells(14).Value.ToString = "" Then
    '            PictureBox1.Load(Pic_location + "АверсФон.jpg")
    '        Else
    '            PictureBox1.Load(Pic_location + rus_forein + "Реверсы\" + DataGridView1.CurrentRow.Cells(14).Value)
    '        End If
    '        If DataGridView1.CurrentRow.Cells(15).Value.ToString = "" Then
    '            PictureBox2.Load(Pic_location + "АверсФон.jpg")
    '        Else
    '            PictureBox2.Load(Pic_location + rus_forein + "Аверсы\" + DataGridView1.CurrentRow.Cells(15).Value)
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub

    'Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Coins_guide.Click
    '    'DA.Update(DT)
    '    'Con.Close() ' Закрываем соединение
    '    Clear_tables() 'очистка таблиц для запросов
    '    SqlCom = New OleDb.OleDbCommand("SELECT Монеты.Год, * FROM Монеты ORDER BY Монеты.Год, Монеты.Счётчик", Con) ' Указываем строку запроса и привязываем к соединению
    '    Con.Open() ' Открываем соединение
    '    SqlCom.ExecuteNonQuery() 'Выполняем запрос
    '    DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
    '    DA.Fill(DT) ' Заполняем таблицу результатми
    '    bs1.DataSource = DT
    '    BindingNavigator1.BindingSource = bs1 'привязываем навигатор к источнику данных
    '    DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
    '    Con.Close() ' Закрываем соединение
    '    AdjustColumnOrder() 'упорядочиваем столбцы
    '    DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1)
    'End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Addition.Click
    '    Clear_tables() 'очистка таблиц для запросов
    '    SqlCom = New OleDb.OleDbCommand("SELECT Наименование, Качество, [Облаг НДС], [Металл, проба], Масса, [Ном-л], [Каталожный номер] FROM [Последние введённые монеты]", Con) ' Указываем строку запроса и привязываем к соединению
    '    Con.Open() ' Открываем соединение
    '    SqlCom.ExecuteNonQuery() 'Выполняем запрос
    '    DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
    '    DA.Fill(DT2) ' Заполняем таблицу результатми
    '    bs1.DataSource = DT2
    '    BindingNavigator1.BindingSource = bs1 'привязываем навигатор к источнику данных
    '    DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
    '    Con.Close() ' Закрываем соединение
    '    DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1)
    'End Sub

    'Private Sub Catalog_Click(sender As Object, e As EventArgs) Handles Catalog.Click
    '    Clear_tables() 'очистка таблиц для запросов
    '    SqlCom = New OleDb.OleDbCommand("SELECT [Каталожные номера иностранных].* FROM [Каталожные номера иностранных]", Con) ' Указываем строку запроса и привязываем к соединению
    '    Con.Open() ' Открываем соединение
    '    SqlCom.ExecuteNonQuery() 'Выполняем запрос
    '    DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
    '    DA.Fill(DT3) ' Заполняем таблицу результатми
    '    bs1.DataSource = DT3
    '    BindingNavigator1.BindingSource = bs1 'привязываем навигатор к источнику данных
    '    DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
    '    Con.Close() ' Закрываем соединение
    '    DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1)
    'End Sub

    Private Sub G_coins_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Clear_tables() 'очистка таблиц для запросов
        'SqlCom = New OleDb.OleDbCommand("SELECT * FROM Монеты ORDER BY Монеты.Год, Монеты.Счётчик", Con) ' Указываем строку запроса и привязываем к соединению
        'Con.Open() ' Открываем соединение
        'SqlCom.ExecuteNonQuery() 'Выполняем запрос
        'DA = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DA.Fill(МонетыDataSet1.Монеты) ' Заполняем таблицу результатми
        bs1.DataSource = МонетыDataSet1.Монеты
        BindingNavigator1.BindingSource = bs1 'привязываем навигатор к источнику данных
        DataGridView1.DataSource = bs1 ' Привязываем Грид к источнику данных
        'DA.Update(DS)
        'Con.Close() ' Закрываем соединение
        'AdjustColumnOrder()
        'DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.RowCount - 2).Cells(1)
    End Sub

    Private Sub Coins_guide_Click(sender As Object, e As EventArgs) Handles Coins_guide.Click
        'DA.UpdateCommand = OleDbCommandBuilder.
        DA.Update(МонетыDataSet1.Монеты)
    End Sub

    Private Sub Addition_Click(sender As Object, e As EventArgs) Handles Addition.Click

    End Sub
End Class

'Partial Public Class Frm_ref_staff
'    Public Sub Frm_ref_staff()
'        InitializeComponent()
'    End Sub

'    Private Sub Frm_ref_staff_Load(sender As Object, e As EventArgs)
'        'TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet_DBGIS_ASU.Tbl_officer". При необходимости она может быть перемещена или удалена.
'        this.tbl_officerTableAdapter.Fill(this.dataSet_DBGIS_ASU.Tbl_officer)
'    End Sub

'    Private Sub Frm_ref_staff_FormClosed(sender As Object, e As FormClosedEventArgs)
'        ((Frm_VodoCheb)this.MdiParent).staffToolStripMenuItem.Visible = true
'    End Sub

'    Private Sub button1_Click(sender As Object, e As EventArgs) 'кнопка "Сохранить изменения"
'        Try
'            this.tbl_officerTableAdapter.Update(this.dataSet_DBGIS_ASU.Tbl_officer)
'            MessageBox.Show("Изменения в базе данных выполнены!", "Уведомление о результатах", MessageBoxButtons.OK)
'        Catch ex As Exception
'            MessageBox.Show("Изменения в базе данных выполнить не удалось!", "Уведомление о результатах", MessageBoxButtons.OK)
'        End Try
'    End Sub
'End Class