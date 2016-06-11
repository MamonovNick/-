Public Class Form3
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private SqlComForCmb As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private DAh As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private DAt As OleDb.OleDbDataAdapter
    Private DAforCmb As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private table As DataTable = New DataTable()
    Private tableForCmb As DataTable = New DataTable()
    Private tablet As DataTable = New DataTable()

    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        G_coins.Show()
    End Sub

    Private Function GetTable() As DataTable
        SqlCom = New OleDb.OleDbCommand("SELECT Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2) AS [Кач-во], Монеты.Масса, Монеты.Качество, Монеты.Серия
FROM     (Монеты INNER JOIN
                  [Состав наборов] ON Монеты.[Каталожный номер] = [Состав наборов].КатНомерНабора)
GROUP BY Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2), Монеты.Масса, Монеты.Качество, Монеты.Серия, Монеты.Год, 
                  Монеты.Счётчик
ORDER BY Монеты.Год, Монеты.Счётчик", Con) ' Указываем строку запроса и привязываем к соединению

        DAh = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh.Fill(table) ' Заполняем таблицу результатми
    End Function

    Private Function GetTableForCmb() As DataTable
        SqlComForCmb = New OleDb.OleDbCommand("SELECT Монеты.[Каталожный номер] ++ ' ' ++ Монеты.Краткое ++ ' ' ++ Монеты.Металл as txt FROM (SELECT Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2) AS [Кач-во], Монеты.Масса, Монеты.Качество, Монеты.Серия
FROM     (Монеты INNER JOIN
                  [Состав наборов] ON Монеты.[Каталожный номер] = [Состав наборов].КатНомерНабора)
GROUP BY Монеты.[Каталожный номер], Монеты.Краткое, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, LEFT(Монеты.Качество, 2), Монеты.Масса, Монеты.Качество, Монеты.Серия, Монеты.Год, 
                  Монеты.Счётчик
ORDER BY Монеты.Год, Монеты.Счётчик)", Con) ' Указываем строку запроса и привязываем к соединению
        DAforCmb = New OleDb.OleDbDataAdapter(SqlComForCmb) 'Через адаптер получаем результаты запроса
        DAforCmb.Fill(tableForCmb) ' Заполняем таблицу результатми
    End Function

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetTable()
        GetTableForCmb()
        'SELECT Монеты.[Каталожный номер] ++ ' ' ++ Монеты.Краткое ++ ' ' ++ Монеты.Металл as txt
        ComboBox1.DataSource = tableForCmb
        ComboBox1.DisplayMember = "txt"
        ComboBox1_SelectionChangeCommitted(sender, e)
        'Label7.Text = table.Rows(ComboBox1.SelectedIndex)(1)
    End Sub

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
        'заполнение таблицы
        SqlCom = New OleDb.OleDbCommand("SELECT [Состав наборов].КатНомерНабора, [Состав наборов].КатНомерМонеты, Монеты.Краткое, Монеты.Год, Монеты.Металл, Монеты.Номинал, Монеты.Серия, Монеты.Валюта, Монеты.Качество, Монеты.Проба, Монеты.Масса, Монеты.Диаметр, Монеты.Тираж, Монеты.Статус, Монеты.ФотоАверс, Монеты.ФотоРеверс FROM [Состав наборов] LEFT JOIN Монеты ON [Состав наборов].КатНомерМонеты = Монеты.[Каталожный номер] WHERE ((([Состав наборов].КатНомерНабора)=[ComboBox1.SelectedIndex]![Состав наборов]![Каталожный номер])) ORDER BY Монеты.Металл, Монеты.Масса, Монеты.Валюта, Монеты.Номинал, Монеты.Год, Монеты.Краткое;", Con)
        DAt = New OleDb.OleDbDataAdapter(SqlCom)
        DAt.Fill(tablet)
        DataGridView1.DataSource = tablet
    End Sub
End Class