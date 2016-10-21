Module Module1
    'настройка папок и других прараметров перед запуском программы
    Sub Start_Setup()
        If Not IO.Directory.Exists(Application.StartupPath + "\TempReports") Then
            IO.Directory.CreateDirectory(Application.StartupPath + "\TempReports")
        End If
    End Sub

    Sub Del_Tmp()
        If IO.Directory.Exists(Application.StartupPath + "\TempReports") Then
            IO.Directory.Delete(Application.StartupPath + "\TempReports", True)
        End If
    End Sub

    Function GetTable(str As String) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
        Dim Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT [Каталожный номер], Краткое, Год, Металл, Монеты.Номинал, LEFT(Монеты.Качество, 2) AS [Кач-во]
FROM     Монеты
where [Каталожный номер] = """ + str + """", Con) ' Указываем строку запроса и привязываем к соединению

        DAh = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh.Fill(table) ' Заполняем таблицу результатми
        Return table
    End Function
End Module
