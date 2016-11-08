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
        'Получение таблицы из справочника монет
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы
        If str = "" Then
            SqlCom = New OleDb.OleDbCommand("SELECT [Каталожный номер], Краткое, Год, Металл, Монеты.Номинал, Монеты.Качество
FROM     Монеты", Con) ' Указываем строку запроса и привязываем к соединению
        Else
            SqlCom = New OleDb.OleDbCommand("SELECT [Каталожный номер], Краткое, Год, Металл, Монеты.Номинал, LEFT(Монеты.Качество, 2) AS [Кач-во]
FROM     Монеты
where [Каталожный номер] = """ + str + """", Con) ' Указываем строку запроса и привязываем к соединению
        End If

        DAh = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh.Fill(table) ' Заполняем таблицу результатми
        Return table
    End Function

    Function GetTablePrices(CatNum As String, DateM As Date) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT [Остатки по ценам (для перемещения)].Цена, Sum([Остатки по ценам (для перемещения)].[Кол-во]) AS Кол, [Остатки по ценам (для перемещения)].ВхНДС AS НДС, [Остатки по ценам (для перемещения)].Состояние AS Сост, [Остатки по ценам (для перемещения)].Дефекты, [Остатки по ценам (для перемещения)].МестоХранения AS Хран 
FROM (SELECT Цена, Количество as [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Цены
WHERE [Год]=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") AND [Каталожный номер]=""" + CatNum + """

UNION ALL SELECT Цена, Sum(IIf([Вид операции] In (""выдача"",""продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Операции
WHERE (((Year([ДатаМонет]))=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") Or (Year([ДатаМонет])) Is Null) AND ((IIf([Вид операции] In (""покупка"",""приём""),[ДатаМонет],[ДатаДенег]))<=" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ")) AND [Каталожный номер]=""" + CatNum + """
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateM.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CatNum + """
GROUP BY Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateM.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CatNum + """
GROUP BY Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateM.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CatNum + """
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateM.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateM.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CatNum + """
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое) AS [Остатки по ценам (для перемещения)] 

GROUP BY [Остатки по ценам (для перемещения)].Цена, [Остатки по ценам (для перемещения)].ВхНДС, [Остатки по ценам (для перемещения)].Состояние, [Остатки по ценам (для перемещения)].Дефекты, [Остатки по ценам (для перемещения)].МестоХранения 
HAVING (((Sum([Остатки по ценам (для перемещения)].[Кол-во]))<>0)) 
ORDER BY [Остатки по ценам (для перемещения)].Цена DESC;", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function

    Function GetTableStores() As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT Обозначение FROM [Места хранения монет]", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function
End Module
