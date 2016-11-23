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

    Function GetTable2(i As Integer, Store As String, OpType As String) As DataTable
        'Получение таблицы из справочника монет
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        If Store = Nothing Then
            Store = ""
        End If

        If OpType = Nothing Then
            OpType = ""
        End If

        Select Case i
            Case 1
                SqlCom = New OleDb.OleDbCommand("SELECT *
FROM     [Монеты (для ввода)]", Con) ' Указываем строку запроса и привязываем к соединению
            Case 2
                SqlCom = New OleDb.OleDbCommand("SELECT *
FROM     [Монеты (для ввода из заявок)]", Con) ' Указываем строку запроса и привязываем к соединению
        End Select

        DAh = New OleDb.OleDbDataAdapter() 'Через адаптер получаем результаты запроса
        DAh.SelectCommand = SqlCom
        DAh.SelectCommand.Parameters.AddWithValue("1", Store)
        DAh.SelectCommand.Parameters.AddWithValue("2", OpType)

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

    Function GetTablePricesForCond(CatNum As String, DateM As Date, Storage As String) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT [Остатки по ценам (для изменения состояния)].Цена, Sum([Остатки по ценам (для изменения состояния)].[Кол-во]) AS Кол, [Остатки по ценам (для изменения состояния)].ВхНДС AS НДС, [Остатки по ценам (для изменения состояния)].Состояние AS Сост, [Остатки по ценам (для изменения состояния)].Дефекты 
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
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое) AS [Остатки по ценам (для изменения состояния)] 

WHERE ((([Остатки по ценам (для изменения состояния)].МестоХранения)=""" + Storage + """)) 
GROUP BY [Остатки по ценам (для изменения состояния)].Цена, [Остатки по ценам (для изменения состояния)].ВхНДС, [Остатки по ценам (для изменения состояния)].Состояние, [Остатки по ценам (для изменения состояния)].Дефекты 
HAVING (((Sum([Остатки по ценам (для изменения состояния)].[Кол-во]))<>0)) 
ORDER BY [Остатки по ценам (для изменения состояния)].Цена DESC;", Con) ' Указываем строку запроса и привязываем к соединению

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

    Function GetTableRegional() As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT Подразделения.Наименование 
FROM Подразделения 
WHERE (((Подразделения.ВидУчастника)=""Терр. банк""));", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function

    Function GetTableRegional(type As String) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT Подразделения.Наименование 
FROM Подразделения 
WHERE (((Подразделения.ВидУчастника)=""" + type + """) AND ((Подразделения.Активный)=Yes)) 
ORDER BY Подразделения.Наименование;", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function

    Function GetTableContr() As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT *
FROM [Юридические лица]", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function

    Function GetTableExplan(str As String) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT [Разбивка филиалов].[Выделенная группа] as Расшифровка 
FROM [Разбивка филиалов] 
WHERE ((([Разбивка филиалов].Подразделение)=""" + str + """));", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function

    Function GetTablePricesForOperations(CatNum As String, DateM As Date, Storage As String, FullPrice As Boolean, State As String) As DataTable
        Dim SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
        Dim DAh As New OleDb.OleDbDataAdapter
        Dim table As New DataTable() ' таблица с монетами
        Dim Con As New OleDb.OleDbConnection(MainSettings.AppS.ConnStr) ' Переменная для подключения базы

        SqlCom = New OleDb.OleDbCommand("SELECT [Остатки по ценам (для ввода)].Цена, Sum([Остатки по ценам (для ввода)].[Кол-во]) AS Кол, [Остатки по ценам (для ввода)].ВхНДС AS НДС, [Остатки по ценам (для ввода)].Состояние AS Сост, [Остатки по ценам (для ввода)].Дефекты, [Остатки по ценам (для ввода)].МестоХранения AS Хран 
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
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое) AS [Остатки по ценам (для ввода)] 

WHERE (((IIF(ISNULL([Состояние]), """", [Состояние]))=IIf(" + CStr(FullPrice) + "=True Or IIF(ISNULL(""" + State + """), """", """ + State + """)="""",IIF(ISNULL([Состояние]), """", [Состояние]),""" + State + """)) AND (([Остатки по ценам (для ввода)].МестоХранения)=IIf(" + CStr(FullPrice) + ",[МестоХранения],""" + Storage + """))) 

GROUP BY [Остатки по ценам (для ввода)].Цена, [Остатки по ценам (для ввода)].ВхНДС, [Остатки по ценам (для ввода)].Состояние, [Остатки по ценам (для ввода)].Дефекты, [Остатки по ценам (для ввода)].МестоХранения 
HAVING (((Sum([Остатки по ценам (для ввода)].[Кол-во]))<>0)) 
ORDER BY [Остатки по ценам (для ввода)].Цена DESC;", Con) ' Указываем строку запроса и привязываем к соединению

        DAh.SelectCommand = SqlCom
        DAh.Fill(table) ' Заполняем таблицу результатми

        Return table
    End Function
End Module
