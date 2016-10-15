Public Class Form13
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private delCommand As OleDb.OleDbCommand ' переменная для запроса удаления
    Private updCommand As OleDb.OleDbCommand ' переменная для запроса апдейта
    Private insCommand As OleDb.OleDbCommand ' переменная для запроса вставки

    Private DA As New OleDb.OleDbDataAdapter ' адаптер

    Private bs1 As New BindingSource() 'Переменная bindingsourse
    Private tbt As New DataTable() ' переменная таблица для вывода в грид


    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet._Монеты__для_ввода_". При необходимости она может быть перемещена или удалена.
        Me.Монеты__для_ввода_TableAdapter.Fill(Me.МонетыDataSet._Монеты__для_ввода_)
    End Sub

    Private Sub Update_grid2()
        SqlCom = New OleDb.OleDbCommand("SELECT Sum([Остатки по ценам (для расчёта)].[Кол-во]) AS Кол, [Остатки по ценам (для расчёта)].Цена, [Остатки по ценам (для расчёта)].ВхНДС AS НДС, [Остатки по ценам (для расчёта)].Состояние AS Сост, [Остатки по ценам (для расчёта)].МестоХранения AS Хран 
FROM (SELECT Цена, Количество as [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Цены 
WHERE [Год]=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + "

UNION ALL SELECT Цена, Sum(IIf([Вид операции] In (""Выдача"",""Продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Операции 
WHERE (((Year([ДатаМонет]))=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") Or (Year([ДатаМонет])) Is Null) AND ((IIf([Вид операции] In (""покупка"",""приём""),[ДатаМонет],[ДатаДенег]))<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ")) AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + " 
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + "
GROUP BY Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + "
GROUP BY Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + "
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + "
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое) AS [Остатки по ценам (для расчёта)] 
GROUP BY [Остатки по ценам (для расчёта)].Цена, [Остатки по ценам (для расчёта)].ВхНДС, [Остатки по ценам (для расчёта)].Состояние, [Остатки по ценам (для расчёта)].МестоХранения HAVING (((Sum([Остатки по ценам (для расчёта)].[Кол-во]))<>0));", Con)
        DA = New OleDb.OleDbDataAdapter(SqlCom)
        DA.Fill(tbt)
        GridControl2.DataSource = tbt
    End Sub

    Private Sub ComboBoxEdit1_EditValueChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.EditValueChanged
        Update_grid2()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Update_grid2()
    End Sub
End Class