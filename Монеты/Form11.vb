Imports Access = Microsoft.Office.Interop.Access
Public Class Form11
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private DAh As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private table As DataTable = New DataTable()
    Private NullMonet As Boolean = True 'Проверка на пустоту поля выбора монеты

    Private SelectItem As Int16 = 1
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            SelectItem = 1
            Update_Table()
        Else
            SelectItem = 2
            Update_Table()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Показать все монеты
        Dim oAccess As Access.Application
        Dim strdate As String = ""

        oAccess = New Access.Application()
        Try
            oAccess.OpenCurrentDatabase(filepath:="D:\Монеты-Access\Монеты.mdb", Exclusive:=True)
            oAccess.Visible = False
        Catch ex As Exception
            MsgBox("Can't open database file!", MsgBoxStyle.Critical, "Error")
            Return
        End Try

        Try
            oAccess.DoCmd.OpenForm("Цены", Microsoft.Office.Interop.Access.AcFormView.acNormal)
            oAccess.DoCmd.SetProperty("Дата", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
            If SelectItem = 1 Then
                oAccess.DoCmd.SetProperty("ТипОтчёта", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, 1)
            Else
                oAccess.DoCmd.SetProperty("ТипОтчёта", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, 2)
            End If
            oAccess.DoCmd.OpenReport("Остатки", Access.AcView.acViewPreview)
            strdate = Date.Now.ToString("ddMMyyHHmmss")
            oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Остатки", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
            oAccess.DoCmd.Quit()
            Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
            Form8.Show()
        Catch ex As Exception
            MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
            oAccess.DoCmd.Quit()
        End Try
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet._Монеты__для_ввода_". При необходимости она может быть перемещена или удалена.
        Me.Монеты__для_ввода_TableAdapter.Fill(Me.МонетыDataSet._Монеты__для_ввода_)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Просмотр годовых остатков
        Form12.Show()
    End Sub

    Private Sub Update_Table()
        If NullMonet Then Return
        Select Case SelectItem
            Case 1
                SqlCom = New OleDb.OleDbCommand("SELECT Sum([Кол-во]) AS Кол, Цена, ВхНДС AS НДС, Состояние AS Сост, Дефекты, МестоХранения AS Хран 
FROM (
SELECT [Каталожный номер], Цена, Количество as [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Цены
WHERE [Год]=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ")

UNION ALL SELECT [Каталожный номер], Цена, Sum(IIf([Вид операции] In (""выдача"",""продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Операции
WHERE (((Year([ДатаМонет]))>=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") Or (Year([ДатаМонет])) Is Null) AND ((IIf([Вид операции] In (""покупка"",""приём""),[ДатаМонет],[ДатаДенег]))<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + "))
GROUP BY [Каталожный номер], Цена, ВхНДС, Состояние, Дефекты, МестоХранения

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " 
GROUP BY [Каталожный номер], Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " 
GROUP BY [Каталожный номер], Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " 
GROUP BY [Каталожный номер],Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " 
GROUP BY [Каталожный номер], Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое)
WHERE ((([Каталожный номер])=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """)) 
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХранения HAVING (((Sum([Кол-во]))<>0)) 
ORDER BY Цена DESC;", Con)
            Case 2
                SqlCom = New OleDb.OleDbCommand("SELECT Sum([Кол-во]) AS Кол, Цена, ВхНДС AS НДС, Состояние AS Сост, Дефекты, МестоХранения AS Хран 
FROM (
SELECT [Каталожный номер], Цена, Количество AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Цены
WHERE [Год]=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ")

UNION ALL SELECT [Каталожный номер], Цена, Sum(IIf([Вид операции] In (""Выдача"",""Продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Операции
WHERE Year([ДатаМонет])=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND [ДатаМонет]<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, Состояние, Дефекты, МестоХранения 

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения 
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения 

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения 
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Завершено]
GROUP BY [Каталожный номер],Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Завершено]
GROUP BY [Каталожный номер],Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое
)
WHERE ((([Каталожный номер])=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """)) 
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХранения HAVING (((Sum([Кол-во]))<>0)) 
ORDER BY Цена DESC;", Con)
        End Select
        table.Clear()

        DAh = New OleDb.OleDbDataAdapter(SqlCom) 'Через адаптер получаем результаты запроса
        DAh.Fill(table) ' Заполняем таблицу результатми
        DataGridView1.DataSource = table
    End Sub

    Private Sub ComboBoxEdit1_EditValueChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.EditValueChanged
        'Действие на изменение монеты
        NullMonet = False
        Update_Table()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'Действие на изменение даты
        Update_Table()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Запомнить 
        Dim dateNew As Date
        Try
            Dim str As String = CStr(InputBox("Ввведите год, на начало которого нужно сохранить остатки") - 1)
            Dim a As Int32 = CInt(str)
            If ((a > 1500) And (a < 3000)) Then
                dateNew = CDate("31.12." + str)
            Else
                'Throw Exception
                Throw New Exception
            End If
        Catch ex As Exception
            MsgBox("Год введен неверно", MsgBoxStyle.Critical, "Ошибка")
            Return
        End Try
        Dim selCom As New OleDb.OleDbCommand("SELECT * FROM [Цены] WHERE [Год]=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ")+1", Con)
        Dim insCommand As New OleDb.OleDbCommand("INSERT INTO Цены ( [Каталожный номер], Цена, Количество, ВхНДС, Монета, Год, Состояние, Дефекты, МестоХранения )
SELECT [Остатки по ценам (фактические)].[Каталожный номер], [Остатки по ценам (фактические)].Цена, Sum([Остатки по ценам (фактические)].[Кол-во]) AS [Sum_Кол-во], [Остатки по ценам (фактические)].ВхНДС, [Краткое]+' - '+CStr(IIf([Год]=2000,[Год],Right([Год],2)))+', '+Left([Качество],2)+', '+[Металл]+', '+CStr([Номинал]) AS Выражение1, Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ")+1 AS Выражение2, [Остатки по ценам (фактические)].Состояние, [Остатки по ценам (фактические)].Дефекты, [Остатки по ценам (фактические)].МестоХранения
FROM (
SELECT [Каталожный номер], Цена, Количество AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Цены
WHERE [Год]=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ")

UNION ALL SELECT [Каталожный номер], Цена, Sum(IIf([Вид операции] In (""Выдача"",""Продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Операции
WHERE Year([ДатаМонет])=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ") AND [ДатаМонет]<=" + dateNew.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, Состояние, Дефекты, МестоХранения 

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения 
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + dateNew.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения 

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения 
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + dateNew.ToString("#MM\/dd\/yyyy#") + "
GROUP BY [Каталожный номер], Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT [Каталожный номер], Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + dateNew.ToString("#MM\/dd\/yyyy#") + " AND [Завершено]
GROUP BY [Каталожный номер],Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT [Каталожный номер], Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + dateNew.ToString("#MM\/dd\/yyyy#") + " AND [Завершено]
GROUP BY [Каталожный номер],Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое
) as [Остатки по ценам (фактические)] LEFT JOIN Монеты ON [Остатки по ценам (фактические)].[Каталожный номер] = Монеты.[Каталожный номер]
GROUP BY [Остатки по ценам (фактические)].[Каталожный номер], [Остатки по ценам (фактические)].Цена, [Остатки по ценам (фактические)].ВхНДС, [Краткое]+' - '+CStr(IIf([Год]=2000,[Год],Right([Год],2)))+', '+Left([Качество],2)+', '+[Металл]+', '+CStr([Номинал]), [Остатки по ценам (фактические)].Состояние, [Остатки по ценам (фактические)].Дефекты, [Остатки по ценам (фактические)].МестоХранения
HAVING (((Sum([Остатки по ценам (фактические)].[Кол-во]))<>0));", Con)
        Dim delCommand As New OleDb.OleDbCommand("DELETE * FROM [Цены] WHERE [Год] = Year(" + dateNew.ToString("#MM\/dd\/yyyy#") + ")+1", Con)
        Dim da As New OleDb.OleDbDataAdapter(SqlCom)
        Dim tbt As New DataTable

        da.SelectCommand = selCom
        da.InsertCommand = insCommand
        da.DeleteCommand = delCommand

        da.Fill(tbt)
        If tbt.Rows.Count = 0 Then
            'Если нет строк в таблице
            Con.Close()
            da.InsertCommand.Connection = Con
            da.InsertCommand.Connection.Open()
            da.InsertCommand.ExecuteNonQuery()
        Else
            If MsgBox("Остатки на начало указанного года ужесуществуют. Заменить их?", vbOKCancel + vbQuestion) = vbOK Then
                Con.Close()
                da.DeleteCommand.Connection = Con
                da.DeleteCommand.Connection.Open()
                da.DeleteCommand.ExecuteNonQuery()
                Con.Close()
                da.InsertCommand.Connection = Con
                da.InsertCommand.Connection.Open()
                da.InsertCommand.ExecuteNonQuery()
            End If
        End If
    End Sub
End Class