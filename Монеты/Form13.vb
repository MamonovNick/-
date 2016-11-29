Imports Access = Microsoft.Office.Interop.Access
Imports Монеты.MainSettings
Public Class Form13
    Private oAccess As Access.Application
    Private Con As New OleDb.OleDbConnection(AppS.ConnStr) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов

    Private DA As New OleDb.OleDbDataAdapter ' адаптер
    Private DA2 As New OleDb.OleDbDataAdapter ' адаптер

    Private tbt As New DataTable() ' переменная таблица для вывода в грид 2
    Private tbt_main As New DataTable() ' переменная таблица для вывода в грид 1

    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet._Монеты__для_ввода_". При необходимости она может быть перемещена или удалена.
        Me.Монеты__для_ввода_TableAdapter.Fill(Me.МонетыDataSet._Монеты__для_ввода_)

        'Открытие базы данных
        oAccess = New Access.Application()
        Try
            oAccess.OpenCurrentDatabase(filepath:=AppS.FileDBPath, Exclusive:=False)
            oAccess.Visible = False
        Catch ex As Exception
            MsgBox("Can't open database file!", MsgBoxStyle.Critical, "Error")
            Return
        End Try
    End Sub

    Private Sub Update_SumTxtBox()
        ' обновление суммирующих полей для основной таблицы
        Dim sum1 As Integer = 0, sum2 = 0, sum3 = 0
        If tbt_main.Rows.Count = 0 Then
            TextBox6.Text = 0
            TextBox5.Text = 0
            TextBox1.Text = 0
            Return
        End If
        For Each row In tbt_main.Rows
            sum1 += row(3)
            sum2 += row(4)
            sum3 += row(5)
        Next
        TextBox6.Text = sum1
        TextBox5.Text = sum2
        TextBox1.Text = sum3
    End Sub

    Private Sub Update_grid1()
        'обновление основной таблицы формы
        Dim delCommand As New OleDb.OleDbCommand("DELETE * FROM Распределение", Con)
        DA2.DeleteCommand = delCommand
        Con.Close()
        DA2.DeleteCommand.Connection = Con
        DA2.DeleteCommand.Connection.Open()
        DA2.DeleteCommand.ExecuteNonQuery()
        'Call Raspr form Access
        Try
            oAccess.DoCmd.OpenForm("Расчёт распределения", Microsoft.Office.Interop.Access.AcFormView.acNormal)
            oAccess.DoCmd.SetProperty("Каталожный номер", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")))
            oAccess.DoCmd.SetProperty("Дата", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
            oAccess.DoCmd.SetProperty("МаксКол", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, TextBox2.Text)
            oAccess.DoCmd.SetProperty("Минимум", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, TextBox3.Text)
            oAccess.DoCmd.SetProperty("Кратность", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, TextBox4.Text)
            oAccess.Run("Raspr")
        Catch ex As Exception
            MsgBox("Error in database!", MsgBoxStyle.Critical, "Error")
        End Try

        Con.Close()
        SqlCom = New OleDb.OleDbCommand("SELECT * FROM Распределение", Con)
        DA2 = New OleDb.OleDbDataAdapter(SqlCom)
        tbt_main.Clear()
        DA2.Fill(tbt_main)
        GridControl1.DataSource = tbt_main
        Update_SumTxtBox()
    End Sub

    Private Sub Update_grid2()
        SqlCom = New OleDb.OleDbCommand("SELECT Sum([Остатки по ценам (для расчёта)].[Кол-во]) AS Кол, [Остатки по ценам (для расчёта)].Цена, [Остатки по ценам (для расчёта)].ВхНДС AS НДС, [Остатки по ценам (для расчёта)].Состояние AS Сост, [Остатки по ценам (для расчёта)].МестоХранения AS Хран 
FROM (SELECT Цена, Количество as [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения
FROM Цены 
WHERE [Год]=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND [Каталожный номер]=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """

UNION ALL SELECT Цена, Sum(IIf([Вид операции] In (""Выдача"",""Продажа""),-[Количество],[Количество])) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХранения 
FROM Операции 
WHERE (((Year([ДатаМонет]))=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") Or (Year([ДатаМонет])) Is Null) AND ((IIf([Вид операции] In (""покупка"",""приём""),[ДатаМонет],[ДатаДенег]))<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ")) AND [Каталожный номер]=" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + " 
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, СтарСостояние AS Состояние, СтарДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """
GROUP BY Цена, ВхНДС, СтарСостояние, СтарДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, НовСостояние AS Состояние, НовДефекты AS Дефекты, МестоХранения
FROM [Изменение состояния]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """
GROUP BY Цена, ВхНДС, НовСостояние, НовДефекты, МестоХранения

UNION ALL SELECT Цена, Sum(-Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияСтарое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое

UNION ALL SELECT Цена, Sum(Количество) AS [Кол-во], ВхНДС, Состояние, Дефекты, МестоХраненияНовое AS МестоХранения
FROM [Перемещение между хранилищами]
WHERE Year(Дата)=Year(" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + ") AND Дата<=" + DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#") + " AND [Каталожный номер]=""" + CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")) + """
GROUP BY Цена, ВхНДС, Состояние, Дефекты, МестоХраненияНовое) AS [Остатки по ценам (для расчёта)] 
GROUP BY [Остатки по ценам (для расчёта)].Цена, [Остатки по ценам (для расчёта)].ВхНДС, [Остатки по ценам (для расчёта)].Состояние, [Остатки по ценам (для расчёта)].МестоХранения HAVING (((Sum([Остатки по ценам (для расчёта)].[Кол-во]))<>0));", Con)
        DA = New OleDb.OleDbDataAdapter(SqlCom)
        tbt.Clear()
        DA.Fill(tbt)
        GridControl2.DataSource = tbt
    End Sub

    Private Sub ComboBoxEdit1_EditValueChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.EditValueChanged
        'изменение монеты
        Dim table As New DataTable() ' таблица с монетами
        Cursor.Current = Cursors.WaitCursor
        table = Module1.GetTable(CStr(ComboBoxEdit1.GetColumnValue("Каталожный номер")))
        If ComboBoxEdit1.GetColumnValue("Каталожный номер") = "" Then
            Label5.Text = ""
        Else
            Label5.Text = CStr(table.Rows(0)(1)) + " - " + CStr(table.Rows(0)(2)) + ", " + CStr(table.Rows(0)(5)) + ", " + CStr(table.Rows(0)(3)) + ", " + CStr(table.Rows(0)(4))
            If CInt(TextBox2.Text) <> 0 Then
                TextBox2.Text = "0"
            End If
        End If
        Update_grid2()
        Update_grid1()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Cursor.Current = Cursors.WaitCursor
        If ComboBoxEdit1.GetColumnValue("Каталожный номер") <> "" Then
            Update_grid2()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Кнопка расчета
        If ComboBoxEdit1.GetColumnValue("Каталожный номер") = "" Then
            MsgBox("Не выбрана монета.", 48, "Недостаточно информации")
        Else
            Cursor.Current = Cursors.WaitCursor
            Update_grid2()
        End If
    End Sub

    Private Sub Form13_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Закрытие базы, всё успешно
        oAccess.DoCmd.Quit()
    End Sub
End Class