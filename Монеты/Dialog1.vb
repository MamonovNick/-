Imports System.Windows.Forms

Public Class Dialog1
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlComForCmb As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private DAforCmb As New OleDb.OleDbDataAdapter() ' Адаптер для заполнения таблицы после запроса
    Private tableForCmb As DataTable = New DataTable()
    Private FirstOpen_1 As Boolean = True
    Private FirstOpen_2 As Boolean = True

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Class1.PutCat(tableForCmb.Rows(ComboBox1.SelectedIndex)(0))
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub GetTableForCmb()
        SqlComForCmb = New OleDb.OleDbCommand("SELECT  Монеты.[Каталожный номер] as txt, [Монеты].[Краткое]+IIf([Монеты].[Каталожный номер]<>'5216-0060' And [Монеты].[Каталожный номер]<>'5111-0178',' - '+CStr(IIf([Монеты].[Год]=2000,[Монеты].[Год],Right([Монеты].[Год],2))),'') AS Наименование, Монеты.Металл, Монеты.Номинал, Монеты.Валюта, Left([Монеты].[Качество],2) AS [Кач-во], Монеты.Масса, Монеты.Качество, Монеты.Серия, [Состав наборов].КатНомерНабора, [Состав наборов_1].КатНомерМонеты
FROM (Монеты LEFT JOIN [Состав наборов] ON Монеты.[Каталожный номер] = [Состав наборов].КатНомерНабора) LEFT JOIN [Состав наборов] AS [Состав наборов_1] ON Монеты.[Каталожный номер] = [Состав наборов_1].КатНомерМонеты 
WHERE ((" + IIf(RadioButton1.Checked, "Mid([Монеты].[Каталожный номер],2,3)", "'099'") + "='099')  AND ((Монеты.ВыводВСписок)=" + IIf(CheckBox1.Checked, "true", "[ВыводВСписок]") + "))
GROUP BY Монеты.[Каталожный номер], [Монеты].[Краткое]+IIf([Монеты].[Каталожный номер]<>'5216-0060' And [Монеты].[Каталожный номер]<>'5111-0178',' - '+CStr(IIf([Монеты].[Год]=2000,[Монеты].[Год],Right([Монеты].[Год],2))),''), Монеты.Металл, Монеты.Номинал, Монеты.Валюта, Left([Монеты].[Качество],2), Монеты.Масса, Монеты.Качество, Монеты.Серия, [Состав наборов].КатНомерНабора, [Состав наборов_1].КатНомерМонеты, IIf([Монеты].[Валюта]='российск. рубль',2,1), Монеты.Металл, Монеты.Масса, Монеты.Валюта, Монеты.Номинал, Монеты.Год, Монеты.Краткое 
HAVING ((([Состав наборов].КатНомерНабора) Is Null) AND (([Состав наборов_1].КатНомерМонеты) Is Null)) 
ORDER BY IIf([Монеты].[Валюта]='российск. рубль',2,1), Монеты.Металл DESC , Монеты.Масса, Монеты.Валюта, Монеты.Номинал, Монеты.Год, Монеты.Краткое", Con) ' Указываем строку запроса и привязываем к соединению
        DAforCmb.SelectCommand = SqlComForCmb 'Через адаптер получаем результаты запроса
        tableForCmb.Clear() 'Очищаем таблицу
        DAforCmb.Fill(tableForCmb) ' Заполняем таблицу результатми
    End Sub

    Private Sub LoadToCmb(sender As Object, e As EventArgs)
        GetTableForCmb()
        ComboBox1.DisplayMember = "txt" 'Столбец для отображения в combobox
        ComboBox1.DataSource = tableForCmb ' Загрузка таблицы для отображения в combobox
        ComboBox1_SelectionChangeCommitted(sender, e) ' Фиксация изменения строки
    End Sub

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadToCmb(sender, e)
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Label2.Text = tableForCmb.Rows(ComboBox1.SelectedIndex)(1)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If (FirstOpen_1) Then
            FirstOpen_1 = False
        Else
            LoadToCmb(sender, e)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If FirstOpen_2 Then
            FirstOpen_2 = False
        Else
            LoadToCmb(sender, e)
        End If
    End Sub
End Class
