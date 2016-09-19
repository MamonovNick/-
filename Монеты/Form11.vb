Imports Access = Microsoft.Office.Interop.Access
Public Class Form11
    Private ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Монеты-Access\\Монеты.mdb"
    Private Con As New OleDb.OleDbConnection(ConnString) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private SqlComForCmb As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private DAh As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private DAforCmb As OleDb.OleDbDataAdapter ' Адаптер для заполнения таблицы после запроса
    Private table As DataTable = New DataTable()
    Private tableForCmb As DataTable = New DataTable()

    Private SelectItem As Int16 = 1
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            SelectItem = 1
        Else
            SelectItem = 2
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Показать все монеты
        Dim oAccess As Access.Application
        Dim strdate As String = ""

        oAccess = New Access.Application()
        Try
            oAccess.OpenCurrentDatabase(filepath:="D:\Монеты-Access\Монеты.mdb", Exclusive:=True)
            oAccess.Visible = True
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
        Form12.Show()
    End Sub
End Class