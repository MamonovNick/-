Imports Access = Microsoft.Office.Interop.Access
Imports Монеты.MainSettings

Public Class Form7
    Private SelectItem As Int16 = 1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oAccess As Access.Application
        Dim strdate As String = ""

        oAccess = New Access.Application()
        Try
            oAccess.OpenCurrentDatabase(filepath:=AppS.FileDBPath, Exclusive:=True)
            oAccess.Visible = False
        Catch ex As Exception
            MsgBox("Can't open database file!", MsgBoxStyle.Critical, "Error")
            Return
        End Try

        Select Case SelectItem
            Case 1
                'Обороты по дням
                Try
                    oAccess.DoCmd.SetParameter("[forms]![Отчёты]![ДатаС]", DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("[forms]![Отчёты]![ДатаПо]", DateTimePicker2.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("datesin", DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("dateto", DateTimePicker2.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.OpenReport("Обороты по дням", Access.AcView.acViewPreview)
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Обороты по дням", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                    oAccess.DoCmd.Quit()
                    Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
                    Form8.Show()
                Catch ex As Exception
                    MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
                    oAccess.DoCmd.Quit()
                End Try
            Case 2
                'Оперативка
                Try
                    oAccess.DoCmd.OpenForm("Отчёты", Microsoft.Office.Interop.Access.AcFormView.acNormal)
                    oAccess.DoCmd.SetProperty("ДатаС", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
                    oAccess.DoCmd.SetProperty("ДатаПо", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker2.Value.Date))
                    oAccess.DoCmd.OpenReport("Оперативка", Access.AcView.acViewPreview)
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Оперативка", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                    oAccess.DoCmd.Quit()
                    Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
                    Form8.Show()
                Catch ex As Exception
                    MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
                    oAccess.DoCmd.Quit()
                End Try
            Case 3
                'Доход
                Try
                    oAccess.DoCmd.OpenForm("Отчёты", Microsoft.Office.Interop.Access.AcFormView.acNormal)
                    oAccess.DoCmd.SetProperty("ДатаС", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
                    oAccess.DoCmd.SetProperty("ДатаПо", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker2.Value.Date))
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.Run("Profit", False, Application.StartupPath + "\TempReports\Report" + strdate + ".txt")
                    oAccess.DoCmd.Quit()
                Catch ex As Exception
                    MsgBox("Error in database!", MsgBoxStyle.Critical, "Error")
                    oAccess.DoCmd.Quit()
                End Try
                Dim f As FileIO.TextFieldParser
                Dim s As String
                Try
                    f = FileIO.FileSystem.OpenTextFieldParser(Application.StartupPath + "\TempReports\Report" + strdate + ".txt", "/n")
                    s = f.ReadLine()

                    Dim itog_ost As Long = CLng(s)
                    s = f.ReadLine()
                    Dim itog_sr_ost As Long = CLng(s)
                    s = f.ReadLine()
                    Dim itog_doh As Long = CLng(s)
                    s = f.ReadLine()
                    Dim itog_doh_dop As Long = CLng(s)
                    s = f.ReadLine()
                    Dim itog_dn As Long = CLng(s)
                    s = f.ReadLine()
                    Dim days As Long = CLng(s)
                    f.Dispose()
                    s = Nothing

                    MsgBox("Сумма вложений на конец дня " + CStr("ДатаПо") + " г.: " + CStr(itog_ost) + " руб." _
+ Chr(13) + "Средняя сума вложений с начала года: " + CStr(itog_sr_ost) + " руб." _
+ Chr(13) + "Полученный доход: " + CStr(itog_doh - itog_doh_dop) + " + " + CStr(itog_doh_dop) + " = " + CStr(itog_doh) + " руб." _
+ Chr(13) + "ИТОГО доходность: " + CStr(CLng(itog_doh / itog_sr_ost * days / itog_dn * 100 * 100) / 100) + "%")
                Catch ex As Exception
                    MsgBox("Error reading file!", MsgBoxStyle.Critical, "Error")
                End Try
        End Select
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            SelectItem = 1
            Label2.Enabled = True
            DateTimePicker1.Enabled = True
        Else
            If RadioButton2.Checked Then
                SelectItem = 2
                Label2.Enabled = True
                DateTimePicker1.Enabled = True
            Else
                SelectItem = 3
                Label2.Enabled = False
                DateTimePicker1.Enabled = False
            End If
        End If
    End Sub
End Class