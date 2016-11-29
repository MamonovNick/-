Imports Access = Microsoft.Office.Interop.Access
Imports Монеты.MainSettings
Public Class Form9
    Private SelectItem As Int16 = 1
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            SelectItem = 1
        Else
            SelectItem = 2
        End If
    End Sub

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
                'Условные ставки
                Try
                    oAccess.DoCmd.OpenForm("Условные", Microsoft.Office.Interop.Access.AcFormView.acNormal)
                    oAccess.DoCmd.SetProperty("ДатаС", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
                    oAccess.DoCmd.SetProperty("ДатаПо", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker2.Value.Date))
                    oAccess.DoCmd.OpenReport("Условные", Access.AcView.acViewPreview)
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Условные", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                    oAccess.DoCmd.Quit()
                    Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
                    Form8.Show()
                Catch ex As Exception
                    MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
                    oAccess.DoCmd.Quit()
                End Try
            Case 2
                'Условные распоряжения
                Try
                    oAccess.DoCmd.OpenForm("Условные", Microsoft.Office.Interop.Access.AcFormView.acNormal)
                    oAccess.DoCmd.SetProperty("ДатаС", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker1.Value.Date))
                    oAccess.DoCmd.SetProperty("ДатаПо", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(DateTimePicker2.Value.Date))
                    oAccess.DoCmd.OpenReport("Условные по распоряжениям", Access.AcView.acViewPreview)
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Условные по распоряжениям", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                    oAccess.DoCmd.Quit()
                    Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
                    Form8.Show()
                Catch ex As Exception
                    MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
                    oAccess.DoCmd.Quit()
                End Try
        End Select
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.Show()
    End Sub
End Class