Imports Access = Microsoft.Office.Interop.Access

Public Class Form7
    Private SelectItem As Int16 = 1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oAccess As Access.Application
        Dim strdate As String = ""

        Select Case SelectItem
            Case 1
                'Обороты по дням
                oAccess = New Access.Application()
                Try
                    oAccess.OpenCurrentDatabase(filepath:="D:\Монеты-Access\Монеты.mdb", Exclusive:=True)
                Catch ex As Exception
                    MsgBox("Can't open database file!", MsgBoxStyle.Critical, "Error")
                    Return
                End Try

                Try
                    oAccess.Visible = False
                    oAccess.DoCmd.SetParameter("[forms]![Отчёты]![ДатаС]", DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("[forms]![Отчёты]![ДатаПо]", DateTimePicker2.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("datesin", DateTimePicker1.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.SetParameter("dateto", DateTimePicker2.Value.Date.ToString("#MM\/dd\/yyyy#"))
                    oAccess.DoCmd.OpenReport("Обороты по дням", Access.AcView.acViewPreview)
                    strdate = Date.Now.ToString("ddMMyyHHmmss")
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Обороты по дням", "PDF", Application.StartupPath + "\TempReports\Report" + strdate + ".pdf", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
                    oAccess.CloseCurrentDatabase()
                    Class1.setFileWeb(Application.StartupPath + "\TempReports\Report" + strdate + ".pdf")
                    Form8.Show()
                Catch ex As Exception
                    MsgBox("File was not created!", MsgBoxStyle.Critical, "Error")
                    oAccess.CloseCurrentDatabase()
                End Try
            Case 2

            Case 3

        End Select
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            SelectItem = 1
        Else
            If RadioButton2.Checked Then
                SelectItem = 2
            Else
                SelectItem = 3
            End If
        End If
    End Sub
End Class