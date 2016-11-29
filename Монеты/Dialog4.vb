Imports System.Windows.Forms
Imports Монеты.MainSettings
Imports Access = Microsoft.Office.Interop.Access

Public Class Dialog4

    Private Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OK_Button.Click
        DialogResult = DialogResult.OK
        ActionOrd()
        Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Cancel_Button.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

    Private Sub Dialog4_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Class1.getUpdateType() = 0 Then
            RadioButton1.Enabled = False
            RadioButton2.Checked = True
        Else
            RadioButton1.Enabled = True
            RadioButton1.Checked = True
        End If
    End Sub

    Private Sub ActionOrd()
        Cursor.Current = Cursors.WaitCursor
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
        Try
            oAccess.DoCmd.OpenForm("Заявки", Microsoft.Office.Interop.Access.AcFormView.acNormal)
            oAccess.DoCmd.SetProperty("ДатаОт", Microsoft.Office.Interop.Access.AcProperty.acPropertyValue, CDate(Class1.getDate(1)).Date)
            If Class1.getUpdateType = 1 Then
                oAccess.DoCmd.SetProperty("ПоказатьЗаявки", Access.AcProperty.acPropertyValue, 2)
                oAccess.DoCmd.SetProperty("Подразделение", Access.AcProperty.acPropertyValue, Class1.getStoragePlace())
            Else
                oAccess.DoCmd.SetProperty("ПоказатьЗаявки", Access.AcProperty.acPropertyValue, 1)
            End If
            oAccess.DoCmd.OpenForm("К обновлению исполнения заявок", Microsoft.Office.Interop.Access.AcFormView.acNormal)
            If RadioButton1.Checked Then
                'подразделение
                oAccess.DoCmd.SetProperty("ГруппаПодразделений", Access.AcProperty.acPropertyValue, 1)
            Else
                If RadioButton2.Checked Then
                    'ОСБ Москвы
                    oAccess.DoCmd.SetProperty("ГруппаПодразделений", Access.AcProperty.acPropertyValue, 2)
                Else
                    If RadioButton3.Checked Then
                        'терр. банки
                        oAccess.DoCmd.SetProperty("ГруппаПодразделений", Access.AcProperty.acPropertyValue, 3)
                    Else
                        'всем
                        oAccess.DoCmd.SetProperty("ГруппаПодразделений", Access.AcProperty.acPropertyValue, 4)
                    End If
                End If
            End If
            oAccess.Run("ClickButtonOnUpd")
            oAccess.DoCmd.Quit()
        Catch ex As Exception
            MsgBox("Database error!", MsgBoxStyle.Critical, "Error")
            oAccess.DoCmd.Quit()
        End Try
    End Sub
End Class
