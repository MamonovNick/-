Imports DevExpress.XtraGrid.Views.Base
Imports System.Xml.Serialization
Imports Монеты.MainSettings

Public Class MainForm
    Private FirstOpen As Boolean = True
    Private OperationBool As Boolean = False 'Открытие подменю операций
    Private TabNum As Int16 = -1 'Номер текущей таблицы и пункта меню
    Private PrevTabNum As Int16 = -1 'Номер предыдущей таблицы
    Private MenuPosNum As Int16 = 6 ' Количество позиций меню
    Private MonetType As Int16 = 1 ' Тип загружаемой таблицы для выбора монет(в операциях)
    Private PriceNotEditable As Boolean = True 'При работе с операцияи - возможность изменить цену

    Private Con As New OleDb.OleDbConnection(AppS.ConnStr) ' Переменная для подключения базы
    Private SqlCom As OleDb.OleDbCommand ' Переменная для Sql запросов
    Private delCommand As OleDb.OleDbCommand ' переменная для запроса удаления
    Private updCommand As OleDb.OleDbCommand ' переменная для запроса апдейта
    Private insCommand As OleDb.OleDbCommand ' переменная для запроса вставки

    Private DA As New OleDb.OleDbDataAdapter ' адаптер
    Private DA2 As New OleDb.OleDbDataAdapter ' вспомогательный адаптер

    Private bs1 As New BindingSource() 'Переменная bindingsourse
    Private tbt As New DataTable() ' переменная таблица для вывода в грид
    Private tbt2 As New DataTable() ' переменная табоица для проверки задвоений
    Private tbtMonets As New DataTable() ' переменная таблица для монет
    Private tbtPrices As New DataTable() ' переменная таблица для цен
    Private tbtStores As New DataTable() ' переменная таблица для мест хранения
    Private tbtStores2 As New DataTable() ' переменная таблица для мест хранения
    Private tbtExplan As New DataTable() ' переменная таблица рашифровки к подразделению

    Private LookUp1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
    Private LookUp2rep As New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
    Private LookUp3rep As New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
    Private LookUp4rep As New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
    Private Combo1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private Combo2Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private Combo3Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private Text1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    'Private Text2Rep As New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit

    Private Sub Update_table()
        'обновление БД в соответствии с внесенными изменениями
        Try
            Try
                GridView1.UpdateCurrentRow()
            Catch ex As Exception
            End Try
            DA.Update(tbt)
        Catch e As System.Data.DBConcurrencyException
            MsgBox("Изменения не были сохранены!", MsgBoxStyle.Critical, "Внимание")
        Catch e As Exception
            MsgBox("С обновлением БД все плохо(")
        End Try
    End Sub

    Private Sub Clear_Form()
        'Отмена выделения пунктов меню
        ToolStripButton1.Checked = False
        ToolStripButton3.Checked = False
        ToolStripButton4.Checked = False
        ToolStripButton5.Checked = False
        ToolStripButton6.Checked = False
        ToolStripButton10.Checked = False
        If (TabNum = 1) Or (TabNum > 3) Then
            OperationBool = True
            ToolStripButton2_Click(Nothing, Nothing)
        End If
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Panel1.Visible = False
    End Sub

    Private Sub МонетыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МонетыToolStripMenuItem.Click
        'Справочник монет
        G_coins.Show()
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        'Выход
        Me.Close()
    End Sub

    Private Sub ПодразделенияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПодразделенияToolStripMenuItem.Click
        'Справочник подразделений
        Form4.Show()
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'Only for debugging
        'For tech only 
        'Need to be deleted in release ver
    End Sub

    Private Sub КонтрагентыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КонтрагентыToolStripMenuItem.Click
        'Справочник юридических лиц
        Form5.Show()
    End Sub

    Private Sub ВидыВалютToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВидыВалютToolStripMenuItem.Click
        'Справочник валют
        Form6.Show()
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        'Окно о программе
        AboutBox1.ShowDialog()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        'раскрытие пункта меню для опреций
        If OperationBool Then
            ToolStripButton3.Visible = False
            ToolStripButton4.Visible = False
        Else
            ToolStripButton3.Visible = True
            ToolStripButton4.Visible = True
        End If
        OperationBool = Not OperationBool
    End Sub

    Private Sub GridView1_FocusedRowChanged(sender As Object, e As FocusedRowChangedEventArgs) Handles GridView1.FocusedRowChanged
        Select Case TabNum
            Case 1
                Try
                    If GridView1.FocusedColumn.AbsoluteIndex = 4 Then
                        tbtStores.Reset()
                        tbtStores = Module1.GetTableRegional(GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(3)))
                        LookUp2rep.DataSource = tbtStores
                    End If
                Catch ex As Exception
                End Try
            Case 2
                Try
                    Dim a As Date
                    Try
                        a = GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(1))
                    Catch ex As Exception
                        a = Nothing
                    End Try

                    If a <> Nothing Then
                        PriceNotEditable = True
                        GridView1.Columns(0).OptionsColumn.AllowEdit = False
                        GridView1.Columns(2).OptionsColumn.AllowEdit = False
                        GridView1.Columns(3).OptionsColumn.AllowEdit = False
                        GridView1.Columns(4).OptionsColumn.AllowEdit = False
                        GridView1.Columns(5).OptionsColumn.AllowEdit = False
                        GridView1.Columns(6).OptionsColumn.AllowEdit = False
                        GridView1.Columns(7).OptionsColumn.AllowEdit = False
                        GridView1.Columns(8).OptionsColumn.AllowEdit = False
                        GridView1.Columns(13).OptionsColumn.AllowEdit = False
                        GridView1.Columns(15).OptionsColumn.AllowEdit = False
                        GridView1.Columns(16).OptionsColumn.AllowEdit = False
                    Else
                        PriceNotEditable = False
                        GridView1.Columns(0).OptionsColumn.AllowEdit = True
                        GridView1.Columns(2).OptionsColumn.AllowEdit = True
                        GridView1.Columns(3).OptionsColumn.AllowEdit = True
                        GridView1.Columns(4).OptionsColumn.AllowEdit = True
                        GridView1.Columns(5).OptionsColumn.AllowEdit = True
                        GridView1.Columns(6).OptionsColumn.AllowEdit = True
                        GridView1.Columns(7).OptionsColumn.AllowEdit = True
                        GridView1.Columns(8).OptionsColumn.AllowEdit = True
                        GridView1.Columns(13).OptionsColumn.AllowEdit = True
                        GridView1.Columns(15).OptionsColumn.AllowEdit = True
                        GridView1.Columns(16).OptionsColumn.AllowEdit = True
                    End If
                Catch ex As Exception
                End Try

                Try
                    Select Case GridView1.FocusedColumn.AbsoluteIndex
                        Case 4
                            tbtStores.Reset()
                            tbtStores = Module1.GetTableRegional(GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(3)))
                            LookUp2rep.DataSource = tbtStores
                        Case 5
                            tbtExplan.Reset()
                            tbtExplan = Module1.GetTableExplan(GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(4)))
                            LookUp4rep.DataSource = tbtExplan
                    End Select
                Catch ex As Exception
                End Try

                Try
                    If MonetType = 2 Then
                        tbtMonets.Reset()
                        tbtMonets.Columns.Clear()
                        tbtMonets = Module1.GetTable2(MonetType, GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(4)), IIf(RadioButton1.Checked, "выдача", "приём"))
                        LookUp1Rep.DataSource = tbtMonets
                    End If
                Catch ex As Exception
                End Try
            Case 3

                Try
                    If MonetType = 2 Then
                        tbtMonets.Reset()
                        tbtMonets.Columns.Clear()
                        tbtMonets = Module1.GetTable2(MonetType, GridView1.GetRowCellValue(e.FocusedRowHandle, GridView1.Columns(3)), IIf(RadioButton1.Checked, "продажа", "покупка"))
                        LookUp1Rep.DataSource = tbtMonets
                    End If
                Catch ex As Exception
                End Try
        End Select
    End Sub

    Private Sub GridView1_FocusedColumnChanged(sender As Object, e As FocusedColumnChangedEventArgs) Handles GridView1.FocusedColumnChanged
        Select Case TabNum
            Case 1
                Try
                    If e.FocusedColumn.AbsoluteIndex = 4 Then
                        tbtStores.Reset()
                        tbtStores = Module1.GetTableRegional(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(3)))
                        LookUp2rep.DataSource = tbtStores
                    End If
                Catch ex As Exception
                End Try
            Case 2
                Try
                    Select Case e.FocusedColumn.AbsoluteIndex
                        Case 4
                            tbtStores.Reset()
                            tbtStores = Module1.GetTableRegional(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(3)))
                            LookUp2rep.DataSource = tbtStores
                        Case 5
                            tbtExplan.Reset()
                            tbtExplan = Module1.GetTableExplan(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(4)))
                            LookUp4rep.DataSource = tbtExplan
                    End Select
                Catch ex As Exception
                End Try

                'Try
                '    If MonetType = 2 Then
                '        tbtMonets.Reset()
                '        tbtMonets = Module1.GetTable2(MonetType, GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(4)), IIf(RadioButton1.Checked, "выдача", "приём"))
                '        LookUp1Rep.DataSource = tbtMonets
                '    End If
                'Catch ex As Exception
                'End Try
        End Select
    End Sub

    Private Sub GridView1_CellValueChanging(sender As Object, e As CellValueChangedEventArgs) Handles GridView1.CellValueChanging
        Select Case TabNum
            Case 4, 5, 6
                Select Case e.Column.Name
                    Case "colКаталожныйномер"
                        'Обновление краткого описания монеты (столбец №3)
                        Dim InfoRow As DataRowView = LookUp1Rep.GetRowByKeyValue(e.Value)
                        Dim NewStr As String = InfoRow.Item(1) + " - " + Strings.Right(CStr(InfoRow.Item(2)), 2) + ", " + Strings.Left(InfoRow.Item(5), 2) + ", " + Strings.Left(InfoRow.Item(3), 3) + ", " + CStr(InfoRow.Item(4))
                        GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(3), NewStr)
                End Select
            Case 1
                Select Case e.Column.Name
                    Case "colКаталожныйномер"
                        'Обновление краткого описания монеты (столбец №6)
                        Dim InfoRow As DataRowView = LookUp1Rep.GetRowByKeyValue(e.Value)
                        Dim NewStr As String = InfoRow.Item(1) + " - " + Strings.Right(CStr(InfoRow.Item(2)), 2) + ", " + Strings.Left(InfoRow.Item(5), 2) + ", " + Strings.Left(InfoRow.Item(3), 3) + ", " + CStr(InfoRow.Item(4))
                        GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(6), NewStr)
                End Select
            Case 2
                Select Case e.Column.Name
                    Case "colКаталожныйномер"
                        'Обновление краткого описания монеты (столбец №7)
                        Dim InfoRow As DataRowView = LookUp1Rep.GetRowByKeyValue(e.Value)
                        Try
                            Dim NewStr As String = InfoRow.Item(1) + ", " + Strings.Left(CStr(InfoRow.Item(4)), 2) + ", " + Strings.Left(InfoRow.Item(2), 3) + ", " + CStr(InfoRow.Item(3))
                            GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(7), NewStr)
                            If (Not CheckBox3.Checked) Then
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(11), CStr(InfoRow.Item(5)))
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(8), CStr(InfoRow.Item(6)))
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(14), CStr(InfoRow.Item(7)))
                            Else
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(14), "без заявки")
                            End If
                        Catch ee As Exception
                        End Try
                End Select
            Case 3
                Select Case e.Column.Name
                    Case "colКаталожныйномер"
                        'Обновление краткого описания монеты (столбец №5)
                        Dim InfoRow As DataRowView = LookUp1Rep.GetRowByKeyValue(e.Value)
                        Try
                            Dim NewStr As String = InfoRow.Item(1) + ", " + Strings.Left(CStr(InfoRow.Item(4)), 2) + ", " + Strings.Left(InfoRow.Item(2), 3) + ", " + CStr(InfoRow.Item(3))
                            GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(6), NewStr)
                            If (Not CheckBox3.Checked) Then
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(10), CStr(InfoRow.Item(5)))
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(7), CStr(InfoRow.Item(6)))
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(12), CStr(InfoRow.Item(7)))
                            Else
                                GridView1.SetRowCellValue(e.RowHandle, GridView1.Columns(12), "без заявки")
                            End If
                        Catch ee As Exception
                        End Try
                End Select
        End Select
    End Sub

    Private Sub GridControl1_Click(sender As Object, e As EventArgs) Handles GridControl1.Click
        Try
            Select Case TabNum
                Case 4
                    If GridView1.FocusedColumn.AbsoluteIndex = 5 Then
                        Class1.PutCat(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(2)))
                        Class1.setDate(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1)), GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1)))
                        Class1.setPriceType(1)
                        If Dialog3.ShowDialog() = DialogResult.OK Then
                            tbtPrices.Dispose()
                            tbtPrices = New DataTable
                            tbtPrices = Module1.GetTablePrices(Class1.GetCat, Class1.getDate(1))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(5), tbtPrices.Rows(Class1.getSelectedIndex)(0))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(6), tbtPrices.Rows(Class1.getSelectedIndex)(2))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(7), tbtPrices.Rows(Class1.getSelectedIndex)(3))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(8), tbtPrices.Rows(Class1.getSelectedIndex)(4))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(9), tbtPrices.Rows(Class1.getSelectedIndex)(5))
                        End If
                    Else
                        Return
                    End If
                Case 5
                    If GridView1.FocusedColumn.AbsoluteIndex = 5 Then
                        Class1.PutCat(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(2)))
                        Class1.setDate(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(0)), GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(0)))
                        Class1.setPriceType(2)
                        Class1.setStoragePlace(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1)))
                        If Dialog3.ShowDialog() = DialogResult.OK Then
                            tbtPrices.Dispose()
                            tbtPrices = New DataTable
                            tbtPrices = Module1.GetTablePricesForCond(Class1.GetCat, Class1.getDate(1), Class1.getStoragePlace)
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(5), tbtPrices.Rows(Class1.getSelectedIndex)(0))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(6), tbtPrices.Rows(Class1.getSelectedIndex)(2))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(7), tbtPrices.Rows(Class1.getSelectedIndex)(3))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(8), tbtPrices.Rows(Class1.getSelectedIndex)(4))
                        End If
                    Else
                        Return
                    End If
                Case 2
                    If (Not PriceNotEditable) And (GridView1.FocusedColumn.AbsoluteIndex = 9) Then
                        Class1.PutCat(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(6)))
                        Class1.setDate(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(0)), GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(0)))
                        Class1.setPriceType(3)
                        Class1.setStoragePlace(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(2)))
                        Class1.setFullPrice(CheckBox2.Checked)
                        Try
                            Class1.setState(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(11)))
                        Catch ex As Exception
                            Class1.setState("")
                        End Try
                        If Dialog3.ShowDialog() = DialogResult.OK Then
                            tbtPrices.Dispose()
                            tbtPrices = New DataTable
                            tbtPrices = Module1.GetTablePricesForOperations(Class1.GetCat, Class1.getDate(1), Class1.getStoragePlace, Class1.getFullPrice(), Class1.getState())
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(9), tbtPrices.Rows(Class1.getSelectedIndex)(0))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(10), tbtPrices.Rows(Class1.getSelectedIndex)(2))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(11), tbtPrices.Rows(Class1.getSelectedIndex)(3))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(12), tbtPrices.Rows(Class1.getSelectedIndex)(4))
                        End If
                    Else
                        Return
                    End If
                Case 3
                    If GridView1.FocusedColumn.AbsoluteIndex = 8 Then
                        Class1.PutCat(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(5)))
                        Class1.setDate(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1)), GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1)))
                        Class1.setPriceType(3)
                        Class1.setStoragePlace(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(0)))
                        Class1.setFullPrice(CheckBox2.Checked)
                        Try
                            Class1.setState(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(10)))
                        Catch ex As Exception
                            Class1.setState("")
                        End Try
                        If Dialog3.ShowDialog() = DialogResult.OK Then
                            tbtPrices.Dispose()
                            tbtPrices = New DataTable
                            tbtPrices = Module1.GetTablePricesForOperations(Class1.GetCat, Class1.getDate(1), Class1.getStoragePlace, Class1.getFullPrice(), Class1.getState())
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(8), tbtPrices.Rows(Class1.getSelectedIndex)(0))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(9), tbtPrices.Rows(Class1.getSelectedIndex)(2))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(10), tbtPrices.Rows(Class1.getSelectedIndex)(3))
                            GridView1.SetFocusedRowCellValue(GridView1.Columns(11), tbtPrices.Rows(Class1.getSelectedIndex)(4))
                        End If
                    Else
                        Return
                    End If
            End Select
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridControl1_ProcessGridKey(sender As Object, e As KeyEventArgs) Handles GridControl1.ProcessGridKey
        Dim grid = CType(sender, DevExpress.XtraGrid.GridControl)
        Dim view = CType(grid.FocusedView, DevExpress.XtraGrid.Views.Grid.GridView)
        Select Case TabNum
            Case 2
                Dim a As Date
                Try
                    a = GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(1))
                Catch ex As Exception
                    a = Nothing
                End Try

                If (a = Nothing) And (e.KeyData = Keys.Delete) Then
                    view.DeleteSelectedRows()
                    e.Handled = True
                End If
            Case Else
                If (e.KeyData = Keys.Delete) Then
                    view.DeleteSelectedRows()
                    e.Handled = True
                End If
        End Select
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        'Перемещение монет
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If

        TabNum = 4
        Clear_Form()
        GridView1.Columns.Clear()
        ToolStripButton5.Checked = True

        TableMoveCoinsMake()

        Button1.Visible = True
        Button1.Enabled = True
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Button1.Text = "Сделать спецификацию"

        Panel1.Visible = False
        Panel2.Visible = False
        FlowLayoutPanel1.Visible = True
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Закрытие формы: обновление БД
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not FirstOpen Then
            Update_table()
        End If

        Select Case TabNum
            Case 1
                'загрузить таблицу
                TabNum = 7
                tbt.Reset()
                GridView1.Columns.Clear()
                tbt = New DataTable
                DA = New OleDb.OleDbDataAdapter
                SqlCom = New OleDb.OleDbCommand("TRANSFORM IIf(Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))=0,"""",Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))) AS Выражение2 
SELECT IIf(IsNull([Для свода заявок].Монета), NULL, IIf((Монеты.[Валюта] = ""российск. рубль""),""росс."",""иностр."")) AS Вид, Монеты.Металл, Монеты.Масса, [Для свода заявок].Монета, [Для свода заявок].Состояние
FROM     (Монеты RIGHT OUTER JOIN
                      (SELECT Подразделение, ВидУчастника, [Каталожный номер], Монета, Состояние, Количество, Исполнено
                       FROM      Заявки
                       WHERE   (Дата >= ?) AND (Закрыто = False) AND ([Вид операции] = 'Выдача')
                       UNION ALL
                       SELECT Наименование, ВидУчастника, NULL AS [Каталожный номер], NULL AS Монета, NULL AS Состояние, NULL AS Количество, NULL AS Исполнено
                       FROM     Подразделения
                       WHERE  (Активный = True)) [Для свода заявок] ON Монеты.[Каталожный номер] = [Для свода заявок].[Каталожный номер])
GROUP BY IIf(IsNull([Монета]),Null,IIf([Валюта]=""российск. рубль"",""росс."",""иностр."")), Монеты.Металл, Монеты.Масса, Монеты.Год, [Для свода заявок].[Каталожный номер], [Для свода заявок].Монета, [Для свода заявок].Состояние
ORDER BY IIf(IsNull([Монета]),Null,IIf([Валюта]=""российск. рубль"",""росс."",""иностр."")) DESC , Монеты.Металл DESC , Монеты.Масса, Монеты.Год, IIf([ВидУчастника]=""терр. банк"","" ""+[Подразделение],IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]))
PIVOT IIf([ВидУчастника]=""терр. банк"","" ""+[Подразделение],IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]));", Con)
                DA.SelectCommand = SqlCom
                DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
                DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

                DA.Fill(tbt)
                GridControl1.DataSource = tbt
                GridView1.OptionsBehavior.ReadOnly = True
                TableLayoutPanel1.SetRowSpan(GridControl1, 3)
                Panel1.Enabled = False
                Panel2.Visible = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button1.Text = "Назад"
                GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None
            Case 4
                MsgBox("Данное действие еще не реализовано!", MsgBoxStyle.Critical, "Ошибка")
            Case 7, 8, 9
                ToolStripButton1_Click(sender, e)
        End Select
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        'Состояние монет
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If

        TabNum = 5
        Clear_Form()
        GridView1.Columns.Clear()
        ToolStripButton6.Checked = True

        TableCoinsConditionMake()

        Button1.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        FlowLayoutPanel1.Visible = False
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        'Приобретение монет ТБ в ЦБ
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If

        TabNum = 6
        Clear_Form()
        GridView1.Columns.Clear()
        ToolStripButton10.Checked = True

        TableCoinsPurchaseMake()

        Button1.Visible = False
        Panel1.Visible = False
        Panel2.Visible = False
        FlowLayoutPanel1.Visible = False
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'Обработка пункта меню "Заявки"
        'Обновление предыдущей таблицы, в случае если она была
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If
        PrevTabNum = TabNum 'пока не используется

        Clear_Form() 'отмена выделения пункта меню
        GridView1.Columns.Clear()
        TabNum = 1 'номер текущей таблицы
        ToolStripButton1.Checked = True 'выделяем текущий пункт меню

        TableOrdersMake() ' Заполнение адаптера и таблицы и грида

        ComboBoxEdit1.Visible = True
        Button1.Visible = True
        Button1.Enabled = True
        Button2.Visible = True
        Button2.Enabled = True
        Button3.Visible = True
        Button3.Enabled = True
        Button4.Visible = True
        Button4.Enabled = True
        Button1.Text = """Cвод для ""хвостов"
        Button2.Text = "Свод для новых"
        Button3.Text = "Закрыть заявки"
        Button4.Text = "Искать повторы"

        Panel1.Visible = True
        Panel1.Enabled = True
        Panel2.Visible = False
        FlowLayoutPanel1.Visible = True

        Label1.Text = "Показать заявки с"
        Label2.Visible = False
        CheckBox1.Visible = True
        CheckBox1.Text = "Показать только незакрытые позиции"
        RadioButton1.Text = "Все"
        RadioButton2.Text = "Только"
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'Обработка пункта меню "Внутрисистемные операции"
        'Обновление предыдущей таблицы, в случае если она была
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If

        PrevTabNum = TabNum
        TabNum = 2 'номер текущей таблицы
        Clear_Form() 'отмена выделения пункта меню
        GridView1.Columns.Clear()

        ToolStripButton3.Checked = True 'выделяем текущий пункт меню

        TableOperationsMakeInner() ' Заполнение адаптера, таблицы и грида

        Button1.Visible = True
        Button1.Enabled = True
        Button2.Visible = True
        Button2.Enabled = True
        Button3.Visible = True
        Button3.Enabled = True
        Button4.Visible = False
        Button1.Text = "Искать повторы"
        Button2.Text = "Импорт из Excel"
        Button3.Text = "Рассчитать номер спецификации"

        Panel1.Visible = True
        Panel1.Enabled = True
        Panel2.Visible = True
        FlowLayoutPanel1.Visible = True

        Label1.Text = "Показать операции с"
        CheckBox1.Visible = False
        RadioButton1.Text = "Выдача"
        RadioButton2.Text = "Приём"
        ComboBoxEdit1.Visible = False
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'Обработка пункта меню "Внешние операции"
        'Обновление предыдущей таблицы, в случае если она была
        If Not FirstOpen Then
            Update_table()
        Else
            FirstOpen = False
        End If

        PrevTabNum = TabNum
        TabNum = 3 'номер текущей таблицы
        Clear_Form() 'отмена выделения пункта меню
        GridView1.Columns.Clear()

        ToolStripButton4.Checked = True 'выделяем текущий пункт меню

        TableOperationsMakeOut() ' Заполнение адаптера, таблицы и грида

        Button1.Visible = True
        Button1.Enabled = True
        Button2.Visible = False
        Button3.Visible = True
        Button3.Enabled = True
        Button4.Visible = False
        Button1.Text = "Искать повторы"
        Button3.Text = "Рассчитать номер спецификации"

        Panel1.Visible = True
        Panel1.Enabled = True
        Panel2.Visible = True
        FlowLayoutPanel1.Visible = True

        Label1.Text = "Показать операции с"
        CheckBox1.Visible = False
        RadioButton1.Text = "Продажа"
        RadioButton2.Text = "Покупка"
        ComboBoxEdit1.Visible = False
    End Sub

    Private Sub GridControl1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        Select Case TabNum
            Case 1
                'If e.ColumnIndex = GridControl1.Columns("ВидУчастника").Index Then
                '    comboColumn2.Items.Add(tbt.Rows(e.RowIndex)(e.ColumnIndex))
                'End If
            Case 2

        End Select
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: данная строка кода позволяет загрузить данные в таблицу "МонетыDataSet.Подразделения". При необходимости она может быть перемещена или удалена.
        Me.SecDA.Fill(Me.МонетыDataSet.Подразделения)
        If Not IO.File.Exists(Application.StartupPath + "/settings.xml") Then
            'Start smth)
            MsgBox("Отсутствует файл настроек!", MsgBoxStyle.Exclamation)
        Else
            Dim formatter As XmlSerializer = New XmlSerializer(GetType(AppSettings))
            Dim FileForSerialization As New IO.FileStream("settings.xml", IO.FileMode.OpenOrCreate)
            AppS = formatter.Deserialize(FileForSerialization)
        End If

        'AppS.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Монеты.mdb"
        'AppS.LastVisit = Date.Now
        'AppS.Name = "Nick"
        'AppS.Phone = "985-45"
        'AppS.PicPath = "fkvmdf"
        'Dim b As XmlSerializer = New XmlSerializer(GetType(AppSettings))
        'Dim c As New IO.FileStream("settings.xml", IO.FileMode.Create)
        'b.Serialize(c, AppS)
        'Console.WriteLine(AppS.Phone)

        Module1.Start_Setup()
        TableLayoutPanel1.SetRowSpan(GridControl1, 4)
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Not FirstOpen Then
            Update_table()
        End If
        Select Case TabNum
            Case 1
                TableOrdersMake()
            Case 2
                TableOperationsMakeInner()
            Case 3
                TableOperationsMakeOut()
        End Select
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Not FirstOpen Then
            Update_table()
        End If

        If RadioButton1.Checked Then
            ComboBoxEdit1.Enabled = False
        Else
            ComboBoxEdit1.Enabled = True
        End If

        Select Case TabNum
            Case 1
                TableOrdersMake()
            Case 2
                TableOperationsMakeInner()
        End Select
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.EditValueChanged
        If Not FirstOpen Then
            Update_table()
        End If
        TableOrdersMake()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Not FirstOpen Then
            Update_table()
        End If
        TableOrdersMake()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not FirstOpen Then
            Update_table()
        End If

        Select Case TabNum
            Case 1
                PrevTabNum = TabNum
                TabNum = 8
                tbt.Reset()
                GridView1.Columns.Clear()
                tbt = New DataTable
                DA = New OleDb.OleDbDataAdapter
                SqlCom = New OleDb.OleDbCommand("TRANSFORM IIf(Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))=0,"""",Sum([Количество]-IIf(IsNull([Исполнено]),0,[Исполнено]))) AS Выражение1
SELECT IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]) AS Филиал
FROM Монеты RIGHT JOIN (SELECT Подразделение, ВидУчастника, [Каталожный номер], Монета, Состояние, Количество, Исполнено
                       FROM      Заявки
                       WHERE   (Дата >= ?) AND (Закрыто = False) AND ([Вид операции] = 'Выдача')
                       UNION ALL
                       SELECT Наименование, ВидУчастника, NULL AS [Каталожный номер], NULL AS Монета, NULL AS Состояние, NULL AS Количество, NULL AS Исполнено
                       FROM     Подразделения
                       WHERE  (Активный = True)) [Для свода заявок] ON Монеты.[Каталожный номер] = [Для свода заявок].[Каталожный номер]
GROUP BY IIf([Подразделение]=""Операционное управление"",""ОПЕРУ"",[Подразделение]), IIf([ВидУчастника]=""терр. банк"",1,IIf([ВидУчастника]=""Москва"" And [Подразделение]<>""Московский банк"",2,IIf([Подразделение]=""Московский банк"",3,4)))
ORDER BY IIf([ВидУчастника]=""терр. банк"",1,IIf([ВидУчастника]=""Москва"" And [Подразделение]<>""Московский банк"",2,IIf([Подразделение]=""Московский банк"",3,4)))
PIVOT [Монета]+"", ""+[Состояние];", Con)
                DA.SelectCommand = SqlCom
                DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
                DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

                DA.Fill(tbt)
                GridControl1.DataSource = tbt
                GridView1.OptionsBehavior.ReadOnly = True
                Panel1.Enabled = False
                Panel2.Visible = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button1.Text = "Назад"
                GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not FirstOpen Then
            Update_table()
        End If

        Select Case TabNum
            Case 1
                ' Закрыть заявки
                If Dialog2.ShowDialog() = DialogResult.OK Then
                    Dim updCommand As New OleDb.OleDbCommand("UPDATE Заявки 
SET Закрыто = True 
WHERE ((Закрыто = False) AND (Дата Between ? AND ?))", Con)
                    DA.UpdateCommand = updCommand
                    DA.UpdateCommand.Parameters.Clear()
                    DA.UpdateCommand.Parameters.Add("@Дата1", OleDb.OleDbType.Date, -1, "Дата")
                    DA.UpdateCommand.Parameters(0).Value = Class1.getDate(1)
                    DA.UpdateCommand.Parameters.Add("@Дата2", OleDb.OleDbType.Date, -1, "Дата")
                    DA.UpdateCommand.Parameters(1).Value = Class1.getDate(0)

                    Con.Close()
                    DA.UpdateCommand.Connection = Con
                    DA.UpdateCommand.Connection.Open()
                    DA.UpdateCommand.ExecuteNonQuery()

                    ToolStripButton1_Click(sender, e)
                End If
            Case 2
                Label2.Text = "Следующий номер спецификации: " + CStr(MaxNumSpec())
                Label2.Visible = True
        End Select
    End Sub

    Private Function MaxNumSpec() As Integer
        tbt2.Reset()
        SqlCom = New OleDb.OleDbCommand("SELECT ДатаДенег, Спецификация FROM Операции", Con)
        DA2.SelectCommand = SqlCom
        DA2.Fill(tbt2)

        Dim maxSpec As Integer = 0
        For i = 0 To tbt2.Rows.Count
            Try
                If (CInt(Strings.Left(tbt2.Rows(i)(1), InStr(tbt2.Rows(i)(1), "-") - 1)) > maxSpec) And (Year(tbt2.Rows(i)(0)) = Year(Date.Now())) Then
                    maxSpec = CInt(Strings.Left(tbt2.Rows(i)(1), InStr(tbt2.Rows(i)(1), "-") - 1))
                End If
            Catch e As Exception
            End Try
        Next i
        Return maxSpec
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not FirstOpen Then
            Update_table()
        End If

        Select Case TabNum
            Case 1
                'Искать повторы
                tbt2.Reset()
                SqlCom = New OleDb.OleDbCommand("SELECT * FROM [Проверка повторов в заявках]", Con)
                DA2.SelectCommand = SqlCom

                DA2.Fill(tbt2)
                If tbt2.Rows.Count = 0 Then
                    MsgBox("Повторов не найдено",, "Поиск повторяющихся строк")
                Else
                    MsgBox("Повторы найдены!", MsgBoxStyle.Exclamation, "Поиск повторяющихся строк")
                    GridView1.Columns.Clear()
                    GridControl1.DataSource = tbt2
                    GridView1.OptionsBehavior.ReadOnly = True

                    TabNum = 9
                    Panel1.Enabled = False
                    Panel2.Visible = False
                    Button2.Enabled = False
                    Button3.Enabled = False
                    Button4.Enabled = False
                    Button1.Text = "Назад"
                    GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None
                End If
        End Select
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        'Отчеты
        Form7.Show()
    End Sub

    Private Sub MainForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Module1.Del_Tmp()

    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        'Условные доходы
        Form9.Show()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        'Остатки и цены
        Form11.Show()
    End Sub

    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        'Расчёт сумм распределениия
        Form13.Show()
    End Sub

    Private Sub НастройкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НастройкиToolStripMenuItem.Click
        Form14.ShowDialog()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Select Case TabNum
            Case 2, 3
                MonetType = IIf(CheckBox3.Checked, 1, 2)
                If CheckBox3.Checked Then
                    tbtMonets.Reset()
                    'tbtMonets.Columns.Clear()
                    tbtMonets = Module1.GetTable2(MonetType, "", "")
                    LookUp1Rep.DataSource = tbtMonets
                End If
        End Select
    End Sub

    'Запоолнение таблиц для каждого пункта меню--------------

    Private Sub TableOrdersMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 1 Then
            Return
        End If

        tbt.Dispose()
        tbtMonets.Reset()
        tbtStores.Reset()
        LookUp1Rep.Dispose()
        LookUp2rep.Dispose()
        Combo1Rep.Dispose()
        Combo2Rep.Dispose()
        Combo3Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp2rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        Combo1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Combo2Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Combo3Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox

        SqlCom = New OleDb.OleDbCommand("SELECT Дата, Номер, [Вид операции], ВидУчастника, Подразделение, [Каталожный номер], Монета, Количество, Состояние, Исполнено, Закрыто 
FROM Заявки
WHERE ((Дата >= @Дата)" + IIf(CheckBox1.Checked, " AND (Закрыто = False)", "") + IIf(RadioButton2.Checked, " AND (Подразделение = """ + CStr(ComboBoxEdit1.SelectedText) + """)", "") + ")", Con)
        DA.SelectCommand = SqlCom
        DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

        delCommand = New OleDb.OleDbCommand("DELETE FROM Заявки 
WHERE ((Дата = @Дата) OR Дата IS NULL) AND ((Номер = @Номер) OR Номер IS NULL) AND (([Вид операции] = @Операция) OR [Вид операции] IS NULL) AND ((Подразделение = @Подразделение) OR Подразделение IS NULL) AND (([Каталожный номер] = @КатНомер) OR [Каталожный номер] IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.DeleteCommand.Parameters.Add("@Операция", OleDb.OleDbType.VarChar, 17, "Вид операции")
        DA.DeleteCommand.Parameters.Add("@Подразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.DeleteCommand.Parameters.Add("@КатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")

        updCommand = New OleDb.OleDbCommand("UPDATE Заявки 
SET Дата = qДата, Номер = qНомер, [Вид операции] = qОперация, ВидУчастника = qУчастник, Подразделение = qПодразделение, [Каталожный номер] = qКатНомер, Монета = qМонета, Количество = qКолво, Состояние = qСостояние, Исполнено = qИсполнено, Закрыто = qЗакрыто 
WHERE ((Дата = Дата_Orig) AND ((Номер = Номер_Orig) OR Номер IS NULL) AND ([Вид операции] = Операция_Orig) AND (Подразделение = Подразделение_Orig) AND (([Каталожный номер] = КатНомер_Orig) OR [Каталожный номер] IS NULL))", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("qДата", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Add("qНомер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.UpdateCommand.Parameters.Add("qОперация", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.UpdateCommand.Parameters.Add("qУчастник", OleDb.OleDbType.VarChar, 15, "ВидУчастника")
        DA.UpdateCommand.Parameters.Add("qПодразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.UpdateCommand.Parameters.Add("qКатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Add("qМонета", OleDb.OleDbType.VarChar, 60, "Монета")
        DA.UpdateCommand.Parameters.Add("qКолво", OleDb.OleDbType.Integer, -1, "Количество")
        DA.UpdateCommand.Parameters.Add("qСостояние", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.UpdateCommand.Parameters.Add("qИсполнено", OleDb.OleDbType.Integer, -1, "Исполнено")
        DA.UpdateCommand.Parameters.Add("qЗакрыто", OleDb.OleDbType.Boolean, -1, "Закрыто")
        DA.UpdateCommand.Parameters.Add("Дата_Orig", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Item(11).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Номер_Orig", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.UpdateCommand.Parameters.Item(12).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Операция_Orig", OleDb.OleDbType.VarChar, 17, "Вид операции")
        DA.UpdateCommand.Parameters.Item(13).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("Подразделение_Orig", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.UpdateCommand.Parameters.Item(14).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("КатНомер_Orig", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(15).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO Заявки
                          (Дата, Номер, [Вид операции], ВидУчастника, Подразделение, [Каталожный номер], Монета, Количество, Состояние, Исполнено, Закрыто)
        VALUES (@Дата, @Номер, @Операция, @Участник, @Подразделение, @КатНомер, @Монета, @Колво, @Состояние, @Исполнено, @Закрыто)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.InsertCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 17, "Номер")
        DA.InsertCommand.Parameters.Add("@Операция", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.InsertCommand.Parameters.Add("@Участник", OleDb.OleDbType.VarChar, 15, "ВидУчастника")
        DA.InsertCommand.Parameters.Add("@Подразделение", OleDb.OleDbType.VarChar, 41, "Подразделение")
        DA.InsertCommand.Parameters.Add("@КатНомер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.InsertCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета")
        DA.InsertCommand.Parameters.Add("@Колво", OleDb.OleDbType.Integer, -1, "Количество")
        DA.InsertCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.InsertCommand.Parameters.Add("@Исполнено", OleDb.OleDbType.Integer, -1, "Исполнено")
        DA.InsertCommand.Parameters.Add("@Закрыто", OleDb.OleDbType.Boolean, -1, "Закрыто")

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        tbtMonets = Module1.GetTable("")

        Combo1Rep.Items.AddRange({"выдача", "приём"})
        Combo1Rep.AutoComplete = True
        Combo1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(2).ColumnEdit = Combo1Rep

        Combo2Rep.Items.AddRange({"Москва", "терр. банк", "управление"})
        Combo2Rep.AutoComplete = True
        Combo2Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(3).ColumnEdit = Combo2Rep

        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(5).ColumnEdit = LookUp1Rep

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Наименование"
        LookUp2rep.ValueMember = "Наименование"
        LookUp2rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(4).ColumnEdit = LookUp2rep

        Combo3Rep.Items.AddRange({"отл.", "уд.", "неуд."})
        Combo3Rep.AutoComplete = True
        Combo3Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(8).ColumnEdit = Combo3Rep

        TableLayoutPanel1.SetRowSpan(GridControl1, 2)
        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub CleanMakeForTableOperation()
        tbt.Dispose()
        tbtMonets.Reset()
        tbtStores.Reset()
        tbtStores2.Reset()
        tbtExplan.Reset()
        LookUp1Rep.Dispose()
        LookUp2rep.Dispose()
        LookUp3rep.Dispose()
        LookUp4rep.Dispose()
        Combo1Rep.Dispose()
        Combo2Rep.Dispose()
        Combo3Rep.Dispose()
        Text1Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp2rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp3rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp4rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        Combo1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Combo2Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Combo3Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Text1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
    End Sub

    Private Sub TableOperationsMakeInner()
        'Операции внутренние
        Cursor.Current = Cursors.WaitCursor
        If (TabNum <> 2) Then
            Return
        End If

        CleanMakeForTableOperation()

        SqlCom = New OleDb.OleDbCommand("SELECT ДатаДенег, ДатаМонет, МестоХранения, ВидУчастника, Отделение as Участник, РасшифрПодр as [Расшифровка к подразделению], [Каталожный номер], Монета, Количество, Цена, ВхНДС as НДС, Состояние, Дефекты, Спецификация, Заявка, Комиссия, Распоряжение, [Вид операции] 
FROM Операции 
WHERE ((Операции.ДатаДенег >= @Дата) AND (Операции.[Вид операции] = " + IIf(RadioButton1.Checked, """Выдача""", """Приём""") + "))
ORDER BY Операции.ДатаДенег, Операции.МестоХранения, Операции.ВидУчастника, Операции.Отделение, Операции.Спецификация, Операции.Монета", Con)

        DA.SelectCommand = SqlCom
        DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

        delCommand = New OleDb.OleDbCommand("DELETE FROM Операции
WHERE  ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND ((ДатаДенег = ?) OR ДатаДенег IS NULL) AND ((Отделение = ?) OR Отделение IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND 
       ((Спецификация = ?) OR Спецификация IS NULL) AND ((Распоряжение = ?) OR Распоряжение IS NULL) AND ((Заявка = ?) OR Заявка IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND ((Состояние = ?) OR Состояние IS NULL) AND ((Дефекты = ?) OR Дефекты IS NULL) AND 
       ((Комиссия = ?) OR Комиссия IS NULL) AND ((РасшифрПодр = ?) OR РасшифрПодр IS NULL) AND ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Вид операции] = ?) OR [Вид операции] IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("1", OleDb.OleDbType.Date, -1, "ДатаМонет")
        DA.DeleteCommand.Parameters.Item(0).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("2", OleDb.OleDbType.Date, -1, "ДатаДенег")
        DA.DeleteCommand.Parameters.Item(1).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 41, "Участник")
        DA.DeleteCommand.Parameters.Item(2).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("4", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.DeleteCommand.Parameters.Item(3).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("5", OleDb.OleDbType.Integer, -1, "Количество")
        DA.DeleteCommand.Parameters.Item(4).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("6", OleDb.OleDbType.Double, -1, "Цена")
        DA.DeleteCommand.Parameters.Item(5).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("7", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.DeleteCommand.Parameters.Item(6).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("8", OleDb.OleDbType.SmallInt, -1, "Распоряжение")
        DA.DeleteCommand.Parameters.Item(7).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 50, "Заявка")
        DA.DeleteCommand.Parameters.Item(8).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("10", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.DeleteCommand.Parameters.Item(9).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("11", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.DeleteCommand.Parameters.Item(10).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("12", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.DeleteCommand.Parameters.Item(11).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("13", OleDb.OleDbType.Boolean, -1, "Комиссия")
        DA.DeleteCommand.Parameters.Item(12).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("14", OleDb.OleDbType.VarChar, 60, "Расшифровка к подразделению")
        DA.DeleteCommand.Parameters.Item(13).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("15", OleDb.OleDbType.VarChar, 5, "МестоХранения")
        DA.DeleteCommand.Parameters.Item(14).SourceVersion = DataRowVersion.Original
        DA.DeleteCommand.Parameters.Add("16", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.DeleteCommand.Parameters.Item(15).SourceVersion = DataRowVersion.Original

        updCommand = New OleDb.OleDbCommand("UPDATE Операции
SET  ДатаМонет = ?, ДатаДенег = ?, [Вид операции] = ?, ВидУчастника = ?, Отделение = ?, [Каталожный номер] = ?, Монета = ?, Количество = ?, Цена = ?, Спецификация = ?, Распоряжение = ?, Заявка = ?, 
     ВхНДС = ?, Состояние = ?, Дефекты = ?, Комиссия = ?, РасшифрПодр = ?, МестоХранения = ?
WHERE  ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND ((ДатаДенег = ?) OR ДатаДенег IS NULL) AND ((Отделение = ?) OR Отделение IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND 
       ((Спецификация = ?) OR Спецификация IS NULL) AND ((Распоряжение = ?) OR Распоряжение IS NULL) AND ((Заявка = ?) OR Заявка IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND ((Состояние = ?) OR Состояние IS NULL) AND 
       ((Дефекты = ?) OR Дефекты IS NULL) AND ((Комиссия = ?) OR Комиссия IS NULL) AND ((РасшифрПодр = ?) OR РасшифрПодр IS NULL) AND ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Вид операции] = ?) OR [Вид операции] IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("s1", OleDb.OleDbType.Date, -1, "ДатаМонет") '0
        DA.UpdateCommand.Parameters.Add("s2", OleDb.OleDbType.Date, -1, "ДатаДенег") '1
        DA.UpdateCommand.Parameters.Add("s3", OleDb.OleDbType.VarChar, 10, "Вид операции") '2
        DA.UpdateCommand.Parameters.Add("s4", OleDb.OleDbType.VarChar, 15, "ВидУчастника") '3
        DA.UpdateCommand.Parameters.Add("s5", OleDb.OleDbType.VarChar, 41, "Участник") '4
        DA.UpdateCommand.Parameters.Add("s6", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '5
        DA.UpdateCommand.Parameters.Add("s7", OleDb.OleDbType.VarChar, 60, "Монета") '6
        DA.UpdateCommand.Parameters.Add("s8", OleDb.OleDbType.Integer, -1, "Количество") '7
        DA.UpdateCommand.Parameters.Add("s9", OleDb.OleDbType.Double, -1, "Цена") '8
        DA.UpdateCommand.Parameters.Add("s10", OleDb.OleDbType.VarChar, 7, "Спецификация") '9
        DA.UpdateCommand.Parameters.Add("s11", OleDb.OleDbType.SmallInt, -1, "Распоряжение") '10
        DA.UpdateCommand.Parameters.Add("s12", OleDb.OleDbType.VarChar, 50, "Заявка") '11
        DA.UpdateCommand.Parameters.Add("s13", OleDb.OleDbType.VarChar, 12, "НДС") '12
        DA.UpdateCommand.Parameters.Add("s14", OleDb.OleDbType.VarChar, 5, "Состояние") '13
        DA.UpdateCommand.Parameters.Add("s15", OleDb.OleDbType.VarChar, 60, "Дефекты") '14
        DA.UpdateCommand.Parameters.Add("s16", OleDb.OleDbType.Boolean, -1, "Комиссия") '15
        DA.UpdateCommand.Parameters.Add("s17", OleDb.OleDbType.VarChar, 60, "Расшифровка к подразделению") '16
        DA.UpdateCommand.Parameters.Add("s18", OleDb.OleDbType.VarChar, 5, "МестоХранения") '17

        DA.UpdateCommand.Parameters.Add("1", OleDb.OleDbType.Date, -1, "ДатаМонет")
        DA.UpdateCommand.Parameters.Item(18).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("2", OleDb.OleDbType.Date, -1, "ДатаДенег")
        DA.UpdateCommand.Parameters.Item(19).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 41, "Участник")
        DA.UpdateCommand.Parameters.Item(20).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("4", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(21).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("5", OleDb.OleDbType.Integer, -1, "Количество")
        DA.UpdateCommand.Parameters.Item(22).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("6", OleDb.OleDbType.Double, -1, "Цена")
        DA.UpdateCommand.Parameters.Item(23).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("7", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.UpdateCommand.Parameters.Item(24).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("8", OleDb.OleDbType.SmallInt, -1, "Распоряжение")
        DA.UpdateCommand.Parameters.Item(25).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 50, "Заявка")
        DA.UpdateCommand.Parameters.Item(26).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("10", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.UpdateCommand.Parameters.Item(27).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("11", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.UpdateCommand.Parameters.Item(28).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("12", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.UpdateCommand.Parameters.Item(29).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("13", OleDb.OleDbType.Boolean, -1, "Комиссия")
        DA.UpdateCommand.Parameters.Item(30).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("14", OleDb.OleDbType.VarChar, 60, "Расшифровка к подразделению")
        DA.UpdateCommand.Parameters.Item(31).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("15", OleDb.OleDbType.VarChar, 5, "МестоХранения")
        DA.UpdateCommand.Parameters.Item(32).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("16", OleDb.OleDbType.VarChar, 10, "Вид операции")
        DA.UpdateCommand.Parameters.Item(33).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO Операции 
        (ДатаМонет, ДатаДенег, [Вид операции], ВидУчастника, Отделение, [Каталожный номер], Монета, Количество, Цена, Спецификация, Распоряжение, Заявка, ВхНДС, Состояние, Дефекты, Комиссия, 
        РасшифрПодр, МестоХранения) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("s1", OleDb.OleDbType.Date, -1, "ДатаМонет")
        DA.InsertCommand.Parameters.Add("s2", OleDb.OleDbType.Date, -1, "ДатаДенег")
        DA.InsertCommand.Parameters.AddWithValue("s3", IIf(RadioButton1.Checked, "выдача", "приём"))
        DA.InsertCommand.Parameters.Add("s4", OleDb.OleDbType.VarChar, 15, "ВидУчастника")
        DA.InsertCommand.Parameters.Add("s5", OleDb.OleDbType.VarChar, 41, "Участник")
        DA.InsertCommand.Parameters.Add("s6", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.InsertCommand.Parameters.Add("s7", OleDb.OleDbType.VarChar, 60, "Монета")
        DA.InsertCommand.Parameters.Add("s8", OleDb.OleDbType.Integer, -1, "Количество")
        DA.InsertCommand.Parameters.Add("s9", OleDb.OleDbType.Double, -1, "Цена")
        DA.InsertCommand.Parameters.Add("s10", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.InsertCommand.Parameters.Add("s11", OleDb.OleDbType.SmallInt, -1, "Распоряжение")
        DA.InsertCommand.Parameters.Add("s12", OleDb.OleDbType.VarChar, 50, "Заявка")
        DA.InsertCommand.Parameters.Add("s13", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.InsertCommand.Parameters.Add("s14", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.InsertCommand.Parameters.Add("s15", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.InsertCommand.Parameters.Add("s16", OleDb.OleDbType.Boolean, -1, "Комиссия")
        DA.InsertCommand.Parameters.Add("s17", OleDb.OleDbType.VarChar, 60, "Расшифровка к подразделению")
        DA.InsertCommand.Parameters.Add("s18", OleDb.OleDbType.VarChar, 5, "МестоХранения")

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        MonetType = IIf(CheckBox3.Checked, 1, 2)
        tbtMonets = Module1.GetTable2(MonetType, "", "")
        tbtStores2 = Module1.GetTableStores()

        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.NullText = Nothing
        LookUp1Rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(6).ColumnEdit = LookUp1Rep

        LookUp3rep.DataSource = tbtStores2
        LookUp3rep.AutoComplete = True
        LookUp3rep.DisplayMember = "Обозначение"
        LookUp3rep.ValueMember = "Обозначение"
        LookUp3rep.NullText = Nothing
        LookUp3rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp3rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(2).ColumnEdit = LookUp3rep

        Combo2Rep.Items.AddRange({"Москва", "терр. банк", "управление"})
        Combo2Rep.AutoComplete = True
        Combo2Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(3).ColumnEdit = Combo2Rep

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Наименование"
        LookUp2rep.ValueMember = "Наименование"
        LookUp2rep.NullText = Nothing
        LookUp2rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(4).ColumnEdit = LookUp2rep

        LookUp4rep.DataSource = tbtExplan
        LookUp4rep.AutoComplete = True
        LookUp4rep.DisplayMember = "Расшифровка"
        LookUp4rep.ValueMember = "Расшифровка"
        LookUp4rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp4rep.NullText = ""
        LookUp4rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp4rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(5).ColumnEdit = LookUp4rep

        GridView1.Columns(9).OptionsColumn.AllowEdit = False 'запрет редактирования цены, обработка по клику
        GridView1.Columns(10).OptionsColumn.AllowEdit = False 'запрет редактирования НДС
        GridView1.Columns(11).OptionsColumn.AllowEdit = False 'запрет редактирования Состояния
        GridView1.Columns(12).OptionsColumn.AllowEdit = False 'запрет редактирования Дефектов
        GridView1.Columns(14).OptionsColumn.AllowEdit = False 'запрет редактирования Заявок

        If RadioButton1.Checked Then
            Text1Rep.NullText = "выдача"
        Else
            Text1Rep.NullText = "приём"
        End If
        GridView1.Columns(17).ColumnEdit = Text1Rep
        GridView1.Columns(17).OptionsColumn.AllowEdit = False

        TableLayoutPanel1.SetRowSpan(GridControl1, 2)
        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsBehavior.KeepFocusedRowOnUpdate = True
        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub TableOperationsMakeOut()
        'Операции внешние
        Cursor.Current = Cursors.WaitCursor
        If (TabNum <> 3) Then
            Return
        End If

        CleanMakeForTableOperation()

        SqlCom = New OleDb.OleDbCommand("SELECT МестоХранения, ДатаДенег, ДатаМонет, Отделение as Контрагент, Спецификация, [Каталожный номер], Монета, Количество, Цена, ВхНДС as НДС, Состояние, Дефекты, Заявка as Основание 
FROM Операции 
WHERE ((Операции.ДатаДенег >= @Дата) AND (Операции.[Вид операции] = " + IIf(RadioButton1.Checked, """продажа""", """покупка""") + "))
ORDER BY Операции.ДатаДенег, Операции.Спецификация, Операции.Монета", Con)
        DA.SelectCommand = SqlCom
        DA.SelectCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.SelectCommand.Parameters(0).Value = DateTimePicker1.Value.Date()

        delCommand = New OleDb.OleDbCommand("DELETE FROM Операции
WHERE  ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND ((ДатаДенег = ?) OR ДатаДенег IS NULL) AND ((Отделение = ?) OR Отделение IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND 
       ((Спецификация = ?) OR Спецификация IS NULL) AND ((Заявка = ?) OR Заявка IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND ((Состояние = ?) OR Состояние IS NULL) AND ((Дефекты = ?) OR Дефекты IS NULL) AND 
       ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Вид операции] = ?) OR [Вид операции] IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("1", OleDb.OleDbType.Date, -1, "ДатаМонет")
        DA.DeleteCommand.Parameters.Add("2", OleDb.OleDbType.Date, -1, "ДатаДенег")
        DA.DeleteCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 41, "Конртагент")
        DA.DeleteCommand.Parameters.Add("4", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.DeleteCommand.Parameters.Add("5", OleDb.OleDbType.Integer, -1, "Количество")
        DA.DeleteCommand.Parameters.Add("6", OleDb.OleDbType.Double, -1, "Цена")
        DA.DeleteCommand.Parameters.Add("7", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.DeleteCommand.Parameters.Add("8", OleDb.OleDbType.VarChar, 50, "Основание")
        DA.DeleteCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.DeleteCommand.Parameters.Add("10", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.DeleteCommand.Parameters.Add("11", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.DeleteCommand.Parameters.Add("12", OleDb.OleDbType.VarChar, 5, "МестоХранения")
        DA.DeleteCommand.Parameters.AddWithValue("13", IIf(RadioButton1.Checked, "продажа", "покупка"))

        updCommand = New OleDb.OleDbCommand("UPDATE Операции
SET  ДатаМонет = ?, ДатаДенег = ?, [Вид операции] = ?, ВидУчастника = ?, Отделение = ?, [Каталожный номер] = ?, Монета = ?, Количество = ?, Цена = ?, Спецификация = ?, Распоряжение = ?, Заявка = ?, 
     ВхНДС = ?, Состояние = ?, Дефекты = ?, Комиссия = ?, РасшифрПодр = ?, МестоХранения = ?
WHERE  ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND ((ДатаДенег = ?) OR ДатаДенег IS NULL) AND ((Отделение = ?) OR Отделение IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND 
       ((Спецификация = ?) OR Спецификация IS NULL) AND ((Заявка = ?) OR Заявка IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND ((Состояние = ?) OR Состояние IS NULL) AND ((Дефекты = ?) OR Дефекты IS NULL) AND 
       ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Вид операции] = ?) OR [Вид операции] IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("s1", OleDb.OleDbType.Date, -1, "ДатаМонет") '0
        DA.UpdateCommand.Parameters.Add("s2", OleDb.OleDbType.Date, -1, "ДатаДенег") '1
        DA.UpdateCommand.Parameters.AddWithValue("s3", IIf(RadioButton1.Checked, "продажа", "покупка")) '2
        DA.UpdateCommand.Parameters.AddWithValue("s4", "юр. лицо") '3
        DA.UpdateCommand.Parameters.Add("s5", OleDb.OleDbType.VarChar, 41, "Контрагент") '4
        DA.UpdateCommand.Parameters.Add("s6", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '5
        DA.UpdateCommand.Parameters.Add("s7", OleDb.OleDbType.VarChar, 60, "Монета") '6
        DA.UpdateCommand.Parameters.Add("s8", OleDb.OleDbType.Integer, -1, "Количество") '7
        DA.UpdateCommand.Parameters.Add("s9", OleDb.OleDbType.Double, -1, "Цена") '8
        DA.UpdateCommand.Parameters.Add("s10", OleDb.OleDbType.VarChar, 7, "Спецификация") '9
        DA.UpdateCommand.Parameters.AddWithValue("s11", "") '10
        DA.UpdateCommand.Parameters.Add("s12", OleDb.OleDbType.VarChar, 50, "Основание") '11
        DA.UpdateCommand.Parameters.Add("s13", OleDb.OleDbType.VarChar, 12, "НДС") '12
        DA.UpdateCommand.Parameters.Add("s14", OleDb.OleDbType.VarChar, 5, "Состояние") '13
        DA.UpdateCommand.Parameters.Add("s15", OleDb.OleDbType.VarChar, 60, "Дефекты") '14
        DA.UpdateCommand.Parameters.AddWithValue("s16", False) '15
        DA.UpdateCommand.Parameters.AddWithValue("13", "") '16
        DA.UpdateCommand.Parameters.Add("s18", OleDb.OleDbType.VarChar, 5, "МестоХранения") '17

        DA.UpdateCommand.Parameters.Add("1", OleDb.OleDbType.Date, -1, "ДатаМонет")
        DA.UpdateCommand.Parameters.Item(18).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("2", OleDb.OleDbType.Date, -1, "ДатаДенег")
        DA.UpdateCommand.Parameters.Item(19).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 41, "Конртагент")
        DA.UpdateCommand.Parameters.Item(20).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("4", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(21).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("5", OleDb.OleDbType.Integer, -1, "Количество")
        DA.UpdateCommand.Parameters.Item(22).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("6", OleDb.OleDbType.Double, -1, "Цена")
        DA.UpdateCommand.Parameters.Item(23).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("7", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.UpdateCommand.Parameters.Item(24).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("8", OleDb.OleDbType.VarChar, 50, "Основание")
        DA.UpdateCommand.Parameters.Item(25).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.UpdateCommand.Parameters.Item(26).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("10", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.UpdateCommand.Parameters.Item(27).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("11", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.UpdateCommand.Parameters.Item(28).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("12", OleDb.OleDbType.VarChar, 5, "МестоХранения")
        DA.UpdateCommand.Parameters.Item(29).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.AddWithValue("13", IIf(RadioButton1.Checked, "продажа", "покупка"))
        DA.UpdateCommand.Parameters.Item(30).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO Операции 
        (ДатаМонет, ДатаДенег, [Вид операции], ВидУчастника, Отделение, [Каталожный номер], Монета, Количество, Цена, Спецификация, Распоряжение, Заявка, ВхНДС, Состояние, Дефекты, Комиссия, 
        РасшифрПодр, МестоХранения) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("s1", OleDb.OleDbType.Date, -1, "ДатаМонет") '0
        DA.InsertCommand.Parameters.Add("s2", OleDb.OleDbType.Date, -1, "ДатаДенег") '1
        DA.InsertCommand.Parameters.AddWithValue("s3", IIf(RadioButton1.Checked, "продажа", "покупка")) '2
        DA.InsertCommand.Parameters.AddWithValue("s4", "юр. лицо") '3
        DA.InsertCommand.Parameters.Add("s5", OleDb.OleDbType.VarChar, 41, "Контрагент") '4
        DA.InsertCommand.Parameters.Add("s6", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '5
        DA.InsertCommand.Parameters.Add("s7", OleDb.OleDbType.VarChar, 60, "Монета") '6
        DA.InsertCommand.Parameters.Add("s8", OleDb.OleDbType.Integer, -1, "Количество") '7
        DA.InsertCommand.Parameters.Add("s9", OleDb.OleDbType.Double, -1, "Цена") '8
        DA.InsertCommand.Parameters.Add("s10", OleDb.OleDbType.VarChar, 7, "Спецификация") '9
        DA.InsertCommand.Parameters.AddWithValue("s11", "") '10
        DA.InsertCommand.Parameters.Add("s12", OleDb.OleDbType.VarChar, 50, "Основание") '11
        DA.InsertCommand.Parameters.Add("s13", OleDb.OleDbType.VarChar, 12, "НДС") '12
        DA.InsertCommand.Parameters.Add("s14", OleDb.OleDbType.VarChar, 5, "Состояние") '13
        DA.InsertCommand.Parameters.Add("s15", OleDb.OleDbType.VarChar, 60, "Дефекты") '14
        DA.InsertCommand.Parameters.AddWithValue("s16", False) '15
        DA.InsertCommand.Parameters.AddWithValue("13", "") '16
        DA.InsertCommand.Parameters.Add("s18", OleDb.OleDbType.VarChar, 5, "МестоХранения") '17

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        MonetType = IIf(CheckBox3.Checked, 1, 2)
        tbtMonets = Module1.GetTable2(MonetType, "", "")
        tbtStores = Module1.GetTableContr()
        tbtStores2 = Module1.GetTableStores()

        LookUp3rep.DataSource = tbtStores2
        LookUp3rep.AutoComplete = True
        LookUp3rep.DisplayMember = "Обозначение"
        LookUp3rep.ValueMember = "Обозначение"
        LookUp3rep.NullText = Nothing
        LookUp3rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp3rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(0).ColumnEdit = LookUp3rep

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Наименование"
        LookUp2rep.ValueMember = "Наименование"
        LookUp2rep.NullText = Nothing
        LookUp2rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(3).ColumnEdit = LookUp2rep

        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.NullText = Nothing
        LookUp1Rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(5).ColumnEdit = LookUp1Rep

        Combo1Rep.Items.AddRange({"18%", "без номинала", "16.67%", "нет"})
        Combo1Rep.AutoComplete = True
        Combo1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(9).ColumnEdit = Combo1Rep

        Combo2Rep.Items.AddRange({"отл.", "уд.", "неуд."})
        Combo2Rep.AutoComplete = True
        Combo2Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(10).ColumnEdit = Combo2Rep

        GridView1.Columns(8).OptionsColumn.AllowEdit = False 'запрет редактирования цены, обработка по клику

        TableLayoutPanel1.SetRowSpan(GridControl1, 2)
        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsBehavior.KeepFocusedRowOnUpdate = True
        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub TableMoveCoinsMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 4 Then
            Return
        End If

        tbt.Dispose()
        tbtMonets.Reset()
        tbtStores.Reset()
        LookUp1Rep.Dispose()
        LookUp2rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp2rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit

        SqlCom = New OleDb.OleDbCommand("SELECT Завершено, Дата, [Каталожный номер], Монета, Количество, Цена, ВхНДС as НДС, Состояние, Дефекты, МестоХраненияСтарое as [Старое место хранения], МестоХраненияНовое as [Новое место хранения], Спецификация 
FROM [Перемещение между хранилищами]", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM [Перемещение между хранилищами] 
WHERE ((Дата = @Дата) OR Дата IS NULL) AND (([Каталожный номер] = @Номер) OR [Каталожный номер] IS NULL) AND ((Количество = @Колво) OR Количество IS NULL) AND ((Цена = @Цена) OR Цена IS NULL) AND ((ВхНДС = @НДС) OR ВхНДС IS NULL) AND 
      ((Состояние = @Состояние) OR Состояние IS NULL) AND ((Дефекты = @Дефекты) OR Дефекты IS NULL) AND ((МестоХраненияСтарое = @СМХ) OR МестоХраненияСтарое IS NULL) AND ((МестоХраненияНовое = @НМХ) OR МестоХраненияНовое IS NULL) AND ((Спецификация = @Спецификация) OR Спецификация IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.DeleteCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.DeleteCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена")
        DA.DeleteCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.DeleteCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.DeleteCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.DeleteCommand.Parameters.Add("@СМХ", OleDb.OleDbType.VarChar, 5, "Старое место хранения")
        DA.DeleteCommand.Parameters.Add("@НМХ", OleDb.OleDbType.VarChar, 5, "Новое место хранения")
        DA.DeleteCommand.Parameters.Add("@Спецификация", OleDb.OleDbType.VarChar, 7, "Спецификация")

        updCommand = New OleDb.OleDbCommand("UPDATE [Перемещение между хранилищами] 
SET Завершено = @Завершено, Дата = @Дата, [Каталожный номер] = @Номер, Монета = @Монета, Количество = @Колво, Цена = @Цена, ВхНДС = @НДС, Состояние = @Состояние, Дефекты = @Дефекты, МестоХраненияСтарое = @СМХ, МестоХраненияНовое = @НМХ, Спецификация = @Спецификация 
WHERE ((Дата = @ДатаО) OR Дата IS NULL) AND (([Каталожный номер] = @НомерО) OR [Каталожный номер] IS NULL) AND ((Количество = @КолвоО) OR Количество IS NULL) AND ((Цена = @ЦенаО) OR Цена IS NULL) AND ((ВхНДС = @НДСО) OR ВхНДС IS NULL) AND 
      ((Состояние = @СостояниеО) OR Состояние IS NULL) AND ((Дефекты = @ДефектыО) OR Дефекты IS NULL) AND ((МестоХраненияСтарое = @СМХО) OR МестоХраненияСтарое IS NULL) AND ((МестоХраненияНовое = @НМХО) OR МестоХраненияНовое IS NULL) AND ((Спецификация = @СпецификацияО) OR Спецификация IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("@Завершено", OleDb.OleDbType.Boolean, -1, "Завершено") '0
        DA.UpdateCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '1
        DA.UpdateCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '2
        DA.UpdateCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '3
        DA.UpdateCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '4
        DA.UpdateCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '5
        DA.UpdateCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС") '6
        DA.UpdateCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Состояние") '7
        DA.UpdateCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Дефекты") '8
        DA.UpdateCommand.Parameters.Add("@СМХ", OleDb.OleDbType.VarChar, 5, "Старое место хранения") '9
        DA.UpdateCommand.Parameters.Add("@НМХ", OleDb.OleDbType.VarChar, 5, "Новое место хранения") '10
        DA.UpdateCommand.Parameters.Add("@Спецификация", OleDb.OleDbType.VarChar, 7, "Спецификация") '11

        DA.UpdateCommand.Parameters.Add("@ДатаО", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Item(12).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НомерО", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(13).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@КолвоО", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.UpdateCommand.Parameters.Item(14).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ЦенаО", OleDb.OleDbType.Double, -1, "Цена")
        DA.UpdateCommand.Parameters.Item(15).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НДСО", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.UpdateCommand.Parameters.Item(16).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@СостояниеО", OleDb.OleDbType.VarChar, 5, "Состояние")
        DA.UpdateCommand.Parameters.Item(17).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ДефектыО", OleDb.OleDbType.VarChar, 60, "Дефекты")
        DA.UpdateCommand.Parameters.Item(18).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@СМХО", OleDb.OleDbType.VarChar, 5, "Старое место хранения")
        DA.UpdateCommand.Parameters.Item(19).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НМХО", OleDb.OleDbType.VarChar, 5, "Новое место хранения")
        DA.UpdateCommand.Parameters.Item(20).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@СпецификацияО", OleDb.OleDbType.VarChar, 7, "Спецификация")
        DA.UpdateCommand.Parameters.Item(21).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO [Перемещение между хранилищами] 
        (Завершено, Дата, [Каталожный номер], Монета, Количество, Цена, ВхНДС, Состояние, Дефекты, МестоХраненияСтарое, МестоХраненияНовое, Спецификация) 
        VALUES (@Завершено, @Дата, @Номер, @Монета, @Колво, @Цена, @НДС, @Состояние, @Дефекты, @СМХ, @НМХ, @Спецификация)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("@Завершено", OleDb.OleDbType.Boolean, -1, "Завершено") '0
        DA.InsertCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '1
        DA.InsertCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '2
        DA.InsertCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '3
        DA.InsertCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '4
        DA.InsertCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '5
        DA.InsertCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС") '6
        DA.InsertCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Состояние") '7
        DA.InsertCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Дефекты") '8
        DA.InsertCommand.Parameters.Add("@СМХ", OleDb.OleDbType.VarChar, 5, "Старое место хранения") '9
        DA.InsertCommand.Parameters.Add("@НМХ", OleDb.OleDbType.VarChar, 5, "Новое место хранения") '10
        DA.InsertCommand.Parameters.Add("@Спецификация", OleDb.OleDbType.VarChar, 7, "Спецификация") '11

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        tbtMonets = Module1.GetTable("")
        tbtStores = Module1.GetTableStores()

        TableLayoutPanel1.SetRowSpan(GridControl1, 3)
        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(2).ColumnEdit = LookUp1Rep

        GridView1.Columns(5).OptionsColumn.AllowEdit = False 'запрет редактирования цены, обработка по клику
        GridView1.Columns(6).OptionsColumn.AllowEdit = False 'запрет редактирования НДС
        GridView1.Columns(7).OptionsColumn.AllowEdit = False 'запрет редактирования Состояния
        GridView1.Columns(8).OptionsColumn.AllowEdit = False 'запрет редактирования Дефектов
        GridView1.Columns(9).OptionsColumn.AllowEdit = False 'запрет редактирования Старого места хранения

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Обозначение"
        LookUp2rep.ValueMember = "Обозначение"
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(10).ColumnEdit = LookUp2rep

        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub TableCoinsConditionMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 5 Then
            Return
        End If

        tbt.Dispose()
        tbtMonets.Reset()
        tbtStores.Reset()
        LookUp1Rep.Dispose()
        LookUp2rep.Dispose()
        Combo1Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp2rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        Combo1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox

        SqlCom = New OleDb.OleDbCommand("SELECT Дата, МестоХранения AS [Место хранения], [Каталожный номер], Монета, Количество, Цена, ВхНДС AS НДС, СтарСостояние AS Состояние, СтарДефекты AS [Старые дефекты], НовСостояние AS [Новое состояние], НовДефекты AS [Новые дефекты], Основание 
FROM [Изменение состояния]", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM [Изменение состояния] 
WHERE ((Дата = ?) OR Дата IS NULL) AND ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND 
      ((НовСостояние = ?) OR НовСостояние IS NULL) AND ((НовДефекты = ?) OR НовДефекты IS NULL) AND ((Основание = ?) OR Основание IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@Хран", OleDb.OleDbType.VarChar, 5, "Место хранения")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.DeleteCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.DeleteCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена")
        DA.DeleteCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.DeleteCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Новое состояние")
        DA.DeleteCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Новые дефекты")
        DA.DeleteCommand.Parameters.Add("@Основание", OleDb.OleDbType.VarChar, 50, "Основание")

        updCommand = New OleDb.OleDbCommand("UPDATE [Изменение состояния] 
SET Дата = ?, [Каталожный номер] = ?, Монета = ?, Количество = ?, Цена = ?, ВхНДС = ?, СтарСостояние = ?, СтарДефекты = ?, НовСостояние = ?, НовДефекты = ?, Основание = ?, МинуяЦА = False, МестоХранения = ?
WHERE ((Дата = ?) OR Дата IS NULL) AND ((МестоХранения = ?) OR МестоХранения IS NULL) AND (([Каталожный номер] =?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL) AND ((ВхНДС = ?) OR ВхНДС IS NULL) AND 
      ((НовСостояние = ?) OR НовСостояние IS NULL) AND ((НовДефекты = ?) OR НовДефекты IS NULL) AND ((Основание = ?) OR Основание IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '0
        DA.UpdateCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '1
        DA.UpdateCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '2
        DA.UpdateCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '3
        DA.UpdateCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '4
        DA.UpdateCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС") '5
        DA.UpdateCommand.Parameters.Add("@СтарСостояние", OleDb.OleDbType.VarChar, 5, "Состояние") '6
        DA.UpdateCommand.Parameters.Add("@СтарДефекты", OleDb.OleDbType.VarChar, 60, "Старые дефекты") '7
        DA.UpdateCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Новое состояние") '8
        DA.UpdateCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Новые дефекты") '9
        DA.UpdateCommand.Parameters.Add("@Основание", OleDb.OleDbType.VarChar, 50, "Основание") '10
        DA.UpdateCommand.Parameters.Add("@Хран", OleDb.OleDbType.VarChar, 5, "Место хранения") '11
        DA.UpdateCommand.Parameters.Add("@ДатаО", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Item(12).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ХранО", OleDb.OleDbType.VarChar, 5, "Место хранения")
        DA.UpdateCommand.Parameters.Item(13).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НомерО", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(14).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@КолвоО", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.UpdateCommand.Parameters.Item(15).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ЦенаО", OleDb.OleDbType.Double, -1, "Цена")
        DA.UpdateCommand.Parameters.Item(16).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НДСО", OleDb.OleDbType.VarChar, 12, "НДС")
        DA.UpdateCommand.Parameters.Item(17).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@СостояниеО", OleDb.OleDbType.VarChar, 5, "Новое состояние")
        DA.UpdateCommand.Parameters.Item(18).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ДефектыО", OleDb.OleDbType.VarChar, 60, "Новые дефекты")
        DA.UpdateCommand.Parameters.Item(19).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@ОснованиеО", OleDb.OleDbType.VarChar, 50, "Основание")
        DA.UpdateCommand.Parameters.Item(20).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO [Изменение состояния] 
        (Дата, [Каталожный номер], Монета, Количество, Цена, ВхНДС, СтарСостояние, СтарДефекты, НовСостояние, НовДефекты, Основание, МинуяЦА, МестоХранения) 
        VALUES (@Дата, @Номер, @Монета, @Колво, @Цена, @НДС, @СтарСостояние, @СтарДефекты, @Состояние, @Дефекты, @Основание, False, @Хран)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '0
        DA.InsertCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '1
        DA.InsertCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '2
        DA.InsertCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '3
        DA.InsertCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '4
        DA.InsertCommand.Parameters.Add("@НДС", OleDb.OleDbType.VarChar, 12, "НДС") '5
        DA.InsertCommand.Parameters.Add("@СтарСостояние", OleDb.OleDbType.VarChar, 5, "Состояние") '6
        DA.InsertCommand.Parameters.Add("@СтарДефекты", OleDb.OleDbType.VarChar, 60, "Старые дефекты") '7
        DA.InsertCommand.Parameters.Add("@Состояние", OleDb.OleDbType.VarChar, 5, "Новое состояние") '8
        DA.InsertCommand.Parameters.Add("@Дефекты", OleDb.OleDbType.VarChar, 60, "Новые дефекты") '9
        DA.InsertCommand.Parameters.Add("@Основание", OleDb.OleDbType.VarChar, 50, "Основание") '10
        DA.InsertCommand.Parameters.Add("@Хран", OleDb.OleDbType.VarChar, 5, "Место хранения") '11

        tbtMonets = Module1.GetTable("")
        tbtStores = Module1.GetTableStores()

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        TableLayoutPanel1.SetRowSpan(GridControl1, 4)

        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(2).ColumnEdit = LookUp1Rep

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Обозначение"
        LookUp2rep.ValueMember = "Обозначение"
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(1).ColumnEdit = LookUp2rep

        GridView1.Columns(5).OptionsColumn.AllowEdit = False 'запрет редактирования цены, обработка по клику
        GridView1.Columns(6).OptionsColumn.AllowEdit = False 'запрет редактирования НДС
        GridView1.Columns(7).OptionsColumn.AllowEdit = False 'запрет редактирования Состояния
        GridView1.Columns(8).OptionsColumn.AllowEdit = False 'запрет редактирования Дефектов

        Combo1Rep.Items.AddRange({"отл.", "уд.", "неуд."})
        Combo1Rep.AutoComplete = True
        Combo1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(9).ColumnEdit = Combo1Rep

        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub TableCoinsPurchaseMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 6 Then
            Return
        End If

        tbt.Dispose()
        tbtMonets.Reset()
        tbtStores.Reset()
        LookUp1Rep.Dispose()
        LookUp2rep.Dispose()
        'Combo1Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        LookUp2rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit

        SqlCom = New OleDb.OleDbCommand("SELECT ДатаМонет AS Дата, [Наименование ТБ], [Каталожный номер], Монета, Количество, Цена  
FROM [Приобретение монет ТБ в ЦБ]", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM [Приобретение монет ТБ в ЦБ] 
WHERE ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND (([Наименование ТБ] = ?) OR [Наименование ТБ] IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.DeleteCommand.Parameters.Add("@НТБ", OleDb.OleDbType.VarChar, 40, "Наименование ТБ")
        DA.DeleteCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.DeleteCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.DeleteCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена")

        updCommand = New OleDb.OleDbCommand("UPDATE [Приобретение монет ТБ в ЦБ] 
SET ДатаМонет = ?, [Наименование ТБ] =?, [Каталожный номер] = ?, Монета = ?, Количество = ?, Цена = ?
WHERE ((ДатаМонет = ?) OR ДатаМонет IS NULL) AND (([Наименование ТБ] = ?) OR [Наименование ТБ] IS NULL) AND (([Каталожный номер] = ?) OR [Каталожный номер] IS NULL) AND ((Количество = ?) OR Количество IS NULL) AND ((Цена = ?) OR Цена IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '0
        DA.UpdateCommand.Parameters.Add("@НТБ", OleDb.OleDbType.VarChar, 40, "Наименование ТБ") '1
        DA.UpdateCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '2
        DA.UpdateCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '3
        DA.UpdateCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '4
        DA.UpdateCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '5
        DA.UpdateCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата")
        DA.UpdateCommand.Parameters.Item(6).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@НТБ", OleDb.OleDbType.VarChar, 40, "Наименование ТБ")
        DA.UpdateCommand.Parameters.Item(7).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(8).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество")
        DA.UpdateCommand.Parameters.Item(9).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена")
        DA.UpdateCommand.Parameters.Item(10).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO [Приобретение монет ТБ в ЦБ] 
        (ДатаМонет, [Наименование ТБ], [Каталожный номер], Монета, Количество, Цена) 
        VALUES (@Дата, @НТБ, @Номер, @Монета, @Колво, @Цена)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("@Дата", OleDb.OleDbType.Date, -1, "Дата") '0
        DA.InsertCommand.Parameters.Add("@НТБ", OleDb.OleDbType.VarChar, 40, "Наименование ТБ") '1
        DA.InsertCommand.Parameters.Add("@Номер", OleDb.OleDbType.VarChar, 9, "Каталожный номер") '2
        DA.InsertCommand.Parameters.Add("@Монета", OleDb.OleDbType.VarChar, 60, "Монета") '3
        DA.InsertCommand.Parameters.Add("@Колво", OleDb.OleDbType.SmallInt, -1, "Количество") '4
        DA.InsertCommand.Parameters.Add("@Цена", OleDb.OleDbType.Double, -1, "Цена") '5

        tbtMonets = Module1.GetTable("")
        tbtStores = Module1.GetTableRegional()

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        TableLayoutPanel1.SetRowSpan(GridControl1, 4)

        LookUp1Rep.DataSource = tbtMonets
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Каталожный номер"
        LookUp1Rep.ValueMember = "Каталожный номер"
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(2).ColumnEdit = LookUp1Rep

        LookUp2rep.DataSource = tbtStores
        LookUp2rep.AutoComplete = True
        LookUp2rep.DisplayMember = "Наименование"
        LookUp2rep.ValueMember = "Наименование"
        LookUp2rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        LookUp2rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(1).ColumnEdit = LookUp2rep

        GridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        GridView1.OptionsSelection.MultiSelect = True
    End Sub
    '--------------------------------------------------------

End Class