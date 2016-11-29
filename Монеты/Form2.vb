Imports System.ComponentModel
Imports Монеты.MainSettings
Imports DevExpress.XtraGrid.Views.Base
Public Class G_coins
    Private FirstOpen As Boolean = True
    Private OperationBool As Boolean = False 'Открытие подменю операций
    Private TabNum As Int16 = -1 'Номер текущей таблицы и пункта меню
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
    Private tbtCurrency As New DataTable() ' переменная таблица для валюты

    Private LookUp1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
    Private Combo1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private Combo2Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private Combo3Rep As New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Private WithEvents Text1Rep As New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit

    Private Function Update_table() As Boolean
        Try
            GridView1.UpdateCurrentRow()
        Catch ex As Exception
        End Try
        Try
            Select Case TabNum
                Case 1, 2 ' Coins_guide
                    DA.Update(tbt)
            End Select
        Catch ex As Exception
            MsgBox("Неуспешное обновление")
            Return False
        End Try
        Return True
    End Function

    Private Sub G_coins_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FirstOpen = True
        Coins_guide_Click(sender, e)
    End Sub

    Private Sub G_coins_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not Update_table() Then
            e.Cancel = True
            Return
        End If
    End Sub

    Private Sub Coins_guide_Click(sender As Object, e As EventArgs) Handles Coins_guide.Click
        If Not FirstOpen Then
            If Not Update_table() Then
                Return
            End If
        Else
            FirstOpen = False
        End If

        TabNum = 1

        GridView1.Columns.Clear()
        GridView1.OptionsBehavior.ReadOnly = False

        TableCoinsGuideMake()
    End Sub

    Private Sub Addition_Click(sender As Object, e As EventArgs) Handles Addition.Click
        If Not FirstOpen Then
            If Not Update_table() Then
                Return
            End If
        Else
            FirstOpen = False
        End If

        TabNum = 2

        GridView1.Columns.Clear()
        GridView1.OptionsBehavior.ReadOnly = False

        TableLastAddMake()
    End Sub

    Private Sub Catalog_Click(sender As Object, e As EventArgs) Handles Catalog.Click
        If Not FirstOpen Then
            If Not Update_table() Then
                Return
            End If
        Else
            FirstOpen = False
        End If

        TabNum = 3

        GridView1.Columns.Clear()
        GridView1.OptionsBehavior.ReadOnly = True

        TableCatalogMake()
    End Sub

    Private Sub Tr_sets_Click(sender As Object, e As EventArgs) Handles Tr_sets.Click
        Me.Visible = False
        Form3.Show()
    End Sub

    Private Sub Text1Rep_ValueChanging(sender As Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles Text1Rep.EditValueChanging
        Select Case TabNum
            Case 1
                'If (e.NewValue = "") And (e.OldValue = "") Then
                '    MsgBox("No empty text")
                'End If
        End Select
    End Sub

    Private Sub GridView1_FocusedRowChanged(sender As Object, e As FocusedRowChangedEventArgs) Handles GridView1.FocusedRowChanged
        Select Case TabNum
            Case 1
                'Проверка на пустоту поля каталожного номера
                If (e.PrevFocusedRowHandle > 0) And (GridView1.GetRowCellDisplayText(e.PrevFocusedRowHandle, GridView1.Columns(1)) = "") Then
                    GridView1.FocusedRowHandle = e.PrevFocusedRowHandle
                    MsgBox("Поле ""Каталожный номер"" не может содержать пустое значение", MsgBoxStyle.Exclamation, "Внимание")
                End If
                'загрузка изображений монет в picturebox
                Dim rus_forein As String
                Dim Reverse As String
                Dim Obverse As String
                Try
                    rus_forein = IIf(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(8)) = "российск. рубль", "Российские\", "Иностранные\")
                    Try
                        Reverse = CStr(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(15)))
                    Catch ex As Exception
                        Reverse = Nothing
                    End Try
                    Try
                        Obverse = CStr(GridView1.GetRowCellValue(GridView1.FocusedRowHandle, GridView1.Columns(16)))
                    Catch ex As Exception
                        Obverse = Nothing
                    End Try
                    Try
                        If (Reverse = Nothing) Or (Reverse = "") Then
                            PictureBox1.Load(AppS.PicPath + "АверсФон.jpg")
                        Else
                            PictureBox1.Load(AppS.PicPath + rus_forein + "Реверсы\" + Reverse)
                        End If
                        If (Obverse = Nothing) Or (Obverse = "") Then
                            PictureBox2.Load(AppS.PicPath + "АверсФон.jpg")
                        Else
                            PictureBox2.Load(AppS.PicPath + rus_forein + "Аверсы\" + Obverse)
                        End If
                    Catch ex As Exception
                        PictureBox1.Load(AppS.PicPath + "АверсФон.jpg")
                        PictureBox2.Load(AppS.PicPath + "АверсФон.jpg")
                    End Try
                Catch ex As Exception
                    PictureBox1.Load(AppS.PicPath + "АверсФон.jpg")
                    PictureBox2.Load(AppS.PicPath + "АверсФон.jpg")
                End Try
            Case Else
                PictureBox1.Load(AppS.PicPath + "АверсФон.jpg")
                PictureBox2.Load(AppS.PicPath + "АверсФон.jpg")
        End Select
    End Sub

    Private Sub gridControl1_ProcessGridKey(sender As Object, e As KeyEventArgs) Handles GridControl1.ProcessGridKey
        Dim grid = CType(sender, DevExpress.XtraGrid.GridControl)
        Dim view = CType(grid.FocusedView, DevExpress.XtraGrid.Views.Grid.GridView)
        Select Case TabNum
            Case 1
                If (e.KeyData = Keys.Delete) Then
                    view.DeleteSelectedRows()
                    e.Handled = True
                End If
        End Select
    End Sub

    ''''''''''''''''''
    'Заполнение таблиц
    Private Sub TableCoinsGuideMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 1 Then
            Return
        End If

        tbt.Dispose()
        tbtCurrency.Dispose()
        Text1Rep.Dispose()
        Combo1Rep.Dispose()
        Combo2Rep.Dispose()
        LookUp1Rep.Dispose()
        Combo3Rep.Dispose()

        tbt = New DataTable
        tbtCurrency = New DataTable
        DA = New OleDb.OleDbDataAdapter
        Text1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
        Combo1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        Combo2Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox
        LookUp1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit
        Combo3Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox

        SqlCom = New OleDb.OleDbCommand("SELECT ВыводВСписок as Выводить, [Каталожный номер], Краткое as Наименование, Год, Качество, Металл, Проба, Номинал, Валюта, Масса, Диаметр, Тираж, Статус, 
        Серия, Надписи, ФотоРеверс, ФотоАверс 
FROM Монеты", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM Монеты 
WHERE ([Каталожный номер] = ?)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("1", OleDb.OleDbType.VarChar, 9, "Каталожный номер")

        updCommand = New OleDb.OleDbCommand("UPDATE Монеты 
SET ВыводВСписок = ?, [Каталожный номер] = ?, Краткое = ?, Год = ?, Качество = ?, Металл = ?, Проба = ?, Номинал = ?, Валюта = ?, Масса = ?, Диаметр = ?, Тираж = ?, Статус = ?, Серия = ?, Надписи = ?, ФотоРеверс = ?, ФотоАверс = ? 
WHERE ([Каталожный номер] = ?)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("1", OleDb.OleDbType.Boolean, -1, "Выводить")
        DA.UpdateCommand.Parameters.Add("2", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 60, "Наименование")
        DA.UpdateCommand.Parameters.Add("4", OleDb.OleDbType.SmallInt, -1, "Год")
        DA.UpdateCommand.Parameters.Add("5", OleDb.OleDbType.VarChar, 15, "Качество")
        DA.UpdateCommand.Parameters.Add("6", OleDb.OleDbType.VarChar, 8, "Металл")
        DA.UpdateCommand.Parameters.Add("7", OleDb.OleDbType.Single, -1, "Проба")
        DA.UpdateCommand.Parameters.Add("8", OleDb.OleDbType.Single, -1, "Номинал")
        DA.UpdateCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 20, "Валюта")
        DA.UpdateCommand.Parameters.Add("10", OleDb.OleDbType.Single, -1, "Масса")
        DA.UpdateCommand.Parameters.Add("11", OleDb.OleDbType.Single, -1, "Диаметр")
        DA.UpdateCommand.Parameters.Add("12", OleDb.OleDbType.Integer, -1, "Тираж")
        DA.UpdateCommand.Parameters.Add("13", OleDb.OleDbType.VarChar, 6, "Статус")
        DA.UpdateCommand.Parameters.Add("14", OleDb.OleDbType.VarChar, 80, "Серия")
        DA.UpdateCommand.Parameters.Add("15", OleDb.OleDbType.VarChar, 80, "Надписи")
        DA.UpdateCommand.Parameters.Add("16", OleDb.OleDbType.VarChar, 40, "ФотоРеверс")
        DA.UpdateCommand.Parameters.Add("17", OleDb.OleDbType.VarChar, 40, "ФотоАверс")
        DA.UpdateCommand.Parameters.Add("s1", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Item(17).SourceVersion = DataRowVersion.Original

        insCommand = New OleDb.OleDbCommand("INSERT INTO Заявки
        (ВыводВСписок, [Каталожный номер], Краткое, Год, Качество, Металл, Проба, Номинал, Валюта, Масса, Диаметр, Тираж, Статус, Серия, Надписи, ФотоРеверс, ФотоАверс)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", Con)
        DA.InsertCommand = insCommand
        DA.InsertCommand.Parameters.Add("1", OleDb.OleDbType.Boolean, -1, "Выводить")
        DA.InsertCommand.Parameters.Add("2", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.InsertCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 60, "Наименование")
        DA.InsertCommand.Parameters.Add("4", OleDb.OleDbType.SmallInt, -1, "Год")
        DA.InsertCommand.Parameters.Add("5", OleDb.OleDbType.VarChar, 15, "Качество")
        DA.InsertCommand.Parameters.Add("6", OleDb.OleDbType.VarChar, 8, "Металл")
        DA.InsertCommand.Parameters.Add("7", OleDb.OleDbType.Single, -1, "Проба")
        DA.InsertCommand.Parameters.Add("8", OleDb.OleDbType.Single, -1, "Номинал")
        DA.InsertCommand.Parameters.Add("9", OleDb.OleDbType.VarChar, 20, "Валюта")
        DA.InsertCommand.Parameters.Add("10", OleDb.OleDbType.Single, -1, "Масса")
        DA.InsertCommand.Parameters.Add("11", OleDb.OleDbType.Single, -1, "Диаметр")
        DA.InsertCommand.Parameters.Add("12", OleDb.OleDbType.Integer, -1, "Тираж")
        DA.InsertCommand.Parameters.Add("13", OleDb.OleDbType.VarChar, 6, "Статус")
        DA.InsertCommand.Parameters.Add("14", OleDb.OleDbType.VarChar, 80, "Серия")
        DA.InsertCommand.Parameters.Add("15", OleDb.OleDbType.VarChar, 80, "Надписи")
        DA.InsertCommand.Parameters.Add("16", OleDb.OleDbType.VarChar, 40, "ФотоРеверс")
        DA.InsertCommand.Parameters.Add("17", OleDb.OleDbType.VarChar, 40, "ФотоАверс")

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        tbtCurrency = Module1.GetTableCurrency()

        GridView1.Columns(1).ColumnEdit = Text1Rep

        Combo1Rep.Items.AddRange({"proof", "uncirculated", "br.uncirculated"})
        Combo1Rep.AutoComplete = True
        Combo1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(4).ColumnEdit = Combo1Rep

        Combo2Rep.Items.AddRange({"серебро", "золото", "платина", "палладий"})
        Combo2Rep.AutoComplete = True
        Combo2Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(5).ColumnEdit = Combo2Rep

        LookUp1Rep.DataSource = tbtCurrency
        LookUp1Rep.AutoComplete = True
        LookUp1Rep.DisplayMember = "Валюта"
        LookUp1Rep.ValueMember = "Валюта"
        LookUp1Rep.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True
        LookUp1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        LookUp1Rep.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup
        GridView1.Columns(8).ColumnEdit = LookUp1Rep

        Combo3Rep.Items.AddRange({"нах.", "выш.", "иност."})
        Combo3Rep.AutoComplete = True
        Combo3Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(12).ColumnEdit = Combo3Rep

        GridView1.OptionsSelection.MultiSelect = True
    End Sub

    Private Sub TableLastAddMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 2 Then
            Return
        End If

        tbt.Dispose()
        tbtCurrency.Dispose()
        Text1Rep.Dispose()
        Combo1Rep.Dispose()
        Combo2Rep.Dispose()
        LookUp1Rep.Dispose()
        Combo3Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter
        Combo1Rep = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox

        SqlCom = New OleDb.OleDbCommand("SELECT Наименование, Качество, [Облаг НДС] as [Облагается НДС], [Металл, проба], Масса, [Ном-л] as Номинал, [Каталожный номер]
FROM [Последние введённые монеты]", Con)
        DA.SelectCommand = SqlCom

        delCommand = New OleDb.OleDbCommand("DELETE FROM [Последние введённые монеты] 
WHERE ((Наименование = ?) OR Наименование IS NULL) AND (([Облаг НДС] = ?) OR [Облаг НДС] IS NULL) AND (([Металл, проба] = ?) OR [Металл, проба] IS NULL) AND ((Масса = ?) OR Масса IS NULL) AND (([Ном-л] = ?) OR [Ном-л] IS NULL)", Con)
        DA.DeleteCommand = delCommand
        DA.DeleteCommand.Parameters.Add("1", OleDb.OleDbType.VarChar, 255, "Наименование")
        DA.DeleteCommand.Parameters.Add("2", OleDb.OleDbType.VarChar, 255, "Облагается НДС")
        DA.DeleteCommand.Parameters.Add("3", OleDb.OleDbType.VarChar, 255, "Металл, проба")
        DA.DeleteCommand.Parameters.Add("4", OleDb.OleDbType.VarChar, 255, "Масса")
        DA.DeleteCommand.Parameters.Add("5", OleDb.OleDbType.VarChar, 255, "Номинал")

        updCommand = New OleDb.OleDbCommand("UPDATE [Последние введённые монеты] 
SET Качество = ?, [Каталожный номер] = ? 
WHERE ((Наименование = ?) OR Наименование IS NULL) AND (([Облаг НДС] = ?) OR [Облаг НДС] IS NULL) AND (([Металл, проба] = ?) OR [Металл, проба] IS NULL) AND ((Масса = ?) OR Масса IS NULL) AND (([Ном-л] = ?) OR [Ном-л] IS NULL)", Con)
        DA.UpdateCommand = updCommand
        DA.UpdateCommand.Parameters.Add("1", OleDb.OleDbType.VarChar, 15, "Качество")
        DA.UpdateCommand.Parameters.Add("2", OleDb.OleDbType.VarChar, 9, "Каталожный номер")
        DA.UpdateCommand.Parameters.Add("s1", OleDb.OleDbType.VarChar, 255, "Наименование")
        DA.UpdateCommand.Parameters.Item(2).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("s2", OleDb.OleDbType.VarChar, 255, "Облагается НДС")
        DA.UpdateCommand.Parameters.Item(3).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("s3", OleDb.OleDbType.VarChar, 255, "Металл, проба")
        DA.UpdateCommand.Parameters.Item(4).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("s4", OleDb.OleDbType.VarChar, 255, "Масса")
        DA.UpdateCommand.Parameters.Item(5).SourceVersion = DataRowVersion.Original
        DA.UpdateCommand.Parameters.Add("s5", OleDb.OleDbType.VarChar, 255, "Номинал")
        DA.UpdateCommand.Parameters.Item(6).SourceVersion = DataRowVersion.Original

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        GridView1.Columns(0).OptionsColumn.AllowEdit = False
        GridView1.Columns(2).OptionsColumn.AllowEdit = False
        GridView1.Columns(3).OptionsColumn.AllowEdit = False
        GridView1.Columns(4).OptionsColumn.AllowEdit = False
        GridView1.Columns(5).OptionsColumn.AllowEdit = False

        Combo1Rep.Items.AddRange({"proof", "uncirculated", "br.uncirculated"})
        Combo1Rep.AutoComplete = True
        Combo1Rep.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        GridView1.Columns(1).ColumnEdit = Combo1Rep

        GridView1.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False
        GridView1.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False
        GridView1.OptionsSelection.MultiSelect = True
        Try
            GridView1.FocusedRowHandle = GridView1.RowCount - 1
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TableCatalogMake()
        Cursor.Current = Cursors.WaitCursor
        If TabNum <> 3 Then
            Return
        End If

        tbt.Dispose()
        tbtCurrency.Dispose()
        Text1Rep.Dispose()
        Combo1Rep.Dispose()
        Combo2Rep.Dispose()
        LookUp1Rep.Dispose()
        Combo3Rep.Dispose()

        tbt = New DataTable
        DA = New OleDb.OleDbDataAdapter

        SqlCom = New OleDb.OleDbCommand("SELECT * 
FROM [Каталожные номера иностранных]", Con)
        DA.SelectCommand = SqlCom

        DA.Fill(tbt)
        bs1.DataSource = tbt
        GridControl1.DataSource = bs1

        GridView1.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False
        GridView1.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False
        GridView1.OptionsSelection.MultiSelect = True
        Try
            GridView1.FocusedRowHandle = GridView1.RowCount - 1
        Catch ex As Exception
        End Try
    End Sub
    ''''''''''''''''''
End Class