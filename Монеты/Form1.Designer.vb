﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ОсновнаяToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПодключениеКБДToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ВыходToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправочникиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.МонетыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПодразделенияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КонтрагентыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВидыВалютToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СервисToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НастройкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СправкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОПрограммеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton6 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton10 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton7 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton8 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton9 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton11 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton12 = New System.Windows.Forms.ToolStripButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBoxEdit1 = New DevExpress.XtraEditors.LookUpEdit()
        Me.ПодразделенияBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.МонетыDataSet = New Монеты.МонетыDataSet()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.GridControl1 = New DevExpress.XtraGrid.GridControl()
        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.SecDA = New Монеты.МонетыDataSetTableAdapters.SecDA()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.MenuStrip1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.ComboBoxEdit1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ПодразделенияBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОсновнаяToolStripMenuItem, Me.СправочникиToolStripMenuItem, Me.СервисToolStripMenuItem, Me.СправкаToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1204, 28)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ОсновнаяToolStripMenuItem
        '
        Me.ОсновнаяToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПодключениеКБДToolStripMenuItem, Me.ToolStripMenuItem2, Me.ВыходToolStripMenuItem})
        Me.ОсновнаяToolStripMenuItem.Name = "ОсновнаяToolStripMenuItem"
        Me.ОсновнаяToolStripMenuItem.Size = New System.Drawing.Size(90, 24)
        Me.ОсновнаяToolStripMenuItem.Text = "Основная"
        '
        'ПодключениеКБДToolStripMenuItem
        '
        Me.ПодключениеКБДToolStripMenuItem.Name = "ПодключениеКБДToolStripMenuItem"
        Me.ПодключениеКБДToolStripMenuItem.Size = New System.Drawing.Size(215, 26)
        Me.ПодключениеКБДToolStripMenuItem.Text = "Подключение к БД"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(212, 6)
        '
        'ВыходToolStripMenuItem
        '
        Me.ВыходToolStripMenuItem.Name = "ВыходToolStripMenuItem"
        Me.ВыходToolStripMenuItem.Size = New System.Drawing.Size(215, 26)
        Me.ВыходToolStripMenuItem.Text = "Выход"
        '
        'СправочникиToolStripMenuItem
        '
        Me.СправочникиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.МонетыToolStripMenuItem, Me.ПодразделенияToolStripMenuItem, Me.КонтрагентыToolStripMenuItem, Me.ВидыВалютToolStripMenuItem})
        Me.СправочникиToolStripMenuItem.Name = "СправочникиToolStripMenuItem"
        Me.СправочникиToolStripMenuItem.Size = New System.Drawing.Size(115, 24)
        Me.СправочникиToolStripMenuItem.Text = "Справочники"
        '
        'МонетыToolStripMenuItem
        '
        Me.МонетыToolStripMenuItem.Name = "МонетыToolStripMenuItem"
        Me.МонетыToolStripMenuItem.Size = New System.Drawing.Size(194, 26)
        Me.МонетыToolStripMenuItem.Text = "Монеты"
        '
        'ПодразделенияToolStripMenuItem
        '
        Me.ПодразделенияToolStripMenuItem.Name = "ПодразделенияToolStripMenuItem"
        Me.ПодразделенияToolStripMenuItem.Size = New System.Drawing.Size(194, 26)
        Me.ПодразделенияToolStripMenuItem.Text = "Подразделения"
        '
        'КонтрагентыToolStripMenuItem
        '
        Me.КонтрагентыToolStripMenuItem.Name = "КонтрагентыToolStripMenuItem"
        Me.КонтрагентыToolStripMenuItem.Size = New System.Drawing.Size(194, 26)
        Me.КонтрагентыToolStripMenuItem.Text = "Контрагенты"
        '
        'ВидыВалютToolStripMenuItem
        '
        Me.ВидыВалютToolStripMenuItem.Name = "ВидыВалютToolStripMenuItem"
        Me.ВидыВалютToolStripMenuItem.Size = New System.Drawing.Size(194, 26)
        Me.ВидыВалютToolStripMenuItem.Text = "Виды валют"
        '
        'СервисToolStripMenuItem
        '
        Me.СервисToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.НастройкиToolStripMenuItem})
        Me.СервисToolStripMenuItem.Name = "СервисToolStripMenuItem"
        Me.СервисToolStripMenuItem.Size = New System.Drawing.Size(71, 24)
        Me.СервисToolStripMenuItem.Text = "Сервис"
        '
        'НастройкиToolStripMenuItem
        '
        Me.НастройкиToolStripMenuItem.Name = "НастройкиToolStripMenuItem"
        Me.НастройкиToolStripMenuItem.Size = New System.Drawing.Size(159, 26)
        Me.НастройкиToolStripMenuItem.Text = "Настройки"
        '
        'СправкаToolStripMenuItem
        '
        Me.СправкаToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОПрограммеToolStripMenuItem})
        Me.СправкаToolStripMenuItem.Name = "СправкаToolStripMenuItem"
        Me.СправкаToolStripMenuItem.Size = New System.Drawing.Size(79, 24)
        Me.СправкаToolStripMenuItem.Text = "Справка"
        '
        'ОПрограммеToolStripMenuItem
        '
        Me.ОПрограммеToolStripMenuItem.Name = "ОПрограммеToolStripMenuItem"
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(179, 26)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 218.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 350.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 6.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.ToolStrip1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.ToolStrip2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 3, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.GridControl1, 2, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 28)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 46.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1204, 593)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripButton3, Me.ToolStripButton4, Me.ToolStripButton5, Me.ToolStripButton6, Me.ToolStripButton10, Me.ToolStripSeparator1})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.VerticalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 28)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.TableLayoutPanel1.SetRowSpan(Me.ToolStrip1, 4)
        Me.ToolStrip1.Size = New System.Drawing.Size(218, 560)
        Me.ToolStrip1.TabIndex = 3
        Me.ToolStrip1.Text = "Меню главного окна"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ToolStripButton1.Size = New System.Drawing.Size(206, 34)
        Me.ToolStripButton1.Text = "Заявки"
        Me.ToolStripButton1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton2.Size = New System.Drawing.Size(206, 34)
        Me.ToolStripButton2.Text = "Операции"
        Me.ToolStripButton2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Margin = New System.Windows.Forms.Padding(25, 1, 0, 2)
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(191, 24)
        Me.ToolStripButton3.Text = "Внутрисистемные"
        Me.ToolStripButton3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton3.Visible = False
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Margin = New System.Windows.Forms.Padding(25, 1, 0, 2)
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(191, 24)
        Me.ToolStripButton4.Text = "Внешние"
        Me.ToolStripButton4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton4.Visible = False
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton5.Image = CType(resources.GetObject("ToolStripButton5.Image"), System.Drawing.Image)
        Me.ToolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton5.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton5.Size = New System.Drawing.Size(206, 34)
        Me.ToolStripButton5.Tag = ""
        Me.ToolStripButton5.Text = "Перемещение монет"
        Me.ToolStripButton5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton5.ToolTipText = "Перемещение монет между хранилищами"
        '
        'ToolStripButton6
        '
        Me.ToolStripButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton6.Image = CType(resources.GetObject("ToolStripButton6.Image"), System.Drawing.Image)
        Me.ToolStripButton6.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton6.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.ToolStripButton6.Name = "ToolStripButton6"
        Me.ToolStripButton6.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton6.Size = New System.Drawing.Size(206, 34)
        Me.ToolStripButton6.Text = "Состояние монет"
        Me.ToolStripButton6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton6.ToolTipText = "Изменение состояния монет"
        '
        'ToolStripButton10
        '
        Me.ToolStripButton10.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton10.Image = CType(resources.GetObject("ToolStripButton10.Image"), System.Drawing.Image)
        Me.ToolStripButton10.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton10.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.ToolStripButton10.Name = "ToolStripButton10"
        Me.ToolStripButton10.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton10.Size = New System.Drawing.Size(206, 34)
        Me.ToolStripButton10.Text = "Покупка в ЦБ терр. банками"
        Me.ToolStripButton10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton10.ToolTipText = "Самостоятельная покупка в ЦБ терр. банками"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(216, 6)
        '
        'ToolStrip2
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.ToolStrip2, 4)
        Me.ToolStrip2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton7, Me.ToolStripButton8, Me.ToolStripButton9, Me.ToolStripButton11, Me.ToolStripButton12})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(1198, 28)
        Me.ToolStrip2.TabIndex = 5
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'ToolStripButton7
        '
        Me.ToolStripButton7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton7.Image = CType(resources.GetObject("ToolStripButton7.Image"), System.Drawing.Image)
        Me.ToolStripButton7.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton7.Name = "ToolStripButton7"
        Me.ToolStripButton7.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton7.Size = New System.Drawing.Size(131, 25)
        Me.ToolStripButton7.Text = "Остатки и цены"
        '
        'ToolStripButton8
        '
        Me.ToolStripButton8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton8.Image = CType(resources.GetObject("ToolStripButton8.Image"), System.Drawing.Image)
        Me.ToolStripButton8.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton8.Name = "ToolStripButton8"
        Me.ToolStripButton8.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton8.Size = New System.Drawing.Size(101, 25)
        Me.ToolStripButton8.Text = "Документы"
        '
        'ToolStripButton9
        '
        Me.ToolStripButton9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton9.Image = CType(resources.GetObject("ToolStripButton9.Image"), System.Drawing.Image)
        Me.ToolStripButton9.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton9.Name = "ToolStripButton9"
        Me.ToolStripButton9.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton9.Size = New System.Drawing.Size(73, 25)
        Me.ToolStripButton9.Text = "Отчеты"
        '
        'ToolStripButton11
        '
        Me.ToolStripButton11.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton11.Image = CType(resources.GetObject("ToolStripButton11.Image"), System.Drawing.Image)
        Me.ToolStripButton11.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton11.Name = "ToolStripButton11"
        Me.ToolStripButton11.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton11.Size = New System.Drawing.Size(148, 25)
        Me.ToolStripButton11.Text = "Условные доходы"
        '
        'ToolStripButton12
        '
        Me.ToolStripButton12.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton12.Image = CType(resources.GetObject("ToolStripButton12.Image"), System.Drawing.Image)
        Me.ToolStripButton12.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton12.Name = "ToolStripButton12"
        Me.ToolStripButton12.Padding = New System.Windows.Forms.Padding(5)
        Me.ToolStripButton12.Size = New System.Drawing.Size(184, 25)
        Me.ToolStripButton12.Text = "Суммы распределения"
        Me.ToolStripButton12.ToolTipText = "Расчет сумм распределения"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ComboBoxEdit1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(226, 500)
        Me.Panel1.Name = "Panel1"
        Me.TableLayoutPanel1.SetRowSpan(Me.Panel1, 2)
        Me.Panel1.Size = New System.Drawing.Size(344, 85)
        Me.Panel1.TabIndex = 7
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(3, 59)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(285, 21)
        Me.CheckBox1.TabIndex = 5
        Me.CheckBox1.Text = "Показать только незакрытые позиции"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(103, 29)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(76, 21)
        Me.RadioButton2.TabIndex = 3
        Me.RadioButton2.Text = "Только"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(6, 29)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(53, 21)
        Me.RadioButton1.TabIndex = 2
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Все"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(181, 3)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(107, 22)
        Me.DateTimePicker1.TabIndex = 1
        Me.DateTimePicker1.Value = New Date(2015, 10, 1, 23, 29, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Показать заявки с"
        '
        'ComboBoxEdit1
        '
        Me.ComboBoxEdit1.Enabled = False
        Me.ComboBoxEdit1.Location = New System.Drawing.Point(181, 29)
        Me.ComboBoxEdit1.Name = "ComboBoxEdit1"
        Me.ComboBoxEdit1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.ComboBoxEdit1.Properties.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Наименование", "Наименование", 96, DevExpress.Utils.FormatType.None, "", True, DevExpress.Utils.HorzAlignment.Near)})
        Me.ComboBoxEdit1.Properties.DataSource = Me.ПодразделенияBindingSource
        Me.ComboBoxEdit1.Properties.DisplayMember = "Наименование"
        Me.ComboBoxEdit1.Properties.NullText = ""
        Me.ComboBoxEdit1.Properties.PopupSizeable = False
        Me.ComboBoxEdit1.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard
        Me.ComboBoxEdit1.Properties.ValueMember = "Наименование"
        Me.ComboBoxEdit1.Size = New System.Drawing.Size(145, 22)
        Me.ComboBoxEdit1.TabIndex = 8
        '
        'ПодразделенияBindingSource
        '
        Me.ПодразделенияBindingSource.DataMember = "Подразделения"
        Me.ПодразделенияBindingSource.DataSource = Me.МонетыDataSet
        '
        'МонетыDataSet
        '
        Me.МонетыDataSet.DataSetName = "МонетыDataSet"
        Me.МонетыDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button2)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button3)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button4)
        Me.FlowLayoutPanel1.Controls.Add(Me.TextBox1)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(576, 546)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(619, 39)
        Me.FlowLayoutPanel1.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.Location = New System.Drawing.Point(490, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(126, 27)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.Location = New System.Drawing.Point(358, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(126, 27)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.Location = New System.Drawing.Point(226, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(126, 27)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.AutoSize = True
        Me.Button4.Location = New System.Drawing.Point(94, 3)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(126, 27)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TextBox1.Location = New System.Drawing.Point(18, 5)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(70, 22)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Button5)
        Me.Panel2.Controls.Add(Me.CheckBox3)
        Me.Panel2.Controls.Add(Me.CheckBox2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(576, 500)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(619, 40)
        Me.Panel2.TabIndex = 8
        '
        'Button5
        '
        Me.Button5.AutoSize = True
        Me.Button5.Location = New System.Drawing.Point(402, 10)
        Me.Button5.Name = "Button5"
        Me.Button5.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Button5.Size = New System.Drawing.Size(214, 27)
        Me.Button5.TabIndex = 2
        Me.Button5.Text = "Обновить исполнение заявок"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Location = New System.Drawing.Point(18, 19)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(230, 21)
        Me.CheckBox3.TabIndex = 1
        Me.CheckBox3.Text = "Не учитывать наличие заявки"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(18, 0)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(312, 21)
        Me.CheckBox2.TabIndex = 0
        Me.CheckBox2.Text = "Показывать по монете полный список цен"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'GridControl1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.GridControl1, 2)
        Me.GridControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControl1.Location = New System.Drawing.Point(226, 31)
        Me.GridControl1.MainView = Me.GridView1
        Me.GridControl1.Name = "GridControl1"
        Me.TableLayoutPanel1.SetRowSpan(Me.GridControl1, 2)
        Me.GridControl1.Size = New System.Drawing.Size(969, 463)
        Me.GridControl1.TabIndex = 9
        Me.GridControl1.UseEmbeddedNavigator = True
        Me.GridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})
        '
        'GridView1
        '
        Me.GridView1.GridControl = Me.GridControl1
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsView.ShowGroupPanel = False
        '
        'SecDA
        '
        Me.SecDA.ClearBeforeFill = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "Файлы MS Excel| *.xls; *.xlsx"
        Me.OpenFileDialog1.Title = "Выберите файл распрделения"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1204, 621)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MinimumSize = New System.Drawing.Size(1222, 668)
        Me.Name = "MainForm"
        Me.Text = "Монеты"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.ComboBoxEdit1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ПодразделенияBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ОсновнаяToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправочникиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents МонетыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПодразделенияToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КонтрагентыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВидыВалютToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СервисToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СправкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ОПрограммеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПодключениеКБДToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ВыходToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НастройкиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripSeparator
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ToolStripButton4 As ToolStripButton
    Friend WithEvents ToolStripButton5 As ToolStripButton
    Friend WithEvents ToolStripButton6 As ToolStripButton
    Friend WithEvents ToolStrip2 As ToolStrip
    Friend WithEvents ToolStripButton7 As ToolStripButton
    Friend WithEvents ToolStripButton8 As ToolStripButton
    Friend WithEvents ToolStripButton9 As ToolStripButton
    Friend WithEvents ToolStripButton10 As ToolStripButton
    Friend WithEvents ToolStripButton11 As ToolStripButton
    Friend WithEvents ToolStripButton12 As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents МонетыDataSet As МонетыDataSet
    Friend WithEvents Panel2 As Panel
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents GridControl1 As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents ComboBoxEdit1 As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents ПодразделенияBindingSource As BindingSource
    Friend WithEvents SecDA As МонетыDataSetTableAdapters.SecDA
    Friend WithEvents Button5 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
