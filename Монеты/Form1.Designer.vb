<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ОсновнаяToolStripMenuItem, Me.СправочникиToolStripMenuItem, Me.СервисToolStripMenuItem, Me.СправкаToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1131, 28)
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
        Me.ОПрограммеToolStripMenuItem.Size = New System.Drawing.Size(181, 26)
        Me.ОПрограммеToolStripMenuItem.Text = "О программе"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1131, 555)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainForm"
        Me.Text = "Монеты"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
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
End Class
