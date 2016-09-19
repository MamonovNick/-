<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form10
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form10))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.МонетыDataSet = New Монеты.МонетыDataSet()
        Me.ИндивидуальныеСтавкиУсловныхBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Индивидуальные_ставки_условныхTableAdapter = New Монеты.МонетыDataSetTableAdapters.Индивидуальные_ставки_условныхTableAdapter()
        Me.ДатаDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.КаталожныйНомерDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.МонетаDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.СостояниеDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.СтавкаDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ИндивидуальныеСтавкиУсловныхBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 15.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 15.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 15.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 15.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(965, 569)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ДатаDataGridViewTextBoxColumn, Me.КаталожныйНомерDataGridViewTextBoxColumn, Me.МонетаDataGridViewTextBoxColumn, Me.СостояниеDataGridViewTextBoxColumn, Me.СтавкаDataGridViewTextBoxColumn})
        Me.TableLayoutPanel1.SetColumnSpan(Me.DataGridView1, 2)
        Me.DataGridView1.DataSource = Me.ИндивидуальныеСтавкиУсловныхBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(18, 18)
        Me.DataGridView1.Name = "DataGridView1"
        Me.TableLayoutPanel1.SetRowSpan(Me.DataGridView1, 2)
        Me.DataGridView1.Size = New System.Drawing.Size(928, 532)
        Me.DataGridView1.TabIndex = 0
        '
        'МонетыDataSet
        '
        Me.МонетыDataSet.DataSetName = "МонетыDataSet"
        Me.МонетыDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ИндивидуальныеСтавкиУсловныхBindingSource
        '
        Me.ИндивидуальныеСтавкиУсловныхBindingSource.DataMember = "Индивидуальные ставки условных"
        Me.ИндивидуальныеСтавкиУсловныхBindingSource.DataSource = Me.МонетыDataSet
        '
        'Индивидуальные_ставки_условныхTableAdapter
        '
        Me.Индивидуальные_ставки_условныхTableAdapter.ClearBeforeFill = True
        '
        'ДатаDataGridViewTextBoxColumn
        '
        Me.ДатаDataGridViewTextBoxColumn.DataPropertyName = "Дата"
        Me.ДатаDataGridViewTextBoxColumn.HeaderText = "Дата"
        Me.ДатаDataGridViewTextBoxColumn.Name = "ДатаDataGridViewTextBoxColumn"
        '
        'КаталожныйНомерDataGridViewTextBoxColumn
        '
        Me.КаталожныйНомерDataGridViewTextBoxColumn.DataPropertyName = "Каталожный номер"
        Me.КаталожныйНомерDataGridViewTextBoxColumn.HeaderText = "Каталожный номер"
        Me.КаталожныйНомерDataGridViewTextBoxColumn.Name = "КаталожныйНомерDataGridViewTextBoxColumn"
        '
        'МонетаDataGridViewTextBoxColumn
        '
        Me.МонетаDataGridViewTextBoxColumn.DataPropertyName = "Монета"
        Me.МонетаDataGridViewTextBoxColumn.HeaderText = "Монета"
        Me.МонетаDataGridViewTextBoxColumn.Name = "МонетаDataGridViewTextBoxColumn"
        '
        'СостояниеDataGridViewTextBoxColumn
        '
        Me.СостояниеDataGridViewTextBoxColumn.DataPropertyName = "Состояние"
        Me.СостояниеDataGridViewTextBoxColumn.HeaderText = "Состояние"
        Me.СостояниеDataGridViewTextBoxColumn.Name = "СостояниеDataGridViewTextBoxColumn"
        '
        'СтавкаDataGridViewTextBoxColumn
        '
        Me.СтавкаDataGridViewTextBoxColumn.DataPropertyName = "Ставка"
        Me.СтавкаDataGridViewTextBoxColumn.HeaderText = "Ставка"
        Me.СтавкаDataGridViewTextBoxColumn.Name = "СтавкаDataGridViewTextBoxColumn"
        '
        'Form10
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(965, 569)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form10"
        Me.Text = "Индивидуальные ставки условных"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ИндивидуальныеСтавкиУсловныхBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents МонетыDataSet As МонетыDataSet
    Friend WithEvents ИндивидуальныеСтавкиУсловныхBindingSource As BindingSource
    Friend WithEvents Индивидуальные_ставки_условныхTableAdapter As МонетыDataSetTableAdapters.Индивидуальные_ставки_условныхTableAdapter
    Friend WithEvents ДатаDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents КаталожныйНомерDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents МонетаDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents СостояниеDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents СтавкаDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
