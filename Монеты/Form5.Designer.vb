<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form5))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ЮридическиеЛицаBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.МонетыDataSet = New Монеты.МонетыDataSet()
        Me.Юридические_лицаTableAdapter = New Монеты.МонетыDataSetTableAdapters.Юридические_лицаTableAdapter()
        Me.НаименованиеDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ЮридическиеЛицаBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(477, 512)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.НаименованиеDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.ЮридическиеЛицаBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(23, 28)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(431, 456)
        Me.DataGridView1.TabIndex = 0
        '
        'ЮридическиеЛицаBindingSource
        '
        Me.ЮридическиеЛицаBindingSource.DataMember = "Юридические лица"
        Me.ЮридическиеЛицаBindingSource.DataSource = Me.МонетыDataSet
        '
        'МонетыDataSet
        '
        Me.МонетыDataSet.DataSetName = "МонетыDataSet"
        Me.МонетыDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Юридические_лицаTableAdapter
        '
        Me.Юридические_лицаTableAdapter.ClearBeforeFill = True
        '
        'НаименованиеDataGridViewTextBoxColumn
        '
        Me.НаименованиеDataGridViewTextBoxColumn.DataPropertyName = "Наименование"
        Me.НаименованиеDataGridViewTextBoxColumn.HeaderText = "Наименование"
        Me.НаименованиеDataGridViewTextBoxColumn.Name = "НаименованиеDataGridViewTextBoxColumn"
        Me.НаименованиеDataGridViewTextBoxColumn.Width = 648
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(477, 512)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form5"
        Me.Text = "Юридические лица"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ЮридическиеЛицаBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents МонетыDataSet As МонетыDataSet
    Friend WithEvents ЮридическиеЛицаBindingSource As BindingSource
    Friend WithEvents Юридические_лицаTableAdapter As МонетыDataSetTableAdapters.Юридические_лицаTableAdapter
    Friend WithEvents НаименованиеDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
