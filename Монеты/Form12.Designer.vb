<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form12
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form12))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.GridControl1 = New DevExpress.XtraGrid.GridControl()
        Me.ЦеныBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.МонетыDataSet = New Монеты.МонетыDataSet()
        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colГод = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colКаталожныйномер = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colМонета = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colЦена = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colКоличество = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colВхНДС = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colСостояние = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemComboBox1 = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.colДефекты = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colМестоХранения = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.ЦеныTableAdapter = New Монеты.МонетыDataSetTableAdapters.ЦеныTableAdapter()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ЦеныBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.GridControl1, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1040, 613)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'GridControl1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.GridControl1, 2)
        Me.GridControl1.DataSource = Me.ЦеныBindingSource
        Me.GridControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControl1.Location = New System.Drawing.Point(23, 23)
        Me.GridControl1.MainView = Me.GridView1
        Me.GridControl1.Name = "GridControl1"
        Me.GridControl1.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemComboBox1})
        Me.TableLayoutPanel1.SetRowSpan(Me.GridControl1, 2)
        Me.GridControl1.Size = New System.Drawing.Size(994, 566)
        Me.GridControl1.TabIndex = 0
        Me.GridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})
        '
        'ЦеныBindingSource
        '
        Me.ЦеныBindingSource.DataMember = "Цены"
        Me.ЦеныBindingSource.DataSource = Me.МонетыDataSet
        '
        'МонетыDataSet
        '
        Me.МонетыDataSet.DataSetName = "МонетыDataSet"
        Me.МонетыDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'GridView1
        '
        Me.GridView1.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colГод, Me.colКаталожныйномер, Me.colМонета, Me.colЦена, Me.colКоличество, Me.colВхНДС, Me.colСостояние, Me.colДефекты, Me.colМестоХранения})
        Me.GridView1.GridControl = Me.GridControl1
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsView.ShowGroupPanel = False
        '
        'colГод
        '
        Me.colГод.FieldName = "Год"
        Me.colГод.Name = "colГод"
        Me.colГод.Visible = True
        Me.colГод.VisibleIndex = 0
        '
        'colКаталожныйномер
        '
        Me.colКаталожныйномер.FieldName = "Каталожный номер"
        Me.colКаталожныйномер.Name = "colКаталожныйномер"
        Me.colКаталожныйномер.Visible = True
        Me.colКаталожныйномер.VisibleIndex = 1
        '
        'colМонета
        '
        Me.colМонета.FieldName = "Монета"
        Me.colМонета.Name = "colМонета"
        Me.colМонета.Visible = True
        Me.colМонета.VisibleIndex = 2
        '
        'colЦена
        '
        Me.colЦена.FieldName = "Цена"
        Me.colЦена.Name = "colЦена"
        Me.colЦена.Visible = True
        Me.colЦена.VisibleIndex = 3
        '
        'colКоличество
        '
        Me.colКоличество.FieldName = "Количество"
        Me.colКоличество.Name = "colКоличество"
        Me.colКоличество.Visible = True
        Me.colКоличество.VisibleIndex = 4
        '
        'colВхНДС
        '
        Me.colВхНДС.FieldName = "ВхНДС"
        Me.colВхНДС.Name = "colВхНДС"
        Me.colВхНДС.Visible = True
        Me.colВхНДС.VisibleIndex = 5
        '
        'colСостояние
        '
        Me.colСостояние.ColumnEdit = Me.RepositoryItemComboBox1
        Me.colСостояние.FieldName = "Состояние"
        Me.colСостояние.Name = "colСостояние"
        Me.colСостояние.Visible = True
        Me.colСостояние.VisibleIndex = 6
        '
        'RepositoryItemComboBox1
        '
        Me.RepositoryItemComboBox1.AutoHeight = False
        Me.RepositoryItemComboBox1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemComboBox1.Items.AddRange(New Object() {"отл.", "уд.", "неуд."})
        Me.RepositoryItemComboBox1.Name = "RepositoryItemComboBox1"
        '
        'colДефекты
        '
        Me.colДефекты.FieldName = "Дефекты"
        Me.colДефекты.Name = "colДефекты"
        Me.colДефекты.Visible = True
        Me.colДефекты.VisibleIndex = 7
        '
        'colМестоХранения
        '
        Me.colМестоХранения.FieldName = "МестоХранения"
        Me.colМестоХранения.Name = "colМестоХранения"
        Me.colМестоХранения.Visible = True
        Me.colМестоХранения.VisibleIndex = 8
        '
        'ЦеныTableAdapter
        '
        Me.ЦеныTableAdapter.ClearBeforeFill = True
        '
        'Form12
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1040, 613)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form12"
        Me.Text = "Цены"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ЦеныBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.МонетыDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents GridControl1 As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents МонетыDataSet As МонетыDataSet
    Friend WithEvents ЦеныBindingSource As BindingSource
    Friend WithEvents ЦеныTableAdapter As МонетыDataSetTableAdapters.ЦеныTableAdapter
    Friend WithEvents colГод As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colКаталожныйномер As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colМонета As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colЦена As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colКоличество As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colВхНДС As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colСостояние As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colДефекты As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colМестоХранения As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemComboBox1 As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
End Class
