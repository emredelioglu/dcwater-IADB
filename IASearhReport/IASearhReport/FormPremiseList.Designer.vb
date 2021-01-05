<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPremiseList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.DataGridPremises = New System.Windows.Forms.DataGridView
        Me.Column1 = New System.Windows.Forms.DataGridViewImageColumn
        Me.Column2 = New System.Windows.Forms.DataGridViewImageColumn
        Me.PexprmDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexactDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.OwnerDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexsadDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexpstsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexptypDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexuidDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ClsPremiseBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.DataGridPremises, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ClsPremiseBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridPremises
        '
        Me.DataGridPremises.AllowUserToAddRows = False
        Me.DataGridPremises.AllowUserToDeleteRows = False
        Me.DataGridPremises.AutoGenerateColumns = False
        Me.DataGridPremises.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridPremises.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.PexprmDataGridViewTextBoxColumn, Me.PexactDataGridViewTextBoxColumn, Me.OwnerDataGridViewTextBoxColumn, Me.PexsadDataGridViewTextBoxColumn, Me.PexpstsDataGridViewTextBoxColumn, Me.PexptypDataGridViewTextBoxColumn, Me.PexuidDataGridViewTextBoxColumn})
        Me.DataGridPremises.DataSource = Me.ClsPremiseBindingSource
        Me.DataGridPremises.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridPremises.Location = New System.Drawing.Point(0, 0)
        Me.DataGridPremises.MultiSelect = False
        Me.DataGridPremises.Name = "DataGridPremises"
        Me.DataGridPremises.RowHeadersWidth = 10
        Me.DataGridPremises.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridPremises.Size = New System.Drawing.Size(668, 266)
        Me.DataGridPremises.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.HeaderText = ""
        Me.Column1.Image = Global.IAIS.My.Resources.Resources.zoomin
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 20
        '
        'Column2
        '
        Me.Column2.HeaderText = ""
        Me.Column2.Image = Global.IAIS.My.Resources.Resources.pan_1
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 20
        '
        'PexprmDataGridViewTextBoxColumn
        '
        Me.PexprmDataGridViewTextBoxColumn.DataPropertyName = "Pexprm"
        Me.PexprmDataGridViewTextBoxColumn.HeaderText = "Premise No."
        Me.PexprmDataGridViewTextBoxColumn.Name = "PexprmDataGridViewTextBoxColumn"
        '
        'PexactDataGridViewTextBoxColumn
        '
        Me.PexactDataGridViewTextBoxColumn.DataPropertyName = "Pexact"
        Me.PexactDataGridViewTextBoxColumn.HeaderText = "Account No."
        Me.PexactDataGridViewTextBoxColumn.Name = "PexactDataGridViewTextBoxColumn"
        '
        'OwnerDataGridViewTextBoxColumn
        '
        Me.OwnerDataGridViewTextBoxColumn.DataPropertyName = "Owner"
        Me.OwnerDataGridViewTextBoxColumn.HeaderText = "Owner"
        Me.OwnerDataGridViewTextBoxColumn.Name = "OwnerDataGridViewTextBoxColumn"
        '
        'PexsadDataGridViewTextBoxColumn
        '
        Me.PexsadDataGridViewTextBoxColumn.DataPropertyName = "Pexsad"
        Me.PexsadDataGridViewTextBoxColumn.HeaderText = "Address"
        Me.PexsadDataGridViewTextBoxColumn.Name = "PexsadDataGridViewTextBoxColumn"
        '
        'PexpstsDataGridViewTextBoxColumn
        '
        Me.PexpstsDataGridViewTextBoxColumn.DataPropertyName = "Pexpsts"
        Me.PexpstsDataGridViewTextBoxColumn.HeaderText = "Status"
        Me.PexpstsDataGridViewTextBoxColumn.Name = "PexpstsDataGridViewTextBoxColumn"
        '
        'PexptypDataGridViewTextBoxColumn
        '
        Me.PexptypDataGridViewTextBoxColumn.DataPropertyName = "Pexptyp"
        Me.PexptypDataGridViewTextBoxColumn.HeaderText = "Premise Type"
        Me.PexptypDataGridViewTextBoxColumn.Name = "PexptypDataGridViewTextBoxColumn"
        '
        'PexuidDataGridViewTextBoxColumn
        '
        Me.PexuidDataGridViewTextBoxColumn.DataPropertyName = "Pexuid"
        Me.PexuidDataGridViewTextBoxColumn.HeaderText = "Pexuid"
        Me.PexuidDataGridViewTextBoxColumn.Name = "PexuidDataGridViewTextBoxColumn"
        Me.PexuidDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.PexuidDataGridViewTextBoxColumn.Visible = False
        '
        'ClsPremiseBindingSource
        '
        Me.ClsPremiseBindingSource.DataSource = GetType(IAIS.clsPremise)
        '
        'FormPremiseList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(668, 266)
        Me.Controls.Add(Me.DataGridPremises)
        Me.Name = "FormPremiseList"
        Me.Text = "Premises"
        CType(Me.DataGridPremises, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ClsPremiseBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridPremises As System.Windows.Forms.DataGridView
    Friend WithEvents ClsPremiseBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents PexprmDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexactDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OwnerDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexsadDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexpstsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexptypDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexuidDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
