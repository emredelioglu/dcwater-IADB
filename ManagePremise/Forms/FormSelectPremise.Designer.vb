<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSelectPremise
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
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.DataGridPremises = New System.Windows.Forms.DataGridView
        Me.pexuid = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.MasterFlagDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.PexprmDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PexsadDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ImperviousOnlyDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ExemptIABDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ExemptReasonDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ClsPremisePtBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGridPremises, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ClsPremisePtBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 276)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(567, 48)
        Me.Panel1.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(182, 13)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(12, 13)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(136, 23)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "Assign Charge Feature"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.DataGridPremises)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(567, 276)
        Me.Panel2.TabIndex = 4
        '
        'DataGridPremises
        '
        Me.DataGridPremises.AllowUserToAddRows = False
        Me.DataGridPremises.AllowUserToDeleteRows = False
        Me.DataGridPremises.AutoGenerateColumns = False
        Me.DataGridPremises.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridPremises.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MasterFlagDataGridViewCheckBoxColumn, Me.PexprmDataGridViewTextBoxColumn, Me.pexuid, Me.PexsadDataGridViewTextBoxColumn, Me.ImperviousOnlyDataGridViewTextBoxColumn, Me.ExemptIABDataGridViewTextBoxColumn, Me.ExemptReasonDataGridViewTextBoxColumn})
        Me.DataGridPremises.DataSource = Me.ClsPremisePtBindingSource
        Me.DataGridPremises.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridPremises.Location = New System.Drawing.Point(0, 0)
        Me.DataGridPremises.MultiSelect = False
        Me.DataGridPremises.Name = "DataGridPremises"
        Me.DataGridPremises.ReadOnly = True
        Me.DataGridPremises.RowHeadersVisible = False
        Me.DataGridPremises.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridPremises.Size = New System.Drawing.Size(563, 272)
        Me.DataGridPremises.TabIndex = 2
        '
        'pexuid
        '
        Me.pexuid.DataPropertyName = "pexuid"
        Me.pexuid.HeaderText = "Premise ID"
        Me.pexuid.Name = "pexuid"
        Me.pexuid.ReadOnly = True
        '
        'MasterFlagDataGridViewCheckBoxColumn
        '
        Me.MasterFlagDataGridViewCheckBoxColumn.DataPropertyName = "masterFlag"
        Me.MasterFlagDataGridViewCheckBoxColumn.HeaderText = ""
        Me.MasterFlagDataGridViewCheckBoxColumn.Name = "MasterFlagDataGridViewCheckBoxColumn"
        Me.MasterFlagDataGridViewCheckBoxColumn.ReadOnly = True
        Me.MasterFlagDataGridViewCheckBoxColumn.Width = 20
        '
        'PexprmDataGridViewTextBoxColumn
        '
        Me.PexprmDataGridViewTextBoxColumn.DataPropertyName = "pexprm"
        Me.PexprmDataGridViewTextBoxColumn.HeaderText = "Premise Number"
        Me.PexprmDataGridViewTextBoxColumn.Name = "PexprmDataGridViewTextBoxColumn"
        Me.PexprmDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PexsadDataGridViewTextBoxColumn
        '
        Me.PexsadDataGridViewTextBoxColumn.DataPropertyName = "pexsad"
        Me.PexsadDataGridViewTextBoxColumn.HeaderText = "Premise Address"
        Me.PexsadDataGridViewTextBoxColumn.Name = "PexsadDataGridViewTextBoxColumn"
        Me.PexsadDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ImperviousOnlyDataGridViewTextBoxColumn
        '
        Me.ImperviousOnlyDataGridViewTextBoxColumn.DataPropertyName = "imperviousOnly"
        Me.ImperviousOnlyDataGridViewTextBoxColumn.HeaderText = "Impervious Only"
        Me.ImperviousOnlyDataGridViewTextBoxColumn.Name = "ImperviousOnlyDataGridViewTextBoxColumn"
        Me.ImperviousOnlyDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ExemptIABDataGridViewTextBoxColumn
        '
        Me.ExemptIABDataGridViewTextBoxColumn.DataPropertyName = "exemptIAB"
        Me.ExemptIABDataGridViewTextBoxColumn.HeaderText = "Is Exempt"
        Me.ExemptIABDataGridViewTextBoxColumn.Name = "ExemptIABDataGridViewTextBoxColumn"
        Me.ExemptIABDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ExemptReasonDataGridViewTextBoxColumn
        '
        Me.ExemptReasonDataGridViewTextBoxColumn.DataPropertyName = "exemptReason"
        Me.ExemptReasonDataGridViewTextBoxColumn.HeaderText = "Exempt Reason"
        Me.ExemptReasonDataGridViewTextBoxColumn.Name = "ExemptReasonDataGridViewTextBoxColumn"
        Me.ExemptReasonDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ClsPremisePtBindingSource
        '
        Me.ClsPremisePtBindingSource.DataSource = GetType(IAIS.clsPremisePt)
        '
        'FormSelectPremise
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(567, 324)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormSelectPremise"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Premises by Selected Parcel"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGridPremises, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ClsPremisePtBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ClsPremisePtBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents DataGridPremises As System.Windows.Forms.DataGridView
    Friend WithEvents MasterFlagDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents PexprmDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents pexuid As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PexsadDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImperviousOnlyDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExemptIABDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExemptReasonDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
