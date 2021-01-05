<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormIACFList
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.DataGridIACF = New System.Windows.Forms.DataGridView
        Me.sqft = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.sf = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EffDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridIACF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 274)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'DataGridIACF
        '
        Me.DataGridIACF.AllowUserToAddRows = False
        Me.DataGridIACF.AllowUserToDeleteRows = False
        Me.DataGridIACF.AllowUserToOrderColumns = True
        Me.DataGridIACF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridIACF.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.sqft, Me.sf, Me.EffDate, Me.EndDate})
        Me.DataGridIACF.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridIACF.Location = New System.Drawing.Point(13, 13)
        Me.DataGridIACF.MultiSelect = False
        Me.DataGridIACF.Name = "DataGridIACF"
        Me.DataGridIACF.RowHeadersWidth = 20
        Me.DataGridIACF.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridIACF.Size = New System.Drawing.Size(407, 241)
        Me.DataGridIACF.TabIndex = 1
        '
        'sqft
        '
        Me.sqft.HeaderText = "Bill ERU"
        Me.sqft.Name = "sqft"
        Me.sqft.ReadOnly = True
        '
        'sf
        '
        Me.sf.HeaderText = "IA Sqft"
        Me.sf.Name = "sf"
        Me.sf.ReadOnly = True
        '
        'EffDate
        '
        Me.EffDate.HeaderText = "Effective Date"
        Me.EffDate.Name = "EffDate"
        Me.EffDate.ReadOnly = True
        '
        'EndDate
        '
        Me.EndDate.HeaderText = "End Date"
        Me.EndDate.Name = "EndDate"
        Me.EndDate.ReadOnly = True
        '
        'FormIACFList
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(435, 315)
        Me.Controls.Add(Me.DataGridIACF)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormIACFList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "IA Charge History"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DataGridIACF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents DataGridIACF As System.Windows.Forms.DataGridView
    Friend WithEvents sqft As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents sf As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EffDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EndDate As System.Windows.Forms.DataGridViewTextBoxColumn

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
