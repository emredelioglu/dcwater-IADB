<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPremiseLocator
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
        Me.LabelPexprm = New System.Windows.Forms.Label
        Me.TextPremiseNo = New System.Windows.Forms.TextBox
        Me.LabelAccountNo = New System.Windows.Forms.Label
        Me.TextAccountNo = New System.Windows.Forms.TextBox
        Me.LabelOwner = New System.Windows.Forms.Label
        Me.TextOwner = New System.Windows.Forms.TextBox
        Me.btnZoom = New System.Windows.Forms.Button
        Me.btnPan = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnClear = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'LabelPexprm
        '
        Me.LabelPexprm.AutoSize = True
        Me.LabelPexprm.Location = New System.Drawing.Point(12, 28)
        Me.LabelPexprm.Name = "LabelPexprm"
        Me.LabelPexprm.Size = New System.Drawing.Size(64, 13)
        Me.LabelPexprm.TabIndex = 1
        Me.LabelPexprm.Text = "Premise No."
        '
        'TextPremiseNo
        '
        Me.TextPremiseNo.Location = New System.Drawing.Point(108, 25)
        Me.TextPremiseNo.Name = "TextPremiseNo"
        Me.TextPremiseNo.Size = New System.Drawing.Size(118, 20)
        Me.TextPremiseNo.TabIndex = 2
        '
        'LabelAccountNo
        '
        Me.LabelAccountNo.AutoSize = True
        Me.LabelAccountNo.Location = New System.Drawing.Point(12, 61)
        Me.LabelAccountNo.Name = "LabelAccountNo"
        Me.LabelAccountNo.Size = New System.Drawing.Size(67, 13)
        Me.LabelAccountNo.TabIndex = 3
        Me.LabelAccountNo.Text = "Account No."
        '
        'TextAccountNo
        '
        Me.TextAccountNo.Location = New System.Drawing.Point(108, 58)
        Me.TextAccountNo.Name = "TextAccountNo"
        Me.TextAccountNo.Size = New System.Drawing.Size(118, 20)
        Me.TextAccountNo.TabIndex = 4
        '
        'LabelOwner
        '
        Me.LabelOwner.AutoSize = True
        Me.LabelOwner.Location = New System.Drawing.Point(12, 93)
        Me.LabelOwner.Name = "LabelOwner"
        Me.LabelOwner.Size = New System.Drawing.Size(38, 13)
        Me.LabelOwner.TabIndex = 5
        Me.LabelOwner.Text = "Owner"
        '
        'TextOwner
        '
        Me.TextOwner.Location = New System.Drawing.Point(108, 90)
        Me.TextOwner.Name = "TextOwner"
        Me.TextOwner.Size = New System.Drawing.Size(118, 20)
        Me.TextOwner.TabIndex = 6
        '
        'btnZoom
        '
        Me.btnZoom.Enabled = False
        Me.btnZoom.Location = New System.Drawing.Point(12, 137)
        Me.btnZoom.Name = "btnZoom"
        Me.btnZoom.Size = New System.Drawing.Size(67, 23)
        Me.btnZoom.TabIndex = 7
        Me.btnZoom.Text = "Zoom"
        Me.btnZoom.UseVisualStyleBackColor = True
        '
        'btnPan
        '
        Me.btnPan.Enabled = False
        Me.btnPan.Location = New System.Drawing.Point(85, 137)
        Me.btnPan.Name = "btnPan"
        Me.btnPan.Size = New System.Drawing.Size(67, 23)
        Me.btnPan.TabIndex = 8
        Me.btnPan.Text = "Pan"
        Me.btnPan.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(141, 166)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(67, 23)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(54, 166)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(67, 23)
        Me.btnClear.TabIndex = 10
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(159, 137)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(67, 23)
        Me.btnFind.TabIndex = 9
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'FormPremiseLocator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(255, 203)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnPan)
        Me.Controls.Add(Me.btnZoom)
        Me.Controls.Add(Me.TextOwner)
        Me.Controls.Add(Me.LabelOwner)
        Me.Controls.Add(Me.TextAccountNo)
        Me.Controls.Add(Me.LabelAccountNo)
        Me.Controls.Add(Me.TextPremiseNo)
        Me.Controls.Add(Me.LabelPexprm)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormPremiseLocator"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Premise Search"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelPexprm As System.Windows.Forms.Label
    Friend WithEvents TextPremiseNo As System.Windows.Forms.TextBox
    Friend WithEvents LabelAccountNo As System.Windows.Forms.Label
    Friend WithEvents TextAccountNo As System.Windows.Forms.TextBox
    Friend WithEvents LabelOwner As System.Windows.Forms.Label
    Friend WithEvents TextOwner As System.Windows.Forms.TextBox
    Friend WithEvents btnZoom As System.Windows.Forms.Button
    Friend WithEvents btnPan As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button

End Class
