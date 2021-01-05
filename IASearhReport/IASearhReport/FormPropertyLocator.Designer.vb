<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPropertyLocator
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
        Me.TextSquare = New System.Windows.Forms.TextBox
        Me.TextSuffix = New System.Windows.Forms.TextBox
        Me.TextLot = New System.Windows.Forms.TextBox
        Me.LabelSquare = New System.Windows.Forms.Label
        Me.LabelSuffix = New System.Windows.Forms.Label
        Me.LabelLot = New System.Windows.Forms.Label
        Me.btnZoom = New System.Windows.Forms.Button
        Me.btnPan = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnClear = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextSquare
        '
        Me.TextSquare.Location = New System.Drawing.Point(77, 24)
        Me.TextSquare.MaxLength = 4
        Me.TextSquare.Name = "TextSquare"
        Me.TextSquare.Size = New System.Drawing.Size(105, 20)
        Me.TextSquare.TabIndex = 1
        '
        'TextSuffix
        '
        Me.TextSuffix.Location = New System.Drawing.Point(77, 50)
        Me.TextSuffix.MaxLength = 4
        Me.TextSuffix.Name = "TextSuffix"
        Me.TextSuffix.Size = New System.Drawing.Size(105, 20)
        Me.TextSuffix.TabIndex = 2
        '
        'TextLot
        '
        Me.TextLot.Location = New System.Drawing.Point(77, 76)
        Me.TextLot.MaxLength = 15
        Me.TextLot.Name = "TextLot"
        Me.TextLot.Size = New System.Drawing.Size(105, 20)
        Me.TextLot.TabIndex = 3
        '
        'LabelSquare
        '
        Me.LabelSquare.AutoSize = True
        Me.LabelSquare.Location = New System.Drawing.Point(21, 24)
        Me.LabelSquare.Name = "LabelSquare"
        Me.LabelSquare.Size = New System.Drawing.Size(41, 13)
        Me.LabelSquare.TabIndex = 4
        Me.LabelSquare.Text = "Square"
        '
        'LabelSuffix
        '
        Me.LabelSuffix.AutoSize = True
        Me.LabelSuffix.Location = New System.Drawing.Point(21, 50)
        Me.LabelSuffix.Name = "LabelSuffix"
        Me.LabelSuffix.Size = New System.Drawing.Size(33, 13)
        Me.LabelSuffix.TabIndex = 5
        Me.LabelSuffix.Text = "Suffix"
        '
        'LabelLot
        '
        Me.LabelLot.AutoSize = True
        Me.LabelLot.Location = New System.Drawing.Point(21, 76)
        Me.LabelLot.Name = "LabelLot"
        Me.LabelLot.Size = New System.Drawing.Size(22, 13)
        Me.LabelLot.TabIndex = 6
        Me.LabelLot.Text = "Lot"
        '
        'btnZoom
        '
        Me.btnZoom.Location = New System.Drawing.Point(24, 116)
        Me.btnZoom.Name = "btnZoom"
        Me.btnZoom.Size = New System.Drawing.Size(65, 23)
        Me.btnZoom.TabIndex = 4
        Me.btnZoom.Text = "Zoom"
        Me.btnZoom.UseVisualStyleBackColor = True
        '
        'btnPan
        '
        Me.btnPan.Location = New System.Drawing.Point(117, 116)
        Me.btnPan.Name = "btnPan"
        Me.btnPan.Size = New System.Drawing.Size(65, 23)
        Me.btnPan.TabIndex = 5
        Me.btnPan.Text = "Pan"
        Me.btnPan.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(117, 145)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(65, 23)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(22, 145)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(67, 23)
        Me.btnClear.TabIndex = 6
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'FormPropertyLocator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(212, 190)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnPan)
        Me.Controls.Add(Me.btnZoom)
        Me.Controls.Add(Me.LabelLot)
        Me.Controls.Add(Me.LabelSuffix)
        Me.Controls.Add(Me.LabelSquare)
        Me.Controls.Add(Me.TextLot)
        Me.Controls.Add(Me.TextSuffix)
        Me.Controls.Add(Me.TextSquare)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormPropertyLocator"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Property Search"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextSquare As System.Windows.Forms.TextBox
    Friend WithEvents TextSuffix As System.Windows.Forms.TextBox
    Friend WithEvents TextLot As System.Windows.Forms.TextBox
    Friend WithEvents LabelSquare As System.Windows.Forms.Label
    Friend WithEvents LabelSuffix As System.Windows.Forms.Label
    Friend WithEvents LabelLot As System.Windows.Forms.Label
    Friend WithEvents btnZoom As System.Windows.Forms.Button
    Friend WithEvents btnPan As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button

End Class
