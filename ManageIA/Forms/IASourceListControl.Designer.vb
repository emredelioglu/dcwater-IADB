<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IASourceListControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.ComboSourceList = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ComboTargetList = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'ComboSourceList
        '
        Me.ComboSourceList.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboSourceList.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboSourceList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboSourceList.FormattingEnabled = True
        Me.ComboSourceList.Items.AddRange(New Object() {"DCGIS", "Scratch"})
        Me.ComboSourceList.Location = New System.Drawing.Point(51, 0)
        Me.ComboSourceList.MaxDropDownItems = 20
        Me.ComboSourceList.Name = "ComboSourceList"
        Me.ComboSourceList.Size = New System.Drawing.Size(88, 21)
        Me.ComboSourceList.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Source"
        '
        'ComboTargetList
        '
        Me.ComboTargetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboTargetList.FormattingEnabled = True
        Me.ComboTargetList.Items.AddRange(New Object() {"IAAssignPly", "RevIAAssignPly", "AppealIAAssignPly"})
        Me.ComboTargetList.Location = New System.Drawing.Point(189, 0)
        Me.ComboTargetList.Name = "ComboTargetList"
        Me.ComboTargetList.Size = New System.Drawing.Size(150, 21)
        Me.ComboTargetList.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(145, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Target"
        '
        'IASourceListControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboTargetList)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboSourceList)
        Me.Name = "IASourceListControl"
        Me.Size = New System.Drawing.Size(344, 21)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboSourceList As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboTargetList As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
