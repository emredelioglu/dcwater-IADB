<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormPremiseAttributeAdv
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
        Me.chkHAS_LIEN = New System.Windows.Forms.CheckBox
        Me.chkFROM_STANDALONE = New System.Windows.Forms.CheckBox
        Me.cmbLOCTNPRECS = New IAIS.QuietComboBox(Me.components)
        Me.cmbINFORSRC = New IAIS.QuietComboBox(Me.components)
        Me.txtADDRESS_ID = New System.Windows.Forms.TextBox
        Me.btnClearMARId = New System.Windows.Forms.Button
        Me.btnClearMasterPexUid = New System.Windows.Forms.Button
        Me.txtMASTER_PEXUID = New System.Windows.Forms.TextBox
        Me.txtCOMNT = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'chkHAS_LIEN
        '
        Me.chkHAS_LIEN.AutoSize = True
        Me.chkHAS_LIEN.Location = New System.Drawing.Point(31, 13)
        Me.chkHAS_LIEN.Name = "chkHAS_LIEN"
        Me.chkHAS_LIEN.Size = New System.Drawing.Size(68, 17)
        Me.chkHAS_LIEN.TabIndex = 0
        Me.chkHAS_LIEN.Text = "Has Lien"
        Me.chkHAS_LIEN.UseVisualStyleBackColor = True
        '
        'chkFROM_STANDALONE
        '
        Me.chkFROM_STANDALONE.AutoSize = True
        Me.chkFROM_STANDALONE.Location = New System.Drawing.Point(31, 37)
        Me.chkFROM_STANDALONE.Name = "chkFROM_STANDALONE"
        Me.chkFROM_STANDALONE.Size = New System.Drawing.Size(159, 17)
        Me.chkFROM_STANDALONE.TabIndex = 1
        Me.chkFROM_STANDALONE.Text = "Was Created as Standalone"
        Me.chkFROM_STANDALONE.UseVisualStyleBackColor = True
        '
        'cmbLOCTNPRECS
        '
        Me.cmbLOCTNPRECS.FormattingEnabled = True
        Me.cmbLOCTNPRECS.Location = New System.Drawing.Point(31, 88)
        Me.cmbLOCTNPRECS.Name = "cmbLOCTNPRECS"
        Me.cmbLOCTNPRECS.Size = New System.Drawing.Size(232, 21)
        Me.cmbLOCTNPRECS.TabIndex = 2
        '
        'cmbINFORSRC
        '
        Me.cmbINFORSRC.FormattingEnabled = True
        Me.cmbINFORSRC.Location = New System.Drawing.Point(31, 136)
        Me.cmbINFORSRC.Name = "cmbINFORSRC"
        Me.cmbINFORSRC.Size = New System.Drawing.Size(232, 21)
        Me.cmbINFORSRC.TabIndex = 3
        '
        'txtADDRESS_ID
        '
        Me.txtADDRESS_ID.Enabled = False
        Me.txtADDRESS_ID.Location = New System.Drawing.Point(31, 192)
        Me.txtADDRESS_ID.Name = "txtADDRESS_ID"
        Me.txtADDRESS_ID.Size = New System.Drawing.Size(150, 20)
        Me.txtADDRESS_ID.TabIndex = 4
        '
        'btnClearMARId
        '
        Me.btnClearMARId.Location = New System.Drawing.Point(203, 192)
        Me.btnClearMARId.Name = "btnClearMARId"
        Me.btnClearMARId.Size = New System.Drawing.Size(75, 23)
        Me.btnClearMARId.TabIndex = 5
        Me.btnClearMARId.Text = "Clear"
        Me.btnClearMARId.UseVisualStyleBackColor = True
        '
        'btnClearMasterPexUid
        '
        Me.btnClearMasterPexUid.Location = New System.Drawing.Point(203, 236)
        Me.btnClearMasterPexUid.Name = "btnClearMasterPexUid"
        Me.btnClearMasterPexUid.Size = New System.Drawing.Size(75, 23)
        Me.btnClearMasterPexUid.TabIndex = 6
        Me.btnClearMasterPexUid.Text = "Clear"
        Me.btnClearMasterPexUid.UseVisualStyleBackColor = True
        '
        'txtMASTER_PEXUID
        '
        Me.txtMASTER_PEXUID.Enabled = False
        Me.txtMASTER_PEXUID.Location = New System.Drawing.Point(31, 238)
        Me.txtMASTER_PEXUID.Name = "txtMASTER_PEXUID"
        Me.txtMASTER_PEXUID.Size = New System.Drawing.Size(150, 20)
        Me.txtMASTER_PEXUID.TabIndex = 7
        '
        'txtCOMNT
        '
        Me.txtCOMNT.Location = New System.Drawing.Point(31, 288)
        Me.txtCOMNT.Multiline = True
        Me.txtCOMNT.Name = "txtCOMNT"
        Me.txtCOMNT.Size = New System.Drawing.Size(232, 72)
        Me.txtCOMNT.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 69)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Location Precision"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(31, 116)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Information Source"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(31, 173)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "MAR Address ID"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(34, 219)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Master Premise ID"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(34, 269)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Comment"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(50, 377)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 14
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(166, 377)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 15
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'FormPremiseAttributeAdv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(296, 412)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCOMNT)
        Me.Controls.Add(Me.txtMASTER_PEXUID)
        Me.Controls.Add(Me.btnClearMasterPexUid)
        Me.Controls.Add(Me.btnClearMARId)
        Me.Controls.Add(Me.txtADDRESS_ID)
        Me.Controls.Add(Me.cmbINFORSRC)
        Me.Controls.Add(Me.cmbLOCTNPRECS)
        Me.Controls.Add(Me.chkFROM_STANDALONE)
        Me.Controls.Add(Me.chkHAS_LIEN)
        Me.Name = "FormPremiseAttributeAdv"
        Me.Text = "FormPremiseAttributeAdv"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkHAS_LIEN As System.Windows.Forms.CheckBox
    Friend WithEvents chkFROM_STANDALONE As System.Windows.Forms.CheckBox
    Friend WithEvents cmbLOCTNPRECS As IAIS.QuietComboBox
    Friend WithEvents cmbINFORSRC As IAIS.QuietComboBox
    Friend WithEvents txtADDRESS_ID As System.Windows.Forms.TextBox
    Friend WithEvents btnClearMARId As System.Windows.Forms.Button
    Friend WithEvents btnClearMasterPexUid As System.Windows.Forms.Button
    Friend WithEvents txtMASTER_PEXUID As System.Windows.Forms.TextBox
    Friend WithEvents txtCOMNT As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
