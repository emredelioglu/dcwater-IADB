Imports System.Windows.Forms

Public Class FormVariable
    Public editMode As enumEditmode

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If editMode = enumEditmode.enumEditNew Then
            Try
                IAISToolSetting.GetParameterValue(TextName.Text)
                MsgBox("Variable " & TextName.Text & " exists. Please use update or select a different name")
                Return
            Catch ex As Exception
                'Expecting exception
            End Try
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FormVariable_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If editMode = enumEditmode.enumEditUpdate Then
            TextName.Enabled = False
        Else
            TextName.Enabled = True
        End If
    End Sub
End Class

Public Enum enumEditmode As Integer
    enumEditNew = 1
    enumEditUpdate = 2
End Enum