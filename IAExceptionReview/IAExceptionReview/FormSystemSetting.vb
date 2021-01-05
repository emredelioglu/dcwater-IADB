Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase

Public Class FormSystemSetting

    Protected Friend m_table As itable

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Public Sub AddVariable(ByVal name As String, ByVal value As String)
        DataGridVariable.Rows.Add()
        Dim rowindex As Integer = DataGridVariable.RowCount - 1
        DataGridVariable.Rows.Item(rowindex).Cells().Item(0).Value = name
        DataGridVariable.Rows.Item(rowindex).Cells().Item(1).Value = value
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim pForm As FormVariable = New FormVariable
        pForm.ShowDialog()
        If pForm.DialogResult = System.Windows.Forms.DialogResult.OK Then

            '-------------------------------------------------
            '*  not versioned.
            '-------------------------------------------------
            'Dim pDataset As IDataset = m_table
            'Dim pWorkspaceEdit As IWorkspaceEdit = pDataset.Workspace

            'Dim stopEditingFlag As Boolean = False
            'If Not pWorkspaceEdit.IsBeingEdited Then
            '    pWorkspaceEdit.StartEditing(False)
            '    stopEditingFlag = True
            'End If

            'pWorkspaceEdit.StartEditOperation()
            Dim pRow As IRow = m_table.CreateRow
            pRow.Value(pRow.Fields.FindField("NAME")) = pForm.TextName.Text
            pRow.Value(pRow.Fields.FindField("VALUE")) = pForm.TextValue.Text
            pRow.Store()
            'pWorkspaceEdit.StopEditOperation()

            'If stopEditingFlag Then
            '    pWorkspaceEdit.StopEditing(True)
            'End If

            AddVariable(pForm.TextName.Text, pForm.TextValue.Text)
            IAISToolSetting.AddParameter(pForm.TextName.Text, pForm.TextValue.Text)
        End If
        pForm = Nothing
    End Sub


    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If DataGridVariable.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridVariable.SelectedRows.Item(0)
            Dim pForm As FormVariable = New FormVariable
            pForm.TextName.Text = row.Cells().Item(0).Value
            pForm.TextValue.Text = row.Cells().Item(1).Value
            pForm.ShowDialog()

            If pForm.DialogResult = System.Windows.Forms.DialogResult.OK Then
                Try

                    Dim pDataset As IDataset = m_table
                    pDataset.Workspace.ExecuteSQL("UPDATE IAIS_TOOL_SETTING SET NAME='" & _
                                                  pForm.TextName.Text & "', VALUE='" & _
                                                  pForm.TextValue.Text & "' WHERE NAME='" & _
                                                  row.Cells().Item(0).Value & "' ")

                    IAISToolSetting.UpdateParameter(row.Cells().Item(0).Value, pForm.TextName.Text, pForm.TextValue.Text)

                    row.Cells().Item(0).Value = pForm.TextName.Text
                    row.Cells().Item(1).Value = pForm.TextValue.Text
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            End If

            pForm = Nothing
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If DataGridVariable.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridVariable.SelectedRows.Item(0)
            If MsgBox("Delete variable [" & row.Cells().Item(0).Value & "] from settings?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim pDataset As IDataset = m_table
                pDataset.Workspace.ExecuteSQL("DELETE FROM IAIS_TOOL_SETTING WHERE NAME='" & _
                                              row.Cells().Item(0).Value & "' ")
                IAISToolSetting.DeleteParameter(row.Cells().Item(0).Value)
                DataGridVariable.Rows.Remove(row)
            End If
        End If
    End Sub
End Class
