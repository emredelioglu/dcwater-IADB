Public Class Form1

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        OpenFileDialog1.Filter = "Text File (*.txt)|*.txt|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
        Dim res As Windows.Forms.DialogResult = OpenFileDialog1.ShowDialog(Me)
        If Not res = Windows.Forms.DialogResult.Cancel Then
            TextFilePath.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        End
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Module1.startLoading(TextFilePath.Text)
    End Sub
End Class