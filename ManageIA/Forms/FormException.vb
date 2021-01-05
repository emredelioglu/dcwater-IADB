Imports ESRI.ArcGIS.Carto

Public Class FormException
    Public Const exceptionPremiseExport As Integer = 1
    Public Const exceptionPremiseImport As Integer = 2
    Public Const exceptionMAR As Integer = 3
    Public Const exceptionOwnerPly As Integer = 4
    Public Const exceptionIAChargeFile As Integer = 5
    Public Const exceptionIAAssignPly As Integer = 6


    Private m_map As IMap
    Private m_exceptionType As Integer

    Public Property ExceptionType() As Integer
        Get
            Return (m_exceptionType)
        End Get
        Set(ByVal value As Integer)
            m_exceptionType = value
        End Set
    End Property

    Public WriteOnly Property GMap() As IMap
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property


    Private Sub btnZoomTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomTo.Click

    End Sub

    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

    End Sub

    Private Sub chkFixed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFixed.CheckedChanged

    End Sub

    Private Sub DataGridExceptions_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExceptions.CellContentClick

    End Sub
End Class