Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI

Public Class FormIACFList
    Private m_map As IMap
    Private m_doc As IMxDocument

    Private m_version_count As Integer = 0
    Private m_originalVersion As IVersion

    Public WriteOnly Property PremiseMap() As IMap
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property

    Public WriteOnly Property MapDoc() As IMxDocument
        Set(ByVal value As IMxDocument)
            m_doc = value
        End Set
    End Property

    Public ReadOnly Property VersionCount() As Integer
        Get
            Return m_version_count
        End Get
    End Property


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If DataGridIACF.SelectedRows.Count < 1 Then
            Return
        End If



        Dim dTimeStamp As Date
        dTimeStamp = DataGridIACF.SelectedRows.Item(0).Cells(2).Value
        'dTimeStamp = "10/1/2008"

        Dim poldWS As IWorkspace
        Dim pDataset As IDataset

        Dim premiseLayer As ILayer
        premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)

        pDataset = premiseLayer
        poldWS = pDataset.Workspace

        Dim hWS As IHistoricalWorkspace = CType(poldWS, IHistoricalWorkspace)
        Dim hVersion As IHistoricalVersion = hWS.FindHistoricalVersionByTimeStamp(dTimeStamp)

        Dim pNewWS As IFeatureWorkspace = CType(hVersion, IFeatureWorkspace)

        Dim pVersionUtil As VersionUtil = New VersionUtil
        pVersionUtil.ChangeVersion(m_doc, poldWS, pNewWS)

        'Save the original version information
        If m_version_count = 0 Then
            m_originalVersion = poldWS
        End If

        m_version_count = m_version_count + 1

        'Me.DialogResult = System.Windows.Forms.DialogResult.OK
        'Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub AddIACF(ByRef pRow As IRow)
        DataGridIACF.Rows.Add()
        Dim rowindex As Integer = DataGridIACF.RowCount - 1
        DataGridIACF.Rows.Item(rowindex).Cells().Item(0).Value = pRow.Value(pRow.Fields.FindField("IABILLERU"))
        DataGridIACF.Rows.Item(rowindex).Cells().Item(1).Value = pRow.Value(pRow.Fields.FindField("IASQFT"))
        DataGridIACF.Rows.Item(rowindex).Cells().Item(2).Value = pRow.Value(pRow.Fields.FindField("EFFSTARTDT"))
        DataGridIACF.Rows.Item(rowindex).Cells().Item(3).Value = pRow.Value(pRow.Fields.FindField("EFFENDDT"))
    End Sub

    Private Sub FormIACFList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'If m_version_count > 0 Then
        '    If MsgBox("Do you want to switch to the previous version?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
        '        Dim sVerionName As String = m_originalVersion.VersionName
        '        Dim premiseLayer As ILayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
        '        Dim pDataset As IDataset = premiseLayer
        '        Dim pVersionUtil As VersionUtil = New VersionUtil
        '        pVersionUtil.ChangeVersionByName(m_doc, pDataset.Workspace, sVerionName)
        '    End If
        'End If
        'm_version_count = 0
    End Sub
End Class
