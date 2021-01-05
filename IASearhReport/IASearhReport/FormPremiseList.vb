Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display

Public Class FormPremiseList
    Protected Friend m_map As IMap

    Friend Shared formWidth As Integer
    Friend Shared colWidth(6) As Integer

    Private m_selectedFeature As IFeature
    Private m_viewEvent As IActiveViewEvents_Event


    Private Sub flashFeature(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase)

        If Not m_selectedFeature Is Nothing AndAlso phase = esriViewDrawPhase.esriViewForeground Then
            MapUtil.FlashFeature(m_selectedFeature, m_map)
            m_selectedFeature = Nothing

            RemoveHandler m_viewEvent.AfterDraw, AddressOf flashFeature
        End If

    End Sub


    Public Sub SetPremisePts(ByRef premiseList As IList)
        DataGridPremises.DataSource = premiseList
    End Sub

    Private Sub DataGridPremises_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridPremises.MouseClick
        Dim hit As System.Windows.Forms.DataGridView.HitTestInfo = DataGridPremises.HitTest(e.X, e.Y)
        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
        If hit.Type = System.Windows.Forms.DataGridViewHitTestType.Cell Then
            If hit.ColumnIndex = 0 Then
                'zoom to feature
                m_selectedFeature = GetPremise(m_map, DataGridPremises.SelectedRows.Item(0).Cells.Item(8).Value)
                MapUtil.SelectFeature(m_selectedFeature, premiseLayer)
                MapUtil.ZoomToFeature(m_selectedFeature, m_map)

                m_viewEvent = m_map
                AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature

            ElseIf hit.ColumnIndex = 1 Then
                'pan to feature
                m_selectedFeature = GetPremise(m_map, DataGridPremises.SelectedRows.Item(0).Cells.Item(8).Value)
                MapUtil.SelectFeature(m_selectedFeature, premiseLayer)
                MapUtil.PanToFeature(m_selectedFeature, m_map)

                m_viewEvent = m_map
                AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature

            End If
        End If
    End Sub

    Private Function GetPremise(ByRef pMap As IMap, ByVal pexuid As String) As IFeature
        Dim premiseLayer As IFeatureLayer
        premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
        Dim pFilter As IQueryFilter = New QueryFilter
        pFilter.WhereClause = "PEXUID=" & pexuid
        Dim pCursor As IFeatureCursor
        pCursor = premiseLayer.Search(pFilter, False)
        Return pCursor.NextFeature
    End Function

    Private Sub FormPremiseList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        formWidth = Width
        colWidth(0) = DataGridPremises.Columns(2).Width
        colWidth(1) = DataGridPremises.Columns(3).Width
        colWidth(2) = DataGridPremises.Columns(4).Width
        colWidth(3) = DataGridPremises.Columns(5).Width
        colWidth(4) = DataGridPremises.Columns(6).Width
        colWidth(5) = DataGridPremises.Columns(7).Width

    End Sub

    Private Sub FormPremiseList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If formWidth > 0 Then
            Me.Width = formWidth
            DataGridPremises.Columns(2).Width = colWidth(0)
            DataGridPremises.Columns(3).Width = colWidth(1)
            DataGridPremises.Columns(4).Width = colWidth(2)
            DataGridPremises.Columns(5).Width = colWidth(3)
            DataGridPremises.Columns(6).Width = colWidth(4)
            DataGridPremises.Columns(7).Width = colWidth(5)
        End If

    End Sub
End Class