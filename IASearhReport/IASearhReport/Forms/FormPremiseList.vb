Imports System
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Reflection
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display
Imports System.Data
Imports ESRI.ArcGIS.ArcMapUI

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
        End If

    End Sub


    Public Sub SetPremisePts(ByRef premiseList As IList)
        'Dim pList As IBindingList = New BindingList(Of clsPremise)(premiseList)

        DataGridPremises.DataSource = convert2Table(premiseList)
    End Sub

    Private Sub DataGridPremises_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridPremises.MouseClick
        Dim hit As System.Windows.Forms.DataGridView.HitTestInfo = DataGridPremises.HitTest(e.X, e.Y)
        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
        If hit.Type = System.Windows.Forms.DataGridViewHitTestType.Cell Then
            If hit.ColumnIndex = 0 Then
                'zoom to feature
                m_selectedFeature = GetPremise(m_map, DataGridPremises.SelectedRows.Item(0).Cells.Item(8).Value)
                If Not premiseLayer Is Nothing Then
                    MapUtil.SelectFeature(m_selectedFeature, premiseLayer)
                End If
                MapUtil.ZoomToFeature(m_selectedFeature, m_map)


            ElseIf hit.ColumnIndex = 1 Then
                'pan to feature
                m_selectedFeature = GetPremise(m_map, DataGridPremises.SelectedRows.Item(0).Cells.Item(8).Value)
                If Not premiseLayer Is Nothing Then
                    MapUtil.SelectFeature(m_selectedFeature, premiseLayer)
                End If
                MapUtil.PanToFeature(m_selectedFeature, m_map)

            End If
        End If
    End Sub

    Private Function GetPremise(ByRef pMap As IMap, ByVal pexuid As String) As IFeature
        Dim pWS As IFeatureWorkspace = MapUtil.GetPremiseWS(m_map)
        Dim pPremiseFClass As IFeatureClass = pWS.OpenFeatureClass(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"))

        Dim pFilter As IQueryFilter = New QueryFilter
        pFilter.WhereClause = "PEXUID=" & pexuid
        Dim pCursor As IFeatureCursor
        pCursor = pPremiseFClass.Search(pFilter, False)
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

        Try
            Dim pDoc As IMxDocument = IAISApplication.GetInstance.ArcMapApplicaiton.Document
            m_viewEvent = pDoc.ActiveView.FocusMap
            AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature
        Catch ex As Exception
        End Try
    End Sub



    Private Function convert2Table(ByVal list As List(Of clsPremise)) As DataTable
        Dim table As DataTable = New DataTable()

        If list.Count > 0 Then
            Dim properties As Reflection.PropertyInfo() = list(0).GetType().GetProperties()
            Dim columns As List(Of String) = New List(Of String)
            For Each pi In properties
                table.Columns.Add(pi.Name)
                columns.Add(pi.Name)
            Next

            For Each item In list
                Dim cells As Object() = getValues(columns, item)
                table.Rows.Add(cells)
            Next
        End If

        Return table
    End Function

    Private Function getValues(ByVal columns As List(Of String), ByVal instance As Object) As Object()
        Dim ret(columns.Count - 1) As Object
        Dim n As Integer
        For n = 0 To ret.Length - 1
            Dim pi As Reflection.PropertyInfo = instance.GetType().GetProperty(columns(n))
            ret(n) = pi.GetValue(instance, Nothing)
        Next
        Return ret
    End Function

End Class