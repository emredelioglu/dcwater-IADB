Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Windows.Forms.DataGridView
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem


Public Class FormSelectPremise
    Private rowindex As Integer
    Private m_map As IMap
    Private m_app As IApplication

    Private m_chargePexuid As Integer

    Public WriteOnly Property PremiseApp() As IApplication
        Set(ByVal value As IApplication)
            m_app = value
        End Set
    End Property

    Public WriteOnly Property PremiseMap() As IMap
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property

    Public WriteOnly Property ChargePexuid() As Integer
        Set(ByVal value As Integer)
            m_chargePexuid = value
        End Set
    End Property

    Public Sub SetPremisePts(ByRef premiseList As IList)
        rowindex = -1
        DataGridPremises.DataSource = premiseList
        Dim i As Integer
        For i = 0 To premiseList.Count - 1
            Dim pt As clsPremisePt = premiseList.Item(i)
            If pt.masterFlag Then
                rowindex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub FormSelectPremise_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim premiseList As IList = New List(Of clsPremisePt)
        'Dim i As Integer

        'For i = 0 To 4
        '    Dim pt As clsPremisePt = New clsPremisePt()
        '    pt.pexuid = 100 + i
        '    pt.imperviousOnly = "Y"
        '    If i = 2 Then
        '        pt.masterFlag = True
        '        rowindex = i
        '    End If
        '    pt.pexprm = "****"
        '    pt.pexsad = "100 main st"

        '    premiseList.Add(pt)
        'Next

        'DataGridPremises.DataSource = premiseList
        If rowindex >= 0 Then
            DataGridPremises.Rows(rowindex).Selected = True
        End If

        'DataGridPremises.s
        'DataGridPremises.DataBind()

    End Sub


    Private Sub DataGridPremises_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridPremises.MouseUp
        Dim hit As HitTestInfo = DataGridPremises.HitTest(e.X, e.Y)
        Dim pt As clsPremisePt

        If hit.Type = DataGridViewHitTestType.Cell Then
            If hit.ColumnIndex = 0 Then
                Dim premiseList As IList = DataGridPremises.DataSource
                Dim i As Integer
                For i = 0 To premiseList.Count - 1
                    pt = premiseList.Item(i)
                    If i = hit.RowIndex Then
                        pt.masterFlag = True
                        rowindex = i
                    ElseIf pt.masterFlag Then
                        pt.masterFlag = False
                        DataGridPremises.InvalidateRow(i)
                    End If
                Next
            Else
                If rowindex >= 0 Then
                    DataGridPremises.Rows(rowindex).Selected = True
                End If
            End If

            Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
            Dim pFilter As IQueryFilter = New QueryFilter
            pFilter.WhereClause = "PEXUID=" & DataGridPremises.Rows.Item(hit.RowIndex).Cells(2).Value
            Dim pCursor As IFeatureCursor = premiseLayer.Search(pFilter, False)
            Dim ptFeature As IFeature = pCursor.NextFeature
            MapUtil.FlashFeature(ptFeature, m_map)

        End If
    End Sub


    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Cursor = Cursors.WaitCursor
        Try

            'set attribute
            Dim master_pexuid As Integer
            Dim master_pexprm As String

            Dim i As Integer = DataGridPremises.SelectedRows.Item(0).Index
            Dim premiseList As IList(Of clsPremisePt) = DataGridPremises.DataSource
            Dim premise As clsPremisePt = premiseList.Item(i)
            master_pexuid = premise.pexuid
            master_pexprm = premise.pexprm

            Dim pUid As New UID
            pUid.Value = "esriEditor.Editor"
            Dim m_Editor As IEditor = m_app.FindExtensionByCLSID(pUid)

            Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
            'Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerFromWS(pFWS, "PremsInterPt")
            Dim pQFilter As IQueryFilter = New QueryFilter

            IAISApplication.GetInstance.StartToolEditing()
            Try

                'Set attribute and close form
                If master_pexuid <> m_chargePexuid Then
                    If m_chargePexuid > 0 Then
                        If MsgBox("Charge carrying feature exists for this property. Do you want to apply the change?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            IAISApplication.GetInstance.AbortToolEditing()
                            Return
                        Else


                            Dim form As FormPickDate = New FormPickDate
                            form.LabelDate.Text = "Effective End Date"
                            form.Text = "Effective End Date"
                            form.ShowDialog()

                            If form.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                                IAISApplication.GetInstance.AbortToolEditing()
                                Return
                            End If

                            Dim pIACFTable As ITable
                            Dim pFeatWS As IFeatureWorkspace
                            Dim pDataset As IDataset

                            pDataset = premiseLayer
                            pFeatWS = pDataset.Workspace

                            pIACFTable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
                            Dim pRow As IRow
                            pQFilter.WhereClause = "EFFENDDT IS NULL AND PEXUID=" & m_chargePexuid

                            Dim pCursor As ICursor
                            pCursor = pIACFTable.Search(pQFilter, False)
                            pRow = pCursor.NextRow
                            If Not pRow Is Nothing Then
                                pRow.Value(pRow.Fields.FindField("EFFENDDT")) = form.DateTimePicker1.Text
                                pRow.Store()

                                Dim pNewIACF As IRow = pIACFTable.CreateRow
                                pNewIACF.Value(pNewIACF.Fields.FindField("PEXUID")) = master_pexuid
                                pNewIACF.Value(pNewIACF.Fields.FindField("IABILLERU")) = pRow.Value(pRow.Fields.FindField("IABILLERU"))
                                pNewIACF.Value(pNewIACF.Fields.FindField("IASQFT")) = pRow.Value(pRow.Fields.FindField("IASQFT"))
                                pNewIACF.Value(pNewIACF.Fields.FindField("EFFSTARTDT")) = form.DateTimePicker1.Text
                                pNewIACF.Value(pNewIACF.Fields.FindField("IA_SOURCE")) = pRow.Value(pRow.Fields.FindField("IA_SOURCE"))
                                pNewIACF.Value(pNewIACF.Fields.FindField("PARCEL_SOURCE")) = pRow.Value(pRow.Fields.FindField("PARCEL_SOURCE"))

                                pNewIACF.Value(pNewIACF.Fields.FindField("DBSTAMPDT")) = Now

                                pNewIACF.Store()
                            End If

                        End If

                    End If
                End If


                Dim pFeatCursor As IFeatureCursor
                Dim pt As IFeature

                For i = 0 To premiseList.Count - 1
                    premise = premiseList.Item(i)
                    pQFilter.WhereClause = "PEXUID=" & premise.pexuid

                    pFeatCursor = premiseLayer.Search(pQFilter, False)
                    pt = pFeatCursor.NextFeature
                    If Not pt Is Nothing Then
                        If premise.pexuid <> master_pexuid Then
                            pt.Value(pt.Fields.FindField("MASTER_PEXUID")) = master_pexuid
                            pt.Value(pt.Fields.FindField("IS_EXEMPT_IAB")) = "Y"
                            If master_pexprm = "" Then
                                pt.Value(pt.Fields.FindField("EXEMPT_IAB_REASON")) = "DUPLI"
                            Else
                                pt.Value(pt.Fields.FindField("EXEMPT_IAB_REASON")) = "SUBOR"
                            End If
                        Else
                            pt.Value(pt.Fields.FindField("MASTER_PEXUID")) = System.DBNull.Value
                            pt.Value(pt.Fields.FindField("IS_EXEMPT_IAB")) = "N"
                            pt.Value(pt.Fields.FindField("EXEMPT_IAB_REASON")) = System.DBNull.Value
                        End If
                        pt.Store()
                    End If
                Next

                IAISApplication.GetInstance.StopToolEditing("Set IA relationship")
                MsgBox("Charge feature assigned.")

                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()

            Catch ex As Exception
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox(ex.Message)
            End Try

        Catch ex As Exception
            Throw ex
        Finally
            Me.Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'close form without making any changes
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class
