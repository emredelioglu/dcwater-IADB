Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<CLSCompliant(False)> _
<ComClass(IABuldAssignment.ClassId, IABuldAssignment.InterfaceId, IABuldAssignment.EventsId)> _
<ProgId("IAIS.IABuldAssignment")> _
Public Class IABuldAssignment
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "aa4d8609-6b1e-477f-a47a-e60305fbe6e7"
    Public Const InterfaceId As String = "b82cd2a6-afe6-47d9-87f2-4d3949c6602d"
    Public Const EventsId As String = "2b935b38-040e-4dc1-818c-73f573f71ad8"
#End Region

#Region "Component Category Registration"
    ' The below automatically adds the Component Category registration.
    <ComRegisterFunction()> Shared _
      Sub Reg(ByVal regKey As [String])
        MxCommands.Register(regKey)
    End Sub 'Reg

    <ComUnregisterFunction()> Shared _
    Sub Unreg(ByVal regKey As [String])
        MxCommands.Unregister(regKey)
    End Sub


#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_perc_mass.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Building Assignment(Percentage Mass)"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Building Assignment(Percentage Mass)"
        MyBase.m_name = "Building Assignment(Percentage Mass)"
        MyBase.m_toolTip = "Building Assignment(Percentage Mass)"

    End Sub

    Public Overrides Sub OnClick()
        'Get the source and target layer list
        Dim pUid As New UID
        pUid.Value = "IAIS.IASourceListControl"
        Dim commandItem As ICommandItem
        commandItem = m_app.Document.CommandBars.Find(pUid)

        Dim sourceList As IASourceListControl
        sourceList = commandItem.Command

        If sourceList.ComboSourceList.SelectedItem = "" Then
            Exit Sub
        End If

        If sourceList.ComboTargetList.SelectedItem = "" Then
            Exit Sub
        End If

        'Dim bldgLayer As IFeatureLayer

        Dim IsScratch As Boolean = False

        Dim sourceLayer As IFeatureLayer
        Dim targetIALayer As IFeatureLayer

        Dim pMap As IMap
        pMap = m_doc.FocusMap

        If sourceList.ComboSourceList.SelectedItem = "Scratch" Then
            sourceLayer = IAToolUtil.GetSourceLayer(m_app)
            IsScratch = True
        Else
            sourceLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_BLDGPLY"), pMap)
            IsScratch = False
        End If

        If sourceLayer Is Nothing Then
            MsgBox("Can't find building layer from source list.")
            Return
        End If

        targetIALayer = IAToolUtil.GetTargetLayer(m_app)

        Dim pColLayer As IFeatureLayer

        Dim pFeatCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pFeatSel As IFeatureSelection

        pFeatSel = sourceLayer
        If pFeatSel.SelectionSet.Count = 0 Then
            MsgBox("No feature is selected from source layer.")
            Return
        ElseIf pFeatSel.SelectionSet.Count > 1 Then
            MsgBox("Please select only one feature from source layer.")
            Return
        End If
        pFeatSel.SelectionSet.Search(Nothing, False, pFeatCursor)

        pFeature = pFeatCursor.NextFeature
        If Not pFeature Is Nothing Then

            If Not IsScratch AndAlso UCase(MapUtil.GetValue(pFeature, "DESCRIPTION")) = "VOID" Then
                MsgBox("The selected building is 'VOID'.")
                Exit Sub
            ElseIf IsScratch AndAlso MapUtil.GetValue(pFeature, "FEATURETYPE") <> 1 Then
                MsgBox("The selected feature from scratch is not building.")
                Exit Sub
            End If

            If Not MapUtil.GetOverlapShape(pFeature.Shape, targetIALayer) Is Nothing Then
                MsgBox("Topology check failed. This assignment will create a overlap.")
                Return
            End If


            Dim pFilter As ISpatialFilter
            pFilter = New SpatialFilter

            pFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            pFilter.Geometry = pFeature.Shape

            Dim colCursor As IFeatureCursor
            Dim pCol As IFeature

            Dim ptopo As ITopologicalOperator
            ptopo = pFeature.Shape

            Dim impAreaPoly As IPolygon
            Dim parcelPoly As IPolygon
            Dim pIntersectPoly As IPolygon

            Dim maxPct As Double
            maxPct = 0

            pCol = MapUtil.GetSelectedCol(pMap)
            If pCol Is Nothing Then
                MsgBox("No property polygon is selected.")
                Return
            Else

                impAreaPoly = pFeature.Shape
                parcelPoly = pCol.Shape
                pIntersectPoly = ptopo.Intersect(pCol.Shape, esriGeometryDimension.esriGeometry2Dimension)

                If MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(impAreaPoly) >= (CDbl(IAISToolSetting.GetParameterValue("BLDG_MASS_PECENTAGE")) / 100) Then

                    IAISApplication.GetInstance.StartToolEditing()
                    Try
                        Dim iaFeature As IFeature
                        iaFeature = targetIALayer.FeatureClass.CreateFeature

                        iaFeature.Shape = pFeature.ShapeCopy

                        If iaFeature.Fields.FindField("APPEALID") > 0 Then
                            'iaFeature.Value(iaFeature.Fields.FindField("APPEALID")) = iaFeature.OID 'Should we get a sequence id here?
                            iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")
                            iaFeature.Value(iaFeature.Fields.FindField("IATYPE")) = MapUtil.FEATURETYPE_BUILDING
                        Else
                            iaFeature.Value(iaFeature.Fields.FindField("IAID")) = iaFeature.OID 'Should we get a sequence id here?
                            If pCol.Fields.FindField("GIS_ID") > 0 Then
                                iaFeature.Value(iaFeature.Fields.FindField("OWNER_GIS_ID")) = MapUtil.GetValue(pCol, "GIS_ID")
                            End If

                            iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")
                            iaFeature.Value(iaFeature.Fields.FindField("FEATURETYPE")) = MapUtil.FEATURETYPE_BUILDING
                            iaFeature.Value(iaFeature.Fields.FindField("ASSIGNBUILD")) = 3
                            iaFeature.Value(iaFeature.Fields.FindField("ASSIGNBUILDPERC")) = CInt((MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(impAreaPoly) * 100) - 0.5)

                            iaFeature.Value(iaFeature.Fields.FindField("PROCESSDT")) = Now

                            If Not IsScratch Then
                                iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = MapUtil.GetValue(pFeature, "GIS_ID")
                                MapUtil.SetFeatureValue(iaFeature, iaFeature.Fields.FindField("SOURCE"), 3, True)
                            Else
                                iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = pFeature.OID
                            End If
                            End If

                            iaFeature.Value(iaFeature.Fields.FindField("IARSQM")) = MapUtil.GetPlyArea(impAreaPoly)
                            iaFeature.Value(iaFeature.Fields.FindField("IARSQF")) = MapUtil.GetPlyArea(impAreaPoly) * 10.763910417

                            iaFeature.Store()

                            If IsScratch Then
                                pFeature.Delete()
                            End If

                            IAISApplication.GetInstance.StopToolEditing("Building Assignment")
                            m_doc.ActiveView.Refresh()

                            Return
                    Catch ex As Exception
                        IAISApplication.GetInstance.AbortToolEditing()
                        MsgBox(ex.Message)
                        Return
                    End Try
                ElseIf maxPct < MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(impAreaPoly) Then
                    maxPct = MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(impAreaPoly)
                End If

            End If

            If maxPct = 0 Then
                MsgBox("Selected building is outside of property.")
            Else
                MsgBox("can't assign building to property. The actual percentage is " & FormatNumber(maxPct, 2))
            End If
        Else
            MsgBox("No building is selected from source layer " & sourceLayer.Name)
        End If

    End Sub

    ''' <param name="hook">
    ''' A reference to the application in which the command was created.
    '''            The hook may be an IApplication reference (for commands created in ArcGIS Desktop applications)
    '''            or an IHookHelper reference (for commands created on an Engine ToolbarControl).
    ''' </param>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)
    End Sub 'Unreg

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Try

                If Not MapUtil.IsExtentionEnabled("IAIS.IAManageExt", m_app) Then
                    Return False
                End If

                If Not IAISToolSetting.isInitilized Then
                    Return False
                End If

                If Not IAISApplication.GetInstance.IsInJTXApplication() Then
                    Return False
                End If


                If m_Editor Is Nothing Then
                    Return False
                End If

                If m_Editor.EditState = esriEditState.esriStateNotEditing Then
                    Return False
                End If

                If (IAToolUtil.GetTargetLayer(m_app) Is Nothing Or _
                        IAToolUtil.GetSourceLayer(m_app) Is Nothing) Then
                    Return False
                End If
                'Dim pMap As IMap = m_doc.FocusMap

                'Dim pColLayer As IFeatureLayer
                'pColLayer = MapUtil.GetLayerByTableName("BldgPly", pMap, False)

                'If pColLayer Is Nothing Then
                '    Return False
                'End If

                'Dim pFeatSel As IFeatureSelection
                'pFeatSel = pColLayer
                'Return (pFeatSel.SelectionSet.Count = 1)

                Dim pMap As IMap = m_doc.FocusMap
                Return MapUtil.IsOneColSelected(pMap)


            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg
End Class


