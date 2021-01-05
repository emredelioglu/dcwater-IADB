
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.ADF.CATIDs

<CLSCompliant(False)> _
<ComClass(IAAssign.ClassId, IAAssign.InterfaceId, IAAssign.EventsId)> _
<ProgId("IAIS.IAAssign")> _
Public Class IAAssign
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ba0a8e5f-40f2-47ff-bef1-6d31ff46a264"
    Public Const InterfaceId As String = "e0d81d87-bc23-4177-88bb-8d6ea91e95d9"
    Public Const EventsId As String = "e2333a9e-1835-49e7-8371-d0d609993dd7"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_ia_assign.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "IA Assignment"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "IA Assignment"
        MyBase.m_name = "IA Assignment"
        MyBase.m_toolTip = "IA Assignment"
    End Sub


    Public Overrides Sub OnClick()
        Dim pMap As IMap
        pMap = m_doc.FocusMap

        Dim pCol As IFeature = MapUtil.GetSelectedCol(pMap)

        Dim sourceLayer As ILayer
        Dim targetIALayer As IFeatureLayer
        sourceLayer = IAToolUtil.GetSourceLayer(m_app)
        targetIALayer = IAToolUtil.GetTargetLayer(m_app)

        Dim sourceIALayer As IFeatureLayer
        Dim pSFilter As ISpatialFilter = New SpatialFilter
        pSFilter.Geometry = pCol.Shape
        pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

        Dim iaFlag As Boolean = False

        Dim pCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pDataset As IDataset
        Dim isBuild As Boolean = False

        If TypeOf sourceLayer Is GroupLayer Then
            'DCGIS Layer
            Dim iaCompLayer As ICompositeLayer
            iaCompLayer = sourceLayer

            IAISApplication.GetInstance.StartToolEditing()
            Try
                Dim i As Integer
                For i = 0 To iaCompLayer.Count - 1
                    isBuild = False
                    sourceIALayer = iaCompLayer.Layer(i)
                    pDataset = sourceIALayer

                    If MapUtil.GetTableName(pDataset.BrowseName) = "BLDGPLY" Then
                        isBuild = True
                    End If

                    pCursor = sourceIALayer.Search(pSFilter, False)
                    pFeature = pCursor.NextFeature
                    Do While Not pFeature Is Nothing

                        Dim ptopo As ITopologicalOperator
                        ptopo = pFeature.Shape

                        Dim pIntersectPoly As IPolygon
                        Dim pIAPoly As IPolygon = pFeature.Shape
                        pIntersectPoly = ptopo.Intersect(pCol.Shape, esriGeometryDimension.esriGeometry2Dimension)


                    
                        If MapUtil.GetPlyArea(pIntersectPoly) > 0 Then
                            If Not isBuild Or _
                                ((MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(pIAPoly)) > (CDbl(IAISToolSetting.GetParameterValue("BLDG_MASS_PECENTAGE")) / 100)) Then

                                If Not MapUtil.GetOverlapShape(pIntersectPoly, targetIALayer) Is Nothing Then
                                    Throw New Exception("Topology check failed. This assignment will create a overlap.")
                                End If



                                Dim iaFeature As IFeature
                                iaFeature = targetIALayer.FeatureClass.CreateFeature

                                If Not isBuild Then
                                    iaFeature.Shape = pIntersectPoly
                                    iaFeature.Value(iaFeature.Fields.FindField("IARSQM")) = MapUtil.GetPlyArea(pIntersectPoly)
                                    iaFeature.Value(iaFeature.Fields.FindField("IARSQF")) = MapUtil.GetPlyArea(pIntersectPoly) * 10.763910417
                                Else
                                    iaFeature.Shape = pFeature.Shape
                                    iaFeature.Value(iaFeature.Fields.FindField("IARSQM")) = MapUtil.GetPlyArea(pIAPoly)
                                    iaFeature.Value(iaFeature.Fields.FindField("IARSQF")) = MapUtil.GetPlyArea(pIAPoly) * 10.763910417
                                End If

                                If iaFeature.Fields.FindField("APPEALID") > 0 Then
                                    iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")
                                    iaFeature.Value(iaFeature.Fields.FindField("IATYPE")) = IAToolUtil.GetIAType(MapUtil.GetTableName(pDataset.BrowseName))
                                Else
                                    iaFeature.Value(iaFeature.Fields.FindField("IAID")) = iaFeature.OID
                                    If pCol.Fields.FindField("GIS_ID") > 0 Then
                                        iaFeature.Value(iaFeature.Fields.FindField("OWNER_GIS_ID")) = MapUtil.GetValue(pCol, "GIS_ID")
                                    End If

                                    iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")

                                    iaFeature.Value(iaFeature.Fields.FindField("FEATURETYPE")) = IAToolUtil.GetIAType(MapUtil.GetTableName(pDataset.BrowseName))

                                    iaFeature.Value(iaFeature.Fields.FindField("ASSIGN")) = 1

                                    iaFeature.Value(iaFeature.Fields.FindField("PROCESSDT")) = Now
                                    MapUtil.SetFeatureValue(iaFeature, iaFeature.Fields.FindField("SOURCE"), 3, True)

                                    If pFeature.Fields.FindField("GIS_ID") > 0 Then
                                        iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = MapUtil.GetValue(pFeature, "GIS_ID")
                                    Else
                                        iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = pFeature.OID
                                    End If

                                End If

                                iaFeature.Store()
                                iaFlag = True
                            End If
                        End If

                        pFeature = pCursor.NextFeature

                    Loop

                Next

                If iaFlag Then
                    IAISApplication.GetInstance.StopToolEditing("IA Assignment")
                    m_doc.ActiveView.Refresh()
                Else
                    IAISApplication.GetInstance.AbortToolEditing()
                    MsgBox("Selected parcel does not have any IA area from source layer")
                End If
            Catch ex As Exception
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox(ex.Message)
            End Try

        Else
            'Scratch layer
            sourceIALayer = sourceLayer
            pCursor = sourceIALayer.Search(pSFilter, False)
            pFeature = pCursor.NextFeature
            IAISApplication.GetInstance.StartToolEditing()
            Try
                Do While Not pFeature Is Nothing

                    Dim ptopo As ITopologicalOperator
                    ptopo = pFeature.Shape

                    Dim pIntersectPoly As IPolygon
                    pIntersectPoly = ptopo.Intersect(pCol.Shape, esriGeometryDimension.esriGeometry2Dimension)
                    Dim pIAPoly As IPolygon = pFeature.Shape


                    If MapUtil.GetValue(pFeature, "FEATURETYPE") <> "1" Then


                        If MapUtil.GetPlyArea(pIntersectPoly) > 0 Then
                            If Not isBuild Or _
                             (MapUtil.GetPlyArea(pIntersectPoly) / MapUtil.GetPlyArea(pIAPoly)) > (CDbl(IAISToolSetting.GetParameterValue("BLDG_MASS_PECENTAGE", m_app)) / 100) Then


                                If Not MapUtil.GetOverlapShape(CType(pIntersectPoly, ESRI.ArcGIS.Geometry.IGeometry), targetIALayer) Is Nothing Then
                                    Throw New Exception("Topology check failed. This assignment will create a overlap.")
                                End If

                                Dim iaFeature As IFeature
                                iaFeature = targetIALayer.FeatureClass.CreateFeature

                                iaFeature.Shape = pIntersectPoly
                                iaFeature.Value(iaFeature.Fields.FindField("IARSQM")) = MapUtil.GetPlyArea(pIntersectPoly)
                                iaFeature.Value(iaFeature.Fields.FindField("IARSQF")) = MapUtil.GetPlyArea(pIntersectPoly) * 10.763910417

                                If iaFeature.Fields.FindField("APPEALID") > 0 Then
                                    iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")
                                    iaFeature.Value(iaFeature.Fields.FindField("IATYPE")) = MapUtil.GetValue(pFeature, "FEATURETYPE")
                                Else
                                    iaFeature.Value(iaFeature.Fields.FindField("IAID")) = iaFeature.OID
                                    If pCol.Fields.FindField("GIS_ID") > 0 Then
                                        iaFeature.Value(iaFeature.Fields.FindField("OWNER_GIS_ID")) = MapUtil.GetValue(pCol, "GIS_ID")
                                    End If
                                    iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(pCol, "SSL")

                                    iaFeature.Value(iaFeature.Fields.FindField("FEATURETYPE")) = MapUtil.GetValue(pFeature, "FEATURETYPE")

                                    iaFeature.Value(iaFeature.Fields.FindField("ASSIGN")) = 1

                                    iaFeature.Value(iaFeature.Fields.FindField("PROCESSDT")) = Now
                                    iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = pFeature.OID
                                End If
                                iaFeature.Store()
                                iaFlag = True

                            End If

                        End If

                        'pFeature.Delete()
                        pFeature.Shape = ptopo.Difference(pCol.Shape)
                        pFeature.Store()

                    End If

                    pFeature = pCursor.NextFeature
                Loop

                If iaFlag Then
                    IAISApplication.GetInstance.StopToolEditing("IA Assignment")
                    m_doc.ActiveView.Refresh()
                    Return
                Else
                    IAISApplication.GetInstance.AbortToolEditing()
                    MsgBox("Selected parcel does not have any IA area from source layer")
                End If
            Catch ex As Exception
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox(ex.Message)
            End Try
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

    End Sub


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

                If IAToolUtil.GetSourceLayer(m_app) Is Nothing Or _
                    IAToolUtil.GetTargetLayer(m_app) Is Nothing Then
                    Return False
                End If

                Dim pMap As IMap = m_doc.FocusMap
                Return MapUtil.IsOneColSelected(pMap)
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg


End Class


