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
Imports ESRI.ArcGIS.Display

Imports System.Timers

<CLSCompliant(False)> _
<ComClass(IABuldManual.ClassId, IABuldManual.InterfaceId, IABuldManual.EventsId)> _
<ProgId("IAIS.IABuldManual")> _
Public Class IABuldManual
    Inherits BaseTool

    Private m_app As IApplication
    Private m_doc As IMxDocument
    Private m_map As IMap

    Private m_Editor As IEditor

    Private m_buld As IFeature
    Private m_col As IFeature

    Private m_isScratch As Boolean

    Private m_selCursor As System.Windows.Forms.Cursor
    Private m_pLineFeed As INewLineFeedback

    Private m_mousedown As Boolean


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5e8379f8-f999-47cf-8e5b-f9d302bf8a38"
    Public Const InterfaceId As String = "822fcc1c-ea87-4105-b48a-29a1038c76ab"
    Public Const EventsId As String = "1a959e05-ffe2-406a-8d0c-729051b7585f"
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
    End Sub 'Unreg
#End Region
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldManual).Assembly.GetManifestResourceStream("IAIS.ia_building_manual.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        Try
            m_selCursor = New System.Windows.Forms.Cursor(GetType(IABuldManual).Assembly.GetManifestResourceStream("IAIS.CursorBuldManual.cur"))
        Catch ex As Exception
            m_selCursor = Nothing
        End Try

        MyBase.m_caption = "Building Assignment(Manual)"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Building Assignment(Manual)"
        MyBase.m_name = "Building Assignment(Manual)"
        MyBase.m_toolTip = "Building Assignment(Manual)"

    End Sub

    Public Overrides Sub OnClick()
        m_map = m_doc.FocusMap
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

    Public Overrides Function OnContextMenu(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Return True
    End Function

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        Dim pPt As IPoint
        pPt = m_Editor.Display.DisplayTransformation.ToMapPoint(X, Y)

        If Not m_buld Is Nothing Then 'AndAlso Not m_pLineFeed Is Nothing Then
            If Not m_pLineFeed Is Nothing Then
                '    m_pLineFeed = New NewLineFeedback
                '    m_pLineFeed.Display = m_doc.ActiveView.ScreenDisplay

                '    Dim pArea As IArea
                '    pArea = m_buld.Shape
                '    m_doc.ActiveView.Refresh()
                '    m_pLineFeed.Start(pArea.Centroid)
                'Else
                m_pLineFeed.MoveTo(pPt)
            End If
        End If
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'm_doc.ActiveView.Refresh()
        'If Not m_pLineFeed Is Nothing Then
        '    m_pLineFeed.Refresh(0)
        'End If

        m_mousedown = False
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        m_mousedown = True
        Try
            Dim pPt As IPoint
            pPt = m_Editor.Display.DisplayTransformation.ToMapPoint(X, Y)
            Dim pSFilter As ISpatialFilter = New SpatialFilter

            pSFilter.Geometry = pPt
            pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin

            Dim pLayer As IFeatureLayer
            Dim pCursor As IFeatureCursor
            Dim pFeatSel As IFeatureSelection

            If m_buld Is Nothing Then
                m_isScratch = False
                Dim sourceLayer As ILayer = IAToolUtil.GetSourceLayer(m_app)
                If TypeOf sourceLayer Is IFeatureLayer Then
                    pLayer = sourceLayer
                    m_isScratch = True
                Else
                    pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_BLDGPLY"), m_map)
                End If

                pCursor = pLayer.Search(pSFilter, False)
                m_buld = pCursor.NextFeature
                pFeatSel = pLayer
                pFeatSel.Clear()

                If Not m_buld Is Nothing Then
                    MapUtil.FlashFeature(m_buld, m_map)

                    pFeatSel.SelectionSet.Add(m_buld.OID)

                    'If m_pLineFeed Is Nothing Then
                    m_pLineFeed = New NewLineFeedback
                    m_pLineFeed.Display = m_doc.ActiveView.ScreenDisplay
                    m_pLineFeed.Start(pPt)

                    'Dim pArea As IArea
                    'pArea = m_buld.Shape
                    'm_doc.ActiveView.Refresh()
                    'm_pLineFeed.Start(pArea.Centroid)
                    'End If
                Else
                    m_app.StatusBar().Message(esriStatusBarPanes.esriStatusMain) = "No building polygon is selected."
                End If
                Return
            Else
                'find parcel
                'pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
                'pCursor = pLayer.Search(pSFilter, False)
                'm_col = pCursor.NextFeature
                'If m_col Is Nothing Then
                '    pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), m_map)
                '    pCursor = pLayer.Search(pSFilter, False)
                '    m_col = pCursor.NextFeature
                'End If

                m_col = MapUtil.SelectedCol(m_map, pSFilter)

                If Not m_col Is Nothing Then
                    'Assign the building to polygon
                    MapUtil.FlashFeature(m_col, m_map)

                    If MsgBox("Assign building to selected property?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Return
                    End If

                    Dim iaFeature As IFeature
                    Dim targetIALayer As IFeatureLayer = IAToolUtil.GetTargetLayer(m_app)



                    If Not MapUtil.GetOverlapShape(m_buld.Shape, targetIALayer) Is Nothing Then
                        MsgBox("Topology check failed. This assignment will create a overlap.")
                        Return
                    End If

                    IAISApplication.GetInstance.StartToolEditing()
                    Try
                        iaFeature = targetIALayer.FeatureClass.CreateFeature
                        iaFeature.Shape = m_buld.ShapeCopy

                        If iaFeature.Fields.FindField("APPEALID") > 0 Then
                            iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(m_col, "SSL")
                            iaFeature.Value(iaFeature.Fields.FindField("IATYPE")) = MapUtil.FEATURETYPE_BUILDING
                        Else
                            iaFeature.Value(iaFeature.Fields.FindField("IAID")) = iaFeature.OID 'Should we get a sequence id here?
                            If m_col.Fields.FindField("GIS_ID") > 0 Then
                                iaFeature.Value(iaFeature.Fields.FindField("OWNER_GIS_ID")) = MapUtil.GetValue(m_col, "GIS_ID")
                            End If

                            iaFeature.Value(iaFeature.Fields.FindField("SSL")) = MapUtil.GetValue(m_col, "SSL")
                            iaFeature.Value(iaFeature.Fields.FindField("FEATURETYPE")) = MapUtil.FEATURETYPE_BUILDING
                            iaFeature.Value(iaFeature.Fields.FindField("ASSIGNBUILD")) = 2

                            iaFeature.Value(iaFeature.Fields.FindField("PROCESSDT")) = Now

                            If m_isScratch Then
                                iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = m_buld.OID
                            Else
                                iaFeature.Value(iaFeature.Fields.FindField("FEATUREOID")) = MapUtil.GetValue(m_buld, "GIS_ID")
                                MapUtil.SetFeatureValue(iaFeature, iaFeature.Fields.FindField("SOURCE"), 3, True)
                            End If

                        End If


                        Dim impAreaPoly As IPolygon = m_buld.Shape
                        iaFeature.Value(iaFeature.Fields.FindField("IARSQM")) = MapUtil.GetPlyArea(impAreaPoly)
                        iaFeature.Value(iaFeature.Fields.FindField("IARSQF")) = MapUtil.GetPlyArea(impAreaPoly) * 10.763910417
                        iaFeature.Store()

                        If m_isScratch Then
                            m_buld.Delete()
                        End If

                        IAISApplication.GetInstance.StopToolEditing("Building Assignment (manual)")

                        m_buld = Nothing
                        m_col = Nothing
                        m_pLineFeed.Stop()
                        m_pLineFeed = Nothing
                        m_doc.ActiveView.Refresh()

                        Return
                    Catch ex As Exception
                        IAISApplication.GetInstance.AbortToolEditing()
                        MsgBox(ex.Message)
                        m_buld = Nothing
                        m_col = Nothing

                        m_doc.FocusMap.ClearSelection()
                        Return
                    End Try
                Else
                    m_app.StatusBar().Message(esriStatusBarPanes.esriStatusMain) = "No property polygon is selected."
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            m_mousedown = False
        End Try

    End Sub



    Public Overrides ReadOnly Property Cursor() As Integer
        Get
            If (Not m_selCursor Is Nothing) Then
                Return m_selCursor.Handle().ToInt32
            Else
                Return 0
            End If
        End Get
    End Property

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

                Return (Not IAToolUtil.GetSourceLayer(m_app) Is Nothing And _
                        Not IAToolUtil.GetTargetLayer(m_app) Is Nothing)

            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg

    Public Sub FlashFeature(ByVal sender As Object, _
           ByVal e As System.Timers.ElapsedEventArgs)
        If Not m_col Is Nothing Then
            MapUtil.FlashFeature(m_col, m_doc.FocusMap)
            Debug.Print("flash feature")
        End If
    End Sub

    ''' <param name="keyCode">Specifies a virtual key code value of the key pressed on the keyboard. For alpha-numeric keys this corresponds to the ASCII value, for example "A" key returns 65 which is the ASCII value for capital A. Other key codes are F1 to F12 are 112 to 123 respectively.</param>
    ''' <param name="Shift">Specifies an integer corresponding to the state of the SHIFT (bit 0), CTRL (bit 1) and ALT (bit 2) keys. When none, some or all of these keys are pressed none, some or all the bits get set. These bits correspond to the values 1, 2, and 4, respectively. For example, if both SHIFT and ALT were pressed, Shift would be 5.</param>
    Public Overrides Sub OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer)
        If Not m_mousedown AndAlso keyCode = System.Windows.Forms.Keys.Escape Then
            m_buld = Nothing
            If Not m_pLineFeed Is Nothing Then
                m_pLineFeed.Stop()
                m_pLineFeed = Nothing
            End If
            m_doc.FocusMap.ClearSelection()
            m_doc.ActiveView.Refresh()

            Return
        End If
    End Sub

End Class


