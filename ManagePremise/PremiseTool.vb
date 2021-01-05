Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase

Imports IAIS.Windows.UI
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.SystemUI

<CLSCompliant(False)> _
<ComClass(PremiseTool.ClassId, PremiseTool.InterfaceId, PremiseTool.EventsId)> _
Public NotInheritable Class PremiseTool
    Inherits BaseTool
    Implements ICommandSubType

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor

    Private m_objEvents As IObjectClassEvents_Event
    Private m_EditEvents2 As IEditEvents2_Event

    Private m_FormPremiseEvent As PremiseAttribute


    Private m_form As PremiseAttribute
    Private m_flagValidPremisePt As Boolean

    Private m_toolMode As Integer

    Private m_digitCur As System.Windows.Forms.Cursor
    Private m_selCur As System.Windows.Forms.Cursor
    Private m_clickCur As System.Windows.Forms.Cursor


    Private m_pPointFeed As IMovePointFeedback

    Private m_subtype As Integer

    Private m_pPremiseEx As PremiseExt


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "461500d8-d9b1-4065-94c6-67bcc658aa82"
    Public Const InterfaceId As String = "20ac3bbb-9ad5-4580-992f-887d3172a2a2"
    Public Const EventsId As String = "a9a96748-ac58-480d-811c-d7ca04ce3246"
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
            m_digitCur = New System.Windows.Forms.Cursor(GetType(PremiseTool).Assembly.GetManifestResourceStream("IAIS.CursorDigit.cur"))
        Catch ex As Exception
            m_digitCur = Nothing
        End Try

        Try
            m_selCur = New System.Windows.Forms.Cursor(GetType(PremiseTool).Assembly.GetManifestResourceStream("IAIS.CursorSel.cur"))
        Catch ex As Exception
            m_selCur = Nothing
        End Try

        Try
            m_clickCur = New System.Windows.Forms.Cursor(GetType(PremiseTool).Assembly.GetManifestResourceStream("IAIS.CursorClick.cur"))
        Catch ex As Exception
            m_clickCur = Nothing
        End Try


    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)

        m_pPremiseEx = MapUtil.GetExtention("IAIS.PremiseExt", m_app)
        If Not m_pPremiseEx Is Nothing Then
            AddHandler m_pPremiseEx.PremiseExtEvent, AddressOf onPremiseExExtEvent
        End If

    End Sub

    Private Sub onPremiseExExtEvent(ByVal strMsg As String)
        If strMsg = "disable" Then
            If Not m_form Is Nothing AndAlso m_form.Visible Then
                m_form.Close()
            End If

            m_toolMode = 0
            MapUtil.SetCurrentTool("esriArcMapUI.SelectFeaturesTool", m_app)

        End If
    End Sub

    Public Overrides Sub OnClick()

        Dim pMap As IMap = m_doc.FocusMap
        Dim premiseLayer As IFeatureLayer
        premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

        If premiseLayer Is Nothing Then
            MsgBox("Premise layer is not loaded.")
            Return
        End If

        If m_subtype = 2 AndAlso m_toolMode = 0 Then
            Try
                If m_form IsNot Nothing AndAlso m_form.Visible Then
                    If m_form.EditMode = 2 Then
                        Return
                    End If
                    MsgBox("You have not completed the previous premise point.")
                    Return
                End If

                'Dim pSel As IFeatureSelection
                'pSel = premiseLayer
                'Dim pCursor As IFeatureCursor
                'pSel.SelectionSet.Search(Nothing, False, pCursor)

                Dim pPt As IFeature
                pPt = MapUtil.GetSelectedPremisePt(m_doc.FocusMap)

                If pPt Is Nothing Then
                    MsgBox("No premise point is selected")
                End If

                m_form = New PremiseAttribute
                m_form.EditMode = 2

                m_form.PremiseMap = m_doc.FocusMap
                m_form.PremiseApp = m_app
                m_form.PremisePt = pPt


                AddHandler m_form.CommandButtonClicked, AddressOf FormButtonClicked
                'm_form.Show(New ModelessDialog(m_app.hWnd))
                IAISToolSetting.OpenForm(m_form)

                StopMovingPoint()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf m_subtype = 1 AndAlso m_toolMode = 0 Then
            m_toolMode = 1
        End If
    End Sub

    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            If m_toolMode = 0 Then
                Return False
            End If
        End Get
    End Property

    Public Overrides ReadOnly Property Cursor() As Integer
        Get
            If m_toolMode = 1 Then
                If (Not m_digitCur Is Nothing) Then
                    Return m_digitCur.Handle().ToInt32
                Else
                    Return 0
                End If
            ElseIf m_toolMode = 2 Then
                If (Not m_selCur Is Nothing) Then
                    Return m_selCur.Handle().ToInt32
                Else
                    Return 0
                End If
            ElseIf m_toolMode = 3 Then
                If (Not m_clickCur Is Nothing) Then
                    Return m_clickCur.Handle().ToInt32
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get

    End Property

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.PremiseExt", m_app) Then
                Return False
            End If

            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            If Not IAISApplication.GetInstance.IsInJTXApplication() Then
                Return False
            End If

            If Not IAISApplication.GetInstance.IsInEditingSession() Then
                Return False
            End If


            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            If premiseLayer Is Nothing Then
                Return False
            End If

            If m_subtype = 2 Then
                'Dim pSel As IFeatureSelection
                'pSel = premiseLayer
                'If pSel.SelectionSet.Count <> 1 Then
                '    Return False
                'End If
                Return MapUtil.GetSelectedPremisePtCount(pMap) = 1
            End If

            Return True

        End Get
    End Property

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        If m_toolMode = 1 Then
            Dim pPoint As IPoint = m_doc.ActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
            Dim pSnapEnv As ISnapEnvironment
            pSnapEnv = m_Editor
            pSnapEnv.SnapPoint(pPoint)
            ShowMovingPoint(pPoint, 1)
        End If
    End Sub


    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'MyBase.OnMouseDown(Button, Shift, X, Y)
        Dim pPoint As IPoint = m_doc.ActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

        If Button <> 1 Then
            Return
        End If

        Try

            Select Case m_toolMode
                Case 1                'Create the premise point 

                    If m_form IsNot Nothing AndAlso m_form.Visible Then
                        MsgBox("You have not completed the previous premise point.")
                        Return
                    End If

                    Dim pSnapEnv As ISnapEnvironment
                    pSnapEnv = m_Editor
                    pSnapEnv.SnapPoint(pPoint)

                    Dim pMap As IMap = m_doc.FocusMap
                    Dim premiseLayer As IFeatureLayer
                    premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)
                    'Dim premisePoint As IFeature = premiseLayer.FeatureClass.CreateFeatureBuffer()
                    Dim premisePoint As IFeature = premiseLayer.FeatureClass.CreateFeature()
                    premisePoint.Shape = pPoint
                    MapUtil.SetFeatureValue(premisePoint, premisePoint.Fields.FindField("PEXPSTS"), "PN", True)
                    MapUtil.SetFeatureValue(premisePoint, premisePoint.Fields.FindField("IS_EXEMPT_IAB"), "Y", True)
                    MapUtil.SetFeatureValue(premisePoint, premisePoint.Fields.FindField("EXEMPT_IAB_REASON"), "PENDG", True)
                    MapUtil.SetFeatureValue(premisePoint, premisePoint.Fields.FindField("USE_PARSED_ADDRESS"), "Y", True)
                    MapUtil.SetFeatureValue(premisePoint, premisePoint.Fields.FindField("IS_IMPERVIOUS_ONLY"), "N", True)


                    m_form = New PremiseAttribute
                    m_form.PremiseMap = m_doc.FocusMap
                    m_form.PremiseApp = m_app
                    m_form.EditMode = 1
                    m_form.PremisePt = premisePoint


                    AddHandler m_form.CommandButtonClicked, AddressOf FormButtonClicked
                    'm_form.Show(New ModelessDialog(m_app.hWnd))
                    IAISToolSetting.OpenForm(m_form)

                    Dim pSel As IFeatureSelection
                    pSel = premiseLayer
                    pSel.Clear()
                    pSel.Add(premisePoint)

                    m_doc.ActiveView.Refresh()

                Case 2
                    Dim pRubberBand As IRubberBand
                    Dim pActiveView As IActiveView
                    Dim pEnv As IEnvelope

                    pActiveView = m_doc.FocusMap
                    pRubberBand = New RubberEnvelope
                    pEnv = pRubberBand.TrackNew(pActiveView.ScreenDisplay, Nothing)

                    Dim pSFilter As ISpatialFilter = New SpatialFilter

                    If Not pEnv.IsEmpty AndAlso pEnv.Width * pEnv.Height <> 0 Then
                        pSFilter.Geometry = pEnv
                    Else
                        Dim pTopo As ITopologicalOperator = pPoint
                        pSFilter.Geometry = pTopo.Buffer(m_doc.SearchTolerance)
                    End If

                    pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                    Dim addrptLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT"), m_doc.FocusMap)

                    If addrptLayer Is Nothing Then
                        MsgBox("Address point is not loaded.")
                        Return
                    End If
                    'Dim addrptSel As IFeatureSelection = addrptLayer
                    'addrptSel.Clear()
                    Dim pFeatCursor As IFeatureCursor = addrptLayer.Search(pSFilter, False)

                    Dim addPt As IFeature = pFeatCursor.NextFeature
                    Dim formMar As FormMarAddress = New FormMarAddress

                    Dim message As String = ""

                    Do While Not addPt Is Nothing
                        'addrptSel.Add(addPt)
                        Dim marAddr As clsMARAddress = New clsMARAddress

                        marAddr.FullAddress = MapUtil.GetValue(addPt, "FULLADDRESS")
                        marAddr.StreetType = MapUtil.GetValue(addPt, "STREET_TYPE")

                        marAddr.AddressId = MapUtil.GetValue(addPt, "ADDRESS_ID")
                        marAddr.AddrNum = MapUtil.GetValue(addPt, "ADDRNUM")
                        marAddr.StName = MapUtil.GetValue(addPt, "STNAME")

                        marAddr.AddrNumSuf = MapUtil.GetValue(addPt, "ADDRNUMSUFFIX")
                        marAddr.Quadrant = MapUtil.GetValue(addPt, "QUADRANT")
                        marAddr.ZipCode = MapUtil.GetValue(addPt, "ZIPCODE")

                        formMar.ListMARAddress.Items.Add(marAddr)

                        'If Trim(marAddr.AddrNum) = "" Then
                        '    message = "The address point you selected does not have address number."
                        'Else
                        '    formMar.ListMARAddress.Items.Add(marAddr)
                        'End If

                        addPt = pFeatCursor.NextFeature
                    Loop

                    If formMar.ListMARAddress.Items.Count = 0 Then
                        If message <> "" Then
                            MsgBox(message)
                        Else
                            MsgBox("No address point is selected")
                        End If
                        Return
                    ElseIf formMar.ListMARAddress.Items.Count <> 1 Then
                        formMar.ListMARAddress.DisplayMember = "FullAddress"
                        formMar.ShowDialog()
                        If formMar.DialogResult = System.Windows.Forms.DialogResult.OK Then
                            Dim marAddr As clsMARAddress = formMar.ListMARAddress.SelectedItem
                            m_form.SetMarAddress(marAddr.FullAddress, _
                                             marAddr.AddressId, _
                                             marAddr.AddrNum, _
                                             marAddr.StName, _
                                             marAddr.StreetType, _
                                             marAddr.AddrNumSuf, _
                                             marAddr.Quadrant, _
                                             "", _
                                             marAddr.ZipCode)
                        End If

                        m_form.WindowState = System.Windows.Forms.FormWindowState.Normal
                    Else
                        'addrptSel.SelectionSet.Search(Nothing, False, pFeatCursor)
                        'addPt = pFeatCursor.NextFeature
                        'm_form.SetMarAddress(MapUtil.GetValue(addPt, "FULLADDRESS"), _
                        '                     MapUtil.GetValue(addPt, "ADDRESS_ID"), _
                        '                     MapUtil.GetValue(addPt, "ADDRNUM"), _
                        '                     MapUtil.GetValue(addPt, "STNAME"), _
                        '                     MapUtil.GetValue(addPt, "STREET_TYPE"), _
                        '                     MapUtil.GetValue(addPt, "ADDRNUMSUFFIX"), _
                        '                     MapUtil.GetValue(addPt, "QUADRANT"), _
                        '                     "", _
                        '                     MapUtil.GetValue(addPt, "ZIPCODE"))

                        Dim marAddr As clsMARAddress = formMar.ListMARAddress.Items.Item(0)
                        If marAddr.AddrNum = "" Then
                            MsgBox("The address point you selected does not have address number.")
                            Return
                        Else
                            m_form.SetMarAddress(marAddr.FullAddress, _
                                             marAddr.AddressId, _
                                             marAddr.AddrNum, _
                                             marAddr.StName, _
                                             marAddr.StreetType, _
                                             marAddr.AddrNumSuf, _
                                             marAddr.Quadrant, _
                                             "", _
                                             marAddr.ZipCode)

                            m_form.WindowState = System.Windows.Forms.FormWindowState.Normal
                        End If

                    End If

                    'm_doc.ActiveView.Refresh()
                Case 3
                    Dim premiseLayer As IFeatureLayer
                    premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_doc.FocusMap)

                    Dim pDataset As IDataset = premiseLayer
                    Dim ddotAddr As String = MapUtil.GetAddressAtPoint(pPoint, IAISToolSetting.GetParameterValue("DDOT_LOCATOR_NAME"), pDataset.Workspace)

                    Dim ddotForm As FormDOTAddress = New FormDOTAddress
                    ddotForm.TextAddress.Text = ddotAddr

                    ddotForm.ShowDialog()

                    If ddotForm.DialogResult = System.Windows.Forms.DialogResult.OK Then
                        m_form.SetDDotAddress(ddotAddr)
                    End If

            End Select

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Sub FormButtonClicked(ByVal strCommand As String)
        Select Case strCommand
            Case "GetMAR"
                m_toolMode = 2
            Case "GetDDotAddr"
                m_toolMode = 3
            Case "Ok"
                If m_subtype = 1 Then
                    m_toolMode = 1
                Else
                    m_toolMode = 0
                    MapUtil.SetCurrentTool("esriArcMapUI.SelectFeaturesTool", m_app)
                End If

                Dim premiseLayer As IFeatureLayer
                premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_doc.FocusMap)
                m_doc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, premiseLayer, Nothing)
            Case "Cancel"
                If m_subtype = 1 Then
                    m_toolMode = 1
                Else
                    m_toolMode = 0
                    MapUtil.SetCurrentTool("esriArcMapUI.SelectFeaturesTool", m_app)
                End If
            Case "FormClosing"
                If m_subtype = 1 Then
                    m_toolMode = 1
                Else
                    m_toolMode = 0
                    MapUtil.SetCurrentTool("esriArcMapUI.SelectFeaturesTool", m_app)
                End If
        End Select

    End Sub


    Public Function GetCount() As Integer Implements ESRI.ArcGIS.SystemUI.ICommandSubType.GetCount
        Return 2
    End Function

    Public Sub SetSubType(ByVal SubType As Integer) Implements ESRI.ArcGIS.SystemUI.ICommandSubType.SetSubType
        m_subtype = SubType

        Try
            If m_subtype = 1 Then
                MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PremiseTool).Assembly.GetManifestResourceStream("IAIS.PremiseNew.bmp"))
            ElseIf m_subtype = 2 Then
                MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PremiseTool).Assembly.GetManifestResourceStream("IAIS.PremiseUpdate.bmp"))
            End If
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        If m_subtype = 1 Then
            MyBase.m_caption = "Premise New"
            MyBase.m_category = "IAIS Tools"
            MyBase.m_message = "Add new premise point"
            MyBase.m_name = "IAIS Premise New"
            MyBase.m_toolTip = "Create new premise point"

            m_toolMode = 1
        ElseIf m_subtype = 2 Then
            MyBase.m_caption = "Premise Update"
            MyBase.m_category = "IAIS Tools"
            MyBase.m_message = "Update premise point"
            MyBase.m_name = "IAIS Premise Update"
            MyBase.m_toolTip = "Update premise feature"
        End If
    End Sub

    Private Sub ShowMovingPoint(ByVal pt As IPoint, ByVal ptType As Integer)

        Dim symbol As ISymbol
        Dim pSmplMarker As ISimpleMarkerSymbol
        Dim pRGBClr As IRgbColor

        If m_pPointFeed Is Nothing Then
            m_pPointFeed = New MovePointFeedback
            m_pPointFeed.Display = m_doc.ActiveView.ScreenDisplay

            symbol = m_pPointFeed.Symbol
            pSmplMarker = symbol
            pSmplMarker.Size = 6
            pRGBClr = New RgbColor

            pRGBClr.Red = 0
            pRGBClr.Green = 255
            pRGBClr.Blue = 255
            pSmplMarker.Color = pRGBClr
            pSmplMarker.Style = esriSimpleMarkerStyle.esriSMSCircle

            Dim OutlineColor As IColor
            OutlineColor = New RgbColor
            OutlineColor.RGB = RGB(0, 0, 0)

            pSmplMarker.Outline = True
            pSmplMarker.OutlineSize = 0.5
            pSmplMarker.OutlineColor = OutlineColor

            m_pPointFeed.Start(pt, pt)
        Else
            m_pPointFeed.MoveTo(pt)
        End If


    End Sub

    Private Sub StopMovingPoint()
        If Not m_pPointFeed Is Nothing Then
            m_pPointFeed.Stop()
            m_pPointFeed = Nothing
            m_doc.ActivatedView.PartialRefresh(esriViewDrawPhase.esriViewForeground + esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
        End If

    End Sub

    ''' <returns>A boolean value indicating if the tool can be interrupted by other tools.</returns>
    Public Overrides Function Deactivate() As Boolean
        If Not m_form Is Nothing AndAlso m_form.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
            m_form.WindowState = System.Windows.Forms.FormWindowState.Normal
        End If

        StopMovingPoint()
        Return True
    End Function

    Public Function ConvertPixelsToMapUnits(ByVal pActiveView As IActiveView, ByVal pixelUnits As Double) As Double
        Dim realWorldDisplayExtent As Double
        Dim pixelExtent As Integer
        Dim sizeOfOnePixel As Double
        Dim deviceRect As tagRECT

        deviceRect = pActiveView.ScreenDisplay.DisplayTransformation.DeviceFrame
        pixelExtent = deviceRect.right - deviceRect.left
        realWorldDisplayExtent = pActiveView.ScreenDisplay.DisplayTransformation.VisibleBounds.Width
        sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
        ConvertPixelsToMapUnits = pixelUnits * sizeOfOnePixel

    End Function

End Class


