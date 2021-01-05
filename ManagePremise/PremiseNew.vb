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

<CLSCompliant(False)> _
<ComClass(PremiseNew.ClassId, PremiseNew.InterfaceId, PremiseNew.EventsId)> _
Public NotInheritable Class PremiseNew
    Inherits BaseTool

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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PremiseNew).Assembly.GetManifestResourceStream("IAIS.PremiseNew.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Premise New"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Add new premise point"
        MyBase.m_name = "IAIS Premise New"
        MyBase.m_toolTip = "Create new premise point"

        m_digitCur = New System.Windows.Forms.Cursor(GetType(PremiseNew).Assembly.GetManifestResourceStream("IAIS.CursorDigit.cur"))
        m_selCur = New System.Windows.Forms.Cursor(GetType(PremiseNew).Assembly.GetManifestResourceStream("IAIS.CursorSel.cur"))
        m_clickCur = New System.Windows.Forms.Cursor(GetType(PremiseNew).Assembly.GetManifestResourceStream("IAIS.CursorClick.cur"))

        m_toolMode = 0
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)

        'm_objEvents = CType(ObjectClass, IObjectClassEvents)

    End Sub

    Public Overrides Sub OnClick()

        'Dim i As Integer
        'Dim edTask As IEditTask

        'For i = 0 To m_Editor.TaskCount - 1
        '    edTask = m_Editor.Task(i)
        '    If edTask.Name = "Create New Feature" Then
        '        m_Editor.CurrentTask = edTask
        '        Exit For
        '    End If
        'Next i

        'Dim pUid As New UID
        'pUid.Value = "{B479F48A-199D-11D1-9646-0000F8037368}"
        'Dim commandItem As ICommandItem
        'commandItem = m_app.Document.CommandBars.Find(pUid)
        ''m_app.CurrentTool = commandItem
        'If Not commandItem Is Nothing Then
        '    commandItem.Execute()
        'End If

        'Dim pEditLayers As IEditLayers
        'pEditLayers = m_Editor

        'Dim pMap As IMap = m_doc.FocusMap
        'Dim premiseLayer As IFeatureLayer
        'premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", pMap)
        'pEditLayers.SetCurrentLayer(premiseLayer, 0)

        'Dim pObjectClass As IObjectClass
        'pObjectClass = pEditLayers.CurrentLayer.FeatureClass

        'm_objEvents = CType(pObjectClass, IObjectClassEvents_Event)
        'AddHandler m_objEvents.OnCreate, AddressOf OnCreatePremise

        'm_EditEvents2 = CType(m_Editor, IEditEvents_Event)
        'AddHandler m_EditEvents2.OnStopOperation, AddressOf OnStopCreatePremise

    End Sub



    Public Overrides ReadOnly Property Cursor() As Integer
        Get
            If m_toolMode = 0 Then
                If (Not m_digitCur Is Nothing) Then
                    Return m_digitCur.Handle().ToInt32
                Else
                    Return 0
                End If
            ElseIf m_toolMode = 1 Then
                If (Not m_selCur Is Nothing) Then
                    Return m_selCur.Handle().ToInt32
                Else
                    Return 0
                End If
            ElseIf m_toolMode = 2 Then
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
            Enabled = False
            If Not m_Editor Is Nothing AndAlso m_Editor.EditState <> esriEditState.esriStateNotEditing Then
                Enabled = True
            End If
        End Get
    End Property

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'MyBase.OnMouseDown(Button, Shift, X, Y)
        Dim pPoint As IPoint = m_doc.ActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
        Select Case m_toolMode
            Case 0                'Create the premise point 

                If m_form IsNot Nothing AndAlso m_form.Visible Then
                    MsgBox("You have not completed the previous premise point.")
                    Return
                End If


                Dim pMap As IMap = m_doc.FocusMap
                Dim premiseLayer As IFeatureLayer
                premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", pMap)
                Dim premisePoint As IFeature = premiseLayer.FeatureClass.CreateFeature()
                premisePoint.Shape = pPoint

                m_form = New PremiseAttribute
                m_form.PremiseMap = m_doc.FocusMap
                m_form.PremiseApp = m_app
                m_form.PremisePt = premisePoint
                m_form.EditMode = 1

                AddHandler m_form.CommandButtonClicked, AddressOf FormButtonClicked
                m_form.Show(New ModelessDialog(m_app.hWnd))

                Dim pSel As IFeatureSelection
                pSel = premiseLayer
                pSel.Clear()
                pSel.Add(premisePoint)

                m_doc.ActiveView.Refresh()

            Case 1
                Dim pRubberBand As IRubberBand
                Dim pActiveView As IActiveView
                Dim pEnv As IEnvelope

                pActiveView = m_doc.FocusMap
                pRubberBand = New RubberEnvelope
                pEnv = pRubberBand.TrackNew(pActiveView.ScreenDisplay, Nothing)

                Dim pSFilter As ISpatialFilter = New SpatialFilter
                pSFilter.Geometry = pEnv
                pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                Dim addrptLayer As IFeatureLayer = MapUtil.GetLayerByTableName("ADDRESSPT", m_doc.FocusMap)
                Dim addrptSel As IFeatureSelection = addrptLayer
                addrptSel.Clear()
                Dim pFeatCursor As IFeatureCursor = addrptLayer.Search(pSFilter, False)

                Dim addPt As IFeature = pFeatCursor.NextFeature
                Do While Not addPt Is Nothing
                    addrptSel.Add(addPt)
                    addPt = pFeatCursor.NextFeature
                Loop

                If addrptSel.SelectionSet.Count = 0 Then
                    MsgBox("No address point is selected")
                    Return
                ElseIf addrptSel.SelectionSet.Count <> 1 Then
                    MsgBox("Select ony one address point. (" & addrptSel.SelectionSet.Count & " points are selected)")
                Else
                    addrptSel.SelectionSet.Search(Nothing, False, pFeatCursor)
                    addPt = pFeatCursor.NextFeature
                    'm_form.textAddr.Text = MapUtil.GetValue(addPt, "FULLADDRES")
                    m_form.SetMarAddress(MapUtil.GetValue(addPt, "FULLADDRES"), _
                                         MapUtil.GetValue(addPt, "ADDRESS_ID"), _
                                         MapUtil.GetValue(addPt, "ADDRNUM"), _
                                         MapUtil.GetValue(addPt, "STNAME"), _
                                         MapUtil.GetValue(addPt, "STREET_TYP"), _
                                         MapUtil.GetValue(addPt, "ADDRNUMSUF"), _
                                         MapUtil.GetValue(addPt, "QUADRANT"), _
                                         "")

                End If

                m_doc.ActiveView.Refresh()
            Case 2
                Dim premiseLayer As IFeatureLayer
                premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_doc.FocusMap)

                Dim pDataset As IDataset = premiseLayer
                Dim ddotAddr As String = MapUtil.GetAddressAtPoint(pPoint, "IABADMIN.StreetCLTheoric", pDataset.Workspace)

                Dim ddotForm As FormDOTAddress = New FormDOTAddress
                ddotForm.TextAddress.Text = ddotAddr

                ddotForm.ShowDialog()

                If ddotForm.DialogResult = System.Windows.Forms.DialogResult.OK Then
                    m_form.SetDDotAddress(ddotAddr)
                End If

        End Select
    End Sub

    Sub FormButtonClicked(ByVal strCommand As String)
        Select Case strCommand
            Case "GetMAR"
                m_toolMode = 1
            Case "GetDDotAddr"
                m_toolMode = 2
            Case "Ok"
                m_toolMode = 0
            Case "Cancel"
                m_toolMode = 0
        End Select

        'Marshal.ReleaseComObject()
    End Sub


End Class


