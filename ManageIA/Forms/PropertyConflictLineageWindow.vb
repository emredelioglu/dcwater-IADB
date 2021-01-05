Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.JTX
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Editor

<ComClass(PropertyConflictLineageWindow.ClassId, PropertyConflictLineageWindow.InterfaceId, PropertyConflictLineageWindow.EventsId), _
 ProgId("IAIS.PropertyConflictLineageWindow")> _
Public Class PropertyConflictLineageWindow
    Implements IDockableWindowDef


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4467D820-6F41-41de-A08F-AA792D29FCB1"
    Public Const InterfaceId As String = "47ED2B59-E26E-4271-A774-640CC414306F"
    Public Const EventsId As String = "C622197D-89A9-4753-BF7C-9F4073F4F963"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxDockableWindows.Register(regKey)
        GxDockableWindows.Register(regKey)
        SxDockableWindows.Register(regKey)
        GMxDockableWindows.Register(regKey)
    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxDockableWindows.Unregister(regKey)
        GxDockableWindows.Unregister(regKey)
        SxDockableWindows.Unregister(regKey)
        GMxDockableWindows.Unregister(regKey)
    End Sub

#End Region
#End Region

    Private m_application As IApplication

    Private m_map As IMap
    Private m_exceptionType As Integer

    Private m_expFeature As ITopologicalOperator

    Private m_viewEvent As IActiveViewEvents_Event

    Private m_oidIndex As Integer

    Private m_Editor As IEditor
    Private m_bInEditingSession As Boolean = False


    Private Sub flashFeature(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase)

        If Not m_expFeature Is Nothing AndAlso phase = esriViewDrawPhase.esriViewForeground Then
            MapUtil.FlashGeometry(m_expFeature, m_map)
            m_expFeature = Nothing

            RemoveHandler m_viewEvent.AfterDraw, AddressOf flashFeature
        End If

    End Sub

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


    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Caption
        Get
            Return "Property Conflict Lineage Review"
        End Get
    End Property

    Public ReadOnly Property ChildHWND() As Integer Implements ESRI.ArcGIS.Framework.IDockableWindowDef.ChildHWND
        Get
            Return Me.Handle.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Name
        Get
            Return "IAIS_PropertyConflictLineageWindow"
        End Get
    End Property

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnCreate
        m_application = CType(hook, IApplication)
        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_application.FindExtensionByCLSID(pUid)

        AddHandler DirectCast(m_Editor, IEditEvents_Event).OnStartEditing, AddressOf OnStartEditing
        AddHandler DirectCast(m_Editor, IEditEvents_Event).OnStopEditing, AddressOf OnStopEditing
    End Sub

    Public Sub OnStartEditing()

        'CheckBox to see if the version match the Job version
        Dim pVersion As IVersion = TryCast(m_Editor.EditWorkspace, IVersion)
        If pVersion Is Nothing Then
            m_bInEditingSession = False
            Return
        End If

        Dim jtxExt As IJTXExtension = DirectCast(m_application.FindExtensionByName("Workflow Manager"), IJTXExtension)
        If jtxExt.Job Is Nothing Then
            m_bInEditingSession = False
            Return
        End If

        If MapUtil.GetVersionName(pVersion.VersionName) <> MapUtil.GetVersionName(jtxExt.Job.VersionName) Then
            MsgBox("Version of current editing workspace does not match the Job version")
            m_bInEditingSession = False
        Else
            m_bInEditingSession = True
            Me.comboActions.Enabled = True
            Me.btnApply.Enabled = True
        End If
    End Sub

    Public Sub OnStopEditing()
        m_bInEditingSession = False
        Me.comboActions.Enabled = False
        Me.btnApply.Enabled = False

    End Sub

    Public Sub OnDestroy() Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnDestroy
        Me.Dispose()
    End Sub

    Public ReadOnly Property UserData() As Object Implements ESRI.ArcGIS.Framework.IDockableWindowDef.UserData
        Get
            Return Me
        End Get
    End Property

    Private Function GetColIndex(ByRef dg As DataGridView, ByVal colname As String) As Integer
        Dim i As Integer
        For i = 0 To dg.ColumnCount - 1
            Dim col As DataGridViewColumn = DataGridExceptions.Columns().Item(i)
            If UCase(col.Name) = UCase(colname) Then
                Return i
            End If
        Next

        Return -1
    End Function

    Private Sub btnZoomTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomTo.Click
        GotoException("zoom")
    End Sub

    Private Sub GotoException(ByVal functionName As String)
        If DataGridExceptions.SelectedRows.Count = 0 Then
            Return
        End If

        Dim i As Integer = DataGridExceptions.SelectedRows.Item(0).Index

        Dim ssl1, ssl2 As String
        Dim index As Integer = GetColIndex(DataGridExceptions, "INPUTFEATURESSL")
        ssl1 = DataGridExceptions.SelectedRows.Item(0).Cells(index).Value
        index = GetColIndex(DataGridExceptions, "CONFLICTFEATURESSL")
        ssl2 = DataGridExceptions.SelectedRows.Item(0).Cells(index).Value

        Dim pLayer As IFeatureLayer
        Dim pCursor As IFeatureCursor

        Dim pFilter As IQueryFilter = New QueryFilter
        pLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
        If pLayer Is Nothing Then
            MsgBox("OwnerPly layer is not loaded")
            Return
        End If

        pFilter.WhereClause = "SSL='" & ssl1 & "' OR SSL='" & ssl2 & "'"
        pCursor = pLayer.Search(pFilter, False)

        Dim pGeometryBag As IGeometryCollection = New GeometryBag
        Dim pFeature As IFeature
        pFeature = pCursor.NextFeature
        Do While Not pFeature Is Nothing
            If pGeometryBag.GeometryCount = 0 Then
                MapUtil.SelectFeature(pFeature, pLayer)
            Else
                MapUtil.SelectFeature(pFeature, pLayer, False)
            End If

            pGeometryBag.AddGeometry(pFeature.ShapeCopy)
            pFeature = pCursor.NextFeature

        Loop

        m_viewEvent = m_map
        AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature

        If pGeometryBag.GeometryCount > 0 Then
            If pGeometryBag.GeometryCount > 1 Then
                m_expFeature = New Polygon
                m_expFeature.ConstructUnion(pGeometryBag)
            Else
                m_expFeature = pGeometryBag.Geometry(0)
            End If

            If functionName = "zoom" Then
                MapUtil.ZoomToGeometry(m_expFeature, m_map)
            ElseIf functionName = "pan" Then
                MapUtil.PanToGeometry(m_expFeature, m_map)
            Else
                Dim pActiveview As IActiveView = m_map
                'pActiveview.Refresh()
                pActiveview.PartialRefresh(esriViewDrawPhase.esriViewForeground, Nothing, Nothing)
            End If
        Else
            MsgBox("Can't find the ownerPly with SSL " & ssl1 & "," & ssl2)
        End If

    End Sub

    Private Sub DataGridExceptions_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridExceptions.SelectionChanged
        If DataGridExceptions.SelectedRows.Count = 0 Then
            btnZoomTo.Enabled = False
            btnPan.Enabled = False

            Me.comboActions.Enabled = False
            Me.btnApply.Enabled = False

        Else
            If Not m_bInEditingSession Then
                Me.comboActions.Enabled = False
                Me.btnApply.Enabled = False
            Else
                Me.comboActions.Enabled = True
                Me.btnApply.Enabled = True
            End If

            btnZoomTo.Enabled = True
            btnPan.Enabled = True

            GotoException("zoom")

        End If
    End Sub


    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click
        GotoException("pan")
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        If comboActions.SelectedIndex < 0 Then
            Return
        End If

        Dim pLayer As IFeatureLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
        If pLayer Is Nothing Then
            MsgBox("OwnerPly layer is not loaded")
            Return
        End If

        Dim ssl1, ssl2 As String
        Dim index As Integer = GetColIndex(DataGridExceptions, "INPUTFEATURESSL")
        ssl1 = DataGridExceptions.SelectedRows.Item(0).Cells(index).Value
        index = GetColIndex(DataGridExceptions, "CONFLICTFEATURESSL")
        ssl2 = DataGridExceptions.SelectedRows.Item(0).Cells(index).Value


        Dim pFeature1 As IFeature
        Dim pfeature2 As IFeature


        m_Editor.StartOperation()
        Try
            Dim pFilter As IQueryFilter = New QueryFilter

            Dim pCursor As IFeatureCursor

            If comboActions.SelectedIndex <> 1 Then
                pFilter.WhereClause = "SSL='" & ssl1 & "'"
                pCursor = pLayer.Search(pFilter, False)
                pFeature1 = pCursor.NextFeature

                If pFeature1 Is Nothing Then
                    Throw New Exception("Can't find ownerply with SSL " & ssl1)
                End If

                If comboActions.SelectedIndex = 0 Then
                    pFeature1.Delete()
                End If

            End If

            If comboActions.SelectedIndex <> 0 Then
                pFilter.WhereClause = "SSL='" & ssl2 & "'"
                pCursor = pLayer.Search(pFilter, False)
                pfeature2 = pCursor.NextFeature

                If pfeature2 Is Nothing Then
                    Throw New Exception("Can't find ownerply with SSL " & ssl1)
                End If

                If comboActions.SelectedIndex = 1 Then
                    pfeature2.Delete()
                End If
            End If

            Dim pTopo As ITopologicalOperator
            If comboActions.SelectedIndex = 2 Then
                pTopo = pFeature1.Shape
                pFeature1.Shape = pTopo.Difference(pfeature2.Shape)
                pFeature1.Store()
            End If

            If comboActions.SelectedIndex = 3 Then
                pTopo = pfeature2.Shape
                pfeature2.Shape = pTopo.Difference(pFeature1.Shape)
                pfeature2.Store()
            End If

            Dim pTable As ITable = CType(m_Editor.EditWorkspace, IFeatureWorkspace).OpenTable("PROPERTYCONFLICTLINEAGE")
            Dim oid As Integer = DataGridExceptions.SelectedRows.Item(0).Cells(m_oidIndex).Value
            pFilter.WhereClause = pTable.OIDFieldName & "=" & oid

            Dim pCursor2 As ICursor = pTable.Search(pFilter, False)
            Dim prow As IRow = pCursor2.NextRow
            If Not prow Is Nothing Then
                prow.Value(prow.Fields.FindField("CONFLICTRESOLUTION")) = comboActions.SelectedItem
                prow.Value(prow.Fields.FindField("PROCESSDT")) = Now
                prow.Store()
            End If


            m_Editor.StopOperation("Resolve ownerply confiliction")

        Catch ex As Exception
            m_Editor.AbortOperation()
            MsgBox(ex.Message)
            Return
        End Try

        DataGridExceptions.Rows.RemoveAt(DataGridExceptions.SelectedRows.Item(0).Index)
        Dim pActiveView As IActiveView = m_map
        pActiveView.Refresh()

    End Sub

    Private Sub DataGridExceptions_DataBindingComplete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles DataGridExceptions.DataBindingComplete

        Dim i As Integer
        For i = 0 To DataGridExceptions.ColumnCount - 1
            Dim col As DataGridViewColumn = DataGridExceptions.Columns().Item(i)
            col.ReadOnly = True
            If col.Name = "objectid" Then
                col.Visible = False
                m_oidIndex = i
            End If
        Next

        If Not IAISApplication.GetInstance.IsInEditingSession() Then
            Me.comboActions.Enabled = False
            Me.btnApply.Enabled = False
        Else
            Me.comboActions.Enabled = True
            Me.btnApply.Enabled = True
        End If

    End Sub

    Private Sub btnFlash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlash.Click
        GotoException("flash")
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim dockWindowManager As IDockableWindowManager
        dockWindowManager = CType(m_application, IDockableWindowManager)

        Dim windowID As UID = New UIDClass
        windowID.Value = "{4467D820-6F41-41de-A08F-AA792D29FCB1}"

        Dim dockableWindow As IDockableWindow = dockWindowManager.GetDockableWindow(windowID)
        If Not dockableWindow Is Nothing Then
            dockableWindow.Dock(esriDockFlags.esriDockHide)
        End If

    End Sub
End Class
