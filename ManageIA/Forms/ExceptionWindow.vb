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
Imports ESRI.ArcGIS.Editor

<ComClass(ExceptionWindow.ClassId, ExceptionWindow.InterfaceId, ExceptionWindow.EventsId), _
 ProgId("IAIS.ExceptionWindow")> _
Public Class ExceptionWindow
    Implements IDockableWindowDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "630418E7-267B-432f-8622-810771F3BA4C"
    Public Const InterfaceId As String = "E83047C0-9668-4660-AB9D-053701A08480"
    Public Const EventsId As String = "1486383D-7579-458c-8349-7B2681ADCD70"
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

    Public Const exceptionPremiseExport As Integer = 1
    Public Const exceptionPremiseImport As Integer = 2
    Public Const exceptionMAR As Integer = 3
    Public Const exceptionOwnerPly As Integer = 4
    Public Const exceptionIAChargeFile As Integer = 5
    Public Const exceptionIAAssignPly As Integer = 6


    Public m_bReviewStart As Boolean

    Private m_map As IMap
    Private m_exceptionType As Integer

    Private m_expFeature As IFeature = Nothing

    Private m_oidIndex As Integer
    Private m_uidIndex As Integer

    Private m_viewEvent As IActiveViewEvents_Event

    Private m_Editor As IEditor
    Private m_bInEditingSession As Boolean = False


    Private Sub flashFeature(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase)

        If Not m_expFeature Is Nothing AndAlso phase = esriViewDrawPhase.esriViewForeground Then
            MapUtil.FlashFeature(m_expFeature, m_map)
            m_expFeature = Nothing
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
            Return "Exception Review"
        End Get
    End Property

    Public ReadOnly Property ChildHWND() As Integer Implements ESRI.ArcGIS.Framework.IDockableWindowDef.ChildHWND
        Get
            Return Me.Handle.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Name
        Get
            Return "IAIS_ExceptionWindow"
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
        'The job version does not have the owner as the prefix
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
            Me.chkFixed.Enabled = True

        End If
    End Sub

    Public Sub OnStopEditing()
        m_bInEditingSession = False
    End Sub

    Public Sub OnDestroy() Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnDestroy
        Me.Dispose()
    End Sub

    Public ReadOnly Property UserData() As Object Implements ESRI.ArcGIS.Framework.IDockableWindowDef.UserData
        Get
            Return Me
        End Get
    End Property

    Private Sub btnZoomTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GotoException("zoom")
    End Sub

    Private Sub GotoException(ByVal functionName As String)
        If DataGridExceptions.SelectedRows.Count = 0 Then
            Return
        End If

        Dim uid As Integer = DataGridExceptions.SelectedRows.Item(0).Cells("uniqueid").Value
        Dim uidField As String = ""

        Dim pLayer As IFeatureLayer = Nothing
        Dim pCursor As IFeatureCursor

        Dim pFilter As IQueryFilter = New QueryFilter
        If m_exceptionType = exceptionPremiseExport Or _
            m_exceptionType = exceptionPremiseImport Or _
            m_exceptionType = exceptionIAChargeFile Then
            pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
            If pLayer Is Nothing Then
                MsgBox("Premise Point layer is not loaded")
                Return
            End If
            uidField = "PEXUID"
            pFilter.WhereClause = uidField & "=" & uid
        ElseIf m_exceptionType = exceptionIAAssignPly Then
            pLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_IAASSIGNPLY"), m_map)
            If pLayer Is Nothing Then
                MsgBox("IAAssignPly layer is not loaded")
                Return
            End If
            uidField = "IAID"
            pFilter.WhereClause = "IAID='" & uid & "'"
        ElseIf m_exceptionType = exceptionMAR Then
            Dim addressLayerName As String
            If DataGridExceptions.SelectedRows.Item(0).Cells("REFDATASRC").Value = "O" Then
                addressLayerName = IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT") & "Old"
            Else
                addressLayerName = IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT")
            End If
            pLayer = MapUtil.GetLayerByTableName(addressLayerName, m_map)

            If pLayer Is Nothing Then
                MsgBox("Address Point layer " & addressLayerName & "is not loaded")
                Return
            End If
            uidField = "ADDRESS_ID"
            pFilter.WhereClause = "ADDRESS_ID=" & uid
        ElseIf m_exceptionType = exceptionOwnerPly Then
            Dim ssl As String = DataGridExceptions.SelectedRows.Item(0).Cells("SSL").Value
            pLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
            If pLayer Is Nothing Then
                MsgBox("OwnerPly layer is not loaded")
                Return
            End If
            uidField = "SSL"
            pFilter.WhereClause = "SSL='" & ssl & "'"
        End If

        'pFilter.WhereClause = uidField & "=" & uid
        pCursor = pLayer.Search(pFilter, False)

        m_expFeature = pCursor.NextFeature
        If m_expFeature Is Nothing Then
            MsgBox("Can't find feature [" & uidField & "  " & uid & "] in layer " & pLayer.Name)
            Return
        End If


        If Not m_expFeature Is Nothing Then
            If functionName = "zoom" Then
                MapUtil.ZoomToFeature(m_expFeature, m_map)
            Else
                MapUtil.PanToFeature(m_expFeature, m_map)
            End If

            MapUtil.SelectFeature(m_expFeature, pLayer)
        End If

    End Sub

    Private Sub DataGridExceptions_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExceptions.CellClick
        m_bReviewStart = True
        ReviewException()
    End Sub

    Private Sub DataGridExceptions_DataBindingComplete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles DataGridExceptions.DataBindingComplete
        'Change the caption of the columns
        'Datagridview does not honor the column name of the table

        Dim dt As DataTable = DirectCast(DirectCast(DataGridExceptions.DataSource, BindingSource).DataSource, DataTable)

        m_oidIndex = 0
        m_uidIndex = 0
        Dim i As Integer
        For i = 0 To DataGridExceptions.ColumnCount - 1
            Dim col As DataGridViewColumn = DataGridExceptions.Columns().Item(i)
            col.ReadOnly = True
            If col.Name = "objectid" Then
                m_oidIndex = i
                col.Visible = False
            ElseIf col.Name = "uniqueid" Then
                col.Visible = False
                m_uidIndex = i
            End If
            col.HeaderText = dt.Columns(i).Caption
        Next

    End Sub

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

    Private Sub ReviewException()

        If DataGridExceptions.SelectedRows.Count = 0 Then
            chkFixed.Enabled = False
            chkFixed.Checked = False

            Return
        Else

            If rBtnZoom.Checked Then
                GotoException("zoom")
            Else
                GotoException("pan")
            End If

            If Not m_bInEditingSession Then
                Return
            End If

            chkFixed.Enabled = True
            chkFixed.Checked = False

            If m_exceptionType = exceptionIAChargeFile Then
                Return
            End If

            Dim oid As Integer = DataGridExceptions.SelectedRows.Item(0).Cells(m_oidIndex).Value

            Dim jtxExt As IJTXExtension = DirectCast(m_application.FindExtensionByName("Workflow Manager"), IJTXExtension)
            Dim pFWS As IFeatureWorkspace = DirectCast(m_Editor.EditWorkspace, IFeatureWorkspace) 'jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)
            Dim pWE As IWorkspaceEdit = pFWS

            m_Editor.StartOperation()

            Try
                Dim pTable As ITable = Nothing

                Select Case m_exceptionType
                    Case ExceptionWindow.exceptionPremiseExport
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSEXPOUT"))
                    Case ExceptionWindow.exceptionPremiseImport
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSINTERPT"))
                    Case ExceptionWindow.exceptionMAR
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSADDRESSPT"))
                    Case ExceptionWindow.exceptionOwnerPly
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSOWNERPLY"))
                    Case ExceptionWindow.exceptionIAAssignPly
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSIAPLY"))
                    Case ExceptionWindow.exceptionIAChargeFile
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSBILLDET"))
                End Select

                Dim pFilter As IQueryFilter = New QueryFilter
                pFilter.WhereClause = pTable.OIDFieldName & "=" & oid
                Dim pRow As IRow
                Dim pCursor As ICursor = pTable.Search(pFilter, False)
                pRow = pCursor.NextRow
                If Not pRow Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("REVDT")) = Now
                    pRow.Value(pRow.Fields.FindField("REVIEWBY")) = Environment.UserName

                    Dim i As Integer
                    i = GetColIndex(DataGridExceptions, "REVDT")
                    DataGridExceptions.SelectedRows.Item(0).Cells(i).Value = Now
                    i = GetColIndex(DataGridExceptions, "REVIEWBY")
                    DataGridExceptions.SelectedRows.Item(0).Cells(i).Value = Environment.UserName

                    pRow.Store()

                End If

                m_Editor.StopOperation("")

            Catch ex As Exception
                m_Editor.AbortOperation()
                MsgBox(ex.Message & vbCrLf & ex.StackTrace.ToString())
                Return
            End Try


        End If

    End Sub

    Private Sub DataGridExceptions_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridExceptions.SelectionChanged

        If DataGridExceptions.SelectedRows.Count = 0 Then
            chkFixed.Enabled = False
            chkFixed.Checked = False
        End If

        If Not m_bReviewStart Then
            Return
        End If

        If Not Me.Visible Then
            Return
        End If

        ReviewException()

    End Sub

    Private Sub chkFixed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFixed.CheckedChanged
        If chkFixed.Checked Then
            'delete the record from list and table

            Dim oid As Integer = DataGridExceptions.SelectedRows.Item(0).Cells(m_oidIndex).Value

            Dim jtxExt As IJTXExtension = DirectCast(m_application.FindExtensionByName("Workflow Manager"), IJTXExtension)
            Dim pFWS As IFeatureWorkspace = DirectCast(m_Editor.EditWorkspace, IFeatureWorkspace)  'jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)
            Dim pWE As IWorkspaceEdit = pFWS

            m_Editor.StartOperation()
            Try
                Dim pTable As ITable = Nothing

                Select Case m_exceptionType
                    Case ExceptionWindow.exceptionPremiseExport
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSEXPOUT"))
                    Case ExceptionWindow.exceptionPremiseImport
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSINTERPT"))
                    Case ExceptionWindow.exceptionMAR
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSADDRESSPT"))
                    Case ExceptionWindow.exceptionOwnerPly
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSOWNERPLY"))
                    Case ExceptionWindow.exceptionIAChargeFile
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSBILLDET"))
                    Case ExceptionWindow.exceptionIAAssignPly
                        pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSIAPLY"))
                End Select

                Dim pFilter As IQueryFilter = New QueryFilter
                pFilter.WhereClause = pTable.OIDFieldName & "=" & oid
                Dim pRow As IRow
                Dim pCursor As ICursor = pTable.Search(pFilter, False)
                pRow = pCursor.NextRow
                If Not pRow Is Nothing Then
                    'If m_exceptionType = exceptionIAChargeFile Then
                    '    Try
                    '        pRow.Delete()
                    '    Catch ex As Exception
                    '        MsgBox("Delete Exception Row:" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace.ToString())
                    '    End Try
                    'Else
                    Try
                        pRow.Value(pRow.Fields.FindField("FIXDT")) = Now
                        pRow.Value(pRow.Fields.FindField("FIXBY")) = Environment.UserName
                        pRow.Store()
                    Catch ex As Exception
                        MsgBox("Store table " & DirectCast(pTable, IDataset).BrowseName & ":" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace.ToString())
                    End Try

                    If m_exceptionType = ExceptionWindow.exceptionPremiseExport Then
                        'Check to see if all the exceptions are fixed
                        Dim transCode As String = pRow.Value(pRow.Fields.FindField("TRANSACTIONCODE"))
                        Dim pexuid As String = pRow.Value(pRow.Fields.FindField("PEXUID"))

                        Marshal.ReleaseComObject(pCursor)

                        pFilter.WhereClause = "PEXUID=" & pexuid & " AND TRANSACTIONCODE='" & transCode & "' AND FIXDT IS NULL"
                        pCursor = pTable.Search(pFilter, False)
                        pRow = pCursor.NextRow
                        If pRow Is Nothing Then
                            pFilter.WhereClause = "PEXUID=" & pexuid & " AND TRANSACTIONCODE='" & transCode & "' AND NOT EXCEPTIONMSGID IS NULL"
                            Marshal.ReleaseComObject(pCursor)

                            pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_CISOUTPUTPREMEXP"))
                            pCursor = pTable.Search(pFilter, False)
                            pRow = pCursor.NextRow
                            If Not pRow Is Nothing Then
                                Try
                                    pRow.Value(pRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value
                                    pRow.Store()
                                Catch ex As Exception
                                    MsgBox("Store TABLE_CISOUTPUTPREMEXP:" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace.ToString())
                                End Try
                            End If
                        End If
                    ElseIf m_exceptionType = ExceptionWindow.exceptionIAChargeFile Then
                        'Check to see if all the exceptions are fixed
                        Dim pexuid As String = pRow.Value(pRow.Fields.FindField("PEXUID"))

                        Marshal.ReleaseComObject(pCursor)

                        pFilter.WhereClause = "PEXUID=" & pexuid & " AND FIXDT IS NULL"
                        pCursor = pTable.Search(pFilter, False)
                        pRow = pCursor.NextRow
                        If pRow Is Nothing Then
                            pFilter.WhereClause = "PEXUID=" & pexuid & " AND NOT EXCEPTIONMSGID IS NULL"
                            Marshal.ReleaseComObject(pCursor)

                            pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_CISOUTPUTBILLDET"))
                            pCursor = pTable.Search(pFilter, False)
                            pRow = pCursor.NextRow
                            If Not pRow Is Nothing Then
                                Try
                                    pRow.Value(pRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value
                                    pRow.Store()
                                Catch ex As Exception
                                    MsgBox("Store TABLE_CISOUTPUTBILLDET:" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace.ToString())
                                End Try
                            End If
                        End If
                    End If

                    'End If
                Else
                    MsgBox("Can't find row in database " & vbCrLf & pFilter.WhereClause)
                End If

                m_Editor.StopOperation("Exception review")

                DataGridExceptions.Rows.RemoveAt(DataGridExceptions.SelectedRows.Item(0).Index)
                DataGridExceptions.CurrentCell = Nothing
                chkFixed.Enabled = False

            Catch ex As Exception
                m_Editor.AbortOperation()
                MsgBox(ex.Message & vbCrLf & ex.StackTrace.ToString())
                Return
            End Try

        End If
    End Sub



    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GotoException("pan")
    End Sub



    Private Sub BindingNavigatorMoveNextItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorMoveNextItem.Click

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim dockWindowManager As IDockableWindowManager
        dockWindowManager = CType(m_application, IDockableWindowManager)

        Dim windowID As UID = New UIDClass
        windowID.Value = "{630418E7-267B-432f-8622-810771F3BA4C}"

        Dim dockableWindow As IDockableWindow = dockWindowManager.GetDockableWindow(windowID)
        If Not dockableWindow Is Nothing Then
            dockableWindow.Dock(esriDockFlags.esriDockHide)
        End If

    End Sub


    Private Sub ExceptionWindow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not m_map Is Nothing Then
            m_viewEvent = m_map
            AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature
        End If
    End Sub
End Class
