Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display


Public Class IAISApplication : Implements IJTXJobListener

    Private m_application As IApplication
    Private m_editorEvent As IEditEvents_Event
    Private m_editorEvent2 As IEditEvents2_Event
    Private m_pEditor As IEditor

    Private m_docEvents As IDocumentEvents_Event
    Private m_viewEvent As IActiveViewEvents_Event

    Private Shared IAISApp As IAISApplication = Nothing

    Private m_inToolEditingSession As Boolean = False
    Private m_invalidEditing As Boolean = False
    Private m_deleteFeature As Boolean = False

    Private m_editingErrMsg As String

    Private m_flashFeature As IGeometry



    Public Shared Function GetInstance() As IAISApplication
        If IAISApp Is Nothing Then
            IAISApp = New IAISApplication
        End If

        Return IAISApp
    End Function

    Public Sub StartToolEditing()
        m_pEditor.StartOperation()
        m_inToolEditingSession = True
    End Sub

    Public Sub StopToolEditing(Optional ByVal opName As String = "")
        m_pEditor.StopOperation(opName)
        m_inToolEditingSession = False
    End Sub

    Public Sub AbortToolEditing()
        m_pEditor.AbortOperation()
        m_inToolEditingSession = False
    End Sub


    Public Property ArcMapApplicaiton() As IApplication
        Get
            Return m_application
        End Get
        Set(ByVal value As IApplication)
            If m_application Is Nothing Then
                m_application = value


                Dim pUid As New UID
                pUid.Value = "esriEditor.Editor"
                m_pEditor = m_application.FindExtensionByCLSID(pUid)

                If m_editorEvent Is Nothing AndAlso Not m_pEditor Is Nothing Then
                    m_editorEvent = CType(m_pEditor, IEditEvents_Event)
                    AddHandler m_editorEvent.OnCreateFeature, AddressOf OnEditingCheck
                    AddHandler m_editorEvent.OnChangeFeature, AddressOf OnEditingCheck
                    AddHandler m_editorEvent.OnDeleteFeature, AddressOf OnDeleteFeature


                    m_editorEvent2 = CType(m_pEditor, IEditEvents2_Event)
                    AddHandler m_editorEvent2.BeforeStopOperation, AddressOf OnBeforeStopOperation
                    AddHandler m_editorEvent2.BeforeStopEditing, AddressOf OnBeforeStopEditing
                    AddHandler m_editorEvent2.OnStopOperation, AddressOf OnStopOperation

                End If

				Dim jtxExt As ESRI.ArcGIS.JTXExt.JTXExtension = DirectCast(m_application.FindExtensionByName("Workflow Manager"), ESRI.ArcGIS.JTXExt.JTXExtension)
				If Not jtxExt Is Nothing Then
                    jtxExt.AttachListener(Me)

                End If


                m_docEvents = CType(m_application.Document, IDocumentEvents_Event)

                AddHandler m_docEvents.OpenDocument, AddressOf OnOpenDocument
                AddHandler m_docEvents.NewDocument, AddressOf OnNewDocument
                AddHandler m_docEvents.ActiveViewChanged, AddressOf OnActiveViewChanged
                AddHandler m_docEvents.MapsChanged, AddressOf OnMapsChanged

                Dim currentDomain As AppDomain = AppDomain.CurrentDomain
                AddHandler System.Windows.Forms.Application.ThreadException, AddressOf onThreadException
                AddHandler currentDomain.UnhandledException, AddressOf GlobalExpHandle

            End If
        End Set
    End Property

    Private Sub OnOpenDocument()
        'Log("OnOpenDocument")
        OnActiveViewChanged()
    End Sub
    Private Sub OnNewDocument()
        'Log("OnNewDocument")
        OnActiveViewChanged()
    End Sub
    Private Sub OnMapsChanged()
        'Log("OnMapsChanged")
        OnActiveViewChanged()
    End Sub

    Private Sub OnJobChanged()
        'Log("OnJobChanged")
        InitilizeIAISToolSetting()
    End Sub

    Private Sub OnItemAdded()
        'Log("OnItemAdded")
        InitilizeIAISToolSetting()
    End Sub

    Private Sub OnFocusMapChanged()
        'Log("OnFocusMapChanged")
        InitilizeIAISToolSetting()
    End Sub


    Private Sub OnActiveViewChanged()
        Try
            InitilizeIAISToolSetting()
            Dim pMxDoc As IMxDocument = m_application.Document
            AddHandler CType(pMxDoc.FocusMap, IActiveViewEvents_Event).ItemAdded, AddressOf OnItemAdded
            AddHandler CType(pMxDoc.FocusMap, IActiveViewEvents_Event).FocusMapChanged, AddressOf OnFocusMapChanged
            AddHandler CType(pMxDoc.FocusMap, IActiveViewEvents_Event).AfterDraw, AddressOf onAfterDraw
            AddHandler CType(pMxDoc.FocusMap, IActiveViewEvents_Event).SelectionChanged, AddressOf OnFocusMapChanged
            AddHandler CType(pMxDoc.FocusMap, IActiveViewEvents_Event).ViewRefreshed, AddressOf OnFocusMapChanged

        Catch ex As Exception
            Log(ex.Message)
            Log(ex.StackTrace())
        End Try
    End Sub

    Private Sub onAfterDraw(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase)
        If Not m_flashFeature Is Nothing AndAlso phase = esriViewDrawPhase.esriViewForeground Then
            Dim pMxDoc As IMxDocument = m_application.Document
            MapUtil.FlashGeometry(m_flashFeature, pMxDoc.FocusMap)
            m_flashFeature = Nothing
        End If
    End Sub

    Private Sub InitilizeIAISToolSetting()
        Try
            If Not m_application Is Nothing Then
                IAISToolSetting.Initilize(m_application)
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString(), , "Unhandled Exception")
        End Try

    End Sub

    Private Sub OnDeleteFeature(ByVal obj As IObject)
        If TypeOf obj Is IFeature Then
            Dim pFeature As IFeature = CType(obj, IFeature)
            Dim pObjClass As IObjectClass = obj.Class
            Dim pFeatureClass As IFeatureClass = CType(pObjClass, IFeatureClass)
            Dim pDataset As IDataset = pFeatureClass

            If UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT")) Then
                If Not CanDeletePremise(pFeature) Then
                    m_invalidEditing = True
                    m_deleteFeature = True
                Else
                    m_invalidEditing = False
                    m_deleteFeature = True
                End If
            End If
        End If
    End Sub

    Private Function CanDeletePremise(ByVal pFeature As IFeature) As Boolean
        'TABLE_CISOUTPUTPREMEXP
        If MapUtil.GetValue(pFeature, "PEXPRM") <> "" Then
            m_editingErrMsg = "Deleting premise with premise number is not allowed."
            Return False
        End If

        Dim pexuid As String = MapUtil.GetValue(pFeature, "PEXUID")
        Dim pMxDoc As IMxDocument = m_application.Document

        Dim pWS As IFeatureWorkspace = MapUtil.GetPremiseWS(pMxDoc.FocusMap)
        Dim pQueryDef As IQueryDef = pWS.CreateQueryDef
        pQueryDef.Tables = IAISToolSetting.GetParameterValue("TABLE_CISOUTPUTPREMEXP")
        pQueryDef.SubFields = "COUNT(*)"
        pQueryDef.WhereClause = "PEXUID=" & pexuid

        Dim pCursor As ICursor = pQueryDef.Evaluate
        Dim pRow As IRow = pCursor.NextRow

        If pRow.Value(0) > 0 Then
            m_editingErrMsg = "There is a pending transaction for this premise. Can't delete."
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub OnEditingCheck(ByVal obj As IObject)

        Dim registryKey As Microsoft.Win32.RegistryKey
        registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\IAIS")
        Dim tnIAISSetting As String = registryKey.GetValue("ALLOW_PREMISE_EDIT")

        If Not tnIAISSetting Is Nothing Then
            If "Y".Equals(tnIAISSetting) Then
                Return
            End If
        End If

        If TypeOf obj Is IFeature Then
            Dim pFeature As IFeature = CType(obj, IFeature)
            Dim pObjClass As IObjectClass = obj.Class
            Dim pFeatureClass As IFeatureClass = CType(pObjClass, IFeatureClass)
            Dim pDataset As IDataset = pFeatureClass

            If UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT")) Then
                If Not m_inToolEditingSession Then
                    m_invalidEditing = True
                    Return
                End If
            End If
        End If
    End Sub

    Private Sub OnBeforeStopOperation()
        'If m_invalidEditing Then
        '    Dim pMxDocument As IMxDocument
        '    pMxDocument = m_application.Document
        '    Dim pOperationStack As IOperationStack
        '    pOperationStack = pMxDocument.OperationStack
        '    Dim pLastOperation As IOperation = pOperationStack.Item(pOperationStack.Count - 1)
        '    If Not pLastOperation.MenuString = "Move" Then
        '        MsgBox("Editing to premise point layer without tool is not allowed.")
        '        m_pEditor.AbortOperation()
        '    End If
        '    m_invalidEditing = False
        'End If
    End Sub

    Private Sub OnBeforeStopEditing()
        IAISToolSetting.CloseActiveForm()
    End Sub

    Private Sub OnStopOperation()
        Dim pMxDocument As IMxDocument
        pMxDocument = m_application.Document

        Dim pOperationStack As IOperationStack
        pOperationStack = pMxDocument.OperationStack

        Dim pLastOperation As IOperation = pOperationStack.Item(pOperationStack.Count - 1)
        If pLastOperation.MenuString = "Move" Then
            Try
                If MsgBox("You just moved a feature. Are you sure you want to move it?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    pOperationStack.Undo()
                    pOperationStack.Remove(pOperationStack.Count - 1)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf m_invalidEditing Then
            If m_deleteFeature Then
                MsgBox(m_editingErrMsg)
            Else
                MsgBox("Editing to premise point layer without tool is not allowed.")
            End If

            pOperationStack.Undo()
            pOperationStack.Remove(pOperationStack.Count - 1)
        End If

        m_invalidEditing = False
        m_deleteFeature = False

    End Sub

    Private Sub onThreadException(ByVal sender As Object,
       ByVal e As System.Threading.ThreadExceptionEventArgs)
        MsgBox("Unhandled Exception: " & e.Exception.Message & vbCrLf & vbCrLf & e.Exception.StackTrace.ToString(), Title:="Unhandled Exception")
    End Sub

    Private Sub GlobalExpHandle(ByVal sender As Object, ByVal args As UnhandledExceptionEventArgs)
        Dim e As Exception = DirectCast(args.ExceptionObject, Exception)
        MsgBox("Unhandled Exception: " & e.Message & vbCrLf & vbCrLf & e.StackTrace.ToString(), Title:="Unhandled Exception")
    End Sub

    Public Function IsInEditingSession() As Boolean
        If m_pEditor Is Nothing Then
            Return False
        End If
        If m_pEditor.EditState = esriEditState.esriStateNotEditing Then
            Return False
        End If
        Return True
    End Function

    Public Function IsInJTXApplication() As Boolean

        If UCase(IAISToolSetting.GetParameterValue("EDITING_WITHOUT_JTX")) = "TRUE" Then
            Return True
        End If

        Dim jtxExt As ESRI.ArcGIS.JTXExt.JTXExtension = Nothing
        Try
			jtxExt = DirectCast(m_application.FindExtensionByName("Workflow Manager"), ESRI.ArcGIS.JTXExt.JTXExtension)
		Catch ex As Exception

        End Try

        If jtxExt Is Nothing Then
            Return False
        End If

        If jtxExt.Job Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub FlashFeature(ByVal pGeom As IGeometry)
        m_flashFeature = pGeom
    End Sub

    Public Sub Log(ByVal msg As String)
        Dim logFilePath As String = Environment.GetEnvironmentVariable("TEMP")
        If logFilePath Is Nothing Then
            logFilePath = Environment.GetEnvironmentVariable("TMP")
        End If
        If logFilePath Is Nothing Then
            logFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())
        End If

        Dim logFile As String = logFilePath + "\IAISApplication.log"

        Try
            Dim logStream As System.IO.StreamWriter = System.IO.File.AppendText(logFile)
            Try
                logStream.WriteLine(Now & " : " & msg)
            Catch ex As Exception
            Finally
                logStream.Close()
            End Try
        Catch ex As Exception

        End Try

    End Sub

    Public Sub JobChanged(pNewJob As IJTXJob) Implements IJTXJobListener.JobChanged
        InitilizeIAISToolSetting()
        'MsgBox("Job Changed")
    End Sub
End Class
