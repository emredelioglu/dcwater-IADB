Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Carto

<ComClass(IAIAExt.ClassId, IAIAExt.InterfaceId, IAIAExt.EventsId)> _
<ProgId("IAIS.Extension")> _
Public Class IAIAExt
    Implements IExtension
    Implements IExtensionConfig
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4cab507d-c95b-40fb-903c-7b85bd30390e"
    Public Const InterfaceId As String = "1ce327d1-7506-42cd-ba40-79a820d96d25"
    Public Const EventsId As String = "3ca267f1-fab2-4ceb-af00-0830278b33e2"
#End Region

#Region "COM Registration Functions"
    ' Register the Extension in the ESRI MxExtensions Component Category 
    <ComRegisterFunction()> _
    <ComVisible(False)> _
    Private Shared Sub RegisterFunction(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    <ComVisible(False)> _
    Private Shared Sub UnregisterFunction(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Unregister(regKey)
    End Sub
#End Region


    Private m_application As IApplication
    Private m_isMenuPresent As Boolean = False

    Private m_extensionState As esriExtensionState = esriExtensionState.esriESDisabled

    Private m_editorEvent As IEditEvents2_Event
    Private m_pEditor As IEditor




    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "IAIS Extension"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        m_application = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_application = TryCast(initializationData, IApplication)
    End Sub

    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "Extension for IAIS tools"
        End Get
    End Property

    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "IAIS Extension"
        End Get
    End Property

    Public Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
        Get
            Dim pDoc As IMxDocument
            pDoc = m_application.Document
            Dim pLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", pDoc.FocusMap)

            If pLayer Is Nothing Then
                m_extensionState = esriExtensionState.esriESDisabled
            End If

            Return m_extensionState
        End Get
        Set(ByVal value As ESRI.ArcGIS.esriSystem.esriExtensionState)
            'If value = m_extensionState Then
            '    Return
            'End If

            m_extensionState = value

            If m_application Is Nothing Then
                m_extensionState = esriExtensionState.esriESUnavailable
                Return
            End If

            Dim pCommandBar As ICommandBar
            Dim uid As New UID

            If m_extensionState = esriExtensionState.esriESDisabled Or m_extensionState = esriExtensionState.esriESUnavailable Then

                If Not m_editorEvent Is Nothing Then
                    RemoveHandler m_editorEvent.OnStopOperation, AddressOf OnStopOperation
                    m_editorEvent = Nothing
                End If

                uid.Value = "IAIS.ExceptionToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                uid.Value = "IAIS.SearchRptToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                uid.Value = "IAIS.IAManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                uid.Value = "IAIS.PremiseManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If
            ElseIf m_extensionState = esriExtensionState.esriESEnabled Then

                Dim pUid As New UID
                pUid.Value = "esriEditor.Editor"
                m_pEditor = m_application.FindExtensionByCLSID(pUid)

                If Not IAISToolSetting.Initilize(m_application) Then
                    m_extensionState = esriExtensionState.esriESDisabled
                    Return
                End If

                If m_editorEvent Is Nothing AndAlso Not m_pEditor Is Nothing Then
                    m_editorEvent = CType(m_pEditor, IEditEvents_Event)
                    AddHandler m_editorEvent.OnStopOperation, AddressOf OnStopOperation
                End If

                uid.Value = "IAIS.ExceptionToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                uid.Value = "IAIS.SearchRptToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                uid.Value = "IAIS.IAManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                uid.Value = "IAIS.PremiseManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

            End If
        End Set
    End Property


    Private Sub OnStopOperation()
        Dim pMxDocument As IMxDocument
        pMxDocument = m_application.Document

        Dim pOperationStack As IOperationStack
        pOperationStack = pMxDocument.OperationStack

        Dim pLastOperation As IOperation = pOperationStack.Item(pOperationStack.Count - 1)
        If pLastOperation.MenuString = "Move" Then
            Try

                If MsgBox("You just moved a feature. Are you sure you want to move it?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    'm_pEditor.UndoOperation()
                    pOperationStack.Undo()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub


End Class


