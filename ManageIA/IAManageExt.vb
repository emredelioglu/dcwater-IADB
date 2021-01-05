Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Carto

<ComClass(IAManageExt.ClassId, IAManageExt.InterfaceId, IAManageExt.EventsId)> _
<ProgId("IAIS.IAManageExt")> _
Public Class IAManageExt
    Implements IExtension
    Implements IExtensionConfig

    Public Event IAManageExtEvent(ByVal strMsg As String)

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ca1ae6d2-9e05-404c-84fe-2032faf5ac7f"
    Public Const InterfaceId As String = "f40cf906-11dd-48af-9702-9407011e9934"
    Public Const EventsId As String = "8144b656-23c3-4ebf-9e2c-dfc306c24606"
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

    Private m_extensionState As esriExtensionState = esriExtensionState.esriESDisabled
    Private m_application As IApplication

    Private m_editorEvent As IEditEvents2_Event
    Private m_pEditor As IEditor
    Private m_docEvents As IDocumentEvents_Event
    Private m_activeViewEvents As IActiveViewEvents_Event

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "IAIS Manage IA Tools"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        m_application = Nothing
        m_docEvents = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_application = TryCast(initializationData, IApplication)
        IAISApplication.GetInstance().ArcMapApplicaiton = m_application

        Dim pMxDoc As IMxDocument = m_application.Document
        m_docEvents = CType(pMxDoc, IDocumentEvents_Event)
        AddHandler m_docEvents.OpenDocument, AddressOf OnOpenDocument

        IAISToolSetting.SetApplication(m_application)
    End Sub

    Private Sub OnOpenDocument()
        IAISToolSetting.Initilize(m_application)

        Dim dockWindowManager As IDockableWindowManager
        dockWindowManager = CType(m_application, IDockableWindowManager)
        If Not dockWindowManager Is Nothing Then
            Dim windowID As UID = New UIDClass
            windowID.Value = "{4467D820-6F41-41de-A08F-AA792D29FCB1}"
            Dim dockableWindow As IDockableWindow = dockWindowManager.GetDockableWindow(windowID)
            If Not dockableWindow Is Nothing Then
                dockableWindow.Dock(esriDockFlags.esriDockHide)
            End If

            windowID.Value = "{630418E7-267B-432f-8622-810771F3BA4C}"
            dockableWindow = dockWindowManager.GetDockableWindow(windowID)
            If Not dockableWindow Is Nothing Then
                dockableWindow.Dock(esriDockFlags.esriDockHide)
            End If
        End If

    End Sub

    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "Extension for IAIS Manage IA Tools"
        End Get
    End Property

    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "IAIS Manage IA Tools"
        End Get
    End Property

    Public Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
        Get
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

                uid.Value = "IAIS.ExceptionToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                uid.Value = "IAIS.IAManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                RaiseEvent IAManageExtEvent("disable")

            ElseIf m_extensionState = esriExtensionState.esriESEnabled Then

                uid.Value = "IAIS.ExceptionToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                uid.Value = "IAIS.IAManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                RaiseEvent IAManageExtEvent("enable")

            End If
        End Set
    End Property



End Class


