Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Carto

<ComClass(IASearchReportExt.ClassId, IASearchReportExt.InterfaceId, IASearchReportExt.EventsId)>
<ProgId("IAIS.IASearchReportExt")>
Public Class IASearchReportExt
    Implements IExtension
    Implements IExtensionConfig

    Public Event IASearchReportExtEvent(ByVal strMsg As String)

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "9781f06d-e8b4-4fc1-b6f2-85b39bd87ea5"
    Public Const InterfaceId As String = "492902dc-786a-4fd2-b116-f71aa5e08f0c"
    Public Const EventsId As String = "9f01706c-7778-41b1-b7d5-831757b66626"
#End Region

#Region "COM Registration Functions"
    ' Register the Extension in the ESRI MxExtensions Component Category 
    <ComRegisterFunction()>
    <ComVisible(False)>
    Private Shared Sub RegisterFunction(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Register(regKey)
    End Sub

    <ComUnregisterFunction()>
    <ComVisible(False)>
    Private Shared Sub UnregisterFunction(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtension.Unregister(regKey)
    End Sub
#End Region

    Private m_extensionState As esriExtensionState = esriExtensionState.esriESDisabled
    Private m_application As IApplication

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
            Return "IAIS Query Tools"
        End Get
    End Property

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_application = TryCast(initializationData, IApplication)
        IAISApplication.GetInstance().ArcMapApplicaiton = m_application
    End Sub

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        m_application = Nothing
        m_docEvents = Nothing
    End Sub


    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "Extension for IAIS Query Tools"
        End Get
    End Property

    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "IAIS Query Tools"
        End Get
    End Property

    Public Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
        Get
            Return m_extensionState
        End Get
        Set(ByVal value As ESRI.ArcGIS.esriSystem.esriExtensionState)
            m_extensionState = value

            If m_application Is Nothing Then
                m_extensionState = esriExtensionState.esriESUnavailable
                Return
            End If

            Dim pCommandBar As ICommandBar
            Dim uid As New UID

            If m_extensionState = esriExtensionState.esriESDisabled Or m_extensionState = esriExtensionState.esriESUnavailable Then

                uid.Value = "IAIS.IASearchReportExt"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                RaiseEvent IASearchReportExtEvent("disable")
            ElseIf m_extensionState = esriExtensionState.esriESEnabled Then

                uid.Value = "IAIS.IASearchReportExt"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                RaiseEvent IASearchReportExtEvent("enable")

            End If

        End Set
    End Property
End Class


