Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms

<ComClass(PremiseExt.ClassId, PremiseExt.InterfaceId, PremiseExt.EventsId)> _
<ProgId("IAIS.PremiseExt")> _
Public Class PremiseExt
    Implements IExtension
    Implements IExtensionConfig

    Public Event PremiseExtEvent(ByVal strMsg As String)

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d4aa9917-e7a3-4414-a4e8-22c18061a477"
    Public Const InterfaceId As String = "1a55d740-4db2-4756-84dc-01b20f9bf72e"
    Public Const EventsId As String = "93067fc2-4474-4175-af83-91db9ef2182c"
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

    Private m_docEvents As IDocumentEvents_Event

    Private m_editorEvent As IEditEvents_Event
    Private m_pEditor As IEditor


    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "Extension for IAIS Manage Premise Tools "
        End Get
    End Property

    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "IAIS Manage Premise Tools"
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

                uid.Value = "IAIS.PremiseManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing Then
                    pCommandBar.Dock(esriDockFlags.esriDockHide)
                End If

                Dim pForm As Form = IAISToolSetting.GetActiveForm
                If Not pForm Is Nothing Then
                    If TypeOf pForm Is PremiseAttribute Or _
                        TypeOf pForm Is FormSelectPremise Then
                        IAISToolSetting.CloseActiveForm()
                    End If
                End If

                RaiseEvent PremiseExtEvent("disable")

            ElseIf m_extensionState = esriExtensionState.esriESEnabled Then

                uid.Value = "IAIS.PremiseManageToolbar"
                pCommandBar = m_application.Document.CommandBars.Find(uid)
                If Not pCommandBar Is Nothing And Not pCommandBar.IsVisible Then
                    pCommandBar.Dock(esriDockFlags.esriDockShow)
                End If

                RaiseEvent PremiseExtEvent("enable")

            End If

        End Set
    End Property


    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "IAIS Manage Premise Tools"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        m_application = Nothing
        m_docEvents = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_application = TryCast(initializationData, IApplication)
        IAISApplication.GetInstance().ArcMapApplicaiton = m_application
    End Sub

End Class


