Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Carto

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


    Private m_map As IMap
    Private m_exceptionType As Integer

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
    End Sub

    Public Sub OnDestroy() Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnDestroy
        Me.Dispose()
    End Sub

    Public ReadOnly Property UserData() As Object Implements ESRI.ArcGIS.Framework.IDockableWindowDef.UserData
        Get
            Return Me
        End Get
    End Property
End Class
