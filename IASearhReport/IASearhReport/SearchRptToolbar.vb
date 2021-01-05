Imports ESRI.ArcGIS.ADF.CATIDs
Imports System.Runtime.InteropServices

<ComClass(SearchRptToolbar.ClassId, SearchRptToolbar.InterfaceId, SearchRptToolbar.EventsId)> _
<ProgId("IAIS.SearchRptToolbar")> _
Public Class SearchRptToolbar
    Inherits ESRI.ArcGIS.ADF.BaseClasses.BaseToolbar

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e05a9785-8419-436c-962b-b654f90c1b0c"
    Public Const InterfaceId As String = "0cf9ed6c-0fd1-4eff-b7fc-5361f274c0bd"
    Public Const EventsId As String = "a960756f-9dcd-480d-9dcc-a707fa2dabcd"
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
        MxCommandBars.Register(regKey)
    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)
    End Sub

#End Region
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        AddItem("IAIS.PropertyLocator")
        AddItem("IAIS.PremiseLocator")
        AddItem("IAIS.IAChargeHistory")
        AddItem("IAIS.BillDeterminantReport")
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "IAIS Search/Report Toolbar"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "IAIS Search/Report Toolbar"
        End Get
    End Property

End Class


