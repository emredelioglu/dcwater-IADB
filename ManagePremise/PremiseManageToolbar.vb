Imports ESRI.ArcGIS.ADF.CATIDs
Imports System.Runtime.InteropServices

<ComClass(PremiseManageToolbar.ClassId, PremiseManageToolbar.InterfaceId, PremiseManageToolbar.EventsId), _
ProgId("IAIS.PremiseManageToolbar")> _
Public Class PremiseManageToolbar
    Inherits ESRI.ArcGIS.ADF.BaseClasses.BaseToolbar

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


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a2cab203-f81b-4813-8883-8c5990c8a8a9"
    Public Const InterfaceId As String = "575bc767-6d7b-4f61-ac1e-48d2645e84ec"
    Public Const EventsId As String = "1dd11c24-b3bd-4679-8ea1-a0f759f96852"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        AddItem("IAIS.PremiseTool", 1)
        AddItem("IAIS.PremiseTool", 2)
        'AddItem("IAIS.PremiseUpdate")
        AddItem("IAIS.PremiseSelMaster")
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "IAIS Manage Premise"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "IAIS Manage Premise"
        End Get
    End Property


End Class


