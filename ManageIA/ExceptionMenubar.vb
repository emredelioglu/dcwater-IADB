Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework

<ComClass(ExceptionMenubar.ClassId, ExceptionMenubar.InterfaceId, ExceptionMenubar.EventsId)> _
<ProgId("IAIS.ExceptionMenubar")> _
Public Class ExceptionMenubar
    Inherits BaseMenu
    Implements IRootLevelMenu

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "333d4f55-de6c-4039-9aaf-03ce7a685149"
    Public Const InterfaceId As String = "0cc57f40-5f3a-4779-8ce2-40f0a97d8348"
    Public Const EventsId As String = "f92244aa-4d48-465a-bb71-a9b7ac251e8e"
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
        AddItem("IAIS.ExceptionReview", 1)
        AddItem("IAIS.ExceptionReview", 2)
        AddItem("IAIS.ExceptionReview", 3)
        AddItem("IAIS.ExceptionReview", 4)
        AddItem("IAIS.ExceptionReview", 5)
        AddItem("IAIS.ExceptionReview", 6)
        BeginGroup()
        AddItem("IAIS.ManageColConflict")

    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "IAIS Exception Menu"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "IAIS Exception Menu"
        End Get
    End Property

End Class


