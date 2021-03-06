﻿Imports ESRI.ArcGIS.ADF.CATIDs
Imports System.Runtime.InteropServices

<ComClass(IAManageToolbar.ClassId, IAManageToolbar.InterfaceId, IAManageToolbar.EventsId), _
ProgId("IAIS.IAManageToolbar")> _
Public Class IAManageToolbar
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
    Public Const ClassId As String = "B4D4B2FB-F934-41fa-A13A-A47A210BF1F0"
    Public Const InterfaceId As String = "E2CC2BAE-AEBE-4d29-A7BB-C746A3C40A48"
    Public Const EventsId As String = "40F83159-D153-45ae-88C5-8756937AB2C6"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        AddItem("IAIS.IASourceListControl")
        AddItem("IAIS.IABreak")
        AddItem("IAIS.IABuldAssignment")
        AddItem("IAIS.IABuldSplit")
        AddItem("IAIS.IABuldManual")
        AddItem("IAIS.IAAssign")
        AddItem("IAIS.IACFCalc")
        AddItem("IAIS.IADirectAssign")
        AddItem("IAIS.IABreakDirectIA")
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "IAIS Manage IA Toolbar"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "IAIS Manage IA Toolbar"
        End Get
    End Property

End Class


