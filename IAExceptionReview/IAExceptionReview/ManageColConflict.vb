Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.BaseClasses

Imports ESRI.ArcGIS.Carto

<CLSCompliant(False)> _
<ComClass(ManageColConflict.ClassId, ManageColConflict.InterfaceId, ManageColConflict.EventsId)> _
<ProgId("IAIS.ManageColConflict")> _
Public Class ManageColConflict
    Inherits BaseCommand


    Private m_app As IApplication
    Private m_doc As IMxDocument


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "708a193d-09eb-4a2f-ac5c-379fc984a2e2"
    Public Const InterfaceId As String = "b6bc9d3b-6a77-4727-aa25-bde916240f15"
    Public Const EventsId As String = "9a5bc172-74d4-4469-8d3c-12ab8a9452a9"
#End Region

#Region "Component Category Registration"
    ' The below automatically adds the Component Category registration.
    <ComRegisterFunction()> Shared _
      Sub Reg(ByVal regKey As [String])
        MxCommands.Register(regKey)
    End Sub 'Reg

    <ComUnregisterFunction()> Shared _
    Sub Unreg(ByVal regKey As [String])
        MxCommands.Unregister(regKey)
    End Sub 'Unreg
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        MyBase.m_caption = "Manage Conflict Lineage"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Manage Conflict Lineage"
        MyBase.m_name = "Manage Conflict Lineage"
        MyBase.m_toolTip = "Manage Conflict Lineage"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document
    End Sub

End Class


