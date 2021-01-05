Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto

<ComClass(IAISSystemSetting.ClassId, IAISSystemSetting.InterfaceId, IAISSystemSetting.EventsId)> _
<ProgId("IAIS.IAISSystemSetting")> _
Public Class IAISSystemSetting
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ef476759-c0aa-46c9-827b-e186e2621b52"
    Public Const InterfaceId As String = "5613b16f-3b8f-4b76-91e5-2b94e4bdc5fe"
    Public Const EventsId As String = "6289f918-bf81-4dcf-bac6-e2ad5e867b17"
#End Region

    Private m_app As IApplication
    Private m_doc As IMxDocument


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

        MyBase.m_caption = "System Settings"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "System Settings"
        MyBase.m_name = "System Settings"
        MyBase.m_toolTip = "System Settings"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document
    End Sub

    Public Overrides Sub OnClick()

        IAISToolSetting.Initilize(m_app)
        Dim pForm As FormSystemSetting = New FormSystemSetting


        'Get existing system settings
        Dim pLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_doc.FocusMap)
        Dim pDataset As IDataset = pLayer
        Dim pTable As ITable = CType(pDataset.Workspace, IFeatureWorkspace).OpenTable("IAIS_TOOL_SETTING")

        Dim pCursor As ICursor
        Dim pRow As IRow

        pCursor = pTable.Search(Nothing, False)
        pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            pForm.AddVariable(pRow.Value(pRow.Fields.FindField("NAME")), pRow.Value(pRow.Fields.FindField("VALUE")))
            pRow = pCursor.NextRow
        Loop

        pForm.DataGridVariable.AutoResizeColumns()

        pForm.m_table = pTable

        pForm.ShowDialog()

    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.Extension", m_app) Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

End Class


