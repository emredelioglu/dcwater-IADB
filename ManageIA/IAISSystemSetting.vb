Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.JTX

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

    Private m_pIAManageExt As IAManageExt

    Private m_pForm As FormSystemSetting


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

        m_pIAManageExt = MapUtil.GetExtention("IAIS.IAManageExt", m_app)
        If Not m_pIAManageExt Is Nothing Then
            AddHandler m_pIAManageExt.IAManageExtEvent, AddressOf onIAManageExtEvent
        End If
    End Sub

    Private Sub onIAManageExtEvent(ByVal strMsg As String)
        If strMsg = "disable" Then
            If Not m_pForm Is Nothing And m_pForm.Visible Then
                m_pForm.Close()
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()

        Try


            If Not IAISToolSetting.Initilize(m_app, forceFlag:=True) Then
                MsgBox("Can't initilize system setting.")
                Return
            End If



            Dim registryKey As Microsoft.Win32.RegistryKey
            registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\IAIS")
            Dim tnIAISSetting As String = registryKey.GetValue("SYSTEM_TABLE")

            Dim pTable As ITable = Nothing

            Dim jtxExt As IJTXExtension = DirectCast(m_app.FindExtensionByName("Workflow Manager"), IJTXExtension)
            If Not jtxExt.Job Is Nothing Then

                pTable = CType(jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName), IFeatureWorkspace).OpenTable(tnIAISSetting)
            End If

            If pTable Is Nothing Then
                Dim pMap As IMap = m_doc.FocusMap
                pTable = MapUtil.GetPremiseWS(pMap).OpenTable(tnIAISSetting)
            End If

            If pTable Is Nothing Then
                MsgBox("Can't connect to the system setting table")
                Return
            End If

            m_pForm = New FormSystemSetting

            Dim pCursor As ICursor
            Dim pRow As IRow

            pCursor = pTable.Search(Nothing, False)
            pRow = pCursor.NextRow
            Do While Not pRow Is Nothing
                m_pForm.AddVariable(pRow.Value(pRow.Fields.FindField("NAME")), pRow.Value(pRow.Fields.FindField("VALUE")))
                pRow = pCursor.NextRow
            Loop

            m_pForm.DataGridVariable.AutoResizeColumns()

            m_pForm.m_table = pTable

            m_pForm.ShowDialog()
        Catch ex As Exception
            MsgBox("Err" & ex.Source & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.IAManageExt", m_app) Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

End Class


