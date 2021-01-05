Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.BaseClasses

Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.JTX
Imports System.Windows.Forms

<CLSCompliant(False)> _
<ComClass(ManageColConflict.ClassId, ManageColConflict.InterfaceId, ManageColConflict.EventsId)> _
<ProgId("IAIS.ManageColConflict")> _
Public Class ManageColConflict
    Inherits BaseCommand


    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_dockableWindow As IDockableWindow

    Private Const DockableWindowGuid As String = "{4467D820-6F41-41de-A08F-AA792D29FCB1}"

    Private m_pIAManageExt As IAManageExt

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

        If m_dockableWindow Is Nothing Then
            Dim dockWindowManager As IDockableWindowManager
            dockWindowManager = CType(m_app, IDockableWindowManager)
            If Not dockWindowManager Is Nothing Then
                Dim windowID As UID = New UIDClass
                windowID.Value = DockableWindowGuid
                m_dockableWindow = dockWindowManager.GetDockableWindow(windowID)
            End If
        End If

        m_pIAManageExt = MapUtil.GetExtention("IAIS.IAManageExt", m_app)
        If Not m_pIAManageExt Is Nothing Then
            AddHandler m_pIAManageExt.IAManageExtEvent, AddressOf onIAManageExtEvent
        End If
    End Sub

    Private Sub onIAManageExtEvent(ByVal strMsg As String)
        If strMsg = "disable" Then
            If Not m_dockableWindow Is Nothing Then
                m_dockableWindow.Dock(esriDockFlags.esriDockHide)
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()
        If m_dockableWindow Is Nothing Then Return

        Dim pColForm As PropertyConflictLineageWindow = TryCast(m_dockableWindow.UserData, PropertyConflictLineageWindow)

        If m_dockableWindow.IsVisible() Then
            m_dockableWindow.Show(False)
            Return
        End If

        Dim jtxExt As IJTXExtension = DirectCast(m_app.FindExtensionByName("Workflow Manager"), IJTXExtension)
        If jtxExt.Job Is Nothing Then
            MsgBox("No JTX job is open")
            Return
        End If

        Dim pFWS As IFeatureWorkspace = jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)



        Dim pTable As ITable

        Dim pQFilter As IQueryFilter = New QueryFilter
        pQFilter.WhereClause = "PROCESSDT IS NULL"

        pTable = MapUtil.GetTableFromWS(CType(pFWS, IFeatureWorkspace), IAISToolSetting.GetParameterValue("TABLE_PROPERTYCONFLICTLINEAGE"))

        Dim pCursor As ICursor
        Dim pRow As IRow

        Dim dt As DataTable = New DataTable()

        Dim dc As DataColumn
        Dim colIndex As Integer
        For colIndex = 0 To pTable.Fields.FieldCount - 1
            dc = New DataColumn()
            If pTable.Fields.Field(colIndex).Name = pTable.OIDFieldName Then
                dc.ColumnName = "objectid"
            Else
                dc.ColumnName = pTable.Fields.Field(colIndex).Name
            End If

            dc.Caption = pTable.Fields.Field(colIndex).AliasName
            dt.Columns.Add(dc)
        Next

        pCursor = pTable.Search(pQFilter, False)
        pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            Dim dr As DataRow = dt.NewRow()

            For colIndex = 0 To pTable.Fields.FieldCount - 1
                If pTable.Fields.Field(colIndex).Name = pTable.OIDFieldName Then
                    dr.Item(colIndex) = pRow.OID
                Else
                    dr.Item(colIndex) = GetRowValue(pRow, colIndex, True)
                End If
            Next
            dt.Rows.Add(dr)

            pRow = pCursor.NextRow
        Loop

        If dt.Rows.Count = 0 Then
            MsgBox("There is no exception.")
            Return
        End If


        pColForm.GMap = m_doc.FocusMap

        Dim bindingSrc As BindingSource = New BindingSource
        bindingSrc.DataSource = dt

        pColForm.DataGridExceptions.DataSource = bindingSrc
        pColForm.BindingNavigator1.BindingSource = bindingSrc

        pColForm.DataGridExceptions.CurrentCell = Nothing


        If Not m_dockableWindow.IsVisible() Then
            m_dockableWindow.Show(True)
        End If

    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If m_dockableWindow Is Nothing Then
                Return False
            End If


            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            If Not IAISApplication.GetInstance.IsInJTXApplication() Then
                Return False
            End If


            'Return jtxExt.Job.JobType.Name = "Property Update Review"
            Return True

        End Get
    End Property

    Private Function GetRowValue(ByVal pRow As IRow, ByVal fIndex As Integer, Optional ByVal getDomainValue As Boolean = False) As String
        If fIndex < 0 Then
            Return ""
        End If

        If IsDBNull(pRow.Value(fIndex)) Then
            Return ""
        Else
            If Not getDomainValue Then
                Return pRow.Value(fIndex)
            Else
                Dim pDomain As IDomain
                pDomain = pRow.Fields.Field(fIndex).Domain

                If Not pDomain Is Nothing AndAlso TypeOf pDomain Is ICodedValueDomain Then
                    Dim pCodeDomain As ICodedValueDomain = pDomain

                    Dim cindex As Long
                    For cindex = 0 To pCodeDomain.CodeCount - 1
                        If pCodeDomain.Value(cindex) = pRow.Value(fIndex) Then
                            Return pCodeDomain.Name(cindex)
                        End If
                    Next
                    Return pRow.Value(fIndex)
                Else
                    Return pRow.Value(fIndex)
                End If
            End If
        End If
    End Function


End Class


