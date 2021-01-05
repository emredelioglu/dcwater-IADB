Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.SystemUI
Imports IAIS.Windows.UI
Imports System.Windows.Forms
Imports System.Collections.Generic

Imports ESRI.ArcGIS.JTX

<CLSCompliant(False)> _
<ComClass(ExceptionReview.ClassId, ExceptionReview.InterfaceId, ExceptionReview.EventsId)> _
<ProgId("IAIS.ExceptionReview")> _
Public Class ExceptionReview
    Inherits BaseCommand
    Implements ICommandSubType

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_subtype As Integer

    Private m_dockableWindow As IDockableWindow

    Private Const DockableWindowGuid As String = "{630418E7-267B-432f-8622-810771F3BA4C}"

    Private m_pIAManageExt As IAManageExt

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e3197090-4392-412c-8179-3cae7c2bee81"
    Public Const InterfaceId As String = "4013679c-8f6c-416d-b6b0-e678b8c21c75"
    Public Const EventsId As String = "8499b452-5460-4b58-b7a3-8f220f6bc423"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        MyBase.m_category = "IAIS Tools"

    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Select Case m_subtype
                Case ExceptionWindow.exceptionPremiseExport
                    Return "Premise Export Exceptions"
                Case ExceptionWindow.exceptionPremiseImport
                    Return "Premise Import Exceptions"
                Case ExceptionWindow.exceptionMAR
                    Return "MAR Exceptions"
                Case ExceptionWindow.exceptionOwnerPly
                    Return "OwnerPly Exceptions"
                Case ExceptionWindow.exceptionIAChargeFile
                    Return "IA Charge File Exceptions"
                Case ExceptionWindow.exceptionIAAssignPly
                    Return "IAAssignPly Exceptions"
            End Select

            Return ""
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Select Case m_subtype
                Case ExceptionWindow.exceptionPremiseExport
                    Return "Review Premise Export Exceptions"
                Case ExceptionWindow.exceptionPremiseImport
                    Return "Review Premise Import Exceptions"
                Case ExceptionWindow.exceptionMAR
                    Return "Review MAR Exceptions"
                Case ExceptionWindow.exceptionOwnerPly
                    Return "Review OwnerPly Exceptions"
                Case ExceptionWindow.exceptionIAChargeFile
                    Return "Review IA Charge File Exceptions"
                Case ExceptionWindow.exceptionIAAssignPly
                    Return "Review IAAssignPly Exceptions"
            End Select

            Return ""
        End Get
    End Property


    Public Overrides Sub OnCreate(ByVal hook As Object)

        If Not hook Is Nothing Then
            m_app = CType(hook, IApplication)
            m_doc = m_app.Document

        End If

        If m_app IsNot Nothing Then
            SetupExceptionWindow()
            MyBase.m_enabled = m_dockableWindow IsNot Nothing
        Else
            MyBase.m_enabled = False
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

    Public Function GetCount() As Integer Implements ESRI.ArcGIS.SystemUI.ICommandSubType.GetCount
        Return 6
    End Function

    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            If m_dockableWindow Is Nothing Then
                Return False
            End If

            Dim pExpForm As ExceptionWindow = TryCast(m_dockableWindow.UserData, ExceptionWindow)
            If pExpForm.ExceptionType <> m_subtype Then
                Return False
            Else
                Return m_dockableWindow.IsVisible()
            End If
        End Get
    End Property

    Public Sub SetSubType(ByVal SubType As Integer) Implements ESRI.ArcGIS.SystemUI.ICommandSubType.SetSubType
        m_subtype = SubType
    End Sub

    Public Overrides Sub OnClick()
        If m_dockableWindow Is Nothing Then Return

        Dim pExpForm As ExceptionWindow = TryCast(m_dockableWindow.UserData, ExceptionWindow)

        'If m_subtype = pExpForm.ExceptionType AndAlso m_dockableWindow.IsVisible() Then
        '    m_dockableWindow.Show(False)
        '    Return
        'End If
        Dim dt As DataTable = Nothing
        'Dim ds As DataSet = Nothing

        Try
            dt = GetDataTable()

            If dt Is Nothing Then
                Return
            End If

            If dt.Rows.Count = 0 Then
                MsgBox("There is no exception to be reviewed.")
                If m_dockableWindow.IsVisible() Then
                    m_dockableWindow.Dock(esriDockFlags.esriDockHide)
                End If

                Return

            End If

            Dim bindingSrc As BindingSource = New BindingSource
            bindingSrc.DataSource = dt 'ds.Tables("Exceptions")



            pExpForm.ExceptionType = m_subtype
            pExpForm.GMap = m_doc.FocusMap

            pExpForm.m_bReviewStart = False

            pExpForm.DataGridExceptions.DataSource = bindingSrc
            pExpForm.BindingNavigator1.BindingSource = bindingSrc
            pExpForm.DataGridExceptions.CurrentCell = Nothing

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString())
            Return
        End Try

        If Not m_dockableWindow.IsVisible() Then
            m_dockableWindow.Dock(esriDockFlags.esriDockShow)
        End If

    End Sub


    Private Function GetDataTable() As DataTable

        Try
            Dim dt As DataTable = New DataTable()

            Dim jtxExt As IJTXExtension = DirectCast(m_app.FindExtensionByName("Workflow Manager"), IJTXExtension)
            If jtxExt.Job Is Nothing Then
                MsgBox("No JTX job is open")
                Return Nothing
            End If

            Dim pFWS As IFeatureWorkspace = jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)

            Dim pTable As ITable

            Dim pQFilter As IQueryFilter = New QueryFilter
            pQFilter.WhereClause = "FIXDT IS NULL"

            Select Case m_subtype
                Case ExceptionWindow.exceptionPremiseExport
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSEXPOUT"))
                Case ExceptionWindow.exceptionPremiseImport
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSINTERPT"))
                Case ExceptionWindow.exceptionMAR
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSADDRESSPT"))
                Case ExceptionWindow.exceptionOwnerPly
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSOWNERPLY"))
                Case ExceptionWindow.exceptionIAChargeFile
                    'pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_CISOUTPUTBILLDET"))
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSBILLDET"))
                Case ExceptionWindow.exceptionIAAssignPly
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSIAPLY"))
                Case Else
                    Return Nothing
            End Select

            Dim pCursor As ICursor
            Dim pRow As IRow


            Dim dc As DataColumn
            Dim colIndex As Integer
            For colIndex = 0 To pTable.Fields.FieldCount - 1
                dc = New DataColumn()
                If pTable.Fields.Field(colIndex).Name = pTable.OIDFieldName Then
                    dc.ColumnName = "objectid"
                Else
                    If pTable.Fields.Field(colIndex).Name = "PEXUID" AndAlso
                        (m_subtype = ExceptionWindow.exceptionPremiseExport Or
                        m_subtype = ExceptionWindow.exceptionPremiseImport Or
                        m_subtype = ExceptionWindow.exceptionIAChargeFile) Then
                        dc.ColumnName = "uniqueid"
                    ElseIf pTable.Fields.Field(colIndex).Name = "ADDRESS_ID" AndAlso
                        m_subtype = ExceptionWindow.exceptionMAR Then
                        dc.ColumnName = "uniqueid"
                    ElseIf pTable.Fields.Field(colIndex).Name = "OWNER_GIS_ID" AndAlso
                        m_subtype = ExceptionWindow.exceptionOwnerPly Then
                        dc.ColumnName = "uniqueid"
                    ElseIf pTable.Fields.Field(colIndex).Name = "IAID" AndAlso
                        m_subtype = ExceptionWindow.exceptionIAAssignPly Then
                        dc.ColumnName = "uniqueid"
                    Else
                        dc.ColumnName = pTable.Fields.Field(colIndex).Name
                    End If
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

            Return dt

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace.ToString())
            Return Nothing
        End Try




    End Function


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

    Private Function GetExceptions() As IList
        Dim expList As IList = New List(Of clsIAISException)
        Try

            Dim jtxExt As IJTXExtension = DirectCast(m_app.FindExtensionByName("Workflow Manager"), IJTXExtension)
            If jtxExt.Job Is Nothing Then
                MsgBox("No JTX job is open")
                Return expList
            End If


            Dim pFWS As IFeatureWorkspace = jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)

            'Dim pFWS As IFeatureWorkspace = MapUtil.GetPremiseWS(m_doc.FocusMap) ''

            Dim pTable As ITable

            Select Case m_subtype
                Case ExceptionWindow.exceptionPremiseExport
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSEXPOUT"))
                Case ExceptionWindow.exceptionPremiseImport
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSPREMSINTERPT"))
                Case ExceptionWindow.exceptionMAR
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSADDRESSPT"))
                Case ExceptionWindow.exceptionOwnerPly
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSOWNERPLY"))
                Case ExceptionWindow.exceptionIAChargeFile
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSBILLDET"))
                Case ExceptionWindow.exceptionIAAssignPly
                    pTable = MapUtil.GetTableFromWS(pFWS, IAISToolSetting.GetParameterValue("TABLE_EXCEPTIONSIAPLY"))
                Case Else
                    Return Nothing
            End Select


            Dim pCursor As ICursor
            Dim pRow As IRow

            Dim pQFilter As IQueryFilter = New QueryFilter
            If m_subtype = ExceptionWindow.exceptionPremiseExport Then
                pQFilter.WhereClause = "FIXDT IS NULL"
            End If


            pCursor = pTable.Search(pQFilter, False)
            'pCursor = pTable.Search(Nothing, False)
            pRow = pCursor.NextRow
            Do While Not pRow Is Nothing
                Dim iaisExp As clsIAISException = New clsIAISException

                If m_subtype = ExceptionWindow.exceptionPremiseExport Then

                    iaisExp.Excepttyp = MapUtil.GetRowValue(pRow, "EXCEPTIONMSG", True)
                Else
                    iaisExp.Exceptdt = MapUtil.GetRowValue(pRow, "EXCEPTDT", True)
                    iaisExp.Excepttyp = MapUtil.GetRowValue(pRow, "EXCEPTTYP", True)
                    iaisExp.AppName = MapUtil.GetRowValue(pRow, "APPNAME", True)
                    iaisExp.Apptype = MapUtil.GetRowValue(pRow, "APPTYPE", True)
                    iaisExp.Reviewst = MapUtil.GetRowValue(pRow, "REVIEWST", True)
                    iaisExp.Revdt = MapUtil.GetRowValue(pRow, "REVDT", True)
                    iaisExp.Status = MapUtil.GetRowValue(pRow, "STATUS", True)
                    iaisExp.Reviewby = MapUtil.GetRowValue(pRow, "REVIEWBY", True)
                    iaisExp.Fixdt = MapUtil.GetRowValue(pRow, "FIXDT", True)
                    iaisExp.Fixby = MapUtil.GetRowValue(pRow, "FIXBY", True)
                End If


                Select Case m_subtype
                    Case ExceptionWindow.exceptionPremiseExport
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "PEXUID", True)
                    Case ExceptionWindow.exceptionPremiseImport
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "PEXUID", True)
                    Case ExceptionWindow.exceptionMAR
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "ADDRESS_ID", True)
                    Case ExceptionWindow.exceptionOwnerPly
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "OWNER_GIS_ID", True)
                    Case ExceptionWindow.exceptionIAChargeFile
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "PEXUID", True)
                    Case ExceptionWindow.exceptionIAAssignPly
                        iaisExp.UniqueId = MapUtil.GetRowValue(pRow, "IAID", True)
                End Select

                iaisExp.Objectid = pRow.OID

                expList.Add(iaisExp)

                pRow = pCursor.NextRow
            Loop

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace.ToString())
        End Try


        Return expList
    End Function


    Private Sub SetupExceptionWindow()
        If m_dockableWindow Is Nothing Then
            Dim dockWindowManager As IDockableWindowManager
            dockWindowManager = CType(m_app, IDockableWindowManager)
            If Not dockWindowManager Is Nothing Then
                Dim windowID As UID = New UIDClass
                windowID.Value = DockableWindowGuid
                m_dockableWindow = dockWindowManager.GetDockableWindow(windowID)
            End If
        End If

        If Not m_dockableWindow Is Nothing Then
            m_dockableWindow.Dock(esriDockFlags.esriDockHide)
            'm_dockableWindow.Show(False)
        End If
    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.IAManageExt", m_app) Then
                Return False
            End If

            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            'If Not IAISApplication.GetInstance.IsInJTXApplication() Then
            '    Return False
            'End If


            Return True

        End Get
    End Property

End Class


