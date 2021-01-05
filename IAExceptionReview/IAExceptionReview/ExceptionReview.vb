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

        If m_subtype = pExpForm.ExceptionType AndAlso m_dockableWindow.IsVisible() Then
            m_dockableWindow.Show(False)
            Return
        End If

        Dim dt As DataTable = GetDataTable()
        Dim bindingSrc As BindingSource = New BindingSource
        bindingSrc.DataSource = dt

        pExpForm.DataGridExceptions.DataSource = bindingSrc
        pExpForm.BindingNavigator1.BindingSource = bindingSrc
        pExpForm.ExceptionType = m_subtype

        If Not m_dockableWindow.IsVisible() Then
            m_dockableWindow.Show(True)
        End If

    End Sub


    Private Function GetDataTable() As DataTable
        Dim pMap As IMap = m_doc.FocusMap
        Dim pLayer As IFeatureLayer
        pLayer = MapUtil.GetLayerByTableName("PremsInterPt", pMap)
        Dim pDataset As IDataset = pLayer
        Dim pTable As ITable

        Dim pQFilter As IQueryFilter = New QueryFilter
        pQFilter.WhereClause = "STATUS IS NULL OR STATUS <= 1"

        Select Case m_subtype
            Case ExceptionWindow.exceptionPremiseExport
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsPremsInterPt")
            Case ExceptionWindow.exceptionPremiseImport
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsPremsInterPt")
            Case ExceptionWindow.exceptionMAR
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsAddressPt")
            Case ExceptionWindow.exceptionOwnerPly
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsOwnerPly")
            Case ExceptionWindow.exceptionIAChargeFile
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsBillDet")
            Case ExceptionWindow.exceptionIAAssignPly
                pTable = MapUtil.GetTableFromWS(CType(pDataset.Workspace, IFeatureWorkspace), "ExceptionsIAPly")
            Case Else
                Return Nothing
        End Select

        Dim pCursor As ICursor
        Dim pRow As IRow

        Dim dt As DataTable = New DataTable()
        Dim dc As DataColumn
        Dim colIndex As Integer
        For colIndex = 0 To pTable.Fields.FieldCount - 1
            dc = New DataColumn(pTable.Fields.Field(colIndex).AliasName)
            dt.Columns.Add(dc)
        Next

        pCursor = pTable.Search(pQFilter, False)
        pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            Dim dr As DataRow = dt.NewRow()

            For colIndex = 0 To pTable.Fields.FieldCount - 1
                dr.Item(colIndex) = GetRowValue(pRow, colIndex)
            Next
            pRow = pCursor.NextRow
        Loop

        Return dt
    End Function

    Private Function GetRowValue(ByVal pRow As IRow, ByVal colIndex As Integer) As String
        If IsDBNull(pRow.Value(colIndex)) Then
            Return ""
        Else
            Dim pDomain As IDomain
            pDomain = pRow.Fields.Field(colIndex)

            If TypeOf pDomain Is ICodedValueDomain Then
                Dim pCodeDomain As ICodedValueDomain = pDomain

                Dim cindex As Long
                For cindex = 0 To pCodeDomain.CodeCount - 1
                    If pCodeDomain.Value(cindex) = pRow.Value(colIndex) Then
                        Return pCodeDomain.Name(cindex)
                    End If
                Next
                Return pRow.Value(colIndex)
            Else
                Return pRow.Value(colIndex)
            End If
        End If
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
    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.Extension", m_app) Then
                Return False
            End If
        End Get
    End Property

End Class


