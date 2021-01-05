Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Carto
Imports IAIS.Windows.UI

<CLSCompliant(False)> _
<ComClass(IAChargeHistory.ClassId, IAChargeHistory.InterfaceId, IAChargeHistory.EventsId)> _
<ProgId("IAIS.IAChargeHistory")> _
Public Class IAChargeHistory
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "c1365783-584d-4487-ab33-24c1e667730d"
    Public Const InterfaceId As String = "40194ef5-8f56-444e-8905-5afb7310107c"
    Public Const EventsId As String = "3b82af91-7488-4083-961f-870df9e25e79"
#End Region

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor

    Private m_checked2 As Boolean
    Private m_IACFForm As FormIACFList
    Private m_originalVersion As IVersion

    Private m_pIASearchReportExt As IASearchReportExt


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

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IAChargeHistory).Assembly.GetManifestResourceStream("IAIS.search_ia_charge_history.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "IA Charge History"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "IA Charge History"
        MyBase.m_name = "IA Charge History"

    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)

        m_checked2 = False

        m_pIASearchReportExt = MapUtil.GetExtention("IAIS.IASearchReportExt", m_app)
        If Not m_pIASearchReportExt Is Nothing Then
            AddHandler m_pIASearchReportExt.IASearchReportExtEvent, AddressOf onIASearchReportExtEvent
        End If

    End Sub

    Private Sub onIASearchReportExtEvent(ByVal strMsg As String)
        If strMsg = "disable" Then
            ResetVersion()
        End If
    End Sub

    Public Overrides ReadOnly Property Tooltip() As String
        Get
            If m_checked2 Then
                Return "Reset to transactional version"
            Else
                Return "IA Charge History"
            End If
        End Get
    End Property

    Public Overrides Sub OnClick()

        Try


            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As ILayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            Dim pDataset As IDataset

            If m_checked2 Then
                If Not m_IACFForm Is Nothing Then
                    m_IACFForm.Visible = False
                    m_IACFForm.Dispose()
                    m_IACFForm = Nothing
                End If

                'Reset to the original version
                Dim sVerionName As String = m_originalVersion.VersionName

                pDataset = premiseLayer
                Dim pVersionUtil As VersionUtil = New VersionUtil
                Try
                    pVersionUtil.ChangeVersionByName(m_doc, pDataset.Workspace, sVerionName)
                Catch ex As Exception
                    MsgBox("Error: " & ex.Source & vbCrLf & ex.Message)
                End Try


                m_checked2 = False
                Return
            End If


            Dim pSel As IFeatureSelection
            pSel = premiseLayer

            Dim pFeature As IFeature
            Dim pFeatCursor As IFeatureCursor
            pSel.SelectionSet.Search(Nothing, False, pFeatCursor)
            pFeature = pFeatCursor.NextFeature

            'Get a list of IACF record 
            pDataset = premiseLayer

            m_originalVersion = pDataset.Workspace

            Dim iacfTable As ITable = CType(pDataset.Workspace, IFeatureWorkspace).OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
            Dim pQFilter As IQueryFilter = New QueryFilter
            pQFilter.WhereClause = "PEXUID=" & MapUtil.GetValue(pFeature, "PEXUID")


            Dim pTableSort As ITableSort = New TableSort
            With pTableSort
                .Fields = "EFFSTARTDT"
                .setAscending("EFFSTARTDT", True)
                .QueryFilter = pQFilter
                .Table = iacfTable
            End With

            pTableSort.Sort(Nothing)
            Dim pCursor As ICursor = pTableSort.Rows

            Dim pRow As IRow
            pRow = pCursor.NextRow

            If pRow Is Nothing Then
                MsgBox("There is no charge record for selected premise.")
                Return
            End If

            m_IACFForm = New FormIACFList
            m_IACFForm.MapDoc = m_doc
            m_IACFForm.PremiseMap = pMap

            Do While Not pRow Is Nothing
                m_IACFForm.AddIACF(pRow)
                pRow = pCursor.NextRow
            Loop

            AddHandler m_IACFForm.FormClosing, AddressOf IACFFormClosing

            m_IACFForm.Show(New ModelessDialog(m_app.hWnd))

            m_checked2 = True
        Catch ex As Exception
            MsgBox("Error: " & ex.Source & vbCrLf & ex.Message)
        End Try

    End Sub


    Private Sub IACFFormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        If m_IACFForm.VersionCount = 0 Then
            m_checked2 = False
        End If
    End Sub

    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            Return m_checked2
        End Get
    End Property

    Private Sub ResetVersion()
        Try
            If Not m_IACFForm Is Nothing Then
                m_IACFForm.Visible = False
                m_IACFForm.Dispose()
                m_IACFForm = Nothing
            End If

            If m_originalVersion Is Nothing Then
                Return
            End If

            'Reset to the original version
            Dim sVerionName As String = m_originalVersion.VersionName

            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As ILayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            Dim pDataset As IDataset
            pDataset = premiseLayer
            Dim pVersionUtil As VersionUtil = New VersionUtil
            Try
                pVersionUtil.ChangeVersionByName(m_doc, pDataset.Workspace, sVerionName)
            Catch ex As Exception
                MsgBox("Error: " & ex.Source & vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
            End Try


            m_checked2 = False
            Return
        Catch ex As Exception
            MsgBox("Error: " & ex.Source & vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub


    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.IASearchReportExt", m_app) Then
                If m_checked2 Then
                    ResetVersion()
                End If
                Return False
            End If

            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            If Not m_Editor Is Nothing AndAlso m_Editor.EditState <> esriEditState.esriStateNotEditing Then
                Return False
            End If

            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            If premiseLayer Is Nothing Then
                Return False
            End If

            If m_checked2 Then
                Return True
            End If

            Dim pSel As IFeatureSelection
            pSel = premiseLayer
            Return (pSel.SelectionSet.Count = 1)
        End Get
    End Property

End Class


