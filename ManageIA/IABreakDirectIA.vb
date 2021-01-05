
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.ADF.CATIDs

<CLSCompliant(False)> _
<ComClass(IABreakDirectIA.ClassId, IABreakDirectIA.InterfaceId, IABreakDirectIA.EventsId)> _
<ProgId("IAIS.IABreakDirectIA")> _
Public Class IABreakDirectIA
    Inherits BaseCommand
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e3425a78-c870-43be-b74e-97b5f9460021"
    Public Const InterfaceId As String = "737a25af-966b-425b-b24d-6a8a9150a5fe"
    Public Const EventsId As String = "c560b8bf-68b4-426b-ac9b-4e7415802b82"
#End Region

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor


#Region "Component Category Registration"
    ' The below automatically adds the Component Category registration.
    <ComRegisterFunction()> Shared _
      Sub Reg(ByVal regKey As [String])
        MxCommands.Register(regKey)
    End Sub 'Reg

    <ComUnregisterFunction()> Shared _
    Sub Unreg(ByVal regKey As [String])
        MxCommands.Unregister(regKey)
    End Sub


#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_break_ia_direct_assign.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Break Direct IA Assign"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Break Direct IA Assign"
        MyBase.m_name = "Break Direct IA Assign"
        MyBase.m_toolTip = "Break Direct IA Assign"
    End Sub


    Public Overrides Sub OnClick()

        Dim pMap As IMap
        pMap = m_doc.FocusMap

        Dim premiseLayer As IFeatureLayer
        premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

        Dim pSel As IFeatureSelection
        Dim pFeatCursor As IFeatureCursor
        Dim pPt As IFeature

        pSel = premiseLayer
        pSel.SelectionSet.Search(Nothing, False, pFeatCursor)
        pPt = pFeatCursor.NextFeature

        If MapUtil.GetRowValue(pPt, "HAS_DIRECT_IAASSIGN") <> "Y" Then
            MsgBox("Selected premise does not have direct IA assignment")
            Return
        End If

        Dim iaLayer As IFeatureLayer
        Dim ia As IFeature



        IAISApplication.GetInstance.StartToolEditing()
        Try
            'remove the HAS_DIRECT_IAASSIGN for selected premise
            pPt.Value(pPt.Fields.FindField("HAS_DIRECT_IAASSIGN")) = System.DBNull.Value
            pPt.Store()

            'remove PEXUID from ia layers one by one 
            Dim pexuid As String = MapUtil.GetRowValue(pPt, "PEXUID")
            Dim pFilter As IQueryFilter = New QueryFilter
            pFilter.WhereClause = "PEXUID=" & pexuid
            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_IAASSIGNPLY"), pMap)
            If Not iaLayer Is Nothing Then
                pFeatCursor = iaLayer.Search(pFilter, False)
                ia = pFeatCursor.NextFeature
                Do While Not ia Is Nothing
                    ia.Value(ia.Fields.FindField("PEXUID")) = System.DBNull.Value
                    ia.Store()

                    ia = pFeatCursor.NextFeature
                Loop
            End If

            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_REVIAASSIGNPLY"), pMap)
            If Not iaLayer Is Nothing Then
                pFeatCursor = iaLayer.Search(pFilter, False)
                ia = pFeatCursor.NextFeature
                Do While Not ia Is Nothing
                    ia.Value(ia.Fields.FindField("PEXUID")) = System.DBNull.Value
                    ia.Store()

                    ia = pFeatCursor.NextFeature
                Loop
            End If

            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_APPEALIAASSIGNPLY", m_app), pMap)
            If Not iaLayer Is Nothing Then
                pFeatCursor = iaLayer.Search(pFilter, False)
                ia = pFeatCursor.NextFeature
                Do While Not ia Is Nothing
                    ia.Value(ia.Fields.FindField("PEXUID")) = System.DBNull.Value
                    ia.Store()

                    ia = pFeatCursor.NextFeature
                Loop
            End If

            Dim pIACFTable As ITable
            Dim pFeatWS As IFeatureWorkspace
            Dim pDataset As IDataset

            pDataset = premiseLayer
            pFeatWS = pDataset.Workspace

            pIACFTable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
            Dim pRow As IRow
            Dim pQFilter As IQueryFilter = New QueryFilter
            pQFilter.WhereClause = "EFFENDDT IS NULL AND PEXUID=" & MapUtil.GetValue(pPt, "PEXUID")

            Dim pCursor As ICursor
            pCursor = pIACFTable.Search(pQFilter, False)
            pRow = pCursor.NextRow
            If Not pRow Is Nothing Then

                Dim form As FormPickDate = New FormPickDate
                form.LabelDate.Text = "Effective End Date"
                form.Text = "Effective End Date"
                form.ShowDialog()

                If form.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                    IAISApplication.GetInstance.AbortToolEditing()
                    Return
                End If

                pRow.Value(pRow.Fields.FindField("EFFENDDT")) = form.DateTimePicker1.Text
                pRow.Store()
            Else
                MsgBox("Direct IA assignment has been removed for premise " & pexuid)
            End If

            IAISApplication.GetInstance.StopToolEditing("Break Direct IA Assign")


            Return
        Catch ex As Exception
            IAISApplication.GetInstance.AbortToolEditing()
            MsgBox(ex.Message)
        End Try


    End Sub

    ''' <param name="hook">
    ''' A reference to the application in which the command was created.
    '''            The hook may be an IApplication reference (for commands created in ArcGIS Desktop applications)
    '''            or an IHookHelper reference (for commands created on an Engine ToolbarControl).
    ''' </param>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)

    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Try

                If Not MapUtil.IsExtentionEnabled("IAIS.IAManageExt", m_app) Then
                    Return False
                End If

                If Not IAISToolSetting.isInitilized Then
                    Return False
                End If

                If Not IAISApplication.GetInstance.IsInJTXApplication() Then
                    Return False
                End If

                If m_Editor Is Nothing Then
                    Return False
                End If

                If m_Editor.EditState = esriEditState.esriStateNotEditing Then
                    Return False
                End If

                Dim pMap As IMap = m_doc.FocusMap

                Dim premiseLayer As IFeatureLayer
                premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)


                If premiseLayer Is Nothing Then
                    Return False
                End If

                Dim pFeatSel As IFeatureSelection
                pFeatSel = premiseLayer
                Return (pFeatSel.SelectionSet.Count = 1)

            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg
End Class


