
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
<ComClass(IADirectAssign.ClassId, IADirectAssign.InterfaceId, IADirectAssign.EventsId)> _
<ProgId("IAIS.IADirectAssign")> _
Public Class IADirectAssign
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "89c41844-bdc5-4328-8578-ffff9eeb923b"
    Public Const InterfaceId As String = "413c53fa-9bc9-4b78-8300-a824b6d4db3e"
    Public Const EventsId As String = "e5de4911-368e-4799-ae6c-dd6dbd82f7a4"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_ia_assign.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Direct IA Assign"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Direct IA Assign"
        MyBase.m_name = "Direct IA Assign"
        MyBase.m_toolTip = "Direct IA Assign"
    End Sub


    Public Overrides Sub OnClick()
        Dim pMap As IMap
        pMap = m_doc.FocusMap

        Dim premiseLayer As IFeatureLayer
        premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

        Dim pSel As IFeatureSelection
        Dim pFeatCursor As IFeatureCursor

        Dim pPt As IFeature

        Dim iaLayer As IFeatureLayer
        iaLayer = IAToolUtil.GetTargetLayer(m_app)
        pSel = iaLayer
        pSel.SelectionSet.Search(Nothing, False, pFeatCursor)

        Dim ia As IFeature
        ia = pFeatCursor.NextFeature

        IAISApplication.GetInstance.StartToolEditing()
        Try
            'Now check to see if IA has been directly assign to a premise
            Dim pQFilter As IQueryFilter = New QueryFilter

            Dim existPexuid As String = MapUtil.GetValue(ia, "PEXUID")
            If existPexuid <> "" Then
                If MsgBox("Selected IA feature has a direct assignment to premise " & existPexuid & ". " & _
                          "Do you want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    IAISApplication.GetInstance.AbortToolEditing()
                    Return
                End If


                pQFilter.WhereClause = "PEXUID=" & existPexuid

                If iaLayer.FeatureClass.FeatureCount(pQFilter) = 1 Then
                    'reset the HAS_DIRECT_IAASSIGN field of the premise
                    pFeatCursor = premiseLayer.Search(pQFilter, False)
                    pPt = pFeatCursor.NextFeature
                    pPt.Value(pPt.Fields.FindField("HAS_DIRECT_IAASSIGN")) = System.DBNull.Value
                    pPt.Store()
                End If

            End If

            'Now set the HAS_DIRECT_IAASSIGN for selected premise
            pSel = premiseLayer
            pSel.SelectionSet.Search(Nothing, False, pFeatCursor)
            pPt = pFeatCursor.NextFeature

            'Check to see if direct assignment is allowed

            If MapUtil.GetValue(pPt, "IS_EXEMPT_IAB") = "Y" Then
                Throw New Exception("Selected premise is exempt from IA charge.")
            ElseIf MapUtil.GetValue(pPt, "PEXPTYP") = "RES" Then
                Throw New Exception("Direct assignment is not allowed for residential only property.")
            End If


            pPt.Value(pPt.Fields.FindField("HAS_DIRECT_IAASSIGN")) = "Y"
            pPt.Store()

            'set the PEXUID to selected IA feature
            ia.Value(ia.Fields.FindField("PEXUID")) = pPt.Value(pPt.Fields.FindField("PEXUID"))
            ia.Store()

            'Dim pForm As FormPickDate = New FormPickDate
            'pForm.ShowDialog()

            Dim form As FormPickDate = New FormPickDate
            form.LabelDate.Text = "Effective End Date"
            form.Text = "Effective End Date"
            form.ShowDialog()

            If form.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                IAISApplication.GetInstance.AbortToolEditing()
                Return
            End If


            Dim pIACFTable As ITable
            Dim pFeatWS As IFeatureWorkspace
            Dim pDataset As IDataset

            pDataset = premiseLayer
            pFeatWS = pDataset.Workspace

            pIACFTable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
            Dim pRow As IRow
            pQFilter.WhereClause = "EFFENDDT IS NULL AND PEXUID=" & MapUtil.GetValue(pPt, "PEXUID")

            Dim pCursor As ICursor
            pCursor = pIACFTable.Search(pQFilter, False)
            pRow = pCursor.NextRow
            If Not pRow Is Nothing Then
                pRow.Value(pRow.Fields.FindField("EFFENDDT")) = form.DateTimePicker1.Text
                pRow.Store()
            End If


            IAISApplication.GetInstance.StopToolEditing("Direct IA Assign")

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
                If pFeatSel.SelectionSet.Count <> 1 Then
                    Return False
                End If

                Dim iaLayer As IFeatureLayer
                iaLayer = IAToolUtil.GetTargetLayer(m_app)
                pFeatSel = iaLayer
                Return (pFeatSel.SelectionSet.Count = 1)

            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg
End Class


