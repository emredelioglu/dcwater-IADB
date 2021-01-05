Imports System.Runtime.InteropServices
Imports System.Collections.Generic

Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

<CLSCompliant(False)> _
<ComClass(IABreak.ClassId, IABreak.InterfaceId, IABreak.EventsId)> _
<ProgId("IAIS.IABreak")> _
Public Class IABreak
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "316dbc6b-29ae-414d-97bd-3fed58fa8c26"
    Public Const InterfaceId As String = "79d9e99f-56ca-4fd9-8f4a-5fd0a222b930"
    Public Const EventsId As String = "fa75fe6f-f120-471c-ba4f-ecd89deba5bb"
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
    End Sub


#End Region
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABreak).Assembly.GetManifestResourceStream("IAIS.ia_break_ia.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Break IA Assignment"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Break IA Assignment"
        MyBase.m_name = "Break IA Assignment"
        MyBase.m_toolTip = "Break IA Assignment"

    End Sub


    Public Overrides Sub OnClick()
        IAISApplication.GetInstance.StartToolEditing()
        Try

            Dim pFeatCursor As IFeatureCursor
            Dim pQFilter As IQueryFilter = New QueryFilter

            Dim pMap As IMap = m_doc.FocusMap

            Dim scratchLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SCRATCHPLY"), pMap)
            Dim scratchObj As IFeature

            Dim pCol As IFeature
            pCol = MapUtil.GetSelectedCol(pMap)
            If Not pCol Is Nothing Then
                Dim ssl As String
                ssl = MapUtil.GetValue(pCol, "SSL")

                'delete from IAAssignPly
                pQFilter.WhereClause = "SSL='" & ssl & "'"

                Dim pIALayer As IFeatureLayer
                Dim ia As IFeature
                pIALayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_IAASSIGNPLY"), m_doc.FocusMap)
                If pIALayer Is Nothing Then
                    Throw New Exception("IAASSIGNPLY layer is not loaded.")
                End If
                pFeatCursor = pIALayer.Search(pQFilter, False)
                ia = pFeatCursor.NextFeature
                While Not ia Is Nothing
                    If MapUtil.GetValue(ia, "PEXUID") <> "" Then
                        Throw New Exception("There are directly assigned IA. Use Break Direct IA tool first.")
                    End If

                    ia.Delete()
                    ia = pFeatCursor.NextFeature
                End While

                pIALayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_REVIAASSIGNPLY"), m_doc.FocusMap)
                If pIALayer Is Nothing Then
                    Throw New Exception("REVIAASSIGNPLY layer ia not loaded.")
                End If

                pFeatCursor = pIALayer.Search(pQFilter, False)
                ia = pFeatCursor.NextFeature
                While Not ia Is Nothing
                    If MapUtil.GetValue(ia, "PEXUID") <> "" Then
                        Throw New Exception("There are directly assigned IA. Use Break Direct IA tool first.")
                    End If

                    scratchObj = scratchLayer.FeatureClass.CreateFeature
                    scratchObj.Shape = ia.Shape
                    scratchObj.Value(scratchObj.Fields.FindField("FEATURETYPE")) = ia.Value(ia.Fields.FindField("FEATURETYPE"))
                    scratchObj.Value(scratchObj.Fields.FindField("SOURCETYPE")) = ia.Value(ia.Fields.FindField("SOURCE"))
                    scratchObj.Value(scratchObj.Fields.FindField("FEATUREOID")) = ia.Value(ia.Fields.FindField("FEATUREOID"))
                    scratchObj.Store()

                    ia.Delete()
                    ia = pFeatCursor.NextFeature
                End While

                pIALayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_APPEALIAASSIGNPLY"), m_doc.FocusMap)
                If pIALayer Is Nothing Then
                    Throw New Exception("APPEALIAASSIGNPLY layer ia not loaded.")
                End If

                pFeatCursor = pIALayer.Search(pQFilter, False)
                ia = pFeatCursor.NextFeature
                While Not ia Is Nothing
                    If MapUtil.GetValue(ia, "PEXUID") <> "" Then
                        Throw New Exception("There are directly assigned IA. Use Break Direct IA tool first.")
                    End If
                    scratchObj = scratchLayer.FeatureClass.CreateFeature
                    scratchObj.Shape = ia.Shape
                    scratchObj.Value(scratchObj.Fields.FindField("FEATURETYPE")) = ia.Value(ia.Fields.FindField("IATYPE"))
                    'scratchObj.Value(scratchObj.Fields.FindField("SOURCE")) = ia.Value(ia.Fields.FindField("SOURCE"))
                    scratchObj.Store()

                    ia.Delete()
                    ia = pFeatCursor.NextFeature
                End While


                Dim form As FormPickDate = New FormPickDate
                form.LabelDate.Text = "Effective End Date"
                form.Text = "Effective End Date"
                form.ShowDialog()

                If form.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                    IAISApplication.GetInstance.AbortToolEditing()
                    Return
                End If

                'now get the premise point for this parcel
                'If Trim(Mid(ssl, 5, 4)) = "" Then
                '    pQFilter.WhereClause = "PEXSQUARE='" & Left(ssl, 4) & "' AND " & _
                '                            "PEXLOT='" & Mid(ssl, 9) & "' AND (IS_EXEMPT_IAB='N' OR IS_EXEMPT_IAB IS NULL) "
                'Else
                '    pQFilter.WhereClause = "PEXSQUARE='" & Left(ssl, 4) & "' AND " & _
                '                            "PEXSUFFIX='" & Mid(ssl, 5, 4) & "' AND " & _
                '                            "PEXLOT='" & Mid(ssl, 9) & "' AND (IS_EXEMPT_IAB='N' OR IS_EXEMPT_IAB IS NULL) "
                'End If

                Dim pSFilter As ISpatialFilter = New SpatialFilter
                pSFilter.Geometry = pCol.Shape
                pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                pSFilter.WhereClause = "IS_EXEMPT_IAB = 'N'"

                Dim pPremiseLayer As IFeatureLayer
                pPremiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_doc.FocusMap)

                pFeatCursor = pPremiseLayer.Search(pSFilter, False)
                'there should be only one or zero premise point for each parcel

                Dim pt As IFeature
                pt = pFeatCursor.NextFeature
                Do While Not pt Is Nothing
                    'ask user for the effective end date

                    'form.DateTimePicker1.Text

                    Dim pIACFTable As ITable
                    Dim pFeatWS As IFeatureWorkspace
                    Dim pDataset As IDataset

                    pDataset = pIALayer
                    pFeatWS = pDataset.Workspace

                    pIACFTable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
                    Dim pRow As IRow
                    pQFilter.WhereClause = "EFFENDDT IS NULL AND PEXUID=" & MapUtil.GetValue(pt, "PEXUID")

                    Dim pCursor As ICursor
                    pCursor = pIACFTable.Search(pQFilter, False)
                    pRow = pCursor.NextRow
                    If Not pRow Is Nothing Then
                        pRow.Value(pRow.Fields.FindField("EFFENDDT")) = form.DateTimePicker1.Text
                        pRow.Store()
                    End If

                    pt = pFeatCursor.NextFeature



                Loop

                'Reset directly assigned IA

                form = Nothing

                IAISApplication.GetInstance.StopToolEditing("Break IA Assignment")
                m_doc.ActiveView.Refresh()
            Else
                IAISApplication.GetInstance.AbortToolEditing()
            End If

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
    End Sub 'Unreg

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
                Return (MapUtil.GetSelectedColCount(pMap) = 1)

            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg
End Class


