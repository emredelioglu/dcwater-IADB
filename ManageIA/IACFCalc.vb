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
Imports System.Collections.Generic

<CLSCompliant(False)> _
<ComClass(IACFCalc.ClassId, IACFCalc.InterfaceId, IACFCalc.EventsId)> _
<ProgId("IAIS.IACFCalc")> _
Public Class IACFCalc
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument
    Private m_Editor As IEditor

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e3238408-fe03-4b4e-935b-47336cc6f2a2"
    Public Const InterfaceId As String = "b21270e2-9a06-4e87-8542-3863935a0d6a"
    Public Const EventsId As String = "2f884f4e-74be-4a39-9265-0602574521e5"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_calc_charges.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Calculate IA Charges"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Calculate IA Charges"
        MyBase.m_name = "Calculate IA Charges"
        MyBase.m_toolTip = "Calculate IA Charges"

    End Sub


    Public Overrides Sub OnClick()
        Try
            Dim pMap As IMap
            pMap = m_doc.FocusMap

            Dim pLayerParaName As String = "" 'Parcel layer parameter name
            Dim pCol As IFeature = MapUtil.GetSelectedCol(pMap, pLayerParaName)
            Dim ssl As String = MapUtil.GetValue(pCol, "SSL")

            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            'Spatial search to find all NO-EXAMPT premise point

            Dim pSFilter As ISpatialFilter = New SpatialFilter
            pSFilter.Geometry = pCol.ShapeCopy
            pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
            pSFilter.WhereClause = "NOT IS_EXEMPT_IAB='Y'"

            Dim pFeatCursor As IFeatureCursor
            Dim premiseCount As Integer = premiseLayer.FeatureClass.FeatureCount(pSFilter)

            If premiseCount = 0 Then
                MsgBox("Selected parcel do not have Non-Exempted premise point")
                Return
            Else
                pFeatCursor = premiseLayer.Search(pSFilter, False)

                Dim iacfList As ArrayList = New ArrayList()

                Dim pexptypDict As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
                Dim typCount As Integer

                Dim pPt As IFeature = pFeatCursor.NextFeature

                Do While Not pPt Is Nothing
                    Dim pexptyp As String = MapUtil.GetValue(pPt, "PEXPTYP")
                    If pexptypDict.TryGetValue(pexptyp, typCount) Then
                        pexptypDict.Item(pexptyp) = typCount + 1
                    Else
                        pexptypDict.Item(pexptyp) = 1
                    End If

                    Dim iacfItem As IACF = New IACF()
                    iacfItem.Pexptyp = pexptyp
                    iacfItem.Pexuid = MapUtil.GetValue(pPt, "PEXUID")

                    iacfList.Add(iacfItem)

                    pPt = pFeatCursor.NextFeature
                Loop

                Dim flagCheckDirectIA As Boolean = True
                Dim flagHasResca As Boolean = False

                If iacfList.Count > 1 Then
                    If pexptypDict.TryGetValue("RES", typCount) Then
                        If typCount > 1 Then
                            'Check to see if a RESCA exist
                            pSFilter.WhereClause = "EXEMPT_IAB_REASON='RESCA'"
                            If premiseLayer.FeatureClass.FeatureCount(pSFilter) <> 1 Then
                                MsgBox("There are multiple Non-Exempted residential premise point and no RESCA has been created.")
                                Return
                            Else
                                flagHasResca = True
                            End If
                        End If

                        If pexptypDict.Count = 1 Then
                            flagCheckDirectIA = False
                        End If
                    End If
                Else
                    flagCheckDirectIA = False
                End If

                'Check to make sure each non-RES premise point 
                'has direct IA, it will be the same as IA calculation
                'so wait until all the IA are calculated
                'If flagCheckDirectIA Then
                'End If


                Dim i As Integer
                Dim rescaSqft As Double = -1.0
                Dim rescaEru As Double = -1.0

                Dim rescaIaLayerName As String = ""

                Dim pdataset As IDataset
                pdataset = premiseLayer
                Dim iaTierLevel As ITable = MapUtil.GetTableFromWS(CType(pdataset.Workspace, IFeatureWorkspace), IAISToolSetting.GetParameterValue("TABLE_IATIERLEVEL"))

                Dim eruMessage As String = "PEXUID" & vbTab & "ERU" & vbTab & "Square footage"

                For i = 0 To iacfList.Count - 1
                    Dim iacfItem As IACF = iacfList.Item(i)
                    If iacfItem.Pexptyp = "RES" And flagHasResca And rescaSqft >= 0 Then
                        iacfItem.Iasqft = rescaSqft
                        iacfItem.IaLayerName = rescaIaLayerName
                        iacfItem.Iaeru = rescaEru
                    Else
                        Dim iacfSqft As Double = 0
                        If iacfItem.Pexptyp = "RES" And flagHasResca Then

                            iacfSqft = GetPremiseIA(iacfItem.Pexuid, ssl, pMap, iacfItem.IaLayerName, False)
                            rescaSqft = IAToolUtil.GetRoundIASQFT(GetPremiseIA(iacfItem.Pexuid, ssl, pMap, iacfItem.IaLayerName, False) / pexptypDict.Item("RES"), m_app)

                            rescaEru = MapUtil.GetRoundERU(MapUtil.GetBillableERU(iacfItem.Pexptyp, iacfSqft / typCount, iaTierLevel), m_app)
                            rescaIaLayerName = iacfItem.IaLayerName

                            iacfItem.Iasqft = rescaSqft
                            iacfItem.Iaeru = rescaEru

                            eruMessage = eruMessage & vbCrLf & "RESCA" & vbTab & iacfItem.Iaeru & vbTab & iacfItem.Iasqft

                        Else

                            If flagCheckDirectIA And Not iacfItem.Pexptyp = "RES" Then
                                iacfSqft = GetPremiseIA(iacfItem.Pexuid, ssl, pMap, iacfItem.IaLayerName, True)
                                If iacfSqft = 0 Then
                                    MsgBox("Premise " & iacfItem.Pexuid & " does not have direct IA assigment.")
                                    Return
                                End If
                            Else
                                iacfSqft = GetPremiseIA(iacfItem.Pexuid, ssl, pMap, iacfItem.IaLayerName, False)
                            End If

                            iacfItem.Iasqft = iacfSqft
                            iacfItem.Iaeru = MapUtil.GetRoundERU(MapUtil.GetBillableERU(iacfItem.Pexptyp, iacfSqft, iaTierLevel), m_app)

                            eruMessage = eruMessage & vbCrLf & iacfItem.Pexuid & vbTab & iacfItem.Iaeru & vbTab & iacfItem.Iasqft
                        End If
                    End If



                Next


                'Now report result and ask user to pick a date

                'If MsgBox("Charge has been recalculated for Parcel " & ssl & _
                '          vbCrLf & "ERU: " & _
                '          vbCrLf & "Square footage: " & _
                '          vbCrLf & "Update IACF table?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                '    Return
                'End If

                If MsgBox("Charge has been recalculated for Parcel " & ssl & _
                          vbCrLf & eruMessage & _
                          vbCrLf & "Update IACF table?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then
                    Return
                End If

                Dim pForm As FormPickDate
                pForm = New FormPickDate
                pForm.LabelDate.Text = "Effective Start Date"
                pForm.Text = "Effective Start Date"
                pForm.ShowDialog()

                If pForm.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                    Return
                End If

                IAISApplication.GetInstance.StartToolEditing()
                Try
                    Dim iacfTable As ITable = MapUtil.GetTableFromWS(CType(pdataset.Workspace, IFeatureWorkspace), IAISToolSetting.GetParameterValue("TABLE_IACF"))
                    Dim iacfRow As IRow
                    Dim pQFilter As IQueryFilter = New QueryFilter


                    For i = 0 To iacfList.Count - 1
                        Dim iacfItem As IACF = iacfList.Item(i)
                        pQFilter.WhereClause = "PEXUID=" & iacfItem.Pexuid & " AND EFFENDDT IS NULL"
                        Dim pCursor As ICursor = iacfTable.Search(pQFilter, False)
                        iacfRow = pCursor.NextRow
                        If Not iacfRow Is Nothing Then
                            iacfRow.Value(iacfRow.Fields.FindField("EFFENDDT")) = pForm.DateTimePicker1.Text
                            iacfRow.Store()
                        End If

                        iacfRow = iacfTable.CreateRow

                        iacfRow.Value(iacfRow.Fields.FindField("PEXUID")) = iacfItem.Pexuid
                        iacfRow.Value(iacfRow.Fields.FindField("IABILLERU")) = iacfItem.Iaeru
                        iacfRow.Value(iacfRow.Fields.FindField("IASQFT")) = iacfItem.Iasqft

                        iacfRow.Value(iacfRow.Fields.FindField("DBSTAMPDT")) = Now
                        iacfRow.Value(iacfRow.Fields.FindField("EFFSTARTDT")) = pForm.DateTimePicker1.Text
                        iacfRow.Value(iacfRow.Fields.FindField("IA_SOURCE")) = IAISToolSetting.GetParameterValue(iacfItem.IaLayerName)


                        iacfRow.Value(iacfRow.Fields.FindField("PARCEL_SOURCE")) = IAISToolSetting.GetParameterValue(pLayerParaName)

                        iacfRow.Store()

                    Next


                    IAISApplication.GetInstance.StopToolEditing("Calculate Charges")
                Catch ex As Exception
                    IAISApplication.GetInstance.AbortToolEditing()
                    MsgBox(ex.Message)
                    Return
                End Try




            End If



        Catch ex As Exception
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


    Private Function GetPremiseIA(ByVal pexuid As Integer, ByVal ssl As String, ByVal pMap As IMap, ByRef iaLayerParaName As String, ByVal bDirectAssign As Boolean) As Double

        'Search AppealOwnerPly
        Dim pQFilter As IQueryFilter = New QueryFilter

        Dim iaCursor As IFeatureCursor
        Dim iaLayer As IFeatureLayer
        Dim ia As IFeature
        Dim iasqft As Double = 0
        Dim iaeru As Double = 0

        Dim searchFlag As Boolean = True
        Dim searchSql As String
        If bDirectAssign Then
            searchSql = "PEXUID=" & pexuid
        Else
            searchSql = "(SSL='" & ssl & "' AND PEXUID IS NULL) OR PEXUID=" & pexuid
        End If


        If searchFlag Then
            pQFilter.WhereClause = searchSql
            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_APPEALIAASSIGNPLY"), pMap)
            If Not iaLayer Is Nothing AndAlso iaLayer.Visible Then
                iaCursor = iaLayer.Search(pQFilter, False)
                ia = iaCursor.NextFeature
                Do While Not ia Is Nothing
                    iasqft = iasqft + ia.Value(ia.Fields.FindField("IARSQF"))
                    ia = iaCursor.NextFeature

                    searchFlag = False
                Loop
                If Not searchFlag Then
                    iaLayerParaName = "LAYER_APPEALIAASSIGNPLY"
                End If
            End If
        End If

        If searchFlag Then
            pQFilter.WhereClause = searchSql
            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_REVIAASSIGNPLY"), pMap)
            If Not iaLayer Is Nothing AndAlso iaLayer.Visible Then
                iaCursor = iaLayer.Search(pQFilter, False)
                ia = iaCursor.NextFeature
                Do While Not ia Is Nothing
                    iasqft = iasqft + ia.Value(ia.Fields.FindField("IARSQF"))
                    ia = iaCursor.NextFeature

                    searchFlag = False
                Loop
                If Not searchFlag Then
                    iaLayerParaName = "LAYER_REVIAASSIGNPLY"
                End If
            End If
        End If

        If searchFlag Then
            pQFilter.WhereClause = searchSql
            iaLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_IAASSIGNPLY"), pMap)
            If Not iaLayer Is Nothing AndAlso iaLayer.Visible Then
                iaCursor = iaLayer.Search(pQFilter, False)
                ia = iaCursor.NextFeature
                Do While Not ia Is Nothing
                    iasqft = iasqft + ia.Value(ia.Fields.FindField("IARSQF"))
                    ia = iaCursor.NextFeature

                    searchFlag = False
                Loop
                If Not searchFlag Then
                    iaLayerParaName = "LAYER_IAASSIGNPLY"
                End If
            End If
        End If


        Return IAToolUtil.GetRoundIASQFT(iasqft, m_app)


    End Function


End Class


