Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Carto
Imports System.Collections.Generic
Imports IAIS.Windows.UI

<CLSCompliant(False)> _
<ComClass(PremiseSelMaster.ClassId, PremiseSelMaster.InterfaceId, PremiseSelMaster.EventsId)> _
Public Class PremiseSelMaster
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor

    Private m_form As FormSelectPremise
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "543ac4ca-1282-4cb4-8d70-a57fbfdd0fe0"
    Public Const InterfaceId As String = "b821ab59-7ee7-4ec1-a7d7-af3071e54897"
    Public Const EventsId As String = "b7d613ce-f215-4b38-a720-ff1e9423ca4c"
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

        Try
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PremiseSelMaster).Assembly.GetManifestResourceStream("IAIS.PremiseSelMaster.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Manage IA Relationship"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Manage IA Charge Relationship"
        MyBase.m_name = "Manage IA Relationship"
        MyBase.m_toolTip = "Manage IA Charge Relationship"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        Dim pUid As New UID
        pUid.Value = "esriEditor.Editor"
        m_Editor = m_app.FindExtensionByCLSID(pUid)


    End Sub

    Public Overrides Sub OnClick()
        'Get the selected premise point and pass it to the form
        Try


            Dim pPremiseLayer As IFeatureLayer
            pPremiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_doc.FocusMap)

            Dim premiseList As IList = New List(Of clsPremisePt)
            Dim premise As clsPremisePt

            Dim pSFilter As ISpatialFilter
            Dim pQFilter As IQueryFilter

            Dim pFeatCursor As IFeatureCursor
            Dim pt As IFeature

            Dim chargePexuid As Integer = 0

            Dim pCol As IFeature = MapUtil.GetSelectedCol(m_doc.FocusMap)
            If Not pCol Is Nothing Then
                'Search all the premise point in this parcel

                pSFilter = New SpatialFilter
                pSFilter.Geometry = pCol.Shape
                pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                pSFilter.WhereClause = "EXEMPT_IAB_REASON IS NULL OR EXEMPT_IAB_REASON <> 'PURGE'"

                pQFilter = New QueryFilter


                pFeatCursor = pPremiseLayer.Search(pSFilter, False)

                pt = pFeatCursor.NextFeature

                While Not pt Is Nothing
                    premise = New clsPremisePt

                    premise.pexuid = MapUtil.GetValue(pt, "PEXUID")
                    premise.pexprm = MapUtil.GetValue(pt, "PEXPRM")
                    premise.pexsad = MapUtil.GetPremiseAddress(pt)
                    premise.imperviousOnly = MapUtil.GetValue(pt, "IS_IMPERVIOUS_ONLY", True)
                    premise.exemptIAB = MapUtil.GetValue(pt, "IS_EXEMPT_IAB", True)
                    premise.exemptReason = MapUtil.GetValue(pt, "EXEMPT_IAB_REASON", True)

                    'check to see if it is a master premise

                    pQFilter.WhereClause = "MASTER_PEXUID=" & premise.pexuid
                    If pPremiseLayer.FeatureClass.FeatureCount(pQFilter) > 0 Then
                        premise.masterFlag = True
                        chargePexuid = premise.pexuid
                    End If

                    premiseList.Add(premise)

                    pt = pFeatCursor.NextFeature

                End While
            Else
                MsgBox("Np property feature is selected.")
                Return
            End If


            If premiseList.Count > 1 Then

                m_form = New FormSelectPremise
                m_form.SetPremisePts(premiseList)
                m_form.PremiseMap = m_doc.FocusMap
                m_form.PremiseApp = m_app
                m_form.ChargePexuid = chargePexuid
                'm_form.Show(New ModelessDialog(m_app.hWnd))
                IAISToolSetting.OpenForm(m_form)

            Else
                MsgBox("The selected property does not have multiple non-purged premise points.")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.PremiseExt", m_app) Then
                Return False
            End If

            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            If Not IAISApplication.GetInstance.IsInJTXApplication() Then
                Return False
            End If

            If Not IAISApplication.GetInstance.IsInEditingSession() Then
                Return False
            End If


            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            If premiseLayer Is Nothing Then
                Return False
            End If

            Return (MapUtil.GetSelectedColCount(m_doc.FocusMap) = 1)
        End Get
    End Property


End Class


