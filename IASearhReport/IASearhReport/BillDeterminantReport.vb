Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase

Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Carto
Imports System.IO

<CLSCompliant(False)> _
<ComClass(BillDeterminantReport.ClassId, BillDeterminantReport.InterfaceId, BillDeterminantReport.EventsId)> _
<ProgId("IAIS.BillDeterminantReport")> _
Public Class BillDeterminantReport
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f1aeb93d-6612-42ae-953c-1af3b5dde8da"
    Public Const InterfaceId As String = "1702ca59-9c36-4e89-b04e-6cc11fb8bc37"
    Public Const EventsId As String = "5a6448d3-3f37-4a14-b082-252aec0d97ae"
#End Region
    Private m_app As IApplication
    Private m_doc As IMxDocument

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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(BillDeterminantReport).Assembly.GetManifestResourceStream("IAIS.search_bill_determinate.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "View Bill Determinant Report"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "View Bill Determinant Report"
        MyBase.m_name = "View Bill Determinant Report"
        MyBase.m_toolTip = "View Bill Determinant Report"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document
    End Sub

    Public Overrides Sub OnClick()
        Try
            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            Dim pSel As IFeatureSelection
            pSel = premiseLayer
            Dim pCursor As IFeatureCursor
            pSel.SelectionSet.Search(Nothing, False, pCursor)

            Dim pFeature As IFeature = pCursor.NextFeature
            If MapUtil.GetRowValue(pFeature, "IS_EXEMPT_IAB") = "Y" Then
                MsgBox("Selected premise is IA exempt", Title:="IA Search Report")
                Return
            End If
            Dim reportFilPath As String = ""
            Try
                reportFilPath = IAISToolSetting.GetParameterValue("PDF_MAP_PATH") & "\" & MapUtil.GetRowValue(pFeature, "PEXUID") & ".pdf"
                If Not File.Exists(reportFilPath) Then
                    MsgBox("Bill Determinant Report for Premise ID " & MapUtil.GetRowValue(pFeature, "PEXUID") & " is not found.", Title:="IA Search Report")
                    Return
                End If
                System.Diagnostics.Process.Start(reportFilPath)
            Catch ex As Exception
                MsgBox("Can't open report for Premise ID " & MapUtil.GetRowValue(pFeature, "PEXUID") & ". " & reportFilPath & vbCrLf & ex.Message, Title:="IA Search Report")
            End Try

        Catch ex As Exception
            MsgBox("Error: " & ex.Source & vbCrLf & ex.Message)
        End Try


    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If Not MapUtil.IsExtentionEnabled("IAIS.IASearchReportExt", m_app) Then
                Return False
            End If

            If Not IAISToolSetting.isInitilized Then
                Return False
            End If

            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap)

            If premiseLayer Is Nothing Then
                Return False
            End If

            Dim pSel As IFeatureSelection
            pSel = premiseLayer
            Return (pSel.SelectionSet.Count = 1)
        End Get
    End Property

End Class


