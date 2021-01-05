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
Imports System.Drawing

<CLSCompliant(False)> _
<ComClass(PropertyLocator.ClassId, PropertyLocator.InterfaceId, PropertyLocator.EventsId)> _
<ProgId("IAIS.PropertyLocator")> _
Public Class PropertyLocator
    Inherits BaseCommand

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_form As FormPropertyLocator

    Private m_pIASearchReportExt As IASearchReportExt

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "19cec05b-0e4c-453a-9916-c43496497c8a"
    Public Const InterfaceId As String = "e258e1c9-c8a6-4a2d-8895-99ae710423af"
    Public Const EventsId As String = "32973244-7e6c-41f7-8f42-c4d7f7cdd021"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PropertyLocator).Assembly.GetManifestResourceStream("IAIS.search_property_locator.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Property Locator"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Property Locator"
        MyBase.m_name = "Property Locator"
        MyBase.m_toolTip = "Property Locator"
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        m_app = hook
        m_doc = m_app.Document

        m_pIASearchReportExt = MapUtil.GetExtention("IAIS.IASearchReportExt", m_app)
        If Not m_pIASearchReportExt Is Nothing Then
            AddHandler m_pIASearchReportExt.IASearchReportExtEvent, AddressOf onIASearchReportExtEvent
        End If

    End Sub

    Private Sub onIASearchReportExtEvent(ByVal strMsg As String)
        If strMsg = "disable" Then
            If Not m_form Is Nothing AndAlso m_form.Visible Then
                m_form.Close()
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()
        Try
            If Not m_form Is Nothing AndAlso m_form.Visible Then
                'Dim x As Integer = (m_form.Parent.Left + m_form.Parent.Right) / 2
                'Dim y As Integer = (m_form.Parent.Height - m_form.Parent.Height) / 2
                'm_form.Location = New Point(x, y)
                m_form.Activate()
            Else
                m_form = New FormPropertyLocator
                Dim pmap As IMap = m_doc.FocusMap
                m_form.GMap = pmap
                m_form.PremiseApp = m_app
                m_form.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
                m_form.Show(New ModelessDialog(m_app.hWnd))
            End If
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
            Dim colLayer As IFeatureLayer
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap)

            If Not colLayer Is Nothing Then
                Return True
            End If

            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), pMap)
            If Not colLayer Is Nothing Then
                Return True
            End If

            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), pMap)
            If Not colLayer Is Nothing Then
                Return True
            End If

            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), pMap)
            Return Not colLayer Is Nothing

        End Get
    End Property
End Class


