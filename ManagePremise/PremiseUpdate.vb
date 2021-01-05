Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.BaseClasses

Imports IAIS.Windows.UI
Imports ESRI.ArcGIS.Carto

<CLSCompliant(False)> _
<ComClass(PremiseUpdate.ClassId, PremiseUpdate.InterfaceId, PremiseUpdate.EventsId)> _
Public Class PremiseUpdate
    Inherits BaseCommand
    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_Editor As IEditor

    Private m_form As PremiseAttribute

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5b2b2f5a-c0a2-4404-bd9f-318f0ef23980"
    Public Const InterfaceId As String = "78abac6d-4698-4714-8e16-5ceb8b74c237"
    Public Const EventsId As String = "92b9e244-9951-4089-80aa-6fd38f0d7cb4"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(PremiseUpdate).Assembly.GetManifestResourceStream("IAIS.PremiseUpdate.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Premise Update"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Update premise point"
        MyBase.m_name = "IAIS Premise Update"
        MyBase.m_toolTip = "Update premise feature"
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
            If m_form IsNot Nothing AndAlso m_form.Visible Then
                MsgBox("You have not completed the previous premise point.")
                Return
            End If

            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeatureLayer
            premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", pMap)

            Dim pSel As IFeatureSelection
            pSel = premiseLayer
            Dim pCursor As IFeatureCursor
            pSel.SelectionSet.Search(Nothing, False, pCursor)

            Dim pPt As IFeature
            pPt = pCursor.NextFeature

            m_form = New PremiseAttribute
            m_form.PremiseMap = m_doc.FocusMap
            m_form.PremiseApp = m_app
            m_form.PremisePt = pPt

            m_form.EditMode = 2

            m_form.Show(New ModelessDialog(m_app.hWnd))

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If m_Editor Is Nothing Then
                Return False
            End If

            If m_Editor.EditState = esriEditState.esriStateNotEditing Then
                Return False
            End If

            Dim pMap As IMap = m_doc.FocusMap
            Dim premiseLayer As IFeaturelayer
            premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", pMap)

            Dim pSel As IFeatureselection
            pSel = premiseLayer
            Return (pSel.SelectionSet.Count = 1)

        End Get
    End Property


End Class


