Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports System.Runtime.InteropServices

<ComClass(IASourceListControl.ClassId, IASourceListControl.InterfaceId, IASourceListControl.EventsId), _
 ProgId("IAIS.IASourceListControl")> _
Public Class IASourceListControl
    Implements ICommand
    Implements IToolControl

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6ACAE1AA-461E-47d6-9E32-9E2F9A55C339"
    Public Const InterfaceId As String = "6E5A27FF-D669-4d59-9025-51B9316C0DE6"
    Public Const EventsId As String = "0BA940CC-8A75-4ba8-8312-7455FC2D2B3A"
#End Region

    Private m_app As IApplication
    Private m_doc As IMxDocument

    Private m_completeNotify As ICompletionNotify

    Private m_docEvents As IDocumentEvents_Event
    Private m_pActiveViewEvents As IActiveViewEvents_Event


    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

#Region "IToolControl Members"

    Public ReadOnly Property Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get

        End Get
    End Property

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "Source"
        End Get
    End Property

    Public ReadOnly Property Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "IADB Tool"
        End Get
    End Property

    Public ReadOnly Property Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property Enabled1() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            Me.Refresh()
            Return True
        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return String.Empty
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "IA Source dropdown list"
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "IADB_ToolSourceControl"
        End Get
    End Property

    Public Sub OnClick1() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

    End Sub

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_app = TryCast(hook, IApplication)
        m_doc = m_app.Document

        Dim pMap As IMap
        pMap = m_doc.FocusMap

        'Dim player As ILayer
        'Dim i As Integer
        'For i = 0 To pMap.LayerCount - 1
        '    player = pMap.Layer(i)
        '    ComboSourceList.Items.Add(player.Name)
        '    ComboTargetList.Items.Add(player.Name)
        'Next

        'Dim appDocument As IDocument = m_app.Document
        'm_docEvents = CType(appDocument, IDocumentEvents_Event)
        'AddHandler m_docEvents.NewDocument, AddressOf onDocumentEvent
        'AddHandler m_docEvents.OpenDocument, AddressOf onDocumentEvent
        'AddHandler m_docEvents.MapsChanged, AddressOf onDocumentEvent

        'm_pActiveViewEvents = CType(m_doc.FocusMap, IActiveViewEvents_Event)
        'AddHandler m_pActiveViewEvents.ItemAdded, AddressOf RefreshLayersList
        'AddHandler m_pActiveViewEvents.ItemDeleted, AddressOf RefreshLayersList

    End Sub

    Public ReadOnly Property Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "IA Source"
        End Get
    End Property

    Public ReadOnly Property hWnd() As Integer Implements ESRI.ArcGIS.SystemUI.IToolControl.hWnd
        Get
            Return Me.Handle.ToInt32
        End Get
    End Property

    Public Function OnDrop(ByVal barType As ESRI.ArcGIS.SystemUI.esriCmdBarType) As Boolean Implements ESRI.ArcGIS.SystemUI.IToolControl.OnDrop
        Return True
    End Function
#End Region

    Public Sub OnFocus(ByVal complete As ESRI.ArcGIS.SystemUI.ICompletionNotify) Implements ESRI.ArcGIS.SystemUI.IToolControl.OnFocus
        m_completeNotify = complete
    End Sub

    Private Sub ComboSourceList_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboSourceList.DropDownClosed
        If m_completeNotify IsNot Nothing Then
            m_completeNotify.SetComplete()
        End If
    End Sub

    Sub RefreshLayersList()
        Dim source As String = ComboSourceList.SelectedItem
        Dim target As String = ComboTargetList.SelectedItem

        m_doc = m_app.Document
        Dim pMap As IMap
        pMap = m_doc.FocusMap

        ComboSourceList.Items.Clear()
        ComboTargetList.Items.Clear()

        Dim player As ILayer
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1
            player = pMap.Layer(i)
            ComboSourceList.Items.Add(player.Name)
            ComboTargetList.Items.Add(player.Name)
        Next

        ComboSourceList.SelectedItem = source
        ComboTargetList.SelectedItem = target
    End Sub

    Sub onDocumentEvent()
        RefreshLayersList()

        m_doc = m_app.Document

        m_pActiveViewEvents = CType(m_doc.FocusMap, IActiveViewEvents_Event)
        AddHandler m_pActiveViewEvents.ItemAdded, AddressOf RefreshLayersList
        AddHandler m_pActiveViewEvents.ItemDeleted, AddressOf RefreshLayersList

    End Sub


End Class
