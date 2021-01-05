Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor

Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<CLSCompliant(False)> _
<ComClass(IABuldSplit.ClassId, IABuldSplit.InterfaceId, IABuldSplit.EventsId)> _
Public Class IABuldSplit
    Inherits BaseTool

    Private m_app As IApplication
    Private m_doc As IMxDocument
    Private m_map As IMap

    Private m_Editor As IEditor


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d71a3fa9-409e-46a0-9d9f-c20476cc5189"
    Public Const InterfaceId As String = "c6bde655-48ac-4b2a-a6c7-9f2f94507d1b"
    Public Const EventsId As String = "874c377d-472d-4980-b362-91e968e877f7"
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
            MyBase.m_bitmap = New System.Drawing.Bitmap(GetType(IABuldAssignment).Assembly.GetManifestResourceStream("IAIS.ia_building_split.bmp"))
        Catch
            MyBase.m_bitmap = Nothing
        End Try

        MyBase.m_caption = "Building Split"
        MyBase.m_category = "IAIS Tools"
        MyBase.m_message = "Building Split"
        MyBase.m_name = "Building Split"
        MyBase.m_toolTip = "Building Split"

    End Sub

    Public Overrides Sub OnClick()
        m_map = m_doc.FocusMap

        Dim pSel As IFeatureSelection
        Dim scratchLayer As IFeatureLayer
        Dim pScratchIA As IFeature = Nothing

        Dim pUid As New UID

        IAISApplication.GetInstance.StartToolEditing()
        Try

            Dim bldgLayer As IFeatureLayer
            Dim pFeatCursor As IFeatureCursor


            Dim sourceLayer As ILayer = IAToolUtil.GetSourceLayer(m_app)
            If sourceLayer Is Nothing Then
                Return
            End If

            scratchLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SCRATCHPLY"), m_map)
            If scratchLayer Is Nothing Then
                Throw New Exception("ScratchLayer is not loaded.")
            End If

            If TypeOf sourceLayer Is GroupLayer Then
                'Copy feature from building to Scratch

                bldgLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_BLDGPLY"), m_map)

                pSel = bldgLayer
                pSel.SelectionSet.Search(Nothing, False, pFeatCursor)

                Dim pBldg As IFeature
                pBldg = pFeatCursor.NextFeature

                'Clear all the selections

                If Not MapUtil.GetOverlapShape(pBldg.Shape, scratchLayer) Is Nothing Then
                    Throw New Exception("Topology check failed. This assignment will create a overlap in scratch layer.")
                End If

                m_map.ClearSelection()

                'ScratchIAPly
                Dim pEditLayers As IEditLayers
                pEditLayers = m_Editor
                pEditLayers.SetCurrentLayer(scratchLayer, 0)


                pScratchIA = scratchLayer.FeatureClass.CreateFeature

                pScratchIA.Value(pScratchIA.Fields.FindField("FEATURETYPE")) = 1
                pScratchIA.Value(pScratchIA.Fields.FindField("SOURCETYPE")) = 3
                pScratchIA.Value(pScratchIA.Fields.FindField("FEATUREOID")) = MapUtil.GetValue(pBldg, "GIS_ID")
                pScratchIA.Shape = pBldg.ShapeCopy

                pScratchIA.Store()

                MapUtil.SelectFeature(pScratchIA, scratchLayer)

                pSel = scratchLayer
                pSel.SelectionChanged()

            Else



                pSel = scratchLayer
                pSel.SelectionSet.Search(Nothing, False, pFeatCursor)
                pScratchIA = pFeatCursor.NextFeature

                m_map.ClearSelection()
                MapUtil.SelectFeature(pScratchIA, scratchLayer)

                pSel.SelectionChanged()
            End If


            IAISApplication.GetInstance.StopToolEditing("Split Polygon")

        Catch ex As Exception
            IAISApplication.GetInstance.AbortToolEditing()
            MsgBox(ex.Message)
            Return
        End Try

        ' If Not pScratchIA Is Nothing Then
        '        MapUtil.SelectFeature(pScratchIA, scratchLayer)
        'End If

        If m_Editor.CurrentTask.Name <> "Cut Polygon Features" Then
            Dim i As Integer
            Dim edTask As IEditTask

            For i = 0 To m_Editor.TaskCount - 1
                edTask = m_Editor.Task(i)
                If edTask.Name = "Cut Polygon Features" Then
                    m_Editor.CurrentTask = edTask
                    Exit For
                End If
            Next i

            If i = m_Editor.TaskCount Then
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox("Can't find cut polygon edit task")
                Return
            End If
        Else
            Dim i As Integer
            Dim edTask As IEditTask

            For i = 0 To m_Editor.TaskCount - 1
                edTask = m_Editor.Task(i)
                If edTask.Name = "Create New Feature" Then
                    m_Editor.CurrentTask = edTask
                    Exit For
                End If
            Next i

            System.Threading.Thread.Sleep(500)

            For i = 0 To m_Editor.TaskCount - 1
                edTask = m_Editor.Task(i)
                If edTask.Name = "Cut Polygon Features" Then
                    m_Editor.CurrentTask = edTask
                    Exit For
                End If
            Next i



        End If



        pUid.Value = "{B479F48A-199D-11D1-9646-0000F8037368}"
        Dim commandItem As ICommandItem
        commandItem = m_app.Document.CommandBars.Find(pUid)
        If Not commandItem Is Nothing Then
            commandItem.Execute()
        End If

        m_doc.ActiveView.Refresh()

    End Sub

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

                Dim sourceLayer As ILayer = IAToolUtil.GetSourceLayer(m_app)
                If sourceLayer Is Nothing Then
                    Return False
                End If

                Dim pLayer As IFeatureLayer
                Dim pFeatSel As IFeatureSelection
                If TypeOf sourceLayer Is GroupLayer Then
                    pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_BLDGPLY"), pMap)

                    If pLayer Is Nothing Then
                        Return False
                    End If

                    pFeatSel = pLayer
                    Return (pFeatSel.SelectionSet.Count = 1)

                Else
                    pLayer = sourceLayer
                    pFeatSel = pLayer
                    Return (pFeatSel.SelectionSet.Count = 1)
                End If

            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property 'Unreg
End Class


