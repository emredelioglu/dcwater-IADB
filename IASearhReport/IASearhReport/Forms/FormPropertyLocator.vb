Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display
Imports System.Drawing
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

Public Class FormPropertyLocator
    Private m_map As IMap
    Private m_search_result As Integer

    Private m_selectedFeatures As ITopologicalOperator

    Private m_viewEvent As IActiveViewEvents_Event

    Private m_app As IApplication

    Public WriteOnly Property PremiseApp() As IApplication
        Set(ByVal value As IApplication)
            m_app = value
        End Set
    End Property

    Public Property GMap() As IMap
        Get
            Return m_map
        End Get
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property

    Private Sub FindCol(ByVal square As String, ByVal suffix As String, ByVal lot As String)
        m_search_result = 0

        Dim colLayer As IFeatureLayer
        Dim pCursor As IFeatureCursor
        Dim pFeature As IFeature

        Dim SquarePlyFlag As Boolean = False
        Dim pSFilter As IQueryFilter = New QueryFilter

        square = UCase(Trim(square))
        suffix = UCase(Trim(suffix))
        lot = UCase(Trim(lot))

        If square = "" Then
            MsgBox("Square is required.", Title:="IA Search Report")
            m_search_result = 1
            Return
        End If

        If lot = "" Then
            SquarePlyFlag = True
        End If

        'square = square.PadLeft(4, "0")
        'If suffix <> "" Then
        '    suffix = square.PadRight(4, " ")
        'End If


        'If Len(lot) < 4 And lot <> "" Then
        '    lot = lot.PadLeft(4, "0")
        'End If


        If lot = "" Then
            If suffix = "" Then
                pSFilter.WhereClause = "SQUARE='" & square & "'"
            Else
                pSFilter.WhereClause = "SQUARE='" & square & "' AND SUFFIX='" & suffix & "'"
            End If
        Else
            If suffix = "" Then
                pSFilter.WhereClause = "SQUARE='" & square & "' AND LOT='" & lot & "'"
            Else
                pSFilter.WhereClause = "SQUARE='" & square & _
                                        "' AND SUFFIX='" & suffix & _
                                        "' AND LOT='" & lot & "'"
            End If
        End If

        Dim layerloadedFlag As Boolean = False
        If Not SquarePlyFlag Then
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
            If Not colLayer Is Nothing Then
                layerloadedFlag = True
                pCursor = colLayer.Search(pSFilter, False)
                pFeature = pCursor.NextFeature

                If Not pFeature Is Nothing Then
                    m_selectedFeatures = pFeature.Shape
                    MapUtil.SelectFeature(pFeature, colLayer)
                    Return
                End If
            End If
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), m_map)
            If Not colLayer Is Nothing Then
                layerloadedFlag = True
                pCursor = colLayer.Search(pSFilter, False)
                pFeature = pCursor.NextFeature

                If Not pFeature Is Nothing Then
                    m_selectedFeatures = pFeature.Shape
                    MapUtil.SelectFeature(pFeature, colLayer)
                    Return
                End If
            End If

            pSFilter.WhereClause = "SSL='" & MapUtil.GetSSL(square, suffix, lot) & "'"
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), m_map)
            If Not colLayer Is Nothing Then
                layerloadedFlag = True
                pCursor = colLayer.Search(pSFilter, False)
                pFeature = pCursor.NextFeature

                If Not pFeature Is Nothing Then
                    MapUtil.SelectFeature(pFeature, colLayer)
                    m_selectedFeatures = pFeature.Shape
                    Return
                End If
            End If
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), m_map)
            If Not colLayer Is Nothing Then
                layerloadedFlag = True
                pCursor = colLayer.Search(pSFilter, False)
                pFeature = pCursor.NextFeature

                If Not pFeature Is Nothing Then
                    MapUtil.SelectFeature(pFeature, colLayer)
                    m_selectedFeatures = pFeature.Shape
                    Return
                Else
                    m_search_result = 0
                End If
            End If

            If Not layerloadedFlag Then
                MsgBox("No property layer is loaded.", Title:="IA Search Report")
                m_search_result = -2
            End If

        Else
            colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SQUARESPLY"), m_map)
            If colLayer Is Nothing Then
                'Load it from the database
                'Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
                'If premiseLayer Is Nothing Then
                '    MsgBox("SquaresPly layer is not loaded and can't be loaded from database.")
                '    m_search_result = 1
                '    Return
                'End If
                'Dim pDataset As IDataset = premiseLayer

                colLayer = MapUtil.GetLayerFromWS(MapUtil.GetPremiseWS(m_map), _
                                                 IAISToolSetting.GetParameterValue("LAYER_SQUARESPLY"))
                If colLayer Is Nothing Then
                    MsgBox("SquaresPly layer is not loaded and can't be loaded from database.", Title:="IA Search Report")
                    m_search_result = 1
                    Return
                End If
            End If

            pCursor = colLayer.Search(pSFilter, False)
            pFeature = pCursor.NextFeature

            Dim pGeometryBag As IGeometryCollection = New GeometryBag
            Do While Not pFeature Is Nothing
                If pGeometryBag.GeometryCount = 0 Then
                    MapUtil.SelectFeature(pFeature, colLayer)
                Else
                    MapUtil.SelectFeature(pFeature, colLayer, False)
                End If

                pGeometryBag.AddGeometry(pFeature.ShapeCopy)
                pFeature = pCursor.NextFeature
            Loop

            If pGeometryBag.GeometryCount > 0 Then
                m_selectedFeatures = New Polygon
                m_selectedFeatures.ConstructUnion(pGeometryBag)
            Else
                m_search_result = -1
            End If

        End If


    End Sub

    Private Sub flashFeature(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase)

        If Not m_selectedFeatures Is Nothing AndAlso phase = esriViewDrawPhase.esriViewForeground Then
            MapUtil.FlashGeometry(m_selectedFeatures, m_map)
            m_selectedFeatures = Nothing
        End If
    End Sub
    Private Sub btnZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoom.Click
        Try
            m_selectedFeatures = Nothing

            FindCol(TextSquare.Text, TextSuffix.Text, TextLot.Text)
            If m_selectedFeatures Is Nothing Then
                If m_search_result = 0 Then
                    MsgBox("Can't find property with given SSL.", Title:="IA Search Report")
                ElseIf m_search_result = -1 Then
                    MsgBox("Can't find square with given value.", Title:="IA Search Report")
                End If

                Return
            End If

            IAISApplication.GetInstance().FlashFeature(m_selectedFeatures)

            MapUtil.ZoomToGeometry(m_selectedFeatures, m_map)
        Catch ex As Exception
            MsgBox("Error: " & ex.Source & vbCrLf & ex.Message, Title:="IA Search Report")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click
        Try
            m_selectedFeatures = Nothing

            FindCol(TextSquare.Text, TextSuffix.Text, TextLot.Text)
            If m_selectedFeatures Is Nothing Then
                If m_search_result = 0 Then
                    MsgBox("Can't find property with given SSL.", Title:="IA Search Report")
                    Return
                ElseIf m_search_result = -1 Then
                    MsgBox("Can't find square with given value.", Title:="IA Search Report")
                    Return
                End If
            End If

            '**if expression added by amanda please QC
            If TextSquare.Text <> "" Or TextSuffix.Text <> "" Or TextLot.Text <> "" Then
                IAISApplication.GetInstance().FlashFeature(m_selectedFeatures)
                MapUtil.PanToGeometry(m_selectedFeatures, m_map)
            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.Source & vbCrLf & ex.Message, Title:="IA Search Report")
        End Try
    End Sub

    Private Sub TextSquare_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextSquare.TextChanged
        If Trim(TextSquare.Text) <> "" Then
            LabelSquare.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
        Else
            LabelSquare.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        End If
    End Sub

    Private Sub TextSuffix_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextSuffix.TextChanged
        If Trim(TextSuffix.Text) <> "" Then
            LabelSuffix.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
        Else
            LabelSuffix.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        End If
    End Sub

    Private Sub TextLot_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextLot.TextChanged
        If Trim(TextLot.Text) <> "" Then
            LabelLot.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
        Else
            LabelLot.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        TextSquare.Text = ""
        TextSuffix.Text = ""
        TextLot.Text = ""
    End Sub


    Private Sub FormPremiseLocator_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Tab Then
            e.SuppressKeyPress = True
            If e.Modifiers = Keys.Shift Then
                Me.ProcessTabKey(False)
            Else
                Me.ProcessTabKey(True)
            End If
        End If
    End Sub

    'Private Sub FormPropertyLocator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    Dim pDoc As IMxDocument = m_app.Document
    '    m_viewEvent = pDoc.ActiveView
    '    AddHandler m_viewEvent.AfterDraw, AddressOf flashFeature
    'End Sub
End Class
