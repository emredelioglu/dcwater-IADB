Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

Public Class IAToolUtil
    Public Shared Function GetTargetLayer(ByRef app As IApplication) As IFeatureLayer
        Dim pUid As New UID
        pUid.Value = "IAIS.IASourceListControl"
        Dim commandItem As ICommandItem
        commandItem = app.Document.CommandBars.Find(pUid)

        Dim sourceList As IASourceListControl
        sourceList = commandItem.Command

        If sourceList.ComboTargetList.SelectedItem = "" Then
            Return Nothing
        End If

        Dim pDoc As IMxDocument
        pDoc = app.Document
        Dim pLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_" & UCase(sourceList.ComboTargetList.SelectedItem)), pDoc.FocusMap)

        'If Not pLayer.Visible Then
        '    pLayer.Visible = True
        'End If

        Return pLayer

    End Function


    Public Shared Function GetSourceLayer(ByRef app As IApplication) As ILayer
        Dim pUid As New UID
        pUid.Value = "IAIS.IASourceListControl"
        Dim commandItem As ICommandItem
        commandItem = app.Document.CommandBars.Find(pUid)

        Dim sourceList As IASourceListControl
        sourceList = commandItem.Command

        If sourceList.ComboSourceList.SelectedItem = "" Then
            Return Nothing
        End If

        Dim pDoc As IMxDocument
        pDoc = app.Document

        Dim pLayer As IFeatureLayer


        If sourceList.ComboSourceList.SelectedItem = "Scratch" Then
            pLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SCRATCHPLY"), pDoc.FocusMap)
            'If Not pLayer.Visible Then
            '    pLayer.Visible = True
            'End If
            Return pLayer
        Else

            Dim iaGroupLayer As IGroupLayer = New GroupLayer
            Dim ialayer As IFeatureLayer

            ialayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_ROADPLY"), pDoc.FocusMap)
            If Not ialayer Is Nothing AndAlso ialayer.Visible Then
                iaGroupLayer.Add(ialayer)
            End If

            ialayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_RECOUTDOORPLY"), pDoc.FocusMap)
            If Not ialayer Is Nothing AndAlso ialayer.Visible Then
                iaGroupLayer.Add(ialayer)
            End If

            ialayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_STRUCTPLY"), pDoc.FocusMap)
            If Not ialayer Is Nothing AndAlso ialayer.Visible Then
                iaGroupLayer.Add(ialayer)
            End If

            ialayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SIDEWALKPLY"), pDoc.FocusMap)
            If Not ialayer Is Nothing AndAlso ialayer.Visible Then
                iaGroupLayer.Add(ialayer)
            End If

            ialayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_SWMPOOLPLY"), pDoc.FocusMap)
            If Not ialayer Is Nothing AndAlso ialayer.Visible Then
                iaGroupLayer.Add(ialayer)
            End If

            Return iaGroupLayer 'MapUtil.GetGroupLayer("DCGIS IA Layers", pDoc.FocusMap)
        End If
    End Function

    Public Shared Function GetIAType(ByVal tablename As String) As Integer
        Select Case LCase(tablename)
            Case "bldgply"
                Return 1
            Case "sidewalkply"
                Return 2
            Case "structply"
                Return 3
            Case "roadply"
                Return 4
            Case "recoutdoorply"
                Return 5
            Case "swmpoolply"
                Return 6
        End Select

        Return 0
    End Function

    Public Shared Function GetSSL(ByVal square As String, ByVal suffix As String, ByVal lot As String) As String
        If IsNumeric(square) Then
            square = square.PadLeft(4, "0")
        Else
            square = square.PadRight(4, " ")
        End If

        suffix = suffix.PadRight(4, " ")

        If Len(lot) < 4 Then
            lot = lot.PadLeft(4, "0")
        End If

        Return square & suffix & lot
    End Function

    Public Shared Function GetRoundIASQFT(ByVal iasqft As Double, ByRef app As IApplication) As Double
        Dim factor As Double = CDbl(IAISToolSetting.GetParameterValue("SQFT_ROUND_PRECISION"))
        Dim roundmethod As String = IAISToolSetting.GetParameterValue("IASQFT_ROUND_METHOD")
        If roundmethod = "UP" Then
            Return Math.Ceiling(iasqft / factor) * factor
        ElseIf roundmethod = "DOWN" Then
            Return Math.Floor(iasqft / factor) * factor
        Else
            Return Math.Round(iasqft / factor) * factor
        End If

    End Function




End Class
