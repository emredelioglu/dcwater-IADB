Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Location
Imports ESRI.ArcGIS.DataSourcesGDB

<CLSCompliant(False)> _
Public Class MapUtil
    Public Const FEATURETYPE_BUILDING As Integer = 1


    Public Shared Function GetOwnerName(ByVal fullname As String) As String
        Dim iLoc As Integer
        iLoc = InStr(fullname, ".")
        If iLoc <= 0 Then
            GetOwnerName = ""
        Else
            GetOwnerName = UCase(Mid(fullname, 1, iLoc - 1))
        End If

        Exit Function
    End Function



    Public Shared Function GetTableName(ByVal fullname As String) As String
        Dim iLoc As Integer
        iLoc = InStrRev(fullname, ".")
        If iLoc <= 0 Then
            GetTableName = UCase(fullname)
        Else
            GetTableName = UCase(Mid(fullname, iLoc + 1))
        End If

        Exit Function
    End Function

    Public Shared Function GetGroupLayer(ByVal grpLayerName As String, ByRef pMap As IMap) As ILayer
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1

            If pMap.Layer(i).Name = grpLayerName AndAlso TypeOf pMap.Layer(i) Is IGroupLayer Then
                Return pMap.Layer(i)
            End If
        Next i

        Return Nothing
    End Function

    Public Shared Function GetLayerByName(ByVal tablename As String, ByVal pLayer As ILayer, Optional ByVal IsLayerName As Boolean = False) As ILayer
        On Error GoTo ErrorHandler

        If Not TypeOf pLayer Is IGroupLayer Then
            If TypeOf pLayer Is IFeatureLayer Then
                If IsLayerName Then
                    If UCase(pLayer.Name) = UCase(tablename) Then
                        GetLayerByName = pLayer
                        Exit Function
                    End If
                Else
                    Dim pFeatureLayer As IFeatureLayer
                    pFeatureLayer = pLayer
                    Dim pDataset As IDataset
                    pDataset = pFeatureLayer.FeatureClass
                    If GetTableName(pDataset.BrowseName) = UCase(tablename) Then
                        GetLayerByName = pLayer
                        Exit Function
                    End If
                End If
            End If
            GetLayerByName = Nothing
            Exit Function
        End If

        Dim pCompositeLayer As ICompositeLayer
        Dim i As Integer
        If TypeOf pLayer Is IGroupLayer Then
            pCompositeLayer = pLayer
            For i = 0 To pCompositeLayer.Count - 1
                GetLayerByName = GetLayerByName(tablename, pCompositeLayer.Layer(i), IsLayerName)
                If Not GetLayerByName Is Nothing Then
                    Exit Function
                End If
            Next i
        End If

        GetLayerByName = Nothing

        Exit Function
ErrorHandler:
        MsgBox("GetLayerByName " & Err.Description)
        Err.Clear()
    End Function


    Public Shared Function GetLayerByTableName(ByVal tablename As String, ByVal pMap As IMap, Optional ByVal IsLayerName As Boolean = False) As ILayer
        GetLayerByTableName = Nothing
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1
            GetLayerByTableName = GetLayerByName(tablename, pMap.Layer(i), IsLayerName)
            If Not GetLayerByTableName Is Nothing Then
                Exit Function
            End If
        Next i
    End Function

    Public Shared Sub PrintDataSet(ByVal ds As DataSet)

        Dim colIndex As Integer
        ' Print out all tables and their columns
        For Each table As DataTable In ds.Tables
            Debug.Print("TABLE '{0}'", table.TableName)
            Debug.Print("Total # of rows: {0}", table.Rows.Count)
            Debug.Print("---------------------------------------------------------------")

            'For Each column As DataColumn In table.Columns
            '    Debug.Print("- {0} ({1})", column.ColumnName, column.DataType.ToString())
            'Next  ' For Each column

            For Each row As DataRow In table.Rows
                For colIndex = 0 To table.Columns.Count - 1
                    'row.Item(0).ToString()
                    'For Each column As DataColumn In table.Columns
                    Debug.Print("- {0} ({1}) ({2})", table.Columns.Item(colIndex).ColumnName, table.Columns.Item(colIndex).DataType.ToString(), row.Item(colIndex).ToString())
                Next  ' For Each column
            Next

            Debug.Print(System.Environment.NewLine)
        Next  ' For Each table

        ' Print out table relations
        For Each relation As DataRelation In ds.Relations
            Debug.Print("RELATION: {0}", relation.RelationName)
            Debug.Print("---------------------------------------------------------------")
            Debug.Print("Parent: {0}", relation.ParentTable.TableName)
            Debug.Print("Child: {0}", relation.ChildTable.TableName)
            Debug.Print(System.Environment.NewLine)
        Next  ' For Each relation

    End Sub

    Public Shared Function GetValueFromDataSet(ByVal ds As DataSet, ByVal fieldName As String) As String
        If ds.Tables.Count < 1 Then
            Return ""
        End If

        Dim table As DataTable = ds.Tables.Item(0)
        Dim colIndex As Integer
        If table.Rows.Count < 1 Then
            Return ""
        End If

        For colIndex = 0 To table.Columns.Count - 1
            If LCase(table.Columns.Item(colIndex).ColumnName) = LCase(fieldName) Then
                Return table.Rows.Item(0).Item(colIndex).ToString
            End If
        Next

        Return ""
    End Function

    Public Shared Function GetConfidenceLevel(ByVal ds As DataSet) As Double
        If ds.Tables.Count < 1 Then
            Return -1
        End If

        Dim table As DataTable = ds.Tables.Item(0)
        Dim colIndex As Integer
        If table.Rows.Count < 1 Then
            Return -1
        End If

        For colIndex = 0 To table.Columns.Count - 1
            If LCase(table.Columns.Item(colIndex).ColumnName) = "confidencelevel" Then
                Return CDbl(table.Rows.Item(0).Item(colIndex).ToString)
            End If
        Next

        Return -1

    End Function


    Public Shared Sub PopulateDomain(ByRef pCombo As ComboBox, ByVal pCodeDomain As ICodedValueDomain, ByVal allowNull As Boolean, Optional ByVal value As Object = VariantType.Null)

        Dim cindex As Long
        pCombo.Items.Clear()
        If allowNull Then
            pCombo.Items.Add("")
            If value = "" Then
                pCombo.SelectedIndex = 0
            End If
        End If

        For cindex = 0 To pCodeDomain.CodeCount - 1
            pCombo.Items.Add(pCodeDomain.Name(cindex))
            If value <> "" Then
                If value = pCodeDomain.Value(cindex) Then
                    pCombo.SelectedIndex = pCombo.Items.Count - 1
                End If
            End If
        Next

    End Sub

    Public Shared Function GetValue(ByVal pFeature As IFeature, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Return GetRowValue(CType(pFeature, IRow), col, getDomainValue)
    End Function


    Public Shared Function GetRowValue(ByVal pRow As IRow, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Dim fIndex As Integer
        fIndex = pRow.Fields.FindField(col)
        If IsDBNull(pRow.Value(fIndex)) Then
            Return ""
        Else
            If Not getDomainValue Then
                Return pRow.Value(fIndex)
            Else
                Dim pDomain As IDomain
                pDomain = pRow.Fields.Field(fIndex).Domain

                If Not pDomain Is Nothing AndAlso TypeOf pDomain Is ICodedValueDomain Then
                    Dim pCodeDomain As ICodedValueDomain = pDomain

                    Dim cindex As Long
                    For cindex = 0 To pCodeDomain.CodeCount - 1
                        If pCodeDomain.Value(cindex) = pRow.Value(fIndex) Then
                            Return pCodeDomain.Name(cindex)
                        End If
                    Next
                    Return pRow.Value(fIndex)
                Else
                    Return pRow.Value(fIndex)
                End If
            End If
        End If
    End Function

    Public Shared Sub SetComboIndex(ByRef pCombo As ComboBox, ByVal value As String)
        'If value = "" Then
        '    pCombo.SelectedIndex = -1
        '    Exit Sub
        'End If
        pCombo.SelectedIndex = pCombo.FindString(value)

        If pCombo.SelectedIndex = -1 Then
            pCombo.Text = value
        End If
    End Sub


    Public Shared Function SetFeatureValue(ByVal pFeature As IFeature, ByVal fIndex As Integer, ByVal fValue As String, Optional ByVal valueOnly As Boolean = False) As Boolean
        'Try

        If fIndex <= 0 Then
            SetFeatureValue = False
            Exit Function
        End If

        Dim pField As IField

        Dim pDomain As IDomain
        Dim pCodeDomain As ICodedValueDomain

        pField = pFeature.Fields.Field(fIndex)

        pDomain = pField.Domain

        If Not pDomain Is Nothing And Trim(fValue) <> "" Then
            'GetDomainCode
            If Not valueOnly And pDomain.Type = esriDomainType.esriDTCodedValue Then
                pCodeDomain = pDomain
                If IsDBNull(GetDomainCode(pCodeDomain, fValue)) Then
                    If Not pField.IsNullable Then
                        If pField.Type = esriFieldType.esriFieldTypeString Then
                            pFeature.Value(fIndex) = ""
                        Else
                            pFeature.Value(fIndex) = 0
                        End If

                    End If
                Else
                    pFeature.Value(fIndex) = GetDomainCode(pCodeDomain, fValue)
                End If
            Else
                pFeature.Value(fIndex) = fValue
            End If
        Else
            If Trim(fValue) = "" Then
                If Not IsDBNull(pFeature.Value(fIndex)) _
                    AndAlso CStr(pFeature.Value(fIndex)) <> "" _
                    AndAlso pField.IsNullable Then
                    pFeature.Value(fIndex) = System.DBNull.Value
                End If
            ElseIf pField.Type = esriFieldType.esriFieldTypeDate Then
                pFeature.Value(fIndex) = CDate(fValue)
            ElseIf pField.Type = esriFieldType.esriFieldTypeDouble Then
                pFeature.Value(fIndex) = CDbl(fValue)
            ElseIf pField.Type = esriFieldType.esriFieldTypeInteger Then
                pFeature.Value(fIndex) = CLng(fValue)
            ElseIf pField.Type = esriFieldType.esriFieldTypeSingle Then
                pFeature.Value(fIndex) = CSng(fValue)
            ElseIf pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
                pFeature.Value(fIndex) = CInt(fValue)
            ElseIf pField.Type = esriFieldType.esriFieldTypeString Then
                pFeature.Value(fIndex) = fValue
            Else
                pFeature.Value(fIndex) = fValue
            End If
        End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Function

    Public Shared Function GetDomainCode(ByVal pCodeDomain As ICodedValueDomain, ByVal value As String) As Object

        GetDomainCode = Nothing
        If IsDBNull(value) Or value = "" Then
            Exit Function
        End If

        Dim cindex As Long
        For cindex = 0 To pCodeDomain.CodeCount - 1
            If UCase(pCodeDomain.Name(cindex)) = UCase(value) Then
                GetDomainCode = pCodeDomain.Value(cindex)
                Exit For
            End If
        Next

    End Function

    Public Shared Function GetPlyArea(ByVal inpoly As IPolygon) As Double
        If inpoly Is Nothing Then
            GetPlyArea = 0.0#
        Else
            Dim parea As IArea
            parea = inpoly
            GetPlyArea = parea.Area
        End If
    End Function




    Public Function GetIaFeatureType(ByVal tableName As String) As Integer

        tableName = GetTableName(tableName)

        If UCase(tableName) = "BLDGPLY" Then
            Return 1
        ElseIf UCase(tableName) = "SIDEWALKPLY" Then
            Return 2
        ElseIf UCase(tableName) = "STRUCTPLY" Then
            Return 3
        ElseIf UCase(tableName) = "ROADPLY" Then
            Return 4
        ElseIf UCase(tableName) = "RECOUTDOORPLY" Then
            Return 5
        ElseIf UCase(tableName) = "SWMPOOLPLY" Then
            Return 6
        ElseIf UCase(tableName) = "BRGTUNPLY" Then
            Return 7
        Else
            Return 0
        End If

    End Function

    Public Shared Function GetSelectedCol(ByVal pMap As IMap) As IFeature

        Dim pColLayer As IFeatureLayer
        Dim pColSel As IFeatureSelection
        Dim pColCursor As IFeatureCursor

        pColLayer = MapUtil.GetLayerByTableName("OwnerPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing AndAlso pColSel.SelectionSet.Count = 1 Then
            pColSel.SelectionSet.Search(Nothing, False, pColCursor)
            Return pColCursor.NextFeature
        End If

        pColLayer = MapUtil.GetLayerByTableName("OwnerGapPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing AndAlso pColSel.SelectionSet.Count = 1 Then
            pColSel.SelectionSet.Search(Nothing, False, pColCursor)
            Return pColCursor.NextFeature
        End If

        pColLayer = MapUtil.GetLayerByTableName("AppealOwnerPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing AndAlso pColSel.SelectionSet.Count = 1 Then
            pColSel.SelectionSet.Search(Nothing, False, pColCursor)
            Return pColCursor.NextFeature
        End If

        pColLayer = MapUtil.GetLayerByTableName("RevOwnerPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing AndAlso pColSel.SelectionSet.Count = 1 Then
            pColSel.SelectionSet.Search(Nothing, False, pColCursor)
            Return pColCursor.NextFeature
        End If

        Return Nothing


    End Function

    Public Shared Function GetSelectedColCount(ByVal pMap As IMap) As Integer

        Dim pColLayer As IFeatureLayer
        Dim pColSel As IFeatureSelection

        pColLayer = MapUtil.GetLayerByTableName("OwnerPly", pMap)
        pColSel = pColLayer

        Dim fCount As Integer = 0

        If Not pColLayer Is Nothing Then
            fCount = fCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("OwnerGapPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing Then
            fCount = fCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("AppealOwnerPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing Then
            fCount = fCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("RevOwnerPly", pMap)
        pColSel = pColLayer

        If Not pColLayer Is Nothing Then
            fCount = fCount + pColSel.SelectionSet.Count
        End If

        Return fCount

    End Function

    Public Shared Function IsOneColSelected(ByVal pMap As IMap) As Boolean
        Dim selCount As Integer = 0

        Dim pColLayer As IFeatureLayer
        Dim pColSel As IFeatureSelection

        pColLayer = MapUtil.GetLayerByTableName("OwnerPly", pMap)
        If Not pColLayer Is Nothing Then
            pColSel = pColLayer
            selCount = selCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("OwnerGapPly", pMap)
        If Not pColLayer Is Nothing Then
            pColSel = pColLayer
            selCount = selCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("AppealOwnerPly", pMap)
        If Not pColLayer Is Nothing Then
            pColSel = pColLayer
            selCount = selCount + pColSel.SelectionSet.Count
        End If

        pColLayer = MapUtil.GetLayerByTableName("RevOwnerPly", pMap)
        If Not pColLayer Is Nothing Then
            pColSel = pColLayer
            selCount = selCount + pColSel.SelectionSet.Count
        End If

        Return (selCount = 1)

    End Function

    Public Shared Function GetTable(ByVal tableName As String, ByVal pMap As IMap) As ITable
        Dim tables As IStandaloneTableCollection
        tables = pMap
        Dim i As Integer
        Dim pTable As ITable

        For i = 0 To tables.StandaloneTableCount - 1
            pTable = tables.StandaloneTable(i)
            If tables.StandaloneTable(i).Name = tableName Then
                Return pTable
            End If
        Next i

        Return Nothing
    End Function

    Public Shared Function SelectFeature(ByRef pFeature As IFeature, ByRef pLayer As IFeatureLayer, Optional ByVal newSelection As Boolean = True) As Boolean
        Dim pSel As IFeatureSelection
        pSel = pLayer
        If newSelection Then
            pSel.Clear()
        End If

        pSel.Add(pFeature)
    End Function


    Public Shared Function ZoomToFeature(ByRef pFeature As IFeature, ByRef pMap As IMap) As Boolean
        Dim minScale As Double = pMap.ReferenceScale
        Dim pNewMapExtent As IEnvelope
        Dim pActiveView As IActiveView
        Dim pDisplayTransform As IDisplayTransformation

        Try
            Dim pFeatClass As IFeatureClass = pFeature.Class

            pActiveView = pMap
            pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation

            If pFeature.Extent.Width = 0 Then
                'This is a point
                pNewMapExtent = pDisplayTransform.VisibleBounds

                Dim pCenterPoint As IPoint
                pCenterPoint = New Point

                pCenterPoint.X = pFeature.Extent.XMin
                pCenterPoint.Y = pFeature.Extent.YMin

                pNewMapExtent.CenterAt(pCenterPoint)
                pDisplayTransform.VisibleBounds = pNewMapExtent
                pMap.MapScale = minScale
                pActiveView.Refresh()
                Return True
            Else
                pNewMapExtent = pFeature.Extent
                pNewMapExtent.Expand(1.05, 1.05, True)
                pDisplayTransform.VisibleBounds = pNewMapExtent

                'If pMap.MapScale < minScale Then
                '    pMap.MapScale = minScale
                'End If

                pActiveView.Refresh()

                Return True
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function


    Public Shared Function PanToFeature(ByRef pFeature As IFeature, ByRef pMap As IMap) As Boolean
        Dim pNewMapExtent As IEnvelope
        Dim pActiveView As IActiveView
        Dim pDisplayTransform As IDisplayTransformation
        Try
            pActiveView = pMap
            pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation

            pNewMapExtent = pDisplayTransform.VisibleBounds

            Dim pCenterPoint As IPoint
            pCenterPoint = New Point

            pCenterPoint.X = (pFeature.Extent.XMin + pFeature.Extent.XMax) / 2
            pCenterPoint.Y = (pFeature.Extent.YMin + pFeature.Extent.YMax) / 2

            pNewMapExtent.CenterAt(pCenterPoint)
            pDisplayTransform.VisibleBounds = pNewMapExtent

            pActiveView.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function

    Public Shared Sub FlashFeature(ByVal pFeature As IFeature, ByVal pMap As IMap)

        Dim featureIdObj As IFeatureIdentifyObj
        featureIdObj = New FeatureIdentifyObject
        featureIdObj.Feature = pFeature

        Dim pActiveView As IActiveView
        pActiveView = pMap

        Dim idObj As IIdentifyObj
        idObj = featureIdObj
        idObj.Flash(pActiveView.ScreenDisplay)

        System.Windows.Forms.Application.DoEvents()
    End Sub


    Public Shared Function GetTableFromWS(ByRef pFWS As IFeatureWorkspace, ByVal tableName As String) As ITable
        Return pFWS.OpenTable(tableName)
    End Function

    Public Shared Function GetAddressAtPoint(ByVal pPt As IPoint, ByVal locatorName As String, ByRef pWorkspace As IWorkspace) As String

        Dim pLocatorManager As ILocatorManager2
        Dim pReverseGeocoding As IReverseGeocoding
        Dim pAddressGeocoding As IAddressGeocoding

        Dim pMatchFields As IFields
        Dim pShapeField As IField
        Dim pReverseGeocodingProperties As IReverseGeocodingProperties
        Dim pAddressProperties As IPropertySet

        Dim pAddressInputs As IAddressInputs
        Dim pAddressFields As IFields
        Dim pAddressField As IField

        pLocatorManager = New LocatorManager
        Dim pLWS As ILocatorWorkspace
        pLWS = pLocatorManager.GetLocatorWorkspace(pWorkspace)
        Dim pLocator As ILocator
        pLocator = pLWS.GetLocator(locatorName)
        pReverseGeocoding = pLocator


        '+++ create a Point at which to find the address
        pAddressGeocoding = pReverseGeocoding
        pMatchFields = pAddressGeocoding.MatchFields
        pShapeField = pMatchFields.Field(pMatchFields.FindField("Shape"))

        '+++ set the search tolerance for reverse geocoding
        pReverseGeocodingProperties = pReverseGeocoding
        pReverseGeocodingProperties.SearchDistance = 100
        pReverseGeocodingProperties.SearchDistanceUnits = esriUnits.esriMeters

        '+++ find the address nearest the Point
        pAddressProperties = pReverseGeocoding.ReverseGeocode(pPt, False)
        '+++ print the address properties
        pAddressInputs = pReverseGeocoding
        pAddressFields = pAddressInputs.AddressFields

        If pAddressFields.FieldCount > 0 Then
            pAddressField = pAddressFields.Field(0)
            Return pAddressProperties.GetProperty(pAddressField.Name)
            Exit Function
        End If

        Return ""

    End Function

    Public Shared Function SetCurrentTool(ByVal strToolName As String, ByVal app As IApplication) As Boolean

        Dim pUid As New UID
        pUid.Value = "esriArcMapUI.SelectFeaturesTool"
        Dim pArcMapCmd As ICommandItem = app.Document.CommandBars.Find(pUid)
        If Not pArcMapCmd Is Nothing Then
            pArcMapCmd.Execute()
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function GetSSL(ByVal square As String, ByVal suffix As String, ByVal lot As String) As String
        square = square.PadLeft(4, "0")
        suffix = square.PadRight(4, " ")

        If Len(lot) < 4 Then
            lot = lot.PadLeft(4, "0")
        End If

        Return square & suffix & lot
    End Function

End Class


