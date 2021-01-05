Option Strict Off

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
Imports Microsoft.VisualBasic.CompilerServices

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

    Public Shared Function GetLayerByName(ByVal tablename As String, ByVal pLayer As ILayer, Optional ByVal IsLayerName As Boolean = False, Optional ByVal remoteDB As Boolean = False) As ILayer
        On Error GoTo ErrorHandler

        If pLayer Is Nothing Or Not pLayer.Valid Then
            Return Nothing
        End If

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

                    If remoteDB And Not pDataset.Workspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
                        Return Nothing
                    End If

                    If tablename.IndexOf(".") > 0 Then
                        If UCase(pDataset.BrowseName) = UCase(tablename) Then
                            Return pLayer
                            Exit Function
                        End If
                    Else
                        If UCase(GetTableName(pDataset.BrowseName)) = UCase(tablename) Then
                            Return pLayer
                            Exit Function
                        End If
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
                GetLayerByName = GetLayerByName(tablename, pCompositeLayer.Layer(i), IsLayerName, remoteDB)
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


    Public Shared Function GetLayerByTableName(ByVal tablename As String, ByVal pMap As IMap, Optional ByVal IsLayerName As Boolean = False, Optional ByVal remoteDB As Boolean = False) As ILayer
        GetLayerByTableName = Nothing
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1
            GetLayerByTableName = GetLayerByName(tablename, pMap.Layer(i), IsLayerName, remoteDB)
            If Not GetLayerByTableName Is Nothing Then
                Exit Function
            End If
        Next i
    End Function

    Public Overloads Shared Function GetLayerByBrowseName(ByVal tablename As String, ByVal pMap As IMap) As ILayer
        If tablename = "" Then
            Return Nothing
        End If

        GetLayerByBrowseName = Nothing
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1
            GetLayerByBrowseName = GetLayerByBrowseName(tablename, pMap.Layer(i))
            If Not GetLayerByBrowseName Is Nothing Then
                Exit Function
            End If
        Next i
    End Function

    Public Overloads Shared Function GetLayerByBrowseName(ByVal tablename As String, ByVal pLayer As ILayer) As ILayer
        On Error GoTo ErrorHandler

        If pLayer Is Nothing Or Not pLayer.Valid Then
            Return Nothing
        End If

        If Not TypeOf pLayer Is IGroupLayer Then
            If TypeOf pLayer Is IFeatureLayer Then
                Dim pFeatureLayer As IFeatureLayer
                pFeatureLayer = pLayer
                Dim pDataset As IDataset
                pDataset = pFeatureLayer.FeatureClass

                If tablename.IndexOf(".") > 0 Then
                    If UCase(pDataset.BrowseName) = UCase(tablename) Then
                        Return pLayer
                        Exit Function
                    End If
                Else
                    If UCase(GetTableName(pDataset.BrowseName)) = UCase(tablename) Then
                        Return pLayer
                        Exit Function
                    End If
                End If
            End If
            Return Nothing
            Exit Function
        End If

        Dim pCompositeLayer As ICompositeLayer
        Dim i As Integer
        If TypeOf pLayer Is IGroupLayer Then
            pCompositeLayer = pLayer
            For i = 0 To pCompositeLayer.Count - 1
                Dim findLayer As ILayer = GetLayerByBrowseName(tablename, pCompositeLayer.Layer(i))
                If Not findLayer Is Nothing Then
                    Return findLayer
                End If
            Next i
        End If

        Return Nothing

        Exit Function
ErrorHandler:
        MsgBox("GetLayerByBrowseName " & Err.Description)
        Err.Clear()
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


    Public Overloads Shared Function GetDomainCode(ByRef pCombo As ComboBox) As String
        Dim domain As clsGisDomain = pCombo.SelectedItem
        If domain Is Nothing Then
            Return ""
        Else
            Return domain.Value
        End If
    End Function

    Public Shared Sub PopulateDomain(ByRef pCombo As ComboBox, ByVal pCodeDomain As ICodedValueDomain, ByVal allowNull As Boolean, Optional ByVal value As Object = VariantType.Null)

        Dim cindex As Long
        pCombo.Items.Clear()
        If allowNull Then
            Dim domain As clsGisDomain = New clsGisDomain
            domain.Name = ""
            domain.Value = ""

            pCombo.Items.Add(domain)
            If value = "" Then
                pCombo.SelectedIndex = 0
            End If
        End If

        For cindex = 0 To pCodeDomain.CodeCount - 1
            Dim domain As clsGisDomain = New clsGisDomain
            domain.Name = pCodeDomain.Name(cindex)
            domain.Value = pCodeDomain.Value(cindex)

            pCombo.Items.Add(domain)

            If value <> "" Then
                If value = pCodeDomain.Value(cindex) Then
                    pCombo.SelectedIndex = pCombo.Items.Count - 1
                End If
            End If
        Next

        pCombo.DisplayMember = "Name"
        pCombo.ValueMember = "Value"

    End Sub

    Public Shared Function GetValue(ByVal pFeature As IFeature, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Return GetRowValue(CType(pFeature, IRow), col, getDomainValue)
    End Function


    Public Shared Function GetRowValue(ByVal pRow As IRow, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Dim fIndex As Integer
        fIndex = pRow.Fields.FindField(col)
        If fIndex < 0 Then
            MsgBox("Can't find field [" & UCase(col) & "]")
            Return ""
        End If
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

        pCombo.SelectedIndex = pCombo.FindString(value)

        If pCombo.SelectedIndex >= 0 AndAlso value = "" AndAlso Trim(pCombo.SelectedText) <> "" Then
            pCombo.SelectedIndex = -1
            Exit Sub
        End If

        If pCombo.SelectedIndex = -1 Then
            pCombo.Text = value
        End If
    End Sub

    Public Shared Sub SetComboIndex2(ByRef pCombo As ComboBox, ByVal value As String)
        Dim i As Integer
        For i = 0 To pCombo.Items.Count - 1
            Dim domain As clsGisDomain = pCombo.Items.Item(i)
            If UCase(domain.Value) = UCase(value) Then
                pCombo.SelectedIndex = i
                Return
            End If

        Next

        pCombo.SelectedIndex = -1

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

    Public Overloads Shared Function GetDomainCode(ByVal pCodeDomain As ICodedValueDomain, ByVal value As String) As Object

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

	Public Shared Function GetPremiseWS(ByVal pMap As IMap) As IFeatureWorkspace
		Dim jtxExt As ESRI.ArcGIS.JTXExt.JTXExtension = Nothing
		Try
			jtxExt = DirectCast(IAISApplication.GetInstance.ArcMapApplicaiton.FindExtensionByName("Workflow Manager"), ESRI.ArcGIS.JTXExt.JTXExtension)
		Catch ex As Exception

		End Try

		If Not jtxExt Is Nothing Then
			If Not jtxExt.Job Is Nothing Then
				Try
					Return jtxExt.Database.DataWorkspace(jtxExt.Job.VersionName)
				Catch ex As Exception
					'No workspace information from job 
				End Try
			End If
		End If

		Dim pLayer As IFeatureLayer = GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), pMap, remoteDB:=True)
		If pLayer Is Nothing Then
			pLayer = GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap, remoteDB:=True)
		End If

		If pLayer Is Nothing Then
			Throw New Exception("Can't get the connection to database. Premise Point or OwnerPly layer is not loaded.")
		End If

		Dim pDataset As IDataset = pLayer
		Return DirectCast(pDataset.Workspace, IFeatureWorkspace)
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

	'Public Shared Function GetSelectedCol(ByVal pMap As IMap) As IFeature

	'    Dim pColLayer As IFeatureLayer
	'    Dim pColSel As IFeatureSelection
	'    Dim pColCursor As IFeatureCursor

	'    pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap)
	'    pColSel = pColLayer

	'    If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
	'        pColSel.SelectionSet.Search(Nothing, False, pColCursor)
	'        Return pColCursor.NextFeature
	'    End If

	'    pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), pMap)
	'    pColSel = pColLayer

	'    If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
	'        pColSel.SelectionSet.Search(Nothing, False, pColCursor)
	'        Return pColCursor.NextFeature
	'    End If

	'    pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), pMap)
	'    pColSel = pColLayer

	'    If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
	'        pColSel.SelectionSet.Search(Nothing, False, pColCursor)
	'        Return pColCursor.NextFeature
	'    End If

	'    pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), pMap)
	'    pColSel = pColLayer

	'    If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
	'        pColSel.SelectionSet.Search(Nothing, False, pColCursor)
	'        Return pColCursor.NextFeature
	'    End If

	'    Return Nothing


	'End Function

	Public Shared Function GetSelectedCol(ByVal pMap As IMap, Optional ByRef pLayerParaName As String = "") As IFeature

		Dim pColLayer As IFeatureLayer
		Dim pColSel As IFeatureSelection
		Dim pColCursor As IFeatureCursor

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
			pColSel.SelectionSet.Search(Nothing, False, pColCursor)
			pLayerParaName = "LAYER_OWNERPLY"
			Return pColCursor.NextFeature
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
			pColSel.SelectionSet.Search(Nothing, False, pColCursor)
			pLayerParaName = "LAYER_OWNERGAPPLY"
			Return pColCursor.NextFeature
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
			pColSel.SelectionSet.Search(Nothing, False, pColCursor)
			pLayerParaName = "LAYER_APPEALOWNERPLY"
			Return pColCursor.NextFeature
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing AndAlso pColLayer.Visible AndAlso pColSel.SelectionSet.Count = 1 Then
			pColSel.SelectionSet.Search(Nothing, False, pColCursor)
			pLayerParaName = "LAYER_REVOWNERPLY"
			Return pColCursor.NextFeature
		End If

		Return Nothing


	End Function


	Public Shared Function GetColLayers(ByVal pMap As IMap) As List(Of IFeatureLayer)
		Dim colLayerList As List(Of IFeatureLayer) = New List(Of IFeatureLayer)
		Dim i As Integer
		For i = 0 To pMap.LayerCount - 1
			GetColLayer(pMap.Layer(i), colLayerList)
		Next i

		Return colLayerList
	End Function

	Public Shared Sub GetColLayer(ByRef pLayer As ILayer, ByRef layerList As List(Of IFeatureLayer))
		Dim pDataset As IDataset
		If TypeOf pLayer Is IFeatureLayer Then
			pDataset = DirectCast(pLayer, IFeatureLayer).FeatureClass
			If UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY")) Or
				UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY")) Or
				UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY")) Or
				UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY")) Then

				layerList.Add(pLayer)

			End If

		ElseIf TypeOf pLayer Is IGroupLayer Then

			Dim i As Integer
			Dim pCompositeLayer As ICompositeLayer
			pCompositeLayer = pLayer
			For i = 0 To pCompositeLayer.Count - 1
				GetColLayer(pCompositeLayer.Layer(i), layerList)
			Next i

		End If
	End Sub

	Public Shared Function GetSelectedPremisePt(ByVal pMap As IMap) As IFeature
		Dim premiseLayerList As List(Of IFeatureLayer) = New List(Of IFeatureLayer)
		Dim i As Integer
		For i = 0 To pMap.LayerCount - 1
			GetPremiseLayer(pMap.Layer(i), premiseLayerList)
		Next i

		If premiseLayerList.Count = 0 Then
			Return Nothing
		End If


		Dim pSel As IFeatureSelection
		Dim pCursor As IFeatureCursor


		Dim lCount As Integer = 0
		For lCount = 0 To premiseLayerList.Count - 1
			Dim pLayer As IFeatureLayer = premiseLayerList.Item(lCount)
			If pLayer.Visible And pLayer.Selectable Then
				pSel = pLayer
				If pSel.SelectionSet.Count > 0 Then
					pSel.SelectionSet.Search(Nothing, False, pCursor)
					Return pCursor.NextFeature
				End If

			End If
		Next

		Return Nothing

	End Function

	Public Shared Function GetSelectedPremisePtCount(ByVal pMap As IMap) As Integer
		Dim premiseLayerList As List(Of IFeatureLayer) = New List(Of IFeatureLayer)
		Dim i As Integer
		For i = 0 To pMap.LayerCount - 1
			GetPremiseLayer(pMap.Layer(i), premiseLayerList)
		Next i

		If premiseLayerList.Count = 0 Then
			Return 0
		End If

		Dim iSelCount As Integer = 0

		Dim pSel As IFeatureSelection

		Dim lCount As Integer = 0
		For lCount = 0 To premiseLayerList.Count - 1
			Dim pLayer As IFeatureLayer = premiseLayerList.Item(lCount)
			If pLayer.Visible And pLayer.Selectable Then
				pSel = pLayer
				iSelCount = iSelCount + pSel.SelectionSet.Count
			End If
		Next

		Return iSelCount

	End Function

	Public Shared Sub GetPremiseLayer(ByRef pLayer As ILayer, ByRef layerList As List(Of IFeatureLayer))
		Dim pDataset As IDataset
		If TypeOf pLayer Is IFeatureLayer Then
			pDataset = DirectCast(pLayer, IFeatureLayer).FeatureClass
			If UCase(pDataset.BrowseName) = UCase(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT")) Then
				layerList.Add(pLayer)
			End If

		ElseIf TypeOf pLayer Is IGroupLayer Then

			Dim i As Integer
			Dim pCompositeLayer As ICompositeLayer
			pCompositeLayer = pLayer
			For i = 0 To pCompositeLayer.Count - 1
				GetPremiseLayer(pCompositeLayer.Layer(i), layerList)
			Next i

		End If
	End Sub


	Public Shared Function SelectedCol(ByVal pMap As IMap, ByVal pFilter As IQueryFilter, Optional ByRef layerName As String = Nothing) As IFeature

		Dim pColLayerList As List(Of IFeatureLayer) = GetColLayers(pMap)
		Dim pColLayer As IFeatureLayer
		Dim pColCursor As IFeatureCursor
		Dim pCol As IFeature

		Dim lCount As Integer = 0
		For lCount = 0 To pColLayerList.Count - 1
			pColLayer = pColLayerList.Item(lCount)
			If pColLayer.Visible And pColLayer.Selectable Then
				pColCursor = pColLayer.Search(pFilter, False)
				pCol = pColCursor.NextFeature
				If Not pCol Is Nothing Then
					If Not layerName Is Nothing Then
						layerName = DirectCast(pColLayer.FeatureClass, IDataset).BrowseName
					End If
					Return pCol
				End If
			End If
		Next

		Return Nothing

	End Function


	Public Shared Function GetSelectedColCount(ByVal pMap As IMap) As Integer

		Dim pColLayer As IFeatureLayer
		Dim pColSel As IFeatureSelection

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap)
		pColSel = pColLayer

		Dim fCount As Integer = 0

		If Not pColLayer Is Nothing Then
			If pColLayer.Visible And pColLayer.Selectable Then
				fCount = fCount + pColSel.SelectionSet.Count
			End If
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing Then
			If pColLayer.Visible And pColLayer.Selectable Then
				fCount = fCount + pColSel.SelectionSet.Count
			End If

		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing Then
			If pColLayer.Visible And pColLayer.Selectable Then
				fCount = fCount + pColSel.SelectionSet.Count
			End If
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), pMap)
		pColSel = pColLayer

		If Not pColLayer Is Nothing Then
			If pColLayer.Visible And pColLayer.Selectable Then
				fCount = fCount + pColSel.SelectionSet.Count
			End If
		End If

		Return fCount

	End Function

	Public Shared Function IsOneColSelected(ByVal pMap As IMap) As Boolean
		Dim selCount As Integer = 0

		Dim pColLayer As IFeatureLayer
		Dim pColSel As IFeatureSelection

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), pMap)

		If Not pColLayer Is Nothing AndAlso pColLayer.Visible Then
			pColSel = pColLayer
			selCount = selCount + pColSel.SelectionSet.Count
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), pMap)
		If Not pColLayer Is Nothing AndAlso pColLayer.Visible Then
			pColSel = pColLayer
			selCount = selCount + pColSel.SelectionSet.Count
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_APPEALOWNERPLY"), pMap)
		If Not pColLayer Is Nothing AndAlso pColLayer.Visible Then
			pColSel = pColLayer
			selCount = selCount + pColSel.SelectionSet.Count
		End If

		pColLayer = MapUtil.GetLayerByTableName(IAISToolSetting.GetParameterValue("LAYER_REVOWNERPLY"), pMap)
		If Not pColLayer Is Nothing AndAlso pColLayer.Visible Then
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
		If minScale = 0 Then
			Try
				minScale = CDbl(IAISToolSetting.GetParameterValue("MAP_REFERENCESCALE"))
			Catch ex As Exception
				minScale = 1000
			End Try
		End If

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

			System.Windows.Forms.Application.DoEvents()

		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try

	End Function

	Public Shared Function ZoomToGeometry(ByRef pGeom As IGeometry, ByRef pMap As IMap) As Boolean
		Dim minScale As Double = pMap.ReferenceScale
		Dim pNewMapExtent As IEnvelope
		Dim pActiveView As IActiveView
		Dim pDisplayTransform As IDisplayTransformation

		Try

			pActiveView = pMap
			pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation

			If pGeom.Envelope.Width = 0 Then
				'This is a point
				pNewMapExtent = pDisplayTransform.VisibleBounds

				Dim pCenterPoint As IPoint
				pCenterPoint = New Point

				pCenterPoint.X = pGeom.Envelope.XMin
				pCenterPoint.Y = pGeom.Envelope.YMin

				pNewMapExtent.CenterAt(pCenterPoint)
				pDisplayTransform.VisibleBounds = pNewMapExtent
				pMap.MapScale = minScale
				pActiveView.Refresh()
				Return True
			Else
				pNewMapExtent = pGeom.Envelope
				pNewMapExtent.Expand(1.05, 1.05, True)
				pDisplayTransform.VisibleBounds = pNewMapExtent

				'If pMap.MapScale < minScale Then
				'    pMap.MapScale = minScale
				'End If

				pActiveView.Refresh()

				Return True
			End If

			System.Windows.Forms.Application.DoEvents()

		Catch ex As Exception
			MsgBox(ex.Message)
			Return False
		End Try

	End Function

	Public Shared Function PanToGeometry(ByRef pGeom As IGeometry, ByRef pMap As IMap) As Boolean
		Dim pNewMapExtent As IEnvelope
		Dim pActiveView As IActiveView
		Dim pDisplayTransform As IDisplayTransformation
		Try
			pActiveView = pMap
			pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation

			pNewMapExtent = pDisplayTransform.VisibleBounds

			Dim pCenterPoint As IPoint
			pCenterPoint = New Point

			pCenterPoint.X = (pGeom.Envelope.XMin + pGeom.Envelope.XMax) / 2
			pCenterPoint.Y = (pGeom.Envelope.YMin + pGeom.Envelope.YMax) / 2

			pNewMapExtent.CenterAt(pCenterPoint)
			pDisplayTransform.VisibleBounds = pNewMapExtent

			pActiveView.Refresh()
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

		'System.Windows.Forms.Application.DoEvents()
	End Sub

	Public Shared Sub FlashGeometry(ByVal pGeo As IGeometry, ByVal pMap As IMap, Optional ByVal iInterval As Integer = 300)

		Dim pSimpleLineSymbol As ILineSymbol
		Dim pSimpleFillSymbol As ISimpleFillSymbol
		Dim pSimpleMarkersymbol As ISimpleMarkerSymbol
		Dim pActive As IActiveView
		Dim pSymbol As ISymbol
		Dim pDisplay As IScreenDisplay
		Dim pColor As IRgbColor

		pColor = New RgbColor
		pColor.Red = 50
		pColor.Blue = 175
		pColor.Green = 50
		pActive = pMap
		pDisplay = pActive.ScreenDisplay

		pDisplay.StartDrawing(0, esriScreenCache.esriNoScreenCache)

		Select Case pGeo.GeometryType
			Case esriGeometryType.esriGeometryLine
				pSimpleLineSymbol = New SimpleLineSymbol
				pSymbol = pSimpleLineSymbol  'QI
				pSymbol.ROP2 = esriRasterOpCode.esriROPNotXOrPen  'erase itself when drawn twice
				pSimpleLineSymbol.Width = 4
				pSimpleLineSymbol.Color = pColor
				pDisplay.SetSymbol(pSimpleLineSymbol)
				pDisplay.DrawPolyline(pGeo)
				System.Threading.Thread.Sleep(iInterval)
				pDisplay.DrawPolyline(pGeo)
			Case esriGeometryType.esriGeometryPolygon
				pSimpleFillSymbol = New SimpleFillSymbol
				pSymbol = pSimpleFillSymbol
				pSymbol.ROP2 = esriRasterOpCode.esriROPNotXOrPen
				pSimpleFillSymbol.Color = pColor
				pDisplay.SetSymbol(pSimpleFillSymbol)
				pDisplay.DrawPolygon(pGeo)
				System.Threading.Thread.Sleep(iInterval)
				pDisplay.DrawPolygon(pGeo)
			Case esriGeometryType.esriGeometryPoint
				pSimpleMarkersymbol = New SimpleMarkerSymbol
				pSymbol = pSimpleMarkersymbol
				pSymbol.ROP2 = esriRasterOpCode.esriROPNotXOrPen
				pSimpleMarkersymbol.Color = pColor
				pSimpleMarkersymbol.Size = 12
				pDisplay.SetSymbol(pSimpleMarkersymbol)
				pDisplay.DrawPoint(pGeo)
				System.Threading.Thread.Sleep(iInterval)
				pDisplay.DrawPoint(pGeo)
			Case esriGeometryType.esriGeometryMultipoint
				pSimpleMarkersymbol = New SimpleMarkerSymbol
				pSymbol = pSimpleMarkersymbol
				pSymbol.ROP2 = esriRasterOpCode.esriROPNotXOrPen
				pSimpleMarkersymbol.Color = pColor
				pSimpleMarkersymbol.Size = 12
				pDisplay.SetSymbol(pSimpleMarkersymbol)
				pDisplay.DrawMultipoint(pGeo)
				System.Threading.Thread.Sleep(iInterval)
				pDisplay.DrawMultipoint(pGeo)
		End Select

		pDisplay.FinishDrawing()
	End Sub


	Public Shared Function GetTableFromWS(ByRef pFWS As IFeatureWorkspace, ByVal tableName As String) As ITable
		Return pFWS.OpenTable(tableName)
	End Function

	Public Shared Function GetLayerFromWS(ByRef pFWS As IFeatureWorkspace, ByVal tableName As String) As IFeatureLayer
		Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
		Try
			pFeatureLayer.FeatureClass = pFWS.OpenFeatureClass(tableName)
		Catch ex As Exception
			MsgBox(ex.Message)
			pFeatureLayer = Nothing
		End Try

		Return pFeatureLayer
	End Function

    Public Shared Function GetSequenceID(ByRef pFWS As IFeatureWorkspace, ByVal sequenceName As String) As Integer

        Dim pCursor As ICursor = DirectCast(pFWS, ISqlWorkspace).OpenQueryCursor("Select NEXT VALUE FOR " & sequenceName)
        Dim pRow As IRow = pCursor.NextRow
        Return pRow.Value(0)
    End Function

    Public Shared Function AddLayerFromWS(ByRef pFWS As IFeatureWorkspace, ByVal tableName As String, ByRef pMap As IMap) As IFeatureLayer
		Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
		Dim formMsg As FormMessage = New FormMessage()
		formMsg.TextBoxMessage.Text = "Loading layer " & tableName & " ... "
		formMsg.Show()
		Try
			pFeatureLayer.FeatureClass = pFWS.OpenFeatureClass(tableName)
			pFeatureLayer.Name = tableName
			pMap.AddLayer(pFeatureLayer)
		Catch ex As Exception
			MsgBox(ex.Message)
			pFeatureLayer = Nothing
		End Try

		formMsg.Hide()
		formMsg = Nothing
		Return pFeatureLayer

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

	Public Shared Function SetCurrentTool(ByVal strToolName As String, ByRef app As IApplication, Optional ByVal subtype As Integer = 0) As Boolean

        Dim pUid As New UID
        pUid.Value = strToolName
        pUid.SubType = subtype
        Dim pArcMapCmd As ICommandItem = app.Document.CommandBars.Find(pUid)
        If Not pArcMapCmd Is Nothing Then
            pArcMapCmd.Execute()
            Return True
        Else
            Return False
        End If
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


    Public Shared Function IsExtentionEnabled(ByVal extName As String, ByRef app As IApplication) As Boolean
        Dim uid As New UID
        uid.Value = extName
        Dim extension As IExtensionConfig
        extension = app.FindExtensionByCLSID(uid)
        If extension Is Nothing Then
            Return False
        End If

        Return (extension.State = esriExtensionState.esriESEnabled)
    End Function

    Public Shared Function GetExtention(ByVal extName As String, ByRef app As IApplication) As IExtension
        Dim uid As New UID
        uid.Value = extName
        Return app.FindExtensionByCLSID(uid)
    End Function

    Public Shared Function ToggleToolBar(ByVal strToolBarName As String, ByVal visibleFlag As Boolean, ByRef app As IApplication) As Boolean
        Dim pCmdBars As ICommandBars
        pCmdBars = app.Document.CommandBars
        Dim pCmdBar As ICommandBar


        pCmdBar = pCmdBars.Find(strToolBarName)
        If Not pCmdBar Is Nothing Then
            If visibleFlag Then
                If Not pCmdBar.IsVisible() Then
                    pCmdBar.Dock(esriDockFlags.esriDockShow)
                End If
            ElseIf pCmdBar.IsVisible() Then
                pCmdBar.Dock(esriDockFlags.esriDockHide)
            End If
        End If
    End Function



    Public Shared Function IsNumericKey(ByVal KCode As String) As Boolean
        If (KCode >= 48 And KCode <= 57) Or KCode = 8 Or KCode = 3 Or KCode = 22 Or KCode = 37 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Shared Function IsNumericValue(ByVal value As String) As Boolean
        If value = "" Then
            Return True
        End If

        If IsNumeric(value) Then
            Return True
        Else
            MsgBox("Only numeric value is allowed for this field.")
            Return False
        End If

    End Function


    Public Overloads Shared Function GetOverlapShape(ByRef pGeom As IGeometry, ByRef pFeatureLayer As IFeatureLayer) As IGeometry
        Dim pFilter As ISpatialFilter = New SpatialFilter
        pFilter.Geometry = pGeom
        pFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

        Dim pRelOp As IRelationalOperator
        pRelOp = pGeom

        Dim pFeatCursor As IFeatureCursor
        pFeatCursor = pFeatureLayer.Search(pFilter, False)

        Dim pR As IFeature = pFeatCursor.NextFeature
        Do While Not pR Is Nothing

            If pRelOp.Overlaps(pR.Shape) Or _
                pRelOp.Equals(pR.Shape) Or _
                pRelOp.Contains(pR.Shape) Or _
                pRelOp.Within(pR.Shape) Then
                Return pR.Shape
            End If

            pR = pFeatCursor.NextFeature
        Loop

        Return Nothing

    End Function


    Public Shared Function TrimAll(ByVal TextIn As String, Optional ByVal TrimChar As String = " ", Optional ByVal CtrlChar As String = Chr(0)) As String

        Try ' Replace ALL Duplicate Characters in String with a Single Instance
            TrimAll = Replace(TextIn, TrimChar, CtrlChar)   ' In case of CrLf etc.

            While InStr(TrimAll, CtrlChar + CtrlChar) > 0
                TrimAll = TrimAll.Replace(CtrlChar + CtrlChar, CtrlChar)
            End While

            TrimAll = TrimAll.Trim(CtrlChar)                ' Trim Begining and End
            TrimAll = TrimAll.Replace(CtrlChar, TrimChar)   ' Replace with Original Trim Character(s)
        Catch Exp As Exception
            TrimAll = TextIn    ' Justin Case
        End Try

        Return TrimAll

    End Function

    Public Shared Function GetPremiseAddress(ByVal pPremiseRow As IRow)
        Dim addressStr As String = ""
        If GetRowValue(pPremiseRow, "USE_PARSED_ADDRESS") = "Y" Then
            addressStr = TrimAll(GetRowValue(pPremiseRow, "ADDRNUM") & " " & _
                                 GetRowValue(pPremiseRow, "ADDRNUMSUF") & " " & _
                                 GetRowValue(pPremiseRow, "STNAME") & " " & _
                                 GetRowValue(pPremiseRow, "STREET_TYPE") & " " & _
                                 GetRowValue(pPremiseRow, "QUADRANT"))
            Dim unit As String = GetRowValue(pPremiseRow, "UNIT")
            If unit <> "" Then
                addressStr = addressStr + " #" + unit
            End If
        Else
            addressStr = GetRowValue(pPremiseRow, "PEXSAD")
        End If

        Return addressStr

    End Function


    Public Shared Function GetRoundERU(ByVal iasqft As Double, ByRef app As IApplication) As Double
        Dim factor As Integer = CInt(IAISToolSetting.GetParameterValue("ERU_ROUND_PRECISION"))
        Dim eru_sqft As Double = CInt(IAISToolSetting.GetParameterValue("ERU_SQFT"))
        Dim sqft_round_precision As Double = CInt(IAISToolSetting.GetParameterValue("SQFT_ROUND_PRECISION"))
        If iasqft < sqft_round_precision Then
            Return 0
        Else
            Return FormatNumber(iasqft / eru_sqft, factor)
        End If
    End Function

    Public Shared Function GetBillableERU(ByVal premiseType As String, ByVal iasqft As Double, _
                                          ByRef iaTierLevel As ITable) As Double
        Dim billableERU As Double
        Dim sqft_round_precision As Double = CInt(IAISToolSetting.GetParameterValue("SQFT_ROUND_PRECISION"))
        If iasqft < sqft_round_precision Then
            Return 0
        End If

        Dim iaTierLevelRow As IRow
        Dim pQFilter As IQueryFilter = New QueryFilter
        pQFilter.WhereClause = "PEXPTYP='" & premiseType & "' AND LOWERSQFT <= " &
                                iasqft & " AND UPPERSQFT >= " & iasqft &
                                " AND EFFSTARTDT <= GETDATE() AND EFFENDDT IS NULL"
        Dim pCursor As ICursor = iaTierLevel.Search(pQFilter, False)
        iaTierLevelRow = pCursor.NextRow
        If Not iaTierLevelRow Is Nothing Then
            If IsDBNull(iaTierLevelRow.Value(iaTierLevelRow.Fields.FindField("MEDIANSQFT"))) Then
                billableERU = iasqft
            Else
                billableERU = iaTierLevelRow.Value(iaTierLevelRow.Fields.FindField("MEDIANSQFT"))
            End If
        Else
            Throw New Exception("No rate is found for premise type " & premiseType & " ( " & iasqft & ") in Tier Level table.")
        End If
        Return billableERU

    End Function

    Public Shared Function IsEnrolledInRSRCredit(ByVal pexuid As Long, ByVal pMap As IMap) As Boolean
        Dim flag As Boolean = False
        Dim filter As IQueryFilter = New QueryFilterClass
        filter.set_WhereClause(("EFFENDDT IS NULL AND PEXUID=" & Conversions.ToString(pexuid)))
        Dim row As IRow = DirectCast(DirectCast(MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT", DirectCast(Nothing, IApplication)), pMap), IDataset).get_Workspace, IFeatureWorkspace).OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF", DirectCast(Nothing, IApplication))).Search(filter, False).NextRow
        If (Not row Is Nothing) Then
            Dim num1 As Long = row.get_Fields.FindField("RSRCREDERU")
            If Not Information.IsDBNull(row.get_Value(row.get_Fields.FindField("IACREDERU"))) Then
                flag = True
            End If
        End If
        Return flag
    End Function

    Public Shared Function GetVersionName(ByVal fullname As String) As String
        Dim iLoc As Integer
        iLoc = InStrRev(fullname, ".")
        If iLoc <= 0 Then
            GetVersionName = UCase(fullname)
        Else
            GetVersionName = UCase(Mid(fullname, iLoc + 1))
        End If

        Exit Function
    End Function

End Class



