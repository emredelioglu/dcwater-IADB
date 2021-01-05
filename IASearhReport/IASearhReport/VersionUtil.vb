Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

Public Class VersionUtil
    Public Sub ChangeVersionByName(ByVal pMxDoc As IMxDocument, _
                                   ByVal pOldWorkspace As IFeatureWorkspace, _
                                   ByVal sVersionName As String)

        On Error GoTo ErrorHandler

        Dim pVersionedWorkspace As IVersionedWorkspace
        Dim pVersion As IVersion

        If (pMxDoc Is Nothing) Or (pOldWorkspace Is Nothing) Or (Len(sVersionName) = 0) Then
            Err.Raise(vbObjectError + 1, "ChangeVersionByName", "Invalid parameters")
        End If

        If Not TypeOf pOldWorkspace Is IVersionedWorkspace Then
            Err.Raise(vbObjectError + 1, "ChangeVersionByName", "Old Workspace does not support IVersionedWorkspace")
        End If

        pVersionedWorkspace = pOldWorkspace
        pVersion = pVersionedWorkspace.FindVersion(sVersionName)

            If pVersion Is Nothing Then
                Err.Raise(vbObjectError + 1, "ChangeVersionByName", "Could not open version [" & sVersionName & "]")
            End If

            ChangeVersion(pMxDoc, pOldWorkspace, pVersion)

        Exit Sub
ErrorHandler:
        '**if/else structure added by amanda
        If Err.Number = -2147215997 Then
            MsgBox("MXD references historical version.  Please manually set your view to a JTX version.")
        Else
            MsgBox("ChangeVersionByName: " & Err.Description)
        End If

    End Sub

    Public Sub ChangeVersion(ByVal pMxDoc As IMxDocument, _
                             ByVal pOldWorkspace As IFeatureWorkspace, _
                             ByVal pNewWorkspace As IFeatureWorkspace)

        On Error GoTo ErrorHandler

        Dim pVersion As IVersion
        Dim pMaps As IMaps
        Dim iCurrMap As Long
        Dim pMap As IMap
        Dim pMapAdmin2 As IMapAdmin2
        Dim pCollectionTableVersionChanges As ICollectionTableVersionChanges
        Dim pActiveView As IActiveView

        If (pMxDoc Is Nothing) Or (pOldWorkspace Is Nothing) Or (pNewWorkspace Is Nothing) Then
            Err.Raise(vbObjectError + 1, "ChangeVersion", "Invalid parameters")
        End If

        If IsBeingEdited(pOldWorkspace) Then
            Err.Raise(vbObjectError + 1, "ChangeVersion", "Must not be editing to switch versions")
        End If

        'Get a collection to stuff all the FeatureClasses/Tables that changed
        pCollectionTableVersionChanges = New EnumTableVersionChanges

        'Make sure we have the absolute latest state of this version
        'pVersion = pNewWorkspace
        'pVersion.RefreshVersion()

        pMaps = pMxDoc.Maps
        For iCurrMap = 0 To pMaps.Count - 1
            pMap = pMaps.Item(iCurrMap)
            pMap.ClearSelection()

            'Notify clients listening that the version has changed
            pMapAdmin2 = pMap
            pMapAdmin2.FireChangeVersion(pOldWorkspace, pNewWorkspace)

            'Change version of FeatureLayers
            ChangeFeatureLayers(pMap, pOldWorkspace, pNewWorkspace, pCollectionTableVersionChanges)

            'Change Version of StandaloneTables
            ChangeTables(pMap, pOldWorkspace, pNewWorkspace, pCollectionTableVersionChanges)

            'Notify clients that FeatureClasses and Tables changed
            FireChangeDataset(pMap, pCollectionTableVersionChanges)

            pCollectionTableVersionChanges.RemoveAll()

            'Refresh Map
            pActiveView = pMap
            pActiveView.Refresh()

        Next iCurrMap

        'Refresh TOC
        pMxDoc.UpdateContents()

        Exit Sub
ErrorHandler:
        MsgBox("ChangeVersion: " & Err.Description)
    End Sub


    Private Function IsBeingEdited(ByVal pFWorkspace As IFeatureWorkspace) As Boolean

        Dim pWorkspaceEdit As IWorkspaceEdit

        If TypeOf pFWorkspace Is IWorkspaceEdit Then
            pWorkspaceEdit = pFWorkspace
            IsBeingEdited = pWorkspaceEdit.IsBeingEdited
        Else
            IsBeingEdited = False
        End If
    End Function

    Private Sub ChangeFeatureLayers(ByVal pMap As IMap, _
                                    ByVal pOldWorkspace As IFeatureWorkspace, _
                                    ByVal pNewWorkspace As IFeatureWorkspace, _
                                    ByVal pCollectionTableVersionChanges As ICollectionTableVersionChanges)

        On Error GoTo ErrorHandler

        Dim pMapAdmin2 As IMapAdmin2
        Dim pEnumLayer As IEnumLayer
        Dim pLayer As ILayer
        Dim pDisplayTable As IDisplayTable

        Dim pFeatureLayer As IFeatureLayer
        Dim pDataset As IDataset
        Dim pWorkspace As IWorkspace
        Dim sTableName As String
        Dim pNewFeatureClass As IFeatureClass

        pMapAdmin2 = pMap
        pEnumLayer = pMap.Layers(Nothing, True)
        If pEnumLayer Is Nothing Then Exit Sub
        pLayer = pEnumLayer.Next
        While Not pLayer Is Nothing
            'Update the internal tables if joined
            If TypeOf pLayer Is IDisplayTable Then
                pDisplayTable = pLayer
                ChangeRelQueryTable(pDisplayTable.DisplayTable, _
                                    pOldWorkspace, _
                                    pNewWorkspace, _
                                    pCollectionTableVersionChanges)
            End If

            If TypeOf pLayer Is IFeatureLayer Then
                pFeatureLayer = pLayer
                pDataset = pFeatureLayer.FeatureClass
                If Not pDataset Is Nothing Then
                    ChangeRelQueryTable(pDataset, _
                                        pOldWorkspace, _
                                        pNewWorkspace, _
                                        pCollectionTableVersionChanges)

                    pWorkspace = pDataset.Workspace
                    If pWorkspace Is pOldWorkspace Then
                        pNewFeatureClass = OpenDataset(pNewWorkspace, pDataset)
                        If Not pNewFeatureClass Is Nothing Then
                            pFeatureLayer.FeatureClass = pNewFeatureClass
                            pCollectionTableVersionChanges.Add(pDataset, pNewFeatureClass)
                        End If
                    End If
                End If
            End If
            pLayer = pEnumLayer.Next
        End While
        Exit Sub
ErrorHandler:
        MsgBox("ChangeFeatureLayers: " & Err.Description)
    End Sub

    Private Sub ChangeTables(ByVal pTableCollection As IStandaloneTableCollection, _
                             ByVal pOldWorkspace As IFeatureWorkspace, _
                             ByVal pNewWorkspace As IFeatureWorkspace, _
                             ByVal pCollectionTableVersionChanges As ICollectionTableVersionChanges)

        On Error GoTo ErrorHandler

        Dim iCurrTable As Long
        Dim colTables As Collection
        Dim pDataset As IDataset
        Dim pWorkspace As IWorkspace
        Dim pDisplayTable As IDisplayTable
        Dim pActiveView As IActiveView
        Dim pTable As IStandaloneTable
        Dim pNewTable As ITable

        For iCurrTable = 0 To pTableCollection.StandaloneTableCount - 1
            pTable = pTableCollection.StandaloneTable(iCurrTable)
            If TypeOf pTable Is IDisplayTable Then
                pDisplayTable = pTable
                ChangeRelQueryTable(pDisplayTable.DisplayTable, _
                                    pOldWorkspace, _
                                    pNewWorkspace, _
                                    pCollectionTableVersionChanges)
            End If

            pDataset = pTable.Table
            If Not pDataset Is Nothing Then
                ChangeRelQueryTable(pDisplayTable.DisplayTable, _
                                    pOldWorkspace, _
                                    pNewWorkspace, _
                                    pCollectionTableVersionChanges)

                pWorkspace = pDataset.Workspace
                If pWorkspace Is pOldWorkspace Then
                    pNewTable = OpenDataset(pNewWorkspace, pDataset)
                    If Not pNewTable Is Nothing Then
                        pTable.Table = pNewTable
                        pCollectionTableVersionChanges.Add(pDataset, pNewTable)
                    End If
                End If
            End If
        Next iCurrTable
        Exit Sub
ErrorHandler:
        MsgBox("ChangeTables: " & Err.Description)
    End Sub

    Private Sub ChangeRelQueryTable(ByVal pDataset As IDataset, _
                                   ByVal pOldWorkspace As IFeatureWorkspace, _
                                   ByVal pNewWorkspace As IFeatureWorkspace, _
                                   ByVal pCollectionTableVersionChanges As ICollectionTableVersionChanges)

        On Error GoTo ErrorHandler

        Dim pRelQueryTableManage As IRelQueryTableManage

        If TypeOf pDataset Is IRelQueryTableManage Then
            pRelQueryTableManage = pDataset
            pRelQueryTableManage.VersionChanged(pOldWorkspace, pNewWorkspace, pCollectionTableVersionChanges)
            'Next line may not be needed.  Including it just in case
            pCollectionTableVersionChanges.Add(pDataset, pDataset)
        End If
        Exit Sub
ErrorHandler:
        'Just eat errors
    End Sub

    'Little Helper function to polymorphically open the current Dataset
    ' (Table, Featureclasses, etc)
    Private Function OpenDataset(ByVal pFeatureWorkspace As IFeatureWorkspace, _
                                 ByVal pDataset As IDataset) As ITable
        On Error Resume Next
        OpenDataset = pFeatureWorkspace.OpenTable(pDataset.Name)
    End Function

    Private Sub FireChangeDataset(ByVal pMapAdmin2 As IMapAdmin2, _
                                  ByVal pEnumTableVersionChanges As IEnumTableVersionChanges)
        On Error GoTo ErrorHandler

        Dim pOldTable As ITable
        Dim pNewTable As ITable

        'Loop through the collection and FireChange for each FeatureClass/Table
        pEnumTableVersionChanges.Reset()
        pEnumTableVersionChanges.Next(pOldTable, pNewTable)
        Do Until (pOldTable Is Nothing) Or (pNewTable Is Nothing)
            If TypeOf pOldTable Is IFeatureClass Then
                pMapAdmin2.FireChangeFeatureClass(pOldTable, pNewTable)
            Else
                pMapAdmin2.FireChangeTable(pOldTable, pNewTable)
            End If

            pEnumTableVersionChanges.Next(pOldTable, pNewTable)
        Loop
        Exit Sub
ErrorHandler:
        'Just eat errors
    End Sub

End Class
