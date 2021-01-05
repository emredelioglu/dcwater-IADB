Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geometry
Imports System.Data.OleDb
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices


Module Module1

    Private log As System.IO.StreamWriter
    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()

    Private m_jtxJob As IJTXJob = Nothing
    Private m_exitCode As Integer

    <STAThread()> _
    Sub Main()

        m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
        New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})


        'Create or open log file
        If Not File.Exists(Environment.CurrentDirectory & "\ImportOwnerPly.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\ImportOwnerPly.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\ImportOwnerPly.log")
        End If

        Try
            Dim sParams() As String = Environment.GetCommandLineArgs
            If sParams.Length > 1 Then

                log.WriteLine("Parameter routineName:" & sParams(1))

                If sParams(1) = "CreateTempGeoDB" Then
                    CreateTempGeoDB()
                ElseIf sParams(1) = "CheckLocationException" Then
                    CheckLocationException()
                ElseIf sParams(1) = "CheckOtrLocationException" Then
                    CheckOtrLocationException()
                ElseIf sParams(1) = "CheckNoIntersection" Then
                    CheckNoIntersection()
                ElseIf sParams(1) = "CheckOwnerPolyAttribute" Then
                    If sParams.Length > 2 Then
                        CheckOwnerPolyAttribute(CInt(sParams(2)), CInt(sParams(3)))
                    Else
                        CheckOwnerPolyAttribute()
                    End If
                ElseIf sParams(1) = "CheckOtrNoIntersection" Then
                    CheckOtrNoIntersection()
                ElseIf sParams(1) = "CheckDuplicate" Then
                    CheckOwnerPlyDuplicate()
                ElseIf sParams(1) = "CreateOwnerPlyJTXJob" Then
                    CreateOwnerPlyJTXJob()
                ElseIf sParams(1) = "ProcessOwnerplyAdd" Then
                    ProcessOwnerplyAdd()
                ElseIf sParams(1) = "ProcessOwnerplyDelete" Then
                    ProcessOwnerplyDelete()
                ElseIf sParams(1) = "ProcessOwnerplyUpdate" Then
                    ProcessOwnerplyUpdate()
                ElseIf sParams(1) = "ProcessUpdatepropertylineage" Then
                    ProcessUpdatepropertylineage()
                ElseIf sParams(1) = "ProcessExceptions" Then
                    ProcessExceptions()
                Else
                    Console.WriteLine("Usage for UpdateOwnerPly: UpdateOwnerPly CreateTempGeoDB|CheckLocationException|CheckNoIntersection|CheckOwnerPolyAttribute|CheckOtrNoIntersection|CheckDuplicate|CreateOwnerPlyJTXJob|ProcessOwnerplyAdd|ProcessOwnerplyDelete|ProcessOwnerplyUpdate|ProcessUpdatepropertylineage|ProcessExceptions")
                    Throw New Exception("Wrong arguments " & sParams(1))
                End If
            Else
                Console.WriteLine("Usage for UpdateOwnerPly: UpdateOwnerPly CreateTempGeoDB|CheckLocationException|CheckNoIntersection|CheckOwnerPolyAttribute|CheckOtrNoIntersection|CheckDuplicate|CreateOwnerPlyJTXJob|ProcessOwnerplyAdd|ProcessOwnerplyDelete|ProcessOwnerplyUpdate|ProcessUpdatepropertylineage|ProcessExceptions")
                Throw New Exception("Wrong number of arguments")
            End If

            m_exitCode = 0
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())
            Console.WriteLine("Error:" & ex.Message)
            m_exitCode = 1
        End Try

        log.Close()
        m_AOLicenseInitializer.ShutdownApplication()
        Environment.Exit(m_exitCode)
    End Sub

#Region "Create Temp DB"
    Sub CreateTempGeoDB()
        Try
            If File.Exists(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.mdb") Then
                log.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": Database already exists")
                Return
            End If

            Console.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": CreateTempGeoDB ")

            Dim tempDbName As String = "OwnerPlyExceptionTemp.mdb"
            File.Copy(Application.StartupPath & "\" & tempDbName, Environment.CurrentDirectory & "\" & tempDbName)
            log.WriteLine(Now & ": Temporary database created.")

        Catch ex As Exception
            Console.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": CreateTempGeoDB :" & ex.Message)
            Throw ex
        End Try

    End Sub

#End Region

#Region "CheckOwnerPlyDuplicate"

    Private Sub CheckOwnerPlyDuplicate()
        log.WriteLine(Now & ": Check duplicate geometry started.")


        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)
        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")

        Dim tempUpdatepropertylineage As ITable = pTempWorkspace.OpenTable("updatepropertylineage")

        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")


        Dim tableLineage As ITable = pTempWorkspace.OpenTable("PropertyConfilictLineage")

        Dim ownerPlyOId As Integer

        Dim ownerPly As IFeature
        Dim OwnerPlyCursor As IFeatureCursor

        Dim pLineageRow As IRow
        Dim pLineageCursor As ICursor


        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim counter As Integer = 0
        Console.WriteLine(Now & " : starting process")

        Try

            Dim pFilter As IQueryFilter = New QueryFilter
            Dim pSFilter As ISpatialFilter = New SpatialFilter

            'If there are duplicate geometry in OTROwnerPly then the frequency of 
            'IaisOwnerPly will be >= 2
            Dim pQueryDef As IQueryDef
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "IAISOtrPly, IAISOtrPly_Frequency"
            pQueryDef.SubFields = "IAISOtrPly.FID_OtrOwnerPly"
            pQueryDef.WhereClause = "IAISOtrPly.FID_IaisOwnerPly = IAISOtrPly_Frequency.FID_IaisOwnerPly AND IAISOtrPly_Frequency.FREQUENCY > 1 AND IAISOtrPly.FID_OtrOwnerPly >= 0"

            Dim pFreqCursor As ICursor = pQueryDef.Evaluate
            Dim pFreqRow As IRow = pFreqCursor.NextRow

            'Add checked id into this set so we don't check twice for the same feature
            Dim idSet As IFIDSet = New FIDSet

            Dim idxOidField As Integer = pFreqCursor.FindField("IAISOtrPly.FID_OtrOwnerPly")
            Dim bIdExists As Boolean
            Do While Not pFreqRow Is Nothing

                ownerPlyOId = pFreqRow.Value(idxOidField)
                idSet.Find(ownerPlyOId, bIdExists)

                If Not bIdExists Then
                    'OTR Ownerply
                    pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                    OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                    ownerPly = OwnerPlyCursor.NextFeature
                    Try
                        Dim ssl As String = ownerPly.Value(ownerPly.Fields.FindField("SSL"))

                        pSFilter.Geometry = ownerPly.Shape
                        pSFilter.WhereClause = otrOwnerPlyFC.OIDFieldName & " <> " & ownerPlyOId
                        pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                        Dim pOverlapPoly As IPolygon = ownerPly.Shape
                        Dim pRelOp As IRelationalOperator = pOverlapPoly
                        Dim pTopo As ITopologicalOperator

                        Marshal.ReleaseComObject(OwnerPlyCursor)

                        'Console.WriteLine("ownerPly OID:" & ownerPly.OID)

                        OwnerPlyCursor = otrOwnerPlyFC.Search(pSFilter, False)
                        Dim pR As IFeature = OwnerPlyCursor.NextFeature
                        Do While Not pR Is Nothing

                            'If pRelOp.Equals(pR.Shape) Or _
                            '    pRelOp.Contains(pR.Shape) Or _
                            '    pRelOp.Within(pR.Shape) Then
                            'Console.WriteLine(pR.OID)

                            'Console.WriteLine(pR.OID)
                            'End If
                            If pRelOp.Equals(pR.Shape) Or _
                                pRelOp.Contains(pR.Shape) Or _
                                pRelOp.Within(pR.Shape) Then

                                log.WriteLine(Now & ": duplicate geometry found. " & ownerPly.OID)

                                'Check to see if this record is in lineage table 
                                pFilter.WhereClause = "(INPUTFEATURESSL='" & ssl & "' AND CONFLICTFEATURESSL='" & pR.Value(pR.Fields.FindField("SSL")) & "' ) OR " & _
                                                        "(CONFLICTFEATURESSL='" & ssl & "' AND INPUTFEATURESSL='" & pR.Value(pR.Fields.FindField("SSL")) & "' )"

                                pLineageCursor = tableLineage.Search(pFilter, False)
                                pLineageRow = pLineageCursor.NextRow

                                If pLineageRow Is Nothing Then
                                    'Add a exception
                                    Dim pExpRow As IRow = tempExpTable.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review duplicate geometry"
                                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ssl
                                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                                    pExpRow.Store()

                                    pExpRow = tempUpdatepropertylineage.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURESSL")) = ssl
                                    If ownerPly.Fields.FindField("CREATION_D") >= 0 Then
                                        pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURECREATIONDT")) = ownerPly.Value(ownerPly.Fields.FindField("CREATION_D"))
                                    End If
                                    pExpRow.Value(pExpRow.Fields.FindField("CONFLICTFEATURESSL")) = pR.Value(pR.Fields.FindField("SSL"))
                                    If pR.Fields.FindField("CREATION_D") >= 0 Then
                                        pExpRow.Value(pExpRow.Fields.FindField("CONFLICTFEATURECREATIONDT")) = pR.Value(pR.Fields.FindField("CREATION_D"))
                                    End If

                                    pExpRow.Store()

                                    'Record should have been brought in by new/shift check
                                Else
                                    Dim pExpRow As IRow = tempExpTable.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ssl
                                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                                    pExpRow.Store()

                                End If

                                Marshal.ReleaseComObject(pLineageCursor)

                                'Check to see if the geometry exists in IAIS ownerPly
                                Dim pOwnerPlyN As IFeature
                                pFilter.WhereClause = "SSL='" & ssl & "'"
                                If iaisOwnerPlyFC.FeatureCount(pFilter) = 0 Then
                                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                                    pOwnerPlyN.Store()
                                End If

                                pFilter.WhereClause = "SSL='" & pR.Value(pR.Fields.FindField("SSL")) & "'"
                                If iaisOwnerPlyFC.FeatureCount(pFilter) = 0 Then
                                    If ownerPlyAddFC.FeatureCount(pFilter) = 0 Then
                                        pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                                        MapUtil.CopyFeature(pR, pOwnerPlyN, True)
                                        pOwnerPlyN.Store()
                                    End If
                                End If

                                idSet.Add(pR.OID)

                            End If



                            pR = OwnerPlyCursor.NextFeature
                        Loop

                        idSet.Add(ownerPly.OID)
                        Marshal.ReleaseComObject(OwnerPlyCursor)


                    Catch ex As Exception
                        log.WriteLine(Now & ": Error for ownerply. " & ownerPly.OID)
                    End Try

                End If



                counter = counter + 1
                If counter Mod 500 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pFreqRow = pFreqCursor.NextRow
            Loop

            'Geometries that do not intersect with IAIS ownerply
            'Check to see if they overlap themselves
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "OtrIAISPly, OtrIAISPly_Frequency"
            pQueryDef.SubFields = "OtrIAISPly.FID_OtrOwnerPly,OtrIAISPly.GIS_ID,OtrIAISPly.SSL"
            pQueryDef.WhereClause = "OtrIAISPly.FID_OtrOwnerPly = OtrIAISPly_Frequency.FID_OtrOwnerPly AND OtrIAISPly_Frequency.FREQUENCY = 1 AND OtrIAISPly.FID_IaisOwnerPly = -1"

            pFreqCursor = pQueryDef.Evaluate
            pFreqRow = pFreqCursor.NextRow

            Do While Not pFreqRow Is Nothing
                ownerPlyOId = pFreqRow.Value(0)

                idSet.Find(ownerPlyOId, bIdExists)

                If Not bIdExists Then
                    'OTR Ownerply
                    pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                    OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                    ownerPly = OwnerPlyCursor.NextFeature
                    Try
                        Dim ssl As String = ownerPly.Value(ownerPly.Fields.FindField("SSL"))

                        pSFilter.Geometry = ownerPly.Shape
                        pSFilter.WhereClause = otrOwnerPlyFC.OIDFieldName & " <> " & ownerPlyOId
                        pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                        Dim pOverlapPoly As IPolygon = ownerPly.Shape
                        Dim pRelOp As IRelationalOperator = pOverlapPoly
                        Dim pTopo As ITopologicalOperator

                        Marshal.ReleaseComObject(OwnerPlyCursor)

                        'Console.WriteLine("ownerPly OID:" & ownerPly.OID)

                        OwnerPlyCursor = otrOwnerPlyFC.Search(pSFilter, False)
                        Dim pR As IFeature = OwnerPlyCursor.NextFeature
                        Do While Not pR Is Nothing

                            If pRelOp.Equals(pR.Shape) Or _
                                pRelOp.Contains(pR.Shape) Or _
                                pRelOp.Within(pR.Shape) Then

                                log.WriteLine(Now & ": duplicate geometry found. " & ownerPly.OID)

                                'Check to see if this record is in lineage table 
                                pFilter.WhereClause = "(INPUTFEATURESSL='" & ssl & "' AND CONFLICTFEATURESSL='" & pR.Value(pR.Fields.FindField("SSL")) & "' ) OR " & _
                                                        "(CONFLICTFEATURESSL='" & ssl & "' AND INPUTFEATURESSL='" & pR.Value(pR.Fields.FindField("SSL")) & "' )"

                                pLineageCursor = tableLineage.Search(pFilter, False)
                                pLineageRow = pLineageCursor.NextRow

                                If pLineageRow Is Nothing Then
                                    'Add a exception
                                    Dim pExpRow As IRow = tempExpTable.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review duplicate geometry"
                                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ssl
                                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                                    pExpRow.Store()

                                    pExpRow = tempUpdatepropertylineage.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURESSL")) = ssl
                                    If ownerPly.Fields.FindField("CREATION_D") >= 0 Then
                                        pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURECREATIONDT")) = ownerPly.Value(ownerPly.Fields.FindField("CREATION_D"))
                                    End If
                                    pExpRow.Value(pExpRow.Fields.FindField("CONFLICTFEATURESSL")) = pR.Value(pR.Fields.FindField("SSL"))
                                    If pR.Fields.FindField("CREATION_D") >= 0 Then
                                        pExpRow.Value(pExpRow.Fields.FindField("CONFLICTFEATURECREATIONDT")) = pR.Value(pR.Fields.FindField("CREATION_D"))
                                    End If

                                    pExpRow.Store()

                                    'Record should have been brought in by new/shift check
                                Else
                                    Dim pExpRow As IRow = tempExpTable.CreateRow
                                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ssl
                                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                                    pExpRow.Store()

                                End If

                                Marshal.ReleaseComObject(pLineageCursor)

                                'Check to see if the geometry exists in IAIS ownerPly
                                Dim pOwnerPlyN As IFeature
                                pFilter.WhereClause = "SSL='" & ssl & "'"
                                If iaisOwnerPlyFC.FeatureCount(pFilter) = 0 Then
                                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                                    pOwnerPlyN.Store()
                                End If

                                pFilter.WhereClause = "SSL='" & pR.Value(pR.Fields.FindField("SSL")) & "'"
                                If iaisOwnerPlyFC.FeatureCount(pFilter) = 0 Then
                                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                                    MapUtil.CopyFeature(pR, pOwnerPlyN, True)
                                    pOwnerPlyN.Store()
                                End If

                                idSet.Add(pR.OID)
                            End If

                            pR = OwnerPlyCursor.NextFeature
                        Loop

                        idSet.Add(ownerPly.OID)
                        Marshal.ReleaseComObject(OwnerPlyCursor)


                    Catch ex As Exception
                        log.WriteLine(Now & ": Error for ownerply. " & ownerPly.OID)
                    End Try

                End If


                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pFreqRow = pFreqCursor.NextRow
            Loop


            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        Console.WriteLine(Now & " : end process")
        log.WriteLine(Now & ": Check duplicate geometry completed.")
    End Sub


#End Region


#Region "ProcessExceptions"

    Sub ProcessExceptions()
        log.WriteLine(Now & ": Create OwnerPly Exception job started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tableExceptionsOwnerPly As String
        Dim tablePropertyConfilictLineage As String

        Dim tableOtrOwnerPly As String
        Dim tableIAISOwnerPly As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSOWNERPLY" Then
                    tableExceptionsOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PROPERTYCONFILICTLINEAGE" Then
                    tablePropertyConfilictLineage = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OTROWNERPLYLAYER" Then
                    tableOtrOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "IAISOWNERPLYLAYER" Then
                    tableIAISOwnerPly = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)


        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("OwnerPlyExceptionTemp")

        Dim pFilter As IQueryFilter = New QueryFilter

        Dim pCursor As ICursor = tempExpTable.Search(Nothing, False)
        Dim pRow As IRow = pCursor.NextRow

        Dim updatedSSL As Dictionary(Of String, String) = New Dictionary(Of String, String)

        If Not pRow Is Nothing Then
            Dim pSdeWS As IFeatureWorkspace = CType(pJTXDatabase.DataWorkspace(m_jtxJob.VersionName), IFeatureWorkspace)
            Dim pWorkspaceEdit As IWorkspaceEdit = pSdeWS

            Dim pOwnerPlyExpTable As ITable = pSdeWS.OpenTable(tableExceptionsOwnerPly)
            Dim pOwnerPlyExpRow As IRow

            pWorkspaceEdit.StartEditing(False)
            pWorkspaceEdit.StartEditOperation()

            Try


                pCursor = tempExpTable.Search(Nothing, False)
                pRow = pCursor.NextRow

                Dim counter As Integer = 0
                Console.WriteLine(Now & " : starting process")

                Do While Not pRow Is Nothing

                    pOwnerPlyExpRow = pOwnerPlyExpTable.CreateRow
                    MapUtil.SetRowValue(pOwnerPlyExpRow, pOwnerPlyExpRow.Fields.FindField("EXCEPTDT"), Now)
                    MapUtil.SetRowValue(pOwnerPlyExpRow, pOwnerPlyExpRow.Fields.FindField("EXCEPTTYP"), MapUtil.GetRowValue(pRow, "ExceptionType"))
                    MapUtil.SetRowValue(pOwnerPlyExpRow, pOwnerPlyExpRow.Fields.FindField("OWNER_GIS_ID"), MapUtil.GetRowValue(pRow, "OWNER_GIS_ID"))
                    MapUtil.SetRowValue(pOwnerPlyExpRow, pOwnerPlyExpRow.Fields.FindField("SSL"), MapUtil.GetRowValue(pRow, "SSL"))

                    pOwnerPlyExpRow.Store()

                    counter = counter + 1
                    If counter Mod 100 = 0 Then
                        Console.WriteLine(Now & " : processed " & counter & " records")
                    End If

                    pRow = pCursor.NextRow
                Loop

                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)

                log.WriteLine(Now & ": " & counter & " exceptions were added.")
            Catch ex As Exception
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(False)
                Throw ex
            End Try

        End If

        Console.WriteLine(Now & " : end process")
        log.WriteLine(Now & ": Create OwnerPly Exception job completed.")

    End Sub
#End Region

#Region "Check OwnerPoly Attribute Frequency=1 and OTR_OwnerPolyId >=0 "
    '1. No Geomtry change
    Sub CheckOwnerPolyAttribute(Optional ByVal startRecno As Integer = -1, Optional ByVal checkpoint As Integer = 5000)
        log.WriteLine(Now & ": Check OwnerPoly Attribute started.")



        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")


        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")

        Dim ownerPlyOId As Integer

        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()
        Dim counter As Integer = 0

        Dim exceptionCount As Integer = 0

        Try

            Dim pFilter As IQueryFilter = New QueryFilter

            Dim pQueryDef As IQueryDef
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "IAISOtrPly, IAISOtrPly_Frequency, OtrIAISPly_Frequency"
            pQueryDef.SubFields = "*"
            pQueryDef.WhereClause = "IAISOtrPly.FID_IaisOwnerPly = IAISOtrPly_Frequency.FID_IaisOwnerPly AND IAISOtrPly.FID_OTROwnerPly=OtrIAISPly_Frequency.FID_OTROwnerPly AND IAISOtrPly_Frequency.FREQUENCY = 1 AND OtrIAISPly_Frequency.FREQUENCY = 1"

            Dim pIdentityCursor As ICursor = pQueryDef.Evaluate()
            Dim pIdentityRow As IRow = pIdentityCursor.NextRow


            Console.WriteLine(Now & " : starting process")

            Dim idxOidField As Integer = pIdentityCursor.FindField("FID_OtrOwnerPly")
            Do While Not pIdentityRow Is Nothing

                If counter > startRecno Then
                    ownerPlyOId = pIdentityRow.Value(idxOidField)

                    pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                    Dim pCursor As IFeatureCursor = otrOwnerPlyFC.Search(pFilter, False)
                    'Dim pRow = otrOwnerPlyFC.GetFeature(ownerPlyOId) 'pCursor.NextFeature
                    Dim pRow = pCursor.NextFeature
                    Dim bSSLFlag As Boolean = False
                    If compare(pRow, pIdentityRow, "Shape_Area", True) Then
                        'The Geometry has been reshaped.

                        Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                        MapUtil.CopyFeature(pIdentityRow, pOwnerPlyN, True)
                        pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = pIdentityRow.Value(pIdentityRow.Fields.FindField("FID_IaisOwnerPly"))
                        pOwnerPlyN.Store()

                        pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                        MapUtil.CopyFeature(pRow, pOwnerPlyN, True)
                        pOwnerPlyN.Store()

                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = pRow.Value(pRow.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = pRow.Value(pRow.Fields.FindField("SSL"))
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                        pExpRow.Store()

                        If compare(pRow, pIdentityRow, "SSL") Then
                            pExpRow = tempExpTable.CreateRow
                            pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review SSL change"
                            pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = pRow.Value(pRow.Fields.FindField("GIS_ID"))
                            pExpRow.Value(pExpRow.Fields.FindField("SSL")) = pRow.Value(pRow.Fields.FindField("SSL"))
                            pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"
                            pExpRow.Store()
                        End If


                    ElseIf (compare(pRow, pIdentityRow, "SSL") Or _
                        compare(pRow, pIdentityRow, "SQUARE") Or _
                        compare(pRow, pIdentityRow, "SUFFIX") Or _
                        compare(pRow, pIdentityRow, "LOT")) Then

                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review SSL change"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = pRow.Value(pRow.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = pRow.Value(pRow.Fields.FindField("SSL"))
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"
                        pExpRow.Store()

                        'Console.WriteLine("Review SSL change " & pRow.OID)


                        'Push pRow into ownerplyupdate layer
                        Dim pOwnerPlyN As IFeature = ownerPlyUpdateFC.CreateFeature()
                        MapUtil.CopyFeature(pRow, pOwnerPlyN, True)
                        pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = pIdentityRow.Value(pIdentityRow.Fields.FindField("FID_IaisOwnerPly"))
                        pOwnerPlyN.Store()

                        exceptionCount = exceptionCount + 1
                    ElseIf (compare(pRow, pIdentityRow, "PAR") Or _
                            compare(pRow, pIdentityRow, "LOT_TYPE") Or _
                            compare(pRow, pIdentityRow, "PROPTYPE") Or _
                            compare(pRow, pIdentityRow, "USECODE") Or _
                            compare(pRow, pIdentityRow, "PREMISEADD") Or _
                            compare(pRow, pIdentityRow, "OWNERNAME") Or _
                            compare(pRow, pIdentityRow, "ADDRESS1") Or _
                            compare(pRow, pIdentityRow, "ADDRESS2") Or _
                            compare(pRow, pIdentityRow, "CITYSTZIP")) Then

                        'Push pRow into ownerplyupdate layer
                        Dim pOwnerPlyN As IFeature = ownerPlyUpdateFC.CreateFeature()
                        MapUtil.CopyFeature(pRow, pOwnerPlyN, True)
                        pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = pIdentityRow.Value(pIdentityRow.Fields.FindField("FID_IaisOwnerPly"))
                        pOwnerPlyN.Store()

                    End If

                    Marshal.ReleaseComObject(pCursor)
                End If

                    counter = counter + 1
                    If counter Mod 1000 = 0 Then
                        Console.WriteLine(Now & " : processed " & counter & " records")
                        If counter Mod checkpoint = 0 Then
                            pWorkspaceEdit.StopEditOperation()
                            pWorkspaceEdit.StopEditing(True)

                            Console.WriteLine(Now & ": Checkpoint " & counter)
                            log.WriteLine(Now & ": Checkpoint " & counter)

                            pWorkspaceEdit.StartEditing(False)
                            pWorkspaceEdit.StartEditOperation()
                        End If
                    End If

                    pIdentityRow = pIdentityCursor.NextRow

            Loop



            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & exceptionCount & " were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        Console.WriteLine(Now & " : end process")
        log.WriteLine(Now & ": Check OwnerPoly Attribute completed.")
    End Sub


    Function compare(ByVal newRow As IRow, ByVal oldRow As IRow, ByVal fieldName As String, Optional ByVal bAreaTolerance As Boolean = False, Optional ByVal tolerance As Double = 1) As Boolean
        Dim oldValue As Object
        Dim newValue As Object

        Try
            oldValue = oldRow.Value(oldRow.Fields.FindField(fieldName))
            newValue = newRow.Value(newRow.Fields.FindField(fieldName))

            If IsDBNull(oldValue) AndAlso Not IsDBNull(newValue) Then
                If (newValue.ToString().Trim = "") Then
                    Return False
                Else
                    'Console.WriteLine(fieldName & vbTab & oldValue & ":" & newValue)
                    Return True
                End If
            End If

            If Not IsDBNull(oldValue) AndAlso IsDBNull(newValue) Then
                If (oldValue.ToString().Trim = "") Then
                    Return False
                Else
                    'Console.WriteLine(fieldName & vbTab & oldValue & ":" & newValue)
                    Return True
                End If
            End If

            If bAreaTolerance Then
                If (oldValue - newValue) > tolerance Then
                    'Console.WriteLine(fieldName & vbTab & oldValue & ":" & newValue)
                    Return True
                Else
                    Return False
                End If
            ElseIf oldValue.ToString().CompareTo(newValue.ToString()) <> 0 Then
                'Console.WriteLine(fieldName & vbTab & oldValue.ToString() & ":" & newValue.ToString())
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
            Throw ex
        End Try
    End Function

    Function IsRowChanged(ByVal newRow As IRow, ByVal oldRow As IRow, ByRef sslChange As Boolean) As Boolean
        Dim oldValue As Object
        Dim newValue As Object

        Dim i As Integer
        Dim fIndex As Integer

        Try


            If newRow.Value(newRow.Fields.FindField("SSL")) <> oldRow.Value(oldRow.Fields.FindField("SSL")) Then
                sslChange = True
                Return True
            Else
                Dim pField As IField
                For i = 0 To newRow.Fields.FieldCount - 1
                    pField = newRow.Fields.Field(i)

                    If pField.Editable AndAlso Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                        fIndex = oldRow.Fields.FindField(pField.Name)
                        If fIndex >= 0 Then
                            oldValue = newRow.Value(i)
                            newValue = oldRow.Value(fIndex)


                            If IsDBNull(oldValue) AndAlso Not IsDBNull(newValue) Then
                                Return True
                            End If

                            If Not IsDBNull(oldValue) AndAlso IsDBNull(newValue) Then
                                Return True
                            End If


                            If TypeOf oldValue Is IComparable Then
                                If oldValue.ToString().CompareTo(newValue.ToString()) <> 0 Then
                                    'Console.WriteLine(newRow.OID & "\" & pField.Name & "   " & oldValue.ToString() & ":" & newValue.ToString())
                                    Return True
                                End If
                            End If
                        End If
                    End If
                Next

                Return False
            End If

        Catch ex As Exception
            'MsgBox(ex.Message)
            Throw ex
        End Try


    End Function

#End Region

#Region "CheckLocationException frequency > 1"
    Sub CheckLocationException()
        log.WriteLine(Now & ": Check Location Exception started.")


        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)
        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")

        Dim tempUpdatepropertylineage As ITable = pTempWorkspace.OpenTable("updatepropertylineage")

        Console.WriteLine(Now & " : Opening OTR ownerply")
        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Console.WriteLine(Now & " : Opening IAIS ownerply")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")

        Console.WriteLine(Now & " : Opening Lineage table")
        Dim tableLineage As ITable = pTempWorkspace.OpenTable("PropertyConfilictLineage")

        Dim ownerPlyOId As Integer
        Dim ownerPlyOldSSL As String

        Dim ownerPly As IFeature
        Dim OwnerPlyOld As IFeature
        Dim OwnerPlyCursor As IFeatureCursor

        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim counter As Integer = 0
        Dim exceptionCount As Integer = 0

        Console.WriteLine(Now & " : starting process")

        Try

            Dim pFilter As IQueryFilter = New QueryFilter
            pFilter.WhereClause = "FREQUENCY > 1"
            Dim pFreqCursor As ICursor = pTempWorkspace.OpenTable("IAISOtrPly_Frequency").Search(pFilter, False)
            Dim pFreqRow As IRow = pFreqCursor.NextRow

            Dim idxOidField As Integer = pFreqCursor.FindField("FID_IaisOwnerPly")
            Do While Not pFreqRow Is Nothing
                ownerPlyOId = pFreqRow.Value(idxOidField)
                'IAIS Ownerply
                pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                OwnerPlyCursor = iaisOwnerPlyFC.Search(pFilter, False)
                OwnerPlyOld = OwnerPlyCursor.NextFeature
                ownerPlyOldSSL = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("SSL"))

                pFilter.WhereClause = "SSL='" & ownerPlyOldSSL & "'"
                If otrOwnerPlyFC.FeatureCount(pFilter) > 0 Then
                    '1. polygon has been moved. 
                    '2. A Overlap polygon has been created 
                    '3. Both 1 & 2
                    pFilter.WhereClause = "INPUTFEATURESSL='" & ownerPlyOldSSL & "' OR CONFLICTFEATURESSL='" & ownerPlyOldSSL & "'"
                    If tableLineage.RowCount(pFilter) > 0 Then
                        'Review lineage and Geometry revisions
                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                        pExpRow.Store()
                        exceptionCount = exceptionCount + 1
                    Else
                        'Review polygon shift
                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review polygon shift"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                        pExpRow.Store()
                        exceptionCount = exceptionCount + 1
                    End If

                    'Delete the ownerpoly feature from current layer
                    'and add the corresponding feature from new layer

                    Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                    MapUtil.CopyFeature(OwnerPlyOld, pOwnerPlyN, True)
                    pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = OwnerPlyOld.OID
                    pOwnerPlyN.Store()

                    Marshal.ReleaseComObject(OwnerPlyCursor)

                    pFilter.WhereClause = "SSL='" & ownerPlyOldSSL & "'"
                    OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                    ownerPly = OwnerPlyCursor.NextFeature

                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                    pOwnerPlyN.Store()

                Else
                    'Polygon has been deleted
                    Dim pExpRow As IRow = tempExpTable.CreateRow
                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Existing polygon was removed"
                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                    pExpRow.Store()
                    exceptionCount = exceptionCount + 1

                    pFilter.WhereClause = "INPUTFEATURESSL='" & ownerPlyOldSSL & "' OR CONFLICTFEATURESSL='" & ownerPlyOldSSL & "'"
                    If tableLineage.RowCount(pFilter) > 0 Then
                        'Review lineage and Geometry revisions
                        pExpRow = tempUpdatepropertylineage.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURESSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("PROCESS_TYPE")) = "Delete"
                        pExpRow.Store()
                    End If

                    Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                    MapUtil.CopyFeature(OwnerPlyOld, pOwnerPlyN, True)
                    pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = OwnerPlyOld.OID
                    pOwnerPlyN.Store()

                End If

                Marshal.ReleaseComObject(OwnerPlyCursor)
                counter = counter + 1
                If counter Mod 500 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pFreqRow = pFreqCursor.NextRow
            Loop


            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & exceptionCount & " were added.")
        Catch ex As Exception

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        Console.WriteLine(Now & " : end process")
        log.WriteLine(Now & ": Check Location Exception completed.")
    End Sub

    ' Lots have been merged.
    Sub CheckOtrLocationException()
        log.WriteLine(Now & ": Check Otr LocationException started.")

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")


        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")

        Dim ownerPlyOId As Integer

        Dim ownerPly As IFeature
        Dim OwnerPlyOld As IFeature
        Dim OwnerPlyCursor As IFeatureCursor


        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()
        Dim counter As Integer = 0

        Dim exceptionCount As Integer = 0

        Try

            Dim pFilter As IQueryFilter = New QueryFilter

            Dim pQueryDef As IQueryDef
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "IAISOtrPly, IAISOtrPly_Frequency, OtrIAISPly_Frequency"
            pQueryDef.SubFields = "*"
            pQueryDef.WhereClause = "IAISOtrPly.FID_IaisOwnerPly = IAISOtrPly_Frequency.FID_IaisOwnerPly AND IAISOtrPly.FID_OTROwnerPly=OtrIAISPly_Frequency.FID_OTROwnerPly AND IAISOtrPly_Frequency.FREQUENCY = 1 AND OtrIAISPly_Frequency.FREQUENCY > 1"

            Dim pIdentityCursor As ICursor = pQueryDef.Evaluate()
            Dim pIdentityRow As IRow = pIdentityCursor.NextRow


            Console.WriteLine(Now & " : starting process")

            Dim idSet As IFIDSet = New FIDSet

            Dim idxOidField As Integer = pIdentityCursor.FindField("FID_IaisOwnerPly")
            Do While Not pIdentityRow Is Nothing
                ownerPlyOId = pIdentityRow.Value(idxOidField)
                pFilter.WhereClause = "OBJECTID=" & ownerPlyOId

                OwnerPlyCursor = iaisOwnerPlyFC.Search(pFilter, False)
                OwnerPlyOld = OwnerPlyCursor.NextFeature

                'Delete the old geometry
                Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                MapUtil.CopyFeature(OwnerPlyOld, pOwnerPlyN, True)
                pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = OwnerPlyOld.OID
                pOwnerPlyN.Store()

                Dim pExpRow As IRow = tempExpTable.CreateRow
                pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Existing polygon was removed"
                pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                pExpRow.Value(pExpRow.Fields.FindField("SSL")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("SSL"))
                pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                pExpRow.Store()

                exceptionCount = exceptionCount + 1


                'Add new merged geometry 
                Dim bIdExists As Boolean
                idSet.Find(pIdentityRow.Value(pIdentityCursor.FindField("FID_OTROwnerPly")), bIdExists)
                If Not bIdExists Then
                    Marshal.ReleaseComObject(OwnerPlyCursor)

                    pFilter.WhereClause = "OBJECTID=" & pIdentityRow.Value(pIdentityCursor.FindField("FID_OTROwnerPly"))
                    OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                    ownerPly = OwnerPlyCursor.NextFeature

                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                    pOwnerPlyN.Store()
                End If

                Marshal.ReleaseComObject(OwnerPlyCursor)
                counter = counter + 1
                If counter Mod 500 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pIdentityRow = pIdentityCursor.NextRow
            Loop



            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & exceptionCount & " were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        Console.WriteLine(Now & " : end process")
        log.WriteLine(Now & ": Check Otr LocationException completed.")
    End Sub


#End Region

#Region "CheckNoIntersection frequency = 1 and fid_ownerply_1 = -1"
    Sub CheckNoIntersection()
        log.WriteLine(Now & ": Check polygons without intersection started.")


        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")

        Dim tempUpdatepropertylineage As ITable = pTempWorkspace.OpenTable("updatepropertylineage")


        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")
        Dim tableLineage As ITable = pTempWorkspace.OpenTable("PropertyConfilictLineage")

        Dim ownerPlyOId As Integer
        Dim ownerPlyOldSSL As String

        Dim ownerPly As IFeature
        Dim OwnerPlyOld As IFeature
        Dim OwnerPlyCursor As IFeatureCursor


        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try

            Dim pFilter As IQueryFilter = New QueryFilter

            Dim pQueryDef As IQueryDef
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "IAISOtrPly, IAISOtrPly_Frequency"
            pQueryDef.SubFields = "IAISOtrPly.FID_IaisOwnerPly"
            pQueryDef.WhereClause = "IAISOtrPly.FID_IaisOwnerPly = IAISOtrPly_Frequency.FID_IaisOwnerPly AND IAISOtrPly_Frequency.FREQUENCY = 1 AND IAISOtrPly.FID_OtrOwnerPly = -1"

            Dim pFreqCursor As ICursor = pQueryDef.Evaluate
            Dim pFreqRow As IRow = pFreqCursor.NextRow

            Dim counter As Integer = 0
            Dim exceptionCount As Integer = 0

            Console.WriteLine(Now & " : starting process")



            Do While Not pFreqRow Is Nothing
                ownerPlyOId = pFreqRow.Value(0)
                'IAIS Ownerply
                pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                OwnerPlyCursor = iaisOwnerPlyFC.Search(pFilter, False)
                OwnerPlyOld = OwnerPlyCursor.NextFeature
                ownerPlyOldSSL = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("SSL"))

                pFilter.WhereClause = "SSL='" & ownerPlyOldSSL & "'"
                If otrOwnerPlyFC.FeatureCount(pFilter) > 0 Then
                    'polygon has been moved
                    pFilter.WhereClause = "INPUTFEATURESSL='" & ownerPlyOldSSL & "' OR CONFLICTFEATURESSL='" & ownerPlyOldSSL & "'"
                    If tableLineage.RowCount(pFilter) > 0 Then
                        'Review lineage and Geometry revisions
                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                        pExpRow.Store()

                        exceptionCount = exceptionCount + 1
                    Else
                        'Review polygon shift
                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review polygon shift"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                        pExpRow.Store()
                        exceptionCount = exceptionCount + 1
                    End If

                    Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                    MapUtil.CopyFeature(OwnerPlyOld, pOwnerPlyN, True)
                    pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = OwnerPlyOld.OID
                    pOwnerPlyN.Store()

                    Marshal.ReleaseComObject(OwnerPlyCursor)

                    pFilter.WhereClause = "SSL='" & ownerPlyOldSSL & "'"
                    OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                    ownerPly = OwnerPlyCursor.NextFeature

                    pOwnerPlyN = ownerPlyAddFC.CreateFeature()
                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                    pOwnerPlyN.Store()

                Else
                    'Polygon has been deleted
                    Dim pExpRow As IRow = tempExpTable.CreateRow
                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Existing polygon was removed"
                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = OwnerPlyOld.Value(OwnerPlyOld.Fields.FindField("GIS_ID"))
                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlyOldSSL
                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "O"

                    pExpRow.Store()

                    pFilter.WhereClause = "INPUTFEATURESSL='" & ownerPlyOldSSL & "' OR CONFLICTFEATURESSL='" & ownerPlyOldSSL & "'"
                    If tableLineage.RowCount(pFilter) > 0 Then
                        'Review lineage and Geometry revisions
                        pExpRow = tempUpdatepropertylineage.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("INPUTFEATURESSL")) = ownerPlyOldSSL
                        pExpRow.Value(pExpRow.Fields.FindField("PROCESS_TYPE")) = "Delete"
                        pExpRow.Store()

                        exceptionCount = exceptionCount + 1
                    End If

                    'Dim pOwnerPlyN As IFeature = ownerPlyDeleteFC.CreateFeature()
                    'MapUtil.CopyFeature(OwnerPlyOld, pOwnerPlyN, True)
                    'pOwnerPlyN.Value(pOwnerPlyN.Fields.FindField("IAISOWNERPLY_FID")) = OwnerPlyOld.OID
                    'pOwnerPlyN.Store()

                End If

                Marshal.ReleaseComObject(OwnerPlyCursor)

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If


                pFreqRow = pFreqCursor.NextRow
            Loop


            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            Console.WriteLine(Now & " : End process")

            log.WriteLine(Now & ": " & exceptionCount & " were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try


        log.WriteLine(Now & ": Check polygons without intersection completed.")
    End Sub
#End Region

#Region "CheckOtrNoIntersection OTR IAIS frequency <= 2 and fid_ownerply_1 = -1"
    Sub CheckOtrNoIntersection()
        log.WriteLine(Now & ": Check polygons without intersection started (OTR IAIS Identity).")

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")

        Dim otrOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("OtrOwnerPly")
        Dim iaisOwnerPlyFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("IAISOwnerPly")


        Dim tableLineage As ITable = pTempWorkspace.OpenTable("PropertyConfilictLineage")

        Dim ownerPlyOId As Integer
        Dim ownerPlySSL As String

        Dim ownerPly As IFeature
        Dim OwnerPlyCursor As IFeatureCursor


        Dim pWorkspaceEdit As IWorkspaceEdit = pTempWorkspace
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try

            Dim pFilter As IQueryFilter = New QueryFilter

            Dim pQueryDef As IQueryDef
            pQueryDef = pTempWorkspace.CreateQueryDef
            pQueryDef.Tables = "OtrIAISPly, OtrIAISPly_Frequency"
            pQueryDef.SubFields = "OtrIAISPly.FID_OtrOwnerPly,OtrIAISPly.GIS_ID,OtrIAISPly.SSL"
            pQueryDef.WhereClause = "OtrIAISPly.FID_OtrOwnerPly = OtrIAISPly_Frequency.FID_OtrOwnerPly AND OtrIAISPly_Frequency.FREQUENCY = 1 AND OtrIAISPly.FID_IaisOwnerPly = -1"

            Dim pFreqCursor As ICursor = pQueryDef.Evaluate
            Dim pFreqRow As IRow = pFreqCursor.NextRow


            Dim counter As Integer = 0
            Dim exceptionCount As Integer = 0

            Console.WriteLine(Now & " : starting process")

            Do While Not pFreqRow Is Nothing
                ownerPlyOId = pFreqRow.Value(0)
                pFilter.WhereClause = "OBJECTID=" & ownerPlyOId
                OwnerPlyCursor = otrOwnerPlyFC.Search(pFilter, False)
                ownerPly = OwnerPlyCursor.NextFeature
                ownerPlySSL = ownerPly.Value(ownerPly.Fields.FindField("SSL"))

                pFilter.WhereClause = "SSL='" & ownerPlySSL & "'"
                If iaisOwnerPlyFC.FeatureCount(pFilter) > 0 Then
                    'New geomtry for existing SSL.
                    pFilter.WhereClause = "INPUTFEATURESSL='" & ownerPlySSL & "' OR CONFLICTFEATURESSL='" & ownerPlySSL & "'"
                    If tableLineage.RowCount(pFilter) > 0 Then
                        'Review lineage and Geometry revisions CheckNoIntersection
                        Dim pExpRow As IRow = tempExpTable.CreateRow
                        pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review lineage and Geometry revisions"
                        pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                        pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlySSL
                        pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                        pExpRow.Store()
                    Else
                        'If this has not been cought by 
                        pFilter.WhereClause = "SSL='" & ownerPlySSL & "'"
                        If ownerPlyAddFC.FeatureCount(pFilter) = 0 Then
                            'Review polygon shift
                            Dim pExpRow As IRow = tempExpTable.CreateRow
                            pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "Review polygon shift"
                            pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                            pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlySSL
                            pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                            pExpRow.Store()

                            Dim pOwnerPlyN As IFeature = ownerPlyAddFC.CreateFeature()
                            MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                            pOwnerPlyN.Store()

                        End If
                    End If

                    Marshal.ReleaseComObject(OwnerPlyCursor)

                Else
                    'This is a new polygon
                    Dim pExpRow As IRow = tempExpTable.CreateRow
                    pExpRow.Value(pExpRow.Fields.FindField("ExceptionType")) = "New polygon"
                    pExpRow.Value(pExpRow.Fields.FindField("OWNER_GIS_ID")) = ownerPly.Value(ownerPly.Fields.FindField("GIS_ID"))
                    pExpRow.Value(pExpRow.Fields.FindField("SSL")) = ownerPlySSL
                    pExpRow.Value(pExpRow.Fields.FindField("REFDATASERC")) = "N"

                    pExpRow.Store()
                    exceptionCount = exceptionCount + 1

                    Dim pOwnerPlyN As IFeature = ownerPlyAddFC.CreateFeature()
                    MapUtil.CopyFeature(ownerPly, pOwnerPlyN, True)
                    pOwnerPlyN.Store()

                End If

                Marshal.ReleaseComObject(OwnerPlyCursor)

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pFreqRow = pFreqCursor.NextRow
            Loop


            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & exceptionCount & " were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try


        log.WriteLine(Now & ": Check polygons without intersection completed.")
    End Sub
#End Region
#Region "JTX Job"

    Function GetJTXJob() As IJTXJob
        Dim jobid As String = ""

        If File.Exists(Environment.CurrentDirectory & "\jtxjob.txt") Then
            Dim oRead As System.IO.StreamReader = File.OpenText(Environment.CurrentDirectory & "\jtxjob.txt")
            jobid = oRead.ReadLine
        Else
            log.WriteLine(Now & ": JTX job has not been created")
            Return Nothing
        End If

        Try
            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode

            Dim jtxDB As String
            Dim jtxDBWorkspace As String

            m_xmld = New XmlDocument()
            m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

            m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

            If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
                m_node = m_nodelist.Item(0)
                For i = 0 To m_node.ChildNodes.Count - 1
                    If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                        jtxDB = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                        jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                    End If
                Next i
            End If

            Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
            Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
            Dim pJTXJobMan As IJTXJobManager = pJTXDatabase.JobManager

            Return pJTXJobMan.GetJob(jobid)
        Catch ex As Exception
            log.WriteLine(Now & ": Can't find exception job with job ID " & jobid)
            Throw ex
        End Try

    End Function

    Sub CreateOwnerPlyJTXJob()
        log.WriteLine(Now & ": Create OwnerPly WEEKLY UPDATE Exception job started.")

        If File.Exists(Environment.CurrentDirectory & "\jtxjob.txt") Then
            Dim oRead As System.IO.StreamReader = File.OpenText(Environment.CurrentDirectory & "\jtxjob.txt")
            Dim jobid As String = oRead.ReadLine
            If Trim(jobid) <> "" Then
                log.WriteLine(Now & ": JTX job " & jobid & "exists")
                Return
            End If
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String
        Dim jtxJobParentVersion As String
        Dim exceptionType As String

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OWNERPLY_EXCEPTION_TYPE" Then
                    exceptionType = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

        m_jtxJob = CreateJob(pJTXDatabase, exceptionType, jtxDBWorkspace, jtxJobParentVersion)

        Console.WriteLine("JTX job was created. JTX job ID " & m_jtxJob.ID)
        log.WriteLine(Now & ":JTX job was created. JTX job ID " & m_jtxJob.ID)

        Try
            Dim oWriter As System.IO.StreamWriter = File.CreateText(Environment.CurrentDirectory & "\jtxjob.txt")
            oWriter.WriteLine(m_jtxJob.ID)
            oWriter.Close()
        Catch ex As Exception
            If Not m_jtxJob Is Nothing Then
                m_jtxJob.DeleteVersion()
                pJTXDatabase.JobManager.DeleteJob(m_jtxJob.ID)
                Console.WriteLine("JTX job was deleted because of system exception")
                log.WriteLine(Now & ":JTX job was deleted because of system exception")
            End If
            Throw ex
        End Try

        log.WriteLine(Now & ": Create OwnerPly WEEKLY UPDATE Exception job completed. New Job ID " & m_jtxJob.ID)
    End Sub

    Function CreateJob(ByRef pJTXDatabase As IJTXDatabase, ByVal jobTypeName As String, ByVal workspaceName As String, ByVal parentVersion As String) As IJTXJob

        Dim pJTXJobMan As IJTXJobManager = Nothing
        Dim pNewJob As IJTXJob2 = Nothing

        Try

            pJTXJobMan = pJTXDatabase.JobManager

            Dim jconfigMan As JTXConfigurationManager
            jconfigMan = pJTXDatabase.ConfigurationManager

            Dim pJtxJobTypePremiseExceptions As IJTXJobType = jconfigMan.GetJobType(jobTypeName)
            If pJtxJobTypePremiseExceptions Is Nothing Then
                Throw New Exception("Can't find job type (" & jobTypeName & ")")
            End If
            pNewJob = pJTXJobMan.CreateJob(pJtxJobTypePremiseExceptions, 0, True)

            Dim jtxWorkflowUtil As IJTXWorkflowUtilities = New JTXUtility
            Dim strWorkflowXML As String = jtxWorkflowUtil.GetWorkflowXML(pJTXDatabase, pJtxJobTypePremiseExceptions)
            If strWorkflowXML <> "" Then
                jtxWorkflowUtil.CopyWorkflowXML(pJTXDatabase, pNewJob)
            End If

            Dim pActType As IJTXActivityType = jconfigMan.GetActivityType("CreateJob")
            If Not pActType Is Nothing Then
                pNewJob.LogJobAction(pActType, Nothing, "")
            End If

            Dim pDataWorkspace As IJTXWorkspaceConfiguration = GetDbConnection(workspaceName, pJTXDatabase)
            pNewJob.SetActiveDatabase(pDataWorkspace.DatabaseID)
            pNewJob.ParentVersion = parentVersion

            pNewJob.Status = jconfigMan.GetStatus("ReadyToWork")
            pNewJob.Store()

            pNewJob.CreateVersion(esriVersionAccess.esriVersionAccessProtected)

            Dim jtxSystemUtil As IJTXSystemUtilities = New JTXUtility
            jtxSystemUtil.SendNotification("JobCreated", pJTXDatabase, pNewJob, Nothing)

            Return pNewJob

        Catch ex As Exception
            If Not pNewJob Is Nothing Then
                pNewJob.DeleteVersion()
                pJTXJobMan.DeleteJob(pNewJob.ID, True)
            End If

            Throw ex
        End Try

    End Function

    Function GetDbConnection(ByVal dbname As String, ByVal pJTXDatabase As IJTXDatabase) As IJTXWorkspaceConfiguration
        Dim pDBConnManager As IJTXDatabaseConnectionManager = New JTXDatabaseConnectionManager
        Dim pDBConnection As IJTXDatabaseConnection = pDBConnManager.GetConnection(pJTXDatabase.Alias)

        Dim i As Integer
        For i = 0 To pDBConnection.DataWorkspaceCount - 1
            Dim pDataWorkspace As IJTXWorkspaceConfiguration = pDBConnection.DataWorkspace(i)
            If pDataWorkspace.Name = dbname Then
                Return pDataWorkspace
            End If
        Next

        Return Nothing

    End Function

#End Region

#Region "Process changes from ownerpolyadd"
    Sub ProcessOwnerplyAdd()
        log.WriteLine(Now & ": Process changes from ownerpolyadd started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tableExceptionsOwnerPly As String
        Dim tablePropertyConfilictLineage As String

        Dim tableOtrOwnerPly As String
        Dim tableIAISOwnerPly As String

        Dim jtxJobParentVersion As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PROPERTYCONFILICTLINEAGE" Then
                    tablePropertyConfilictLineage = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OTROWNERPLYLAYER" Then
                    tableOtrOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "IAISOWNERPLYLAYER" Then
                    tableIAISOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
        Dim pSdeWS As IFeatureWorkspace = CType(pJTXDatabase.DataWorkspace(m_jtxJob.VersionName), IFeatureWorkspace)

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        'Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        'Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")


        'Dim otrOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableOtrOwnerPly)
        Dim iaisOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableIAISOwnerPly)


        'Dim tableLineage As ITable = pSdeWS.OpenTable(tablePropertyConfilictLineage)

        Dim pWorkspaceEdit As IWorkspaceEdit = pSdeWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try


            Dim pCursor As ICursor = ownerPlyAddFC.Search(Nothing, False)
            Dim pRow As IRow = pCursor.NextRow
            Dim counter As Integer = 0
            Console.WriteLine(Now & " : starting process")

            Do While Not pRow Is Nothing

                Dim pOwnerPlyN As IFeature = iaisOwnerPlyFC.CreateFeature()
                MapUtil.CopyFeature(pRow, pOwnerPlyN, True)
                pOwnerPlyN.Store()

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pRow = pCursor.NextRow
            Loop

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counter & " ownerply were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Process changes from ownerpolyadd completed.")
    End Sub
#End Region

#Region "Process changes from ownerpolydelete"
    Sub ProcessOwnerplyDelete()
        log.WriteLine(Now & ": Process changes from ownerpolydelete started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tablePropertyConfilictLineage As String

        Dim tableOtrOwnerPly As String
        Dim tableIAISOwnerPly As String

        Dim jtxJobParentVersion As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PROPERTYCONFILICTLINEAGE" Then
                    tablePropertyConfilictLineage = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OTROWNERPLYLAYER" Then
                    tableOtrOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "IAISOWNERPLYLAYER" Then
                    tableIAISOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
        Dim pSdeWS As IFeatureWorkspace = CType(pJTXDatabase.DataWorkspace(m_jtxJob.VersionName), IFeatureWorkspace)

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        'Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        'Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        'Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")


        'Dim otrOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableOtrOwnerPly)
        Dim iaisOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableIAISOwnerPly)


        'Dim tableLineage As ITable = pSdeWS.OpenTable(tablePropertyConfilictLineage)

        Dim pWorkspaceEdit As IWorkspaceEdit = pSdeWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try


            Dim pCursor As ICursor = ownerPlyDeleteFC.Search(Nothing, False)
            Dim pRow As IRow = pCursor.NextRow
            Dim pFilter As IQueryFilter = New QueryFilter

            Dim counter As Integer = 0
            Console.WriteLine(Now & " : starting process")

            Do While Not pRow Is Nothing
                pFilter.WhereClause = "OBJECTID=" & pRow.Value(pRow.Fields.FindField("IAISOWNERPLY_FID"))
                Dim ownerPlyCursor As IFeatureCursor = iaisOwnerPlyFC.Search(pFilter, False)
                Dim pOwnerPly As IFeature = ownerPlyCursor.NextFeature
                If Not pOwnerPly Is Nothing Then
                    pOwnerPly.Delete()
                End If

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                Marshal.ReleaseComObject(ownerPlyCursor)

                pRow = pCursor.NextRow
            Loop

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counter & " ownerply were deleted.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Process changes from ownerpolydelete completed.")
    End Sub
#End Region

#Region "Process changes from ownerpolyupdate"
    Sub ProcessOwnerplyUpdate()
        log.WriteLine(Now & ": Process changes from ownerpolyupdate started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tablePropertyConfilictLineage As String

        Dim tableOtrOwnerPly As String
        Dim tableIAISOwnerPly As String

        Dim jtxJobParentVersion As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PROPERTYCONFILICTLINEAGE" Then
                    tablePropertyConfilictLineage = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OTROWNERPLYLAYER" Then
                    tableOtrOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "IAISOWNERPLYLAYER" Then
                    tableIAISOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
        Dim pSdeWS As IFeatureWorkspace = CType(pJTXDatabase.DataWorkspace(m_jtxJob.VersionName), IFeatureWorkspace)

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        'Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        'Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        'Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")


        'Dim otrOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableOtrOwnerPly)
        Dim iaisOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableIAISOwnerPly)


        'Dim tableLineage As ITable = pSdeWS.OpenTable(tablePropertyConfilictLineage)

        Dim pWorkspaceEdit As IWorkspaceEdit = pSdeWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try


            Dim pCursor As ICursor = ownerPlyUpdateFC.Search(Nothing, False)
            Dim pRow As IRow = pCursor.NextRow
            Dim pFilter As IQueryFilter = New QueryFilter

            Dim counter As Integer = 0
            Console.WriteLine(Now & " : starting process")

            Do While Not pRow Is Nothing
                pFilter.WhereClause = "OBJECTID=" & pRow.Value(pRow.Fields.FindField("IAISOWNERPLY_FID"))
                Dim ownerPlyCursor As IFeatureCursor = iaisOwnerPlyFC.Search(pFilter, False)
                Dim pOwnerPly As IFeature = ownerPlyCursor.NextFeature
                If Not pOwnerPly Is Nothing Then
                    MapUtil.CopyFeature(pRow, pOwnerPly)
                    pOwnerPly.Store()
                End If

                Marshal.ReleaseComObject(ownerPlyCursor)

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pRow = pCursor.NextRow
            Loop

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counter & " ownerply were updated.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Process changes from ownerpolyupdate completed.")
    End Sub
#End Region

#Region "Process changes from updatepropertylineage"
    Sub ProcessUpdatepropertylineage()
        log.WriteLine(Now & ": Process changes from updatepropertylineage started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tablePropertyConfilictLineage As String

        Dim tableOtrOwnerPly As String
        Dim tableIAISOwnerPly As String

        Dim jtxJobParentVersion As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateOwnerPly")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PROPERTYCONFILICTLINEAGE" Then
                    tablePropertyConfilictLineage = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "OTROWNERPLYLAYER" Then
                    tableOtrOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "IAISOWNERPLYLAYER" Then
                    tableIAISOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
        Dim pSdeWS As IFeatureWorkspace = CType(pJTXDatabase.DataWorkspace(m_jtxJob.VersionName), IFeatureWorkspace)

        Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.OpenFromFile(Environment.CurrentDirectory & "\OwnerPlyExceptionTemp.gdb", 0)

        Dim pTempWorkspace As IFeatureWorkspace = pWorkspace

        'Dim tempExpTable As ITable = pTempWorkspace.OpenTable("OwnerPlyExceptionTemp")
        'Dim ownerPlyUpdateFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyUpdate")
        'Dim ownerPlyDeleteFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyDelete")
        'Dim ownerPlyAddFC As IFeatureClass = pTempWorkspace.OpenFeatureClass("ownerplyAdd")
        Dim tempUpdatepropertylineage As ITable = pTempWorkspace.OpenTable("updatepropertylineage")


        'Dim otrOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableOtrOwnerPly)
        'Dim iaisOwnerPlyFC As IFeatureClass = pSdeWS.OpenFeatureClass(tableIAISOwnerPly)


        Dim tableLineage As ITable = pSdeWS.OpenTable(tablePropertyConfilictLineage)

        Dim pWorkspaceEdit As IWorkspaceEdit = pSdeWS
        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Try


            Dim pCursor As ICursor = tempUpdatepropertylineage.Search(Nothing, False)
            Dim pRow As IRow = pCursor.NextRow
            Dim pFilter As IQueryFilter = New QueryFilter

            Dim counter As Integer = 0

            Dim counterAdd As Integer = 0
            Dim counterRemove As Integer = 0
            Console.WriteLine(Now & " : starting process")

            Do While Not pRow Is Nothing
                If MapUtil.GetRowValue(pRow, "PROCESS_TYPE") = "Delete" Then
                    Dim ssl As String = pRow.Value(pRow.Fields.FindField("INPUTFEATURESSL"))
                    pFilter.WhereClause = "INPUTFEATURESSL='" & ssl & "' OR CONFLICTFEATURESSL='" & ssl & "'"
                    Dim lineageCursor As ICursor = tableLineage.Search(pFilter, False)
                    Dim ptempRow As IRow = lineageCursor.NextRow
                    Do While Not ptempRow Is Nothing
                        ptempRow.Delete()
                        counterRemove = counterRemove + 1
                        ptempRow = lineageCursor.NextRow
                    Loop

                    Marshal.ReleaseComObject(lineageCursor)
                Else
                    Dim ptempRow As IRow = tableLineage.CreateRow
                    MapUtil.CopyRow(pRow, ptempRow)
                    ptempRow.Store()
                    counterAdd = counterAdd + 1
                End If

                counter = counter + 1
                If counter Mod 100 = 0 Then
                    Console.WriteLine(Now & " : processed " & counter & " records")
                End If

                pRow = pCursor.NextRow
            Loop

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counterAdd & " rows were added.")
            log.WriteLine(Now & ": " & counterRemove & " rows were added.")

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Process changes from updatepropertylineage completed.")
    End Sub
#End Region
End Module


