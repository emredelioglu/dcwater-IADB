Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.esriSystem
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms

Module Module1

    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()
    Private m_exitCode As Integer

    <STAThread()> _
    Sub Main()

        Dim log As System.IO.StreamWriter

        If Not File.Exists(Environment.CurrentDirectory & "\AdministrativeRoutine.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\AdministrativeRoutine.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\AdministrativeRoutine.log")
        End If

        log.WriteLine(Now & "Administrative routine started")

        Try
            'ESRI License Initializer generated code.
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
            New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})
            'ESRI License Initializer generated code.
            'Do not make any call to ArcObjects after ShutDownApplication()

            Dim sParams() As String = Environment.GetCommandLineArgs
            Dim routineName As String = ""
            Dim transDate As String = ""

            If sParams.Length > 1 Then
                routineName = sParams(1)
                log.WriteLine("Parameter routineName:" & routineName)
                If sParams.Length > 2 Then
                    transDate = sParams(2)
                    log.WriteLine("Parameter transDate:" & transDate)
                End If
            Else
                Console.WriteLine("Usage for AdministrativeRoutine: AdministrativeRoutine [ALL|CISOUTPUTPREMEXP|CISOUTPUTBILLDET|EXCEPTIONSADDRESSPT|EXCEPTIONSOWNERPLY|EXCEPTIONSPREMSINTERPT|EXCEPTIONSIAPLY|EXCEPTIONSPREMSEXPOUT|EXCEPTIONSBILLDET] [TRANSACTION_DATE]")
                Throw New Exception("Wrong number of arguments")
            End If

            If (routineName.Equals("ALL") Or routineName.Equals("CISOUTPUTPREMEXP")) _
                And transDate = "" Then
                Console.WriteLine("[TRANSACTION_DATE] is required for CISOUTPUTPREMEXP")
                Throw New Exception("[TRANSACTION_DATE] is required for CISOUTPUTPREMEXP")
            End If

            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode

            Dim jtxDB As String
            Dim jtxDBWorkspace As String
            Dim jtxJobParentVersion As String

            Dim tableCisoutputpremexp As String = ""
            Dim tableCisoutputbilldet As String = ""
            Dim tableExceptionsaddresspt As String = ""
            Dim tableExceptionsownerply As String = ""
            Dim tableExceptionspremsinterpt As String = ""
            Dim tableExceptionsiaply As String = ""
            Dim tableEXCEPTIONSPREMSEXPOUT As String = ""
            Dim tableExceptionsbilldet As String = ""

            m_xmld = New XmlDocument()
            m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

            m_nodelist = m_xmld.SelectNodes("IAISConfig/AdministrativeRoutine")

            If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
                m_node = m_nodelist.Item(0)
                For i = 0 To m_node.ChildNodes.Count - 1
                    If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                        jtxDB = m_node.ChildNodes(i).InnerText
                        log.WriteLine("JTX_DB_NAME:" & jtxDB)
                    ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                        jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                        log.WriteLine("JTX_WORKSPACE:" & jtxDBWorkspace)
                    ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                        jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                        log.WriteLine("VERSION:" & jtxJobParentVersion)
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTPREMEXP" Then
                        tableCisoutputpremexp = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTBILLDET" Then
                        tableCisoutputbilldet = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSADDRESSPT" Then
                        tableExceptionsaddresspt = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSOWNERPLY" Then
                        tableExceptionsownerply = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSPREMSEXPOUT" Then
                        tableEXCEPTIONSPREMSEXPOUT = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSPREMSINTERPT" Then
                        tableExceptionspremsinterpt = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSIAPLY" Then
                        tableExceptionsiaply = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "EXCEPTIONSBILLDET" Then
                        tableExceptionsbilldet = m_node.ChildNodes(i).InnerText
                    End If
                Next i
            End If


            Dim pFWS As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, jtxJobParentVersion)
            Dim pWorkspaceEdit As IWorkspaceEdit = pFWS

            pWorkspaceEdit.StartEditing(False)
            pWorkspaceEdit.StartEditOperation()

            Try
                'Open Exception tables and remove all corrected records
                Dim pTable As ITable
                Dim pCursor As ICursor
                Dim pRow As IRow
                Dim pFilter As IQueryFilter = New QueryFilter

                Dim rowCount As Integer = 0

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("CISOUTPUTPREMEXP") Then
                    pTable = pFWS.OpenTable(tableCisoutputpremexp)
                    pFilter.WhereClause = "EXCEPTIONMSGID IS NULL AND TRANSACTIONDT < to_date('" & transDate & "', 'mm/dd/yyyy') "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop

                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableCisoutputpremexp)

                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("CISOUTPUTBILLDET") Then
                    pTable = pFWS.OpenTable(tableCisoutputbilldet)
                    pFilter.WhereClause = "EXCEPTIONMSGID IS NULL AND TRANSACTIONDT < to_date('" & transDate & "', 'mm/dd/yyyy') "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop

                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableCisoutputbilldet)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSADDRESSPT") Then
                    pTable = pFWS.OpenTable(tableExceptionsaddresspt)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop

                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionsaddresspt)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSOWNERPLY") Then
                    pTable = pFWS.OpenTable(tableExceptionsownerply)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop

                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionsownerply)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSPREMSINTERPT") Then
                    pTable = pFWS.OpenTable(tableExceptionspremsinterpt)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop

                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionspremsinterpt)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSIAPLY") Then
                    pTable = pFWS.OpenTable(tableExceptionsiaply)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop
                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionsiaply)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSPREMSEXPOUT") Then
                    pTable = pFWS.OpenTable(tableExceptionsiaply)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop
                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionsiaply)
                End If

                rowCount = 0
                If routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSBILLDET") Then
                    pTable = pFWS.OpenTable(tableExceptionsbilldet)
                    pFilter.WhereClause = "NOT FIXDT IS NULL "
                    pCursor = pTable.Search(pFilter, False)
                    pRow = pCursor.NextRow
                    Do While Not pRow Is Nothing
                        pRow.Delete()
                        rowCount = rowCount + 1
                        pRow = pCursor.NextRow
                    Loop
                    log.WriteLine(Now & " Removed " & rowCount & " record(s) from table " & tableExceptionsbilldet)
                End If

                Dim pVersionEdit As IVersionEdit
                pVersionEdit = pFWS
                Dim pEnumConflictClass As IEnumConflictClass
                Dim pConflictClass As IConflictClass
                Dim pDataset As IDataset

                pEnumConflictClass = pVersionEdit.ConflictClasses
                pEnumConflictClass.Reset()

                pConflictClass = pEnumConflictClass.Next
                Dim conflictFlag As Boolean = False
                Do Until pConflictClass Is Nothing And Not conflictFlag
                    pDataset = pConflictClass
                    If pConflictClass.HasConflicts Then
                        If (routineName.Equals("ALL") Or routineName.Equals("CISOUTPUTPREMEXP")) _
                            And UCase(pDataset.BrowseName) = UCase(tableCisoutputpremexp) Then
                            log.WriteLine(Now & "Table " & tableCisoutputpremexp & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("CISOUTPUTBILLDET")) _
                            And UCase(pDataset.BrowseName) = UCase(tableCisoutputbilldet) Then
                            log.WriteLine(Now & "Table " & tableCisoutputbilldet & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSADDRESSPT")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionsaddresspt) Then
                            log.WriteLine(Now & "Table " & tableExceptionsaddresspt & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSOWNERPLY")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionsownerply) Then
                            log.WriteLine(Now & "Table " & tableExceptionsownerply & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSPREMSINTERPT")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionspremsinterpt) Then
                            log.WriteLine(Now & "Table " & tableExceptionspremsinterpt & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSPREMSINTERPT")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionspremsinterpt) Then
                            log.WriteLine(Now & "Table " & tableExceptionspremsinterpt & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSPREMSEXPOUT")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionsiaply) Then
                            log.WriteLine(Now & "Table " & tableEXCEPTIONSPREMSEXPOUT & " has conflicts.")
                            conflictFlag = True
                        ElseIf (routineName.Equals("ALL") Or routineName.Equals("EXCEPTIONSBILLDET")) _
                            And UCase(pDataset.BrowseName) = UCase(tableExceptionsiaply) Then
                            log.WriteLine(Now & "Table " & tableEXCEPTIONSPREMSEXPOUT & " has conflicts.")
                            conflictFlag = True
                        End If
                    End If
                    pConflictClass = pEnumConflictClass.Next
                Loop

                If conflictFlag Then
                    Throw New Exception("Conflicts found.")
                End If

                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)

                log.WriteLine(Now & "Administrative routine completed")

            Catch ex As Exception
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(False)

                log.WriteLine(Now & ": (Error:)" & ex.Message)
                log.WriteLine(Now & ": (Error:)" & ex.StackTrace())


                Throw ex
            End Try

            m_exitCode = 0
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            m_exitCode = 1
        Finally
            m_AOLicenseInitializer.ShutdownApplication()
        End Try

        log.Close()

        Environment.Exit(m_exitCode)

    End Sub

End Module
