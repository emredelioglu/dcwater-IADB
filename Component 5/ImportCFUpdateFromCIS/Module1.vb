Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geodatabase

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms


Module Module1

    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()
    <STAThread()> _
    Sub Main()
        Dim log As System.IO.StreamWriter
        Dim exitCode As Integer = 0

        Dim folderName As String = Replace(FormatDateTime(Now, DateFormat.ShortDate), "/", "_")

        If Not File.Exists(Environment.CurrentDirectory & "\IACFException.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\IACFException.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\IACFException.log")
        End If

        log.WriteLine(Now & ": Load charge file exception started")

        Try

            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
            New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})

            Dim inFileName As String = Nothing
            Dim transdate As String = Nothing
            Dim sParams() As String = Environment.GetCommandLineArgs
            If sParams.Length > 2 Then
                inFileName = sParams(1)
                transdate = sParams(2)

                log.WriteLine("Parameter inputfile:" & inFileName)
                log.WriteLine("Parameter transdate:" & transdate)

                If Not IsDate(transdate) Then
                    Console.WriteLine("Usage for ImportCISPremsExp: ImportCISPremsExp [inputfile] [transaction date(MM/DD/YYYY)]")
                    Throw New Exception("Wrong arguments : " & sParams(2))
                    'Else
                    '    transdate = FormatDateTime(CDate(transdate), DateFormat.ShortDate)
                End If
            ElseIf sParams.Length > 1 Then
                inFileName = sParams(1)
                log.WriteLine("Parameter inputfile:" & inFileName)
                log.WriteLine("Parameter transdate: <none> ")
            Else
                Console.WriteLine("Usage for ImportCISPremsExp: ImportCISPremsExp [inputfile] [transaction date(MM/DD/YYYY)]")
                Throw New Exception("Wrong number of arguments")
            End If

            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode


            Dim jtxDB As String
            Dim jtxDBWorkspace As String
            Dim jtxJobParentVersion As String
            Dim cfExceptionType As String

            Dim tableCISOutputBillDet As String
            Dim tableCISUMSGF As String

            Dim tableEXCEPTIONSBILLDET As String


            m_xmld = New XmlDocument()
            m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

            m_nodelist = m_xmld.SelectNodes("IAISConfig/CFExportExceptionFromCIS")

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
                        log.WriteLine("JOB_PARENT_VERSION:" & jtxJobParentVersion)
                    ElseIf m_node.ChildNodes(i).Name = "EXCEPTION_TYPE" Then
                        cfExceptionType = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTBILLDET" Then
                        tableCISOutputBillDet = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSBILLDET" Then
                        tableEXCEPTIONSBILLDET = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_CISUMSGF" Then
                        tableCISUMSGF = m_node.ChildNodes(i).InnerText
                    End If

                Next i

            End If


            Dim fParser As New FileIO.TextFieldParser(inFileName)
            fParser.TextFieldType = FileIO.FieldType.Delimited
            fParser.SetDelimiters(",")
            fParser.HasFieldsEnclosedInQuotes = True
            fParser.TrimWhiteSpace = True

            Dim bExceptionFlag As Boolean = False

            Dim pFWS As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, jtxJobParentVersion)
            Dim pFilter As IQueryFilter = New QueryFilter
            Dim pTable As ITable = pFWS.OpenTable(tableCISOutputBillDet)

            Dim countRowMissing As Integer = 0
            Dim countRowException As Integer = 0

            While Not fParser.EndOfData And Not bExceptionFlag
                Dim fields() As String = fParser.ReadFields()
                Dim ciscode As String = fields((fields.Length - 1))
                If ciscode <> "" Then
                    bExceptionFlag = True
                    countRowException = countRowException + 1

                    If transdate Is Nothing Then
                        pFilter.WhereClause = "PEXUID=" & fields(1)
                    Else
                        pFilter.WhereClause = "PEXUID=" & fields(1) & " AND TRANSACTIONDT BETWEEN (To_Date(' " & transdate _
                                                & "', 'mm/dd/yyyy')-1) AND To_Date('" & transdate & "', 'mm/dd/yyyy') "
                    End If
                    If pTable.RowCount(pFilter) = 0 Then
                        countRowMissing = countRowMissing + 1
                    End If
                End If
            End While
            fParser.Close()

            Marshal.ReleaseComObject(pFWS)
            Marshal.ReleaseComObject(pTable)

            If Not bExceptionFlag Then
                log.WriteLine(Now & ": No exception was returned")
            ElseIf countRowMissing = countRowException Then
                log.WriteLine(Now & " : " & countRowException & " exceptions were returned but those records are not in the database")
            Else

                If countRowMissing > 0 Then
                    log.WriteLine(Now & " : " & countRowException & " exceptions were returned. " & countRowMissing & " records are not in the database")
                End If

                fParser = New FileIO.TextFieldParser(inFileName)
                fParser.TextFieldType = FileIO.FieldType.Delimited
                fParser.SetDelimiters(",")
                fParser.HasFieldsEnclosedInQuotes = True
                fParser.TrimWhiteSpace = True

                Dim exceptionCount As Integer = 0
                Dim removedRowCount As Integer = 0


                Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
                Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

                Dim pNewJob As IJTXJob2

                'Create job or get the job created

                If File.Exists(Environment.CurrentDirectory & "\jtxjob.txt") Then
                    Dim oRead As System.IO.StreamReader = File.OpenText(Environment.CurrentDirectory & "\jtxjob.txt")
                    Dim jobid As String = oRead.ReadLine
                    If Trim(jobid) <> "" Then
                        Dim pJTXJobMan As IJTXJobManager = pJTXDatabase.JobManager
                        pNewJob = pJTXJobMan.GetJob(jobid)

                        If pNewJob Is Nothing Then
                            log.WriteLine("Job no longer exists in JTX.  Delete your jtxjob.txt file.")
                            Return
                        Else
                            If Not pNewJob.VersionExists Then
                                mdlJTX.CreateJobVersion(pNewJob, log)
                            End If
                        End If

                    End If
                Else
                    pNewJob = mdlJTX.CreateJob(pJTXDatabase, cfExceptionType, jtxDBWorkspace, jtxJobParentVersion, log)
                    log.WriteLine(Now & ": Job " & pNewJob.ID & " is created")

                    Try
                        Dim oWriter As System.IO.StreamWriter = File.CreateText(Environment.CurrentDirectory & "\jtxjob.txt")
                        oWriter.WriteLine(pNewJob.ID)
                        oWriter.Close()
                    Catch ex As Exception
                        If Not pNewJob Is Nothing Then
                            'pJTXDatabase.JobManager.DeleteJob(pNewJob.ID)
                            mdlJTX.DeleteJob(pJTXDatabase, pNewJob, Environment.CurrentDirectory, log)
                            Console.WriteLine("JTX job was deleted because of system exception")
                            log.WriteLine(Now & ":JTX job was deleted because of system exception")
                        End If

                        log.WriteLine(Now & ": (Error:)" & ex.Message)
                        log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

                        Throw ex
                    End Try

                    mdlJTX.CreateJobVersion(pNewJob, log)

                End If

                Try
                    Dim pCursor As ICursor
                    Dim pRow As IRow

                    pFWS = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, pNewJob.VersionName)
                    pTable = pFWS.OpenTable(tableCISOutputBillDet)

                    Dim pCISUMSGFTable As ITable = pFWS.OpenTable(tableCISUMSGF)
                    Dim pEXCEPTIONSBILLDETTable As ITable = pFWS.OpenTable(tableEXCEPTIONSBILLDET)

                    Dim cisexceptdt As Date = Now

                    Dim pWorkspaceEdit As IWorkspaceEdit = pFWS
                    pWorkspaceEdit.StartEditing(False)
                    pWorkspaceEdit.StartEditOperation()

                    Try
                        While Not fParser.EndOfData
                            Dim fields() As String = fParser.ReadFields()
                            If transdate Is Nothing Then
                                pFilter.WhereClause = "PEXUID=" & fields(1)
                            Else
                                pFilter.WhereClause = "PEXUID=" & fields(1) & " AND TRANSACTIONDT BETWEEN (To_Date(' " & transdate _
                                                        & "', 'mm/dd/yyyy')-1) AND To_Date('" & transdate & "', 'mm/dd/yyyy') "
                            End If
                            pCursor = pTable.Search(pFilter, False)
                            pRow = pCursor.NextRow
                            If Not pRow Is Nothing Then

                                Dim ciscode As String = fields((fields.Length - 1))

                                Do While Not pRow Is Nothing
                                    pRow.Value(pRow.Fields.FindField("EXCEPTIONDT")) = Now
                                    pRow.Value(pRow.Fields.FindField("EXCEPTIONMSGID")) = ciscode
                                    pRow.Store()

                                    pRow = pCursor.NextRow
                                Loop


                                Dim cismessage As String = ""

                                pFilter.WhereClause = "MMSGID='" & ciscode & "'"
                                Dim pCISUMSGFCursor As ICursor = pCISUMSGFTable.Search(pFilter, False)
                                Dim pMsgRow As IRow = pCISUMSGFCursor.NextRow
                                If Not pMsgRow Is Nothing Then
                                    cismessage = pMsgRow.Value(pMsgRow.Fields.FindField("MMSG"))
                                End If

                                Marshal.ReleaseComObject(pCISUMSGFCursor)

                                Dim pExpRow As IRow = pEXCEPTIONSBILLDETTable.CreateRow()
                                pExpRow.Value(pExpRow.Fields.FindField("PEXUID")) = fields(1)

                                pExpRow.Value(pExpRow.Fields.FindField("CISEXCEPTDT")) = cisexceptdt
                                pExpRow.Value(pExpRow.Fields.FindField("CISMESSAGE")) = cismessage
                                pExpRow.Value(pExpRow.Fields.FindField("CISCODE")) = ciscode

                                pExpRow.Store()

                                exceptionCount = exceptionCount + 1
                            Else
                                log.WriteLine("Can't find matching row " & fields(1) & " in " & tableCISOutputBillDet)

                            End If

                            Marshal.ReleaseComObject(pCursor)

                        End While

                        fParser.Close()

                        removedRowCount = 0

                        If transdate Is Nothing Then
                            pFilter.WhereClause = "EXCEPTIONDT IS NULL"
                        Else
                            pFilter.WhereClause = "EXCEPTIONDT IS NULL AND TRANSACTIONDT BETWEEN (To_Date(' " & transdate _
                                                & "', 'mm/dd/yyyy')-1) AND To_Date('" & transdate & "', 'mm/dd/yyyy') "
                        End If
                        pCursor = pTable.Search(pFilter, False)
                        pRow = pCursor.NextRow
                        Do While Not pRow Is Nothing
                            pRow.Delete()
                            removedRowCount = removedRowCount + 1
                            pRow = pCursor.NextRow
                        Loop

                        pWorkspaceEdit.StopEditOperation()
                        pWorkspaceEdit.StopEditing(True)

                    Catch ex As Exception
                        pWorkspaceEdit.StopEditOperation()
                        pWorkspaceEdit.StopEditing(False)

                        log.WriteLine(Now & ": (Error:)" & ex.Message)
                        log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

                        Throw ex
                    Finally
                        Marshal.ReleaseComObject(pFWS)
                        Marshal.ReleaseComObject(pCursor)
                        Marshal.ReleaseComObject(pCISUMSGFTable)
                        Marshal.ReleaseComObject(pTable)
                    End Try
                Catch ex As Exception
                    'If Not pNewJob Is Nothing Then
                    '    pNewJob.DeleteVersion()
                    '    pJTXDatabase.JobManager.DeleteJob(pNewJob.ID, True)
                    '    log.WriteLine(Now & ": Job " & pNewJob.ID & " is removed")
                    'End If

                    log.WriteLine(Now & ": (Error:)" & ex.Message)
                    log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

                    Throw ex
                Finally
                    fParser.Close()
                End Try

                If exceptionCount = 0 Then
                    log.WriteLine(Now & ": No exception was returned")
                    If Not pNewJob Is Nothing Then
                        pNewJob.DeleteVersion()
                        mdlJTX.DeleteJob(pJTXDatabase, pNewJob, Environment.CurrentDirectory, log)
                        'pJTXDatabase.JobManager.DeleteJob(pNewJob.ID, True)
                        log.WriteLine(Now & ": Job " & pNewJob.ID & " is removed")
                    End If
                Else
                    log.WriteLine(Now & ": " & exceptionCount & " exception(s) were returned")
                End If

                log.WriteLine(Now & ": " & removedRowCount & " rows(s) were removed")
            End If


            log.WriteLine(Now & ": Load charge file exception completed ")
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            exitCode = 1
        End Try

        log.Close()
        m_AOLicenseInitializer.ShutdownApplication()
        Environment.Exit(exitCode)

    End Sub


    Public Function GetRowValue(ByVal pRow As IRow, ByVal fIndex As Integer) As String
        If IsDBNull(pRow.Value(fIndex)) Then
            Return ""
        Else
            Dim pField As IField
            pField = pRow.Fields.Field(fIndex)
            If pField.Type = esriFieldType.esriFieldTypeDate Then
                Return FormatDateTime(pRow.Value(fIndex), DateFormat.ShortDate)
            Else
                Return pRow.Value(fIndex)
            End If
        End If
    End Function


    Public Function SetRowValue(ByVal pRow As IRow, ByVal fIndex As Integer, ByVal fValue As String) As Boolean

        If fIndex <= 0 Then
            Return False
        End If

        Dim pField As IField
        pField = pRow.Fields.Field(fIndex)

        If Trim(fValue) = "" Then
            If Not IsDBNull(pRow.Value(fIndex)) _
                AndAlso CStr(pRow.Value(fIndex)) <> "" _
                AndAlso pField.IsNullable Then
                pRow.Value(fIndex) = System.DBNull.Value
            End If
        ElseIf pField.Type = esriFieldType.esriFieldTypeDate Then
            pRow.Value(fIndex) = CDate(fValue)
        ElseIf pField.Type = esriFieldType.esriFieldTypeDouble Then
            pRow.Value(fIndex) = CDbl(fValue)
        ElseIf pField.Type = esriFieldType.esriFieldTypeInteger Then
            pRow.Value(fIndex) = CLng(fValue)
        ElseIf pField.Type = esriFieldType.esriFieldTypeSingle Then
            pRow.Value(fIndex) = CSng(fValue)
        ElseIf pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
            pRow.Value(fIndex) = CInt(fValue)
        ElseIf pField.Type = esriFieldType.esriFieldTypeString Then
            pRow.Value(fIndex) = fValue
        Else
            pRow.Value(fIndex) = fValue
        End If

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

End Module
