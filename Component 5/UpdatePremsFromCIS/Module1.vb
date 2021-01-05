Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geodatabase
Imports System.Xml
Imports System.IO
Imports System.Data.OracleClient
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.Windows.Forms

Module Module1

    Private log As System.IO.StreamWriter
    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()

    Private m_jtxJob As IJTXJob = Nothing
    Private m_exitCode As Integer

    <STAThread()> _
    Sub Main()

        If Not File.Exists(Environment.CurrentDirectory & "\UpdatePremiseFromCIS.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\UpdatePremiseFromCIS.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\UpdatePremiseFromCIS.log")
        End If

        Try
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
            New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})

            If File.Exists(Environment.CurrentDirectory & "\jtxjob.txt") Then
                Dim oRead As System.IO.StreamReader = File.OpenText(Environment.CurrentDirectory & "\jtxjob.txt")
                Dim jobid As String = oRead.ReadLine
                m_jtxJob = GetJTXJob(jobid)

                If m_jtxJob Is Nothing Then
                    log.WriteLine("Job no longer exists in JTX.  Delete your jtxjob.txt file.")
                    Return
                End If

                If Not m_jtxJob.VersionExists Then
                    mdlJTX.CreateJobVersion(m_jtxJob, log)
                End If
            End If

            Dim sParams() As String = Environment.GetCommandLineArgs
            If sParams.Length > 1 Then

                log.WriteLine("Parameter routineName:" & sParams(1))

                If sParams(1) = "CreateTempGeoDB" Then
                    CreateTempGeoDB()
                ElseIf sParams(1) = "CheckPremiseException" Then
                    GetPremsExceptions()
                ElseIf sParams(1) = "CreateJTXJob" Then
                    CreatePremiseJTXJob()
                Else
                    If m_jtxJob Is Nothing Then
                        Throw New Exception("JTX job has not been created")
                    End If
                    If sParams(1) = "UpdateCISAccountInfo" Then
                        UpdateCISAcccountInfo()
                    ElseIf sParams(1) = "CreatePremiseExceptions" Then
                        CreatePremiseExceptions()
                    Else
                        Console.WriteLine("Usage for UpdatePremsFromCIS: UpdatePremsFromCIS CreateTempGeoDB|CheckPremiseException|CreateJTXJob|UpdateCISAccountInfo|CreatePremiseExceptions")
                        Throw New Exception("Wrong arguments " & sParams.ToString)
                    End If
                End If
            Else
                Console.WriteLine("Usage for UpdatePremsFromCIS: UpdatePremsFromCIS CreateTempGeoDB|CheckPremiseException|CreateJTXJob|UpdateCISAccountInfo|CreatePremiseExceptions")
                Throw New Exception("No arguments")
            End If

            m_exitCode = 0
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())
            m_exitCode = 1
        End Try

        log.Close()
        m_AOLicenseInitializer.ShutdownApplication()
        Environment.Exit(m_exitCode)
    End Sub

    Sub CreateTempGeoDB()
        Try
            If File.Exists(Environment.CurrentDirectory & "\PremiseExceptionTemp.mdb") Then
                log.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": Database already exists")
                Return
            End If

            Console.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": CreateTempGeoDB ")

            Dim tempDbName As String = "PremiseExceptionTemp.mdb"
            File.Copy(Application.StartupPath & "\" & tempDbName, Environment.CurrentDirectory & "\" & tempDbName)
            log.WriteLine(Now & ": Temporary database created.")

        Catch ex As Exception
            Console.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": CreateTempGeoDB :" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try
    End Sub

    Sub GetPremsExceptions()
        log.WriteLine(Now & ": Checked premise exceptions started.")
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode


        Dim oracleDatasource As String = ""
        Dim username As String = ""
        Dim password As String = ""
        Dim integratedSecurity As String = ""

        Dim jtxJobParentVersion As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/PremiseUpdateFromCIS")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
               If m_node.ChildNodes(i).Name = "ORACON" Then
                    Dim attCol As XmlAttributeCollection = m_node.ChildNodes(i).Attributes

                    oracleDatasource = attCol.GetNamedItem("DataSource").InnerText
                    username = attCol.GetNamedItem("username").InnerText
                    password = attCol.GetNamedItem("password").InnerText
                    integratedSecurity = attCol.GetNamedItem("integratedsecurity").InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If


        Dim oOracleConn As OracleConnection = New OracleConnection
        If integratedSecurity <> "" Then
            oOracleConn.ConnectionString = "Data Source=" & oracleDatasource & ";Integrated Security=" & integratedSecurity
        Else
            oOracleConn.ConnectionString = "Data Source=" & oracleDatasource & ";User ID=" & username & ";Password=" & password
        End If
        oOracleConn.Open()

        Dim Ds As New DataSet
        Try
            Dim oraCmd As New OracleCommand
            oraCmd.Connection = oOracleConn
            oraCmd.CommandText = "iais_cis_pk.GetCISExportExceptions"
            oraCmd.CommandType = CommandType.StoredProcedure
            oraCmd.Parameters.Add(New OracleParameter("versionName", OracleClient.OracleType.VarChar)).Value = jtxJobParentVersion
            oraCmd.Parameters.Add(New OracleParameter("io_cursor", OracleType.Cursor)).Direction = ParameterDirection.Output

            Dim oraDA As New OracleDataAdapter(oraCmd)

            oraDA.Fill(Ds)
            oraDA.Dispose()
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        Finally
            oOracleConn.Close()
            oOracleConn = Nothing
        End Try

        Dim tempDbPath As String = Environment.CurrentDirectory & "\PremiseExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("PremiseExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pWorkspace
        pWSEdit.StartEditing(False)

        Try
            Dim count As Integer = 0
            Dim table As DataTable = Ds.Tables(0)
            For Each row As DataRow In table.Rows
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "ANY CIS FIELD UPDATE NOT PROCESSED"
                pRow.Value(pRow.Fields.FindField("UniqueID")) = row.Item(0).ToString
                pRow.Store()
                count = count + 1
            Next
            pWSEdit.StopEditing(True)
            log.WriteLine(Now & ": " & count & " records created in the temp database.")
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try

        log.WriteLine(Now & ": Checked premise exceptions completed.")

    End Sub

    Sub UpdateCISAcccountInfo()
        log.WriteLine(Now & ": Update CIS account infromation started.")
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode


        Dim oracleDatasource As String = ""
        Dim username As String = ""
        Dim password As String = ""
        Dim integratedSecurity As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/PremiseUpdateFromCIS")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "ORACON" Then
                    Dim attCol As XmlAttributeCollection = m_node.ChildNodes(i).Attributes

                    oracleDatasource = attCol.GetNamedItem("DataSource").InnerText
                    username = attCol.GetNamedItem("username").InnerText
                    password = attCol.GetNamedItem("password").InnerText
                    integratedSecurity = attCol.GetNamedItem("integratedsecurity").InnerText

                End If
            Next i
        End If


        Dim oOracleConn As OracleConnection = New OracleConnection
        If integratedSecurity <> "" Then
            oOracleConn.ConnectionString = "Data Source=" & oracleDatasource & ";Integrated Security=" & integratedSecurity
        Else
            oOracleConn.ConnectionString = "Data Source=" & oracleDatasource & ";User ID=" & username & ";Password=" & password
        End If
        oOracleConn.Open()
        Try
            Dim Ds As New DataSet
            Dim oraCmd As New OracleCommand
            oraCmd.Connection = oOracleConn
            oraCmd.CommandText = "iais_cis_pk.UpdateCISAccountInfo"
            oraCmd.CommandType = CommandType.StoredProcedure
            oraCmd.Parameters.Add(New OracleParameter("versionName", OracleClient.OracleType.VarChar)).Value = m_jtxJob.VersionName

            oraCmd.ExecuteNonQuery()
            oOracleConn.Close()

        Catch ex As Exception
            oOracleConn.Close()
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try

        log.WriteLine(Now & ": Update CIS account infromation completed.")

    End Sub


    Function GetJTXJob(ByVal jobid As Integer) As IJTXJob

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
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try

    End Function

    Sub CreatePremiseJTXJob()
        log.WriteLine(Now & ": Create premise exception job started")

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String
        Dim jtxJobParentVersion As String
        Dim exceptionType As String

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/PremiseUpdateFromCIS")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "EXCEPTION_TYPE" Then
                    exceptionType = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

        m_jtxJob = mdlJTX.CreateJob(pJTXDatabase, exceptionType, jtxDBWorkspace, jtxJobParentVersion, log)

        Dim oWriter As System.IO.StreamWriter = File.CreateText(Environment.CurrentDirectory & "\jtxjob.txt")
        oWriter.WriteLine(m_jtxJob.ID)
        oWriter.Close()

        mdlJTX.CreateJobVersion(m_jtxJob, log)

        log.WriteLine(Now & ": Create premise exception job completed. New Job ID " & m_jtxJob.ID)
    End Sub

    Sub CreatePremiseExceptions()
        log.WriteLine(Now & ": Create premise exception job started")


        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/PremiseUpdateFromCIS")

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

        Dim tempDbPath As String = Environment.CurrentDirectory & "\PremiseExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("PremiseExceptionTemp")

        Dim pCursor As ICursor = tempExpTable.Search(Nothing, False)
        Dim pRow As IRow = pCursor.NextRow
        If Not pRow Is Nothing Then
            Dim pFWS As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, m_jtxJob.VersionName)

            Dim pExpRow As IRow
            Dim pExpTable As ITable = pFWS.OpenTable("EXCEPTIONSPREMSINTERPT")
            Dim pWorkspaceEdit As IWorkspaceEdit = pFWS

            pWorkspaceEdit.StartEditing(False)
            pWorkspaceEdit.StartEditOperation()

            Try
                Dim count As Integer = 0
                Do While Not pRow Is Nothing
                    pExpRow = pExpTable.CreateRow
                    SetRowValue(pExpRow, pExpRow.Fields.FindField("EXCEPTDT"), Now)
                    SetRowValue(pExpRow, pExpRow.Fields.FindField("EXCEPTTYP"), 4, True)
                    SetRowValue(pExpRow, pExpRow.Fields.FindField("PEXUID"), GetRowValue(pRow, "UniqueID"))

                    pExpRow.Store()
                    count = count + 1

                    pRow = pCursor.NextRow
                Loop
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)

                log.WriteLine(Now & ": " & count & " exceptions were created.")
            Catch ex As Exception
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(False)
                log.WriteLine(Now & ": (Error:)" & ex.Message)
                log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

                Throw ex
            End Try

        Else
            If Not m_jtxJob Is Nothing Then
                m_jtxJob.DeleteVersion()
                mdlJTX.DeleteJob(pJTXDatabase, m_jtxJob, Environment.CurrentDirectory, log)
                'pJTXDatabase.JobManager.DeleteJob(m_jtxJob.ID, True)
                log.WriteLine(Now & ": No exception was return. JTX job is removed")
            End If
        End If
        log.WriteLine(Now & ": Create premise exception job completed")

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

    Private Function GetRowValue(ByVal pRow As IRow, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Dim fIndex As Integer
        fIndex = pRow.Fields.FindField(col)
        If fIndex < 0 Then
            'MsgBox("Can't find field [" & UCase(col) & "]")
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


    Private Function SetRowValue(ByVal pRow As IRow, ByVal fIndex As Integer, ByVal fValue As String, Optional ByVal valueOnly As Boolean = False) As Boolean
        'Try

        If fIndex <= 0 Then
            Return False
        End If

        Dim pField As IField

        Dim pDomain As IDomain
        Dim pCodeDomain As ICodedValueDomain

        pField = pRow.Fields.Field(fIndex)

        pDomain = pField.Domain

        If Not pDomain Is Nothing And Trim(fValue) <> "" Then
            'GetDomainCode
            If Not valueOnly And pDomain.Type = esriDomainType.esriDTCodedValue Then
                pCodeDomain = pDomain
                If IsDBNull(GetDomainCode(pCodeDomain, fValue)) Then
                    If Not pField.IsNullable Then
                        If pField.Type = esriFieldType.esriFieldTypeString Then
                            pRow.Value(fIndex) = ""
                        Else
                            pRow.Value(fIndex) = 0
                        End If

                    End If
                Else
                    pRow.Value(fIndex) = GetDomainCode(pCodeDomain, fValue)
                End If
            Else
                pRow.Value(fIndex) = fValue
            End If
        Else
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
        End If

    End Function

    Private Function GetDomainCode(ByVal pCodeDomain As ICodedValueDomain, ByVal value As String) As Object

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
End Module
