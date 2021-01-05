Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geometry
Imports System.Data.OleDb

Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.Xml
Imports System.Windows.Forms

Imports System.Data.OracleClient
Imports System.IO
Imports System.Runtime.InteropServices


Module Module1

    Private m_ToWS As IWorkspace
    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()
    Private log As System.IO.StreamWriter

    Private m_jtxJob As IJTXJob
    Private m_exitCode As Integer

#Region "Main entrance"

    <STAThread()> _
    Sub Main()
        m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
        New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})

        'Dim folderName As String = Replace(FormatDateTime(Now, DateFormat.ShortDate), "/", "_")
        'If Not Directory.Exists(Environment.CurrentDirectory & "\" & folderName) Then
        '    Directory.CreateDirectory(Environment.CurrentDirectory & "\" & folderName)
        'End If

        'Create or open log file
        If Not File.Exists(Environment.CurrentDirectory & "\ImportMAR.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\ImportMAR.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\ImportMAR.log")
        End If

        Try
            Dim sParams() As String = Environment.GetCommandLineArgs
            If sParams.Length > 1 Then
                log.WriteLine("Parameter routineName:" & sParams(1))

                If sParams(1) = "CreateTempGeoDB" Then
                    CreateTempGeoDB()
                ElseIf sParams(1) = "CheckMarLocation" Then
                    CheckMarLocation()
                ElseIf sParams(1) = "CheckMARExp" Then
                    CheckMARExp()
                ElseIf sParams(1) = "CheckMARRetired" Then
                    CheckMARRetired()
                ElseIf sParams(1) = "CheckAddrUpdateNoExp" Then
                    CheckAddrUpdateNoExp()
                ElseIf sParams(1) = "CheckAddrUpdateExp" Then
                    CheckAddrUpdateExp()
                ElseIf sParams(1) = "MARAddrMatch" Then
                    MARAddrMatch()
                ElseIf sParams(1) = "MARPotentialMatch" Then
                    MARPotentialMatch()
                ElseIf sParams(1) = "CreateMARJTXJob" Then

                    CreateMARJTXJob()

                ElseIf sParams(1) = "UpdatePremise" Then
                    UpdatePremise()
                ElseIf sParams(1) = "CreateMarUpdateException" Then
                    CreateMarUpdateException()
                Else
                    Console.WriteLine("Usage for UpdateMAR: UpdateMAR CreateTempGeoDB|CheckMarLocation|CheckMARExp|CheckMARRetired|CheckAddrUpdateNoExp|CheckAddrUpdateExp|MARAddrMatch|MARPotentialMatch|CreateMARJTXJob|UpdatePremise|CreateMarUpdateException")
                    Throw New Exception("Wrong arguments : " & sParams(1))
                End If
            Else
                Console.WriteLine("Usage for UpdateMAR: UpdateMAR CreateTempGeoDB|CheckMarLocation|CheckMARExp|CheckMARRetired|CheckAddrUpdateNoExp|CheckAddrUpdateExp|MARAddrMatch|MARPotentialMatch|CreateMARJTXJob|UpdatePremise|CreateMarUpdateException")
                Throw New Exception("Wrong number of arguments")
            End If
            m_exitCode = 0

        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(ex.StackTrace)
            Console.WriteLine("Error:" & ex.Message)
            m_exitCode = 1
        End Try

        log.Close()
        m_AOLicenseInitializer.ShutdownApplication()
        Environment.Exit(m_exitCode)
    End Sub
#End Region

#Region "Create Temp DB"
    Sub CreateTempGeoDB()
        Try
            If File.Exists(Environment.CurrentDirectory & "\MarExceptionTemp.mdb") Then
                log.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": Database already exists")
                Return
            End If

            Dim tempDbName As String = "MarExceptionTemp.mdb"
            File.Copy(Application.StartupPath & "\" & tempDbName, Environment.CurrentDirectory & "\" & tempDbName)
            log.WriteLine(Now & ": Temporary database created.")

        Catch ex As Exception
            Console.WriteLine(FormatDateTime(Now, DateFormat.LongTime) & ": CreateTempGeoDB :" & ex.Message)
            Throw ex
        End Try

    End Sub

#End Region

#Region "CheckMARLocation - Check MAR point that are removed or with location changed"


    Sub CheckMarLocation()
        log.WriteLine(Now & ": Checked address point location started.")
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateMAR")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "TOLERANCE" Then
                    tolerance = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If


        If tolerance = "" Or Not IsNumeric(tolerance) Then
            Throw New Exception("Invalid tolerance value")
        End If


        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef

            pQueryDef.Tables = "AddressPtOld,AddressPtRev,PremsInterPt"
            pQueryDef.SubFields = "DISTINCT AddressPtOld.ADDRESS_ID"
            pQueryDef.WhereClause = "AddressPtOld.ADDRESS_ID=AddressPtRev.ADDRESS_ID AND " & _
                                    "AddressPtOld.ADDRESS_ID=PremsInterPt.ADDRESS_ID  AND " & _
                                    "( ( AddressPtOld.XCOORD - AddressPtRev.XCOORD ) * ( AddressPtOld.XCOORD - AddressPtRev.XCOORD ) " & _
                                    "+ ( AddressPtOld.YCOORD - AddressPtRev.YCOORD ) * ( AddressPtOld.YCOORD - AddressPtRev.YCOORD ) > " & _
                                    tolerance * tolerance & ")"

            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate

            Dim pOldMar As IRow
            pOldMar = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Changed MAR Location"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "C"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(0)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pQueryDef.Tables = "AddressPtOld,PremsInterPt"
            pQueryDef.SubFields = "DISTINCT AddressPtOld.ADDRESS_ID"
            pQueryDef.WhereClause = "AddressPtOld.ADDRESS_ID = PremsInterPt.ADDRESS_ID And " & _
                                    "NOT EXISTS ( SELECT * FROM AddressPtRev WHERE AddressPtRev.ADDRESS_ID = AddressPtOld.ADDRESS_ID )"

            pCursor = pQueryDef.Evaluate


            pOldMar = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Not in MAR"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "O"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(0)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop


            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " exceptions.")
        Else
            log.WriteLine(Now & ": No exceptions.")
        End If

        log.WriteLine(Now & ": Checked address point location completed.")
    End Sub

#End Region

#Region "Compare address components of premise point with old Address point"
    Sub CheckMARExp()

        log.WriteLine(Now & ": Checked MAR Exception started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef


            pQueryDef.Tables = "AddressPtOld,PremsInterPt,STREET_TYPE"
            pQueryDef.SubFields = "AddressPtOld.ADDRESS_ID"
            pQueryDef.WhereClause = "PremsInterPt.address_id =AddressPtOld.address_id " & _
                                    "AND AddressPtOld.STREET_TYPE = STREET_TYPE.STREET_TYPE " & _
                                    "AND NOT PremsInterPt.PEXPSTS IN ('PG','ST') " & _
                                    "AND PremsInterPt.USE_PARSED_ADDRESS = 'Y' " & _
                                    "AND " & _
                                    "( NOT Iif(IsNull(PremsInterPt.ADDRNUM),""0"",PremsInterPt.ADDRNUM) = Iif(IsNull(AddressPtOld.ADDRNUM),""0"",CStr(AddressPtOld.ADDRNUM)) " & _
                                    "OR NOT Iif(IsNull(PremsInterPt.STNAME),""0"",PremsInterPt.STNAME) = Iif(IsNull(AddressPtOld.STNAME),""0"",AddressPtOld.STNAME) " & _
                                    "OR NOT Iif(IsNull(PremsInterPt.STREET_TYPE),""0"",PremsInterPt.STREET_TYPE) = STREET_TYPE.USPS_ST_TYPE " & _
                                    "OR NOT Iif(IsNull(PremsInterPt.QUADRANT),""0"",PremsInterPt.QUADRANT) = Iif(IsNull(AddressPtOld.QUADRANT),""0"",AddressPtOld.QUADRANT) " & _
                                    "OR ( " & _
                                    "    NOT Iif(IsNull(PremsInterPt.ADDRNUMSUF),""0"",PremsInterPt.ADDRNUMSUF) = Iif(IsNull(AddressPtOld.ADDRNUMSUFFIX),""0"",AddressPtOld.ADDRNUMSUFFIX) " & _
                                    "    AND NOT(PremsInterPt.ADDRNUMSUF IS NOT NULL AND AddressPtOld.ADDRNUMSUFFIX IS NULL) " & _
                                    "    ) )"


            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate


            Dim pOldMar As IRow = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Inconsistent MAR refernce"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "O"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(0)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " exceptions.")
        Else
            log.WriteLine(Now & ": No exceptions.")
        End If

        log.WriteLine(Now & ": Checked MAR Exception completed.")

    End Sub
#End Region

#Region "Check address point labeled as retired"
    Sub CheckMARRetired()
        log.WriteLine(Now & ": Check MAR Retired started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pQueryDef As IQueryDef
        pQueryDef = pFeatureWorkspace.CreateQueryDef

        pQueryDef.Tables = "AddressPtRev,PremsInterPt"
        pQueryDef.SubFields = "AddressPtRev.ADDRESS_ID"
        pQueryDef.WhereClause = "AddressPtRev.ADDRESS_ID=PremsInterPt.ADDRESS_ID AND " & _
                                "AddressPtRev.STATUS='RETIRE'  AND " & _
                                "PremsInterPt.USE_PARSED_ADDRESS='Y' AND " & _
                                "NOT PremsInterPt.PEXPSTS IN ('PG','ST')"

        Dim pCursor As ICursor
        pCursor = pQueryDef.Evaluate

        Dim pWSEdit As IWorkspaceEdit = pWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pOldMar As IRow = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Retired MAR"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "C"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(0)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " exceptions.")
        Else
            log.WriteLine(Now & ": No exceptions.")
        End If

        log.WriteLine(Now & ": Check MAR Retired ended.")
    End Sub
#End Region

#Region "Check address component update"

    Sub CheckAddrUpdateNoExp()
        log.WriteLine(Now & ": Check Address Update Without Exception started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempMarUpdateTable As ITable = pFeatureWorkspace.OpenTable("MARPremiseUpdate")

        Dim pWSEdit As IWorkspaceEdit = pFeatureWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef

            pQueryDef.Tables = "AddressPtRev,PremsInterPt,STREET_TYPE"
            pQueryDef.SubFields = "PremsInterPt.PEXUID,PremsInterPt.ADDRESS_ID,AddressPtRev.ADDRNUM,AddressPtRev.ADDRNUMSUFFIX,AddressPtRev.STNAME,STREET_TYPE.USPS_ST_TYPE,AddressPtRev.QUADRANT"
            pQueryDef.WhereClause = "PremsInterPt.address_id = AddressPtRev.address_id " & _
                                    "AND AddressPtRev.STREET_TYPE = STREET_TYPE.STREET_TYPE " & _
                                    "AND NOT PremsInterPt.PEXPSTS IN ('PG','ST') " & _
                                    "AND USE_PARSED_ADDRESS = 'Y' " & _
                                    "AND " & _
                                    "( " & _
                                    "NOT Iif(IsNull(PremsInterPt.ADDRNUM),""0"",PremsInterPt.ADDRNUM) = Iif(IsNull(AddressPtRev.ADDRNUM),""0"",CStr(AddressPtRev.ADDRNUM)) " & _
                                    "OR NOT Iif(IsNull(PremsInterPt.STNAME),""0"",PremsInterPt.STNAME) = Iif(IsNull(AddressPtRev.STNAME),""0"",AddressPtRev.STNAME) " & _
                                    "OR NOT STREET_TYPE.USPS_ST_TYPE = Iif(IsNull(PremsInterPt.STREET_TYPE),""0"",PremsInterPt.STREET_TYPE) " & _
                                    "OR NOT Iif(IsNull(PremsInterPt.ADDRNUMSUF), ""0"",PremsInterPt.ADDRNUMSUF) = Iif(IsNull(AddressPtRev.ADDRNUMSUFFIX), ""0"",AddressPtRev.ADDRNUMSUFFIX) " & _
                                    ") " & _
                                    "AND Iif(IsNull(PremsInterPt.QUADRANT),""0"",PremsInterPt.QUADRANT) = Iif(IsNull(AddressPtRev.QUADRANT),""0"",AddressPtRev.QUADRANT) " & _
                                    "AND " & _
                                    "( " & _
                                    "Iif(IsNull(PremsInterPt.ADDRNUMSUF), ""0"",PremsInterPt.ADDRNUMSUF) = Iif(IsNull(AddressPtRev.ADDRNUMSUFFIX), ""0"",AddressPtRev.ADDRNUMSUFFIX) " & _
                                    "OR PremsInterPt.ADDRNUMSUF IS NULL " & _
                                    ")"


            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate

            Dim pOldMar As IRow = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempMarUpdateTable.CreateRow
                pRow.Value(pRow.Fields.FindField("PEXUID")) = pOldMar.Value(0)
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(1)
                pRow.Value(pRow.Fields.FindField("ADDRNUM")) = pOldMar.Value(2)
                pRow.Value(pRow.Fields.FindField("ADDRNUMSUF")) = pOldMar.Value(3)
                pRow.Value(pRow.Fields.FindField("STNAME")) = pOldMar.Value(4)
                pRow.Value(pRow.Fields.FindField("STREET_TYPE")) = pOldMar.Value(5)
                pRow.Value(pRow.Fields.FindField("QUADRANT")) = pOldMar.Value(6)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " update.")
        Else
            log.WriteLine(Now & ": No update.")
        End If

        log.WriteLine(Now & ": Check Address Update Without Exception ended.")
    End Sub

    Sub CheckAddrUpdateExp()
        log.WriteLine(Now & ": Check Address Update With Exception started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pFeatureWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef

            pQueryDef.Tables = "AddressPtRev,PremsInterPt,STREET_TYPE"
            pQueryDef.SubFields = "AddressPtRev.ADDRESS_ID"
            pQueryDef.WhereClause = "PremsInterPt.address_id = AddressPtRev.address_id " & _
                                    "AND PremsInterPt.STREET_TYPE = STREET_TYPE.USPS_ST_TYPE " & _
                                    "AND NOT PremsInterPt.PEXPSTS IN ('PG','ST') " & _
                                    "AND USE_PARSED_ADDRESS = 'Y' " & _
                                    "AND (" & _
                                    "NOT Iif(IsNull(PremsInterPt.QUADRANT),""0"",PremsInterPt.QUADRANT) = Iif(IsNull(AddressPtRev.QUADRANT),""0"",AddressPtRev.QUADRANT) " & _
                                    "OR ( NOT Iif(IsNull(PremsInterPt.ADDRNUMSUF),""0"",PremsInterPt.ADDRNUMSUF) = Iif(IsNull(AddressPtRev.ADDRNUMSUFFIX),""0"",AddressPtRev.ADDRNUMSUFFIX) " & _
                                    "AND PremsInterPt.ADDRNUMSUF IS NOT NULL) )"

            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate

            Dim pOldMar As IRow = pCursor.NextRow
            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Review MAR Address Update"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "C"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(0)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " exceptions.")
        Else
            log.WriteLine(Now & ": No exceptions.")
        End If

        log.WriteLine(Now & ": Check Address Update With Exception ended.")
    End Sub
#End Region


    Sub MARAddrMatch()
        log.WriteLine(Now & ": MAR Address Match started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempMarUpdateTable As ITable = pFeatureWorkspace.OpenTable("MARPremiseUpdate")
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pFeatureWorkspace
        pWSEdit.StartEditing(False)
        Dim counter1 As Integer = 0
        Dim counter2 As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef

            pQueryDef.Tables = "AddressPtRev,PremsInterPt,STREET_TYPE"
            pQueryDef.SubFields = "PremsInterPt.PEXUID,Min(AddressPtRev.ADDRESS_ID),Count(*),AddressPtRev.ADDRNUM,AddressPtRev.ADDRNUMSUFFIX,AddressPtRev.STNAME,STREET_TYPE.USPS_ST_TYPE,AddressPtRev.QUADRANT"
            pQueryDef.WhereClause = "PremsInterPt.STREET_TYPE = STREET_TYPE.USPS_ST_TYPE " & _
                                    "AND PremsInterPt.QUADRANT = AddressPtRev.QUADRANT  " & _
                                    "AND PremsInterPt.STNAME = AddressPtRev.STNAME  " & _
                                    "AND Iif(IsNull(PremsInterPt.ADDRNUMSUF),""0"",PremsInterPt.ADDRNUMSUF) = Iif(IsNull(AddressPtRev.ADDRNUMSUFFIX),""0"",AddressPtRev.ADDRNUMSUFFIX) " & _
                                    "AND PremsInterPt.ADDRNUM = Iif(IsNull(AddressPtRev.ADDRNUM),""0"",CStr(AddressPtRev.ADDRNUM) )  " & _
                                    "AND STREET_TYPE.STREET_TYPE = AddressPtRev.STREET_TYPE  " & _
                                    "AND (PremsInterPt.ADDRESS_ID IS NULL OR PremsInterPt.ADDRESS_ID=0) " & _
                                    "AND PremsInterPt.USE_PARSED_ADDRESS='Y' " & _
                                    "AND NOT AddressPtRev.STATUS='RETIRED' " & _
                                    "AND NOT PremsInterPt.PEXPSTS IN ('PG','ST') " & _
                                    "GROUP BY PremsInterPt.PEXUID,AddressPtRev.ADDRNUM,AddressPtRev.ADDRNUMSUFFIX,AddressPtRev.STNAME,STREET_TYPE.USPS_ST_TYPE,AddressPtRev.QUADRANT"

            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate

            Dim pOldMar As IRow = pCursor.NextRow

            Do While Not pOldMar Is Nothing
                If pOldMar.Value(2) = 1 Then

                    Dim pRow As IRow = tempMarUpdateTable.CreateRow
                    pRow.Value(pRow.Fields.FindField("PEXUID")) = pOldMar.Value(0)
                    pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(1)
                    pRow.Value(pRow.Fields.FindField("ADDRNUM")) = pOldMar.Value(3)
                    pRow.Value(pRow.Fields.FindField("ADDRNUMSUF")) = pOldMar.Value(4)
                    pRow.Value(pRow.Fields.FindField("STNAME")) = pOldMar.Value(5)
                    pRow.Value(pRow.Fields.FindField("STREET_TYPE")) = pOldMar.Value(6)
                    pRow.Value(pRow.Fields.FindField("QUADRANT")) = pOldMar.Value(7)
                    pRow.Value(pRow.Fields.FindField("UPDATE_TYPE")) = "MAR_MATCH"
                    pRow.Store()

                    counter1 = counter1 + 1
                Else
                    Dim pRow As IRow = tempExpTable.CreateRow
                    pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Multiple MAR Potential Match"
                    pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "C"
                    pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(1)
                    pRow.Store()

                    counter2 = counter2 + 1
                End If

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": " & counter1 & " update.")
        log.WriteLine(Now & ": " & counter2 & " exceptions.")

        log.WriteLine(Now & ": MAR Address Match ended.")
    End Sub

    Sub MARPotentialMatch()
        log.WriteLine(Now & ": MAR Potential Match started.")

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempMarUpdateTable As ITable = pFeatureWorkspace.OpenTable("MARPremiseUpdate")
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        Dim pWSEdit As IWorkspaceEdit = pFeatureWorkspace
        pWSEdit.StartEditing(False)
        Dim counter As Integer = 0
        Try

            Dim pQueryDef As IQueryDef
            pQueryDef = pFeatureWorkspace.CreateQueryDef

            pQueryDef.Tables = "AddressPtRev,PremsInterPt,STREET_TYPE"
            pQueryDef.SubFields = "PremsInterPt.PEXUID,Min(AddressPtRev.ADDRESS_ID)"
            pQueryDef.WhereClause = "PremsInterPt.STREET_TYPE = STREET_TYPE.USPS_ST_TYPE " & _
                                    "AND PremsInterPt.QUADRANT = AddressPtRev.QUADRANT  " & _
                                    "AND PremsInterPt.STNAME = AddressPtRev.STNAME  " & _
                                    "AND PremsInterPt.ADDRNUMSUF IS NOT NULL " & _
                                    "AND AddressPtRev.ADDRNUMSUFFIX IS NULL " & _
                                    "AND PremsInterPt.ADDRNUM = Iif(IsNull(AddressPtRev.ADDRNUM),""0"",CStr(AddressPtRev.ADDRNUM) )  " & _
                                    "AND STREET_TYPE.STREET_TYPE = AddressPtRev.STREET_TYPE  " & _
                                    "AND (PremsInterPt.ADDRESS_ID IS NULL OR PremsInterPt.ADDRESS_ID=0) " & _
                                    "AND PremsInterPt.USE_PARSED_ADDRESS='Y' " & _
                                    "AND NOT AddressPtRev.STATUS='RETIRED' " & _
                                    "AND NOT PremsInterPt.PEXPSTS IN ('PG','ST') " & _
                                    "GROUP BY PremsInterPt.PEXUID"

            Dim pCursor As ICursor
            pCursor = pQueryDef.Evaluate

            Dim pOldMar As IRow = pCursor.NextRow

            Do While Not pOldMar Is Nothing
                Dim pRow As IRow = tempExpTable.CreateRow
                pRow.Value(pRow.Fields.FindField("ExceptionType")) = "Review MAR Potential Match"
                pRow.Value(pRow.Fields.FindField("REFDATASRC")) = "C"
                pRow.Value(pRow.Fields.FindField("ADDRESS_ID")) = pOldMar.Value(1)
                pRow.Store()

                counter = counter + 1

                pOldMar = pCursor.NextRow
            Loop

            pWSEdit.StopEditing(True)
        Catch ex As Exception
            pWSEdit.StopEditing(False)
            Throw ex
        End Try

        If counter > 0 Then
            log.WriteLine(Now & ": Found " & counter & " exceptions.")
        Else
            log.WriteLine(Now & ": No exceptions.")
        End If

        log.WriteLine(Now & ": MAR Potential Match ended.")
    End Sub


#Region "Update Premise layer"

    Sub UpdatePremise()
        log.WriteLine(Now & ": Update Premise layer started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        If Not m_jtxJob.VersionExists Then
            m_jtxJob.CreateVersion(esriVersionAccess.esriVersionAccessProtected)
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String
        Dim jtxJobParentVersion As String

        Dim tableExceptionsaddresspt As String
        Dim tableCisoutputPremExp As String
        Dim tablePremsInterPt As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateMAR")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSADDRESSPT" Then
                    tableExceptionsaddresspt = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTPREMEXP" Then
                    tableCisoutputPremExp = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "PREMISEPTLAYER" Then
                    tablePremsInterPt = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARPremiseUpdate")

        Dim pWS As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, m_jtxJob.VersionName)
        Dim pWorkspaceEdit As IWorkspaceEdit = pWS

        Dim pPremsInterPtTable As ITable = CType(pWS, IFeatureWorkspace).OpenTable(tablePremsInterPt)
        Dim pPrmsFilter As IQueryFilter = New QueryFilter
        Dim pPrmsCursor As ICursor

        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim pCursor As ICursor = tempExpTable.Search(Nothing, False)
        Dim pRow As IRow = pCursor.NextRow

        Try
            Dim counter As Integer = 0
            Do While Not pRow Is Nothing

                pPrmsFilter.WhereClause = "PEXUID=" & MapUtil.GetRowValue(pRow, "PEXUID")
                pPrmsCursor = pPremsInterPtTable.Search(pPrmsFilter, False)
                Dim pPremsPt As IRow = pPrmsCursor.NextRow
                If MapUtil.GetRowValue(pRow, "UPDATE_TYPE") = "MAR_MATCH" Then
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("ADDRESS_ID"), MapUtil.GetRowValue(pRow, "ADDRESS_ID"))
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("LOCTNPRECS"), "6", True)
                Else
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("ADDRNUM"), MapUtil.GetRowValue(pRow, "ADDRNUM"))
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("ADDRNUMSUF"), MapUtil.GetRowValue(pRow, "ADDRNUMSUF"))
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("STNAME"), MapUtil.GetRowValue(pRow, "STNAME"))
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("STREET_TYPE"), MapUtil.GetRowValue(pRow, "STREET_TYPE"))
                    MapUtil.SetRowValue(pPremsPt, pPremsPt.Fields.FindField("QUADRANT"), MapUtil.GetRowValue(pRow, "QUADRANT"))
                End If

                pPremsPt.Store()
                Marshal.ReleaseComObject(pPrmsCursor)

                counter = counter + 1
                pRow = pCursor.NextRow
            Loop
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counter & " rows updated.")
        Catch ex As Exception
            log.WriteLine(ex.StackTrace)
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Update Premise layer completed.")

    End Sub

#End Region

#Region "Write exceptions into DB"

    Sub CreateMarUpdateException()
        log.WriteLine(Now & ": Create MAR weekly update exceptions started.")

        m_jtxJob = GetJTXJob()
        If m_jtxJob Is Nothing Then
            Throw New Exception("JTX job has not been created")
        End If

        If Not m_jtxJob.VersionExists Then
            m_jtxJob.CreateVersion(esriVersionAccess.esriVersionAccessProtected)
        End If

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String
        Dim jtxJobParentVersion As String

        Dim tableExceptionsaddresspt As String
        Dim tableCisoutputPremExp As String
        Dim tablePremsInterPt As String

        Dim tolerance As String = ""

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateMAR")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_EXCEPTIONSADDRESSPT" Then
                    tableExceptionsaddresspt = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTPREMEXP" Then
                    tableCisoutputPremExp = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "PREMISEPTLAYER" Then
                    tablePremsInterPt = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)

        Dim tempDbPath As String = Environment.CurrentDirectory & "\MarExceptionTemp.mdb"
        Dim pPropset As IPropertySet = New PropertySet
        pPropset.SetProperty("DATABASE", tempDbPath)

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactory
        Dim pWorkspace As IWorkspace = workspaceFactory.Open(pPropset, 0)

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
        Dim tempExpTable As ITable = pFeatureWorkspace.OpenTable("MARExceptionTemp")

        'Dim pFilter As IQueryFilter = New QueryFilter
        'pFilter.WhereClause = "ExceptionType <> 'Premise Address Update' "
        Dim pCursor As ICursor = tempExpTable.Search(Nothing, False)
        Dim pRow As IRow = pCursor.NextRow

        Dim pWS As IWorkspace = pJTXDatabase.DataWorkspace(m_jtxJob.VersionName)
        Dim pWorkspaceEdit As IWorkspaceEdit = pWS

        Dim pMarExpTable As ITable = CType(pWS, IFeatureWorkspace).OpenTable(tableExceptionsaddresspt)
        Dim pCisExpTable As ITable = CType(pWS, IFeatureWorkspace).OpenTable(tableCisoutputPremExp)

        Dim pPremsInterPtTable As ITable = CType(pWS, IFeatureWorkspace).OpenTable(tablePremsInterPt)
        Dim pPrmsFilter As IQueryFilter = New QueryFilter
        Dim pPrmsCursor As ICursor

        Dim pMarExpRow As IRow

        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()

        Dim marExpCount As Integer = 0
        Dim cisExpCount As Integer = 0

        Dim idxExceptionType As Integer = tempExpTable.FindField("ExceptionType")

        Try
            Dim counter As Integer = 0
            Do While Not pRow Is Nothing
                pMarExpRow = pMarExpTable.CreateRow
                MapUtil.SetRowValue(pMarExpRow, pMarExpRow.Fields.FindField("EXCEPTDT"), Now)
                MapUtil.SetRowValue(pMarExpRow, pMarExpRow.Fields.FindField("EXCEPTTYP"), pRow.Value(idxExceptionType))
                MapUtil.SetRowValue(pMarExpRow, pMarExpRow.Fields.FindField("ADDRESS_ID"), MapUtil.GetRowValue(pRow, "ADDRESS_ID"))
                MapUtil.SetRowValue(pMarExpRow, pMarExpRow.Fields.FindField("REFDATASRC"), MapUtil.GetRowValue(pRow, "REFDATASRC"))

                pMarExpRow.Store()
                counter = counter + 1

                pRow = pCursor.NextRow
            Loop
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            log.WriteLine(Now & ": " & counter & " exception(s) were added to the database.")
        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)
            Throw ex
        End Try

        log.WriteLine(Now & ": Create MAR weekly update exceptions completed.")
    End Sub

#End Region

#Region "JTX Job"


    Sub CreateMARJTXJob()
        log.WriteLine(Now & ": Create MAR WEEKLY UPDATE Exception job started.")

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
        Dim marExceptionType As String

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateMAR")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JOB_PARENT_VERSION" Then
                    jtxJobParentVersion = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "MAR_EXCEPTION_TYPE" Then
                    marExceptionType = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        log.WriteLine(Now & ": Creating jtx job with following parameters.")
        log.WriteLine(vbTab & "EXCEPTION TYPE: " & vbTab & marExceptionType)
        log.WriteLine(vbTab & "JTX DATABASE: " & vbTab & jtxDB)
        log.WriteLine(vbTab & "JTX DB WORKSPACE: " & vbTab & jtxDBWorkspace)
        log.WriteLine(vbTab & "PARENT VERSION: " & vbTab & jtxJobParentVersion)

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)


        m_jtxJob = mdlJTX.CreateJob(pJTXDatabase, marExceptionType, jtxDBWorkspace, jtxJobParentVersion, log)
        Console.WriteLine("JTX job was created. JTX job ID " & m_jtxJob.ID)
        log.WriteLine(Now & ":JTX job was created. JTX job ID " & m_jtxJob.ID)

        Try
            Dim oWriter As System.IO.StreamWriter = File.CreateText(Environment.CurrentDirectory & "\jtxjob.txt")
            oWriter.WriteLine(m_jtxJob.ID)
            oWriter.Close()
        Catch ex As Exception
            If Not m_jtxJob Is Nothing Then
                pJTXDatabase.JobManager.DeleteJob(m_jtxJob.ID)
                Console.WriteLine("JTX job was deleted because of system exception")
                log.WriteLine(Now & ":JTX job was deleted because of system exception")
            End If

            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        End Try



        log.WriteLine(Now & ": Create MAR WEEKLY UPDATE Exception job completed. New Job ID " & m_jtxJob.ID)
    End Sub

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

            m_nodelist = m_xmld.SelectNodes("IAISConfig/UpdateMAR")

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


    Public Function CreateRelQueryTable(ByVal targetTable As ITable, ByVal joinTable As ITable, ByVal fromField As String, ByVal toField As String) As ITable

        ' Build a memory relationship class.
        Dim memRelClassFactory As IMemoryRelationshipClassFactory = New MemoryRelationshipClassFactoryClass()
        Dim relationshipClass As IRelationshipClass = memRelClassFactory.Open("MemRelClass", CType(targetTable, IObjectClass), fromField, CType(joinTable, IObjectClass), toField, "forward", "backward", esriRelCardinality.esriRelCardinalityOneToOne)

        ' Open the RelQueryTable as a feature class.
        Dim rqtFactory As IRelQueryTableFactory = New RelQueryTableFactoryClass()
        Dim relQueryTable As ITable = CType(rqtFactory.Open(relationshipClass, True, Nothing, Nothing, "", False, False), ITable)

        Return relQueryTable

    End Function


End Module
