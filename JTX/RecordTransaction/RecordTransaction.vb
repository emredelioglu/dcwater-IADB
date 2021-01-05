Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.JTX.Utilities
Imports ESRI.ArcGIS.Geodatabase


<ComClass(RecordTransaction.ClassId, RecordTransaction.InterfaceId, RecordTransaction.EventsId)>
<System.Serializable()>
Public Class RecordTransaction
    Implements IJTXCustomStep


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f95bf789-9474-4059-8aa1-616c3077f2fd"
    Public Const InterfaceId As String = "d916f70e-f283-4a01-a4ac-9e48880a2209"
    Public Const EventsId As String = "67182a26-f8bc-4a46-a94e-7166ee00298a"
#End Region

#Region "Registration Code"
    <Runtime.InteropServices.ComRegisterFunction()>
    Private Shared Sub Reg(ByVal regKey As String)
        ESRI.ArcGIS.JTX.Utilities.JTXUtilities.RegisterJTXCustomStep(regKey)
    End Sub

    <Runtime.InteropServices.ComUnregisterFunction()>
    Private Shared Sub Unreg(ByVal regKey As String)
        ESRI.ArcGIS.JTX.Utilities.JTXUtilities.UnregisterJTXCustomStep(regKey)
    End Sub
#End Region


    Private m_ipDatabase As IJTXDatabase

    Private m_checkCISException As Boolean

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    <System.Runtime.InteropServices.DispId(1)> Public ReadOnly Property ArgumentDescriptions() As String Implements ESRI.ArcGIS.JTX.IJTXCustomStep.ArgumentDescriptions
        Get
            Return "Check CIS Exception"
        End Get
    End Property

    Public Function Execute(ByVal JobID As Integer, ByVal StepID As Integer, ByRef argv() As Object, ByRef ipFeedback As ESRI.ArcGIS.JTX.IJTXCustomStepFeedback) As Integer Implements ESRI.ArcGIS.JTX.IJTXCustomStep.Execute


        m_checkCISException = False

        If argv Is Nothing Then
            m_checkCISException = False
        End If

        If argv.Length = 0 Then
            m_checkCISException = False
        End If

        If argv.Length > 0 Then
            If UCase(argv(0)) = "/CHECKCISEXCEPTIONS" Then
                m_checkCISException = True
            Else
                m_checkCISException = False
            End If
        End If

        Dim transMananger As IJTXTransactionManager = m_ipDatabase.TransactionManager
        Dim pJobManager As IJTXJobManager = m_ipDatabase.JobManager

        Dim pJob As IJTXJob2 = pJobManager.GetJob(JobID)
        Dim pJobDataWorkspace As IJTXDataWorkspaceName = pJob.ActiveDatabase



        If Not pJob.VersionExists Then
            MsgBox("Job version has not been defined")
            Return 1
        End If

        Dim registryKey As Microsoft.Win32.RegistryKey
        registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\IAIS")
        If registryKey Is Nothing Then
            MsgBox("Can't find system registry for SYSTEM_TABLE")
            Return 1
        End If
        Dim tnIAISSetting As String = registryKey.GetValue("SYSTEM_TABLE")

        If tnIAISSetting Is Nothing Or tnIAISSetting = "" Then
            MsgBox("Invalid SYSTEM_TABLE value in system registry")
            Return 1
        End If

        Dim ids As Dictionary(Of String, String) = New Dictionary(Of String, String)

        Dim pJTXSeurityManager As IJTXSecurityManager2 = New JTXSecurityManager()
        Dim pFWS As IFeatureWorkspace
        pFWS = pJTXSeurityManager.GetDataWorkspace(pJob.ActiveDatabase.DatabaseID, pJob.ParentVersion)


        Dim jtxWorkspace As IFeatureWorkspace = pJTXSeurityManager.GetJTXWorkspace("", False)
        Dim jtxHistoryTable As ITable = jtxWorkspace.OpenTable("JTX_HISTORY_SESSIONS")

        'Dim pFWS As IFeatureWorkspace = m_ipDatabase.DataWorkspace(pJob.ParentVersion)
        Dim pWorkspaceEdit As IWorkspaceEdit = pFWS

        pWorkspaceEdit.StartEditing(False)
        pWorkspaceEdit.StartEditOperation()
        Try

            Dim transactionDate As Date = Now
            Dim pFilter As IQueryFilter = New QueryFilter
            Dim SystemSettingTable As ITable = pFWS.OpenTable(tnIAISSetting)

            Dim tnCISOUTPUTPREMEXP As String = Nothing
            Dim tnPREMSINTERPT As String = Nothing
            Dim tnIACF As String = Nothing
            Dim tnCISOUTPUTBILLDET As String = Nothing

            Dim tnSTCROSSWALK As String = Nothing

            Dim tnEXCEPTIONSPREMSEXPOUT As String = Nothing

            Dim expFlag As Boolean = False

            pFilter.WhereClause = "NAME IN ('TABLE_IACF','TABLE_CISOUTPUTBILLDET','TABLE_CISOUTPUTPREMEXP','LAYER_PREMSINTERPT','TABLE_STCROSSWALK','TABLE_EXCEPTIONSPREMSEXPOUT')"
            Dim pCursor As ICursor = SystemSettingTable.Search(pFilter, False)
            Dim pRow As IRow = pCursor.NextRow
            While Not pRow Is Nothing
                Dim paraName As String = GetRowValue(pRow, "NAME")
                Dim paraValue As String = GetRowValue(pRow, "VALUE")

                If paraName = "TABLE_IACF" Then
                    tnIACF = paraValue
                ElseIf paraName = "TABLE_CISOUTPUTBILLDET" Then
                    tnCISOUTPUTBILLDET = paraValue
                ElseIf paraName = "TABLE_CISOUTPUTPREMEXP" Then
                    tnCISOUTPUTPREMEXP = paraValue
                ElseIf paraName = "LAYER_PREMSINTERPT" Then
                    tnPREMSINTERPT = paraValue
                ElseIf paraName = "TABLE_STCROSSWALK" Then
                    tnSTCROSSWALK = paraValue
                ElseIf paraName = "TABLE_EXCEPTIONSPREMSEXPOUT" Then
                    tnEXCEPTIONSPREMSEXPOUT = paraValue
                End If

                pRow = pCursor.NextRow
            End While


            Marshal.ReleaseComObject(pCursor)

            If tnIACF Is Nothing Then
                MsgBox("Can't find IACF table name in system setting table")
                expFlag = True
            End If

            If tnCISOUTPUTBILLDET Is Nothing Then
                MsgBox("Can't find CISOUTPUTBILLDET table name in system setting table")
                expFlag = True
            End If

            If tnCISOUTPUTPREMEXP Is Nothing Then
                MsgBox("Can't find CISOUTPUTPREMEXP table name in system setting table")
                expFlag = True
            End If

            If tnPREMSINTERPT Is Nothing Then
                MsgBox("Can't find PREMSINTERPT table name in system setting table")
                expFlag = True
            End If

            If tnSTCROSSWALK Is Nothing Then
                MsgBox("Can't find STCROSSWALK table name in system setting table")
                expFlag = True
            End If

            If tnEXCEPTIONSPREMSEXPOUT Is Nothing Then
                MsgBox("Can't find EXCEPTIONSPREMSEXPOUT table name in system setting table")
                expFlag = True
            End If

            If expFlag Then
                Throw New Exception("System setting table exception")
            End If

            Dim premsTableH As ITable = pFWS.OpenTable(tnPREMSINTERPT + "_H")
            Dim premsTable As ITable = pFWS.OpenTable(tnPREMSINTERPT)

            Dim newFilter As IQueryFilter = New QueryFilter
            newFilter.WhereClause = "JOB_ID=" & JobID
            Dim now2 As Date = Now
            Dim historyCursor As ICursor = jtxHistoryTable.Search(newFilter, False)

            Dim stcrosswalkTable As ITable = pFWS.OpenTable(tnSTCROSSWALK)
            Dim iacfTable As ITable = pFWS.OpenTable(tnIACF)
            Dim cisoutputbilldetTable As ITable = pFWS.OpenTable(tnCISOUTPUTBILLDET)
            Dim expTable As ITable = pFWS.OpenTable(tnCISOUTPUTPREMEXP)
            'Dim transactionSets As IJTXTransactionSet = transMananger.GetLoggedTransactions(pFilter)



            Dim tran As IJTXTransaction2

            Dim expCursor As ICursor

            Dim expRow As IRow


            Dim premsRow As IRow
            Dim premsMasterRow As IRow


            Dim stname As String = ""
            Dim ia_charge As String = ""



            Dim countTrans As Integer = 0
            Dim countMADDR As Integer = 0
            Dim countMSSL As Integer = 0
            Dim countMSTAT As Integer = 0
            Dim countMTYPE As Integer = 0
            Dim countMXMPT As Integer = 0
            Dim countMXREF As Integer = 0
            Dim countMNPRM As Integer = 0
            Dim countIACF As Integer = 0

            Dim where As String = ""
            Dim historyRow As IRow = historyCursor.NextRow
            While Not historyRow Is Nothing
                Dim time As DateTime = historyRow.Value(historyRow.Fields.FindField("ARCHIVE_TIME"))

                where = where + "GDB_FROM_DATE = CONVERT(DATETIME,'" + time.ToString("yyyy-MM-dd HH:mm:ss") + "') or GDB_TO_DATE = CONVERT(DATETIME,'" + time.ToString("yyyy-MM-dd HH:mm:ss") + "') or "

                historyRow = historyCursor.NextRow
            End While


            If where.Trim.Length = 0 Then
                Throw New Exception("No Transactions found. Post was not done against default.")
            End If
            Dim qf As IQueryFilter = New QueryFilter
            qf.WhereClause = where.Substring(0, where.Length - 4)
            Dim qfd As IQueryFilterDefinition = qf
            qfd.PostfixClause = "ORDER BY OBJECTID, GDB_FROM_DATE, GDB_TO_DATE"

            Dim premsCursor As ICursor = premsTableH.Search(qfd, False)



            ia_charge = ""
            premsRow = premsCursor.NextRow

            While Not premsRow Is Nothing
                Dim pexuid As String = GetRowValue(premsRow, "PEXUID")
                Dim master_pexuid As String = GetRowValue(premsRow, "MASTER_PEXUID")

                If master_pexuid <> "" And master_pexuid <> "0" Then
                    pFilter.WhereClause = "PEXUID=" & master_pexuid
                    pCursor = premsTable.Search(pFilter, False)
                    premsMasterRow = pCursor.NextRow
                    ia_charge = GetRowValue(premsMasterRow, "PEXPRM")
                    Marshal.ReleaseComObject(pCursor)
                End If

                Dim useParsedAddress As String = GetRowValue(premsRow, "USE_PARSED_ADDRESS")

                Dim row3 As IRow = premsCursor.NextRow


                ' todo check this
                If row3 Is Nothing Or GetRowValue(premsRow, "OBJECTID", False) <> GetRowValue(row3, "OBJECTID", False) Then

                    Dim gdbDateTime As DateTime = GetRowValue(premsRow, "GDB_TO_DATE", False)
                    If gdbDateTime <> now2 Then


                        countMNPRM = countMNPRM + 1

                        pFilter.WhereClause = "PEXUID=" & GetRowValue(premsRow, "PEXUID") & " AND TRANSACTIONCODE='MNPRM' "
                        expCursor = expTable.Search(pFilter, False)
                        expRow = expCursor.NextRow
                        If expRow Is Nothing Then
                            expRow = expTable.CreateRow
                        Else
                            'check to see if the EXCEPTIONMSGID is null
                            If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                Throw New Exception("Premise " & pexuid & " with transaction code MNPRM has unresolved exception")
                            End If
                        End If

                        expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MNPRM"
                        expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))
                        expRow.Value(expRow.Fields.FindField("PEXPTYP")) = premsRow.Value(premsRow.Fields.FindField("PEXPTYP"))
                        expRow.Value(expRow.Fields.FindField("PEXPSTS")) = premsRow.Value(premsRow.Fields.FindField("PEXPSTS"))
                        expRow.Value(expRow.Fields.FindField("ADDRESSSPLITFLAG")) = premsRow.Value(premsRow.Fields.FindField("USE_PARSED_ADDRESS"))

                        If GetRowValue(premsRow, "USE_PARSED_ADDRESS") = "Y" Then

                            expRow.Value(expRow.Fields.FindField("ADDRNUM")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUM"))
                            If (IsDBNull(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")))) Then
                                expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF"))
                            Else
                                expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = Left(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")), 3)
                            End If
                            stname = GetUSPSST(GetRowValue(premsRow, "STNAME"), stcrosswalkTable)
                            expRow.Value(expRow.Fields.FindField("STNAME")) = stname
                            expRow.Value(expRow.Fields.FindField("STREET_TYPE")) = premsRow.Value(premsRow.Fields.FindField("STREET_TYPE"))
                            expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                            expRow.Value(expRow.Fields.FindField("UNIT")) = premsRow.Value(premsRow.Fields.FindField("UNIT"))
                        Else
                            expRow.Value(expRow.Fields.FindField("PEXSAD")) = premsRow.Value(premsRow.Fields.FindField("PEXSAD"))
                            expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                        End If

                        expRow.Value(expRow.Fields.FindField("PEXPZIP")) = premsRow.Value(premsRow.Fields.FindField("PEXPZIP"))
                        expRow.Value(expRow.Fields.FindField("IS_IMPERVIOUS_ONLY")) = premsRow.Value(premsRow.Fields.FindField("IS_IMPERVIOUS_ONLY"))
                        expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = premsRow.Value(premsRow.Fields.FindField("IS_EXEMPT_IAB"))
                        expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = premsRow.Value(premsRow.Fields.FindField("EXEMPT_IAB_REASON"))

                        If ia_charge = "" Then
                            expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                        Else
                            expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                        End If

                        expRow.Value(expRow.Fields.FindField("PEXSAD")) = premsRow.Value(premsRow.Fields.FindField("PEXSAD"))

                        expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                        expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                        expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                        expRow.Store()



                        Marshal.ReleaseComObject(expCursor)
                    End If

                    premsRow = row3

                Else


                    Dim rowValue4 As String = GetRowValue(row3, "MASTER_PEXUID")

                    If rowValue4 <> "" And rowValue4 <> "0" Then
                        pFilter.WhereClause = "PEXUID=" & rowValue4
                        pCursor = premsTable.Search(pFilter, False)
                        premsMasterRow = pCursor.NextRow
                        ia_charge = GetRowValue(premsMasterRow, "PEXPRM")
                        Marshal.ReleaseComObject(pCursor)
                    End If


                    If GetRowValue(row3, "PEXPRM") = "" Then
                        'A pending premise updated

                        countMNPRM = countMNPRM + 1

                        pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MNPRM' "
                        expCursor = expTable.Search(pFilter, False)
                        expRow = expCursor.NextRow
                        If expRow Is Nothing Then
                            expRow = expTable.CreateRow
                        Else
                            'check to see if the EXCEPTIONMSGID is null
                            If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                Throw New Exception("Premise " & pexuid & " with transaction code MNPRM has unresolved exception")
                            End If
                        End If

                        expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MNPRM"
                        expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))
                        expRow.Value(expRow.Fields.FindField("PEXPTYP")) = row3.Value(row3.Fields.FindField("PEXPTYP"))
                        expRow.Value(expRow.Fields.FindField("PEXPSTS")) = row3.Value(row3.Fields.FindField("PEXPSTS"))
                        expRow.Value(expRow.Fields.FindField("ADDRESSSPLITFLAG")) = row3.Value(row3.Fields.FindField("USE_PARSED_ADDRESS"))

                        If GetRowValue(row3, "USE_PARSED_ADDRESS") = "Y" Then

                            expRow.Value(expRow.Fields.FindField("ADDRNUM")) = row3.Value(row3.Fields.FindField("ADDRNUM"))
                            If (IsDBNull(row3.Value(row3.Fields.FindField("ADDRNUMSUF")))) Then
                                expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = row3.Value(row3.Fields.FindField("ADDRNUMSUF"))
                            Else
                                expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = Left(row3.Value(row3.Fields.FindField("ADDRNUMSUF")), 3)
                            End If
                            stname = GetUSPSST(GetRowValue(row3, "STNAME"), stcrosswalkTable)
                            expRow.Value(expRow.Fields.FindField("STNAME")) = stname
                            expRow.Value(expRow.Fields.FindField("STREET_TYPE")) = row3.Value(row3.Fields.FindField("STREET_TYPE"))
                            expRow.Value(expRow.Fields.FindField("QUADRANT")) = row3.Value(row3.Fields.FindField("QUADRANT"))
                            expRow.Value(expRow.Fields.FindField("UNIT")) = row3.Value(row3.Fields.FindField("UNIT"))
                        Else
                            expRow.Value(expRow.Fields.FindField("PEXSAD")) = row3.Value(row3.Fields.FindField("PEXSAD"))
                            expRow.Value(expRow.Fields.FindField("QUADRANT")) = row3.Value(row3.Fields.FindField("QUADRANT"))
                        End If



                        expRow.Value(expRow.Fields.FindField("PEXPZIP")) = row3.Value(row3.Fields.FindField("PEXPZIP"))
                        expRow.Value(expRow.Fields.FindField("IS_IMPERVIOUS_ONLY")) = row3.Value(row3.Fields.FindField("IS_IMPERVIOUS_ONLY"))
                        expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = row3.Value(row3.Fields.FindField("IS_EXEMPT_IAB"))
                        expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = row3.Value(row3.Fields.FindField("EXEMPT_IAB_REASON"))
                        If ia_charge = "" Then
                            expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                        Else
                            expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                        End If


                        expRow.Value(expRow.Fields.FindField("PEXSAD")) = row3.Value(row3.Fields.FindField("PEXSAD"))

                        expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                        expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                        expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                        expRow.Store()

                        Marshal.ReleaseComObject(expCursor)

                    Else
                        'This is a update
                        'MADDR
                        'MSSL
                        'MSTAT
                        'MTYPE
                        'MXMPT
                        'MXREF

                        If HasTxType("MADDR", premsRow, row3, useParsedAddress) Then

                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MADDR' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MADDR has unresolved exception")
                                End If
                            End If

                            countMADDR = countMADDR + 1

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MADDR"
                            expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                            expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))

                            expRow.Value(expRow.Fields.FindField("ADDRESSSPLITFLAG")) = row3.Value(row3.Fields.FindField("USE_PARSED_ADDRESS"))

                            If GetRowValue(row3, "USE_PARSED_ADDRESS") = "Y" Then
                                expRow.Value(expRow.Fields.FindField("ADDRNUM")) = row3.Value(row3.Fields.FindField("ADDRNUM"))
                                If (IsDBNull(row3.Value(row3.Fields.FindField("ADDRNUMSUF")))) Then
                                    expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = row3.Value(row3.Fields.FindField("ADDRNUMSUF"))
                                Else
                                    expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = Left(row3.Value(row3.Fields.FindField("ADDRNUMSUF")), 3)
                                End If

                                stname = GetUSPSST(GetRowValue(row3, "STNAME"), stcrosswalkTable)
                                expRow.Value(expRow.Fields.FindField("STNAME")) = stname
                                expRow.Value(expRow.Fields.FindField("STREET_TYPE")) = row3.Value(row3.Fields.FindField("STREET_TYPE"))
                                expRow.Value(expRow.Fields.FindField("QUADRANT")) = row3.Value(row3.Fields.FindField("QUADRANT"))
                                expRow.Value(expRow.Fields.FindField("UNIT")) = row3.Value(row3.Fields.FindField("UNIT"))
                            Else
                                expRow.Value(expRow.Fields.FindField("PEXSAD")) = row3.Value(row3.Fields.FindField("PEXSAD"))
                                expRow.Value(expRow.Fields.FindField("QUADRANT")) = row3.Value(row3.Fields.FindField("QUADRANT"))
                            End If

                            expRow.Value(expRow.Fields.FindField("PEXPZIP")) = row3.Value(row3.Fields.FindField("PEXPZIP"))


                            expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                            expRow.Store()
                            Marshal.ReleaseComObject(expCursor)
                        End If

                        If HasTxType("MSSL", premsRow, row3) Then
                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MSSL' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MSSL has unresolved exception")
                                End If
                            End If

                            countMSSL = countMSSL + 1

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MSSL"
                            expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                            expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))
                            '**added otr use code
                            If IsDBNull(row3.Value(row3.Fields.FindField("OTRUSECODE"))) Then
                                expRow.Value(expRow.Fields.FindField("OTRUSECODE")) = "0"
                            Else
                                expRow.Value(expRow.Fields.FindField("OTRUSECODE")) = row3.Value(row3.Fields.FindField("OTRUSECODE"))
                            End If
                            expRow.Value(expRow.Fields.FindField("PEXSQUARE")) = row3.Value(row3.Fields.FindField("PEXSQUARE"))
                            expRow.Value(expRow.Fields.FindField("PEXLOT")) = row3.Value(row3.Fields.FindField("PEXLOT"))
                            expRow.Value(expRow.Fields.FindField("PEXSUFFIX")) = row3.Value(row3.Fields.FindField("PEXSUFFIX"))
                            expRow.Value(expRow.Fields.FindField("WARD")) = row3.Value(row3.Fields.FindField("WARD"))

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                            expRow.Store()
                            Marshal.ReleaseComObject(expCursor)
                        End If

                        If HasTxType("MSTAT", premsRow, row3) Then

                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MSTAT' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MSTAT has unresolved exception")
                                End If
                            End If

                            If row3.Value(row3.Fields.FindField("PEXPSTS")) = "PG" Then
                                countMSTAT = countMSTAT + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MSTAT"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))

                                expRow.Value(expRow.Fields.FindField("PEXPSTS")) = row3.Value(row3.Fields.FindField("PEXPSTS"))
                                expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = row3.Value(row3.Fields.FindField("IS_EXEMPT_IAB"))
                                expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = row3.Value(row3.Fields.FindField("EXEMPT_IAB_REASON"))

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                                Marshal.ReleaseComObject(expCursor)
                            End If
                        End If

                        If HasTxType("MTYPE", premsRow, row3) Then
                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MTYPE' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MTYPE has unresolved exception")
                                End If
                            End If

                            countMTYPE = countMTYPE + 1

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MTYPE"
                            expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                            expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))

                            expRow.Value(expRow.Fields.FindField("PEXPTYP")) = row3.Value(row3.Fields.FindField("PEXPTYP"))

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                            expRow.Store()
                            Marshal.ReleaseComObject(expCursor)
                        End If

                        If HasTxType("MXMPT", premsRow, row3) Then
                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MXMPT' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MXMPT has unresolved exception")
                                End If
                            End If

                            countMXMPT = countMXMPT + 1

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MXMPT"
                            expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                            expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))

                            expRow.Value(expRow.Fields.FindField("IS_IMPERVIOUS_ONLY")) = row3.Value(row3.Fields.FindField("IS_IMPERVIOUS_ONLY"))
                            expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = row3.Value(row3.Fields.FindField("IS_EXEMPT_IAB"))
                            expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = row3.Value(row3.Fields.FindField("EXEMPT_IAB_REASON"))
                            expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                            expRow.Store()
                            Marshal.ReleaseComObject(expCursor)
                        End If

                        If HasTxType("MXREF", premsRow, row3) Then
                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND TRANSACTIONCODE='MXREF' "
                            expCursor = expTable.Search(pFilter, False)
                            expRow = expCursor.NextRow
                            If expRow Is Nothing Then
                                expRow = expTable.CreateRow
                            Else
                                'check to see if the EXCEPTIONMSGID is null
                                If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                                    Throw New Exception("Premise " & pexuid & " with transaction code MXREF has unresolved exception")
                                End If
                            End If

                            countMXREF = countMXREF + 1

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MXREF"
                            expRow.Value(expRow.Fields.FindField("PEXPRM")) = row3.Value(row3.Fields.FindField("PEXPRM"))
                            expRow.Value(expRow.Fields.FindField("PEXUID")) = row3.Value(row3.Fields.FindField("PEXUID"))

                            If ia_charge = "" Then
                                expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                            Else
                                expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                            End If

                            expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                            expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                            expRow.Store()
                            Marshal.ReleaseComObject(expCursor)
                        End If

                        If HasTxType("IACF", premsRow, row3) Then
                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID") & " AND EFFENDDT IS NULL"
                            Dim pIacfCursor As ICursor = iacfTable.Search(pFilter, False)
                            Dim iacfRow As IRow = pIacfCursor.NextRow
                            If Not iacfRow Is Nothing Then
                                iacfRow.Value(iacfRow.Fields.FindField("EFFENDDT")) = Now
                                iacfRow.Store()
                            End If

                            Marshal.ReleaseComObject(pIacfCursor)

                            pFilter.WhereClause = "PEXUID=" & GetRowValue(row3, "PEXUID")
                            pIacfCursor = cisoutputbilldetTable.Search(pFilter, False)
                            iacfRow = pIacfCursor.NextRow
                            If iacfRow Is Nothing Then
                                'Create a new record with IASQFT = 0
                                iacfRow = cisoutputbilldetTable.CreateRow
                            End If

                            iacfRow.Value(iacfRow.Fields.FindField("PEXPRM")) = GetRowValue(row3, "PEXPRM")
                            iacfRow.Value(iacfRow.Fields.FindField("PEXUID")) = GetRowValue(row3, "PEXUID")
                            iacfRow.Value(iacfRow.Fields.FindField("IMPERV_BILLED_ERU")) = 0
                            iacfRow.Value(iacfRow.Fields.FindField("IMPERV_CREDIT_ERU")) = 0
                            iacfRow.Value(iacfRow.Fields.FindField("IMPERV_AREA_SQFT")) = 0
                            iacfRow.Value(iacfRow.Fields.FindField("TRANSACTIONDT")) = Now

                            iacfRow.Store()

                            countIACF = countIACF + 1

                            Marshal.ReleaseComObject(pIacfCursor)
                        End If
                    End If

                    premsRow = premsCursor.NextRow

                End If

            End While

            Marshal.ReleaseComObject(pCursor)



            'If pJob.JobType().Name = "Daily – Premise Export Exceptions" Then
            If m_checkCISException Then

                Dim pQueryDef As IQueryDef = pFWS.CreateQueryDef
                pQueryDef.Tables = tnEXCEPTIONSPREMSEXPOUT
                pQueryDef.SubFields = "DISTINCT PEXUID,TRANSACTIONCODE,CISEXCEPTDT"
                Dim pExpPremsOutCursor = pQueryDef.Evaluate

                Dim pexuid As String
                Dim transactionCode As String
                Dim cisexpdt As Date

                Dim pExpPremsOutRow As IRow = pExpPremsOutCursor.NextRow
                Do While Not pExpPremsOutRow Is Nothing

                    ia_charge = ""

                    pexuid = pExpPremsOutRow.Value(0)
                    transactionCode = pExpPremsOutRow.Value(1)
                    cisexpdt = pExpPremsOutRow.Value(2)

                    'find the old transaction with CISEXPDT if it has not been picked up by previous JTX transactions
                    pFilter.WhereClause = "PEXUID =" & pexuid & " AND TRANSACTIONCODE='" & transactionCode & "' AND CONVERT(VARCHAR(10), EXCEPTIONDT, 111) = '" & String.Format("{0:yyyy/MM/dd}", cisexpdt) & "'"
                    pCursor = expTable.Search(pFilter, False)
                    expRow = pCursor.NextRow

                    Marshal.ReleaseComObject(pCursor)

                    If Not expRow Is Nothing Then
                        If Not IsDBNull(expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID"))) Then
                            Throw New Exception("Premise " & pexuid & " with transaction code " & transactionCode & " has unresolved exception")
                        End If

                        pFilter.WhereClause = "PEXUID =" & pexuid
                        pCursor = premsTable.Search(pFilter, False)
                        premsRow = pCursor.NextRow

                        Dim master_pexuid As String = GetRowValue(premsRow, "MASTER_PEXUID")

                        If master_pexuid <> "" Then
                            pFilter.WhereClause = "PEXUID=" & master_pexuid
                            pCursor = premsTable.Search(pFilter, False)
                            premsMasterRow = pCursor.NextRow
                            ia_charge = GetRowValue(premsMasterRow, "PEXPRM")
                        End If


                        If Not premsRow Is Nothing Then
                            If transactionCode = "MNPRM" Then

                                countMNPRM = countMNPRM + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MNPRM"
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))
                                expRow.Value(expRow.Fields.FindField("PEXPTYP")) = premsRow.Value(premsRow.Fields.FindField("PEXPTYP"))
                                expRow.Value(expRow.Fields.FindField("PEXPSTS")) = premsRow.Value(premsRow.Fields.FindField("PEXPSTS"))
                                expRow.Value(expRow.Fields.FindField("ADDRESSSPLITFLAG")) = premsRow.Value(premsRow.Fields.FindField("USE_PARSED_ADDRESS"))

                                If GetRowValue(premsRow, "USE_PARSED_ADDRESS") = "Y" Then

                                    expRow.Value(expRow.Fields.FindField("ADDRNUM")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUM"))
                                    If (IsDBNull(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")))) Then
                                        expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF"))
                                    Else
                                        expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = Left(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")), 3)
                                    End If
                                    stname = GetUSPSST(GetRowValue(premsRow, "STNAME"), stcrosswalkTable)
                                    expRow.Value(expRow.Fields.FindField("STNAME")) = stname
                                    expRow.Value(expRow.Fields.FindField("STREET_TYPE")) = premsRow.Value(premsRow.Fields.FindField("STREET_TYPE"))
                                    expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                                    expRow.Value(expRow.Fields.FindField("UNIT")) = premsRow.Value(premsRow.Fields.FindField("UNIT"))
                                Else
                                    expRow.Value(expRow.Fields.FindField("PEXSAD")) = premsRow.Value(premsRow.Fields.FindField("PEXSAD"))
                                    expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                                End If

                                expRow.Value(expRow.Fields.FindField("PEXPZIP")) = premsRow.Value(premsRow.Fields.FindField("PEXPZIP"))
                                expRow.Value(expRow.Fields.FindField("IS_IMPERVIOUS_ONLY")) = premsRow.Value(premsRow.Fields.FindField("IS_IMPERVIOUS_ONLY"))
                                expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = premsRow.Value(premsRow.Fields.FindField("IS_EXEMPT_IAB"))
                                expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = premsRow.Value(premsRow.Fields.FindField("EXEMPT_IAB_REASON"))
                                If ia_charge = "" Then
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                                Else
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                                End If

                                expRow.Value(expRow.Fields.FindField("PEXSAD")) = premsRow.Value(premsRow.Fields.FindField("PEXSAD"))

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                            ElseIf transactionCode = "MADDR" Then

                                countMADDR = countMADDR + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MADDR"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                expRow.Value(expRow.Fields.FindField("ADDRESSSPLITFLAG")) = premsRow.Value(premsRow.Fields.FindField("USE_PARSED_ADDRESS"))

                                If GetRowValue(premsRow, "USE_PARSED_ADDRESS") = "Y" Then

                                    expRow.Value(expRow.Fields.FindField("ADDRNUM")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUM"))
                                    If (IsDBNull(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")))) Then
                                        expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF"))
                                    Else
                                        expRow.Value(expRow.Fields.FindField("ADDRNUMSUF")) = Left(premsRow.Value(premsRow.Fields.FindField("ADDRNUMSUF")), 3)
                                    End If
                                    stname = GetUSPSST(GetRowValue(premsRow, "STNAME"), stcrosswalkTable)
                                    expRow.Value(expRow.Fields.FindField("STNAME")) = stname
                                    expRow.Value(expRow.Fields.FindField("STREET_TYPE")) = premsRow.Value(premsRow.Fields.FindField("STREET_TYPE"))
                                    expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                                    expRow.Value(expRow.Fields.FindField("UNIT")) = premsRow.Value(premsRow.Fields.FindField("UNIT"))
                                Else
                                    expRow.Value(expRow.Fields.FindField("PEXSAD")) = premsRow.Value(premsRow.Fields.FindField("PEXSAD"))
                                    expRow.Value(expRow.Fields.FindField("QUADRANT")) = premsRow.Value(premsRow.Fields.FindField("QUADRANT"))
                                End If

                                expRow.Value(expRow.Fields.FindField("PEXPZIP")) = premsRow.Value(premsRow.Fields.FindField("PEXPZIP"))


                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                            ElseIf transactionCode = "MSSL" Then

                                countMSSL = countMSSL + 1
                                '** added transacation code
                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MSSL"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                expRow.Value(expRow.Fields.FindField("PEXSQUARE")) = premsRow.Value(premsRow.Fields.FindField("PEXSQUARE"))
                                expRow.Value(expRow.Fields.FindField("PEXLOT")) = premsRow.Value(premsRow.Fields.FindField("PEXLOT"))
                                expRow.Value(expRow.Fields.FindField("PEXSUFFIX")) = premsRow.Value(premsRow.Fields.FindField("PEXSUFFIX"))
                                expRow.Value(expRow.Fields.FindField("WARD")) = premsRow.Value(premsRow.Fields.FindField("WARD"))
                                If IsDBNull(premsRow.Value(premsRow.Fields.FindField("OTRUSECODE"))) Then
                                    expRow.Value(expRow.Fields.FindField("OTRUSECODE")) = "0"
                                Else
                                    expRow.Value(expRow.Fields.FindField("OTRUSECODE")) = premsRow.Value(premsRow.Fields.FindField("OTRUSECODE"))
                                End If


                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                            ElseIf transactionCode = "MSTAT" Then
                                If premsRow.Value(premsRow.Fields.FindField("PEXPSTS")) = "PG" Then

                                    countMSTAT = countMSTAT + 1

                                    expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MSTAT"
                                    expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                    expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                    expRow.Value(expRow.Fields.FindField("PEXPSTS")) = premsRow.Value(premsRow.Fields.FindField("PEXPSTS"))
                                    expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = premsRow.Value(premsRow.Fields.FindField("IS_EXEMPT_IAB"))
                                    expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = premsRow.Value(premsRow.Fields.FindField("EXEMPT_IAB_REASON"))

                                    expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                    expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                    expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                    expRow.Store()
                                End If
                            ElseIf transactionCode = "MTYPE" Then

                                countMTYPE = countMTYPE + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MTYPE"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                expRow.Value(expRow.Fields.FindField("PEXPTYP")) = premsRow.Value(premsRow.Fields.FindField("PEXPTYP"))

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                            ElseIf transactionCode = "MXMPT" Then

                                countMXMPT = countMXMPT + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MXMPT"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                expRow.Value(expRow.Fields.FindField("IS_IMPERVIOUS_ONLY")) = premsRow.Value(premsRow.Fields.FindField("IS_IMPERVIOUS_ONLY"))
                                expRow.Value(expRow.Fields.FindField("IS_EXEMPT_IAB")) = premsRow.Value(premsRow.Fields.FindField("IS_EXEMPT_IAB"))
                                expRow.Value(expRow.Fields.FindField("EXEMPT_IAB_REASON")) = premsRow.Value(premsRow.Fields.FindField("EXEMPT_IAB_REASON"))

                                If ia_charge = "" Then
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                                Else
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                                End If

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value




                                expRow.Store()
                            ElseIf transactionCode = "MXREF" Then

                                countMXREF = countMXREF + 1

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONCODE")) = "MXREF"
                                expRow.Value(expRow.Fields.FindField("PEXPRM")) = premsRow.Value(premsRow.Fields.FindField("PEXPRM"))
                                expRow.Value(expRow.Fields.FindField("PEXUID")) = premsRow.Value(premsRow.Fields.FindField("PEXUID"))

                                If ia_charge = "" Then
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = System.DBNull.Value
                                Else
                                    expRow.Value(expRow.Fields.FindField("IA_CHARGE")) = ia_charge
                                End If

                                expRow.Value(expRow.Fields.FindField("TRANSACTIONDT")) = transactionDate
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONDT")) = System.DBNull.Value
                                expRow.Value(expRow.Fields.FindField("EXCEPTIONMSGID")) = System.DBNull.Value

                                expRow.Store()
                            End If

                        End If

                        Marshal.ReleaseComObject(pCursor)


                    End If

                    pExpPremsOutRow = pExpPremsOutCursor.NextRow

                Loop


            End If

            countTrans = countMADDR + countMSSL + countMSTAT + countMTYPE + countMXMPT + countMXREF + countMNPRM


            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            MsgBox("Number of transactions created: " & countTrans & vbCrLf &
                   "MNPRM: " & vbTab & countMNPRM & vbCrLf &
                   "MSSL : " & vbTab & countMSSL & vbCrLf &
                   "MSTAT: " & vbTab & countMSTAT & vbCrLf &
                   "MTYPE: " & vbTab & countMTYPE & vbCrLf &
                   "MXMPT: " & vbTab & countMXMPT & vbCrLf &
                   "MXREF: " & vbTab & countMXREF & vbCrLf &
                   "MADDR: " & vbTab & countMADDR & vbCrLf &
                   "IACF: " & vbTab & countIACF)

        Catch ex As Exception
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(False)

            MsgBox(" Error : " & ex.Message & vbCrLf & ex.StackTrace())

            Return 2
        End Try


        Return 0

    End Function

    Private Function GetUSPSST(ByVal stname As String, ByVal pTable As ITable) As String
        Dim pFilter As IQueryFilter = New QueryFilter
        Dim quoteStName = stname.Trim().Replace("'", "''''")
        pFilter.WhereClause = "UPPER(STNAME)='" & UCase(quoteStName) & "'"
        Dim pCursor As ICursor = pTable.Search(pFilter, False)
        Dim pRow As IRow = pCursor.NextRow
        If pRow Is Nothing Then
            Return stname
        Else
            Return GetRowValue(pRow, "UPSST")
        End If
        Marshal.ReleaseComObject(pCursor)
        Marshal.ReleaseComObject(pFilter)
    End Function

    Private Function HasTxType(ByVal txcode As String, ByVal pOldRow As IRow, ByVal pNewRow As IRow, Optional ByVal useParsedAddress As String = "") As Boolean

        Dim flag As Boolean

        Dim i As Integer
        For i = 0 To pOldRow.Fields.FieldCount - 1
            If pNewRow.Value(i).ToString <> pOldRow.Value(i).ToString Then
                If txcode = "MADDR" AndAlso (
                (pNewRow.Fields.Field(i).Name = "PEXSAD" AndAlso useParsedAddress <> "Y" AndAlso Not IsDBNull(pNewRow.Value(i)) AndAlso Not pNewRow.Value(i).ToString = "") Or
                (pNewRow.Fields.Field(i).Name = "ADDRNUM" AndAlso useParsedAddress = "Y") Or
                (pNewRow.Fields.Field(i).Name = "ADDRNUMSUF" AndAlso useParsedAddress = "Y") Or
                (pNewRow.Fields.Field(i).Name = "STNAME" AndAlso useParsedAddress = "Y") Or
                (pNewRow.Fields.Field(i).Name = "STREET_TYPE" AndAlso useParsedAddress = "Y") Or
                (pNewRow.Fields.Field(i).Name = "UNIT" AndAlso useParsedAddress = "Y") Or
                (pNewRow.Fields.Field(i).Name = "QUADRANT") AndAlso useParsedAddress = "Y") Then
                    Return True
                End If

                If txcode = "MSSL" AndAlso (
                    pNewRow.Fields.Field(i).Name = "PEXSQUARE" Or
                    pNewRow.Fields.Field(i).Name = "PEXSUFFIX" Or
                    pNewRow.Fields.Field(i).Name = "PEXLOT" Or
                    pNewRow.Fields.Field(i).Name = "OTRUSECODE" Or
                    pNewRow.Fields.Field(i).Name = "WARD") Then
                    Return True
                End If

                If txcode = "MSTAT" AndAlso pNewRow.Fields.Field(i).Name = "PEXPSTS" Then
                    Return True
                End If

                If txcode = "MTYPE" AndAlso pNewRow.Fields.Field(i).Name = "PEXPTYP" Then
                    Return True
                End If

                If txcode = "MXMPT" AndAlso (
                    pNewRow.Fields.Field(i).Name = "IS_EXEMPT_IAB" Or
                    pNewRow.Fields.Field(i).Name = "IS_IMPERVIOUS_ONLY" Or
                    pNewRow.Fields.Field(i).Name = "EXEMPT_IAB_REASON") Then
                    Return True
                End If

                If txcode = "MXREF" AndAlso pNewRow.Fields.Field(i).Name = "MASTER_PEXUID" Then
                    Return True
                End If

                If txcode = "IACF" AndAlso pNewRow.Fields.Field(i).Name = "IS_EXEMPT_IAB" Then
                    If pNewRow.Value(i).ToString = "Y" And pOldRow.Value(i).ToString = "N" Then
                        Return True
                    End If
                End If
            End If
        Next

    End Function

    Public Function InvokeEditor(ByVal hWndParent As Integer, ByVal argsIn() As Object) As Object() Implements ESRI.ArcGIS.JTX.IJTXCustomStep.InvokeEditor
        Return Nothing
    End Function

    Public Sub OnCreate(ByVal ipDatabase As ESRI.ArcGIS.JTX.IJTXDatabase) Implements ESRI.ArcGIS.JTX.IJTXCustomStep.OnCreate
        m_ipDatabase = ipDatabase
    End Sub

    Public Function ValidateArguments(ByRef argv() As Object) As Boolean Implements ESRI.ArcGIS.JTX.IJTXCustomStep.ValidateArguments
        'If argv Is Nothing Then
        '    m_checkCISException = False
        'End If

        'If argv.Length = 0 Then
        '    m_checkCISException = False
        'End If

        'If argv.Length > 0 Then
        '    If argv(0) = "Y" Then
        '        m_checkCISException = True
        '    ElseIf argv(0) = "N" Then
        '        m_checkCISException = False
        '    Else
        '        Return False
        '    End If
        'End If

        Return True

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
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

    End Function

    Private Function GetRowValue(ByVal pRow As IRow, ByVal col As String, Optional ByVal getDomainValue As Boolean = False) As String
        Dim fIndex As Integer

        If pRow Is Nothing Then
            Return ""
        End If

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
End Class


