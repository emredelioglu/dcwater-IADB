Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.JTX
Imports ESRI.ArcGIS.Geodatabase
Imports System.Xml
Imports System.IO
Imports System.Data.OracleClient
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Module Module1

    Private log As System.IO.StreamWriter
    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()

    Private m_jtxJob As IJTXJob = Nothing
    Private m_exitCode As Integer

    <STAThread()> _
    Sub Main()

        If Not File.Exists(Environment.CurrentDirectory & "\ExportCFUpdateToCIS.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\ExportCFUpdateToCIS.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\ExportCFUpdateToCIS.log")
        End If

        Try
            log.WriteLine(Now & " : Update CISOUTPUTBILLDET routine started")

            'ESRI License Initializer generated code.
            m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
            New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})
            'ESRI License Initializer generated code.
            'Do not make any call to ArcObjects after ShutDownApplication()


            Dim sParams() As String = Environment.GetCommandLineArgs
            Dim transDate As String = ""

            If sParams.Length > 1 Then
                transDate = sParams(1)
                log.WriteLine("Parameter transDate:" & transDate)
            Else
                log.WriteLine("Parameter transDate: <none> ")
                Console.WriteLine("Usage for ExportCFUpdateToCIS: ExportCFUpdateToCIS [TRANSACTION_DATE MM/DD/YYYY]")
                Throw New Exception("Wrong number of arguments")
            End If

            If Not IsDate(transDate) Then
                Console.WriteLine("Usage for ExportCFUpdateToCIS: ExportCFUpdateToCIS [TRANSACTION_DATE MM/DD/YYYY]")
                Throw New Exception("Wrong arguments : " & sParams(1))
            End If


            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode

            Dim jtxDB As String
            Dim jtxDBWorkspace As String
            Dim jtxJobParentVersion As String

            Dim tableCISOUTPUTBILLDET As String = ""
            Dim tableIACF As String = ""
            Dim tablePREMSINTERPT As String = ""

            m_xmld = New XmlDocument()
            m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

            m_nodelist = m_xmld.SelectNodes("IAISConfig/ExportCFUpdateToCIS")

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
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTBILLDET" Then
                        tableCISOUTPUTBILLDET = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_IACF" Then
                        tableIACF = m_node.ChildNodes(i).InnerText
                    ElseIf m_node.ChildNodes(i).Name = "TABLE_PREMSINTERPT" Then
                        tablePREMSINTERPT = m_node.ChildNodes(i).InnerText
                    End If
                Next i
            End If

            Dim rowcount As Integer = 0

            Dim pFWS As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, jtxJobParentVersion)
            Dim pWorkspaceEdit As IWorkspaceEdit = pFWS

            pWorkspaceEdit.StartEditing(False)
            pWorkspaceEdit.StartEditOperation()

            Try
                'Open Exception tables and remove all corrected records
                Dim pTableIACF As ITable = pFWS.OpenTable(tableIACF)
                Dim pTableCISOUTPUTBILLDET As ITable = pFWS.OpenTable(tableCISOUTPUTBILLDET)
                Dim pTablePREMSINTERPT As ITable = pFWS.OpenTable(tablePREMSINTERPT)

                Dim pCursor As ICursor
                Dim pRow As IRow
                Dim pFilter As IQueryFilter = New QueryFilter
                pFilter.WhereClause = "LASTEXPDT IS NULL AND EFFSTARTDT <= to_date('" & transDate & "', 'mm/dd/yyyy') AND EFFENDDT IS NULL"

                pCursor = pTableIACF.Search(pFilter, False)
                pRow = pCursor.NextRow

                Do While Not pRow Is Nothing
                    pRow.Value(pRow.Fields.FindField("LASTEXPDT")) = Now
                    pRow.Store()

                    Dim pExpRow As IRow = pTableCISOUTPUTBILLDET.CreateRow

                    pFilter.WhereClause = "PEXUID=" & pRow.Value(pRow.Fields.FindField("PEXUID"))
                    Dim premsCursor As ICursor = pTablePREMSINTERPT.Search(pFilter, False)
                    Dim premsPt As IRow = premsCursor.NextRow

                    pExpRow.Value(pExpRow.Fields.FindField("PEXPRM")) = premsPt.Value(premsPt.Fields.FindField("PEXPRM"))

                    Marshal.releaseComobject(premsCursor)

                    pExpRow.Value(pExpRow.Fields.FindField("PEXUID")) = pRow.Value(pRow.Fields.FindField("PEXUID"))
                    pExpRow.Value(pExpRow.Fields.FindField("IMPERV_BILLED_ERU")) = pRow.Value(pRow.Fields.FindField("IABILLERU"))
                    pExpRow.Value(pExpRow.Fields.FindField("IMPERV_CREDIT_ERU")) = pRow.Value(pRow.Fields.FindField("IACREDERU"))
                    pExpRow.Value(pExpRow.Fields.FindField("IMPERV_AREA_SQFT")) = pRow.Value(pRow.Fields.FindField("IASQFT"))
                    pExpRow.Value(pExpRow.Fields.FindField("TRANSACTIONDT")) = CDate(transDate).AddMinutes(10)

                    pExpRow.Store()

                    rowcount = rowcount + 1

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
            End Try


            log.WriteLine(Now & " : " & rowcount & " rows added")
            log.WriteLine(Now & " : Update CISOUTPUTBILLDET routine completed")
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
