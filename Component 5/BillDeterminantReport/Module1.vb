Imports ESRI.ArcGIS.esriSystem
Imports System.IO
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Output
Imports stdole
Imports System.Xml
Imports ESRI.ArcGIS.JTX

Module Module1
    Private m_AOLicenseInitializer As LicenseInitializer = New LicenseInitializer()
    Private log As System.IO.StreamWriter

    Private m_exitCode As Integer
    Private pexuidFileName As String

#Region "Main entrance"

    <STAThread()> _
    Sub Main()
        m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor, esriLicenseProductCode.esriLicenseProductCodeArcInfo}, _
        New esriLicenseExtensionCode() {esriLicenseExtensionCode.esriLicenseExtensionCodeJTX})

        If Not File.Exists(Environment.CurrentDirectory & "\BillDeterminantReports.log") Then
            log = File.CreateText(Environment.CurrentDirectory & "\BillDeterminantReports.log")
        Else
            log = File.AppendText(Environment.CurrentDirectory & "\BillDeterminantReports.log")
        End If

        Try
            Dim sParams() As String = Environment.GetCommandLineArgs
            If sParams.Length < 4 Then
                Console.WriteLine("Usage for BDRGenerator: ")
                Console.WriteLine("BDRGenerator CreatePremiseList [TRANSACTION_DATE] [FILE_NAME]")
                Console.WriteLine("BDRGenerator CreatePDF [MXD_PATH] [PDF_PATH] [FILE_NAME] <CONTINUE_ON_ERROR Y|N >")
                Throw New Exception("Wrong number of arguments")
            End If

            log.WriteLine("Parameter routineName:" & sParams(1))

            If sParams(1) = "CreatePremiseList" Then
                If sParams.Length < 4 Then
                    Console.WriteLine("Usage for BDRGenerator: BDRGenerator CreatePremiseList [TRANSACTION_DATE] [FILE_NAME]")
                    Throw New Exception("Wrong number of arguments")
                Else
                    pexuidFileName = sParams(3)
                End If

                GeneratePremiseList(sParams(2))
            ElseIf sParams(1) = "CreatePDF" Then
                If sParams.Length < 5 Then
                    Console.WriteLine("Usage for BDRGenerator: BDRGenerator CreatePDF [MXD_PATH] [PDF_PATH] [FILE_NAME] <CONTINUE_ON_ERROR Y|N >")
                    Throw New Exception("Wrong number of arguments")
                Else
                    pexuidFileName = sParams(4)
                End If

                Dim bContunieOnError As Boolean = False
                If sParams.Length > 5 Then
                    If sParams(5) = "Y" Then
                        bContunieOnError = True
                    End If
                End If

                GeneratePDFs(sParams(2), sParams(3), bContunieOnError)

            Else
                Console.WriteLine("Usage for BDRGenerator: ")
                Console.WriteLine("BDRGenerator CreatePremiseList [TRANSACTION_DATE] [FILE_NAME]")
                Console.WriteLine("BDRGenerator CreatePDF [MXD_PATH] [PDF_PATH] [FILE_NAME] <CONTINUE_ON_ERROR Y|N >")
                Throw New Exception("Wrong arguments")
            End If

            m_exitCode = 0
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            Console.WriteLine("Error:" & ex.Message)
            m_exitCode = 1
        End Try

        log.Close()
        m_AOLicenseInitializer.ShutdownApplication()
        Environment.Exit(m_exitCode)
    End Sub
#End Region

    Public Sub GeneratePremiseList(ByVal transdate As String)
        log.WriteLine(Now & " : Generating premise list with following parameters")
        log.WriteLine("transdate : " & transdate)
        log.WriteLine("fileName : " & pexuidFileName)

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim jtxDB As String
        Dim jtxDBWorkspace As String

        Dim tblCISOUTPUTBILLDET As String
        Dim versionName As String

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")

        m_nodelist = m_xmld.SelectNodes("IAISConfig/BillDeterminantReport")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "JTX_DB_NAME" Then
                    jtxDB = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "JTX_WORKSPACE" Then
                    jtxDBWorkspace = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_CISOUTPUTBILLDET" Then
                    tblCISOUTPUTBILLDET = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "VERSION" Then
                    versionName = m_node.ChildNodes(i).InnerText
                End If
            Next i
        End If

        Dim pJTXDBMan As IJTXDatabaseManager = New JTXDatabaseManager
        Dim pJTXDatabase As IJTXDatabase = pJTXDBMan.GetDatabase(jtxDB)
        Dim pSdeWorkspace As IFeatureWorkspace = mdlJTX.GetFeatureWorkspace(jtxDB, jtxDBWorkspace, versionName)

        Dim pQueryDef As IQueryDef = pSdeWorkspace.CreateQueryDef
        pQueryDef.Tables = tblCISOUTPUTBILLDET
        pQueryDef.SubFields = "COUNT(*)"
        pQueryDef.WhereClause = "TRANSACTIONDT BETWEEN To_Date(' " & transdate _
                                                & "', 'mm/dd/yyyy') AND (To_Date('" & transdate & "', 'mm/dd/yyyy') + 1)"

        Dim pCursor As ICursor = pQueryDef.Evaluate
        Dim pRow As IRow = pCursor.NextRow

        Dim oWrite As StreamWriter
        oWrite = File.CreateText(pexuidFileName)

        Try
            If pRow.Value(0) = 0 Then
                log.WriteLine(Now & ": No record in CISOUTPUTBILLDET need to be processed")
            Else
                pQueryDef.SubFields = "*"
                pCursor = pQueryDef.Evaluate
                pRow = pCursor.NextRow
                Do While Not pRow Is Nothing
                    oWrite.WriteLine(MapUtil.GetRowValue(pRow, "PEXUID"))
                    pRow = pCursor.NextRow
                Loop

            End If
        Catch ex As Exception
            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        Finally
            oWrite.Close()
        End Try

        log.WriteLine(Now & " : Completed generating premise list")
    End Sub

    Public Sub GeneratePDFs(ByVal mxdpath As String, ByVal pdfpath As String, ByVal bContunieOnError As Boolean)

        log.WriteLine(Now & " : Generating pdf with following parameters")
        log.WriteLine("mxdpath : " & mxdpath)
        log.WriteLine("pdfpath : " & pdfpath)
        log.WriteLine("fileName : " & pexuidFileName)
        log.WriteLine("bContunieOnError : " & bContunieOnError)

        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode

        Dim tblIACF As String
        Dim tblPremisePt As String
        Dim tblOwnerPly As String
        Dim tblOwnerGapPly As String
        Dim tblRevOwnerPly As String
        Dim tblAppealOwnerPly As String


        Dim tblIAAssignPly As String
        Dim tblRevIAAssignPly As String
        Dim tblAppealIAAssignPly As String
        Dim showResInfo As Boolean = False

        m_xmld = New XmlDocument()
        m_xmld.Load(Application.StartupPath & "\IAISConfig.xml")


        m_nodelist = m_xmld.SelectNodes("IAISConfig/BillDeterminantReport")

        If Not m_nodelist Is Nothing AndAlso m_nodelist.Count > 0 Then
            m_node = m_nodelist.Item(0)
            For i = 0 To m_node.ChildNodes.Count - 1
                If m_node.ChildNodes(i).Name = "TABLE_IACF" Then
                    tblIACF = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_PRMSINTERPT" Then
                    tblPremisePt = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_OWNERPLY" Then
                    tblOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_OWNERGAPPLY" Then
                    tblOwnerGapPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_REVOWNERPLY" Then
                    tblRevOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_APPEALOWNERPLY" Then
                    tblAppealOwnerPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_IAASSIGNPLY" Then
                    tblIAAssignPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_REVIAASSIGNPLY" Then
                    tblRevIAAssignPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "TABLE_APPEALIAASSIGNPLY" Then
                    tblAppealIAAssignPly = m_node.ChildNodes(i).InnerText
                ElseIf m_node.ChildNodes(i).Name = "SHOWRESINFO" Then
                    showResInfo = Boolean.Parse(m_node.ChildNodes(i).InnerText)
                End If
            Next i
        End If

        Dim mapdoc As IMapDocument = Nothing

        Dim oread As System.IO.StreamReader

        If Not File.Exists(pexuidFileName) Then
            Throw New Exception("Can't find file " & pexuidFileName)
        End If
        oread = File.OpenText(pexuidFileName)

        Try
            mapdoc = New MapDocumentClass
            mapdoc.Open(mxdpath)

            Dim pActiveView As IActiveView
            Dim pMap As IMap = mapdoc.Map(0)

            Dim appHwnd As Long = GetHWndForCurrentApplication()

            Dim i As Integer
            For i = 0 To mapdoc.MapCount - 1
                pActiveView = mapdoc.Map(i)
                pActiveView.Activate(appHwnd)
            Next

            pActiveView = mapdoc.PageLayout
            pActiveView.Activate(appHwnd)

            Dim pGC As IGraphicsContainer = mapdoc.PageLayout
            pGC.Reset()


            Dim pLayer As IFeatureLayer = MapUtil.GetLayerByTableName(tblOwnerPly, pMap)
            Dim pFeature As IFeature
            Dim pFeatCursor As IFeatureCursor
            Dim pFilter As IQueryFilter = New QueryFilter

            Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByTableName(tblPremisePt, pMap)

            Dim pDataset As IDataset
            pDataset = premiseLayer

            Dim pIacfTable As ITable
            pIacfTable = MapUtil.GetTableFromWS(pDataset.Workspace, tblIACF)

            'Read pexuid from the text file
            Dim pexuid As String

            Dim square, suffix, lot As String

            Dim rc As RotatingCalipers = New RotatingCalipers
            Dim premiseInfo As clsPremiseInfo = New clsPremiseInfo

            Dim uc As IUnitConverter = New UnitConverter
            Dim minimumWidth = uc.ConvertUnits(500, esriUnits.esriFeet, esriUnits.esriMeters)
            Dim minimumHeight = uc.ConvertUnits(233, esriUnits.esriFeet, esriUnits.esriMeters)
            Dim eruWidth = uc.ConvertUnits(Math.Sqrt(1000), esriUnits.esriFeet, esriUnits.esriMeters)

            Dim _pdfCount As Integer = 0

            log.WriteLine("Start time: " & Now)

            Dim pGeomBag As IGeometryCollection = New GeometryBag
            Dim ownerPly As IFeature

            Dim pCursor As ICursor
            Dim pRow As IRow

            While oread.Peek <> -1
                pexuid = oread.ReadLine()

                Try
                    pGeomBag.RemoveGeometries(0, pGeomBag.GeometryCount)
                    pFilter.WhereClause = "PEXUID =" & pexuid & " AND EFFENDDT IS NULL"

                    pCursor = pIacfTable.Search(pFilter, False)
                    pRow = pCursor.NextRow

                    If Not pRow Is Nothing Then
                        premiseInfo.Reset()

                        Dim pFLDef As IFeatureLayerDefinition
                        pFLDef = premiseLayer
                        pFLDef.DefinitionExpression = "PEXUID=" & pexuid

                        pFeatCursor = premiseLayer.Search(Nothing, False)
                        pFeature = pFeatCursor.NextFeature

                        premiseInfo.IAERU = MapUtil.GetRowValue(pRow, "IABILLERU")
                        premiseInfo.IAArea = Math.Floor(CDbl(MapUtil.GetRowValue(pRow, "IASQFT")))
                        premiseInfo.EffectiveDate = MapUtil.GetRowValue(pRow, "EFFSTARTDT")

                        'Get the IA and Parcel layer information from the IACF record
                        Dim iaLayerParaName As String = MapUtil.GetRowValue(pRow, "IA_SOURCE")
                        Dim pLayerParaName As String = MapUtil.GetRowValue(pRow, "PARCEL_SOURCE")
                        If (iaLayerParaName = "" Or pLayerParaName = "") Then
                            log.WriteLine("IA_SOURCE or PARCEL_SOURCE not set in IACF")
                        End If

                        'find the correct SSL based on spatial query
                        Dim pSFilter As ISpatialFilter = New SpatialFilter
                        pSFilter.Geometry = pFeature.Shape
                        pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin

                        Dim OwnerPlyType As String = ""
                        ownerPly = Nothing

                        pLayer = MapUtil.GetLayerByTableName(tblAppealOwnerPly, pMap)
                        Dim pFSel As IFeatureSelection
                        pFSel = pLayer
                        pFSel.Clear()
                        'pFilter.WhereClause = sslquery
                        If UCase(pLayerParaName) = UCase(tblAppealOwnerPly) Then
                            pFeatCursor = pLayer.Search(pSFilter, False)
                            ownerPly = pFeatCursor.NextFeature
                            If Not ownerPly Is Nothing Then
                                OwnerPlyType = "A"
                                pFSel.Add(ownerPly)
                            End If
                        End If

                        pLayer = MapUtil.GetLayerByTableName(tblRevOwnerPly, pMap)
                        pFSel = pLayer
                        pFSel.Clear()
                        If UCase(pLayerParaName) = UCase(tblRevOwnerPly) Then
                            If ownerPly Is Nothing Then
                                pFeatCursor = pLayer.Search(pSFilter, False)
                                ownerPly = pFeatCursor.NextFeature
                                If Not ownerPly Is Nothing Then
                                    OwnerPlyType = "R"
                                    pFSel.Add(ownerPly)
                                End If
                            End If
                        End If

                        pLayer = MapUtil.GetLayerByTableName(tblOwnerPly, pMap)
                        pFSel = pLayer
                        pFSel.Clear()
                        If UCase(pLayerParaName) = UCase(tblOwnerPly) Then
                            If ownerPly Is Nothing Then
                                pFeatCursor = pLayer.Search(pSFilter, False)
                                ownerPly = pFeatCursor.NextFeature
                                If Not ownerPly Is Nothing Then
                                    OwnerPlyType = "B"
                                    pFSel.Add(ownerPly)
                                End If
                            End If
                        End If

                        pLayer = MapUtil.GetLayerByTableName(tblOwnerGapPly, pMap)
                        pFSel = pLayer
                        pFSel.Clear()
                        If UCase(pLayerParaName) = UCase(tblOwnerGapPly) Then
                            If ownerPly Is Nothing Then
                                pFeatCursor = pLayer.Search(pSFilter, False)
                                ownerPly = pFeatCursor.NextFeature
                                If Not ownerPly Is Nothing Then
                                    OwnerPlyType = "G"
                                    pFSel.Add(ownerPly)
                                End If
                            End If
                        End If


                        If Not ownerPly Is Nothing Then

                            Dim sslquery As String = MapUtil.GetValue(ownerPly, "SSL")

                            premiseInfo.PremiseNumber = MapUtil.GetValue(pFeature, "PEXPRM")

                            square = MapUtil.GetValue(pFeature, "PEXSQUARE")
                            suffix = MapUtil.GetValue(pFeature, "PEXSUFFIX")
                            lot = MapUtil.GetValue(pFeature, "PEXLOT")

                            premiseInfo.SSL = MapUtil.GetSSL(square, suffix, lot)


                            premiseInfo.ServiceAddress1 = GetSAD(pFeature)
                            premiseInfo.ServiceAddress2 = "Washington, DC " & MapUtil.getFormatedZip(MapUtil.GetValue(pFeature, "PEXPZIP"))

                            premiseInfo.Pexptyp = MapUtil.GetValue(pFeature, "PEXPTYP")
                            premiseInfo.DirectAssign = MapUtil.GetValue(pFeature, "HAS_DIRECT_IAASSIGN")

                            pFLDef = premiseLayer
                            pFLDef.DefinitionExpression = "PEXUID=" & pexuid & " OR MASTER_PEXUID=" & pexuid

                            pFeatCursor = premiseLayer.Search(Nothing, False)
                            pFeature = pFeatCursor.NextFeature
                            Do While Not pFeature Is Nothing
                                If MapUtil.GetValue(pFeature, "MASTER_PEXUID") <> "" Then
                                    Dim pt As clsPremisePt = New clsPremisePt
                                    pt.pexnum = MapUtil.GetValue(pFeature, "PEXPRM")
                                    pt.pexptyp = MapUtil.GetValue(pFeature, "PEXPTYP")
                                    pt.pexsad = GetSAD(pFeature)
                                    pt.exemptreason = MapUtil.GetValue(pFeature, "EXEMPT_IAB_REASON", True)

                                    premiseInfo.addPremisePt(pt)
                                End If

                                pFeature = pFeatCursor.NextFeature
                            Loop

                            Dim pEnv As IEnvelope = ownerPly.Shape.Envelope
                            Dim rotationAngle As Double = 0

                            pGeomBag.AddGeometry(ownerPly.ShapeCopy)

                            Dim definitionSQL As String = ""
                            If premiseInfo.DirectAssign = "Y" Then
                                definitionSQL = "(SSL='" & MapUtil.GetSSL(square, suffix, lot) & "' AND PEXUID IS NULL) OR PEXUID=" & pexuid
                            Else
                                definitionSQL = "SSL='" & MapUtil.GetSSL(square, suffix, lot) & "' AND PEXUID IS NULL"
                            End If

                            Dim pIALayer As IFeatureLayer = MapUtil.GetLayerByTableName(tblIAAssignPly, pMap)
                            'If OwnerPlyType = "G" Or OwnerPlyType = "B" Then
                            If UCase(tblIAAssignPly) = UCase(iaLayerParaName) Then
                                If Not pIALayer.Visible Then
                                    pIALayer.Visible = True
                                End If

                                pFLDef = pIALayer
                                pFLDef.DefinitionExpression = definitionSQL
                                pFeatCursor = pIALayer.Search(Nothing, False)
                            Else
                                If pIALayer.Visible Then
                                    pIALayer.Visible = False
                                End If
                            End If

                            pIALayer = MapUtil.GetLayerByTableName(tblRevIAAssignPly, pMap)
                            'If OwnerPlyType = "R" Then
                            If UCase(tblRevIAAssignPly) = UCase(iaLayerParaName) Then
                                If Not pIALayer.Visible Then
                                    pIALayer.Visible = True
                                End If

                                pFLDef = pIALayer
                                pFLDef.DefinitionExpression = definitionSQL
                                pFeatCursor = pIALayer.Search(Nothing, False)
                            Else
                                If pIALayer.Visible Then
                                    pIALayer.Visible = False
                                End If
                            End If

                            pIALayer = MapUtil.GetLayerByTableName(tblAppealIAAssignPly, pMap)
                            'If OwnerPlyType = "A" Then
                            If UCase(tblAppealIAAssignPly) = UCase(iaLayerParaName) Then
                                If Not pIALayer.Visible Then
                                    pIALayer.Visible = True
                                End If

                                pFLDef = pIALayer
                                pFLDef.DefinitionExpression = definitionSQL
                                pFeatCursor = pIALayer.Search(Nothing, False)
                            Else
                                If pIALayer.Visible Then
                                    pIALayer.Visible = False
                                End If
                            End If

                            'The cursor has been returned from the code above
                            pFeature = pFeatCursor.NextFeature
                            Dim pRelOp As IRelationalOperator = ownerPly.Shape
                            Do While Not pFeature Is Nothing
                                If Not pRelOp.Contains(pFeature.ShapeCopy) Then
                                    pGeomBag.AddGeometry(pFeature.ShapeCopy)
                                End If

                                If pFeature.Fields.FindField("FEATURETYPE") > 0 Then
                                    premiseInfo.SetIAArea(pFeature.Value(pFeature.Fields.FindField("IARSQF")), pFeature.Value(pFeature.Fields.FindField("FEATURETYPE")))
                                Else
                                    premiseInfo.SetIAArea(pFeature.Value(pFeature.Fields.FindField("IARSQF")), pFeature.Value(pFeature.Fields.FindField("IATYPE")))
                                End If

                                pFeature = pFeatCursor.NextFeature
                            Loop

                            Dim mer As IPolygon = rc.GetMER(pGeomBag)
                            Dim mertrans As ITransform2D = mer

                            Dim pSegCol As ISegmentCollection = mer
                            Dim pLine As ILine = pSegCol.Segment(0)

                            pEnv = mer.Envelope

                            If pLine.Length < pSegCol.Segment(1).Length Then
                                pEnv.Height = pLine.Length
                                pLine = pSegCol.Segment(1)
                                pEnv.Width = pSegCol.Segment(1).Length
                            Else
                                pEnv.Width = pLine.Length
                                pEnv.Height = pSegCol.Segment(1).Length
                            End If
                            Dim pArea As IArea = mer
                            pEnv.CenterAt(pArea.Centroid)


                            rotationAngle = Math.Abs((180 * pLine.Angle) / (4 * Math.Atan(1)))
                            If rotationAngle > 270 Then
                                rotationAngle = rotationAngle - 360
                            ElseIf rotationAngle > 90 Then
                                rotationAngle = rotationAngle - 180
                            End If

                            pActiveView = pMap

                            pActiveView.Extent = pEnv
                            pEnv = pActiveView.Extent

                            If pEnv.Width < minimumWidth And pEnv.Height < minimumHeight Then
                                Dim ratio As Double = minimumWidth / pEnv.Width
                                If ratio > minimumHeight / pEnv.Height Then
                                    ratio = minimumHeight / pEnv.Height
                                End If
                                pEnv.Expand(ratio, ratio, True)
                                pActiveView.Extent = pEnv
                            Else
                                pEnv.Expand(1.05, 1.05, True)
                            End If

                            pActiveView = pMap
                            pActiveView.Extent = pEnv
                            pActiveView.ScreenDisplay.DisplayTransformation.Rotation = rotationAngle
                            pActiveView.ScreenDisplay.DisplayTransformation.VisibleBounds = pEnv

                            '************************
                            pGC = mapdoc.PageLayout
                            pGC.Reset()

                            Dim pEl As IElement = pGC.Next
                            Dim tblX, tblY As Double
                            Do While Not pEl Is Nothing

                                Dim pEP As IElementProperties = pEl
                                'If pEP.Name = "eruSquare" Then

                                '    Dim pERUEnv As IEnvelope = pEl.Geometry.Envelope
                                '    Dim eruLength As Double = eruWidth / (pEnv.Width / 7.5)
                                '    pERUEnv.PutCoords(pERUEnv.XMin, 5.8 - eruLength / 2, pERUEnv.XMin + eruLength, 5.8 + eruLength / 2)
                                '    pEl.Geometry = PolygonFromEnvelope(pERUEnv)
                                'Else
                                If pEP.Name = "ScaleBar" Then
                                    'Dim pGraphicsComposite As IGraphicsComposite = pEP
                                    'Dim pScaleBar As IMapSurround = pGraphicsComposite

                                    Dim pERUEnv As IEnvelope = pEl.Geometry.Envelope
                                    pERUEnv.PutCoords(3, 5.6, 5.5, 5.86)
                                    pEl.Geometry = PolygonFromEnvelope(pERUEnv)
                                End If

                                If TypeOf pEl Is ITextElement Then

                                    Dim pTE As ITextElement = pEl

                                    If pEP.Name = "PEXADDR1" Then
                                        pTE.Text = premiseInfo.ServiceAddress1
                                    ElseIf pEP.Name = "PEXADDR2" Then
                                        pTE.Text = premiseInfo.ServiceAddress2
                                    ElseIf pEP.Name = "PremiseNumber" Then
                                        pTE.Text = premiseInfo.PremiseNumber
                                    ElseIf pEP.Name = "SSL" Then
                                        pTE.Text = premiseInfo.SSL
                                    ElseIf pEP.Name = "IAHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = " "
                                        Else
                                            pTE.Text = "Total Impervious Area ="
                                        End If
                                    ElseIf pEP.Name = "IAArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = " "
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.IAArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "ERU" Then
                                        pTE.Text = FormatNumber(premiseInfo.IAERU, 1) & " ERU"
                                    ElseIf pEP.Name = "EffectiveDate" Then
                                        pTE.Text = FormatDateTime(premiseInfo.EffectiveDate, DateFormat.ShortDate)
                                    ElseIf pEP.Name = "bldgArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.BldgArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "bldgHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = "Building"
                                        Else
                                            pTE.Text = "Building ="
                                        End If
                                    ElseIf pEP.Name = "roadArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.RoadArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "roadHeader" And Not showResInfo Then
                                        If premiseInfo.Pexptyp = "RES" Then
                                            pTE.Text = "Road/Drive/Parking Lot"
                                        Else
                                            pTE.Text = "Road/Drive/Parking Lot ="
                                        End If
                                    ElseIf pEP.Name = "sidewalkArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.SidewalkArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "sidewalkHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = "Sidewalk"
                                        Else
                                            pTE.Text = "Sidewalk ="
                                        End If
                                    ElseIf pEP.Name = "stairArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.StairsArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "stairsHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = "Stairs"
                                        Else
                                            pTE.Text = "Stairs ="
                                        End If
                                    ElseIf pEP.Name = "swmpoolArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.SwmpoolArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "swmpoolHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = "Swimming Pool"
                                        Else
                                            pTE.Text = "Swimming Pool ="
                                        End If
                                    ElseIf pEP.Name = "outdoorArea" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = ""
                                        Else
                                            pTE.Text = FormatNumber(premiseInfo.OutdoorArea, 0) & " sq. ft"
                                        End If
                                    ElseIf pEP.Name = "outdoorHeader" Then
                                        If premiseInfo.Pexptyp = "RES" And Not showResInfo Then
                                            pTE.Text = "Outdoor Rec Facility"
                                        Else
                                            pTE.Text = "Outdoor Rec Facility ="
                                        End If
                                    ElseIf pEP.Name = "prmsListHeader" Then
                                        tblY = pEl.Geometry.Envelope.YMin - 0.12
                                        tblX = pEl.Geometry.Envelope.XMin + 0.07
                                    End If
                                ElseIf TypeOf pEl Is ILineElement Then
                                    Dim pLE As ILineElement = pEl
                                    If pEP.Name = "Line1" Then
                                        Debug.Print(pEl.Geometry.GeometryType)
                                    End If

                                    Dim pPL As IPolyline = pEl.Geometry

                                End If


                                pEl = pGC.Next
                            Loop


                            Dim pList As List(Of clsPremisePt) = premiseInfo.ListOtherPremises

                            If Not pList Is Nothing AndAlso pList.Count > 0 Then
                                Dim lPt As IPoint
                                Dim cellHeight As Double = 0.15
                                Dim tableYOffset As Double = 0.12

                                If pList.Count <= 25 Then
                                    tblX = tblX + 1.93
                                End If

                                lPt = New Point
                                lPt.X = tblX
                                lPt.Y = tblY
                                pEl = MakeATextElement("ptlistPRM", lPt, "Premise", 8, False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 0.47
                                lPt.Y = tblY
                                pEl = MakeATextElement("ptlistPTYP", lPt, "Type", 8, False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 0.9
                                lPt.Y = tblY
                                pEl = MakeATextElement("ptlistSAD", lPt, "Service Address", 8, False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 2.6
                                lPt.Y = tblY
                                pEl = MakeATextElement("ptlistER", lPt, "Exempt Reason", 8, False)
                                pGC.AddElement(pEl, 0)

                                If pList.Count > 25 Then
                                    lPt = New Point
                                    lPt.X = tblX + 4
                                    lPt.Y = tblY
                                    pEl = MakeATextElement("ptlistPRM", lPt, "Premise", 8, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 4.47
                                    lPt.Y = tblY
                                    pEl = MakeATextElement("ptlistPTYP", lPt, "Type", 8, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 4.9
                                    lPt.Y = tblY
                                    pEl = MakeATextElement("ptlistSAD", lPt, "Service Address", 8, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 6.6
                                    lPt.Y = tblY
                                    pEl = MakeATextElement("ptlistER", lPt, "Exempt Reason", 8, False)
                                    pGC.AddElement(pEl, 0)
                                End If

                                For i = 0 To pList.Count - 1
                                    Dim lx, ly As Double

                                    If i < 25 Then
                                        lx = tblX
                                        ly = tblY - (i + 1) * cellHeight
                                    Else
                                        lx = tblX + 4
                                        ly = tblY - (i - 24) * cellHeight
                                    End If

                                    Dim pt As clsPremisePt = pList.Item(i)
                                    lPt = New Point
                                    lPt.X = lx
                                    lPt.Y = ly
                                    pEl = MakeATextElement("ptlistPRM" & i, lPt, pt.pexnum, 7, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = lx + 0.47
                                    lPt.Y = ly
                                    pEl = MakeATextElement("ptlistPTYP" & i, lPt, pt.pexptyp, 7, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = lx + 0.9
                                    lPt.Y = ly
                                    pEl = MakeATextElement("ptlistSAD" & i, lPt, pt.pexsad, 7, False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = lx + 2.6
                                    lPt.Y = ly
                                    pEl = MakeATextElement("ptlistER" & i, lPt, pt.exemptreason, 7, False)
                                    pGC.AddElement(pEl, 0)

                                Next

                                Dim lineCount As Integer
                                If pList.Count > 25 Then
                                    lineCount = pList.Count + 3
                                Else
                                    lineCount = pList.Count + 1
                                End If

                                For i = 0 To lineCount
                                    lPt = New Point
                                    If i <= 26 Then
                                        lPt.X = tblX - 0.07
                                        lPt.Y = tblY + tableYOffset - i * cellHeight
                                    Else
                                        lPt.X = tblX + 3.93
                                        lPt.Y = tblY + tableYOffset - (i - 27) * cellHeight
                                    End If

                                    pEl = MakeALineElement("ptlistHLine" & i, lPt, 3.5)
                                    pGC.AddElement(pEl, 0)
                                Next


                                If pList.Count <= 25 Then
                                    lineCount = pList.Count
                                Else
                                    lineCount = 25
                                End If

                                lPt = New Point
                                lPt.X = tblX - 0.07
                                lPt.Y = tblY + tableYOffset
                                pEl = MakeALineElement("ptlistVLine1", lPt, cellHeight * (lineCount + 1), False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 0.43
                                lPt.Y = tblY + tableYOffset
                                pEl = MakeALineElement("ptlistVLine2", lPt, cellHeight * (lineCount + 1), False)
                                pGC.AddElement(pEl, 0)


                                lPt = New Point
                                lPt.X = tblX + 0.88
                                lPt.Y = tblY + tableYOffset
                                pEl = MakeALineElement("ptlistVLine3", lPt, cellHeight * (lineCount + 1), False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 2.53
                                lPt.Y = tblY + tableYOffset
                                pEl = MakeALineElement("ptlistVLine4", lPt, cellHeight * (lineCount + 1), False)
                                pGC.AddElement(pEl, 0)

                                lPt = New Point
                                lPt.X = tblX + 3.43
                                lPt.Y = tblY + tableYOffset
                                pEl = MakeALineElement("ptlistVLine5", lPt, cellHeight * (lineCount + 1), False)
                                pGC.AddElement(pEl, 0)

                                If pList.Count > 25 Then
                                    lineCount = pList.Count - 25

                                    lPt = New Point
                                    lPt.X = tblX + 3.93
                                    lPt.Y = tblY + tableYOffset
                                    pEl = MakeALineElement("ptlistVLine1", lPt, cellHeight * (lineCount + 1), False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 4.43
                                    lPt.Y = tblY + tableYOffset
                                    pEl = MakeALineElement("ptlistVLine2", lPt, cellHeight * (lineCount + 1), False)
                                    pGC.AddElement(pEl, 0)


                                    lPt = New Point
                                    lPt.X = tblX + 4.88
                                    lPt.Y = tblY + tableYOffset
                                    pEl = MakeALineElement("ptlistVLine3", lPt, cellHeight * (lineCount + 1), False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 6.53
                                    lPt.Y = tblY + tableYOffset
                                    pEl = MakeALineElement("ptlistVLine4", lPt, cellHeight * (lineCount + 1), False)
                                    pGC.AddElement(pEl, 0)

                                    lPt = New Point
                                    lPt.X = tblX + 7.43
                                    lPt.Y = tblY + tableYOffset
                                    pEl = MakeALineElement("ptlistVLine5", lPt, cellHeight * (lineCount + 1), False)
                                    pGC.AddElement(pEl, 0)
                                End If

                            End If


                            '************************

                            pActiveView.Refresh()

                            mapdoc.Map(1).MapScale = pMap.MapScale
                            pActiveView = mapdoc.Map(1)
                            pActiveView.Refresh()

                            CreatePDF(mapdoc.PageLayout, pdfpath & "\" & MapUtil.GetRowValue(pRow, "PEXUID") & ".pdf", 300, True, True)
                            _pdfCount = _pdfCount + 1

                            pGC.Reset()

                            pEl = pGC.Next
                            Do While Not pEl Is Nothing
                                Dim pEP As IElementProperties = pEl
                                If Left(pEP.Name, 6) = "ptlist" Then
                                    pGC.DeleteElement(pEl)
                                    pGC.Reset()
                                End If
                                pEl = pGC.Next
                            Loop
                        Else
                            log.WriteLine("Error: Pexuid: " & pexuid & " Can't find ownerply (" & square & "|" & suffix & "|" & lot & ") in layer " & pLayerParaName)
                        End If

                    Else
                        log.WriteLine("Premise " & pexuid & " does not exist in table " & tblIACF)
                    End If


                Catch ex As Exception
                    log.WriteLine("Error: Pexuid: " & pexuid & "- " & ex.Message)
                    log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

                    If Not bContunieOnError Then
                        Throw ex
                    End If
                End Try

            End While

            mapdoc.Save()

        Catch ex As Exception
            log.WriteLine("Error: " & ex.Message)
            If Not mapdoc Is Nothing Then
                mapdoc.Close()
            End If

            log.WriteLine(Now & ": (Error:)" & ex.Message)
            log.WriteLine(Now & ": (Error:)" & ex.StackTrace())

            Throw ex
        Finally
            oread.Close()
            log.WriteLine("End time: " & Now)
        End Try

        log.WriteLine(Now & " : Completed generating pdf ")

    End Sub

    Public Function MakeALineElement(ByVal name As String, ByVal pPoint As IPoint, ByVal length As Double, Optional ByVal horizontal As Boolean = True) As IElement
        Dim pLineElement As ILineElement
        Dim pElement As IElement = New LineElement


        Dim pPlolyline As ISegmentCollection = New Polyline
        Dim ppoint2 As IPoint = New Point
        If horizontal Then
            ppoint2.PutCoords(pPoint.X + length, pPoint.Y)
        Else
            ppoint2.PutCoords(pPoint.X, pPoint.Y - length)
        End If

        Dim pline As ILine = New Line
        pline.PutCoords(pPoint, ppoint2)
        Dim pSeg As ISegment = pline
        pPlolyline.AddSegment(pSeg)


        pElement.Geometry = pPlolyline

        Dim pColor As IColor
        Dim pLS As ISimpleLineSymbol

        pColor = New RgbColor
        With pColor
            .UseWindowsDithering = True
            .RGB = RGB(0, 0, 0)
        End With
        pLS = New SimpleLineSymbol
        pLS.Color = pColor
        pLS.Width = 1

        pLineElement = pElement
        pLineElement.Symbol = pLS

        Dim pEP As IElementProperties = pElement
        pEP.Name = name

        Return pElement
    End Function

    Public Function MakeATextElement(ByVal name As String, ByVal pPoint As IPoint, ByVal strText As String, ByVal fontsize As Integer, ByVal bold As Boolean) As IElement
        'Pointers needed to make text element    
        Dim pRGBcolor As IRgbColor
        Dim pTextElement As ITextElement
        Dim pTextSymbol As ITextSymbol

        Dim pElement As IElement
        'First setup a color.  We'll use RGB red   
        pRGBcolor = New RgbColor
        pRGBcolor.Blue = 0
        pRGBcolor.Red = 0
        pRGBcolor.Green = 0
        'Next, cocreate a new TextElement    
        pTextElement = New TextElement
        'Query Interface (QI) to an IElement pointer and set    
        'the geometry that was passed in   
        pElement = pTextElement
        pElement.Geometry = pPoint
        'Next, setup a font    

        Dim pFontDisp As IFontDisp
        pFontDisp = New StdFont
        pFontDisp.Name = "Arial"
        pFontDisp.Bold = bold
        'pFontDisp.Underline = True
        'Next, setup a TextSymbol that the TextElement will draw with    
        pTextSymbol = New TextSymbol
        pTextSymbol.Font = pFontDisp
        pTextSymbol.Color = pRGBcolor
        pTextSymbol.Size = fontsize
        pTextSymbol.HorizontalAlignment = esriTextHorizontalAlignment.esriTHALeft
        'set the size of the text symbol here, rather than on the font        
        'Next, Give the TextSymbol and text string to the TextElement    
        pTextElement.Symbol = pTextSymbol
        pTextElement.Text = strText
        'Finally, hand back the new element as the output of this function    

        Dim pEP As IElementProperties = pElement
        pEP.Name = name

        MakeATextElement = pTextElement


    End Function


#Region "Create Default FontDisp"
    ' ArcGIS Snippet Title:
    ' Create Default FontDisp
    ' 
    ' Long Description:
    ' Generate a default FontDisp object.
    ' 
    ' Add the following references to the project:
    ' stdole
    ' 
    ' Intended ArcGIS Products for this snippet:
    ' ArcGIS Desktop (ArcEditor, ArcInfo, ArcView)
    ' ArcGIS Engine
    ' ArcGIS Server
    ' 
    ' Applicable ArcGIS Product Versions:
    ' 9.2
    ' 9.3
    ' 
    ' Required ArcGIS Extensions:
    ' (NONE)
    ' 
    ' Notes:
    ' This snippet is intended to be inserted at the base level of a Class.
    ' It is not intended to be nested within an existing Function or Sub.
    ' 

    '''<summary>Generate a default FontDisp object.</summary>
    ''' 
    '''<returns>An stdole.IFontDisp interface</returns>
    '''  
    '''<remarks></remarks>
    Public Function CreateDefaultFontDisp() As stdole.IFontDisp

        Dim fontDisp As stdole.IFontDisp = CType(New stdole.StdFont, stdole.IFontDisp) ' Explicit Cast
        With fontDisp
            .Bold = False
            .Name = "Arial"
            .Italic = False
            .Underline = False
            .Size = 16
        End With

        Return fontDisp

    End Function
#End Region


    Private Sub CreatePDF( _
      ByVal pPageLayout As IPageLayout, _
      ByVal sFullPathName As String, _
      ByVal lngOutputResolution As Long, _
      ByVal blnEmbedFonts As Boolean, _
      ByVal blnPolygonizeMarkers As Boolean)

        Dim pExporter As IExport
        Dim pExportPDF As IExportPDF
        Dim pExportVectorOptions As IExportVectorOptions
        Dim pPixelEnv As IEnvelope
        Dim tExpRect As tagRECT
        Dim hdc As Long
        Dim pAV As IActiveView

        ' pixel envelope
        pPixelEnv = New Envelope
        pPixelEnv.PutCoords( _
          0, _
          0, _
          lngOutputResolution * PageExtent(pPageLayout).UpperRight.X, _
          lngOutputResolution * PageExtent(pPageLayout).UpperRight.Y)


        ' exporter object
        pExporter = New ExportPDF

        With pExporter
            .PixelBounds = pPixelEnv
            .Resolution = lngOutputResolution
            .ExportFileName = sFullPathName
        End With

        pExportPDF = pExporter
        pExportPDF.EmbedFonts = blnEmbedFonts

        pExportVectorOptions = pExporter
        pExportVectorOptions.PolygonizeMarkers = blnPolygonizeMarkers

        ' device coordinates origin is upper left, ypositive is down
        With tExpRect
            .left = pExporter.PixelBounds.LowerLeft.X
            .bottom = pExporter.PixelBounds.UpperRight.Y
            .right = pExporter.PixelBounds.UpperRight.X
            .top = pExporter.PixelBounds.LowerLeft.Y
        End With

        ' export
        hdc = pExporter.StartExporting
        pAV = pPageLayout
        pAV.Output(hdc, lngOutputResolution, tExpRect, Nothing, Nothing)
        pExporter.FinishExporting()

    End Sub

    Private Function PageExtent(ByVal pPageLayout As IPageLayout) As IEnvelope
        Dim dWidth As Double, dHeight As Double
        pPageLayout.Page.QuerySize(dWidth, dHeight)
        Dim pEnv As IEnvelope
        pEnv = New Envelope
        pEnv.PutCoords(0.0#, 0.0#, dWidth, dHeight)
        PageExtent = pEnv
    End Function

    Private Function PolygonFromEnvelope(ByVal pEnv As IEnvelope) As IPolygon
        Dim pPtColl As IPointCollection

        pPtColl = New Polygon
        With pPtColl
            .AddPoint(pEnv.LowerLeft)
            .AddPoint(pEnv.UpperLeft)
            .AddPoint(pEnv.UpperRight)
            .AddPoint(pEnv.LowerRight)
            'Close the polygon
            .AddPoint(pEnv.LowerLeft)
        End With

        PolygonFromEnvelope = pPtColl
    End Function




    Private Function GetSAD(ByRef pFeature As IFeature) As String
        Dim sad As String = ""
        If MapUtil.GetValue(pFeature, "USE_PARSED_ADDRESS") = "Y" Then
            '            address(number, address) suffix, street name, street type, quadrant, unit
            'ADDRNUM   ADDRNUMSUF   STNAME  STREET_TYPE   QUADRANT  UNIT
            sad = MapUtil.GetValue(pFeature, "ADDRNUM")
            sad = Trim(sad & " " & MapUtil.GetValue(pFeature, "ADDRNUMSUF"))
            sad = Trim(sad & " " & MapUtil.GetValue(pFeature, "STNAME"))
            sad = Trim(sad & " " & MapUtil.GetValue(pFeature, "STREET_TYPE"))
            sad = Trim(sad & " " & MapUtil.GetValue(pFeature, "QUADRANT", True))
            If MapUtil.GetValue(pFeature, "UNIT") <> "" Then
                sad = Trim(sad & " #" & MapUtil.GetValue(pFeature, "UNIT"))
            End If
        Else
            sad = MapUtil.GetValue(pFeature, "PEXSAD")
        End If

        Return sad
    End Function


#Region "Get hWnd for Current Application"
    ' ArcGIS Snippet Title:
    ' Get hWnd for Current Application
    ' 
    ' Long Description:
    ' Get the Windows Handle (hWnd) of the application that is currently running.
    ' 
    ' Add the following references to the project:
    ' System
    ' 
    ' Intended ArcGIS Products for this snippet:
    ' ArcGIS Desktop (ArcEditor, ArcInfo, ArcView)
    ' ArcGIS Engine
    ' ArcGIS Server
    ' 
    ' Applicable ArcGIS Product Versions:
    ' 9.2
    ' 9.3
    ' 
    ' Required ArcGIS Extensions:
    ' (NONE)
    ' 
    ' Notes:
    ' This snippet is intended to be inserted at the base level of a Class.
    ' It is not intended to be nested within an existing Function or Sub.
    ' 

    '''<summary>Get the Windows Handle (hWnd) of the application that is currently running.</summary>
    '''
    '''<returns>A System.Int32 that is windows handle of the running application.</returns>
    ''' 
    '''<remarks>Note: If the code is running inside an ArcGIS desktop application, you should use the IApplicaiton.hWnd to get the Windows Handle.</remarks>
    Public Function GetHWndForCurrentApplication() As System.Int32

        Return System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0)).ToInt32()

    End Function
#End Region



End Module
