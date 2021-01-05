Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Collections.Generic

<CLSCompliant(False)> _
Public Class PremiseAttribute
    Private m_premiseResult As Integer
    ' Listen to the selected changed events
    Private m_map As IMap
    Private m_app As IApplication

    Private m_premsPt As IFeature

    Private m_flagGetMarSelected As Boolean
    Private m_editMode As Integer

    Private m_bAddressValidated As Boolean
    Private m_bValidAddress As Boolean


    'temp variables for advanced attr
    Public inforsrc As String
    Public loctnprecs As String

    Public address_id As String
    Public master_pexuid As String
    Public comnt As String

    Public from_standalone As String
    Public has_lien As String


    Private m_db_pexptyp As String

    Public Event CommandButtonClicked(ByVal strCommand As String)

    Public Property premiseResult() As Integer
        Get
            premiseResult = m_premiseResult
        End Get
        Set(ByVal value As Integer)
            m_premiseResult = value
        End Set
    End Property

    Public WriteOnly Property PremiseMap() As IMap
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property

    Public WriteOnly Property PremiseApp() As IApplication
        Set(ByVal value As IApplication)
            m_app = value
        End Set
    End Property

    Public Property PremisePt() As IFeature
        Get
            PremisePt = m_premsPt
        End Get
        Set(ByVal value As IFeature)
            m_premsPt = value
            updateAttribute()
            m_bAddressValidated = True
        End Set
    End Property

    Public Property EditMode() As Integer
        Get
            EditMode = m_editMode
        End Get
        Set(ByVal value As Integer)
            m_editMode = value
            UpdateUI()
        End Set
    End Property

    Private Sub UpdateUI()
        If m_editMode = 1 Then
            'btnGetDDotAddr.Enabled = False
            btnGetSSLWard.Enabled = False

            TextSQUARE.Enabled = False
            TextSUFFIX.Enabled = False
            TextLOT.Enabled = False
            ComboWARD.Enabled = False
            ComboUSECODE.Enabled = False

        ElseIf m_editMode = 2 Then
            'btnGetDDotAddr.Enabled = True
            btnGetSSLWard.Enabled = True

            TextSQUARE.Enabled = True
            TextSUFFIX.Enabled = True
            TextLOT.Enabled = True
            ComboWARD.Enabled = True
            ComboUSECODE.Enabled = True

        End If
    End Sub

    Private Function InputValid() As Boolean

        If ComboPexPtyp.Text = "" Then
            MsgBox("Premise Type is required")
            Return False
        End If

        If Trim(TextZip5.Text) = "" Then
            MsgBox("Zip Code is required")
            Return False
        End If

        If Me.ComboQUADRANT.SelectedItem Is Nothing Or Trim(Me.ComboQUADRANT.Text) = "" Then
            MsgBox("Direction is required")
            Return False
        End If

        If chkSplitAddr.Checked Then
            If Trim(TextAddrnum.Text) = "" Then
                MsgBox("House Number is required")
                Return False
            End If

            If Trim(TextSTNAME.Text) = "" Then
                MsgBox("Street Name is required")
                Return False
            End If

            If ComboSTREET_TYPE.Text = "" Then
                MsgBox("Street Type is required")
                Return False
            ElseIf ComboSTREET_TYPE.SelectedItem Is Nothing Then
                MsgBox("Street Type is not formalized")
                Return False
            End If

        Else
            If Trim(textAddr.Text) = "" Then
                MsgBox("Full Address is required")
                Return False
            End If
        End If

        If ComboIAB_EXEMPT.SelectedItem Is Nothing Then
            MsgBox("Exemption Status is required")
            Return False
        ElseIf UCase(MapUtil.GetDomainCode(ComboIAB_EXEMPT)) = "Y" AndAlso ComboEXEMPT_REASON.SelectedItem Is Nothing Then
            MsgBox("Exemption Reason is required")
            Return False
        End If

        If ComboIA_ONLY.SelectedItem Is Nothing Then
            MsgBox("Impervious Only is required")
            Return False
        End If

        Return True
    End Function

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            m_premiseResult = 1
            'Save premise point 

            If Not InputValid() Then
                Return
            End If

            If Not m_bAddressValidated Then
                MsgBox("The address has been changed and was not validated. ")
                Return
            End If

            If Not m_bValidAddress Then
                If MsgBox("You have entered a invalid address, do you want to save the premise with this address?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Return
                End If
            End If

            Dim pUid As New UID
            pUid.Value = "esriEditor.Editor"
            Dim m_Editor As IEditor = m_app.FindExtensionByCLSID(pUid)

            IAISApplication.GetInstance.StartToolEditing()
            Try
                updateFeature()
                'Check to see if Premise Type was changed
                Dim pexptyp As String = MapUtil.GetValue(m_premsPt, "PEXPTYP")
                If EditMode = 2 AndAlso pexptyp <> m_db_pexptyp AndAlso (pexptyp = "RES" Or m_db_pexptyp = "RES") Then


                    Dim pFeatWS As IFeatureWorkspace = DirectCast(m_Editor.EditWorkspace, IFeatureWorkspace)
                    Dim pIACFTable As ITable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IACF"))
                    Dim pQFilter As IQueryFilter = New QueryFilter

                    Dim pRow As IRow
                    pQFilter.WhereClause = "EFFENDDT IS NULL AND PEXUID=" & MapUtil.GetValue(m_premsPt, "PEXUID")

                    Dim pCursor As ICursor
                    pCursor = pIACFTable.Search(pQFilter, False)
                    pRow = pCursor.NextRow
                    If Not pRow Is Nothing Then

                        If MsgBox("Premise type was changed and IA charge exists." & vbLf & " Do you want to save the change and transfer the IA charge?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            IAISApplication.GetInstance.AbortToolEditing()
                            Return
                        End If

                        Dim form As FormPickDate = New FormPickDate
                        form.LabelDate.Text = "Effective End Date"
                        form.Text = "Effective End Date"
                        form.ShowDialog()

                        If form.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
                            IAISApplication.GetInstance.AbortToolEditing()
                            Return
                        End If

                        pRow.Value(pRow.Fields.FindField("EFFENDDT")) = Form.DateTimePicker1.Text
                        pRow.Store()

                        Dim iaTierLevel As ITable = pFeatWS.OpenTable(IAISToolSetting.GetParameterValue("TABLE_IATIERLEVEL"))
                        Dim pNewIACF As IRow = pIACFTable.CreateRow
                        pNewIACF.Value(pNewIACF.Fields.FindField("PEXUID")) = pRow.Value(pRow.Fields.FindField("PEXUID"))
                        pNewIACF.Value(pNewIACF.Fields.FindField("IASQFT")) = pRow.Value(pRow.Fields.FindField("IASQFT"))
                        pNewIACF.Value(pNewIACF.Fields.FindField("IABILLERU")) = MapUtil.GetRoundERU(MapUtil.GetBillableERU(pexptyp, pRow.Value(pRow.Fields.FindField("IASQFT")), iaTierLevel), m_app)
                        pNewIACF.Value(pNewIACF.Fields.FindField("EFFSTARTDT")) = Form.DateTimePicker1.Text
                        pNewIACF.Value(pNewIACF.Fields.FindField("IA_SOURCE")) = pRow.Value(pRow.Fields.FindField("IA_SOURCE"))
                        pNewIACF.Value(pNewIACF.Fields.FindField("PARCEL_SOURCE")) = pRow.Value(pRow.Fields.FindField("PARCEL_SOURCE"))

                        pNewIACF.Value(pNewIACF.Fields.FindField("DBSTAMPDT")) = Now

                        pNewIACF.Store()
                    End If

                End If

                m_premsPt.Store()
                If EditMode = 1 Then
                    IAISApplication.GetInstance.StopToolEditing("Create Premise Point")
                Else
                    IAISApplication.GetInstance.StopToolEditing("Update Premise Point")
                End If

                RaiseEvent CommandButtonClicked("Ok")
                Me.Close()
            Catch ex As Exception
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            Throw ex
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_premiseResult = 0
        'remove the point if edit mode=1 (new premise point)
        If m_editMode = 1 Then
            'm_premsPt.Delete()
            Dim pUid As New UID
            pUid.Value = "esriEditor.Editor"
            Dim m_Editor As IEditor = m_app.FindExtensionByCLSID(pUid)

            IAISApplication.GetInstance.StartToolEditing()
            Try
                m_premsPt.Delete()
                IAISApplication.GetInstance.StopToolEditing("Abort Premise Point")
            Catch ex As Exception
                IAISApplication.GetInstance.AbortToolEditing()
                MsgBox(ex.Message)
                Return
            End Try

            Dim pMxDoc As IMxDocument = m_app.Document
            pMxDoc.OperationStack.Remove(pMxDoc.OperationStack.Count - 1)

        End If

        Dim pActiveview As IActiveView
        pActiveview = m_map
        pActiveview.Refresh()


        RaiseEvent CommandButtonClicked("Cancel")
        Me.Close()
    End Sub

    Private Sub btnGetMAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetMAR.Click
        MapUtil.SetCurrentTool("IAIS.PremiseTool", m_app, Me.EditMode)
        RaiseEvent CommandButtonClicked("GetMAR")
        'Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Function ValidateAddrWithMARTable() As Boolean
        'Dim addrptLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT"), m_map)
        Dim pWS As IFeatureWorkspace = MapUtil.GetPremiseWS(m_map)
        Dim pTable As ITable = MapUtil.GetTableFromWS(pWS, IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT"))

        Dim whereStr As String = ""
        'Dim fulladdress As String = TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text
        whereStr = "ADDRNUM=" & MapUtil.TrimAll(TextAddrnum.Text)
        If MapUtil.TrimAll(TextADDRNUMSUF.Text) <> "" Then
            whereStr = whereStr & " AND ADDRNUMSUFFIX='" & TextAddrnum.Text & "'"
        End If
        whereStr = whereStr & " AND STNAME='" & MapUtil.TrimAll(TextSTNAME.Text) & "'"
        If MapUtil.TrimAll(ComboSTREET_TYPE.Text) <> "" Then
            'Different schema
            If pTable.FindField("STREET_TYPE") >= 0 Then
                whereStr = whereStr & " AND STREET_TYPE='" & MapUtil.TrimAll(ComboSTREET_TYPE.Text) & "'"
            End If
        End If
        If ComboQUADRANT.Text <> "" Then
            whereStr = whereStr & " AND QUADRANT='" & ComboQUADRANT.Text & "'"
        End If

        whereStr = whereStr & " AND ZIPCODE=" & TextZip5.Text

        Dim pQueryDef As IQueryDef = pWS.CreateQueryDef
        pQueryDef.Tables = IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT")
        pQueryDef.SubFields = "COUNT(*)"
        pQueryDef.WhereClause = whereStr

        Dim pCursor As ICursor = pQueryDef.Evaluate()
        Dim pRow As IRow = pCursor.NextRow
        If pRow.Value(0) = 0 Then
            Return False
        Else
            MsgBox("Address is validated in MAR table")
            Return True
        End If

    End Function

    Private Function ValidateAddrWithMARWebService() As Boolean
        Dim MarServiceResult As gov.dc.citizenatlas.ReturnObject
        Dim score As Double
        Dim LocationVerifier As gov.dc.citizenatlas.LocationVerifier = New gov.dc.citizenatlas.LocationVerifier()
        Try
            Dim fulladdress As String = TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text & " " & Me.TextZip5.Text

            MarServiceResult = LocationVerifier.verifyDCAddress(TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text & " " & Me.TextZip5.Text)
            'MarServiceResult = LocationVerifier.verifyDCAddress(textAddr.Text)
            If MarServiceResult.returnDataset Is Nothing Then
                MarServiceResult = LocationVerifier.verifyDCAddress(TextAddrnum.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text & " " & Me.TextZip5.Text)
                If MarServiceResult.returnDataset Is Nothing Then
                    'MsgBox("Can't validate this address with MAR web service")
                    Return False
                Else
                    TextADDRNUMSUF.Text = ""
                    TextUNIT.Text = TextADDRNUMSUF.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

        'MapUtil.PrintDataSet(MarServiceResult.returnDataset)

        Dim formMar As FormMarAddress = New FormMarAddress
        PopulateMarList(formMar.ListMARAddress, MarServiceResult.returnDataset)

        Dim marAddra As clsMARAddress

        If formMar.ListMARAddress.Items.Count = 1 Then
            marAddra = formMar.ListMARAddress.Items.Item(0)
            If marAddra.Score = 100 Then
                Me.TextAddrnum.Text = marAddra.AddrNum
                Me.TextSTNAME.Text = marAddra.StName
                Me.TextADDRNUMSUF.Text = marAddra.AddrNumSuf
                Me.TextZip5.Text = marAddra.ZipCode
                MapUtil.SetComboIndex(Me.ComboSTREET_TYPE, marAddra.StreetType)
                MapUtil.SetComboIndex(Me.ComboQUADRANT, marAddra.Quadrant)

                m_bAddressValidated = True
                m_bValidAddress = True

                MsgBox("Address is validated DCGIS MAR webservice.")
                formMar.Dispose()
                Return True
            End If
        End If

        formMar.ShowDialog()
        If formMar.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
            Return False
        End If

        If formMar.ListMARAddress.SelectedItem Is Nothing Then
            Return False
        End If


        marAddra = formMar.ListMARAddress.SelectedItem

        Me.TextAddrnum.Text = marAddra.AddrNum
        Me.TextSTNAME.Text = marAddra.StName
        Me.TextADDRNUMSUF.Text = marAddra.AddrNumSuf
        Me.TextZip5.Text = marAddra.ZipCode
        MapUtil.SetComboIndex(Me.ComboSTREET_TYPE, marAddra.StreetType)
        MapUtil.SetComboIndex(Me.ComboQUADRANT, marAddra.Quadrant)

        formMar.Dispose()

        Return True
    End Function

    Private Function ValidateAddrWithStreetTable() As Boolean
        Dim pWS As IFeatureWorkspace = MapUtil.GetPremiseWS(m_map)
        Dim pTable As ITable = MapUtil.GetTableFromWS(pWS, IAISToolSetting.GetParameterValue("TABLE_VALIDSTREET"))

        Dim whereStr As String = ""
        whereStr = "STNAME='" & MapUtil.TrimAll(TextSTNAME.Text) & "'"
        Dim sttype As String = GetUSPSStType(ComboSTREET_TYPE.Text)
        If MapUtil.TrimAll(ComboSTREET_TYPE.Text) <> "" Then
            'Different schema
            If pTable.FindField("STREET_TYPE") >= 0 Then
                whereStr = whereStr & " AND STREET_TYPE='" & sttype & "'"
            End If
        End If
        If ComboQUADRANT.Text <> "" Then
            whereStr = whereStr & " AND QUADRANT='" & ComboQUADRANT.Text & "'"
        End If

        Dim pQueryDef As IQueryDef = pWS.CreateQueryDef
        pQueryDef.Tables = IAISToolSetting.GetParameterValue("TABLE_VALIDSTREET")
        pQueryDef.SubFields = "COUNT(*)"
        pQueryDef.WhereClause = whereStr

        Dim pCursor As ICursor = pQueryDef.Evaluate()
        Dim pRow As IRow = pCursor.NextRow
        If pRow.Value(0) = 0 Then
            Return False
        Else
            MsgBox("Address is validated with table " & IAISToolSetting.GetParameterValue("TABLE_VALIDSTREET"))
            Return True
        End If
    End Function

    Private Sub btnValidateAddr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidateAddr.Click

        Me.Cursor = Cursors.WaitCursor
        Try

            If Not chkSplitAddr.Checked Then

                Me.textAddr.Text = MapUtil.TrimAll(Me.textAddr.Text)
                m_bValidAddress = True
                m_bAddressValidated = True

                MsgBox("Limited validation of non-split address completed")

                Return
            End If

            If Trim(TextAddrnum.Text) = "" Then
                MsgBox("Address Number is required")
                Return
            End If

            If Trim(TextSTNAME.Text) = "" Then
                MsgBox("Street Name is required")
                Return
            End If

            If Trim(TextZip5.Text) = "" Then
                MsgBox("Zipcode is required")
                Return
            End If

            If ComboSTREET_TYPE.Text = "" Then
                MsgBox("Street Type is required")
                Return
            End If

            If ValidateAddrWithMARTable() Then
                m_bValidAddress = True
                m_bAddressValidated = True
                Return
            End If

            If ValidateAddrWithMARWebService() Then
                m_bValidAddress = True
                m_bAddressValidated = True
                Return
            End If

            If ValidateAddrWithStreetTable() Then
                m_bValidAddress = True
                m_bAddressValidated = True
                Return
            End If

            MsgBox("Invalid address")
            m_bAddressValidated = True

        Catch ex As Exception
            Throw ex
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub updateAttribute()
        'PopulateDomain
        Dim fields As IFields
        fields = m_premsPt.Fields

        TextPexprm.Text = MapUtil.GetValue(m_premsPt, "PEXPRM")
        Dim pexpsts As String = UCase(MapUtil.GetValue(m_premsPt, "PEXPSTS"))

        If EditMode = 1 Then
            MapUtil.PopulateDomain(Me.ComboPexSts, fields.Field(fields.FindField("PEXPSTS")).Domain, True, MapUtil.GetValue(m_premsPt, "PEXPSTS"))
        Else
            Dim pCodedDomain As ICodedValueDomain = DirectCast(m_premsPt.Fields.Field(m_premsPt.Fields.FindField("PEXPSTS")).Domain, ICodedValueDomain)
            If pexpsts = "IA" Or pexpsts = "PG" Or pexpsts = "ST" Then
                ComboPexSts.Enabled = True
                ComboPexSts.Items.Clear()
                Dim cindex As Long
                For cindex = 0 To pCodedDomain.CodeCount - 1
                    If pCodedDomain.Value(cindex) = "IA" Or pCodedDomain.Value(cindex) = "PG" Or pCodedDomain.Value(cindex) = "ST" Then
                        Dim domain As clsGisDomain = New clsGisDomain
                        domain.Name = pCodedDomain.Name(cindex)
                        domain.Value = pCodedDomain.Value(cindex)

                        ComboPexSts.Items.Add(domain)
                    End If
                Next
            Else
                ComboPexSts.Enabled = False
                MapUtil.PopulateDomain(Me.ComboPexSts, fields.Field(fields.FindField("PEXPSTS")).Domain, True, MapUtil.GetValue(m_premsPt, "PEXPSTS"))
                'ComboPexSts.Items.Clear()
                'Dim cindex As Long
                'For cindex = 0 To pCodedDomain.CodeCount - 1
                '    If pCodedDomain.Value(cindex) = pexpsts Then
                '        Dim domain As clsGisDomain = New clsGisDomain
                '        domain.Name = pCodedDomain.Name(cindex)
                '        domain.Value = pCodedDomain.Value(cindex)

                '        ComboPexSts.Items.Add(domain)
                '    End If
                'Next

            End If
        End If


        ComboPexSts.DisplayMember = "Name"
        ComboPexSts.ValueMember = "Value"

        'If pexpsts = "AC" Then
        '    ComboPexPtyp.Enabled = False
        'End If

        MapUtil.SetComboIndex(Me.ComboPexSts, MapUtil.GetValue(m_premsPt, "PEXPSTS", True))
        MapUtil.PopulateDomain(Me.ComboPexPtyp, fields.Field(fields.FindField("PEXPTYP")).Domain, False, MapUtil.GetValue(m_premsPt, "PEXPTYP"))

        m_db_pexptyp = MapUtil.GetValue(m_premsPt, "PEXPTYP")

        Me.TextAddrnum.Text = MapUtil.GetValue(m_premsPt, "ADDRNUM")
        Me.TextADDRNUMSUF.Text = MapUtil.GetValue(m_premsPt, "ADDRNUMSUF")
        Me.TextSTNAME.Text = MapUtil.GetValue(m_premsPt, "STNAME")

        Me.TextUNIT.Text = MapUtil.GetValue(m_premsPt, "UNIT")

        Me.textAddr.Text = MapUtil.GetValue(m_premsPt, "PEXSAD")

        PopulateStType()
        MapUtil.SetComboIndex(Me.ComboSTREET_TYPE, GetStType(MapUtil.GetValue(m_premsPt, "STREET_TYPE")))

        If MapUtil.GetValue(m_premsPt, "USE_PARSED_ADDRESS") = "Y" Then
            chkSplitAddr.Checked = True
            textAddr.Text = ""
        Else
            chkSplitAddr.Checked = False
        End If

        ResetAddressComponents()

        Dim pexzip As String = MapUtil.GetValue(m_premsPt, "PEXPZIP")
        pexzip = Replace(pexzip, "-", "")

        Me.TextZip5.Text = Mid(pexzip, 1, 5)

        If Len(pexzip) > 5 Then
            Me.TextZip4.Text = Mid(pexzip, 6, 4)
        End If

        MapUtil.PopulateDomain(Me.ComboQUADRANT, fields.Field(fields.FindField("QUADRANT")).Domain, True, MapUtil.GetValue(m_premsPt, "QUADRANT"))

        Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
        Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
        Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")

        MapUtil.PopulateDomain(Me.ComboWARD, fields.Field(fields.FindField("WARD")).Domain, True, MapUtil.GetValue(m_premsPt, "WARD"))
        'MapUtil.SetComboIndex(Me.ComboWARD, MapUtil.GetValue(m_premsPt, "WARD"))

        MapUtil.PopulateDomain(Me.ComboIAB_EXEMPT, fields.Field(fields.FindField("IS_EXEMPT_IAB")).Domain, False, MapUtil.GetValue(m_premsPt, "IS_EXEMPT_IAB"))
        If MapUtil.GetValue(m_premsPt, "IS_EXEMPT_IAB") = "Y" Then
            Me.ComboEXEMPT_REASON.Enabled = True
        Else
            Me.ComboEXEMPT_REASON.Enabled = False
        End If

        MapUtil.PopulateDomain(Me.ComboEXEMPT_REASON, fields.Field(fields.FindField("EXEMPT_IAB_REASON")).Domain, False, MapUtil.GetValue(m_premsPt, "EXEMPT_IAB_REASON"))
        MapUtil.PopulateDomain(Me.ComboIA_ONLY, fields.Field(fields.FindField("IS_IMPERVIOUS_ONLY")).Domain, False, MapUtil.GetValue(m_premsPt, "IS_IMPERVIOUS_ONLY"))

        MapUtil.PopulateDomain(Me.ComboUSECODE, fields.Field(fields.FindField("OTRUSECODE")).Domain, True, MapUtil.GetValue(m_premsPt, "OTRUSECODE"))

        If Trim(TextPexprm.Text) = "" Then
            TextSQUARE.Enabled = False
            TextSUFFIX.Enabled = False
            TextLOT.Enabled = False

            ComboWARD.Enabled = False
            ComboUSECODE.Enabled = False

            btnGetSSLWard.Enabled = False
        End If

        If m_editMode = 1 Then
            ComboPexSts.Enabled = False
            ComboIAB_EXEMPT.Enabled = False
            ComboEXEMPT_REASON.Enabled = False
            m_bValidAddress = False
        Else
            If pexpsts = "PN" Or pexpsts = "PG" Then
                ComboIAB_EXEMPT.Enabled = False
                ComboEXEMPT_REASON.Enabled = False
            End If
            m_bValidAddress = True
        End If


        address_id = MapUtil.GetValue(m_premsPt, "ADDRESS_ID")
        master_pexuid = MapUtil.GetValue(m_premsPt, "MASTER_PEXUID")
        comnt = MapUtil.GetValue(m_premsPt, "COMNT")

        loctnprecs = MapUtil.GetValue(m_premsPt, "LOCTNPRECS")
        inforsrc = MapUtil.GetValue(m_premsPt, "INFORSRC")

        from_standalone = MapUtil.GetValue(m_premsPt, "FROM_STANDALONE")
        has_lien = MapUtil.GetValue(m_premsPt, "HAS_LIEN")
    End Sub


    Private Sub updateFeature()
        'PopulateDomain
        Dim fields As IFields
        fields = m_premsPt.Fields

        UpdateFeatureValue(m_premsPt, fields.FindField("PEXPSTS"), MapUtil.GetDomainCode(Me.ComboPexSts), True)
        UpdateFeatureValue(m_premsPt, fields.FindField("PEXPTYP"), MapUtil.GetDomainCode(Me.ComboPexPtyp), True)

        UpdateFeatureValue(m_premsPt, fields.FindField("ADDRNUM"), Me.TextAddrnum.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("ADDRNUMSUF"), Me.TextADDRNUMSUF.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("STNAME"), Me.TextSTNAME.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("STREET_TYPE"), GetUSPSStType(Me.ComboSTREET_TYPE.Text))

        If chkSplitAddr.Checked Then
            UpdateFeatureValue(m_premsPt, fields.FindField("USE_PARSED_ADDRESS"), "Y", True)
        Else
            UpdateFeatureValue(m_premsPt, fields.FindField("USE_PARSED_ADDRESS"), "N", True)
        End If

        UpdateFeatureValue(m_premsPt, fields.FindField("UNIT"), Me.TextUNIT.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("PEXSAD"), Me.textAddr.Text)

        UpdateFeatureValue(m_premsPt, fields.FindField("PEXPZIP"), Me.TextZip5.Text & Me.TextZip4.Text)

        UpdateFeatureValue(m_premsPt, fields.FindField("QUADRANT"), MapUtil.GetDomainCode(ComboQUADRANT), True)

        UpdateFeatureValue(m_premsPt, fields.FindField("PEXSQUARE"), Me.TextSQUARE.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("PEXSUFFIX"), Me.TextSUFFIX.Text)
        UpdateFeatureValue(m_premsPt, fields.FindField("PEXLOT"), Me.TextLOT.Text)

        UpdateFeatureValue(m_premsPt, fields.FindField("WARD"), MapUtil.GetDomainCode(Me.ComboWARD), True)

        UpdateFeatureValue(m_premsPt, fields.FindField("IS_EXEMPT_IAB"), MapUtil.GetDomainCode(Me.ComboIAB_EXEMPT), True)
        UpdateFeatureValue(m_premsPt, fields.FindField("EXEMPT_IAB_REASON"), MapUtil.GetDomainCode(Me.ComboEXEMPT_REASON), True)
        UpdateFeatureValue(m_premsPt, fields.FindField("IS_IMPERVIOUS_ONLY"), MapUtil.GetDomainCode(Me.ComboIA_ONLY), True)

        UpdateFeatureValue(m_premsPt, fields.FindField("OTRUSECODE"), MapUtil.GetDomainCode(Me.ComboUSECODE), True)


        m_premsPt.Value(fields.FindField("XCOOR")) = m_premsPt.Shape.Envelope.XMax
        m_premsPt.Value(fields.FindField("YCOOR")) = m_premsPt.Shape.Envelope.YMax

        If MapUtil.GetRowValue(m_premsPt, "PEXUID") = "" Then
            m_premsPt.Value(fields.FindField("PEXUID")) = MapUtil.GetSequenceID(MapUtil.GetPremiseWS(Me.m_map), IAISToolSetting.GetParameterValue("SEQUENCE_PEXUID"))
        End If

        If MapUtil.GetRowValue(m_premsPt, "IS_EXEMPT_IAB") = "N" Or _
            (MapUtil.GetRowValue(m_premsPt, "IS_EXEMPT_IAB") = "Y" And _
             Not (MapUtil.GetRowValue(m_premsPt, "EXEMPT_IAB_REASON") = "SUBOR" Or MapUtil.GetRowValue(m_premsPt, "EXEMPT_IAB_REASON") = "DUPLI")) Then
            m_premsPt.Value(fields.FindField("MASTER_PEXUID")) = System.DBNull.Value
        End If

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("LOCTNPRECS"), loctnprecs, True)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("INFORSRC"), inforsrc, True)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("ADDRESS_ID"), address_id)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("MASTER_PEXUID"), master_pexuid)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("COMNT"), comnt)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("FROM_STANDALONE"), from_standalone, True)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("HAS_LIEN"), has_lien, True)

    End Sub


    Private Function UpdateFeatureValue(ByVal pFeature As IFeature, ByVal fIndex As Integer, ByVal fValue As String, Optional ByVal valueOnly As Boolean = False) As Boolean
        'Try

        If fIndex <= 0 Then
            Throw New Exception("Field index is invalid : " & fIndex)
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
                If IsDBNull(MapUtil.GetDomainCode(pCodeDomain, fValue)) Then
                    If Not pField.IsNullable Then
                        If pField.Type = esriFieldType.esriFieldTypeString Then
                            pFeature.Value(fIndex) = ""
                        Else
                            pFeature.Value(fIndex) = 0
                        End If
                    Else
                        pFeature.Value(fIndex) = System.DBNull.Value
                    End If
                Else
                    pFeature.Value(fIndex) = MapUtil.GetDomainCode(pCodeDomain, fValue)
                End If
            Else
                pFeature.Value(fIndex) = fValue
            End If
        Else
            If Trim(fValue) = "" Then
                If pField.Type = esriFieldType.esriFieldTypeString Then
                    If IsDBNull(pFeature.Value(fIndex)) Then
                        Return False
                    ElseIf fValue = pFeature.Value(fIndex) Then
                        Return False
                    Else
                        pFeature.Value(fIndex) = System.DBNull.Value
                    End If
                Else
                    If Not IsDBNull(pFeature.Value(fIndex)) _
                        AndAlso CStr(pFeature.Value(fIndex)) <> "" _
                        AndAlso pField.IsNullable Then
                        pFeature.Value(fIndex) = System.DBNull.Value
                    End If
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
                If IsDBNull(pFeature.Value(fIndex)) Then
                    pFeature.Value(fIndex) = MapUtil.TrimAll(fValue)
                ElseIf fValue = pFeature.Value(fIndex) Then
                    Return False
                Else
                    pFeature.Value(fIndex) = MapUtil.TrimAll(fValue)
                End If
            Else
                pFeature.Value(fIndex) = fValue
            End If
        End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Function

    Private Sub ComboIAB_EXEMPT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboIAB_EXEMPT.SelectedIndexChanged
        'Debug.Print("ComboIAB_EXEMPT_SelectedIndexChanged: " & ComboIAB_EXEMPT.SelectedText)
        If UCase(MapUtil.GetDomainCode(ComboIAB_EXEMPT)) = "Y" Then
            Me.ComboEXEMPT_REASON.Enabled = True
        Else
            Me.ComboEXEMPT_REASON.SelectedIndex = -1
            Me.ComboEXEMPT_REASON.Enabled = False
        End If
    End Sub


    Private Sub btnGetSSLWard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetSSLWard.Click
        Dim wardLayer As IFeatureLayer
        Dim colLayer As IFeatureLayer

        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
        Dim pDataset As IDataset = premiseLayer


        wardLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_WARDPLY"), m_map)
        Dim pSFilter As ISpatialFilter
        pSFilter = New SpatialFilter

        pSFilter.Geometry = m_premsPt.Shape
        pSFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin

        Dim pFeatCursor As IFeatureCursor
        Dim pFeature As IFeature

        If wardLayer Is Nothing Then

            wardLayer = MapUtil.GetLayerFromWS(pDataset.Workspace, _
                                             IAISToolSetting.GetParameterValue("LAYER_WARDPLY"))

            If wardLayer Is Nothing Then
                MsgBox("Ward layer is not loaded and can't be loaded from database.")
            End If
        End If

        If Not wardLayer Is Nothing Then
            pFeatCursor = wardLayer.Search(pSFilter, False)
            pFeature = pFeatCursor.NextFeature
            If Not pFeature Is Nothing Then
                m_premsPt.Value(m_premsPt.Fields.FindField("WARD")) = MapUtil.GetValue(pFeature, "WARD_ID")
                MapUtil.SetComboIndex(Me.ComboWARD, MapUtil.GetValue(m_premsPt, "WARD", True))
            End If
        End If

        Dim pLayerList As List(Of IFeatureLayer) = MapUtil.GetColLayers(m_map)

        Dim pCol As IFeature = MapUtil.SelectedCol(m_map, pSFilter)
        If pCol Is Nothing Then
            'MsgBox("")
            Return
        End If

        If pCol.Fields.FindField("USECODE") >= 0 Then
            If Trim(MapUtil.GetValue(pCol, "USECODE")) = "" Then
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), "000", True)
            Else
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), MapUtil.GetValue(pCol, "USECODE"), True)
            End If
        Else
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), "000", True)
        End If

        Dim ssl As String = MapUtil.GetValue(pCol, "SSL")
        If ssl.Length < 12 Then
            MsgBox("Property feature has invalid SSL: " & ssl)
        Else
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSQUARE"), ssl.Substring(0, 4))
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSUFFIX"), ssl.Substring(4, 4))
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXLOT"), ssl.Substring(8))
        End If

        Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
        Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
        Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")
        MapUtil.SetComboIndex(ComboUSECODE, MapUtil.GetValue(m_premsPt, "OTRUSECODE", True))

        'colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
        'If colLayer Is Nothing Then
        '    colLayer = MapUtil.GetLayerFromWS(pDataset.Workspace, _
        '                                                 IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"))
        'End If
        'If Not colLayer Is Nothing Then
        '    pFeatCursor = colLayer.Search(pSFilter, False)
        '    pFeature = pFeatCursor.NextFeature
        '    If Not pFeature Is Nothing Then
        '        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSQUARE"), MapUtil.GetValue(pFeature, "SQUARE"))
        '        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSUFFIX"), MapUtil.GetValue(pFeature, "SUFFIX"))
        '        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXLOT"), MapUtil.GetValue(pFeature, "LOT"))
        '        If Trim(MapUtil.GetValue(pFeature, "USECODE")) = "" Then
        '            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), "000", True)
        '        Else
        '            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), MapUtil.GetValue(pFeature, "USECODE"), True)
        '        End If


        '        Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
        '        Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
        '        Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")
        '        MapUtil.SetComboIndex(ComboUSECODE, MapUtil.GetValue(m_premsPt, "OTRUSECODE", True))

        '        Return
        '    End If
        'End If

        'colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), m_map)
        'If colLayer Is Nothing Then
        '    colLayer = MapUtil.GetLayerFromWS(pDataset.Workspace, _
        '                                                 IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"))
        'End If

        'pFeatCursor = colLayer.Search(pSFilter, False)
        'pFeature = pFeatCursor.NextFeature
        'If Not pFeature Is Nothing Then
        '    MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSQUARE"), MapUtil.GetValue(pFeature, "SQUARE"))
        '    MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSUFFIX"), MapUtil.GetValue(pFeature, "SUFFIX"))
        '    MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXLOT"), MapUtil.GetValue(pFeature, "LOT"))

        '    Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
        '    Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
        '    Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")

        'End If

    End Sub

    Private Sub btnGetDDotAddr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetDDotAddr.Click
        MapUtil.SetCurrentTool("IAIS.PremiseTool", m_app, Me.EditMode)
        RaiseEvent CommandButtonClicked("GetDDotAddr")
    End Sub

    Private Sub textAddr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textAddr.TextChanged
        If Not Me.chkSplitAddr.Checked Then
            If textAddr.Text.Length >= 2 Then
                m_bAddressValidated = False
                m_bValidAddress = False

                Me.btnValidateAddr.Enabled = True
            Else
                m_bAddressValidated = True
                m_bValidAddress = True

                Me.btnValidateAddr.Enabled = False
            End If
        End If
    End Sub

    Public Sub SetDDotAddress(ByVal fulladdr As String)
        textAddr.Text = fulladdr
        textAddr.Enabled = False

        TextAddrnum.Enabled = True
        TextSTNAME.Enabled = True

        ComboSTREET_TYPE.Enabled = True
        ComboQUADRANT.Enabled = True
        TextADDRNUMSUF.Enabled = True
        address_id = ""

        m_bAddressValidated = False

        ParseAddress()
        chkSplitAddr.Checked = True
        btnValidateAddr.Enabled = True

    End Sub


    Public Sub SetMarAddress(ByVal fulladdr As String, _
                             ByVal addrid As String, _
                             ByVal addrnum As String, _
                             ByVal stname As String, _
                             ByVal sttype As String, _
                             ByVal addrsuf As String, _
                             ByVal quadrant As String, _
                             ByVal unit As String, _
                             ByVal zipcode As String)

        'textAddr.Text = fulladdr
        textAddr.Text = ""
        textAddr.Enabled = False
        TextAddrnum.Text = addrnum
        TextAddrnum.Enabled = False
        TextSTNAME.Text = stname
        TextSTNAME.Enabled = False

        TextZip5.Text = zipcode
        MapUtil.SetComboIndex(ComboSTREET_TYPE, sttype)
        ComboSTREET_TYPE.Enabled = False
        MapUtil.SetComboIndex(ComboQUADRANT, quadrant)
        ComboQUADRANT.Enabled = False

        TextADDRNUMSUF.Text = addrsuf
        TextADDRNUMSUF.Enabled = False

        TextUNIT.Text = unit

        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("ADDRESS_ID"), addrid)
        address_id = addrid
        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("LOCTNPRECS"), 6, True)
        loctnprecs = 6
        m_bAddressValidated = True
        m_bValidAddress = True

        btnValidateAddr.Enabled = False
        chkSplitAddr.Checked = True

    End Sub

    Private Sub ParseAddress()

        Dim FullAddress As String
        FullAddress = textAddr.Text
        textAddr.Text = ""
        textAddr.Enabled = False

        If Trim(FullAddress) = "" Then
            Return
        End If

        Dim lAddrNum As String
        Dim sAddrSufx As String
        Dim strStName As String
        Dim strStTYPE As String
        Dim strQuad As String
        Dim strUnit As String
        strUnit = ""
        sAddrSufx = ""



        Dim pos As Integer
        pos = InStrRev(FullAddress, "#")
        If (pos > 0) Then
            strUnit = Trim(Mid(FullAddress, pos + 1))
            FullAddress = Trim(Mid(FullAddress, 1, pos - 1))
        End If

        pos = InStrRev(FullAddress, " ")
        strQuad = UCase(Mid(FullAddress, pos + 1))
        If strQuad = "NW" Or strQuad = "NE" Or strQuad = "SE" Or strQuad = "SW" Then
            FullAddress = Trim(Mid(FullAddress, 1, pos))
        Else
            strQuad = ""
        End If

        pos = InStrRev(FullAddress, " ")
        strStTYPE = Mid(FullAddress, pos + 1)
        'Check it it is valid street type

        FullAddress = Trim(Mid(FullAddress, 1, pos))
        pos = InStr(FullAddress, " ")

        If IsNumeric(Mid(FullAddress, 1, pos)) Then
            lAddrNum = Trim(Mid(FullAddress, 1, pos))
            FullAddress = Trim(Mid(FullAddress, pos))
        Else
            MsgBox("There is no house number")
        End If

        'Check to see if we have fraction
        pos = InStr(FullAddress, " ")
        If pos > 0 Then
            If Mid(FullAddress, 1, pos).Contains("/") Then
                sAddrSufx = Mid(FullAddress, 1, pos)
                FullAddress = Trim(Mid(FullAddress, pos))
            End If
        End If


        strStName = FullAddress

        TextAddrnum.Text = lAddrNum
        TextSTNAME.Text = strStName
        MapUtil.SetComboIndex(ComboSTREET_TYPE, strStTYPE)
        If ComboSTREET_TYPE.SelectedIndex = -1 Then
            MapUtil.SetComboIndex(ComboSTREET_TYPE, Me.GetStType(strStTYPE))
        End If

        MapUtil.SetComboIndex(ComboQUADRANT, strQuad)

        TextADDRNUMSUF.Text = sAddrSufx
        TextUNIT.Text = strUnit

        chkSplitAddr.Checked = True
        textAddr.Text = ""
    End Sub
    Private Sub ResetAddressComponents()

        If chkSplitAddr.Checked Then
            textAddr.Enabled = False

            TextAddrnum.Enabled = True
            TextSTNAME.Enabled = True
            ComboSTREET_TYPE.Enabled = True
            TextADDRNUMSUF.Enabled = True
            TextUNIT.Enabled = True

            btnValidateAddr.Enabled = True
        Else
            textAddr.Enabled = True

            TextAddrnum.Enabled = False
            TextSTNAME.Enabled = False
            ComboSTREET_TYPE.Enabled = False
            TextADDRNUMSUF.Enabled = False

            TextUNIT.Enabled = False
        End If

    End Sub

    Private Sub PremiseAttribute_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        RaiseEvent CommandButtonClicked("FormClosing")
    End Sub

    Private Sub PremiseAttribute_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Tab Then
            e.SuppressKeyPress = True
            If e.Modifiers = Keys.Shift Then
                Me.ProcessTabKey(False)
            Else
                Me.ProcessTabKey(True)
            End If
        End If
    End Sub

    Private Sub PopulateStType()
        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
        Dim pDataset As IDataset = premiseLayer
        Dim pTable As ITable = MapUtil.GetTableFromWS(pDataset.Workspace, IAISToolSetting.GetParameterValue("TABLE_STREET_TYPE"))

        Dim pTableSort As ITableSort = New TableSort
        With pTableSort
            .Fields = "STREET_TYPE"
            .QueryFilter = Nothing
            .Table = pTable
        End With
        pTableSort.Ascending("STREET_TYPE") = True



        pTableSort.Sort(Nothing)
        Dim pCursor As ICursor = pTableSort.Rows

        Dim pRow As IRow
        pRow = pCursor.NextRow

        Me.ComboSTREET_TYPE.Items.Clear()
        Me.ComboSTREET_TYPE.Items.Add("")


        Do While Not pRow Is Nothing
            Me.ComboSTREET_TYPE.Items.Add(pRow.Value(pRow.Fields.FindField("STREET_TYPE")))
            pRow = pCursor.NextRow
        Loop
    End Sub

    Private Function GetUSPSStType(ByVal sttype As String) As String
        If Trim(sttype) = "" Then
            Return ""
        End If

        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
        Dim pDataset As IDataset = premiseLayer
        Dim pTable As ITable = MapUtil.GetTableFromWS(pDataset.Workspace, IAISToolSetting.GetParameterValue("TABLE_STREET_TYPE"))
        Dim pFilter As IQueryFilter = New QueryFilter

        pFilter.WhereClause = "STREET_TYPE='" & sttype & "'"

        Dim pCursor As ICursor = pTable.Search(pFilter, False)

        Dim pRow As IRow
        pRow = pCursor.NextRow

        If Not pRow Is Nothing Then
            Return pRow.Value(pRow.Fields.FindField("USPS_ST_TYPE"))
        Else
            Return sttype
        End If
    End Function

    Private Function GetStType(ByVal uspssttype As String) As String
        If Trim(uspssttype) = "" Then
            Return ""
        End If

        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_PREMSINTERPT"), m_map)
        Dim pDataset As IDataset = premiseLayer
        Dim pTable As ITable = MapUtil.GetTableFromWS(pDataset.Workspace, IAISToolSetting.GetParameterValue("TABLE_STREET_TYPE"))
        Dim pFilter As IQueryFilter = New QueryFilter

        pFilter.WhereClause = "USPS_ST_TYPE='" & uspssttype & "'"

        Dim pCursor As ICursor = pTable.Search(pFilter, False)

        Dim pRow As IRow
        pRow = pCursor.NextRow

        If Not pRow Is Nothing Then
            Return pRow.Value(pRow.Fields.FindField("STREET_TYPE"))
        Else
            Return uspssttype
        End If
    End Function

    Private Sub TextAddrnum_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextAddrnum.KeyDown
        'If Not e.Control Then
        '    e.Handled = IsNumericKey(e.KeyCode)
        'Else
        '    e.Handled = False
        'End If

    End Sub

    Private Sub TextAddrnum_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextAddrnum.KeyPress
        e.Handled = MapUtil.IsNumericKey(Asc(e.KeyChar))
    End Sub


    'Private Function IsNumericKey(ByVal KCode As String) As Boolean
    '    If (KCode >= 48 And KCode <= 57) Or KCode = 8 Or KCode = 3 Or KCode = 22 Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function

    Private Sub TextZip5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextZip5.KeyPress
        e.Handled = MapUtil.IsNumericKey(Asc(e.KeyChar))
    End Sub

    Private Sub TextAddrnum_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextAddrnum.TextChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Private Sub TextADDRNUMSUF_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextADDRNUMSUF.TextChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Private Sub TextSTNAME_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextSTNAME.TextChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Private Sub ComboSTREET_TYPE_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboSTREET_TYPE.SelectedIndexChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Private Sub ComboQUADRANT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboQUADRANT.SelectedIndexChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Public Shared Sub PopulateMarList(ByRef pList As ListBox, ByVal ds As DataSet)
        If ds.Tables.Count < 1 Then
            Return
        End If

        Dim table As DataTable = ds.Tables.Item(0)
        For Each row As DataRow In table.Rows
            Dim marAddr As clsMARAddress = New clsMARAddress
            Dim colIndex As Integer


            For colIndex = 0 To table.Columns.Count - 1
                'Debug.Print(table.Columns.Item(colIndex).ColumnName)

                If UCase(table.Columns.Item(colIndex).ColumnName) = "ADDRNUM" Then
                    marAddr.AddrNum = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "STNAME" Then
                    marAddr.StName = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "STREET_TYPE" Then
                    marAddr.StreetType = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "QUADRANT" Then
                    marAddr.Quadrant = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "CONFIDENCELEVEL" Then
                    marAddr.Score = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "FULLADDRESS" Then
                    marAddr.FullAddress = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "ZIPCODE" Then
                    marAddr.ZipCode = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "ADDRNUMSUFFIX" Then
                    marAddr.AddrNumSuf = row.Item(colIndex).ToString
                ElseIf UCase(table.Columns.Item(colIndex).ColumnName) = "ADDRESS_ID" Then
                    marAddr.AddressId = row.Item(colIndex).ToString
                End If
            Next

            pList.Items.Add(marAddr)
        Next

        pList.DisplayMember = "FullAddress"
        'pList.Sorted = True
    End Sub

    Private Sub chkSplitAddr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSplitAddr.Click
        Dim bValidAddress As Boolean = m_bAddressValidated
        If chkSplitAddr.Checked Then
            ParseAddress()
        Else
            '**removed & " " & ComboQUADRANT.Text from statement below
            textAddr.Text = Trim(TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text)
            If TextUNIT.Text <> "" Then
                textAddr.Text = textAddr.Text & " #" & TextUNIT.Text
            End If

            TextAddrnum.Text = ""
            TextADDRNUMSUF.Text = ""
            TextSTNAME.Text = ""
            ComboSTREET_TYPE.SelectedIndex = -1
            'ComboQUADRANT.SelectedIndex = -1
            TextUNIT.Text = ""

            bValidAddress = True
        End If

        ResetAddressComponents()
        m_bAddressValidated = bValidAddress
    End Sub

    Private Sub TextAddrnum_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextAddrnum.Validating
        e.Cancel = (Not MapUtil.IsNumericValue(TextAddrnum.Text))
    End Sub

    Private Sub TextZip5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextZip5.Validating
        e.Cancel = (Not MapUtil.IsNumericValue(TextZip5.Text))
    End Sub

    Private Sub TextZip4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextZip4.Validating
        e.Cancel = (Not MapUtil.IsNumericValue(TextZip4.Text))
    End Sub

    Private Sub ComboPexPtyp_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboPexPtyp.SelectedIndexChanged
        If Me.EditMode = 1 Or MapUtil.GetValue(m_premsPt, "PEXPSTS") = "PN" Then
            Dim pexptyp As clsGisDomain = DirectCast(ComboPexPtyp.SelectedItem, clsGisDomain)
            If UCase(pexptyp.Value) = "INTER" Then
                MapUtil.SetComboIndex2(Me.ComboIAB_EXEMPT, "Y")
                MapUtil.SetComboIndex2(Me.ComboPexSts, "ST")
                MapUtil.SetComboIndex2(Me.ComboEXEMPT_REASON, "INTER")
                MapUtil.SetComboIndex2(Me.ComboIA_ONLY, "N")
            ElseIf UCase(pexptyp.Value) = "TMPMT" Then
                MapUtil.SetComboIndex2(Me.ComboIAB_EXEMPT, "Y")
                MapUtil.SetComboIndex2(Me.ComboEXEMPT_REASON, "TMPMT")
                MapUtil.SetComboIndex2(Me.ComboPexSts, "PN")
                MapUtil.SetComboIndex2(Me.ComboIA_ONLY, "N")
            Else
                MapUtil.SetComboIndex2(Me.ComboPexSts, "PN")
                MapUtil.SetComboIndex2(Me.ComboIAB_EXEMPT, "Y")
                MapUtil.SetComboIndex2(Me.ComboEXEMPT_REASON, "PENDG")
            End If
        End If
    End Sub

    Private Sub btnAdvanced_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdvanced.Click
        Dim formAdv As FormPremiseAttributeAdv = New FormPremiseAttributeAdv
        formAdv.PremiseForm = Me 
        'formAdv.PremisePt = m_premsPt
        'Set firm values

        Dim fields As IFields
        fields = m_premsPt.Fields

        formAdv.txtADDRESS_ID.Text = address_id
        formAdv.txtMASTER_PEXUID.Text = master_pexuid
        formAdv.txtCOMNT.Text = comnt

        MapUtil.PopulateDomain(formAdv.cmbLOCTNPRECS, fields.Field(fields.FindField("LOCTNPRECS")).Domain, True, loctnprecs)
        MapUtil.PopulateDomain(formAdv.cmbINFORSRC, fields.Field(fields.FindField("INFORSRC")).Domain, True, inforsrc)

        If from_standalone = "Y" Then
            formAdv.chkFROM_STANDALONE.Checked = True
        Else
            formAdv.chkFROM_STANDALONE.Checked = False
        End If

        If has_lien = "Y" Then
            formAdv.chkHAS_LIEN.Checked = True
        Else
            formAdv.chkHAS_LIEN.Checked = False
        End If

        formAdv.Show(Me)
    End Sub
End Class
