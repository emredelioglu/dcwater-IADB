Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.ArcMapUI

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
        If chkSplitAddr.Checked Then
            If Trim(TextAddrnum.Text) = "" Then
                MsgBox("House Number is required")
                Return False
            End If
            If Trim(TextSTNAME.Text) = "" Then
                MsgBox("Street Name is required")
                Return False
            End If
            If ComboSTREET_TYPE.SelectedItem Is Nothing Then
                MsgBox("Street Type is required")
                Return False
            End If

            If Trim(TextZip5.Text) = "" Then
                MsgBox("Zip Code is required")
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
        ElseIf UCase(ComboIAB_EXEMPT.SelectedItem) = "YES" AndAlso ComboEXEMPT_REASON.SelectedItem Is Nothing Then
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
        m_Editor.StartOperation()
        Try
            updateFeature()
            m_premsPt.Store()
            If EditMode = 1 Then
                m_Editor.StopOperation("Create Premise Point")
            Else
                m_Editor.StopOperation("Update Premise Point")
            End If
        Catch ex As Exception
            m_Editor.AbortOperation()
            MsgBox(ex.Message)
        End Try
        RaiseEvent CommandButtonClicked("Ok")
        Me.Close()
    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_premiseResult = 0
        'remove the point if edit mode=1 (new premise point)
        If m_editMode = 1 Then
            'm_premsPt.Delete()
            Dim pUid As New UID
            pUid.Value = "esriEditor.Editor"
            Dim m_Editor As IEditor = m_app.FindExtensionByCLSID(pUid)

            m_Editor.StartOperation()
            Try
                m_premsPt.Delete()
                m_Editor.StopOperation("Abort Premise Point")
            Catch ex As Exception
                m_Editor.AbortOperation()
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
        RaiseEvent CommandButtonClicked("GetMAR")
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub btnValidateAddr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidateAddr.Click
        Dim MarServiceResult As gov.dc.citizenatlas.ReturnObject
        Dim score As Double
        Dim LocationVerifier As gov.dc.citizenatlas.LocationVerifier = New gov.dc.citizenatlas.LocationVerifier()
        Try
            If Trim(TextAddrnum.Text) = "" Then
                MsgBox("Address Number is required")
                Return
            End If

            If Trim(TextSTNAME.Text) = "" Then
                MsgBox("Street Name is required")
                Return
            End If

            Dim fulladdress As String = TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text

            MarServiceResult = LocationVerifier.verifyDCAddress(TextAddrnum.Text & " " & TextADDRNUMSUF.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text)
            'MarServiceResult = LocationVerifier.verifyDCAddress(textAddr.Text)
            If MarServiceResult.returnDataset Is Nothing Then
                MarServiceResult = LocationVerifier.verifyDCAddress(TextAddrnum.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text)
                If MarServiceResult.returnDataset Is Nothing Then
                    MsgBox("Can't validate this address with MAR web service")
                    Return
                Else
                    TextADDRNUMSUF.Text = ""
                    TextUNIT.Text = TextADDRNUMSUF.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return
        End Try


        MapUtil.PrintDataSet(MarServiceResult.returnDataset)

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

                MsgBox("Address validated.")
                formMar.Dispose()
                Return
            End If
        End If

        formMar.ShowDialog()
        If formMar.DialogResult = System.Windows.Forms.DialogResult.Cancel Then
            Return
        End If

        If formMar.ListMARAddress.SelectedItem Is Nothing Then
            Return
        End If


        marAddra = formMar.ListMARAddress.SelectedItem

        Me.TextAddrnum.Text = marAddra.AddrNum
        Me.TextSTNAME.Text = marAddra.StName
        Me.TextADDRNUMSUF.Text = marAddra.AddrNumSuf
        Me.TextZip5.Text = marAddra.ZipCode
        MapUtil.SetComboIndex(Me.ComboSTREET_TYPE, marAddra.StreetType)
        MapUtil.SetComboIndex(Me.ComboQUADRANT, marAddra.Quadrant)

        m_bAddressValidated = True
        m_bValidAddress = True

        formMar.Dispose()
    End Sub

    Private Sub updateAttribute()
        'PopulateDomain
        Dim fields As IFields
        fields = m_premsPt.Fields

        TextPexprm.Text = MapUtil.GetValue(m_premsPt, "PEXPRM")

        Dim pexpsts As String = UCase(MapUtil.GetValue(m_premsPt, "PEXPSTS"))
        If pexpsts = "IA" Then
            ComboPexSts.Enabled = True
            ComboPexSts.Items.Clear()
            ComboPexSts.Items.Add("Inactive")
            ComboPexSts.Items.Add("Purged")
        Else
            ComboPexSts.Enabled = False
        End If

        If pexpsts = "AC" Then
            ComboPexPtyp.Enabled = False
        End If

        MapUtil.SetComboIndex(Me.ComboPexSts, MapUtil.GetValue(m_premsPt, "PEXPSTS", True))
        MapUtil.PopulateDomain(Me.ComboPexPtyp, fields.Field(fields.FindField("PEXPTYP")).Domain, False, MapUtil.GetValue(m_premsPt, "PEXPTYP"))

        Me.TextAddrnum.Text = MapUtil.GetValue(m_premsPt, "ADDRNUM")
        Me.TextSTNAME.Text = MapUtil.GetValue(m_premsPt, "STNAME")

        Me.TextUNIT.Text = MapUtil.GetValue(m_premsPt, "UNIT")

        Me.textAddr.Text = MapUtil.GetValue(m_premsPt, "PEXSAD")

        MapUtil.SetComboIndex(Me.ComboSTREET_TYPE, MapUtil.GetValue(m_premsPt, "STREET_TYPE"))

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
        Me.TextZip4.Text = Mid(pexzip, 5, 5)

        MapUtil.PopulateDomain(Me.ComboQUADRANT, fields.Field(fields.FindField("QUADRANT")).Domain, True, MapUtil.GetValue(m_premsPt, "QUADRANT"))

        Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
        Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
        Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")

        MapUtil.SetComboIndex(Me.ComboWARD, MapUtil.GetValue(m_premsPt, "WARD"))

        MapUtil.PopulateDomain(Me.ComboIAB_EXEMPT, fields.Field(fields.FindField("IS_EXEMPT_IAB")).Domain, False, MapUtil.GetValue(m_premsPt, "IS_EXEMPT_IAB"))
        If MapUtil.GetValue(m_premsPt, "IS_EXEMPT_IAB") = "Y" Then
            Me.ComboEXEMPT_REASON.Enabled = True
        Else
            Me.ComboEXEMPT_REASON.Enabled = False
        End If

        MapUtil.PopulateDomain(Me.ComboEXEMPT_REASON, fields.Field(fields.FindField("EXEMPT_IAB_REASON")).Domain, False, MapUtil.GetValue(m_premsPt, "EXEMPT_IAB_REASON"))
        MapUtil.PopulateDomain(Me.ComboIA_ONLY, fields.Field(fields.FindField("IS_IMPERVIOUS_ONLY")).Domain, False, MapUtil.GetValue(m_premsPt, "IS_IMPERVIOUS_ONLY"))

        MapUtil.SetComboIndex(Me.ComboUSECODE, MapUtil.GetValue(m_premsPt, "OTRUSECODE"))

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

    End Sub


    Private Sub updateFeature()
        'PopulateDomain
        Dim fields As IFields
        fields = m_premsPt.Fields

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXPSTS"), Me.ComboPexSts.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXPTYP"), Me.ComboPexPtyp.Text)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("ADDRNUM"), Me.TextAddrnum.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("STNAME"), Me.TextSTNAME.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("STREET_TYPE"), GetUSPSStType(Me.ComboSTREET_TYPE.Text))

        If chkSplitAddr.Checked Then
            MapUtil.SetFeatureValue(m_premsPt, fields.FindField("USE_PARSED_ADDRESS"), "Y", True)
        End If

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("UNIT"), Me.TextUNIT.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXSAD"), Me.textAddr.Text)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXPZIP"), Me.TextZip5.Text & Me.TextZip4.Text)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("QUADRANT"), Me.ComboQUADRANT.SelectedItem)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXSQUARE"), Me.TextSQUARE.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXSUFFIX"), Me.TextSUFFIX.Text)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("PEXLOT"), Me.TextLOT.Text)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("WARD"), Me.ComboWARD.SelectedItem)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("IS_EXEMPT_IAB"), Me.ComboIAB_EXEMPT.SelectedItem)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("EXEMPT_IAB_REASON"), Me.ComboEXEMPT_REASON.SelectedItem)
        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("IS_IMPERVIOUS_ONLY"), Me.ComboIA_ONLY.SelectedItem)

        MapUtil.SetFeatureValue(m_premsPt, fields.FindField("OTRUSECODE"), Me.ComboUSECODE.Text)

        m_premsPt.Value(fields.FindField("XCOOR")) = m_premsPt.Shape.Envelope.XMax
        m_premsPt.Value(fields.FindField("YCOOR")) = m_premsPt.Shape.Envelope.YMax

        m_premsPt.Value(fields.FindField("PEXUID")) = m_premsPt.OID

        'Parse address?

    End Sub

    Private Sub ComboIAB_EXEMPT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboIAB_EXEMPT.SelectedIndexChanged
        Debug.Print("ComboIAB_EXEMPT_SelectedIndexChanged: " & ComboIAB_EXEMPT.SelectedText)
        If UCase(ComboIAB_EXEMPT.SelectedItem) = "YES" Then
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
                MapUtil.SetComboIndex(Me.ComboWARD, MapUtil.GetValue(m_premsPt, "WARD"))
            End If
        End If

        colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"), m_map)
        If colLayer Is Nothing Then
            colLayer = MapUtil.GetLayerFromWS(pDataset.Workspace, _
                                                         IAISToolSetting.GetParameterValue("LAYER_OWNERPLY"))
        End If
        If Not colLayer Is Nothing Then
            pFeatCursor = colLayer.Search(pSFilter, False)
            pFeature = pFeatCursor.NextFeature
            If Not pFeature Is Nothing Then
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSQUARE"), MapUtil.GetValue(pFeature, "SQUARE"))
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSUFFIX"), MapUtil.GetValue(pFeature, "SUFFIX"))
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXLOT"), MapUtil.GetValue(pFeature, "LOT"))
                MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("OTRUSECODE"), MapUtil.GetValue(pFeature, "USECODE"))

                Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
                Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
                Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")
                MapUtil.SetComboIndex(ComboUSECODE, MapUtil.GetValue(m_premsPt, "OTRUSECODE"))

                Return
            End If
        End If

        colLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"), m_map)
        If colLayer Is Nothing Then
            colLayer = MapUtil.GetLayerFromWS(pDataset.Workspace, _
                                                         IAISToolSetting.GetParameterValue("LAYER_OWNERGAPPLY"))
        End If

        pFeatCursor = colLayer.Search(pSFilter, False)
        pFeature = pFeatCursor.NextFeature
        If Not pFeature Is Nothing Then
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSQUARE"), MapUtil.GetValue(pFeature, "SQUARE"))
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXSUFFIX"), MapUtil.GetValue(pFeature, "SUFFIX"))
            MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("PEXLOT"), MapUtil.GetValue(pFeature, "LOT"))

            Me.TextSQUARE.Text = MapUtil.GetValue(m_premsPt, "PEXSQUARE")
            Me.TextSUFFIX.Text = MapUtil.GetValue(m_premsPt, "PEXSUFFIX")
            Me.TextLOT.Text = MapUtil.GetValue(m_premsPt, "PEXLOT")

        End If

    End Sub

    Private Sub btnGetDDotAddr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetDDotAddr.Click
        RaiseEvent CommandButtonClicked("GetDDotAddr")
    End Sub

    Private Sub textAddr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textAddr.TextChanged
        m_bAddressValidated = False
        m_bValidAddress = False
    End Sub

    Public Sub SetDDotAddress(ByVal fulladdr As String)
        textAddr.Text = fulladdr
        textAddr.Enabled = True

        TextAddrnum.Enabled = True
        TextSTNAME.Enabled = True

        ComboSTREET_TYPE.Enabled = True
        ComboQUADRANT.Enabled = True
        TextADDRNUMSUF.Enabled = True

        m_bAddressValidated = False
        btnValidateAddr.Enabled = True

        ParseAddress()
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
        MapUtil.SetFeatureValue(m_premsPt, m_premsPt.Fields.FindField("LOCTNPRECS"), "INFERRED - MAR")

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

        Dim lAddrNum As Long
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
            strUnit = Mid(FullAddress, 1, pos)
            FullAddress = Trim(Mid(FullAddress, 0, pos))
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
        lAddrNum = Mid(FullAddress, 1, pos)


        FullAddress = Trim(Mid(FullAddress, pos))
        strStName = FullAddress

        TextAddrnum.Text = lAddrNum
        TextSTNAME.Text = strStName
        MapUtil.SetComboIndex(ComboSTREET_TYPE, strStTYPE)
        MapUtil.SetComboIndex(ComboQUADRANT, strQuad)

        TextADDRNUMSUF.Text = sAddrSufx
        TextUNIT.Text = strUnit

        chkSplitAddr.Checked = True
        textAddr.Text = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSplitAddr.CheckedChanged
        Dim bValidAddress As Boolean = m_bAddressValidated
        If chkSplitAddr.Checked Then
            ParseAddress()
        Else
            textAddr.Text = Trim(TextAddrnum.Text & " " & TextSTNAME.Text & " " & ComboSTREET_TYPE.Text & " " & ComboQUADRANT.Text)
            If TextUNIT.Text <> "" Then
                textAddr.Text = textAddr.Text & " #" & TextUNIT.Text
            End If

            TextAddrnum.Text = ""
            TextADDRNUMSUF.Text = ""
            TextSTNAME.Text = ""
            ComboSTREET_TYPE.SelectedIndex = -1
            ComboQUADRANT.SelectedIndex = -1
            TextUNIT.Text = ""

        End If

        ResetAddressComponents()
        m_bAddressValidated = bValidAddress
    End Sub

    Private Sub ResetAddressComponents()

        If chkSplitAddr.Checked Then
            textAddr.Enabled = False

            TextAddrnum.Enabled = True
            TextSTNAME.Enabled = True
            ComboSTREET_TYPE.Enabled = True
            ComboQUADRANT.Enabled = True
            TextADDRNUMSUF.Enabled = True
            TextUNIT.Enabled = True

            btnValidateAddr.Enabled = True
        Else
            textAddr.Enabled = True

            TextAddrnum.Enabled = False
            TextSTNAME.Enabled = False
            ComboSTREET_TYPE.Enabled = False
            ComboQUADRANT.Enabled = False
            TextADDRNUMSUF.Enabled = False

            TextUNIT.Enabled = False

            btnValidateAddr.Enabled = False
        End If

    End Sub

    Private Sub PremiseAttribute_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Tab Then
            If e.Modifiers = Keys.Shift Then
                Me.ProcessTabKey(False)
            Else
                Me.ProcessTabKey(True)
            End If
        End If
    End Sub


    Private Sub PremiseAttribute_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Populate the street type
        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
        Dim pDataset As IDataset = premiseLayer
        Dim pTable As ITable = MapUtil.GetTableFromWS(pDataset.Workspace, "STREET_TYPE")

        Dim pTableSort As ITableSort = New TableSort
        With pTableSort
            .Fields = "STREET_TYPE"
            .Ascending("STREET_TYPE") = True
            .QueryFilter = Nothing
            .Table = pTable
        End With

        pTableSort.Sort(Nothing)
        Dim pCursor As ICursor = pTableSort.Rows

        Dim pRow As IRow
        pRow = pCursor.NextRow

        Me.ComboSTREET_TYPE.Items.Clear()

        Do While Not pRow Is Nothing
            Me.ComboSTREET_TYPE.Items.Add(pRow.Value(pRow.Fields.FindField("STREET_TYPE")))
            pRow = pCursor.NextRow
        Loop

    End Sub

    Private Function GetUSPSStType(ByVal sttype As String) As String
        If Trim(sttype) = "" Then
            Return ""
        End If

        Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
        Dim pDataset As IDataset = premiseLayer
        Dim pTable As ITable = MapUtil.GetTableFromWS(pDataset.Workspace, "STREET_TYPE")
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

    Private Sub TextAddrnum_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextAddrnum.KeyPress
        e.Handled = IsNumericKey(Asc(e.KeyChar))
    End Sub


    Private Function IsNumericKey(ByVal KCode As String) As Boolean
        If (KCode >= 48 And KCode <= 57) Or KCode = 8 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub TextZip5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextZip5.KeyPress
        e.Handled = IsNumericKey(Asc(e.KeyChar))
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

    Private Sub ComboPexSts_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboPexSts.DrawItem

    End Sub

    Private Sub ComboPexSts_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboPexSts.DropDown

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
                Debug.Print(table.Columns.Item(colIndex).ColumnName)

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

    Private Sub ComboPexSts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboPexSts.SelectedIndexChanged

    End Sub
End Class
