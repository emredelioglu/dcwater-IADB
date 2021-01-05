<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PremiseAttribute
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnGetMAR = New System.Windows.Forms.Button
        Me.btnGetDDotAddr = New System.Windows.Forms.Button
        Me.btnValidateAddr = New System.Windows.Forms.Button
        Me.textAddr = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnGetSSLWard = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextAddrnum = New System.Windows.Forms.TextBox
        Me.TextADDRNUMSUF = New System.Windows.Forms.TextBox
        Me.TextSTNAME = New System.Windows.Forms.TextBox
        Me.lbl1 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextUNIT = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TextZip5 = New System.Windows.Forms.TextBox
        Me.TextZip4 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.TextSQUARE = New System.Windows.Forms.TextBox
        Me.TextSUFFIX = New System.Windows.Forms.TextBox
        Me.TextLOT = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.chkSplitAddr = New System.Windows.Forms.CheckBox
        Me.TextPexprm = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboSTREET_TYPE = New IAIS.QuietComboBox(Me.components)
        Me.ComboQUADRANT = New IAIS.QuietComboBox(Me.components)
        Me.ComboUSECODE = New IAIS.QuietComboBox(Me.components)
        Me.ComboIA_ONLY = New IAIS.QuietComboBox(Me.components)
        Me.ComboEXEMPT_REASON = New IAIS.QuietComboBox(Me.components)
        Me.ComboIAB_EXEMPT = New IAIS.QuietComboBox(Me.components)
        Me.ComboWARD = New IAIS.QuietComboBox(Me.components)
        Me.ComboPexPtyp = New IAIS.QuietComboBox(Me.components)
        Me.ComboPexSts = New IAIS.QuietComboBox(Me.components)
        Me.btnAdvanced = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(23, 401)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(132, 401)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnGetMAR
        '
        Me.btnGetMAR.Location = New System.Drawing.Point(23, 227)
        Me.btnGetMAR.Name = "btnGetMAR"
        Me.btnGetMAR.Size = New System.Drawing.Size(90, 43)
        Me.btnGetMAR.TabIndex = 13
        Me.btnGetMAR.Text = "Get MAR Address"
        Me.btnGetMAR.UseVisualStyleBackColor = True
        '
        'btnGetDDotAddr
        '
        Me.btnGetDDotAddr.Location = New System.Drawing.Point(134, 227)
        Me.btnGetDDotAddr.Name = "btnGetDDotAddr"
        Me.btnGetDDotAddr.Size = New System.Drawing.Size(90, 43)
        Me.btnGetDDotAddr.TabIndex = 14
        Me.btnGetDDotAddr.Text = "Get DDOT Address"
        Me.btnGetDDotAddr.UseVisualStyleBackColor = True
        '
        'btnValidateAddr
        '
        Me.btnValidateAddr.Location = New System.Drawing.Point(242, 227)
        Me.btnValidateAddr.Name = "btnValidateAddr"
        Me.btnValidateAddr.Size = New System.Drawing.Size(90, 43)
        Me.btnValidateAddr.TabIndex = 15
        Me.btnValidateAddr.Text = "Validate Address"
        Me.btnValidateAddr.UseVisualStyleBackColor = True
        '
        'textAddr
        '
        Me.textAddr.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.textAddr.Location = New System.Drawing.Point(100, 192)
        Me.textAddr.Name = "textAddr"
        Me.textAddr.Size = New System.Drawing.Size(229, 20)
        Me.textAddr.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 195)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Full Address"
        '
        'btnGetSSLWard
        '
        Me.btnGetSSLWard.Location = New System.Drawing.Point(186, 329)
        Me.btnGetSSLWard.Name = "btnGetSSLWard"
        Me.btnGetSSLWard.Size = New System.Drawing.Size(143, 23)
        Me.btnGetSSLWard.TabIndex = 21
        Me.btnGetSSLWard.Text = "Get SSL/Ward"
        Me.btnGetSSLWard.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Premise Status"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Premise Type"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "House #"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(84, 103)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(22, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Frc"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(131, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(35, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Street"
        '
        'TextAddrnum
        '
        Me.TextAddrnum.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextAddrnum.Location = New System.Drawing.Point(26, 120)
        Me.TextAddrnum.Name = "TextAddrnum"
        Me.TextAddrnum.Size = New System.Drawing.Size(45, 20)
        Me.TextAddrnum.TabIndex = 3
        '
        'TextADDRNUMSUF
        '
        Me.TextADDRNUMSUF.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextADDRNUMSUF.Location = New System.Drawing.Point(82, 120)
        Me.TextADDRNUMSUF.Name = "TextADDRNUMSUF"
        Me.TextADDRNUMSUF.Size = New System.Drawing.Size(40, 20)
        Me.TextADDRNUMSUF.TabIndex = 4
        '
        'TextSTNAME
        '
        Me.TextSTNAME.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextSTNAME.Location = New System.Drawing.Point(128, 120)
        Me.TextSTNAME.Name = "TextSTNAME"
        Me.TextSTNAME.Size = New System.Drawing.Size(111, 20)
        Me.TextSTNAME.TabIndex = 5
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Location = New System.Drawing.Point(253, 103)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(62, 13)
        Me.lbl1.TabIndex = 19
        Me.lbl1.Text = "Street Type"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(26, 145)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(49, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Direction"
        '
        'TextUNIT
        '
        Me.TextUNIT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextUNIT.Location = New System.Drawing.Point(82, 164)
        Me.TextUNIT.Name = "TextUNIT"
        Me.TextUNIT.Size = New System.Drawing.Size(40, 20)
        Me.TextUNIT.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(82, 145)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(30, 13)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "Apt#"
        '
        'TextZip5
        '
        Me.TextZip5.Location = New System.Drawing.Point(128, 164)
        Me.TextZip5.MaxLength = 5
        Me.TextZip5.Name = "TextZip5"
        Me.TextZip5.Size = New System.Drawing.Size(45, 20)
        Me.TextZip5.TabIndex = 9
        '
        'TextZip4
        '
        Me.TextZip4.Location = New System.Drawing.Point(193, 164)
        Me.TextZip4.MaxLength = 4
        Me.TextZip4.Name = "TextZip4"
        Me.TextZip4.Size = New System.Drawing.Size(35, 20)
        Me.TextZip4.TabIndex = 10
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(179, 167)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(10, 13)
        Me.Label10.TabIndex = 26
        Me.Label10.Text = "-"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(128, 145)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(50, 13)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Zip Code"
        '
        'TextSQUARE
        '
        Me.TextSQUARE.Location = New System.Drawing.Point(24, 293)
        Me.TextSQUARE.Name = "TextSQUARE"
        Me.TextSQUARE.Size = New System.Drawing.Size(51, 20)
        Me.TextSQUARE.TabIndex = 16
        '
        'TextSUFFIX
        '
        Me.TextSUFFIX.Location = New System.Drawing.Point(82, 293)
        Me.TextSUFFIX.Name = "TextSUFFIX"
        Me.TextSUFFIX.Size = New System.Drawing.Size(58, 20)
        Me.TextSUFFIX.TabIndex = 17
        '
        'TextLOT
        '
        Me.TextLOT.Location = New System.Drawing.Point(147, 293)
        Me.TextLOT.Name = "TextLOT"
        Me.TextLOT.Size = New System.Drawing.Size(60, 20)
        Me.TextLOT.TabIndex = 18
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(23, 273)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(41, 13)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Square"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(86, 273)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(33, 13)
        Me.Label13.TabIndex = 33
        Me.Label13.Text = "Suffix"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(148, 273)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(22, 13)
        Me.Label14.TabIndex = 34
        Me.Label14.Text = "Lot"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(247, 273)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(33, 13)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "Ward"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(23, 355)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(89, 13)
        Me.Label16.TabIndex = 37
        Me.Label16.Text = "Exemption Status"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(119, 355)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(96, 13)
        Me.Label17.TabIndex = 38
        Me.Label17.Text = "Exemption Reason"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(247, 355)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(82, 13)
        Me.Label18.TabIndex = 40
        Me.Label18.Text = "Impervious Only"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(21, 314)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(54, 13)
        Me.Label19.TabIndex = 42
        Me.Label19.Text = "Use Code"
        '
        'chkSplitAddr
        '
        Me.chkSplitAddr.AutoSize = True
        Me.chkSplitAddr.Location = New System.Drawing.Point(242, 166)
        Me.chkSplitAddr.Name = "chkSplitAddr"
        Me.chkSplitAddr.Size = New System.Drawing.Size(87, 17)
        Me.chkSplitAddr.TabIndex = 11
        Me.chkSplitAddr.Text = "Split Address"
        Me.chkSplitAddr.UseVisualStyleBackColor = True
        '
        'TextPexprm
        '
        Me.TextPexprm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextPexprm.Enabled = False
        Me.TextPexprm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextPexprm.Location = New System.Drawing.Point(128, 13)
        Me.TextPexprm.Name = "TextPexprm"
        Me.TextPexprm.Size = New System.Drawing.Size(140, 20)
        Me.TextPexprm.TabIndex = 47
        Me.TextPexprm.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(23, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(84, 13)
        Me.Label7.TabIndex = 48
        Me.Label7.Text = "Premise Number"
        '
        'ComboSTREET_TYPE
        '
        Me.ComboSTREET_TYPE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboSTREET_TYPE.FormattingEnabled = True
        Me.ComboSTREET_TYPE.Items.AddRange(New Object() {"", "ALY", "ANX", "ARC", "AVE", "BCH", "BG", "BGS", "BLF", "BLFS", "BLVD", "BND", "BR", "BRG", "BRK", "BRKS", "BTM", "BYP", "BYU", "CIR", "CIRS", "CLB", "CLF", "CLFS", "CMN", "COR", "CORS", "CP", "CPE", "CRES", "CRK", "CRSE", "CRST", "CSWY", "CT", "CTR", "CTRS", "CTS", "CURV", "CV", "CVS", "CYN", "DL", "DM", "DR", "DRS", "DV", "EST", "ESTS", "EXPY", "EXT", "EXTS", "FALL", "FLD", "FLDS", "FLS", "FLT", "FLTS", "FRD", "FRDS", "FRG", "FRGS", "FRK", "FRKS", "FRST", "FRY", "FT", "FWY", "GDN", "GDNS", "GLN", "GLNS", "GRN", "GRNS", "GRV", "GRVS", "GTWY", "HBR", "HBRS", "HL", "HLS", "HOLW", "HTS", "HVN", "HWY", "INLT", "IS", "ISLE", "ISS", "JCT", "JCTS", "KNL", "KNLS", "KY", "KYS", "LAND", "LCK", "LCKS", "LDG", "LF", "LGT", "LGTS", "LK", "LKS", "LN", "LNDG", "LOOP", "MALL", "MDW", "MDWS", "MEWS", "ML", "MLS", "MNR", "MNRS", "MSN", "MT", "MTN", "MTNS", "MTWY", "NCK", "OPAS", "ORCH", "OVAL", "PARK", "PASS", "PATH", "PIKE", "PKWY", "PL", "PLN", "PLNS", "PLZ", "PNE", "PNES", "PR", "PRT", "PRTS", "PSGE", "PT", "PTS", "RADL", "RAMP", "RD", "RDG", "RDGS", "RDS", "RIV", "RNCH", "ROW", "RPD", "RPDS", "RST", "RTE", "RUE", "RUN", "SHL", "SHLS", "SHR", "SHRS", "SKWY", "SMT", "SPG", "SPGS", "SPUR", "SQ", "SQS", "ST", "STA", "STRA", "STRM", "STS", "TER", "TPKE", "TRAK", "TRCE", "TRFY", "TRL", "TRWY", "TUNL", "UN", "UNS", "UPAS", "VIA", "VIS", "VL", "VLG", "VLGS", "VLY", "VLYS", "VW", "VWS", "WALK", "WALL", "WAY", "WAYS", "WL", "WLS", "XING", "XRD"})
        Me.ComboSTREET_TYPE.Location = New System.Drawing.Point(245, 120)
        Me.ComboSTREET_TYPE.Name = "ComboSTREET_TYPE"
        Me.ComboSTREET_TYPE.Size = New System.Drawing.Size(84, 21)
        Me.ComboSTREET_TYPE.TabIndex = 6
        '
        'ComboQUADRANT
        '
        Me.ComboQUADRANT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboQUADRANT.FormattingEnabled = True
        Me.ComboQUADRANT.Location = New System.Drawing.Point(26, 164)
        Me.ComboQUADRANT.Name = "ComboQUADRANT"
        Me.ComboQUADRANT.Size = New System.Drawing.Size(45, 21)
        Me.ComboQUADRANT.TabIndex = 7
        '
        'ComboUSECODE
        '
        Me.ComboUSECODE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboUSECODE.FormattingEnabled = True
        Me.ComboUSECODE.Items.AddRange(New Object() {"", "001", "002", "003", "004", "005", "006", "007", "008", "011", "012", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023", "024", "025", "026", "027", "028", "029", "031", "032", "033", "034", "035", "036", "037", "038", "039", "041", "042", "043", "044", "045", "046", "047", "048", "049", "051", "052", "053", "056", "057", "058", "059", "061", "062", "063", "064", "065", "066", "067", "068", "069", "071", "072", "073", "074", "075", "076", "078", "079", "081", "082", "083", "084", "085", "086", "087", "088", "089", "091", "092", "093", "094", "095", "096", "097", "116", "117", "126", "127", "165", "189", "214", "216", "217", "265", "316", "365", "416", "417", "465", "516"})
        Me.ComboUSECODE.Location = New System.Drawing.Point(24, 330)
        Me.ComboUSECODE.Name = "ComboUSECODE"
        Me.ComboUSECODE.Size = New System.Drawing.Size(82, 21)
        Me.ComboUSECODE.TabIndex = 20
        '
        'ComboIA_ONLY
        '
        Me.ComboIA_ONLY.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboIA_ONLY.FormattingEnabled = True
        Me.ComboIA_ONLY.Location = New System.Drawing.Point(250, 371)
        Me.ComboIA_ONLY.Name = "ComboIA_ONLY"
        Me.ComboIA_ONLY.Size = New System.Drawing.Size(79, 21)
        Me.ComboIA_ONLY.TabIndex = 24
        '
        'ComboEXEMPT_REASON
        '
        Me.ComboEXEMPT_REASON.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboEXEMPT_REASON.FormattingEnabled = True
        Me.ComboEXEMPT_REASON.Location = New System.Drawing.Point(122, 372)
        Me.ComboEXEMPT_REASON.Name = "ComboEXEMPT_REASON"
        Me.ComboEXEMPT_REASON.Size = New System.Drawing.Size(110, 21)
        Me.ComboEXEMPT_REASON.TabIndex = 23
        '
        'ComboIAB_EXEMPT
        '
        Me.ComboIAB_EXEMPT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboIAB_EXEMPT.FormattingEnabled = True
        Me.ComboIAB_EXEMPT.Location = New System.Drawing.Point(23, 372)
        Me.ComboIAB_EXEMPT.Name = "ComboIAB_EXEMPT"
        Me.ComboIAB_EXEMPT.Size = New System.Drawing.Size(75, 21)
        Me.ComboIAB_EXEMPT.TabIndex = 22
        '
        'ComboWARD
        '
        Me.ComboWARD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboWARD.FormattingEnabled = True
        Me.ComboWARD.ItemHeight = 13
        Me.ComboWARD.Location = New System.Drawing.Point(246, 292)
        Me.ComboWARD.Name = "ComboWARD"
        Me.ComboWARD.Size = New System.Drawing.Size(79, 21)
        Me.ComboWARD.TabIndex = 19
        '
        'ComboPexPtyp
        '
        Me.ComboPexPtyp.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboPexPtyp.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboPexPtyp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboPexPtyp.FormattingEnabled = True
        Me.ComboPexPtyp.Location = New System.Drawing.Point(128, 76)
        Me.ComboPexPtyp.Name = "ComboPexPtyp"
        Me.ComboPexPtyp.Size = New System.Drawing.Size(140, 21)
        Me.ComboPexPtyp.TabIndex = 2
        '
        'ComboPexSts
        '
        Me.ComboPexSts.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboPexSts.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboPexSts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboPexSts.FormattingEnabled = True
        Me.ComboPexSts.Items.AddRange(New Object() {"Pending", "Street Intersection", "Purged"})
        Me.ComboPexSts.Location = New System.Drawing.Point(128, 43)
        Me.ComboPexSts.Name = "ComboPexSts"
        Me.ComboPexSts.Size = New System.Drawing.Size(140, 21)
        Me.ComboPexSts.TabIndex = 1
        '
        'btnAdvanced
        '
        Me.btnAdvanced.Location = New System.Drawing.Point(250, 401)
        Me.btnAdvanced.Name = "btnAdvanced"
        Me.btnAdvanced.Size = New System.Drawing.Size(75, 23)
        Me.btnAdvanced.TabIndex = 49
        Me.btnAdvanced.Text = "Advanced"
        Me.btnAdvanced.UseVisualStyleBackColor = True
        '
        'PremiseAttribute
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(349, 436)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnAdvanced)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TextPexprm)
        Me.Controls.Add(Me.chkSplitAddr)
        Me.Controls.Add(Me.ComboSTREET_TYPE)
        Me.Controls.Add(Me.ComboQUADRANT)
        Me.Controls.Add(Me.ComboUSECODE)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.ComboIA_ONLY)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.ComboEXEMPT_REASON)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.ComboIAB_EXEMPT)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.ComboWARD)
        Me.Controls.Add(Me.TextLOT)
        Me.Controls.Add(Me.TextSUFFIX)
        Me.Controls.Add(Me.TextSQUARE)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TextZip4)
        Me.Controls.Add(Me.TextZip5)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.TextUNIT)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.TextSTNAME)
        Me.Controls.Add(Me.TextADDRNUMSUF)
        Me.Controls.Add(Me.TextAddrnum)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboPexPtyp)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboPexSts)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnGetSSLWard)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.textAddr)
        Me.Controls.Add(Me.btnValidateAddr)
        Me.Controls.Add(Me.btnGetDDotAddr)
        Me.Controls.Add(Me.btnGetMAR)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PremiseAttribute"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Premise Attributes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnGetMAR As System.Windows.Forms.Button
    Friend WithEvents btnGetDDotAddr As System.Windows.Forms.Button
    Friend WithEvents btnValidateAddr As System.Windows.Forms.Button
    Friend WithEvents textAddr As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnGetSSLWard As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboPexSts As IAIS.QuietComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboPexPtyp As IAIS.QuietComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextAddrnum As System.Windows.Forms.TextBox
    Friend WithEvents TextADDRNUMSUF As System.Windows.Forms.TextBox
    Friend WithEvents TextSTNAME As System.Windows.Forms.TextBox
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextUNIT As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextZip5 As System.Windows.Forms.TextBox
    Friend WithEvents TextZip4 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextSQUARE As System.Windows.Forms.TextBox
    Friend WithEvents TextSUFFIX As System.Windows.Forms.TextBox
    Friend WithEvents TextLOT As System.Windows.Forms.TextBox
    Friend WithEvents ComboWARD As IAIS.QuietComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents ComboIAB_EXEMPT As IAIS.QuietComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ComboEXEMPT_REASON As IAIS.QuietComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents ComboIA_ONLY As IAIS.QuietComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents ComboUSECODE As IAIS.QuietComboBox
    Friend WithEvents ComboQUADRANT As IAIS.QuietComboBox
    Friend WithEvents ComboSTREET_TYPE As IAIS.QuietComboBox
    Friend WithEvents chkSplitAddr As System.Windows.Forms.CheckBox
    Friend WithEvents TextPexprm As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnAdvanced As System.Windows.Forms.Button

End Class
