Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports IAIS.Windows.UI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Display
Imports System.Drawing

Public Class FormPremiseLocator

    Private m_app As IApplication
    Private m_map As IMap

    Dim m_form As FormPremiseList


    Dim b_searchFlag As Boolean
    Dim b_searchingFlag As Boolean

    Public WriteOnly Property PremiseApp() As IApplication
        Set(ByVal value As IApplication)
            m_app = value
        End Set
    End Property

    Public Property GMap() As IMap
        Get
            Return m_map
        End Get
        Set(ByVal value As IMap)
            m_map = value
        End Set
    End Property

    Private Function FindPremises(ByVal pexprm As String, ByVal pexact As String, ByVal ownername As String, ByVal zoompan As String) As IList

        pexprm = Trim(pexprm)
        pexact = Trim(pexact)
        ownername = Trim(ownername)

        Dim premiseLayer As IFeatureLayer
        Dim pCursor As IFeatureCursor
        Dim pFeature As IFeature = Nothing

        Dim pQFilter As IQueryFilter = New QueryFilter
        Dim pTableSort As ITableSort = New TableSort

        

        premiseLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)

        If zoompan <> "" Then
            If pexprm <> "" Then
                pQFilter.WhereClause = "PEXPRM = '" & pexprm & "'"
                pTableSort.Fields = "PEXPRM"
                pTableSort.Ascending("PEXPRM") = True
            ElseIf pexact <> "" Then
                pQFilter.WhereClause = "PEXACT = '" & pexact & "'"
                pTableSort.Fields = "PEXACT"
                pTableSort.Ascending("PEXACT") = True
            ElseIf ownername <> "" Then
                pQFilter.WhereClause = "UPPER(PEXNAM) LIKE '" & UCase(ownername) & "%'"
                pTableSort.Fields = "PEXNAM"
                pTableSort.Ascending("PEXNAM") = True
            Else
                MsgBox("All fields are empty.")
                Return Nothing
            End If

            pTableSort.QueryFilter = pQFilter
            pTableSort.Table = CType(premiseLayer.FeatureClass, ITable)
            pTableSort.Sort(Nothing)
            pCursor = pTableSort.Rows
            pFeature = pCursor.NextFeature

        End If



        If pFeature Is Nothing Then
            If pexprm <> "" Then
                pQFilter.WhereClause = "PEXPRM LIKE '" & pexprm & "%'"
                pTableSort.Fields = "PEXPRM"
                pTableSort.Ascending("PEXPRM") = True
            ElseIf pexact <> "" Then
                pQFilter.WhereClause = "PEXACT LIKE '" & pexact & "%'"
                pTableSort.Fields = "PEXACT"
                pTableSort.Ascending("PEXACT") = True
            ElseIf ownername <> "" Then
                pQFilter.WhereClause = "UPPER(PEXNAM) LIKE '" & UCase(ownername) & "%'"
                pTableSort.Fields = "PEXNAM"
                pTableSort.Ascending("PEXNAM") = True
            End If

            pTableSort.QueryFilter = pQFilter
            pTableSort.Table = CType(premiseLayer.FeatureClass, ITable)
            pTableSort.Sort(Nothing)
            pCursor = pTableSort.Rows
            pFeature = pCursor.NextFeature

        End If


        Dim premiseList As IList = New List(Of clsPremise)

        Dim premise As clsPremise
        Dim maxcount As Integer

        Try
            maxcount = CInt(IAISToolSetting.GetParameterValue("IAIS_MAX_RESULT_COUNT"))
        Catch ex As Exception
            maxcount = 2000
        End Try


        b_searchFlag = True

        Dim counter As Integer

        Do While Not pFeature Is Nothing AndAlso b_searchFlag AndAlso counter <= maxcount
            counter = counter + 1
            System.Windows.Forms.Application.DoEvents()

            premise = New clsPremise

            premise.Pexuid = MapUtil.GetValue(pFeature, "PEXUID")
            premise.Pexprm = MapUtil.GetValue(pFeature, "PEXPRM")
            premise.Pexsad = MapUtil.GetValue(pFeature, "PEXSAD")
            premise.Owner = MapUtil.GetValue(pFeature, "PEXNAM")
            premise.Pexact = MapUtil.GetValue(pFeature, "PEXACT")
            premise.Pexptyp = MapUtil.GetValue(pFeature, "PEXPTYP", True)
            premise.Pexpsts = MapUtil.GetValue(pFeature, "PEXPSTS", True)


            premiseList.Add(premise)

            pFeature = pCursor.NextFeature
        Loop

        If Not pFeature Is Nothing And counter > maxcount Then
            MsgBox("Your search returned more then " & maxcount & " record. " & vbCrLf _
                   & "Only first " & maxcount & " were displayed. " & vbCrLf _
                    & "Please refine your search.")
        End If

        Return premiseList

    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        b_searchFlag = False
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoom.Click
        LocatePremise("zoom")
    End Sub

    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click
        LocatePremise("pan")
    End Sub

    Private Sub LocatePremise(ByVal zoompan As String)
        Me.Cursor = Cursors.WaitCursor
        If b_searchingFlag Then
            Return
        End If

        b_searchingFlag = True
        Try
            Dim premiseList As IList = FindPremises(TextPremiseNo.Text, TextAccountNo.Text, TextOwner.Text, zoompan)

            If premiseList Is Nothing Then
                Throw New Exception("LocatePremise Exception")
            End If

            If premiseList.Count = 0 Then
                MsgBox("No premise was found.")
                Throw New Exception("LocatePremise Exception")
            End If

            If premiseList.Count = 1 AndAlso zoompan <> "" Then

                If Not m_form Is Nothing AndAlso m_form.Visible Then
                    m_form.Hide()
                End If

                Dim premiseLayer As IFeatureLayer = MapUtil.GetLayerByTableName("PremsInterPt", m_map)
                Dim premise As clsPremise = premiseList.Item(0)
                Dim pQFilter As IQueryFilter = New QueryFilter
                pQFilter.WhereClause = "PEXUID=" & premise.Pexuid
                Dim pCursor As IFeatureCursor = premiseLayer.Search(pQFilter, False)
                Dim pFeature As IFeature = pCursor.NextFeature
                MapUtil.SelectFeature(pFeature, premiseLayer)
                Select Case zoompan
                    Case "zoom"
                        MapUtil.ZoomToFeature(pFeature, m_map)
                    Case "pan"
                        MapUtil.PanToFeature(pFeature, m_map)
                End Select

            Else
                If Not m_form Is Nothing AndAlso m_form.Visible Then
                    m_form.SetPremisePts(premiseList)
                Else
                    m_form = New FormPremiseList
                    m_form.SetPremisePts(premiseList)
                    m_form.m_map = m_map
                    m_form.Show(New ModelessDialog(Me.Handle))
                End If
            End If

        Catch ex As Exception

        End Try

        b_searchingFlag = False
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub TextPremiseNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextPremiseNo.KeyPress
        Dim KCode As String = Asc(e.KeyChar)
        If (KCode >= 48 And KCode <= 57) Or KCode = 8 Or KCode = 37 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextPremiseNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextPremiseNo.TextChanged
        ResetSearchUI()
    End Sub

    Private Sub ResetSearchUI()
        If Trim(TextPremiseNo.Text) <> "" Then
            LabelPexprm.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
            LabelAccountNo.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelOwner.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

            If Len(Trim(TextPremiseNo.Text)) <> 7 Then
                btnZoom.Enabled = False
                btnPan.Enabled = False
            Else
                btnZoom.Enabled = True
                btnPan.Enabled = True
            End If
        ElseIf Trim(TextAccountNo.Text) <> "" Then
            LabelPexprm.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelAccountNo.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)
            LabelOwner.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

            If Len(Trim(TextAccountNo.Text)) < 5 Then
                btnZoom.Enabled = False
                btnPan.Enabled = False
            Else
                btnZoom.Enabled = True
                btnPan.Enabled = True
            End If

        ElseIf Trim(TextOwner.Text) <> "" Then
            LabelPexprm.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelAccountNo.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelOwner.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Bold)

            btnZoom.Enabled = False
            btnPan.Enabled = False

        Else
            LabelPexprm.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelAccountNo.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)
            LabelOwner.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

            btnZoom.Enabled = False
            btnPan.Enabled = False
        End If
    End Sub

    Private Sub TextAccountNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextAccountNo.KeyPress
        Dim KCode As String = Asc(e.KeyChar)
        If (KCode >= 48 And KCode <= 57) Or KCode = 8 Or KCode = 37 Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextAccountNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextAccountNo.TextChanged
        ResetSearchUI()
    End Sub

    Private Sub TextOwner_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextOwner.TextChanged
        ResetSearchUI()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        TextPremiseNo.Text = ""
        TextAccountNo.Text = ""
        TextOwner.Text = ""

        If Not m_form Is Nothing AndAlso m_form.Visible Then
            m_form.DataGridPremises.DataSource = Nothing
            m_form.DataGridPremises.Refresh()
        End If

    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        LocatePremise("")
    End Sub

    Private Sub FormPremiseLocator_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Tab Then
            If e.Modifiers = Keys.Shift Then
                Me.ProcessTabKey(False)
            Else
                Me.ProcessTabKey(True)
            End If
        End If

        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            Debug.Print("Set search mode to false ")
            b_searchFlag = False
        End If
    End Sub

    Private Sub FormPremiseLocator_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(System.Windows.Forms.Keys.Escape) Then
            Debug.Print("Set search mode to false ")
            b_searchFlag = False
        End If
    End Sub

End Class
