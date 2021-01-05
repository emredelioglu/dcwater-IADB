Imports ESRI.ArcGIS.Geodatabase

Public Class FormPremiseAttributeAdv

    Private m_premiseForm As PremiseAttribute

    Public WriteOnly Property PremiseForm() As PremiseAttribute
        Set(ByVal value As PremiseAttribute)
            m_premiseForm = value
        End Set
    End Property

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        'Save attribute 

        m_premiseForm.loctnprecs = MapUtil.GetDomainCode(Me.cmbLOCTNPRECS)
        m_premiseForm.inforsrc = MapUtil.GetDomainCode(Me.cmbINFORSRC)

        m_premiseForm.address_id = Me.txtADDRESS_ID.Text
        m_premiseForm.master_pexuid = Me.txtMASTER_PEXUID.Text
        m_premiseForm.comnt = Me.txtCOMNT.Text

        m_premiseForm.loctnprecs = MapUtil.GetDomainCode(Me.cmbLOCTNPRECS)


        If chkFROM_STANDALONE.Checked Then
            m_premiseForm.from_standalone = "Y"
        Else
            m_premiseForm.from_standalone = "N"
        End If

        If chkHAS_LIEN.Checked Then
            m_premiseForm.has_lien = "Y"
        Else
            m_premiseForm.has_lien = "N"
        End If

        Me.Close()

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnClearMasterPexUid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearMasterPexUid.Click
        txtMASTER_PEXUID.Text = ""
    End Sub

    Private Sub btnClearMARId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearMARId.Click
        txtADDRESS_ID.Text = ""
    End Sub
End Class