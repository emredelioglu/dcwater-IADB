Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI

Public Class FormMarAddress

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim marAddr As clsMARAddress = ListMARAddress.SelectedItem
        If marAddr.AddrNum = "" Then
            MsgBox("The address point you selected does not have address number.")
            Return
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ListMARAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListMARAddress.SelectedIndexChanged
        If Not ListMARAddress.SelectedItem Is Nothing Then
            OK_Button.Enabled = True

            'Flash the feature
            'Get the feature first
            Dim marAddr As clsMARAddress = ListMARAddress.SelectedItem

            Dim pDoc As IMxDocument = IAISApplication.GetInstance().ArcMapApplicaiton.Document

            Dim addrptLayer As IFeatureLayer = MapUtil.GetLayerByBrowseName(IAISToolSetting.GetParameterValue("LAYER_ADDRESSPT"), pDoc.FocusMap)
            If addrptLayer Is Nothing Then
                Return
            End If
            Dim pFilter As IQueryFilter = New QueryFilter
            pFilter.WhereClause = "ADDRESS_ID=" & marAddr.AddressId
            Dim pFeatCursor As IFeatureCursor = addrptLayer.Search(pFilter, False)

            Dim addPt As IFeature = pFeatCursor.NextFeature
            If Not addPt Is Nothing Then
                MapUtil.FlashFeature(addPt, pDoc.FocusMap)
            End If

        Else
            OK_Button.Enabled = False
        End If
    End Sub
End Class
