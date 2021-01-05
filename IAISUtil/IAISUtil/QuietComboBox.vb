Imports System.Windows.Forms

Public Class QuietComboBox
    Inherits ComboBox


    ''' <param name="e">A <see cref="T:System.Windows.Forms.KeyPressEventArgs" /> that contains the event data.</param>
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        e.Handled = True
    End Sub
End Class
