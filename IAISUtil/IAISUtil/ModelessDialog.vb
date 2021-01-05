
Namespace Windows.UI

    ''' <summary>
    ''' Wrapper class that allows a non-.NET form handle to be passed into Form.Show(), making the form a modeless dialog.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ModelessDialog
        Implements System.Windows.Forms.IWin32Window
        Private hwnd As System.IntPtr

        ''' <summary>
        ''' Returns the form handle.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return hwnd
            End Get
        End Property

        ''' <summary>
        ''' Constructor.
        ''' </summary>
        ''' <param name="handle">The form handle passed as an IntPtr</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal handle As System.IntPtr)
            hwnd = handle
        End Sub

        ''' <summary>
        ''' Constructor.
        ''' </summary>
        ''' <param name="handle">The form handle passed as an integer</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal handle As Integer)
            hwnd = CType(handle, IntPtr)
        End Sub
    End Class
End Namespace
