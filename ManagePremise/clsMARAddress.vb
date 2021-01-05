Public Class clsMARAddress

    Private _AddrNum As String
    Private _StName As String
    Private _StreetType As String
    Private _Quadrant As String
    Private _FullAddress As String
    Private _ZipCode As String
    Private _Score As Double
    Private _AddrNumSuf As String
    Private _AddressId As String

    Public Property AddressId() As String
        Get
            Return _AddressId
        End Get
        Set(ByVal value As String)
            _AddressId = value
        End Set
    End Property

    Public Property AddrNum() As String
        Get
            Return _AddrNum
        End Get
        Set(ByVal value As String)
            _AddrNum = value
        End Set
    End Property

    Public Property StName() As String
        Get
            Return _StName
        End Get
        Set(ByVal value As String)
            _StName = value
        End Set
    End Property

    Public Property StreetType() As String
        Get
            Return _StreetType
        End Get
        Set(ByVal value As String)
            _StreetType = value
        End Set
    End Property

    Public Property Quadrant() As String
        Get
            Return _Quadrant
        End Get
        Set(ByVal value As String)
            _Quadrant = value
        End Set
    End Property

    Public Property FullAddress() As String
        Get
            If Trim(_FullAddress) = "" Then
                Return MapUtil.TrimAll(_AddrNum & " " & _AddrNumSuf & " " & _StName & " " & _StreetType & " " & _Quadrant)
            End If
            Return _FullAddress
        End Get
        Set(ByVal value As String)
            _FullAddress = value
        End Set
    End Property

    Public Property Score() As Double
        Get
            Return _Score
        End Get
        Set(ByVal value As Double)
            _Score = value
        End Set
    End Property

    Public Property ZipCode() As Integer
        Get
            Return _ZipCode
        End Get
        Set(ByVal value As Integer)
            _ZipCode = value
        End Set
    End Property

    Public Property AddrNumSuf() As String
        Get
            Return _AddrNumSuf
        End Get
        Set(ByVal value As String)
            _AddrNumSuf = value
        End Set
    End Property
End Class
