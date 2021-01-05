
Public Class clsPremisePt

    Private _masterFlag As Boolean
    Private _pexprm As String
    Private _pexuid As Integer
    Private _pexsad As String
    Private _imperviousOnly As String
    Private _exemptIAB As String
    Private _exemptReason As String


    Public Property masterFlag() As Boolean
        Get
            masterFlag = _masterFlag
        End Get
        Set(ByVal value As Boolean)
            _masterFlag = value
        End Set
    End Property

    Public Property pexprm() As String
        Get
            pexprm = _pexprm
        End Get
        Set(ByVal value As String)
            _pexprm = value
        End Set
    End Property

    Public Property pexuid() As Integer
        Get
            pexuid = _pexuid
        End Get
        Set(ByVal value As Integer)
            _pexuid = value
        End Set
    End Property

    Public Property pexsad() As String
        Get
            pexsad = _pexsad
        End Get
        Set(ByVal value As String)
            _pexsad = value
        End Set
    End Property

    Public Property imperviousOnly() As String
        Get
            imperviousOnly = _imperviousOnly
        End Get
        Set(ByVal value As String)
            _imperviousOnly = value
        End Set
    End Property

    Public Property exemptIAB() As String
        Get
            Return _exemptIAB
        End Get
        Set(ByVal value As String)
            _exemptIAB = value
        End Set
    End Property

    Public Property exemptReason() As String
        Get
            Return _exemptReason
        End Get
        Set(ByVal value As String)
            _exemptReason = value
        End Set
    End Property
End Class


