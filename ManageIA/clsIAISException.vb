Public Class clsIAISException
    Private _objectid As Integer
    Private _exceptdt As Date
    Private _excepttyp As String
    Private _appname As String
    Private _apptype As String
    Private _reviewst As String
    Private _revdt As Date
    Private _status As String
    Private _reviewby As String
    Private _fixdt As Date
    Private _fixby As String
    Private _uniqueId As String

    Public Property AppName() As String
        Get
            Return _appname
        End Get
        Set(ByVal value As String)
            _appname = value
        End Set
    End Property


    Public Property Objectid() As String
        Get
            Return _Objectid
        End Get
        Set(ByVal value As String)
            _Objectid = value
        End Set
    End Property

    Public Property Exceptdt() As Date
        Get
            Return _exceptdt
        End Get
        Set(ByVal value As Date)
            _exceptdt = value
        End Set
    End Property

    Public Property Excepttyp() As String
        Get
            Return _Excepttyp
        End Get
        Set(ByVal value As String)
            _Excepttyp = value
        End Set
    End Property

    Public Property Apptype() As String
        Get
            Return _Apptype
        End Get
        Set(ByVal value As String)
            _Apptype = value
        End Set
    End Property

    Public Property Reviewst() As String
        Get
            Return _Reviewst
        End Get
        Set(ByVal value As String)
            _Reviewst = value
        End Set
    End Property

    Public Property Revdt() As Date
        Get
            Return _revdt
        End Get
        Set(ByVal value As Date)
            _revdt = value
        End Set
    End Property

    Public Property Status() As String
        Get
            Return _Status
        End Get
        Set(ByVal value As String)
            _Status = value
        End Set
    End Property

    Public Property Reviewby() As String
        Get
            Return _Reviewby
        End Get
        Set(ByVal value As String)
            _Reviewby = value
        End Set
    End Property

    Public Property Fixdt() As Date
        Get
            Return _fixdt
        End Get
        Set(ByVal value As Date)
            _fixdt = value
        End Set
    End Property

    Public Property Fixby() As String
        Get
            Return _Fixby
        End Get
        Set(ByVal value As String)
            _Fixby = value
        End Set
    End Property

    Public Property UniqueId() As String
        Get
            Return _uniqueId
        End Get
        Set(ByVal value As String)
            _uniqueId = value
        End Set
    End Property


End Class
