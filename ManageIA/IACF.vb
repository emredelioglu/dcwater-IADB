Public Class IACF
    Private _pexuid As Integer
    Private _iasqft As Double
    Private _iaeru As Double


    Private _pexptyp As String

    Private _ialayername As String

    Public Property Pexuid() As Integer
        Get
            Return _pexuid
        End Get
        Set(ByVal value As Integer)
            _pexuid = value
        End Set
    End Property

    Public Property Iasqft() As Double
        Get
            Return _iasqft
        End Get
        Set(ByVal value As Double)
            _iasqft = value
        End Set
    End Property

    Public Property Iaeru() As Double
        Get
            Return _iaeru
        End Get
        Set(ByVal value As Double)
            _iaeru = value
        End Set
    End Property

    Public Property Pexptyp() As String
        Get
            Return _pexptyp
        End Get
        Set(ByVal value As String)
            _pexptyp = value
        End Set
    End Property

    Public Property IaLayerName() As String
        Get
            Return _ialayername
        End Get
        Set(ByVal value As String)
            _ialayername = value
        End Set
    End Property


End Class
