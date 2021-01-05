Public Class clsPremiseInfo
    Private _serviceAddress1 As String
    Private _ServiceAddress2 As String


    Private _ListOtherPremises As List(Of clsPremisePt)
    Private _PremiseNumber As String
    Private _ssl As String
    Private _IAArea As Integer
    Private _IAERU As Double

    Private _EffectiveDate As Date
    Private _BldgArea As Integer
    Private _RoadArea As Integer
    Private _OutdoorArea As Integer
    Private _SwmpoolArea As Integer
    Private _SidewalkArea As Integer
    Private _StairsArea As Integer

    Private _directAssign As String
    Private _pexptyp As String

    Public Property Pexptyp() As String
        Get
            Return _pexptyp
        End Get
        Set(ByVal value As String)
            _pexptyp = value
        End Set
    End Property

    Public Property DirectAssign() As String
        Get
            Return _directAssign
        End Get
        Set(ByVal value As String)
            _directAssign = value
        End Set
    End Property


    Public Property ServiceAddress1() As String
        Get
            Return _serviceAddress1
        End Get
        Set(ByVal value As String)
            _serviceAddress1 = value
        End Set
    End Property

    Public Property ServiceAddress2() As String
        Get
            Return _ServiceAddress2
        End Get
        Set(ByVal value As String)
            _ServiceAddress2 = value
        End Set
    End Property

    Public Property PremiseNumber() As String
        Get
            Return _PremiseNumber
        End Get
        Set(ByVal value As String)
            _PremiseNumber = value
        End Set
    End Property

    Public Property SSL() As String
        Get
            Return _ssl
        End Get
        Set(ByVal value As String)
            _ssl = value
        End Set
    End Property

    Public Property IAArea() As Integer
        Get
            Dim areaD As Double = Math.Floor(_IAArea / 100) * 100
            Return CInt(areaD)
        End Get
        Set(ByVal value As Integer)
            _IAArea = value
        End Set
    End Property

    Public Property IAERU() As Double
        Get
            Return _IAERU
        End Get
        Set(ByVal value As Double)
            _IAERU = value
        End Set
    End Property

    Public Property EffectiveDate() As Date
        Get
            Return _EffectiveDate
        End Get
        Set(ByVal value As Date)
            _EffectiveDate = value
        End Set
    End Property

    Public Property BldgArea() As Integer
        Get
            Return _BldgArea
        End Get
        Set(ByVal value As Integer)
            _BldgArea = value
        End Set
    End Property

    Public Property RoadArea() As Integer
        Get
            Return _RoadArea
        End Get
        Set(ByVal value As Integer)
            _RoadArea = value
        End Set
    End Property

    Public Property StairsArea() As Integer
        Get
            Return _StairsArea
        End Get
        Set(ByVal value As Integer)
            _StairsArea = value
        End Set
    End Property

    Public Property SidewalkArea() As Integer
        Get
            Return _SidewalkArea
        End Get
        Set(ByVal value As Integer)
            _SidewalkArea = value
        End Set
    End Property

    Public Property SwmpoolArea() As Integer
        Get
            Return _SwmpoolArea
        End Get
        Set(ByVal value As Integer)
            _SwmpoolArea = value
        End Set
    End Property

    Public Property OutdoorArea() As Integer
        Get
            Return _OutdoorArea
        End Get
        Set(ByVal value As Integer)
            _OutdoorArea = value
        End Set
    End Property

    Public Property ListOtherPremises() As List(Of clsPremisePt)
        Get
            Return _ListOtherPremises
        End Get
        Set(ByVal value As List(Of clsPremisePt))
            _ListOtherPremises = value
        End Set
    End Property


    Public Sub SetIAArea(ByVal iaarea As Integer, ByVal featuretype As Integer)
       If featuretype = 1 Then
            _BldgArea = _BldgArea + iaarea
        ElseIf featuretype = 2 Then
            _SidewalkArea = _SidewalkArea + iaarea
        ElseIf featuretype = 3 Then
            _StairsArea = _StairsArea + iaarea
        ElseIf featuretype = 4 Then
            _RoadArea = _RoadArea + iaarea
        ElseIf featuretype = 5 Then
            _OutdoorArea = _OutdoorArea + iaarea
        ElseIf featuretype = 6 Then
            _SwmpoolArea = _SwmpoolArea + iaarea
        End If
    End Sub

    Public Function addPremisePt(ByVal pt As clsPremisePt)
        If _ListOtherPremises Is Nothing Then
            _ListOtherPremises = New List(Of clsPremisePt)
        End If
        _ListOtherPremises.Add(pt)

    End Function

    Public Sub Reset()
        _serviceAddress1 = ""
        _ServiceAddress2 = ""
        _ListOtherPremises = Nothing
        _PremiseNumber = ""
        _ssl = ""
        _IAArea = 0
        _IAERU = 0

        _EffectiveDate = Nothing
        _BldgArea = 0
        _RoadArea = 0
        _OutdoorArea = 0
        _SwmpoolArea = 0
        _SidewalkArea = 0
        _StairsArea = 0
    End Sub
End Class
