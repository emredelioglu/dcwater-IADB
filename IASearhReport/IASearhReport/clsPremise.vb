Public Class clsPremise

    Private m_pexprm As String
    Private m_pexact As String
    Private m_owner As String
    Private m_pexsad As String
    Private m_pexpsts As String
    Private m_pexptyp As String
    Private m_pexuid As Long

    Public Property Pexprm() As String
        Get
            Return m_pexprm
        End Get
        Set(ByVal value As String)
            m_pexprm = value
        End Set
    End Property

    Public Property Pexact() As String
        Get
            Return m_pexact
        End Get
        Set(ByVal value As String)
            m_pexact = value
        End Set
    End Property

    Public Property Owner() As String
        Get
            Return m_owner
        End Get
        Set(ByVal value As String)
            m_owner = value
        End Set
    End Property

    Public Property Pexsad() As String
        Get
            Return m_pexsad
        End Get
        Set(ByVal value As String)
            m_pexsad = value
        End Set
    End Property

    Public Property Pexpsts() As String
        Get
            Return m_pexpsts
        End Get
        Set(ByVal value As String)
            m_pexpsts = value
        End Set
    End Property

    Public Property Pexptyp() As String
        Get
            Return m_pexptyp
        End Get
        Set(ByVal value As String)
            m_pexptyp = value
        End Set
    End Property

    Public Property Pexuid() As Long
        Get
            Return m_pexuid
        End Get
        Set(ByVal value As Long)
            m_pexuid = value
        End Set
    End Property
End Class
