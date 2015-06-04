Public Class Cliente
    Private _correoLaboral As String
    Public Property CorreoLaboral() As String
        Get
            Return _correoLaboral
        End Get
        Set(ByVal value As String)
            _correoLaboral = value
        End Set
    End Property

    Private _correoPersonal As String
    Public Property CorreoPersonal() As String
        Get
            Return _correoPersonal
        End Get
        Set(ByVal value As String)
            _correoPersonal = value
        End Set
    End Property


End Class
