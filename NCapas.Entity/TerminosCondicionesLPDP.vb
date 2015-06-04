<Serializable()>
Public Class TerminosCondicionesLPDP
    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _titulo As String
    Public Property Titulo() As String
        Get
            Return _titulo
        End Get
        Set(ByVal value As String)
            _titulo = value
        End Set
    End Property

    Private _cabecera As String
    Public Property Cabecera() As String
        Get
            Return _cabecera
        End Get
        Set(ByVal value As String)
            _cabecera = value
        End Set
    End Property

    Private _condiciones As String
    Public Property Condiciones() As String
        Get
            Return _condiciones
        End Get
        Set(ByVal value As String)
            _condiciones = value
        End Set
    End Property

    Private _confidencialidad As String
    Public Property Confidencialidad() As String
        Get
            Return _confidencialidad
        End Get
        Set(ByVal value As String)
            _confidencialidad = value
        End Set
    End Property

    Private _caducidad As String
    Public Property Caducidad() As String
        Get
            Return _caducidad
        End Get
        Set(ByVal value As String)
            _caducidad = value
        End Set
    End Property



End Class
