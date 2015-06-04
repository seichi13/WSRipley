<Serializable()>
Public Class TerminosCondiciones
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

    Private _responsabilidadInf As String
    Public Property ResponsabilidadInf() As String
        Get
            Return _responsabilidadInf
        End Get
        Set(ByVal value As String)
            _responsabilidadInf = value
        End Set
    End Property

    Private _acceso As String
    Public Property Acceso() As String
        Get
            Return _acceso
        End Get
        Set(ByVal value As String)
            _acceso = value
        End Set
    End Property

    Private _fallas As String
    Public Property Fallas() As String
        Get
            Return _fallas
        End Get
        Set(ByVal value As String)
            _fallas = value
        End Set
    End Property

    Private _propiedad As String
    Public Property Propiedad() As String
        Get
            Return _propiedad
        End Get
        Set(ByVal value As String)
            _propiedad = value
        End Set
    End Property

    Private _responsabilidad As String
    Public Property Responsabilidad() As String
        Get
            Return _responsabilidad
        End Get
        Set(ByVal value As String)
            _responsabilidad = value
        End Set
    End Property

    Private _otros As String
    Public Property Otros() As String
        Get
            Return _otros
        End Get
        Set(ByVal value As String)
            _otros = value
        End Set
    End Property

    Private _anexo As String
    Public Property Anexo() As String
        Get
            Return _anexo
        End Get
        Set(ByVal value As String)
            _anexo = value
        End Set
    End Property


End Class
