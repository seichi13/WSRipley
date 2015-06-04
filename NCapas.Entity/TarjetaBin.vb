Public Class TarjetaBin
    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _bin As String
    Public Property Bin() As String
        Get
            Return _bin
        End Get
        Set(ByVal value As String)
            _bin = value
        End Set
    End Property

    Private _tipo As String
    Public Property Tipo() As String
        Get
            Return _tipo
        End Get
        Set(ByVal value As String)
            _tipo = value
        End Set
    End Property


    Private _nombre As String
    Public Property Nombre() As String
        Get
            Return _nombre
        End Get
        Set(ByVal value As String)
            _nombre = value
        End Set
    End Property

    Private _metodo As String
    Public Property Metodo() As String
        Get
            Return _metodo
        End Get
        Set(ByVal value As String)
            _metodo = value
        End Set
    End Property

    Private _tipoPasado As String
    Public Property TipoPasado() As String
        Get
            Return _tipoPasado
        End Get
        Set(ByVal value As String)
            _tipoPasado = value
        End Set
    End Property

    Private _tipoSEF As String
    Public Property TipoSEF() As String
        Get
            Return _tipoSEF
        End Get
        Set(ByVal value As String)
            _tipoSEF = value
        End Set
    End Property

    Private _tipoSimuladorSEF As String
    Public Property TipoSimuladorSEF() As String
        Get
            Return _tipoSimuladorSEF
        End Get
        Set(ByVal value As String)
            _tipoSimuladorSEF = value
        End Set
    End Property

    Private _tipoAbiertoRsat As Integer
    Public Property TipoAbiertoRsat() As Integer
        Get
            Return _tipoAbiertoRsat
        End Get
        Set(ByVal value As Integer)
            _tipoAbiertoRsat = value
        End Set
    End Property


End Class
