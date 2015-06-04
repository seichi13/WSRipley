Public Class MetodoTipoDocumento

    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property


    Private _tipoDocumento As Integer
    Public Property TipoDocumento() As Integer
        Get
            Return _tipoDocumento
        End Get
        Set(ByVal value As Integer)
            _tipoDocumento = value
        End Set
    End Property

    Private _idTipoTarjeta As Integer
    Public Property IdTipoTarjeta() As Integer
        Get
            Return _idTipoTarjeta
        End Get
        Set(ByVal value As Integer)
            _idTipoTarjeta = value
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

    Private _estado As Integer
    Public Property Estado() As Integer
        Get
            Return _estado
        End Get
        Set(ByVal value As Integer)
            _estado = value
        End Set
    End Property

End Class
