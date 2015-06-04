Public Class SimuladorSEF

    Private _LNK_FINANCIADO As String
    Public Property LNK_FINANCIADO() As String
        Get
            Return _LNK_FINANCIADO
        End Get
        Set(ByVal value As String)
            _LNK_FINANCIADO = value
        End Set
    End Property

    Private _LNK_CUOTAS As String
    Public Property LNK_CUOTAS() As String
        Get
            Return _LNK_CUOTAS
        End Get
        Set(ByVal value As String)
            _LNK_CUOTAS = value
        End Set
    End Property

    Private _LNK_DIFERIDO As String
    Public Property LNK_DIFERIDO() As String
        Get
            Return _LNK_DIFERIDO
        End Get
        Set(ByVal value As String)
            _LNK_DIFERIDO = value
        End Set
    End Property

    Private _LNK_TASA_ANUAL As String
    Public Property LNK_TASA_ANUAL() As String
        Get
            Return _LNK_TASA_ANUAL
        End Get
        Set(ByVal value As String)
            _LNK_TASA_ANUAL = value
        End Set
    End Property

    Private _LNK_FECHA_CONSUMO As String
    Public Property LNK_FECHA_CONSUMO() As String
        Get
            Return _LNK_FECHA_CONSUMO
        End Get
        Set(ByVal value As String)
            _LNK_FECHA_CONSUMO = value
        End Set
    End Property

    Private _LNK_FECHA_FACTURACION As String
    Public Property LNK_FECHA_FACTURACION() As String
        Get
            Return _LNK_FECHA_FACTURACION
        End Get
        Set(ByVal value As String)
            _LNK_FECHA_FACTURACION = value
        End Set
    End Property

    Private _LNK_ESTADO As String
    Public Property LNK_ESTADO() As String
        Get
            Return _LNK_ESTADO
        End Get
        Set(ByVal value As String)
            _LNK_ESTADO = value
        End Set
    End Property

    Private _LNK_TABLA As String
    Public Property LNK_TABLA() As String
        Get
            Return _LNK_TABLA
        End Get
        Set(ByVal value As String)
            _LNK_TABLA = value
        End Set
    End Property

    Private _resultado As Integer
    Public Property Resultado() As Integer
        Get
            Return _resultado
        End Get
        Set(ByVal value As Integer)
            _resultado = value
        End Set
    End Property
    Private _Mensaje As String
    Public Property Mensaje() As String
        Get
            Return _Mensaje
        End Get
        Set(ByVal value As String)
            _Mensaje = value
        End Set
    End Property
End Class
