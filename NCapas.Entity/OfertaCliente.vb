Public Class OfertaCliente
    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _reenganche As Integer
    Public Property Reenganche() As Integer
        Get
            Return _reenganche
        End Get
        Set(ByVal value As Integer)
            _reenganche = value
        End Set
    End Property

    Private _estado As String
    Public Property Estado() As String
        Get
            Return _estado
        End Get
        Set(ByVal value As String)
            _estado = value
        End Set
    End Property

    Private _existe As String
    Public Property Existe() As String
        Get
            Return _existe
        End Get
        Set(ByVal value As String)
            _existe = value
        End Set
    End Property

    Private _tipoProducto As String
    Public Property TipoProducto() As String
        Get
            Return _tipoProducto
        End Get
        Set(ByVal value As String)
            _tipoProducto = value
        End Set
    End Property

    Private _nombreCliente As String
    Public Property NombreCliente() As String
        Get
            Return _nombreCliente
        End Get
        Set(ByVal value As String)
            _nombreCliente = value
        End Set
    End Property

    Private _numeroCuenta As String
    Public Property NumeroCuenta() As String
        Get
            Return _numeroCuenta
        End Get
        Set(ByVal value As String)
            _numeroCuenta = value
        End Set
    End Property

    Private _lineaOferta As String
    Public Property LineaOferta() As String
        Get
            Return _lineaOferta
        End Get
        Set(ByVal value As String)
            _lineaOferta = value
        End Set
    End Property

    Private _plazo As String
    Public Property Plazo() As String
        Get
            Return _plazo
        End Get
        Set(ByVal value As String)
            _plazo = value
        End Set
    End Property

    Private _tasa As String
    Public Property Tasa() As String
        Get
            Return _tasa
        End Get
        Set(ByVal value As String)
            _tasa = value
        End Set
    End Property

    Private _cuota As String
    Public Property Cuota() As String
        Get
            Return _cuota
        End Get
        Set(ByVal value As String)
            _cuota = value
        End Set
    End Property

    Private _tem As String
    Public Property Tem() As String
        Get
            Return _tem
        End Get
        Set(ByVal value As String)
            _tem = value
        End Set
    End Property

    Private _tea As String
    Public Property Tea() As String
        Get
            Return _tea
        End Get
        Set(ByVal value As String)
            _tea = value
        End Set
    End Property

    Private _fechaInicialVigencia As String
    Public Property FechaInicialVigencia() As String
        Get
            Return _fechaInicialVigencia
        End Get
        Set(ByVal value As String)
            _fechaInicialVigencia = value
        End Set
    End Property

    Private _fechaFinVigencia As String
    Public Property FechaFinVigencia() As String
        Get
            Return _fechaFinVigencia
        End Get
        Set(ByVal value As String)
            _fechaFinVigencia = value
        End Set
    End Property

    Private _lineaOfertaInicial As String
    Public Property LineaOfertaInicial() As String
        Get
            Return _lineaOfertaInicial
        End Get
        Set(ByVal value As String)
            _lineaOfertaInicial = value
        End Set
    End Property

    Private _plazoOfertaInicial As String
    Public Property PlazoOfertaInicial() As String
        Get
            Return _plazoOfertaInicial
        End Get
        Set(ByVal value As String)
            _plazoOfertaInicial = value
        End Set
    End Property

    Private _tasaOfertaInicial As String
    Public Property TasaOfertaInicial() As String
        Get
            Return _tasaOfertaInicial
        End Get
        Set(ByVal value As String)
            _tasaOfertaInicial = value
        End Set
    End Property

    Private _cuotaOfertaInicial As String
    Public Property CuotaOfertaInicial() As String
        Get
            Return _cuotaOfertaInicial
        End Get
        Set(ByVal value As String)
            _cuotaOfertaInicial = value
        End Set
    End Property

    Private _codigoOferta As String
    Public Property CodigoOferta() As String
        Get
            Return _codigoOferta
        End Get
        Set(ByVal value As String)
            _codigoOferta = value
        End Set
    End Property


End Class
