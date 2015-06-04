Public Class OfertaCambioProducto

    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _contratoTarjeta As String
    Public Property ContratoTarjeta() As String
        Get
            Return _contratoTarjeta
        End Get
        Set(ByVal value As String)
            _contratoTarjeta = value
        End Set
    End Property

    Private _contratoSEF As String
    Public Property ContratoSEF() As String
        Get
            Return _contratoSEF
        End Get
        Set(ByVal value As String)
            _contratoSEF = value
        End Set
    End Property

    Private _codCambioTarjeta As String
    Public Property CodCambioTarjeta() As String
        Get
            Return _codCambioTarjeta
        End Get
        Set(ByVal value As String)
            _codCambioTarjeta = value
        End Set
    End Property

    Private _CodCambioSEF As String
    Public Property CodCambioSEF() As String
        Get
            Return _CodCambioSEF
        End Get
        Set(ByVal value As String)
            _CodCambioSEF = value
        End Set
    End Property

    Private _tipoDocumento As String
    Public Property TipoDocumento() As String
        Get
            Return _tipoDocumento
        End Get
        Set(ByVal value As String)
            _tipoDocumento = value
        End Set
    End Property

    Private _numeroDocumento As String
    Public Property NumeroDocumento() As String
        Get
            Return _numeroDocumento
        End Get
        Set(ByVal value As String)
            _numeroDocumento = value
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

    Private _fechaInicioVigencia As String
    Public Property FechaInicioVigencia() As String
        Get
            Return _fechaInicioVigencia
        End Get
        Set(ByVal value As String)
            _fechaInicioVigencia = value
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

    Private _estado As String
    Public Property Estado() As String
        Get
            Return _estado
        End Get
        Set(ByVal value As String)
            _estado = value
        End Set
    End Property

    Private _idSucursal As Integer
    Public Property IdSucursal() As Integer
        Get
            Return _idSucursal
        End Get
        Set(ByVal value As Integer)
            _idSucursal = value
        End Set
    End Property

    Private _numeroCaja As String
    Public Property NumeroCaja() As String
        Get
            Return _numeroCaja
        End Get
        Set(ByVal value As String)
            _numeroCaja = value
        End Set
    End Property


    Private _nombreKiosko As String
    Public Property NombreKiosko() As String
        Get
            Return _nombreKiosko
        End Get
        Set(ByVal value As String)
            _nombreKiosko = value
        End Set
    End Property

    Private _fechaImpresion As String
    Public Property FechaImpresion() As String
        Get
            Return _fechaImpresion
        End Get
        Set(ByVal value As String)
            _fechaImpresion = value
        End Set
    End Property

    Private _datosTarjeta As String
    Public Property DatosTarjeta() As String
        Get
            Return _datosTarjeta
        End Get
        Set(ByVal value As String)
            _datosTarjeta = value
        End Set
    End Property

    Private _datosSEF As String
    Public Property DatosSEF() As String
        Get
            Return _datosSEF
        End Get
        Set(ByVal value As String)
            _datosSEF = value
        End Set
    End Property

    Private _ofertaVigente As String
    Public Property OfertaVigente() As String
        Get
            Return _ofertaVigente
        End Get
        Set(ByVal value As String)
            _ofertaVigente = value
        End Set
    End Property

    Private _mensajeError As String
    Public Property MensajeError() As String
        Get
            Return _mensajeError
        End Get
        Set(ByVal value As String)
            _mensajeError = value
        End Set
    End Property

End Class
