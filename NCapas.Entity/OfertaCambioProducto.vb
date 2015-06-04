<Serializable()>
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

    Private _codSucursalBanco As String
    Public Property CodSucursalBanco() As String
        Get
            Return _codSucursalBanco
        End Get
        Set(ByVal value As String)
            _codSucursalBanco = value
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

    Private _tipoContrato As String
    Public Property TipoContrato() As String
        Get
            Return _tipoContrato
        End Get
        Set(ByVal value As String)
            _tipoContrato = value
        End Set
    End Property

    Private _sistema As String
    Public Property Sistema() As String
        Get
            Return _sistema
        End Get
        Set(ByVal value As String)
            _sistema = value
        End Set
    End Property

    Private _codOficina As String
    Public Property CodOficina() As String
        Get
            Return _codOficina
        End Get
        Set(ByVal value As String)
            _codOficina = value
        End Set
    End Property

    Private _codUsuario As String
    Public Property CodUsuario() As String
        Get
            Return _codUsuario
        End Get
        Set(ByVal value As String)
            _codUsuario = value
        End Set
    End Property

    Private _numeroIp As String
    Public Property NumeroIp() As String
        Get
            Return _numeroIp
        End Get
        Set(ByVal value As String)
            _numeroIp = value
        End Set
    End Property

    Private _codCambio As String
    Public Property CodCambio() As String
        Get
            Return _codCambio
        End Get
        Set(ByVal value As String)
            _codCambio = value
        End Set
    End Property

    Private _productoOrigenTarjeta As String
    Public Property ProductoOrigenTarjeta() As String
        Get
            Return _productoOrigenTarjeta
        End Get
        Set(ByVal value As String)
            _productoOrigenTarjeta = value
        End Set
    End Property

    Private _productoDestinoTarjeta As String
    Public Property ProductoDestinoTarjeta() As String
        Get
            Return _productoDestinoTarjeta
        End Get
        Set(ByVal value As String)
            _productoDestinoTarjeta = value
        End Set
    End Property

    Private _descripcionProductoOrigenTarjeta As String
    Public Property DescripcionProductoOrigenTarjeta() As String
        Get
            Return _descripcionProductoOrigenTarjeta
        End Get
        Set(ByVal value As String)
            _descripcionProductoOrigenTarjeta = value
        End Set
    End Property

    Private _descripcionProductoDestinoTarjeta As String
    Public Property DescripcionProductoDestinoTarjeta() As String
        Get
            Return _descripcionProductoDestinoTarjeta
        End Get
        Set(ByVal value As String)
            _descripcionProductoDestinoTarjeta = value
        End Set
    End Property

    Private _carpetaDestino As String
    Public Property CarpetaDestino() As String
        Get
            Return _carpetaDestino
        End Get
        Set(ByVal value As String)
            _carpetaDestino = value
        End Set
    End Property


End Class

