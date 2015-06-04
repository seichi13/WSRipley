Public Class SuperEfectivo

    Public Class Salida

        Private _nWSResult As Integer
        Public Property nWSResult() As Integer
            Get
                Return _nWSResult
            End Get
            Set(ByVal value As Integer)
                _nWSResult = value
            End Set
        End Property
        Private _sWSMessage As String
        Public Property sWSMessage() As String
            Get
                Return _sWSMessage
            End Get
            Set(ByVal value As String)
                _sWSMessage = value
            End Set
        End Property
        Private _sNumeroCuenta As String
        Public Property sNumeroCuenta() As String
            Get
                Return _sNumeroCuenta
            End Get
            Set(ByVal value As String)
                _sNumeroCuenta = value
            End Set
        End Property
        Private _sNumeroTarjeta As String
        Public Property sNumeroTarjeta() As String
            Get
                Return _sNumeroTarjeta
            End Get
            Set(ByVal value As String)
                _sNumeroTarjeta = value
            End Set
        End Property

    End Class

    Public Class Parametros
        Private _sCodigoOficina As String
        Private _sTerminalFisico As String
        Private _sCodigoCanal As String
        Private _sCodUsuario As String

        Private _sNombreCliente As String
        Private _sTipoDocumento As String
        Private _sNumeroDocumento As String
        Private _sTipoProducto As String
        Private _sCodigoEntidad As String
        Private _sCentroAlta As String
        Private _sCuenta As String
        Private _nMontoAceptado As String
        Private _nPlazoMes As String
        Private _nCuotaMes As String
        Private _nTEA As String
        Private _nTEM As String
        Private _sMoneda As String
        Private _sFechaVencimiento As String
        Private _sCanal As String
        Private _sSucural As String
        Private _sCodigoVendedor As String
        Private _sCodigoUsuario As String
        Private _sCodigoAtencion As String
        Private _sAcplicaITF As String
        Private _sIndicadorCommit As String

        Public Property sCodigoOficina() As String
            Get
                Return _sCodigoOficina
            End Get
            Set(ByVal value As String)
                _sCodigoOficina = value
            End Set
        End Property
        Public Property sTerminalFisico() As String
            Get
                Return _sTerminalFisico
            End Get
            Set(ByVal value As String)
                _sTerminalFisico = value
            End Set
        End Property
        Public Property sCodigoCanal() As String
            Get
                Return _sCodigoCanal
            End Get
            Set(ByVal value As String)
                _sCodigoCanal = value
            End Set
        End Property
        Public Property sCodUsuario() As String
            Get
                Return _sCodUsuario
            End Get
            Set(ByVal value As String)
                _sCodUsuario = value
            End Set
        End Property

        Public Property sNombreCliente As String
            Get
                Return _sNombreCliente
            End Get
            Set(ByVal value As String)
                _sNombreCliente = value
            End Set
        End Property
        Public Property sTipoDocumento As String
            Get
                Return _sTipoDocumento
            End Get
            Set(ByVal value As String)
                _sTipoDocumento = value
            End Set
        End Property
        Public Property sNumeroDocumento As String
            Get
                Return _sNumeroDocumento
            End Get
            Set(ByVal value As String)
                _sNumeroDocumento = value
            End Set
        End Property
        Public Property sTipoProducto() As String
            Get
                Return _sTipoProducto
            End Get
            Set(ByVal value As String)
                _sTipoProducto = value
            End Set
        End Property
        Public Property sCodigoEntidad() As String
            Get
                Return _sCodigoEntidad
            End Get
            Set(ByVal value As String)
                _sCodigoEntidad = value
            End Set
        End Property
        Public Property sCentroAlta() As String
            Get
                Return _sCentroAlta
            End Get
            Set(ByVal value As String)
                _sCentroAlta = value
            End Set
        End Property
        Public Property sCuenta() As String
            Get
                Return _sCuenta
            End Get
            Set(ByVal value As String)
                _sCuenta = value
            End Set
        End Property
        Public Property nMontoAceptado() As String
            Get
                Return _nMontoAceptado
            End Get
            Set(ByVal value As String)
                _nMontoAceptado = value
            End Set
        End Property
        Public Property nPlazoMes() As String
            Get
                Return _nPlazoMes
            End Get
            Set(ByVal value As String)
                _nPlazoMes = value
            End Set
        End Property
        Public Property nCuotaMes() As String
            Get
                Return _nCuotaMes
            End Get
            Set(ByVal value As String)
                _nCuotaMes = value
            End Set
        End Property
        Public Property nTEA() As String
            Get
                Return _nTEA
            End Get
            Set(ByVal value As String)
                _nTEA = value
            End Set
        End Property
        Public Property nTEM() As String
            Get
                Return _nTEM
            End Get
            Set(ByVal value As String)
                _nTEM = value
            End Set
        End Property
        Public Property sMoneda() As String
            Get
                Return _sMoneda
            End Get
            Set(ByVal value As String)
                _sMoneda = value
            End Set
        End Property
        Public Property sFechaVencimiento() As String
            Get
                Return _sFechaVencimiento
            End Get
            Set(ByVal value As String)
                _sFechaVencimiento = value
            End Set
        End Property
        Public Property sCanal() As String
            Get
                Return _sCanal
            End Get
            Set(ByVal value As String)
                _sCanal = value
            End Set
        End Property
        Public Property sSucursal() As String
            Get
                Return _sSucural
            End Get
            Set(ByVal value As String)
                _sSucural = value
            End Set
        End Property
        Public Property sCodigoVendedor() As String
            Get
                Return _sCodigoVendedor
            End Get
            Set(ByVal value As String)
                _sCodigoVendedor = value
            End Set
        End Property
        Public Property sCodigoUsuario() As String
            Get
                Return _sCodigoUsuario
            End Get
            Set(ByVal value As String)
                _sCodigoUsuario = value
            End Set
        End Property
        Public Property sCodigoAtencion() As String
            Get
                Return _sCodigoAtencion
            End Get
            Set(ByVal value As String)
                _sCodigoAtencion = value
            End Set
        End Property
        Public Property sAplicaITF() As String
            Get
                Return _sAcplicaITF
            End Get
            Set(ByVal value As String)
                _sAcplicaITF = value
            End Set
        End Property
        Public Property sIndicadorCommit() As String
            Get
                Return _sIndicadorCommit
            End Get
            Set(ByVal value As String)
                _sIndicadorCommit = value
            End Set
        End Property

    End Class

End Class
