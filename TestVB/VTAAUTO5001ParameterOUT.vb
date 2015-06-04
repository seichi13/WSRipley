Public Class VTAAUTO5001ParameterOUT

    Private _cuentaENC As String
    Public Property CuentaENC() As String
        Get
            Return _cuentaENC
        End Get
        Set(ByVal value As String)
            _cuentaENC = value
        End Set
    End Property

    Private _productoO As String
    Public Property ProductoO() As String
        Get
            Return _productoO
        End Get
        Set(ByVal value As String)
            _productoO = value
        End Set
    End Property


    Private _codCambioENC As String
    Public Property CodCambioENC() As String
        Get
            Return _codCambioENC
        End Get
        Set(ByVal value As String)
            _codCambioENC = value
        End Set
    End Property

    Private _cuentaPlataformaEnc As String
    Public Property CuentaPlataformaEnc() As String
        Get
            Return _cuentaPlataformaEnc
        End Get
        Set(ByVal value As String)
            _cuentaPlataformaEnc = value
        End Set
    End Property


    Private _tieneSEF As String
    Public Property TieneSEF() As String
        Get
            Return _codCambioENC
        End Get
        Set(ByVal value As String)
            _codCambioENC = value
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

    Private _cambioOK As String
    Public Property CambioOK() As String
        Get
            Return _cambioOK
        End Get
        Set(ByVal value As String)
            _cambioOK = value
        End Set
    End Property

    Private _paso As String
    Public Property Paso() As String
        Get
            Return _paso
        End Get
        Set(ByVal value As String)
            _paso = value
        End Set
    End Property

    Private _codRpta As String
    Public Property CodRpta() As String
        Get
            Return _codRpta
        End Get
        Set(ByVal value As String)
            _codRpta = value
        End Set
    End Property

    Private _desRpta As String
    Public Property DesRpta() As String
        Get
            Return _desRpta
        End Get
        Set(ByVal value As String)
            _desRpta = value
        End Set
    End Property

    Private _resultado As String
    Public Property Resultado() As String
        Get
            Return _resultado
        End Get
        Set(ByVal value As String)
            _resultado = value
        End Set
    End Property

    Private _mensaje As String
    Public Property Mensaje() As String
        Get
            Return _mensaje
        End Get
        Set(ByVal value As String)
            _mensaje = value
        End Set
    End Property


    Sub New(ByVal tramaRespuesta As String)

        CuentaENC = tramaRespuesta.Substring(691, 1)
        ProductoO = tramaRespuesta.Substring(692, 2)
        CodCambioENC = tramaRespuesta.Substring(694, 1)
        CuentaPlataformaEnc = tramaRespuesta.Substring(695, 1)
        TieneSEF = tramaRespuesta.Substring(696, 1)
        ContratoSEF = tramaRespuesta.Substring(697, 20)
        CambioOK = tramaRespuesta.Substring(717, 1)

        Paso = tramaRespuesta.Substring(715, 2)
        CodRpta = tramaRespuesta.Substring(717, 2)
        DesRpta = tramaRespuesta.Substring(719, 200)

    End Sub

    Sub New()
        CuentaENC = String.Empty
        ProductoO = String.Empty
        CodCambioENC = String.Empty
        CuentaPlataformaEnc = String.Empty
        TieneSEF = String.Empty
        ContratoSEF = String.Empty
        CambioOK = String.Empty

        Paso = String.Empty
        CodRpta = String.Empty
        DesRpta = String.Empty
    End Sub

    Public Overrides Function ToString() As String

        Dim tramaParameterOUT As String = String.Empty

        tramaParameterOUT &= CuentaENC
        tramaParameterOUT &= ProductoO
        tramaParameterOUT &= CodCambioENC
        tramaParameterOUT &= CuentaPlataformaEnc
        tramaParameterOUT &= TieneSEF
        tramaParameterOUT &= ContratoSEF
        tramaParameterOUT &= CambioOK

        tramaParameterOUT &= Paso
        tramaParameterOUT &= CodRpta
        tramaParameterOUT &= DesRpta

        Return tramaParameterOUT

    End Function

End Class
