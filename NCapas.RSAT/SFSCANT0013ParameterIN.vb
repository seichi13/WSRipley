Public Class SFSCANT0013ParameterIN

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

    Public Overrides Function ToString() As String

        Dim tramaParameterIN As String = String.Empty

        tramaParameterIN &= LNK_FINANCIADO
        tramaParameterIN &= LNK_CUOTAS
        tramaParameterIN &= LNK_DIFERIDO
        tramaParameterIN &= LNK_TASA_ANUAL
        tramaParameterIN &= LNK_FECHA_CONSUMO
        tramaParameterIN &= LNK_FECHA_FACTURACION

        Return tramaParameterIN

    End Function


End Class
